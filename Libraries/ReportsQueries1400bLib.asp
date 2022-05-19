<%
Function BuildReport1431(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de registros de cuentas bancarias
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1431"
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
	Dim sCondition3
	Dim sQuery

	Dim lCurrentPaymentCenterID
	Dim sCurrentPaymentCenterName
	Dim asStateNames
	Dim asBanksNames
	Dim asPath
	Dim asTemp
	Dim iCount
	Dim aiBanksTotals
	Dim aiBanksGrandTotals
	Dim iIndex
	Dim sBanksShortName
	Dim sConceptStatus
	Dim bFirst
	Dim lTotal
	Dim iMin
	Dim iMax

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(Replace(sCondition, "Companies.", "Employees."), "EmployeeTypes.", "Employees.")
	If (InStr(1, oRequest, "AccountStartDate", vbBinaryCompare) > 0) Or (InStr(1, oRequest, "AccountEndDate", vbBinaryCompare) > 0) Then Call GetStartAndEndDatesFromURL("AccountStartDate", "AccountEndDate", "XXXDate", False, sCondition2)
	sCondition2 = Replace(sCondition2, " And ", "", 1, 1, vbBinaryCompare)
	If Len(sCondition2) > 0 Then sCondition2 = " And ((" & Replace(sCondition2, "XXX", "BankAccounts.Start") & ") Or (" & Replace(sCondition2, "XXX", "BankAccounts.End") & ") Or (" & Replace(Replace(sCondition2, "XXX", "BankAccounts.End", 1, 1, vbBinaryCompare), "XXX", "BankAccounts.Start", 1, 1, vbBinaryCompare) & "))"

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ZoneName From Zones Where (ZoneID>-1) And (ParentID=-1) Order By ZoneID", "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select MAX(BankID) As Max From Banks", "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then	
		If Not oRecordset.EOF Then
			iMax = CInt(oRecordset.Fields("Max").Value) + 1
		End If
	End If
	For iMin = 0 To iMax
		asBanksNames = asBanksNames & LIST_SEPARATOR & ""
		aiBanksTotals = aiBanksTotals & LIST_SEPARATOR & "0"
		aiBanksGrandTotals = aiBanksGrandTotals & LIST_SEPARATOR & "0"
	Next
	asBanksNames = Split(asBanksNames, LIST_SEPARATOR)
	aiBanksTotals = Split(aiBanksTotals, LIST_SEPARATOR)
	aiBanksGrandTotals = Split(aiBanksGrandTotals, LIST_SEPARATOR)
	For iIndex = 0 To iMax
		aiBanksTotals(iIndex) = 0
		aiBanksGrandTotals(iIndex) = 0
	Next
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select BankID, BankShortName From Banks Order By BankID", "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		iCount=0
		Do While Not oRecordset.EOF
			asBanksNames(CInt(oRecordset.Fields("BankID").Value) + 1) = SizeText(CStr(CleanStringForHTML(oRecordset.Fields("BankShortName").Value)), " ", 19, 1)
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
		oRecordset.Close
	End If

	oStartDate = Now()

	sQuery = "Select Employees.EmployeeID, EmployeeNumber, Employees.PaymentCenterID, EmployeeName + ' ' + EmployeeLastName  + ' ' + EmployeeLastName2 As EmployeeFullName, BankName," & _
			 " BankAccounts.BankID, AccountNumber, BankAccounts.StartDate, BankAccounts.EndDate, Users.UserLastName + ' ' + Users.UserName As UserFullName, BankAccounts.Active," & _
			 " PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Zones.ZonePath" & _
			 " From Employees, BankAccounts, Banks, Users, Areas, Areas As PaymentCenters, Jobs, Zones As AreasZones, Zones As ParentZones, Zones, Companies" & _
			 " Where (Employees.JobID=Jobs.JobID)" & _
			 " And (Employees.PaymentCenterID=PaymentCenters.AreaID)" & _
			 " And (Jobs.AreaID=Areas.AreaID)" & _
			 " And (Areas.ZoneID=AreasZones.ZoneID)" & _
			 " And (AreasZones.ParentID=ParentZones.ZoneID)" & _
			 " And (PaymentCenters.ZoneID=Zones.ZoneID)" & _
			 " And (Employees.CompanyID=Companies.CompanyID)" & _
			 " And (Employees.PaymentCenterID=PaymentCenters.AreaID)" & _
			 " And (Employees.EmployeeID=BankAccounts.EmployeeID)" & _
			 " And (Banks.BankID=BankAccounts.BankID)" & _
			 " And (Users.UserID=BankAccounts.UserID)" & _
			 " And (Employees.CompanyID=Companies.CompanyID)" & _
			 sCondition & sCondition2 & _
			 " Order By Employees.EmployeeID, BankAccounts.StartDate"
	sErrorDescription = "No se pudieron obtener los registros de cuentas bancarias de los empleados."

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: " & sQuery & " -->" & vbNewLine
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
					sHeaderContents = Replace(sHeaderContents, "<TITLE />", CleanStringForHTML("Reporte de cuentas bancarias de los empleados."))
					sHeaderContents = Replace(sHeaderContents, "<MONTH_ID />", CleanStringForHTML(asMonthNames_es(iMonth)))
					sHeaderContents = Replace(sHeaderContents, "<YEAR_ID />", iYear)
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, 1))
				End If
				lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
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
							sRowContents = sRowContents & "<TD>CLAVE DEL BANCO</TD>"
							sRowContents = sRowContents & "<TD>TOTAL</TD>"
							sRowContents = sRowContents & "</FONT></TR>"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
							For iIndex = 0 To UBound(aiBanksTotals)
								lTotal = CInt(aiBanksTotals(iIndex))
								If lTotal > 0 Then
									sBanksShortName = Trim(asBanksNames(CInt(iIndex)))
									sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
										sRowContents = sRowContents & "<TD>" & sBanksShortName & "</TD>"
										sRowContents = sRowContents & "<TD>" & lTotal & "</TD>"
									sRowContents = sRowContents & "</FONT></TR>"
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
								End If
							Next
							For iIndex = 0 To UBound(aiBanksTotals)
								aiBanksTotals(iIndex) = 0
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
						sRowContents = "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
							sRowContents = sRowContents & "<TR>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">No.Emp.</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre del empleado</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Numero de cuenta</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Banco</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Fecha de fin</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Clave del centro de trabajo</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Nombre del centro de trabajo</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Usuario que capturo</FONT></TD>"
								sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">Status del registro</FONT></TD>"
							sRowContents = sRowContents & "</TR>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					End If
					lCurrentPaymentCenterID = CLng(oRecordset.Fields("PaymentCenterID").Value)
					sCurrentPaymentCenterName = CStr(oRecordset.Fields("PaymentCenterName").Value)
					sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeFullName").Value)) & "</FONT></TD>"
						If (InStr(1, CStr(oRecordset.Fields("AccountNumber").Value), ".", vbBinaryCompare) > 0) Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML("=T(""Cheque"")") & "</FONT></TD>"
						Else
							asTemp = Split(CStr(oRecordset.Fields("AccountNumber").Value), LIST_SEPARATOR)
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML("=T(""" & asTemp(0)) & """)</FONT></TD>"
						End If
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("BankName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)) & "</FONT></TD>"
						If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">A la fecha</FONT></TD>"
						Else
							sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & "</FONT></TD>"
						End If
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterName").Value)) & "</FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(CStr(oRecordset.Fields("UserFullName").Value)) & "</FONT></TD>"
						Select Case CInt(oRecordset.Fields("Active").Value)
							Case 0
								sConceptStatus = "En proceso"
							Case 1
								sConceptStatus = "Activo"
						End Select
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & sConceptStatus & "</FONT></TD>"
					sRowContents = sRowContents & "</FONT></TR>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					aiBanksTotals(CInt(oRecordset.Fields("BankID").Value)+1) = aiBanksTotals(CInt(oRecordset.Fields("BankID").Value)+1) + 1
					aiBanksGrandTotals(CInt(oRecordset.Fields("BankID").Value)+1) = aiBanksGrandTotals(CInt(oRecordset.Fields("BankID").Value)+1) + 1
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
					sRowContents = sRowContents & "<TD>CLAVE DEL BANCO</TD>"
					sRowContents = sRowContents & "<TD>TOTAL</TD>"
					sRowContents = sRowContents & "</FONT></TR>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					For iIndex = 0 To UBound(aiBanksTotals)
						lTotal = CInt(aiBanksTotals(iIndex))
						If lTotal > 0 Then
							sBanksShortName = Trim(asBanksNames(CInt(iIndex)))
							sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
								sRowContents = sRowContents & "<TD>" & sBanksShortName & "</TD>"
								sRowContents = sRowContents & "<TD>" & lTotal & "</TD>"
							sRowContents = sRowContents & "</FONT></TR>"
							lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						End If
					Next
					For iIndex = 0 To UBound(aiBanksTotals)
						aiBanksTotals(iIndex) = 0
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
				sRowContents = sRowContents & "<TD>CLAVE DEL BANCO</TD>"
				sRowContents = sRowContents & "<TD>TOTAL</TD>"
				sRowContents = sRowContents & "</FONT></TR>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				For iIndex = 0 To UBound(aiBanksGrandTotals)
					lTotal = CInt(aiBanksGrandTotals(iIndex))
					If lTotal > 0 Then
						sBanksShortName = Trim(asBanksNames(CInt(iIndex)))
						sRowContents = "<TR><FONT FACE=""Arial"" SIZE=""2"">"
							sRowContents = sRowContents & "<TD>" & sBanksShortName & "</TD>"
							sRowContents = sRowContents & "<TD>" & lTotal & "</TD>"
						sRowContents = sRowContents & "</FONT></TR>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					End If
				Next
				For iIndex = 0 To UBound(aiBanksTotals)
					aiBanksTotals(iIndex) = 0
				Next
				sRowContents = "</TABLE>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)				
				oRecordset.Close
				lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudieron guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
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
	BuildReport1431 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1432(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de registros de cuentas bancarias
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1432"
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
	Dim sCondition
	Dim sCondition2
	Dim sCondition3
	Dim sQuery
	Dim sBankAccountData
	Dim lCount
	Dim sBankAnt
	Dim sBankAct
	Dim sCuentaAnt
	Dim sCuentaAct
	Dim lFechaAnt
	Dim lFechaAct
	Dim lEmployee
	Dim sEmployeeName
	Dim sDocumentName
	Dim asTemp

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)

	sCondition = Replace(Replace(Replace(sCondition, "Companies.", "Employees."), "EmployeeTypes.", "Employees."), "Areas.ZoneID", "Zones.ZoneID")

	If (InStr(1, oRequest, "RegistrationStartDate", vbBinaryCompare) > 0) Or (InStr(1, oRequest, "RegistrationEndDate", vbBinaryCompare) > 0) Then Call GetStartAndEndDatesFromURL("RegistrationStartDate", "RegistrationEndDate", "XXXDate", False, sCondition3)
	sCondition3 = Replace(sCondition3, " And ", "", 1, 1, vbBinaryCompare)
	If Len(sCondition3) > 0 Then sCondition3 = " And ((" & Replace(sCondition3, "XXX", "BankAccounts.Registration") & ") Or (" & Replace(sCondition3, "XXX", "BankAccounts.Registration") & ") Or (" & Replace(Replace(sCondition3, "XXX", "BankAccounts.Registration", 1, 1, vbBinaryCompare), "XXX", "BankAccounts.Registration", 1, 1, vbBinaryCompare) & "))"

	oStartDate = Now()

	sBankAccountData = GetFileContents(Server.MapPath("Templates\HeaderForReport_1432.htm"), sErrorDescription)

	sQuery = "Select Distinct BankAccounts.EmployeeID, BankAccounts.RegistrationDate" & _
			 " From Employees, BankAccounts, Banks, Jobs, Areas, Zones, Zones As Zones2, Zones As Zones3" & _
			 " Where (Employees.JobID=Jobs.JobID)" & _
			 " And (Jobs.AreaID=Areas.AreaID)" & _
			 " And (Areas.ZoneID=Zones3.ZoneID)" & _
			 " And (Zones3.ParentID=Zones2.ZoneID)" & _
			 " And (Zones2.ParentID=Zones.ZoneID)" & _
			 " And (Employees.EmployeeID=BankAccounts.EmployeeID)" & _
			 " And (BankAccounts.BankID=Banks.BankID) " & sCondition & sCondition3
	sErrorDescription = "No se pudieron obtener los registros de cuentas bancarias de los empleados."

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: " & sQuery & " -->" & vbNewLine
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
					sHeaderContents = Replace(sHeaderContents, "<TITLE />", CleanStringForHTML("Listado de actualización de cuentas bancarias"))
					sHeaderContents = Replace(sHeaderContents, "<MONTH_ID />", CleanStringForHTML(asMonthNames_es(iMonth)))
					sHeaderContents = Replace(sHeaderContents, "<YEAR_ID />", iYear)
					sHeaderContents = Replace(sHeaderContents, "<CURRENT_DATE />", DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, 1))
				End If
				lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
				sRowContents = "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>No.Emp.</B></FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>Nombre del empleado</B></FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>Numero de cuenta anterior</B></FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>Numero de cuenta actual</B></FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>Fecha de inicio anterior</B></FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>Fecha de inicio actual</B></FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>Banco anterior</B></FONT></TD>"
						sRowContents = sRowContents & "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>Banco actual</B></FONT></TD>"
					sRowContents = sRowContents & "</TR>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				Do While Not oRecordset.EOF
					sQuery = "Select TOP 2 Employees.EmployeeID, EmployeeNumber, EmployeeName + ' ' + EmployeeLastName  + ' ' + EmployeeLastName2 As EmployeeFullName, BankShortName, BankName," & _
							 " AccountNumber, BankAccounts.StartDate, BankAccounts.EndDate, ZoneName" & _
							 " From Employees, BankAccounts, Banks, Jobs, Areas, Zones" & _
							 " Where (Employees.JobID = Jobs.JobID)" & _
							 " And (Jobs.AreaID = Areas.AreaID)" & _
							 " And (Areas.ZoneID = Zones.ZoneID)" & _
							 " And (Employees.EmployeeID = BankAccounts.EmployeeID)" & _
							 " And (BankAccounts.BankID = Banks.BankID) " & _
							 " And (BankAccounts.EmployeeID =" & CLng(oRecordset.Fields("EmployeeID").Value) & ")" & _
							 " And (BankAccounts.RegistrationDate<=" & CLng(oRecordset.Fields("RegistrationDate").Value) & ")" & _
							 " Order By BankAccounts.StartDate Desc" ' & sCondition & sCondition2 & sCondition3 & _
					sErrorDescription = "No se pudieron obtener los registros de cuentas bancarias de los empleados."

					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset1)
					If lErrorNumber = 0 Then
						If Not oRecordset1.EOF Then
							lCount = 0
							sContents = sBankAccountData
							Do While Not oRecordset1.EOF
								asTemp = Split(CStr(oRecordset1.Fields("AccountNumber").Value), LIST_SEPARATOR)
								If lCount = 0 Then
									lEmployee = "=T(""" & Right("000000" & CLng(oRecordset1.Fields("EmployeeID").Value), Len("000000")) & """)"
									sEmployeeName = CStr(oRecordset1.Fields("EmployeeFullName").Value)
									sCuentaAct = "=T(""" & Right("0000000000000000" & asTemp(0), Len("0000000000000000")) & """)"
									lFechaAct = CLng(oRecordset1.Fields("StartDate").Value)
									sBankAct = CStr(oRecordset1.Fields("BankShortName").Value)
								Else
									sCuentaAnt = "=T(""" & Right("0000000000000000" & asTemp(0), Len("0000000000000000")) & """)"
									lFechaAnt = CLng(oRecordset1.Fields("StartDate").Value)
									sBankAnt = CStr(oRecordset1.Fields("BankShortName").Value)
								End If
								lCount = 1
								oRecordset1.MoveNext
								If Err.number <> 0 Then Exit Do
							Loop
						End If
						sContents = Replace(sContents, "<EMPLOYEE_NUMBER />", CleanStringForHTML(CStr(lEmployee)))
						sContents = Replace(sContents, "<EMPLOYEE_FULL_NAME />", CleanStringForHTML(CStr(sEmployeeName)))
						If (InStr(1, sCuentaAct, ".", vbBinaryCompare) > 0) Then
							sContents = Replace(sContents, "<CUENTAACT />", CleanStringForHTML("Cheque"))
						Else
							sContents = Replace(sContents, "<CUENTAACT />", CleanStringForHTML(sCuentaAct))
						End If
						sContents = Replace(sContents, "<FECHAACT />", CleanStringForHTML(DisplayDateFromSerialNumber(CLng(lFechaAct), -1, -1, -1)))
						sContents = Replace(sContents, "<BANKACT />", CleanStringForHTML(sBankAct))
						If Len(sCuentaAnt) > 0 Then
							If (InStr(1, sCuentaAnt, ".", vbBinaryCompare) > 0) Then
								sContents = Replace(sContents, "<CUENTAANT />", CleanStringForHTML("Cheque"))								
								sContents = Replace(sContents, "<FECHAANT />", CleanStringForHTML(DisplayDateFromSerialNumber(CLng(lFechaAnt), -1, -1, -1)))
								sContents = Replace(sContents, "<BANKANT />", CleanStringForHTML(sBankAnt))
							Else
								sContents = Replace(sContents, "<CUENTAANT />", CleanStringForHTML(sCuentaAnt))
								sContents = Replace(sContents, "<FECHAANT />", CleanStringForHTML(DisplayDateFromSerialNumber(CLng(lFechaAnt), -1, -1, -1)))
								sContents = Replace(sContents, "<BANKANT />", CleanStringForHTML(sBankAnt))
							End If
						Else
							sContents = Replace(sContents, "<CUENTAACT />", CleanStringForHTML("Ninguna"))
							sContents = Replace(sContents, "<FECHAANT />", CleanStringForHTML("NA"))
							sContents = Replace(sContents, "<BANKANT />", CleanStringForHTML("NA"))
						End If
						sContents = "<TR>" & sContents & "</TR>"
						lErrorNumber = AppendTextToFile(sDocumentName, sContents, sErrorDescription)
					End If
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			End If
			oRecordset.Close
			lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudieron guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
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
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1432 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1433(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de estímulos
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1433"
	Dim aiDays
	Dim sCondition
	Dim sTables
	Dim lPayrollID
	Dim lForPayrollID
	Dim lPayrollID2
	Dim oRecordset
	Dim lCurrentID
	Dim sCurrentID
	Dim bDisplay
	Dim sDate
	Dim sFileName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	oStartDate = Now()
	aiDays = Split("0,0", ",")
	aiDays(0) = 0
	aiDays(1) = 0

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	lPayrollID2 = AddDaysToSerialDate(lForPayrollID, -30)
	sCondition = Replace(Replace(Replace(Replace(Replace(sCondition, "Banks.", "BankAccounts."), "Companies", "EmployeesHistoryList"), "Employees.", "EmployeesHistoryList."), "EmployeeTypes", "EmployeesHistoryList"), "PaymentCenters", "EmployeesHistoryList")
	If InStr(1, sCondition, "BankAccounts", vbBinaryCompare) > 0 Then
		sTables = ", BankAccounts"
		sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=BankAccounts.EmployeeID) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1)"
	End If

	sErrorDescription = "No se pudo limpiar el repositorio temporal del reporte."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Report1433 Where (UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ")", "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	If lErrorNumber = 0 Then
		sErrorDescription = "No se pudo llenar el repositorio temporal del reporte."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Report1433 (UserID, EmployeeID, AntiquityDays) Select Distinct " & aLoginComponent(N_USER_ID_LOGIN) & " As UserID, EmployeesHistoryList.EmployeeID, 0 As AntiquityDays From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas, Zones, Zones As Zones2, Zones As ParentZones" & sTables & ", Positions Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID)  And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (ConceptID In (40,41,42,43,50)) " & sCondition, "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		Response.Write vbNewLine & "<!-- Query: Insert Into Report1433 (UserID, EmployeeID, AntiquityDays) Select Distinct " & aLoginComponent(N_USER_ID_LOGIN) & " As UserID, EmployeesHistoryList.EmployeeID, 0 As AntiquityDays From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas, Zones, Zones As Zones2, Zones As ParentZones" & sTables & ", Positions Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID)  And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (ConceptID In (40,41,42,43,50)) " & sCondition & " -->" & vbNewLine
	End If
	If lErrorNumber = 0 Then
		sErrorDescription = "No se pudieron obtener las antigüedades de los empleados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeDate, EmployeesHistoryList.EndDate, StatusEmployees.Active, Reasons.ActiveEmployeeID From Report1433, EmployeesHistoryList, StatusEmployees, Reasons Where (Report1433.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (StatusEmployees.Active=1) And (ActiveEmployeeID=1) And (EmployeesHistoryList.EmployeeDate<=" & lPayrollID2 & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (Report1433.UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ") Order By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeDate Desc", "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeDate, EmployeesHistoryList.EndDate, StatusEmployees.Active, Reasons.ActiveEmployeeID From Report1433, EmployeesHistoryList, StatusEmployees, Reasons Where (Report1433.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (StatusEmployees.Active=1) And (ActiveEmployeeID=1) And (EmployeesHistoryList.EmployeeDate<=" & lPayrollID2 & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (Report1433.UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ") Order By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeDate Desc -->" & vbNewLine
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
				Do While Not oRecordset.EOF
					If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
						sErrorDescription = "No se pudo actualizar la antigüedad del empleado."
						lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, "Update Report1433 Set AntiquityDays=" & aiDays(1) & " Where (EmployeeID=" & lCurrentID & ") And (UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ")", "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription)
						aiDays(0) = 0
						aiDays(1) = 0
						lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
					End If
					If CLng(oRecordset.Fields("EndDate").Value) > lPayrollID2 Then
						aiDays(1) = aiDays(1) + Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value)), GetDateFromSerialNumber(lPayrollID2))) + 1
					Else
						aiDays(1) = aiDays(1) + Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value)), GetDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)))) + 1
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				sErrorDescription = "No se pudo actualizar la antigüedad del empleado."
				lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, "Update Report1433 Set AntiquityDays=" & aiDays(1) & " Where (EmployeeID=" & lCurrentID & ") And (UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ")", "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription)
			End If
		End If
	End If

	aiDays = Split("0,0", ",")
	aiDays(0) = 0
	aiDays(1) = 0
	sErrorDescription = "No se pudieron obtener las antigüedades de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesAntiquitiesLKP.EmployeeID, AntiquityYears, AntiquityMonths, EmployeesAntiquitiesLKP.AntiquityDays From Report1433, EmployeesAntiquitiesLKP, EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons Where (Report1433.EmployeeID=EmployeesAntiquitiesLKP.EmployeeID) And (EmployeesAntiquitiesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (Report1433.UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ") Order By EmployeesAntiquitiesLKP.EmployeeID", "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select EmployeesAntiquitiesLKP.EmployeeID, AntiquityYears, AntiquityMonths, EmployeesAntiquitiesLKP.AntiquityDays From Report1433, EmployeesAntiquitiesLKP, EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons Where (Report1433.EmployeeID=EmployeesAntiquitiesLKP.EmployeeID) And (EmployeesAntiquitiesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (Report1433.UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ") Order By EmployeesAntiquitiesLKP.EmployeeID -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
			Do While Not oRecordset.EOF
				If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
					lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, "Update Report1433 Set AntiquityDays=AntiquityDays+" & aiDays(1) & " Where (EmployeeID=" & lCurrentID & ") And (UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ")", "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription)
					aiDays(0) = 0
					aiDays(1) = 0
					lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
				End If
				aiDays(1) = aiDays(1) + (CInt(oRecordset.Fields("AntiquityYears").Value) * 365) + Int(CInt(oRecordset.Fields("AntiquityMonths").Value) * 30.4) + CInt(oRecordset.Fields("AntiquityDays").Value)
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
			lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, "Update Report1433 Set AntiquityDays=AntiquityDays+" & aiDays(1) & " Where (EmployeeID=" & lCurrentID & ") And (UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ")", "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription)
		End If
	End If

	aiDays = Split("0,0", ",")
	aiDays(0) = 0
	aiDays(1) = 0
	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesAbsencesLKP.EmployeeID, EmployeesAbsencesLKP.OcurredDate, EmployeesAbsencesLKP.EndDate From Report1433, EmployeesAbsencesLKP, EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons Where (Report1433.EmployeeID=EmployeesAbsencesLKP.EmployeeID) And (EmployeesAbsencesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID2 & ") And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesAbsencesLKP.AbsenceID In (10,95)) And (EmployeesAbsencesLKP.OcurredDate<=" & lPayrollID2 & ") And (Report1433.UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ") Order By EmployeesAbsencesLKP.EmployeeID", "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select EmployeesAbsencesLKP.EmployeeID, EmployeesAbsencesLKP.OcurredDate, EmployeesAbsencesLKP.EndDate From Report1433, EmployeesAbsencesLKP, EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons Where (Report1433.EmployeeID=EmployeesAbsencesLKP.EmployeeID) And (EmployeesAbsencesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID2 & ") And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesAbsencesLKP.AbsenceID In (10,95)) And (EmployeesAbsencesLKP.OcurredDate<=" & lPayrollID2 & ") And (Report1433.UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ") Order By EmployeesAbsencesLKP.EmployeeID -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
			Do While Not oRecordset.EOF
				If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
					lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, "Update Report1433 Set AntiquityDays=AntiquityDays-" & aiDays(1) & " Where (EmployeeID=" & lCurrentID & ") And (UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ")", "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription)
					aiDays(0) = 0
					aiDays(1) = 0
					lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
				End If

				If CLng(oRecordset.Fields("EndDate").Value) > lPayrollID2 Then
					aiDays(1) = aiDays(1) + Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("OcurredDate").Value)), lPayrollID2)) + 1
				Else
					aiDays(1) = aiDays(1) + Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("OcurredDate").Value)), GetDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)))) + 1
				End If
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
			lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, "Update Report1433 Set AntiquityDays=AntiquityDays-" & aiDays(1) & " Where (EmployeeID=" & lCurrentID & ") And (UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ")", "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription)
		End If
	End If

	sErrorDescription = "No se pudieron obtener los registros de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, EmployeeNumber, Areas.AreaCode, PositionShortName, RecordDate, ConceptID, ConceptAmount, AntiquityDays From Report1433, Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas, Zones, Zones As Zones2, Zones As ParentZones" & sTables & ", Positions Where (Report1433.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (ConceptID In (1,4,5,6,7,8,40,41,42,43,50)) And (EmployeesHistoryList.EmployeeID In (Select Distinct EmployeeID From Payroll_" & lPayrollID & " Where (ConceptID In (40,41,42,43,50)))) And (Report1433.UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ") " & sCondition & " Order By Areas.AreaCode, EmployeeNumber, RecordDate, ConceptID", "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryList.EmployeeID, EmployeeNumber, Areas.AreaCode, PositionShortName, RecordDate, ConceptID, ConceptAmount, AntiquityDays From Report1433, Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas, Zones, Zones As Zones2, Zones As ParentZones" & sTables & ", Positions Where (Report1433.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (ConceptID In (1,4,5,6,7,8,40,41,42,43,50)) And (EmployeesHistoryList.EmployeeID In (Select Distinct EmployeeID From Payroll_" & lPayrollID & " Where (ConceptID In (40,41,42,43,50)))) And (Report1433.UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ") " & sCondition & " Order By Areas.AreaCode, EmployeeNumber, RecordDate, ConceptID -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFileName = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls")
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Replace(sFileName, ".xls", ".zip") & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			Response.Flush()

			lErrorNumber = AppendTextToFile(sFileName, "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">", sErrorDescription)
				asColumnsTitles = Split("Empleado,Centro de trabajo,Puesto,Antigüedad,Fecha de pago,Imputación,Concepto 01,Concepto 04,Concepto 05,Concepto 06,Concepto 07,Concepto 08,Concepto 37,Concepto 38,Concepto 39,Concepto 40,Concepto 49", ",", -1, vbBinaryCompare)
				asCellWidths = Split("100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFileName, GetTableHeaderPlainText(asColumnsTitles, True, ""), sErrorDescription)

				sCurrentID = ""
				bDisplay = False
				asCellAlignments = Split(",,,,,,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					If StrComp(sCurrentID, CStr(oRecordset.Fields("EmployeeID").Value) & "," & CStr(oRecordset.Fields("RecordDate").Value), vbBinaryCompare) <> 0 Then
						If (Len(sCurrentID) > 0) And bDisplay Then
							sRowContents = Replace(sRowContents, "<CONCEPT_01 />", "0.00")
							sRowContents = Replace(sRowContents, "<CONCEPT_04 />", "0.00")
							sRowContents = Replace(sRowContents, "<CONCEPT_05 />", "0.00")
							sRowContents = Replace(sRowContents, "<CONCEPT_06 />", "0.00")
							sRowContents = Replace(sRowContents, "<CONCEPT_07 />", "0.00")
							sRowContents = Replace(sRowContents, "<CONCEPT_08 />", "0.00")
							sRowContents = Replace(sRowContents, "<CONCEPT_40 />", "0.00")
							sRowContents = Replace(sRowContents, "<CONCEPT_41 />", "0.00")
							sRowContents = Replace(sRowContents, "<CONCEPT_42 />", "0.00")
							sRowContents = Replace(sRowContents, "<CONCEPT_43 />", "0.00")
							sRowContents = Replace(sRowContents, "<CONCEPT_50 />", "0.00")
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							lErrorNumber = AppendTextToFile(sFileName, GetTableRowText(asRowContents, True, ""), sErrorDescription)
						End If
						sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value)) & "&nbsp;"
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
						aiDays = Split("0,0,0", ",")
						aiDays(2) = CLng(oRecordset.Fields("AntiquityDays").Value) - DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("RecordDate").Value)), GetDateFromSerialNumber(lForPayrollID))
						aiDays(0) = Int(aiDays(2) / 365)
						aiDays(2) = aiDays(2) Mod 365
						aiDays(1) = Int(aiDays(2) / 30.4)
						aiDays(2) = Int(aiDays(2) - (aiDays(1) * 30.4))
						sRowContents = sRowContents & TABLE_SEPARATOR & aiDays(0) & "-" & Right(("00" & aiDays(1)), Len("00")) & "-" & Right(("00" & aiDays(2)), Len("00")) & "&nbsp;"
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(lForPayrollID, -1, -1, -1)
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("RecordDate").Value), -1, -1, -1)
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_01 />"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_04 />"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_05 />"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_06 />"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_07 />"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_08 />"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_40 />"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_41 />"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_42 />"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_43 />"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_50 />"

						sCurrentID = CStr(oRecordset.Fields("EmployeeID").Value) & "," & CStr(oRecordset.Fields("RecordDate").Value)
						bDisplay = False
					End If
					sRowContents = Replace(sRowContents, "<CONCEPT_" & Right(("00" & CStr(oRecordset.Fields("ConceptID").Value)), Len("00")) & " />", FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True))
					If (InStr(1, ",40,41,42,43,50,", "," & CStr(oRecordset.Fields("ConceptID").Value) & ",", vbBinaryCompare) > 0) Then bDisplay = True
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				If bDisplay Then
					sRowContents = Replace(sRowContents, "<CONCEPT_01 />", "0.00")
					sRowContents = Replace(sRowContents, "<CONCEPT_04 />", "0.00")
					sRowContents = Replace(sRowContents, "<CONCEPT_05 />", "0.00")
					sRowContents = Replace(sRowContents, "<CONCEPT_06 />", "0.00")
					sRowContents = Replace(sRowContents, "<CONCEPT_07 />", "0.00")
					sRowContents = Replace(sRowContents, "<CONCEPT_08 />", "0.00")
					sRowContents = Replace(sRowContents, "<CONCEPT_40 />", "0.00")
					sRowContents = Replace(sRowContents, "<CONCEPT_41 />", "0.00")
					sRowContents = Replace(sRowContents, "<CONCEPT_42 />", "0.00")
					sRowContents = Replace(sRowContents, "<CONCEPT_43 />", "0.00")
					sRowContents = Replace(sRowContents, "<CONCEPT_50 />", "0.00")
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(sFileName, GetTableRowText(asRowContents, True, ""), sErrorDescription)
				End If
			lErrorNumber = AppendTextToFile(sFileName, "</TABLE>", sErrorDescription)

			lErrorNumber = ZipFolder(sFileName, Replace(sFileName, ".xls", ".zip"), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudieron guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				lErrorNumber = DeleteFile(sFileName, sErrorDescription)
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
	BuildReport1433 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1434(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de incidencias
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1434"
	Dim sCondition
	Dim sTables
	Dim lPayrollID
	Dim lForPayrollID
	Dim lPayrollNumber
	Dim oRecordset
	Dim lCurrentID
	Dim lAbsenceID
	Dim dJourneyFactor
	Dim dShiftsWorkingHours
	Dim lJourneyTypeID
	Dim sDates
	Dim iDays
	Dim sDate
	Dim lEndDate
	Dim dAmount
	Dim adTotalAmount
	Dim dConcept01
	Dim dConcept01a
	Dim dConcept04
	Dim dConcept07
	Dim dConcept08
	Dim lConceptID
	Dim adTotal
	Dim iTotalHours
	Dim sFileName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	oStartDate = Now()
	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	lPayrollNumber = (CInt(Left(lForPayrollID, Len("YYYY"))) * 100) + CInt(GetPayrollNumber(lForPayrollID))
	sCondition = Replace(Replace(Replace(Replace(Replace(sCondition, "Banks.", "BankAccounts."), "Companies", "EmployeesHistoryList"), "Employees.", "EmployeesHistoryList."), "EmployeeTypes", "EmployeesHistoryList"), "PaymentCenters", "EmployeesHistoryList")
	If InStr(1, sCondition, "BankAccounts", vbBinaryCompare) > 0 Then
		sTables = ", BankAccounts"
		sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=BankAccounts.EmployeeID) And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1)"
	End If

	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct EmployeesHistoryList.EmployeeID, Employees.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, Areas.AreaCode, PositionShortName, Shifts.JourneyTypeID, JourneyTypes.JourneyFactor, Shifts.WorkingHours As ShiftsWorkingHours, RecordDate, ConceptID, ConceptAmount, Absences.AbsenceID, AbsenceShortName, ConceptsIDs, EmployeesAbsencesLKP.OcurredDate, EmployeesAbsencesLKP.EndDate From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Employees, Journeys, Shifts, JourneyTypes, Areas, Zones, Positions, EmployeesAbsencesLKP, Absences" & sTables & " Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (Shifts.JourneyTypeID=JourneyTypes.JourneyTypeID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.EmployeeID=EmployeesAbsencesLKP.EmployeeID) And (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (ConceptID In (1,4,7,8,52,71)) And (Absences.AbsenceID In (3,10,11,16,18,19,20,24,25,26,28,92,93,94, 1,2,4,5,21,23,27)) And (AppliedDate In (0," & lPayrollID & ")) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (JourneyTypes.JourneyFactor>0) And (EmployeesAbsencesLKP.Active=1) And (EmployeesHistoryList.EmployeeID In (Select Distinct EmployeeID From Payroll_" & lPayrollID & " Where (ConceptID In (52,71)))) And (EmployeesAbsencesLKP.AppliedDate=" & lPayrollID & ") " & sCondition & " Order By Employees.EmployeeNumber, Absences.AbsenceID, RecordDate, ConceptID", "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select Distinct EmployeesHistoryList.EmployeeID, Employees.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, Areas.AreaCode, PositionShortName, Shifts.JourneyTypeID, JourneyTypes.JourneyFactor, Shifts.WorkingHours As ShiftsWorkingHours, RecordDate, ConceptID, ConceptAmount, Absences.AbsenceID, AbsenceShortName, ConceptsIDs, EmployeesAbsencesLKP.OcurredDate, EmployeesAbsencesLKP.EndDate From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Employees, Journeys, Shifts, JourneyTypes, Areas, Zones, Positions, EmployeesAbsencesLKP, Absences" & sTables & " Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (Shifts.JourneyTypeID=JourneyTypes.JourneyTypeID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.EmployeeID=EmployeesAbsencesLKP.EmployeeID) And (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (ConceptID In (1,4,7,8,52,71)) And (Absences.AbsenceID In (3,10,11,16,18,19,20,24,25,26,28,92,93,94, 1,2,4,5,21,23,27)) And (AppliedDate In (0," & lPayrollID & ")) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (JourneyTypes.JourneyFactor>0) And (EmployeesAbsencesLKP.Active=1) And (EmployeesHistoryList.EmployeeID In (Select Distinct EmployeeID From Payroll_" & lPayrollID & " Where (ConceptID In (52,71)))) And (EmployeesAbsencesLKP.AppliedDate=" & lPayrollID & ") " & sCondition & " Order By Employees.EmployeeNumber, Absences.AbsenceID, RecordDate, ConceptID -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFileName = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls")
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Replace(sFileName, ".xls", ".zip") & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			Response.Flush()

			lErrorNumber = AppendTextToFile(sFileName, "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">", sErrorDescription)
				asColumnsTitles = Split("Empleado,Centro de trabajo,Quincena,Puesto,Jornada,Sueldo,Riesgos profesionales,Turno opcional,Percepción adicional,Nombre,Diferido,Concepto,Importe,Fechas,Monto aplicado,Días,Diferido próxima quincena", ",", -1, vbBinaryCompare)
				asCellWidths = Split("100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
				lErrorNumber = AppendTextToFile(sFileName, GetTableHeaderPlainText(asColumnsTitles, True, ""), sErrorDescription)

				iTotalHours = 0
				adTotalAmount = Split("0,0", ",")
				adTotalAmount(0) = 0
				adTotalAmount(1) = 0
				lCurrentID = -2
				lAbsenceID = -2
				asCellAlignments = Split(",,,,CENTER,RIGHT,RIGHT,RIGHT,RIGHT,,RIGHT,LEFT,RIGHT,LEFT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					If (lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value)) Or (lAbsenceID <> CLng(oRecordset.Fields("AbsenceID").Value)) Or ((InStr(1, ",52,71,", "," & CStr(oRecordset.Fields("ConceptID").Value) & ",", vbBinaryCompare) > 0) And (lConceptID <> -2) And (lConceptID <> CLng(oRecordset.Fields("ConceptID").Value))) Then
						If ((lConceptID = 52) And (InStr(1, ",3,10,11,16,18,19,20,24,25,26,28,92,93,94,", "," & lAbsenceID & ",", vbBinaryCompare) > 0)) _
						Or ((lConceptID = 71) And (InStr(1, ",1,2,4,5,21,23,27,", "," & lAbsenceID & ",", vbBinaryCompare) > 0)) Then
							sRowContents = Replace(sRowContents, "<CONCEPT_01 />", "0.00")
							sRowContents = Replace(sRowContents, "<CONCEPT_04 />", "0.00")
							sRowContents = Replace(sRowContents, "<CONCEPT_07 />", "0.00")
							sRowContents = Replace(sRowContents, "<CONCEPT_08 />", "0.00")
							adTotal = Split("0,0", ",")
							adTotal(0) = 0
							adTotal(1) = 0
							Select Case lAbsenceID
								Case 1
									Select Case lJourneyTypeID
										Case 1
											adTotal(0) = adTotal(0) + (Int(iTotalHours / 3) / dJourneyFactor / 4)
										Case 2, 3
											adTotal(0) = adTotal(0) + (Int(iTotalHours / 2) / dJourneyFactor / 4)
										Case 4
											adTotal(0) = adTotal(0) + (iTotalHours / dJourneyFactor / 4)
									End Select
								Case 2, 23, 27
									adTotal(0) = adTotal(0) + (iTotalHours / dJourneyFactor / 4)
								Case 4
									adTotal(0) = adTotal(0) + (iTotalHours / dJourneyFactor / 2)
								Case 3, 11, 18, 19, 20, 24, 25, 26, 28, 93, 94
									adTotal(0) = adTotal(0) + (iTotalHours / dJourneyFactor)
								Case 5
									adTotal(0) = adTotal(0) + (iTotalHours / dJourneyFactor / 6)
								Case 10
									If lJourneyTypeID <> 1 Then
										adTotal(0) = adTotal(0) + (iTotalHours / dJourneyFactor)
									Else
										adTotal(1) = adTotal(1) + (iTotalHours / dJourneyFactor)
									End If
								Case 16
									adTotal(0) = adTotal(0) + (iTotalHours / dJourneyFactor / 3)
								Case 21
									adTotal(0) = adTotal(0) + (iTotalHours / (dJourneyFactor * dShiftsWorkingHours * 2))
								Case 92
									adTotal(0) = adTotal(0) + (iTotalHours / dJourneyFactor * 2 / 5)
							End Select
							dAmount = 0
							If adTotal(1) > 0 Then 
								dAmount = ((dConcept01 + dConcept04) * adTotal(1) * 1.4) + ((dConcept07 + dConcept08) * adTotal(1))
							Else
								dAmount = (dConcept01 + dConcept04 + dConcept07 + dConcept08) * adTotal(0)
							End If
							sRowContents = Replace(sRowContents, "<ABSENCE_AMOUNT />", FormatNumber(dAmount, 2, True, False, True))
							If lConceptID = 52 Then
								adTotalAmount(0) = adTotalAmount(0) + dAmount
							Else
								adTotalAmount(1) = adTotalAmount(1) + dAmount
							End If
							sRowContents = Replace(sRowContents, "<CONCEPT_50_70 />", "0.00")
							sRowContents = Replace(sRowContents, "<CONCEPT_99 />", "0.00")
							If Len(sDates) > 0 Then sDates = Left(sDates, (Len(sDates) - Len(", ")))
							sRowContents = Replace(sRowContents, "<OCURRED_DATES />", sDates)
							sRowContents = Replace(sRowContents, "<DAYS />", iDays)
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							lErrorNumber = AppendTextToFile(sFileName, GetTableRowText(asRowContents, True, ""), sErrorDescription)
						End If
						If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
							sRowContents = GetFileContents(sFileName, sErrorDescription)
							If adTotalAmount(0) > (dConcept01a * 0.30) Then
								sRowContents = Replace(sRowContents, "<CONCEPT_999_52 />", FormatNumber((adTotalAmount(0) - (dConcept01a * 0.30)), 2, True, False, True))
							Else
								sRowContents = Replace(sRowContents, "<CONCEPT_999_52 />", "0.00")
							End If
							If adTotalAmount(1) > (dConcept01a * 0.30) Then
								sRowContents = Replace(sRowContents, "<CONCEPT_999_71 />", FormatNumber((adTotalAmount(1) - (dConcept01a * 0.30)), 2, True, False, True))
							Else
								sRowContents = Replace(sRowContents, "<CONCEPT_999_71 />", "0.00")
							End If
							lErrorNumber = SaveTextToFile(sFileName, sRowContents, sErrorDescription)
						End If
						If (lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value)) Or (lAbsenceID <> CLng(oRecordset.Fields("AbsenceID").Value)) Then
							sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value)) & """)"
							sRowContents = sRowContents & TABLE_SEPARATOR & lPayrollNumber
							sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value)) & """)"
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("JourneyTypeID").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_01 />"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_04 />"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_07 />"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_08 />"
							dConcept01 = 0
							dConcept01a = 0
							dConcept04 = 0
							dConcept07 = 0
							dConcept08 = 0
							lConceptID = -2
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " "
								If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
									sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & " "
								End If
								sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_99 />"
							sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("AbsenceShortName").Value)) & """)"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<ABSENCE_AMOUNT />"
							sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & "<OCURRED_DATES />" & """)"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_50_70 />"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<DAYS />"
							If InStr(1, ",3,10,11,16,18,19,20,24,25,26,28,92,93,94,", "," & CStr(oRecordset.Fields("AbsenceID").Value) & ",", vbBinaryCompare) > 0 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_999_52 />"
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_999_71 />"
							End If

							If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
								adTotalAmount(0) = 0
								adTotalAmount(1) = 0
							End If
							lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
							lAbsenceID = CLng(oRecordset.Fields("AbsenceID").Value)
							dJourneyFactor = CDbl(oRecordset.Fields("JourneyFactor").Value)
							dShiftsWorkingHours = CDbl(oRecordset.Fields("ShiftsWorkingHours").Value)
							lJourneyTypeID = CInt(oRecordset.Fields("JourneyTypeID").Value)
							sDates = ""
							iDays = 0
							iTotalHours = 0
						End If
					End If
					If InStr(1, ",52,71,", "," & CStr(oRecordset.Fields("ConceptID").Value) & ",", vbBinaryCompare) > 0 Then
						lConceptID = CLng(oRecordset.Fields("ConceptID").Value)
						If ((lConceptID = 52) And (InStr(1, ",3,10,11,16,18,19,20,24,25,26,28,92,93,94,", "," & lAbsenceID & ",", vbBinaryCompare) > 0)) _
						Or ((lConceptID = 71) And (InStr(1, ",1,2,4,5,21,23,27,", "," & lAbsenceID & ",", vbBinaryCompare) > 0)) Then
							sRowContents = Replace(sRowContents, "<CONCEPT_50_70 />", FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True))
							If CLng(oRecordset.Fields("OcurredDate").Value) = CLng(oRecordset.Fields("EndDate").Value) Then
								sDates = sDates & CStr(oRecordset.Fields("OcurredDate").Value) & ", "
								iTotalHours = iTotalHours + 1
								iDays = iDays + 1
							Else
								If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
									lEndDate = lForPayrollID
								Else
									lEndDate = CLng(oRecordset.Fields("EndDate").Value)
								End If
								sDates = sDates & CStr(oRecordset.Fields("OcurredDate").Value) & "-" & lEndDate & ", "
								iTotalHours = iTotalHours + 1
								iDays = iDays + Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("OcurredDate").Value)), GetDateFromSerialNumber(lEndDate))) + 1
							End If
						End If
					Else
						Select Case CLng(oRecordset.Fields("ConceptID").Value)
							Case 1
								If InStr(1, "," & CStr(oRecordset.Fields("ConceptsIDs").Value) & ",", ",1,", vbBinaryCompare) > 0 Then dConcept01 = CDbl(oRecordset.Fields("ConceptAmount").Value)
								dConcept01a = CDbl(oRecordset.Fields("ConceptAmount").Value)
							Case 4
								If InStr(1, "," & CStr(oRecordset.Fields("ConceptsIDs").Value) & ",", ",4,", vbBinaryCompare) > 0 Then dConcept04 = CDbl(oRecordset.Fields("ConceptAmount").Value)
							Case 7
								If InStr(1, "," & CStr(oRecordset.Fields("ConceptsIDs").Value) & ",", ",7,", vbBinaryCompare) > 0 Then dConcept07 = CDbl(oRecordset.Fields("ConceptAmount").Value)
							Case 8
								If InStr(1, "," & CStr(oRecordset.Fields("ConceptsIDs").Value) & ",", ",8,", vbBinaryCompare) > 0 Then dConcept08 = CDbl(oRecordset.Fields("ConceptAmount").Value)
						End Select
						sRowContents = Replace(sRowContents, "<CONCEPT_" & Right(("00" & CStr(oRecordset.Fields("ConceptID").Value)), Len("00")) & " />", FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True))
					End If
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				If ((lConceptID = 52) And (InStr(1, ",3,10,11,16,18,19,20,24,25,26,28,92,93,94,", "," & lAbsenceID & ",", vbBinaryCompare) > 0)) _
				Or ((lConceptID = 71) And (InStr(1, ",1,2,4,5,21,23,27,", "," & lAbsenceID & ",", vbBinaryCompare) > 0)) Then
					sRowContents = Replace(sRowContents, "<CONCEPT_01 />", "0.00")
					sRowContents = Replace(sRowContents, "<CONCEPT_04 />", "0.00")
					sRowContents = Replace(sRowContents, "<CONCEPT_07 />", "0.00")
					sRowContents = Replace(sRowContents, "<CONCEPT_08 />", "0.00")
					adTotal = Split("0,0", ",")
					adTotal(0) = 0
					adTotal(1) = 0
					Select Case lAbsenceID
						Case 1
							Select Case lJourneyTypeID
								Case 1
									adTotal(0) = adTotal(0) + (Int(iTotalHours / 3) / dJourneyFactor / 4)
								Case 2, 3
									adTotal(0) = adTotal(0) + (Int(iTotalHours / 2) / dJourneyFactor / 4)
								Case 4
									adTotal(0) = adTotal(0) + (iTotalHours / dJourneyFactor / 4)
							End Select
						Case 2, 23, 27
							adTotal(0) = adTotal(0) + (iTotalHours / dJourneyFactor / 4)
						Case 4
							adTotal(0) = adTotal(0) + (iTotalHours / dJourneyFactor / 2)
						Case 3, 11, 18, 19, 20, 24, 25, 26, 28, 93, 94
							adTotal(0) = adTotal(0) + (iTotalHours / dJourneyFactor)
						Case 5
							adTotal(0) = adTotal(0) + (iTotalHours / dJourneyFactor / 6)
						Case 10
							If lJourneyTypeID <> 1 Then
								adTotal(0) = adTotal(0) + (iTotalHours / dJourneyFactor)
							Else
								adTotal(1) = adTotal(1) + (iTotalHours / dJourneyFactor)
							End If
						Case 16
							adTotal(0) = adTotal(0) + (iTotalHours / dJourneyFactor / 3)
						Case 21
							adTotal(0) = adTotal(0) + (iTotalHours / (dJourneyFactor * dShiftsWorkingHours * 2))
						Case 92
							adTotal(0) = adTotal(0) + (iTotalHours / dJourneyFactor * 2 / 5)
					End Select
					dAmount = 0
					If adTotal(1) > 0 Then 
						dAmount = ((dConcept01 + dConcept04) * adTotal(1) * 1.4) + ((dConcept07 + dConcept08) * adTotal(1))
					Else
						dAmount = (dConcept01 + dConcept04 + dConcept07 + dConcept08) * adTotal(0)
					End If
					sRowContents = Replace(sRowContents, "<ABSENCE_AMOUNT />", FormatNumber(dAmount, 2, True, False, True))
					If lConceptID = 52 Then
						adTotalAmount(0) = adTotalAmount(0) + dAmount
					Else
						adTotalAmount(1) = adTotalAmount(1) + dAmount
					End If
					sRowContents = Replace(sRowContents, "<CONCEPT_50_70 />", "0.00")
					sRowContents = Replace(sRowContents, "<CONCEPT_99 />", "0.00")
					If Len(sDates) > 0 Then sDates = Left(sDates, (Len(sDates) - Len(", ")))
					sRowContents = Replace(sRowContents, "<OCURRED_DATES />", sDates)
					sRowContents = Replace(sRowContents, "<DAYS />", iDays)
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = AppendTextToFile(sFileName, GetTableRowText(asRowContents, True, ""), sErrorDescription)
				End If
				sRowContents = GetFileContents(sFileName, sErrorDescription)
				If adTotalAmount(0) > (dConcept01a * 0.30) Then
					sRowContents = Replace(sRowContents, "<CONCEPT_999_52 />", FormatNumber((adTotalAmount(0) - (dConcept01a * 0.30)), 2, True, False, True))
				Else
					sRowContents = Replace(sRowContents, "<CONCEPT_999_52 />", "0.00")
				End If
				If adTotalAmount(1) > (dConcept01a * 0.30) Then
					sRowContents = Replace(sRowContents, "<CONCEPT_999_71 />", FormatNumber((adTotalAmount(1) - (dConcept01a * 0.30)), 2, True, False, True))
				Else
					sRowContents = Replace(sRowContents, "<CONCEPT_999_71 />", "0.00")
				End If
				lErrorNumber = SaveTextToFile(sFileName, sRowContents, sErrorDescription)
			lErrorNumber = AppendTextToFile(sFileName, "</TABLE>", sErrorDescription)

			lErrorNumber = ZipFolder(sFileName, Replace(sFileName, ".xls", ".zip"), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudieron guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				lErrorNumber = DeleteFile(sFileName, sErrorDescription)
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
	BuildReport1434 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1435(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the ConceptValues for Concepts
'Inputs:  oRequest, oADODBConnection, iSelectedTab, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1435"
	Dim sCondition
	Dim sCondition2
	Dim lPayrollID
	Dim lForPayrollID
	Dim lPayrollNumber
	Dim sDate
	Dim sStartDate
	Dim sEndDate
	Dim sFilePath
	Dim lReportID
	Dim sTemp
	Dim lCurrentID
	Dim dTotal
	Dim oRecordset
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asColumnsTitles
	Dim asCellWidths
	Dim asCellAlignments
	Dim sColumnsTitles
	Dim sCellWidths
	Dim sCellAlignments
	Dim lErrorNumber
	Dim sBoldBegin
	Dim sBoldEnd
	Dim sFontBegin
	Dim sFontEnd
	Dim bFirst
	Dim bContinue
	Dim sConceptIDs
	DIm iStatusID
	Dim sRecordIDs
	Dim sStartDateCondition

	Dim iLevelID
	Dim iAntiquityID
	Dim iAntiquityID2
	Dim iEconomicZoneID
	Dim iGroupGradeLevelID
	Dim iClassificationID
	Dim iIntegrationID
	Dim iPositionTypeID
	Dim sPositionTypeShortName
	Dim sPositionShortName
	Dim sLevelShortName
	Dim sWorkingHours
	Dim sPositionName
	Dim sGroupGradeLevelShortName
	Dim lCurrentPositionID

	Dim dConcept
	Dim lConcept_RecordID
	Dim lConcept_StartDate
	Dim bActiveConcept

	Dim dConcept_01
	Dim lConcept_01_RecordID
	Dim lConcept_01_StartDate
	Dim bActiveConcept_01

	Dim dConcept_03
	Dim lConcept_03_RecordID
	Dim lConcept_03_StartDate
	Dim bActiveConcept_03

	Dim dConcept_12
	Dim lConcept_12_RecordID
	Dim lConcept_12_StartDate
	Dim bActiveConcept_12

	Dim dConcept_35
	Dim lConcept_35_RecordID
	Dim lConcept_35_StartDate
	Dim bActiveConcept_35

	Dim dConcept_36
	Dim lConcept_36_RecordID
	Dim lConcept_36_StartDate
	Dim bActiveConcept_36

	Dim dConcept_48
	Dim lConcept_48_RecordID
	Dim lConcept_48_StartDate
	Dim bActiveConcept_48

	Dim dConcept_B2
	Dim lConcept_B2_RecordID
	Dim lConcept_B2_StartDate
	Dim bActiveConcept_B2

	Dim dConcept_Z3
	Dim lConcept_Z3_RecordID
	Dim lConcept_Z3_StartDate
	Dim bActiveConcept_Z3

	Dim dConcept_01_Z3
	Dim lConcept_01_Z3_RecordID
	Dim lConcept_01_Z3_StartDate
	Dim bActiveConcept_01_Z3

	Dim dConcept_03_Z3
	Dim lConcept_03_Z3_RecordID
	Dim lConcept_03_Z3_StartDate
	Dim bActiveConcept_03_Z3

	Dim dConcept_12_Z3
	Dim lConcept_12_Z3_RecordID
	Dim lConcept_12_Z3_StartDate
	Dim bActiveConcept_12_Z3

	Dim dConcept_35_Z3
	Dim lConcept_35_Z3_RecordID
	Dim lConcept_35_Z3_StartDate
	Dim bActiveConcept_35_Z3

	Dim dConcept_36_Z3
	Dim lConcept_36_Z3_RecordID
	Dim lConcept_36_Z3_StartDate
	Dim bActiveConcept_36_Z3

	Dim dConcept_48_Z3
	Dim lConcept_48_Z3_RecordID
	Dim lConcept_48_Z3_StartDate
	Dim bActiveConcept_48_Z3

	Dim dConcept_B2_Z3
	Dim lConcept_B2_Z3_RecordID
	Dim lConcept_B2_Z3_StartDate
	Dim bActiveConcept_B2_Z3

	sDate = Left(GetSerialNumberForDate(""), Len("00000000"))

	If aConceptComponent(N_STATUS_ID_CONCEPT) = 1 Then
		Call GetStartAndEndDatesFromURL("StartForValue", "EndForValue", "ConceptsValues.StartDate", False, sCondition)
		sStartDateCondition = sCondition
	End If
	If (Len(oRequest("PositionID").Item) > 0) And (aConceptComponent(N_STATUS_ID_CONCEPT) > 0) Then
		If CInt(oRequest("PositionID").Item) <> -1 Then
			sCondition = sCondition & " And (Positions.PositionID In (" & oRequest("PositionID").Item & "))"
		End If
	Else
		If (aConceptComponent(N_STATUS_ID_CONCEPT) = 1) And (Not bForExport) Then
				sCondition = sCondition & " And (Positions.PositionID In (0))"
		End If
	End If
	'If (CInt(oRequest("StartForValueYear").Item) > 0) And (CInt(oRequest("StartForValueMonth").Item) > 0) And (CInt(oRequest("StartForValueDay").Item) > 0) And (CInt(oRequest("EndForValueYear").Item) > 0) And (CInt(oRequest("EndForValueYearMonth").Item) > 0) And (CInt(oRequest("EndForValueYearDay").Item) > 0) Then Call GetStartAndEndDatesFromURL("StartForValue", "EndForValue", "ConceptsValues.StartDate", True, sCondition)
	'If (CInt(oRequest("StartJobEndYear").Item) > 0) And (CInt(oRequest("StartJobEndMonth").Item) > 0) And (CInt(oRequest("StartJobEndDay").Item) > 0) And (CInt(oRequest("EndJobEndYear").Item) > 0) And (CInt(oRequest("EndJobEndMonth").Item) > 0) And (CInt(oRequest("EndJobEndDay").Item) > 0) Then Call GetStartAndEndDatesFromURL("StartJobEnd", "EndJobEnd", "Jobs.EndDate", False, sCondition)
	sErrorDescription = "No se pudieron obtener los montos pagados."
	iStatusID = aConceptComponent(N_STATUS_ID_CONCEPT)
	sCondition = sCondition & " And (ConceptsValues.StatusID=" & iStatusID & ")"
	If Len(oRequest("ConceptID").Item) > 0 Then
		lErrorNumber = GetConcept(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
		sStartDate = aConceptComponent(N_START_DATE_CONCEPT)
		sEndDate = aConceptComponent(N_END_DATE_CONCEPT)
		Select Case iSelectedTab
			Case 0
				If InStr(1, ",1,38,49,", "," & CStr(oRequest("ConceptID").Item) & ",", vbBinaryCompare) > 0 Then
					sCondition = sCondition & " And (ConceptID IN (1, 38, 49))"
				Else
					sCondition = sCondition & " And (ConceptID=" & CStr(oRequest("ConceptID").Item) & ")"
				End If
			Case 1, 2, 3, 4
				If InStr(1, ",1,3,", "," & CStr(oRequest("ConceptID").Item) & ",", vbBinaryCompare) > 0 Then
					sCondition = sCondition & " And (ConceptID IN (1, 3))"
				Else
					sCondition = sCondition & " And (ConceptID=" & CStr(oRequest("ConceptID").Item) & ")"
				End If
			Case 5
				If InStr(1, ",39,89,", "," & CStr(oRequest("ConceptID").Item) & ",", vbBinaryCompare) > 0 Then
					sCondition = sCondition & " And (ConceptID IN (39, 89))"
				Else
					sCondition = sCondition & " And (ConceptID=" & CStr(oRequest("ConceptID").Item) & ")"
				End If
			Case 6
				If InStr(1, ",14,", "," & CStr(oRequest("ConceptID").Item) & ",", vbBinaryCompare) > 0 Then
					sCondition = sCondition & " And (ConceptID IN (14))"
				Else
					sCondition = sCondition & " And (ConceptID=" & CStr(oRequest("ConceptID").Item) & ")"
				End If
			Case Else
				sCondition = sCondition & " And (ConceptID=" & CStr(oRequest("ConceptID").Item) & ")"
		End Select
	Else
		sStartDate = sDate
		sEndDate = sDate
		Select Case iSelectedTab
			Case 0
				sCondition = sCondition & " And (ConceptID IN (1, 38, 49))"
			Case 1, 2, 3, 4
				sCondition = sCondition & " And (ConceptID IN (1, 3))"
			Case 5
				sCondition = sCondition & " And (ConceptID IN (39, 89))"
			Case 6
				sCondition = sCondition & " And (ConceptID IN (14))"
		End Select
	End If
	'sCondition = sCondition & " And (((ConceptsValues.StartDate>=" & sStartDate & ") And (ConceptsValues.StartDate<=" & sEndDate & ")) Or ((ConceptsValues.EndDate>=" & sStartDate & ") And (ConceptsValues.EndDate<=" & sEndDate & ")) Or ((ConceptsValues.EndDate>=" & sStartDate & ") And (ConceptsValues.StartDate<=" & sEndDate & ")))"
	If Len(oRequest("ConceptID").Item) > 0 Then	
		sCondition = sCondition & " And (((Positions.StartDate>="& sStartDate & ") And (Positions.StartDate<=" & sEndDate & ")) Or ((Positions.EndDate>=" & sStartDate & ") And (Positions.EndDate<=" & sEndDate & ")) Or ((Positions.EndDate>=" & sStartDate & ") And (Positions.StartDate<=" & sEndDate & ")))"
		sCondition = sCondition & " And (((PositionTypes.StartDate>=" & sStartDate & ") And (PositionTypes.StartDate<=" & sEndDate & ")) Or ((PositionTypes.EndDate>=" & sStartDate & ") And (PositionTypes.EndDate<=" & sEndDate & ")) Or ((PositionTypes.EndDate>=" & sStartDate & ") And (PositionTypes.StartDate<=" & sEndDate & ")))"
		sCondition = sCondition & " And (((GroupGradeLevels.StartDate>=" & sStartDate & ") And (GroupGradeLevels.StartDate<=" & sEndDate & ")) Or ((GroupGradeLevels.EndDate>=" & sStartDate & ") And (GroupGradeLevels.EndDate<=" & sEndDate & ")) Or ((GroupGradeLevels.EndDate>=" & sStartDate & ") And (GroupGradeLevels.StartDate<=" & sEndDate & ")))"
		sCondition = sCondition & " And (((Levels.StartDate>=" & sStartDate & ") And (Levels.StartDate<=" & sEndDate & ")) Or ((Levels.EndDate>=" & sStartDate & ") And (Levels.EndDate<=" & sEndDate & ")) Or ((Levels.EndDate>=" & sStartDate & ") And (Levels.StartDate<=" & sEndDate & ")))"
		sCondition = sCondition & " And (ConceptsValues.EmployeeTypeID IN (-1, " & iSelectedTab & "))" & " And (Positions.EmployeeTypeID IN (-1, " & iSelectedTab & "))"
	Else
		sCondition = sCondition & " And (((Positions.StartDate>=ConceptsValues.StartDate) And (Positions.StartDate<=ConceptsValues.EndDate)) Or ((Positions.EndDate>=ConceptsValues.StartDate) And (Positions.EndDate<=ConceptsValues.EndDate)) Or ((Positions.EndDate>=ConceptsValues.StartDate) And (Positions.StartDate<=ConceptsValues.EndDate)))"
		sCondition = sCondition & " And (((PositionTypes.StartDate>=ConceptsValues.StartDate) And (PositionTypes.StartDate<=ConceptsValues.EndDate)) Or ((PositionTypes.EndDate>=ConceptsValues.StartDate) And (PositionTypes.EndDate<=ConceptsValues.EndDate)) Or ((PositionTypes.EndDate>=ConceptsValues.StartDate) And (PositionTypes.StartDate<=ConceptsValues.EndDate)))"
		sCondition = sCondition & " And (((GroupGradeLevels.StartDate>=ConceptsValues.StartDate) And (GroupGradeLevels.StartDate<=ConceptsValues.EndDate)) Or ((GroupGradeLevels.EndDate>=ConceptsValues.StartDate) And (GroupGradeLevels.EndDate<=ConceptsValues.EndDate)) Or ((GroupGradeLevels.EndDate>=ConceptsValues.StartDate) And (GroupGradeLevels.StartDate<=ConceptsValues.EndDate)))"
		sCondition = sCondition & " And (((Levels.StartDate>=ConceptsValues.StartDate) And (Levels.StartDate<=ConceptsValues.EndDate)) Or ((Levels.EndDate>=ConceptsValues.StartDate) And (Levels.EndDate<=ConceptsValues.EndDate)) Or ((Levels.EndDate>=ConceptsValues.StartDate) And (Levels.StartDate<=ConceptsValues.EndDate)))"
		sCondition = sCondition & " And (ConceptsValues.EmployeeTypeID IN (-1, " & iSelectedTab & "))" & " And (Positions.EmployeeTypeID IN (-1, " & iSelectedTab & "))"
	End If
	'sCondition = sCondition & " And (ConceptsValues.PositionID=2)"
	aConceptComponent(S_QUERY_CONDITION_CONCEPT) = sCondition

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptsValues.RecordID, ConceptsValues.ConceptID, ConceptsValues.ConceptAmount, ConceptsValues.StartDate, ConceptsValues.EndDate, ConceptsValues.StatusID, Positions.PositionID, ConceptsValues.LevelID, Levels.LevelShortName, ConceptsValues.EconomicZoneID, ConceptsValues.ClassificationID, ConceptsValues.IntegrationID, ConceptsValues.GroupGradeLevelID, GroupGradeLevels.GroupGradeLevelShortName, ConceptsValues.WorkingHours, ConceptsValues.AntiquityID, ConceptsValues.Antiquity2ID, Positions.PositionShortName, Positions.PositionName, PositionTypes.PositionTypeID, PositionTypes.PositionTypeShortName, PositionTypes.PositionTypeName From ConceptsValues, Positions, PositionTypes, GroupGradeLevels, Levels Where (ConceptsValues.PositionID=Positions.PositionID) And (Positions.PositionTypeID=PositionTypes.PositionTypeID) And (ConceptsValues.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (ConceptsValues.LevelID=Levels.LevelID)" & sCondition & " Order by ConceptsValues.EmployeeTypeID, PositionShortName, StartDate, LevelID, ClassificationID, IntegrationID, GroupGradeLevelID, WorkingHours, ConceptsValues.PositionTypeID, EconomicZoneID", "ReportsQueries1400bLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select ConceptsValues.RecordID, ConceptsValues.ConceptID, ConceptsValues.ConceptAmount, ConceptsValues.StartDate, ConceptsValues.EndDate, ConceptsValues.StatusID, Positions.PositionID, ConceptsValues.LevelID, Levels.LevelShortName, ConceptsValues.EconomicZoneID, ConceptsValues.ClassificationID, ConceptsValues.IntegrationID, ConceptsValues.GroupGradeLevelID, GroupGradeLevels.GroupGradeLevelShortName, ConceptsValues.WorkingHours, ConceptsValues.AntiquityID, ConceptsValues.Antiquity2ID, Positions.PositionShortName, Positions.PositionName, PositionTypes.PositionTypeID, PositionTypes.PositionTypeShortName, PositionTypes.PositionTypeName From ConceptsValues, Positions, PositionTypes, GroupGradeLevels, Levels Where (ConceptsValues.PositionID=Positions.PositionID) And (Positions.PositionTypeID=PositionTypes.PositionTypeID) And (ConceptsValues.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (ConceptsValues.LevelID=Levels.LevelID)" & sCondition & " Order by ConceptsValues.EmployeeTypeID, PositionShortName, StartDate, LevelID, ClassificationID, IntegrationID, GroupGradeLevelID, WorkingHours, ConceptsValues.PositionTypeID, EconomicZoneID -->" & vbNewLine
	If iStatusID = 0 Then Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""sQuery"" ID=""sQueryHdn"" VALUE=""" & "Select ConceptsValues.RecordID, ConceptsValues.ConceptID, ConceptsValues.ConceptAmount, ConceptsValues.StartDate, ConceptsValues.EndDate, ConceptsValues.StatusID, Positions.PositionID, ConceptsValues.LevelID, Levels.LevelShortName, ConceptsValues.EconomicZoneID, ConceptsValues.ClassificationID, ConceptsValues.IntegrationID, ConceptsValues.GroupGradeLevelID, GroupGradeLevels.GroupGradeLevelShortName, ConceptsValues.WorkingHours, ConceptsValues.AntiquityID, ConceptsValues.Antiquity2ID, Positions.PositionShortName, Positions.PositionName, PositionTypes.PositionTypeID, PositionTypes.PositionTypeShortName, PositionTypes.PositionTypeName From ConceptsValues, Positions, PositionTypes, GroupGradeLevels, Levels Where (ConceptsValues.PositionID=Positions.PositionID) And (Positions.PositionTypeID=PositionTypes.PositionTypeID) And (ConceptsValues.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (ConceptsValues.LevelID=Levels.LevelID)" & sCondition & " Order by ConceptsValues.EmployeeTypeID, PositionShortName, StartDate, LevelID, ClassificationID, IntegrationID, GroupGradeLevelID, WorkingHours, ConceptsValues.PositionTypeID, EconomicZoneID" & """ />"
	sErrorDescription = "No se pudieron obtener los tabuladores."
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			bFirst = False
			bActiveConcept = False
			bActiveConcept_01 = False
			bActiveConcept_03 = False
			bActiveConcept_12 = False
			bActiveConcept_35 = False
			bActiveConcept_36 = False
			bActiveConcept_48 = False
			bActiveConcept_B2 = False
			bActiveConcept_Z3 = False
			bActiveConcept_01_Z3 = False
			bActiveConcept_03_Z3 = False
			bActiveConcept_12_Z3 = False
			bActiveConcept_35_Z3 = False
			bActiveConcept_36_Z3 = False
			bActiveConcept_48_Z3 = False
			bActiveConcept_B2_Z3 = False
			sRecordIDs = ""

			Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><TABLE BORDER="""
			If bForExport Then
				Response.Write "1"
			Else
				Response.Write "0"
			End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			If (Len(oRequest("ConceptID").Item) > 0) And (InStr(1, ",1,3,14,38,39,49,89,", "," & CStr(oRequest("ConceptID").Item) & ",", vbBinaryCompare) = 0) Then
				sColumnsTitles = "Tipo puesto,Código,Nivel,Jornada,Denominación del puesto,Fecha Inicio vigencia,Importe"
				sCellWidths = ",,,,,,,"
				sCellAlignments = "CENTER,RIGHT,RIGHT,RIGHT,LEFT,RIGHT,RIGHT"
			Else
				Select Case iSelectedTab
					Case 0
						asColumnsTitles = Split("<SPAN COLS=""6"">&nbsp;,<SPAN COLS=""4"">Zona 2,<SPAN COLS=""4"">Zona 3", ",", -1, vbBinaryCompare)
						asCellWidths = Split(",,,,,,,,,,,,,", ",", -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
						Else
							If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
								lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
							Else
								lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
							End If
						End If
						sColumnsTitles = "Tipo puesto,Código,Nivel,Jornada,Denominación del puesto,Fecha Inicio vigencia,Fecha Fin vigencia,Sueldo,Asignación médica,Gastos de actualización,Total,Sueldo,Asignación médica,Gastos de actualización,Total"
						sCellWidths = ",,,,,,,,,,,,,"
						sCellAlignments = "CENTER,RIGHT,RIGHT,RIGHT,LEFT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT"
					Case 1
						sColumnsTitles = "Zona Económica,Código,Denominación del puesto,Grupo grado nivel salarial,Clasificación,Integración,Fecha de inicio vigencia,Fecha Fin vigencia,Sueldo base,Compensación garantizada,Sueldo integrado"
						sCellWidths = ",,,,,,,,,,,,,"
						sCellAlignments = "CENTER,CENTER,LEFT,CENTER,CENTER,CENTER,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT"
					Case 2,4
						If iSelectedTab = 2 Then
							asColumnsTitles = Split("<SPAN COLS=""5"">&nbsp;,<SPAN COLS=""3"">Zona 2,<SPAN COLS=""3"">Zona 3", ",", -1, vbBinaryCompare)
							asCellWidths = Split(",,,,,,,,,,", ",", -1, vbBinaryCompare)
						Else
							asColumnsTitles = Split("<SPAN COLS=""4"">&nbsp;,<SPAN COLS=""3"">Zona 2,<SPAN COLS=""3"">Zona 3", ",", -1, vbBinaryCompare)
							asCellWidths = Split(",,,,,,,,,", ",", -1, vbBinaryCompare)
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
						If iSelectedTab = 2 Then
							sColumnsTitles = "Tipo de puesto,Código,Nivel,Denominación del puesto,Fecha Inicio vigencia,Fecha Fin vigencia,Sueldo,Compensación garantizada,Total,Sueldo,Compensación garantizada,Total"
							sCellWidths = ",,,,,,,,,,"
							sCellAlignments = "CENTER,CENTER,CENTER,,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT"
						Else
							sColumnsTitles = "Código,Nivel,Denominación del puesto,Fecha Inicio vigencia,Fecha Fin vigencia,Sueldo,Compensación garantizada,Total,Sueldo,Compensación garantizada,Total"
							sCellWidths = ",,,,,,,,,"
							sCellAlignments = "CENTER,CENTER,,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT"
						End If
					Case 3
						sColumnsTitles = "Código,Nivel,Denominación del puesto,Fecha Inicio vigencia,Fecha Fin vigencia,Sueldo base,Compensación garantizada,Total mensual bruto"
						sCellWidths = ",,,,,,,,,"
						sCellAlignments = "LEFT,CENTER,LEFT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT"
					Case 5
						asColumnsTitles = Split("<SPAN COLS=""3"">&nbsp;,<SPAN COLS=""3"">Zona 2,<SPAN COLS=""3"">Zona 3", ",", -1, vbBinaryCompare)
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
						sColumnsTitles = "Código,Denominación del puesto,Fecha Inicio vigencia,Fecha Fin vigencia,Beca,Complemento de beca,Total,Beca,Complemento de beca,Total"
						sCellWidths = ",,,,,,,,"
						sCellAlignments = ",,,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT"
					Case 6
						sColumnsTitles = "Código,&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Denominación del puesto&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;,Nivel,Zona económica,Fecha Inicio vigencia,Fecha Fin vigencia,Beca"
						sCellWidths = ",,,,,"
						sCellAlignments = "LEFT,LEFT,,CENTER,CENTER,RIGHT"
				End Select
			End If
			If (Not bForExport) And (iStatusID=0) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Or (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
				If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
					sColumnsTitles = sColumnsTitles & ",Acciones"
					sCellWidths = sCellWidths & ",80"
					sCellAlignments = sCellAlignments & ",CENTER"
				End If
			End If
			asColumnsTitles = Split(sColumnsTitles, ",", -1, vbBinaryCompare)
			asCellWidths = Split(sCellWidths, ",", -1, vbBinaryCompare)
			asCellAlignments = Split(sCellAlignments, ",", -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
			Else
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If
			End If
			lCurrentPositionID = -2
			dTotal = 0
			sBoldBegin = "<B>"
			sBoldEnd = "</B>"
			sFontBegin = ""
			sFontEnd = ""
			Do While Not oRecordset.EOF
				bContinue = False
				If (bFirst) And ((lCurrentPositionID <> CLng(oRecordset.Fields("PositionID").Value)) Or _
					(CLng(sStartDate) <> CLng(oRecordset.Fields("StartDate").Value)) Or _
					(CLng(iLevelID) <> CLng(oRecordset.Fields("LevelID").Value)) Or _
					(CLng(iClassificationID) <> CLng(oRecordset.Fields("ClassificationID").Value)) Or _
					(CLng(iIntegrationID) <> CLng(oRecordset.Fields("IntegrationID").Value)) Or _
					(CLng(iGroupGradeLevelID) <> CLng(oRecordset.Fields("GroupGradeLevelID").Value)) Or _
					(CSng(sWorkingHours) <> CSng(oRecordset.Fields("WorkingHours").Value)) Or _
					(CLng(iPositionTypeID) <> CLng(oRecordset.Fields("PositionTypeID").Value)) Or _
					(CInt(iAntiquityID) <> CInt(oRecordset.Fields("AntiquityID").Value)) Or _
					(CInt(iAntiquityID2) <> CInt(oRecordset.Fields("Antiquity2ID").Value))) _
				Then
					aConceptComponent(N_ID_CONCEPT) = CInt(oRecordset.Fields("ConceptID").Value)
					Select Case aConceptComponent(N_ID_CONCEPT)
						Case 1, 3, 14, 38, 39, 49, 89
							Select Case iSelectedTab
								Case 0
									sBoldBegin = ""
									sBoldEnd = ""
									If (StrComp(CStr(lConcept_01_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_35_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_48_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_01_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_35_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_48_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) _
									Then
										sBoldBegin = "<B>"
										sBoldEnd = "</B>"
									End If
									sFontBegin = ""
									sFontEnd = ""
									If (Not bActiveConcept_01) And (Not bActiveConcept_35) And (Not bActiveConcept_48) And (Not bActiveConcept_01_Z3) And (Not bActiveConcept_35_Z3) And (Not bActiveConcept_48_Z3) Then
										sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
										sFontEnd = "</FONT>"
									End If
									sRowContents =  sFontBegin & sBoldBegin & CleanStringForHTML(sPositionTypeShortName) & sBoldEnd & sFontEnd
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
									If iLevelID = -1 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & sBoldEnd & sFontEnd
									End If
									If CSng(sWorkingHours) = -1 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sWorkingHours) & " Hrs." & sBoldEnd & sFontEnd
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionName) & sBoldEnd & sFontEnd
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
									If CLng(sEndDate) = 30000000 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
									End If
									If bActiveConcept_01 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=0&RecordID=" & lConcept_01_RecordID & "&ConceptID=1&StartDate=" & lConcept_01_StartDate & """"
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_01), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_01), 2, True, False, True) & sBoldEnd & sFontEnd
										If Not IsEmpty(lConcept_01_RecordID) Then sRecordIDs = sRecordIDs & lConcept_01_RecordID & ","
									End If
									If bActiveConcept_35 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=0&RecordID=" & lConcept_35_RecordID & "&ConceptID=38&StartDate=" & lConcept_35_StartDate & """"
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_35), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_35), 2, True, False, True) & sBoldEnd & sFontEnd
										If Not IsEmpty(lConcept_35_RecordID) Then sRecordIDs = sRecordIDs & lConcept_35_RecordID & ","
									End If
									If bActiveConcept_48 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=0&RecordID=" & lConcept_48_RecordID & "&ConceptID=49&StartDate=" & lConcept_48_StartDate & """"
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_48), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_48), 2, True, False, True) & sBoldEnd & sFontEnd
										If Not IsEmpty(lConcept_48_RecordID) Then sRecordIDs = sRecordIDs & lConcept_48_RecordID & ","
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(dConcept_01) + CDbl(dConcept_35) + CDbl(dConcept_48), 2, True, False, True)
									If bActiveConcept_01_Z3 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=0&RecordID=" & lConcept_01_Z3_RecordID & "&ConceptID=1&StartDate=" & lConcept_01_Z3_StartDate & """"
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_01_Z3), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_01_Z3), 2, True, False, True) & sBoldEnd & sFontEnd
										If Not IsEmpty(lConcept_01_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_01_Z3_RecordID & ","
									End If
									If bActiveConcept_35_Z3 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=0&RecordID=" & lConcept_35_Z3_RecordID & "&ConceptID=38&StartDate=" & lConcept_35_Z3_StartDate & """"
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_35_Z3), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_35_Z3), 2, True, False, True) & sBoldEnd & sFontEnd
										If Not IsEmpty(lConcept_35_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_35_Z3_RecordID & ","
									End If
									If bActiveConcept_48_Z3 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=0&RecordID=" & lConcept_48_Z3_RecordID & "&ConceptID=49&StartDate=" & lConcept_48_Z3_StartDate & """"
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_48_Z3), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_48_Z3), 2, True, False, True) & sBoldEnd & sFontEnd
										If Not IsEmpty(lConcept_48_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_48_Z3_RecordID & ","
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(dConcept_01_Z3) + CDbl(dConcept_35_Z3) + CDbl(dConcept_48_Z3), 2, True, False, True)
									If (Not bActiveConcept_01) And (Not bActiveConcept_35) And (Not bActiveConcept_48) And (Not bActiveConcept_01_Z3) And (Not bActiveConcept_35_Z3) And (Not bActiveConcept_48_Z3) Then
										If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
											If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
											sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
											If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
												sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
													sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
												sRowContents = sRowContents & "</A>&nbsp;"
												sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""RecordIDChk"" ID=""RecordIDChk"" Value=""" & sRecordIDs & """ CHECKED=""1"" />"
											End If
										End If
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR
									asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If (dConcept_01 + dConcept_35 + dConcept_48 + lConcept_01_Z3_RecordID + dConcept_35_Z3 + dConcept_48_Z3) > 0 Then
										If bForExport Then
											lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
										Else
											lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
										End If
									End If
								Case 1
									sBoldBegin = ""
									sBoldEnd = ""
									If (StrComp(CStr(lConcept_01_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_03_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) _
									Then
										sBoldBegin = "<B>"
										sBoldEnd = "</B>"
									End If
									sFontBegin = ""
									sFontEnd = ""
									If (Not bActiveConcept_01) And (Not bActiveConcept_03) Then
										sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
										sFontEnd = "</FONT>"
									End If
									If iEconomicZoneID = 0 Then
										sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
									Else
										sRowContents = sFontBegin & sBoldBegin & CStr(iEconomicZoneID) & sBoldEnd & sFontEnd
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionName) & sBoldEnd & sFontEnd
									If iGroupGradeLevelID = -1 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sGroupGradeLevelShortName) & sBoldEnd & sFontEnd
									End If
									If iClassificationID = -1 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(iClassificationID) & sBoldEnd & sFontEnd
									End If
									If iIntegrationID = -1 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(iIntegrationID) & sBoldEnd & sFontEnd
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
									If CLng(sEndDate) = 30000000 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
									End If
									If bActiveConcept_01 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=1&RecordID=" & lConcept_01_RecordID & "&ConceptID=1&StartDate=" & lConcept_01_StartDate & """"
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_01), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_01), 2, True, False, True) & sBoldEnd & sFontEnd
										If Not IsEmpty(lConcept_01_RecordID) Then sRecordIDs = sRecordIDs & lConcept_01_RecordID & ","
									End If
									If bActiveConcept_03 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=1&RecordID=" & lConcept_03_RecordID & "&ConceptID=3&StartDate=" & lConcept_03_StartDate & """"
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_03), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_03), 2, True, False, True) & sBoldEnd & sFontEnd
										If Not IsEmpty(lConcept_03_RecordID) Then sRecordIDs = sRecordIDs & lConcept_03_RecordID & ","
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(dConcept_01) + CDbl(dConcept_03), 2, True, False, True)
									If (Not bActiveConcept_01) And (Not bActiveConcept_03) Then
										If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
											If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
											sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
											If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
												sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
													sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
												sRowContents = sRowContents & "</A>&nbsp;"
												sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""RecordIDChk"" ID=""RecordIDChk"" Value=""" & sRecordIDs & """ CHECKED=""1"" />"
											End If
										End If
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR
									asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If (dConcept_01 + dConcept_03) > 0 Then
										If bForExport Then
											lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
										Else
											lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
										End If
									End If
								Case 2,4
									sBoldBegin = ""
									sBoldEnd = ""
									If (StrComp(CStr(lConcept_01_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_03_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_01_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_03_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) _
									Then
										sBoldBegin = "<B>"
										sBoldEnd = "</B>"
									End If
									sFontBegin = ""
									sFontEnd = ""
									If (Not bActiveConcept_01) And (Not bActiveConcept_03) And (Not bActiveConcept_01_Z3) And (Not bActiveConcept_03_Z3) Then
										sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
										sFontEnd = "</FONT>"
									End If
									If iSelectedTab = 2 Then
										sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(sPositionTypeShortName) & sBoldEnd & sFontEnd
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
									Else
										sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
									End If
									If iLevelID = -1 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & sBoldEnd & sFontEnd
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionName) & sBoldEnd & sFontEnd
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
									If CLng(sEndDate) = 30000000 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
									End If
									If bActiveConcept_01 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_01_RecordID & "&ConceptID=1&StartDate=" & lConcept_01_StartDate & """"
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_01), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_01), 2, True, False, True) & sBoldEnd & sFontEnd
										If Not IsEmpty(lConcept_01_RecordID) Then sRecordIDs = sRecordIDs & lConcept_01_RecordID & ","
									End If
									If bActiveConcept_03 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_03_RecordID & "&ConceptID=3&StartDate=" & lConcept_03_StartDate & """"
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_03), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_03), 2, True, False, True) & sBoldEnd & sFontEnd
										If Not IsEmpty(lConcept_03_RecordID) Then sRecordIDs = sRecordIDs & lConcept_03_RecordID & ","
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(dConcept_01) + CDbl(dConcept_03), 2, True, False, True)
									If bActiveConcept_01_Z3 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_01_Z3_RecordID & "&ConceptID=1&StartDate=" & lConcept_01_Z3_StartDate & """"
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_01_Z3), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_01_Z3), 2, True, False, True) & sBoldEnd & sFontEnd
										If Not IsEmpty(lConcept_01_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_01_Z3_RecordID & ","
									End If
									If bActiveConcept_03_Z3 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_03_Z3_RecordID & "&ConceptID=3&StartDate=" & lConcept_03_Z3_StartDate & """"
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_03_Z3), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_03_Z3), 2, True, False, True) & sBoldEnd & sFontEnd
										If Not IsEmpty(lConcept_03_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_03_Z3_RecordID & ","
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(dConcept_01_Z3) + CDbl(dConcept_03_Z3), 2, True, False, True)
									If (Not bActiveConcept_01) And (Not bActiveConcept_03) And (Not bActiveConcept_01_Z3) And (Not bActiveConcept_03_Z3) Then
										If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
											If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
											sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
											If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
												sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
													sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
												sRowContents = sRowContents & "</A>&nbsp;"
												sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""RecordIDChk"" ID=""RecordIDChk"" Value=""" & sRecordIDs & """ CHECKED=""1"" />"
											End If
										End If
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR
									asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If (dConcept_01 + dConcept_03 + dConcept_01_Z3 + dConcept_03_Z3) > 0 Then
										If bForExport Then
											lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
										Else
											lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
										End If
									End If
								Case 3
									sBoldBegin = ""
									sBoldEnd = ""
									If (StrComp(CStr(lConcept_01_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_03_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) _
									Then
										sBoldBegin = "<B>"
										sBoldEnd = "</B>"
									End If
									sFontBegin = ""
									sFontEnd = ""
									If (Not bActiveConcept_01) And (Not bActiveConcept_03) Then
										sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
										sFontEnd = "</FONT>"
									End If
									sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
									If iLevelID = -1 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & sBoldEnd & sFontEnd
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionName) & sBoldEnd & sFontEnd
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
									If CLng(sEndDate) = 30000000 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
									End If
									If bActiveConcept_01 Then
										sFontBegin = ""
										sFontEnd = ""
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=3&RecordID=" & lConcept_01_RecordID & "&ConceptID=1&StartDate=" & lConcept_01_StartDate & """"
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_01), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_01), 2, True, False, True) & sBoldEnd & sFontEnd
										If Not IsEmpty(lConcept_01_RecordID) Then sRecordIDs = sRecordIDs & lConcept_01_RecordID & ","
									End If
									If bActiveConcept_03 Then
										sFontBegin = ""
										sFontEnd = ""
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=3&RecordID=" & lConcept_03_RecordID & "&ConceptID=3&StartDate=" & lConcept_03_StartDate & """"
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_03), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_03), 2, True, False, True) & sBoldEnd & sFontEnd
										If Not IsEmpty(lConcept_03_RecordID) Then sRecordIDs = sRecordIDs & lConcept_03_RecordID & ","
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(dConcept_01) + CDbl(dConcept_03), 2, True, False, True)
									If (Not bActiveConcept_01) And (Not bActiveConcept_03) Then
										If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
											If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
											sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
											If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
												sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
													sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
												sRowContents = sRowContents & "</A>&nbsp;"
												sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""RecordIDChk"" ID=""RecordIDChk"" Value=""" & sRecordIDs & """ CHECKED=""1"" />"
											End If
										End If
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR
									asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If (dConcept_01 + dConcept_03) > 0 Then
										If bForExport Then
											lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
										Else
											lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
										End If
									End If
								Case 5
									sBoldBegin = ""
									sBoldEnd = ""
									If (StrComp(CStr(lConcept_B2_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_36_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_B2_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_36_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) _
									Then
										sBoldBegin = "<B>"
										sBoldEnd = "</B>"
									End If
									sFontBegin = ""
									sFontEnd = ""
									If (Not bActiveConcept_B2) And (Not bActiveConcept_36) And (Not bActiveConcept_B2_Z3) And (Not bActiveConcept_36_Z3) Then
										sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
										sFontEnd = "</FONT>"
									End If
									sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionName) & sBoldEnd & sFontEnd
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
									If CLng(sEndDate) = 30000000 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
									End If
									If bActiveConcept_B2 Then
										sFontBegin = ""
										sFontEnd = ""
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=5&RecordID=" & lConcept_B2_RecordID & "&ConceptID=89&StartDate=" & lConcept_B2_StartDate & """"
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_B2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_B2), 2, True, False, True) & sBoldEnd & sFontEnd
										If Not IsEmpty(lConcept_B2_RecordID) Then sRecordIDs = sRecordIDs & lConcept_B2_RecordID & ","
									End If
									If bActiveConcept_36 Then
										sFontBegin = ""
										sFontEnd = ""
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=5&RecordID=" & lConcept_36_RecordID & "&ConceptID=39&StartDate=" & lConcept_36_StartDate & """"
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_36), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_36), 2, True, False, True) & sBoldEnd & sFontEnd
										If Not IsEmpty(lConcept_36_RecordID) Then sRecordIDs = sRecordIDs & lConcept_36_RecordID & ","
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(dConcept_B2) + CDbl(dConcept_36), 2, True, False, True)
									If bActiveConcept_B2_Z3 Then
										sFontBegin = ""
										sFontEnd = ""
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=5&RecordID=" & lConcept_B2_Z3_RecordID & "&ConceptID=89&StartDate=" & lConcept_B2_StartDate & """"
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_B2_Z3), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_B2), 2, True, False, True) & sBoldEnd & sFontEnd
										If Not IsEmpty(lConcept_B2_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_B2_Z3_RecordID & ","
									End If
									If bActiveConcept_36_Z3 Then
										sFontBegin = ""
										sFontEnd = ""
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=5&RecordID=" & lConcept_36_Z3_RecordID & "&ConceptID=39&StartDate=" & lConcept_36_StartDate & """"
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_36_Z3), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_36), 2, True, False, True) & sBoldEnd & sFontEnd
										If Not IsEmpty(lConcept_36_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_36_Z3_RecordID & ","
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(dConcept_B2_Z3) + CDbl(dConcept_36_Z3), 2, True, False, True)
									If (Not bActiveConcept_B2) And (Not bActiveConcept_36) And (Not bActiveConcept_B2_Z3) And (Not bActiveConcept_36_Z3) Then
										If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
											If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
											sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
											If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
												sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
													sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
												sRowContents = sRowContents & "</A>&nbsp;"
												sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""RecordIDChk"" ID=""RecordIDChk"" Value=""" & sRecordIDs & """ CHECKED=""1"" />"
											End If
										End If
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR
									asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If (dConcept_B2 + dConcept_36) > 0 Then
										If bForExport Then
											lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
										Else
											lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
										End If
									End If
								Case 6
									sBoldBegin = ""
									sBoldEnd = ""
									If (StrComp(CStr(lConcept_12_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Then
										sBoldBegin = "<B>"
										sBoldEnd = "</B>"
									End If
									sFontBegin = ""
									sFontEnd = ""
									If (Not bActiveConcept_12) Then
										sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
										sFontEnd = "</FONT>"
									End If
									sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionName) & sBoldEnd & sFontEnd
									If iLevelID = -1 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & sBoldEnd & sFontEnd
									End If
									If iEconomicZoneID = 0 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(iEconomicZoneID) & sBoldEnd & sFontEnd
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
									If CLng(sEndDate) = 30000000 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
									End If
									If bActiveConcept_12 Then
										sFontBegin = ""
										sFontEnd = ""
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=6&RecordID=" & lConcept_12_RecordID & "&ConceptID=14&StartDate=" & lConcept_12_StartDate & """"
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_12), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_12), 2, True, False, True) & sBoldEnd & sFontEnd
										If Not IsEmpty(lConcept_12_RecordID) Then sRecordIDs = sRecordIDs & lConcept_12_RecordID & ","
									End If
									If Not bActiveConcept_12 Then
										If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
											If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
											sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
											If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
												sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
													sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
												sRowContents = sRowContents & "</A>&nbsp;"
												sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""RecordIDChk"" ID=""RecordIDChk"" Value=""" & sRecordIDs & """ CHECKED=""1"" />"
											End If
										End If
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR
									asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If dConcept_12 > 0 Then
										If bForExport Then
											lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
										Else
											lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
										End If
									End If
							End Select
						Case Else
							sBoldBegin = ""
							sBoldEnd = ""
							If (StrComp(CStr(lConcept_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Then
								sBoldBegin = "<B>"
								sBoldEnd = "</B>"
							End If
							sFontBegin = ""
							sFontEnd = ""
							If (Not bActiveConcept) Then
								sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
								sFontEnd = "</FONT>"
							End If
							If CInt(oRecordset.Fields("PositionTypeID").Value) = -1 Then
								sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
							Else
								sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PositionTypeShortName").Value)) & sBoldEnd & sFontEnd
							End If
							If CInt(oRecordset.Fields("PositionID").Value) = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(oRecordset.Fields("PositionShortName").Value) & sBoldEnd & sFontEnd
							End If
							If CInt(oRecordset.Fields("LevelID").Value) = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(oRecordset.Fields("LevelShortName").Value) & sBoldEnd & sFontEnd
							End If
							If CSng(oRecordset.Fields("WorkingHours").Value) = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("WorkingHours").Value)) & " Hrs." & sBoldEnd & sFontEnd
							End If
							If CInt(oRecordset.Fields("PositionID").Value) = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(oRecordset.Fields("PositionName").Value) & sBoldEnd & sFontEnd
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
							If CLng(sEndDate) = 30000000 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
							End If
							If bActiveConcept Then
								sFontBegin = ""
								sFontEnd = ""
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_RecordID & "&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&StartDate=" & lConcept_StartDate & """"
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_RecordID) Then sRecordIDs = sRecordIDs & lConcept_RecordID & ","
							End If
							If Not bActiveConcept Then
								If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
									If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
									sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
									If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
										sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
											sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
										sRowContents = sRowContents & "</A>&nbsp;"
										sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""RecordIDChk"" ID=""RecordIDChk"" Value=""" & sRecordIDs & """ CHECKED=""1"" />"
									End If
								End If
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If dConcept > 0 Then
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
							End If
					End Select
					dConcept = 0
					dConcept_01 = 0
					dConcept_03 = 0
					dConcept_12 = 0
					dConcept_35 = 0
					dConcept_36 = 0
					dConcept_48 = 0
					dConcept_B2 = 0
					dConcept_Z3 = 0
					dConcept_01_Z3 = 0
					dConcept_03_Z3 = 0
					dConcept_12_Z3 = 0
					dConcept_35_Z3 = 0
					dConcept_36_Z3 = 0
					dConcept_48_Z3 = 0
					dConcept_B2_Z3 = 0
					lConcept_RecordID=0
					lConcept_01_RecordID=0
					lConcept_03_RecordID=0
					lConcept_12_RecordID=0
					lConcept_35_RecordID=0
					lConcept_36_RecordID=0
					lConcept_48_RecordID=0
					lConcept_B2_RecordID=0
					lConcept_Z3_RecordID=0
					lConcept_01_Z3_RecordID=0
					lConcept_03_Z3_RecordID=0
					lConcept_12_Z3_RecordID=0
					lConcept_35_Z3_RecordID=0
					lConcept_36_Z3_RecordID=0
					lConcept_48_Z3_RecordID=0
					lConcept_B2_Z3_RecordID=0
					bActiveConcept = False
					bActiveConcept_01 = False
					bActiveConcept_03 = False
					bActiveConcept_12 = False
					bActiveConcept_35 = False
					bActiveConcept_36 = False
					bActiveConcept_48 = False
					bActiveConcept_B2 = False
					bActiveConcept_Z3 = False
					bActiveConcept_01_Z3 = False
					bActiveConcept_03_Z3 = False
					bActiveConcept_12_Z3 = False
					bActiveConcept_35_Z3 = False
					bActiveConcept_36_Z3 = False
					bActiveConcept_48_Z3 = False
					bActiveConcept_B2_Z3 = False
					sRecordIDs = ""
				End If
				bFirst = True
				lCurrentPositionID = CLng(oRecordset.Fields("PositionID").Value)
				If CInt(oRecordset.Fields("EconomicZoneID").Value) = 3 Then
					Select Case CLng(oRecordset.Fields("ConceptID").Value)
						Case 1
							If Not bContinue Then
								If dConcept_01_Z3 > 0 Then
									bContinue = False
								Else
									dConcept_01_Z3 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_01_Z3_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_01_Z3_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_01_Z3 = True
									Else
										bActiveConcept_01_Z3 = False
									End If
								End If
							End If
						Case 3
							If Not bContinue Then
								If dConcept_03_Z3 > 0 Then
									bContinue = False
								Else
									dConcept_03_Z3 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_03_Z3_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_03_Z3_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_03_Z3 = True
									Else
										bActiveConcept_03_Z3 = False
									End If
								End If
							End If
						Case 14
							If Not bContinue Then
								If dConcept_12_Z3 > 0 Then
									bContinue = False
								Else
									dConcept_12_Z3 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_12_Z3_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_12_Z3_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_12_Z3 = True
									Else
										bActiveConcept_12_Z3 = False
									End If
								End If
							End If
						Case 38
							If Not bContinue Then
								If dConcept_35_Z3 > 0 Then
									bContinue = False
								Else
									dConcept_35_Z3 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_35_Z3_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_35_Z3_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_35_Z3 = True
									Else
										bActiveConcept_35_Z3 = False
									End If
								End If
							End If
						Case 39
							If Not bContinue Then
								If dConcept_36_Z3 > 0 Then
									bContinue = False
								Else
									dConcept_36_Z3 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_36_Z3_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_36_Z3_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_36_Z3 = True
									Else
										bActiveConcept_36_Z3 = False
									End If
								End If
							End If
						Case 49
							If Not bContinue Then
								If dConcept_48_Z3 > 0 Then
									bContinue = False
								Else
									dConcept_48_Z3 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_48_Z3_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_48_Z3_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_48_Z3 = True
									Else
										bActiveConcept_48_Z3 = False
									End If
								End If
							End If
						Case 89
							If Not bContinue Then
								If dConcept_B2_Z3 > 0 Then
									bContinue = False
								Else
									dConcept_B2_Z3 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_B2_Z3_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_B2_Z3_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_B2_Z3 = True
									Else
										bActiveConcept_B2_Z3 = False
									End If
								End If
							End If
						Case Else
							If Not bContinue Then
								If dConcept_Z3 > 0 Then
									bContinue = False
								Else
									dConcept_Z3 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_Z3_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_Z3_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_Z3 = True
									Else
										bActiveConcept_Z3 = False
									End If
								End If
							End If
					End Select
				Else
					Select Case CLng(oRecordset.Fields("ConceptID").Value)
						Case 1
							If Not bContinue Then
								If dConcept_01 > 0 Then
									bContinue = False
								Else
									dConcept_01 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_01_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_01_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_01 = True
									Else
										bActiveConcept_01 = False
									End If
								End If
							End If
						Case 3
							If Not bContinue Then
								If dConcept_03 > 0 Then
									bContinue = False
								Else
									dConcept_03 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_03_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_03_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_03 = True
									Else
										bActiveConcept_03 = False
									End If
								End If
							End If
						Case 14
							If Not bContinue Then
								If dConcept_12 > 0 Then
									bContinue = False
								Else
									dConcept_12 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_12_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_12_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_12 = True
									Else
										lConcept_12_RecordID = CLng(oRecordset.Fields("RecordID").Value)
										bActiveConcept_12 = False
									End If
								End If
							End If
						Case 38
							If Not bContinue Then
								If dConcept_35 > 0 Then
									bContinue = False
								Else
									dConcept_35 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_35_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_35_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_35 = True
									Else
										bActiveConcept_35 = False
									End If
								End If
							End If
						Case 39
							If Not bContinue Then
								If dConcept_36 > 0 Then
									bContinue = False
								Else
									dConcept_36 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_36_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_36_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_36 = True
									Else
										bActiveConcept_36 = False
									End If
								End If
							End If
						Case 49
							If Not bContinue Then
								If dConcept_48 > 0 Then
									bContinue = False
								Else
									dConcept_48 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_48_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_48_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_48 = True
									Else
										bActiveConcept_48 = False
									End If
								End If
							End If
						Case 89
							If Not bContinue Then
								If dConcept_B2 > 0 Then
									bContinue = False
								Else
									dConcept_B2 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_B2_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_B2_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_B2 = True
									Else
										bActiveConcept_B2 = False
									End If
								End If
							End If
						Case Else
							If Not bContinue Then
								If dConcept > 0 Then
									bContinue = False
								Else
									dConcept = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept = True
									Else
										bActiveConcept = False
									End If
								End If
							End If
					End Select
				End If
				aConceptComponent(N_ID_CONCEPT) = CInt(oRecordset.Fields("ConceptID").Value)
				iLevelID = CInt(oRecordset.Fields("LevelID").Value)
				iAntiquityID = CInt(oRecordset.Fields("AntiquityID").Value)
				iAntiquityID2 = CInt(oRecordset.Fields("Antiquity2ID").Value)
				iEconomicZoneID = CInt(oRecordset.Fields("EconomicZoneID").Value)
				iGroupGradeLevelID = CInt(oRecordset.Fields("GroupGradeLevelID").Value)
				iClassificationID = CInt(oRecordset.Fields("ClassificationID").Value)
				iIntegrationID = CInt(oRecordset.Fields("IntegrationID").Value)
				iPositionTypeID = CInt(oRecordset.Fields("PositionTypeID").Value)
				sPositionTypeShortName = CStr(oRecordset.Fields("PositionTypeShortName").Value)
				sPositionShortName = CStr(oRecordset.Fields("PositionShortName").Value)
				sLevelShortName = CStr(oRecordset.Fields("LevelShortName").Value)
				sWorkingHours = CStr(oRecordset.Fields("WorkingHours").Value)
				sPositionName = CStr(oRecordset.Fields("PositionName").Value)
				sStartDate = CStr(oRecordset.Fields("StartDate").Value)
				sEndDate = CStr(oRecordset.Fields("EndDate").Value)
				sGroupGradeLevelShortName = CStr(oRecordset.Fields("GroupGradeLevelShortName").Value)
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			Select Case aConceptComponent(N_ID_CONCEPT)
				Case 1, 3, 14, 38, 39, 49, 89
					Select Case iSelectedTab
						Case 0
							sBoldBegin = ""
							sBoldEnd = ""
							If (StrComp(CStr(lConcept_01_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_35_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_48_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_01_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_35_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_48_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) _
							Then
								sBoldBegin = "<B>"
								sBoldEnd = "</B>"
							End If
							sFontBegin = ""
							sFontEnd = ""
							If (Not bActiveConcept_01) And (Not bActiveConcept_35) And (Not bActiveConcept_48) And (Not bActiveConcept_01_Z3) And (Not bActiveConcept_35_Z3) And (Not bActiveConcept_48_Z3) Then
								sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
								sFontEnd = "</FONT>"
							End If
							sRowContents =  sFontBegin & sBoldBegin & CleanStringForHTML(sPositionTypeShortName) & sBoldEnd & sFontEnd
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
							If iLevelID = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & sBoldEnd & sFontEnd
							End If
							If CSng(sWorkingHours) = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sWorkingHours) & " Hrs." & sBoldEnd & sFontEnd
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionName) & sBoldEnd & sFontEnd
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
							If CLng(sEndDate) = 30000000 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
							End If
							If bActiveConcept_01 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=0&RecordID=" & lConcept_01_RecordID & "&ConceptID=1&StartDate=" & lConcept_01_StartDate & """"
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_01), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_01), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_01_RecordID) Then sRecordIDs = sRecordIDs & lConcept_01_RecordID & ","
							End If
							If bActiveConcept_35 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=0&RecordID=" & lConcept_35_RecordID & "&ConceptID=38&StartDate=" & lConcept_35_StartDate & """"
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_35), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_35), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_35_RecordID) Then sRecordIDs = sRecordIDs & lConcept_35_RecordID & ","
							End If
							If bActiveConcept_48 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=0&RecordID=" & lConcept_48_RecordID & "&ConceptID=49&StartDate=" & lConcept_48_StartDate & """"
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_48), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_48), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_48_RecordID) Then sRecordIDs = sRecordIDs & lConcept_48_RecordID & ","
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(dConcept_01) + CDbl(dConcept_35) + CDbl(dConcept_48), 2, True, False, True)
							If bActiveConcept_01_Z3 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=0&RecordID=" & lConcept_01_Z3_RecordID & "&ConceptID=1&StartDate=" & lConcept_01_Z3_StartDate & """"
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_01_Z3), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_01_Z3), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_01_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_01_Z3_RecordID & ","
							End If
							If bActiveConcept_35_Z3 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=0&RecordID=" & lConcept_35_Z3_RecordID & "&ConceptID=38&StartDate=" & lConcept_35_Z3_StartDate & """"
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_35_Z3), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_35_Z3), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_35_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_35_Z3_RecordID & ","
							End If
							If bActiveConcept_48_Z3 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=0&RecordID=" & lConcept_48_Z3_RecordID & "&ConceptID=49&StartDate=" & lConcept_48_Z3_StartDate & """"
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_48_Z3), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_48_Z3), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_48_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_48_Z3_RecordID & ","
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(dConcept_01_Z3) + CDbl(dConcept_35_Z3) + CDbl(dConcept_48_Z3), 2, True, False, True)
							If (Not bActiveConcept_01) And (Not bActiveConcept_35) And (Not bActiveConcept_48) And (Not bActiveConcept_01_Z3) And (Not bActiveConcept_35_Z3) And (Not bActiveConcept_48_Z3) Then
								If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
									If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
									sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
									If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
										sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
											sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
										sRowContents = sRowContents & "</A>&nbsp;"
										sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""RecordIDChk"" ID=""RecordIDChk"" Value=""" & sRecordIDs & """ CHECKED=""1"" />"
									End If
								End If
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If (dConcept_01 + dConcept_35 + dConcept_48 + lConcept_01_Z3_RecordID + dConcept_35_Z3 + dConcept_48_Z3) > 0 Then
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
							End If
						Case 1
							sBoldBegin = ""
							sBoldEnd = ""
							If (StrComp(CStr(lConcept_01_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_03_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) _
							Then
								sBoldBegin = "<B>"
								sBoldEnd = "</B>"
							End If
							sFontBegin = ""
							sFontEnd = ""
							If (Not bActiveConcept_01) And (Not bActiveConcept_03) Then
								sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
								sFontEnd = "</FONT>"
							End If
							If iEconomicZoneID = 0 Then
								sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
							Else
								sRowContents = sFontBegin & sBoldBegin & CStr(iEconomicZoneID) & sBoldEnd & sFontEnd
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionName) & sBoldEnd & sFontEnd
							If iGroupGradeLevelID = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sGroupGradeLevelShortName) & sBoldEnd & sFontEnd
							End If
							If iClassificationID = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(iClassificationID) & sBoldEnd & sFontEnd
							End If
							If iIntegrationID = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(iIntegrationID) & sBoldEnd & sFontEnd
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
							If CLng(sEndDate) = 30000000 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
							End If
							If bActiveConcept_01 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=1&RecordID=" & lConcept_01_RecordID & "&ConceptID=1&StartDate=" & lConcept_01_StartDate & """"
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_01), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_01), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_01_RecordID) Then sRecordIDs = sRecordIDs & lConcept_01_RecordID & ","
							End If
							If bActiveConcept_03 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=1&RecordID=" & lConcept_03_RecordID & "&ConceptID=3&StartDate=" & lConcept_03_StartDate & """"
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_03), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_03), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_03_RecordID) Then sRecordIDs = sRecordIDs & lConcept_03_RecordID & ","
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(dConcept_01) + CDbl(dConcept_03), 2, True, False, True)
							If (Not bActiveConcept_01) And (Not bActiveConcept_03) Then
								If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
									If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
									sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
									If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
										sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
											sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
										sRowContents = sRowContents & "</A>&nbsp;"
										sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""RecordIDChk"" ID=""RecordIDChk"" Value=""" & sRecordIDs & """ CHECKED=""1"" />"
									End If
								End If
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If (dConcept_01 + dConcept_03) > 0 Then
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
							End If
						Case 2,4
							sBoldBegin = ""
							sBoldEnd = ""
							If (StrComp(CStr(lConcept_01_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_03_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_01_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_03_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) _
							Then
								sBoldBegin = "<B>"
								sBoldEnd = "</B>"
							End If
							sFontBegin = ""
							sFontEnd = ""
							If (Not bActiveConcept_01) And (Not bActiveConcept_03) And (Not bActiveConcept_01_Z3) And (Not bActiveConcept_03_Z3) Then
								sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
								sFontEnd = "</FONT>"
							End If
							If iSelectedTab = 2 Then
								sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(sPositionTypeShortName) & sBoldEnd & sFontEnd
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
							Else
								sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
							End If
							If iLevelID = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & sBoldEnd & sFontEnd
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionName) & sBoldEnd & sFontEnd
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
							If CLng(sEndDate) = 30000000 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
							End If
							If bActiveConcept_01 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_01_RecordID & "&ConceptID=1&StartDate=" & lConcept_01_StartDate & """"
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_01), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_01), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_01_RecordID) Then sRecordIDs = sRecordIDs & lConcept_01_RecordID & ","
							End If
							If bActiveConcept_03 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_03_RecordID & "&ConceptID=3&StartDate=" & lConcept_03_StartDate & """"
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_03), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_03), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_03_RecordID) Then sRecordIDs = sRecordIDs & lConcept_03_RecordID & ","
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(dConcept_01) + CDbl(dConcept_03), 2, True, False, True)
							If bActiveConcept_01_Z3 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_01_Z3_RecordID & "&ConceptID=1&StartDate=" & lConcept_01_Z3_StartDate & """"
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_01_Z3), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_01_Z3), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_01_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_01_Z3_RecordID & ","
							End If
							If bActiveConcept_03_Z3 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_03_Z3_RecordID & "&ConceptID=3&StartDate=" & lConcept_03_Z3_StartDate & """"
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_03_Z3), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_03_Z3), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_03_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_03_Z3_RecordID & ","
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(dConcept_01_Z3) + CDbl(dConcept_03_Z3), 2, True, False, True)
							If (Not bActiveConcept_01) And (Not bActiveConcept_03) And (Not bActiveConcept_01_Z3) And (Not bActiveConcept_03_Z3) Then
								If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
									If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
									sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
									If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
										sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
											sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
										sRowContents = sRowContents & "</A>&nbsp;"
										sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""RecordIDChk"" ID=""RecordIDChk"" Value=""" & sRecordIDs & """ CHECKED=""1"" />"
									End If
								End If
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If (dConcept_01 + dConcept_03 + dConcept_01_Z3 + dConcept_03_Z3) > 0 Then
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
							End If
						Case 3
							sBoldBegin = ""
							sBoldEnd = ""
							If (StrComp(CStr(lConcept_01_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_03_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) _
							Then
								sBoldBegin = "<B>"
								sBoldEnd = "</B>"
							End If
							sFontBegin = ""
							sFontEnd = ""
							If (Not bActiveConcept_01) And (Not bActiveConcept_03) Then
								sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
								sFontEnd = "</FONT>"
							End If
							sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
							If iLevelID = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & sBoldEnd & sFontEnd
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionName) & sBoldEnd & sFontEnd
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
							If CLng(sEndDate) = 30000000 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
							End If
							If bActiveConcept_01 Then
								sFontBegin = ""
								sFontEnd = ""
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=3&RecordID=" & lConcept_01_RecordID & "&ConceptID=1&StartDate=" & lConcept_01_StartDate & """"
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_01), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_01), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_01_RecordID) Then sRecordIDs = sRecordIDs & lConcept_01_RecordID & ","
							End If
							If bActiveConcept_03 Then
								sFontBegin = ""
								sFontEnd = ""
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=3&RecordID=" & lConcept_03_RecordID & "&ConceptID=3&StartDate=" & lConcept_03_StartDate & """"
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_03), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_03), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_03_RecordID) Then sRecordIDs = sRecordIDs & lConcept_03_RecordID & ","
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(dConcept_01) + CDbl(dConcept_03), 2, True, False, True)
							If (Not bActiveConcept_01) And (Not bActiveConcept_03) Then
								If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
									If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
									sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
									If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
										sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
											sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
										sRowContents = sRowContents & "</A>&nbsp;"
										sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""RecordIDChk"" ID=""RecordIDChk"" Value=""" & sRecordIDs & """ CHECKED=""1"" />"
									End If
								End If
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If (dConcept_01 + dConcept_03) > 0 Then
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
							End If
						Case 5
							sBoldBegin = ""
							sBoldEnd = ""
							If (StrComp(CStr(lConcept_B2_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_36_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_B2_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_36_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) _
							Then
								sBoldBegin = "<B>"
								sBoldEnd = "</B>"
							End If
							sFontBegin = ""
							sFontEnd = ""
							If (Not bActiveConcept_B2) And (Not bActiveConcept_36) And (Not bActiveConcept_B2_Z3) And (Not bActiveConcept_36_Z3) Then
								sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
								sFontEnd = "</FONT>"
							End If
							sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionName) & sBoldEnd & sFontEnd
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
							If CLng(sEndDate) = 30000000 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
							End If
							If bActiveConcept_B2 Then
								sFontBegin = ""
								sFontEnd = ""
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=5&RecordID=" & lConcept_B2_RecordID & "&ConceptID=89&StartDate=" & lConcept_B2_StartDate & """"
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_B2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_B2), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_B2_RecordID) Then sRecordIDs = sRecordIDs & lConcept_B2_RecordID & ","
							End If
							If bActiveConcept_36 Then
								sFontBegin = ""
								sFontEnd = ""
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=5&RecordID=" & lConcept_36_RecordID & "&ConceptID=39&StartDate=" & lConcept_36_StartDate & """"
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_36), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_36), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_36_RecordID) Then sRecordIDs = sRecordIDs & lConcept_36_RecordID & ","
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(dConcept_B2) + CDbl(dConcept_36), 2, True, False, True)
							If bActiveConcept_B2_Z3 Then
								sFontBegin = ""
								sFontEnd = ""
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=5&RecordID=" & lConcept_B2_Z3_RecordID & "&ConceptID=89&StartDate=" & lConcept_B2_StartDate & """"
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_B2_Z3), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_B2), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_B2_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_B2_Z3_RecordID & ","
							End If
							If bActiveConcept_36_Z3 Then
								sFontBegin = ""
								sFontEnd = ""
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=5&RecordID=" & lConcept_36_Z3_RecordID & "&ConceptID=39&StartDate=" & lConcept_36_StartDate & """"
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_36_Z3), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_36), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_36_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_36_Z3_RecordID & ","
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(dConcept_B2_Z3) + CDbl(dConcept_36_Z3), 2, True, False, True)
							If (Not bActiveConcept_B2) And (Not bActiveConcept_36) And (Not bActiveConcept_B2_Z3) And (Not bActiveConcept_36_Z3) Then
								If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
									If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
									sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
									If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
										sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
											sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
										sRowContents = sRowContents & "</A>&nbsp;"
										sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""RecordIDChk"" ID=""RecordIDChk"" Value=""" & sRecordIDs & """ CHECKED=""1"" />"
									End If
								End If
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If (dConcept_B2 + dConcept_36) > 0 Then
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
							End If
						Case 6
							sBoldBegin = ""
							sBoldEnd = ""
							If (StrComp(CStr(lConcept_12_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Then
								sBoldBegin = "<B>"
								sBoldEnd = "</B>"
							End If
							sFontBegin = ""
							sFontEnd = ""
							If (Not bActiveConcept_12) Then
								sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
								sFontEnd = "</FONT>"
							End If
							sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionName) & sBoldEnd & sFontEnd
							If iLevelID = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & sBoldEnd & sFontEnd
							End If
							If iEconomicZoneID = 0 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(iEconomicZoneID) & sBoldEnd & sFontEnd
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
							If CLng(sEndDate) = 30000000 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
							End If
							If bActiveConcept_12 Then
								sFontBegin = ""
								sFontEnd = ""
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=6&RecordID=" & lConcept_12_RecordID & "&ConceptID=14&StartDate=" & lConcept_12_StartDate & """"
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_12), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept_12), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_12_RecordID) Then sRecordIDs = sRecordIDs & lConcept_12_RecordID & ","
							End If
							If Not bActiveConcept_12 Then
								If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
									If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
									sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
									If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
										sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
											sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
										sRowContents = sRowContents & "</A>&nbsp;"
										sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""RecordIDChk"" ID=""RecordIDChk"" Value=""" & sRecordIDs & """ CHECKED=""1"" />"
									End If
								End If
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If dConcept_12 > 0 Then
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
							End If
					End Select
				Case Else
					sBoldBegin = ""
					sBoldEnd = ""
					If (StrComp(CStr(lConcept_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sFontBegin = ""
					sFontEnd = ""
					If (Not bActiveConcept) Then
						sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
						sFontEnd = "</FONT>"
					End If
					If CInt(oRecordset.Fields("PositionTypeID").Value) = -1 Then
						sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
					Else
						sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PositionTypeShortName").Value)) & sBoldEnd & sFontEnd
					End If
					If CInt(oRecordset.Fields("PositionID").Value) = -1 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(oRecordset.Fields("PositionShortName").Value) & sBoldEnd & sFontEnd
					End If
					If CInt(oRecordset.Fields("LevelID").Value) = -1 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(oRecordset.Fields("LevelShortName").Value) & sBoldEnd & sFontEnd
					End If
					If CSng(oRecordset.Fields("WorkingHours").Value) = -1 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("WorkingHours").Value)) & " Hrs." & sBoldEnd & sFontEnd
					End If
					If CInt(oRecordset.Fields("PositionID").Value) = -1 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(oRecordset.Fields("PositionName").Value) & sBoldEnd & sFontEnd
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
					If CLng(sEndDate) = 30000000 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
					End If
					If bActiveConcept Then
						sFontBegin = ""
						sFontEnd = ""
						sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
						sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_RecordID & "&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&StartDate=" & lConcept_StartDate & """"
						sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
					Else
						sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
						sFontEnd = "</FONT>"
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(dConcept), 2, True, False, True) & sBoldEnd & sFontEnd
						If Not IsEmpty(lConcept_RecordID) Then sRecordIDs = sRecordIDs & lConcept_RecordID & ","
					End If
					If Not bActiveConcept Then
						If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
							If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
							sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;"
							If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
								sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;"
								sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""RecordIDChk"" ID=""RecordIDChk"" Value=""" & sRecordIDs & """ CHECKED=""1"" />"
							End If
						End If
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If dConcept > 0 Then
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
					End If
			End Select
			Response.Write "</TABLE></DIV><BR /><BR />"
		Else
			If (iStatusID = 1) Then
				If (Len(sStartDateCondition) = 0) And (InStr(1, sCondition, "Positions.PositionID In (0)", vbBinaryCompare) > 0) Then
					lErrorNumber = L_ERR_NO_RECORDS
					sErrorDescription = "Seleccione un rango de fechas (por lo menos la fecha de inicio, si no se indica la fecha de fin se buscaran con fecha indefinida) o un puesto del filtro para poder consultar los tabuladores existentes."
				Else
					lErrorNumber = L_ERR_NO_RECORDS
					sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
				End If
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
			End If
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1435 = lErrorNumber
	Err.Clear
End Function
%>