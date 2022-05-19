<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
%>
<!-- #include file="Libraries/GlobalVariables.asp" -->
<!-- #include file="Libraries/EmployeeSupportLib.asp" -->
<!-- #include file="Libraries/LoginComponent.asp" -->
<!-- #include file="Libraries/ZoneComponent.asp" -->
<!-- #include file="Libraries/EmployeeSpecialJourneyComponent.asp" -->
<%
Dim lRecordID
Dim sRecordID
Dim sFieldID
Dim sAction
Dim asTemp
Dim bEmpty
Dim sCondition
Dim sFullAreaCondition
Dim sAreaCondition
Dim oRecordset
Dim lStartDate
Dim lEndDate
Dim dCounter
Dim dTotal
Dim JobType1
Dim JobType2
Dim JobDesc1
Dim JobDesc2
Dim sOwnerIDs
Dim sQuery
Dim oConceptsRecordset
Dim lBranchId
Dim lCenterTypeID
Dim lRiskLevel
Dim lPositionID2
Dim lServiceID2
Dim lUR
Dim lCT
Dim lAux
Dim lUca
Dim lEmployeeTypeID

If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
	sFullAreaCondition = " And ((Employees.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (Areas.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
	sAreaCondition = " And (Areas.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))"
End If
If Len(oRequest) > 0 Then
	sAction = oRequest("Action").Item
	Select Case sAction
		Case "AreasForCompany"
		Case "BankAccounts"
		Case "Budget_Level2", "Budget_Level3"
		Case "EmployeeAccessKey"
			lErrorNumber = GetNameFromTable(oADODBConnection, "EmployeesIDFromAccessKey", "'" & Replace(oRequest("RecordID").Item, "'", "") & "'", "", "", sRecordID, sErrorDescription)
		Case "EmployeeGender"
			lErrorNumber = GetNameFromTable(oADODBConnection, "EmployeesGenders", Replace(oRequest("RecordID").Item, "'", ""), "", "", sRecordID, sErrorDescription)
		Case "EmployeeConcept"
			sRecordID = ""
			sErrorDescription = "No se pudo obtener la información del registro."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesConceptsLKP Where (EmployeeID=" & Replace(oRequest("RecordID").Item, "'", "") & ") And (ConceptID=" & oRequest("ConceptID").Item & ") And (EndDate=30000000)", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sRecordID = sRecordID & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True)
					sRecordID = sRecordID & ", del " & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
					If CLng(oRecordset.Fields("EndDate").Value) < 30000000 Then
						sRecordID = sRecordID & " al " & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1)
					Else
						sRecordID = sRecordID & " a la fecha"
					End If
				End If
				oRecordset.Close
			End If
		Case "EmployeeHeadNumber"
			sErrorDescription = "No se pudo obtener la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, EmployeeName, EmployeeLastName, EmployeeLastName2 From Employees Where (EmployeeID=" & Replace(oRequest("RecordID").Item, "'", "") & ") And (EmployeeTypeID = 1)", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sFieldID = CStr(oRecordset.Fields("EmployeeID").Value)
					If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
						sRecordID = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value)) & "</B><BR />"
					Else
						sRecordID = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value)) & "</B><BR />"
					End If
				End If
				oRecordset.Close
			End If
		Case "EmployeesGyS"
		Case "EmployeesInfo"
			sErrorDescription = "No se pudo obtener la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.*, PositionShortName, PositionName, JourneyShortName, JourneyName, ShiftShortName, ShiftName, Areas.AreaCode, Areas.AreaName, ServiceShortName, ServiceName, BranchShortName, BranchName, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, EmployeeTypeName From Employees, Jobs, Positions, Journeys, Shifts, Areas, Services, Branches, Areas As PaymentCenters, EmployeeTypes Where (Employees.JobID=Jobs.JobID) And (Jobs.PositionID=Positions.PositionID) And (Employees.JourneyID=Journeys.JourneyID) And (Employees.ShiftID=Shifts.ShiftID) And (Jobs.AreaID=Areas.AreaID) And (Employees.ServiceID=Services.ServiceID) And (Areas.BranchID=Branches.BranchID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Employees.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeeID=" & Replace(oRequest("RecordID").Item, "'", "") & ")" & sFullAreaCondition, "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sFieldID = CStr(oRecordset.Fields("EmployeeID").Value)
					sRecordID = "<B>Información del empleado:</B><BR /><B>No. de empleado: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & "<BR />"
					If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
						sRecordID = sRecordID & "<B>Nombre: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value)) & "<BR />"
					Else
						sRecordID = sRecordID & "<B>Nombre: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value)) & "<BR />"
					End If
					sRecordID = sRecordID & "<B>RFC: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)) & "<BR />"
					sRecordID = sRecordID & "<B>CURP: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value)) & "<BR />"
					If Len(oRequest("Full").Item) > 0 Then sRecordID = sRecordID & "<B>Plaza: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("JobID").Value)) & "<BR />"
					sRecordID = sRecordID & "<B>Puesto: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value)) & "<BR />"
					If Len(oRequest("Full").Item) > 0 Then
						sRecordID = sRecordID & "<B>Turno: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("JourneyShortName").Value) & ". " & CStr(oRecordset.Fields("JourneyName").Value)) & "<BR />"
						sRecordID = sRecordID & "<B>Horario: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("ShiftShortName").Value) & ". " & CStr(oRecordset.Fields("ShiftName").Value)) & "<BR />"
						sRecordID = sRecordID & "<B>Adscripción: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value) & ". " & CStr(oRecordset.Fields("AreaName").Value)) & "<BR />"
						sRecordID = sRecordID & "<B>Servicio: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("ServiceShortName").Value) & ". " & CStr(oRecordset.Fields("ServiceName").Value)) & "<BR />"
						sRecordID = sRecordID & "<B>Rama: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("BranchShortName").Value) & ". " & CStr(oRecordset.Fields("BranchName").Value)) & "<BR />"
					Else
						sRecordID = sRecordID & "<B>Centro de pago: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value) & ". " & CStr(oRecordset.Fields("PaymentCenterName").Value)) & "<BR />"
						sRecordID = sRecordID & "<B>Fecha de ingreso: </B>" & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1) & "<BR />"
						sRecordID = sRecordID & "<B>Tipo de empleado: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeName").Value)) & "<BR />"
					End If
				End If
				oRecordset.Close
			End If
		Case "EmployeesNameFromNumber"
			lErrorNumber = GetNameFromTable(oADODBConnection, "EmployeesNameFromNumber", "'" & Right("000000" & Replace(oRequest("RecordID").Item, "'", ""), Len("000000")) & "'", "", "", sRecordID, sErrorDescription)
		Case "EmployeeNumber"
			'lErrorNumber = GetNameFromTable(oADODBConnection, "EmployeesIDFromNumber", "'" & Right("000000" & Replace(oRequest("RecordID").Item, "'", ""), Len("000000")) & "'", "", "", sRecordID, sErrorDescription)
			sErrorDescription = "No se pudo obtener la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, EmployeeName, EmployeeLastName, EmployeeLastName2 From Employees Where (EmployeeID=" & Replace(oRequest("RecordID").Item, "'", "") & ")", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sFieldID = CStr(oRecordset.Fields("EmployeeID").Value)
					If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
						sRecordID = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value)) & "</B><BR />"
					Else
						sRecordID = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value)) & "</B><BR />"
					End If
				End If
				oRecordset.Close
			End If
		Case "EmployeePayment"
			sErrorDescription = "No se pudo obtener la información del pago del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Payments.EmployeeID, CheckNumber From EmployeesChangesLKP, EmployeesHistoryList, Payments Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeID=Payments.EmployeeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=Payments.PaymentDate) And (Payments.PaymentDate=" & oRequest("PaymentDate").Item  & ") And (EmployeesHistoryList.EmployeeID=" & Replace(oRequest("RecordID").Item, "'", "") & ")", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sFieldID = CStr(oRecordset.Fields("EmployeeID").Value)
					sRecordID = CStr(oRecordset.Fields("CheckNumber").Value)
				End If
				oRecordset.Close
			End If
		Case "EmployeeRFC"
			lErrorNumber = GetNameFromTable(oADODBConnection, "EmployeesIDFromRFC", "'" & Replace(oRequest("RecordID").Item, "'", "") & "'", "", "", sRecordID, sErrorDescription)
		Case "ExternalGyS"
		Case "JobSwap"
			sErrorDescription = "No se pudo obtener la información del empleado"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PositionShortName, PositionLongName from Positions, (Select PositionID From Jobs Where jobID = " & oRequest("JobID1").Item & ") as TipoPlaza Where Positions.PositionID = TipoPlaza.PositionID", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				JobType1 = CStr(oRecordset.Fields("PositionShortName").Value)
				JobDesc1 = CStr(oRecordset.Fields("PositionLongName").Value)
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.*, PositionShortName, PositionLongName, PositionName, PositionTypeName, PaymentCenters.AreaShortName As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, EmployeeTypeName, JourneyName, ServiceName From Employees, Jobs, Positions, PositionTypes, Areas As PaymentCenters, EmployeeTypes, Journeys, Services Where (Employees.JobID=Jobs.JobID) And (Jobs.PositionID=Positions.PositionID) And (Positions.PositionTypeID=PositionTypes.PositionTypeID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Employees.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeeID=" & Replace(oRequest("RecordID").Item, "'", "") & ")And (Employees.ServiceID = Services.ServiceID) And (Employees.JourneyID = Journeys.JourneyID)", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						JobType2 = CStr(oRecordset.Fields("PositionShortName").Value)
						JobDesc2 = CStr(oRecordset.Fields("PositionLongName").Value)
						If (StrComp(JobType1, JobType2, vbBinaryCompare) = 0) Then
							sFieldID = CStr(oRecordset.Fields("EmployeeID").Value)
							sRecordID = "<B>Información del empleado:</B><BR />No. de empleado: " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & "<BR />"
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
								sRecordID = sRecordID & "<B>Nombre: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value)) & "<BR />"
							Else
								sRecordID = sRecordID & "<B>Nombre: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value)) & "<BR />"
							End If
							sRecordID = sRecordID & "<B>RFC: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)) & "<BR />"
							sRecordID = sRecordID & "<B>CURP: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value)) & "<BR />"
							sRecordID = sRecordID & "<B>Fecha de ingreso: </B>" & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1) & "<BR />"
							sRecordID = sRecordID & "<B>Tipo de empleado: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeName").Value)) & "<BR />"
							sRecordID = sRecordID & "<B>Plaza: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("JobID").Value)) & "<BR />"
							sRecordID = sRecordID & "<B>Puesto: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value)) & "<BR />"
							sRecordID = sRecordID & "<B>Tipo de puesto: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("PositionTypeName").Value)) & "<BR />"
							sRecordID = sRecordID & "<B>Adscripción: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value) & ". " & CStr(oRecordset.Fields("PaymentCenterName").Value)) & "<BR />"
							sRecordID = sRecordID & "<B>Turno: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("JournyeID").Value) & ". " & CStr(oRecordset.Fields("JourneyName").Value)) & "<BR />"
							sRecordID = sRecordID & "<B>Servicio: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("ServiceID").Value) & ". " & CStr(oRecordset.Fields("ServiceName").Value)) & "<BR />"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesConceptsLKP.ConceptID, ConceptShortName, ConceptName, ConceptAmount From EmployeesConceptsLKP, Concepts Where EmployeeID = " & Replace(oRequest("RecordID").Item, "'", "") & " And (EmployeesConceptsLKP.StartDate < " & CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) & ") And (EmployeesConceptsLKP.EndDate > " & CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) & ") And (EmployeesConceptsLKP.ConceptID In (4,7)) And (EmployeesConceptsLKP.ConceptID = Concepts.ConceptID)", "SearchRecord.asp", "_root", 000, sErrorDescription, oConceptsRecordset)
							If lErrorNumber = 0 then
								If Not oConceptsRecordset.EOF Then
									Do While Not oConceptsRecordset.EOF
										If CInt(oConceptsRecordset.Fields("ConceptID").Value) = 4 Then
											sRecordID = sRecordID & "<B>" & CleanStringForHTML(CStr(oConceptsRecordset.Fields("ConceptShortName").Value) & "  " & CStr(oConceptsRecordset.Fields("ConceptName").Value)) & "</B> (" & CStr(oConceptsRecordset.Fields("ConceptAmount").Value) & ")" & "<BR />"
										ElseIf CInt(oConceptsRecordset.Fields("ConceptID").Value) = 7 Then
											sRecordID = sRecordID & "<B>" & CleanStringForHTML(CStr(oConceptsRecordset.Fields("ConceptShortName").Value) & "  " & CStr(oConceptsRecordset.Fields("ConceptName").Value)) & "</B> (" & Mid(CStr(oConceptsRecordset.Fields("ConceptAmount").Value),1,5) & ")" & "<BR />"
										End If
										oConceptsRecordset.MoveNext
									Loop
									oConceptsRecordset.Close
								End If
							End If
						Else
							sRecordID = "Las plazas no son iguales: </BR>"
							sRecordID = sRecordID & CleanStringForHTML(JobDesc1) & " - " & CleanStringForHTML(JobDesc2)
						End If
					End If
					oRecordset.Close
				End If
			End If
		Case "DocumentsForLicenses"
			sErrorDescription = "No se pudo obtener la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.*, PositionShortName, PositionName, PositionTypeName, PaymentCenters.AreaShortName As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, EmployeeTypeName From Employees, Jobs, Positions, PositionTypes, Areas As PaymentCenters, EmployeeTypes Where (Employees.JobID=Jobs.JobID) And (Jobs.PositionID=Positions.PositionID) And (Positions.PositionTypeID=PositionTypes.PositionTypeID) And (Employees.PaymentCenterID=PaymentCenters.AreaID) And (Employees.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeeID=" & Replace(oRequest("RecordID").Item, "'", "") & ")", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sFieldID = CStr(oRecordset.Fields("EmployeeID").Value)
					sRecordID = "<B>Información del empleado:</B><BR />No. de empleado: " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & "<BR />"
					If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
						sRecordID = sRecordID & "<B>Nombre: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value)) & "<BR />"
					Else
						sRecordID = sRecordID & "<B>Nombre: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value)) & "<BR />"
					End If
					sRecordID = sRecordID & "<B>RFC: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)) & "<BR />"
					sRecordID = sRecordID & "<B>CURP: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value)) & "<BR />"
					sRecordID = sRecordID & "<B>Puesto: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value)) & "<BR />"
					sRecordID = sRecordID & "<B>Tipo de puesto: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("PositionTypeName").Value)) & "<BR />"
					sRecordID = sRecordID & "<B>Adscripción: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value) & ". " & CStr(oRecordset.Fields("PaymentCenterName").Value)) & "<BR />"
					sRecordID = sRecordID & "<B>Fecha de ingreso: </B>" & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1) & "<BR />"
					sRecordID = sRecordID & "<B>Tipo de empleado: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeName").Value)) & "<BR />"
				End If
				oRecordset.Close
			End If
		Case "JobID", "JobNumber"
			If (CLng(oRequest("lReasonID").Item) = 13) Then
				If (oRequest("StartYear").Item = 0) Or (oRequest("StartMonth").Item = 0) Or (oRequest("StartDay").Item = 0) _
					Or (oRequest("EndYear").Item = 0) Or (oRequest("EndMonth").Item = 0) Or (oRequest("EndDay").Item = 0) Then
					lErrorNumber = -1
					sErrorDescription = "Escriba las fechas inicial y final de la vigencia"
				End If
				If (lErrorNumber = 0) Then
					lStartDate = oRequest("StartYear").Item & oRequest("StartMonth").Item & oRequest("StartDay").Item
					lEndDate = oRequest("EndYear").Item & oRequest("EndMonth").Item & oRequest("EndDay").Item
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "select StatusId, JobDate, EndDate From JobsHistoryList Where JobID = " & oRequest("RecordId").Item & " And (JobDate <= " & lStartDate & " And EndDate >= " & lEndDate & ")", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
					If Not oRecordset.EOF Then
						If CInt(oRecordset.Fields("StatusID").Value) = 1 Then
							lErrorNumber = -1
							sErrorDescription = "Las vigencia del interinato incluyen un periodo no vacante de la plaza. Verifique la vacancia o la vigencia del interinato"
						End If
					End If
				End If
			End If
			If (lErrorNumber = 0) Then
				sErrorDescription = "No se pudo obtener la información de la plaza."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JobID, JobNumber, CompanyName, JobTypeShortName, JobTypeName, JourneyShortName, JourneyName, Jobs.GroupGradeLevelID, Jobs.ClassificationID, Jobs.IntegrationID, Jobs.StartDate, Jobs.EndDate, Jobs.WorkingHours, ServiceShortName, ServiceName, Jobs.LevelID, LevelShortName, LevelName, GroupGradeLevelName, StatusShortName, StatusName, PositionShortName, PositionName, PositionTypeShortName, PositionTypeName, EmployeeTypeShortName, EmployeeTypeName, ShiftShortName, ShiftName, ZoneCode, ZoneName, Areas.AreaCode, Areas.AreaName, PaymentCenters.AreaCode PaymentCenterShortName, PaymentCenters.AreaName PaymentCenterName From Areas, Companies, EmployeeTypes, GroupGradeLevels, Jobs, JobTypes, Journeys, Levels, Areas PaymentCenters, Positions, PositionTypes, Services, Shifts, StatusJobs, Zones Where (Jobs.CompanyID = Companies.CompanyID) And (Jobs.ZoneID = Zones.ZoneID) And (Jobs.AreaID = Areas.AreaID) And (Jobs.PaymentCenterID = PaymentCenters.AreaID) And (Jobs.GroupGradeLevelID = GroupGradeLevels.GroupGradeLevelID) And (PositionTypes.PositionTypeID = Positions.PositionTypeID) And (EmployeeTypes.EmployeeTypeID = Positions.EmployeeTypeID) And (Jobs.PositionID = Positions.PositionID) And (Jobs.PositionID <> -1) And (Jobs.JobTypeID = JobTypes.JobTypeID) And (Jobs.ShiftID = Shifts.ShiftID) And (Jobs.JourneyID = Journeys.JourneyID) And (Jobs.ServiceID = Services.ServiceID) And (Jobs.LevelID = Levels.LevelID) And (Jobs.StatusID = StatusJobs.StatusID) And (JobID=" & Replace(oRequest("RecordID").Item, "'", "") & ")", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
			End If
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sFieldID = CStr(oRecordset.Fields("JobID").Value)
					sRecordID = "<BR /><B>Información de la plaza:</B><BR /><B>No. de plaza:</B> " & CleanStringForHTML(CStr(oRecordset.Fields("JobNumber").Value)) & "<BR />"
					sRecordID = sRecordID & "<B>Empresa:</B> " & CleanStringForHTML(CStr(oRecordset.Fields("CompanyName").Value)) & "<BR />"
					sRecordID = sRecordID & "<B>Puesto: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value)) & "<BR />"
					If CLng(oRecordset.Fields("LevelID").Value) <> -1 Then
						sRecordID = sRecordID & "<B>Nivel: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("LevelName").Value)) & "<BR />"
					Else
						sRecordID = sRecordID & "<B>GGN: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("GroupGradeLevelName").Value)) & "<BR />"
						sRecordID = sRecordID & "<B>Clasificación: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("ClassificationID").Value)) & "<BR />"
						sRecordID = sRecordID & "<B>Integración: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("IntegrationID").Value)) & "<BR />"
					End If
					sRecordID = sRecordID & "<B>Jornada: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("Workinghours").Value)) & "<BR />"
					sRecordID = sRecordID & "<B>Tipo de puesto: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("PositionTypeShortName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("PositionTypeName").Value)) & "<BR />"
					sRecordID = sRecordID & "<B>Centro de trabajo: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value)) & "<BR />"
					sRecordID = sRecordID & "<B>Centro de pago: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterName").Value)) & "<BR />"
					sRecordID = sRecordID & "<B>Servicio: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("ServiceShortName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("ServiceName").Value)) & "<BR />"
					sRecordID = sRecordID & "<B>Turno: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("JourneyShortName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("JourneyName").Value)) & "<BR />"
					sRecordID = sRecordID & "<B>Tipo de plaza: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("JobTypeShortName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("JobTypeName").Value)) & "<BR />"
					sRecordID = sRecordID & "<B>Estatus de la plaza: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value)) & "<BR />"
					If CLng(oRecordset.Fields("StartDate").Value) = 0 Then
						sRecordID = sRecordID & "<B>Fecha inicio: </B>Indefinida<BR />"
					Else
						sRecordID = sRecordID & "<B>Fecha inicio: </B>" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)) & "<BR />"
					End If
					If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
						sRecordID = sRecordID & "<B>Fecha fin: </B>Indefinida<BR />"
					Else
						sRecordID = sRecordID & "<B>Fecha fin: </B>" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & "<BR />"
					End If
					oRecordset.Close
					sErrorDescription = "No se pudo obtener la información del empleado."
					'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusName, Employees.EmployeeNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, JobsHistoryList.JobDate, JobsHistoryList.EndDate From Employees, JobsHistoryList, StatusJobs Where JobsHistoryList.EmployeeID = Employees.EmployeeID And JobsHistoryList.StatusID = StatusJobs.StatusID And JobsHistoryList.JobID=" & Replace(oRequest("RecordID").Item, "'", "") & " Order By JobDate Desc", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusName, JobsHistoryList.EmployeeID, JobsHistoryList.JobDate, JobsHistoryList.EndDate From JobsHistoryList, StatusJobs Where JobsHistoryList.StatusID = StatusJobs.StatusID And JobsHistoryList.JobID=" & Replace(oRequest("RecordID").Item, "'", "") & " Order By JobDate Desc", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							sRecordID = sRecordID & "<BR /><B>Historial de la plaza:</B><BR /><BR />"
							Do While Not oRecordset.EOF
								sRecordID = sRecordID & "<B>Número de empleado: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeID").Value)) & "<BR />"
								sRecordID = sRecordID & "<B>Estatus plaza: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value)) & "<BR />"
								If CLng(oRecordset.Fields("JobDate").Value) = 0 Then
									sRecordID = sRecordID & "<B>Fecha inicio vigencia: </B>Indefinida<BR />"
								Else
									sRecordID = sRecordID & "<B>Fecha inicio vigencia: </B>" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("JobDate").Value)) & "<BR />"
								End If
								If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
									sRecordID = sRecordID & "<B>Fecha fin vigencia: </B>Indefinida<BR />"
								Else
									sRecordID = sRecordID & "<B>Fecha fin vigencia: </B>" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & "<BR />"
								End If
								sRecordID = sRecordID & "<BR />"
								oRecordset.MoveNext
								If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
							Loop
							oRecordset.Close
						End If
					End If
				End If
			End If
		Case "JobsInfo"
			sErrorDescription = "No se pudo obtener la información de la plaza."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Jobs.*, CompanyShortName, CompanyName, AreaCode, AreaName, PositionShortName, PositionName, ShiftShortName, ShiftName, JourneyShortName, JourneyName, ServiceShortName, ServiceName, BranchShortName, BranchName From Jobs, Companies, Areas, Positions, Shifts, Journeys, Services, Branches Where (Jobs.CompanyID=Companies.CompanyID) And (Jobs.AreaID=Areas.AreaID) And (Jobs.PositionID=Positions.PositionID) And (Jobs.ShiftID=Shifts.ShiftID) And (Jobs.JourneyID=Journeys.JourneyID) And (Jobs.ServiceID=Services.ServiceID) And (Areas.BranchID=Branches.BranchID) And (JobID=" & Replace(oRequest("RecordID").Item, "'", "") & ")", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sFieldID = CStr(oRecordset.Fields("JobID").Value)
					If Len(oRequest("SendPosition").Item) > 0 Then lRecordID = CLng(oRecordset.Fields("PositionID").Value)
					sRecordID = "<B>Información de la plaza:</B><BR /><B>No. de plaza: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("JobNumber").Value)) & "<BR />"
					sRecordID = sRecordID & "<B>Compañía: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("CompanyShortName").Value) & ". " & CStr(oRecordset.Fields("CompanyName").Value)) & "<BR />"
					sRecordID = sRecordID & "<B>Adscripción: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value) & ". " & CStr(oRecordset.Fields("AreaName").Value)) & "<BR />"
					sRecordID = sRecordID & "<B>Puesto: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value)) & "<BR />"
					sRecordID = sRecordID & "<B>Horario: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("ShiftShortName").Value) & ". " & CStr(oRecordset.Fields("ShiftName").Value)) & "<BR />"
					sRecordID = sRecordID & "<B>Turno: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("JourneyShortName").Value) & ". " & CStr(oRecordset.Fields("JourneyName").Value)) & "<BR />"
					sRecordID = sRecordID & "<B>Servicio: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("ServiceShortName").Value) & ". " & CStr(oRecordset.Fields("ServiceName").Value)) & "<BR />"
					sRecordID = sRecordID & "<B>Rama: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("BranchShortName").Value) & ". " & CStr(oRecordset.Fields("BranchName").Value)) & "<BR />"
				End If
				oRecordset.Close
			End If
		Case "OwnerParentID"
			sErrorDescription = "No se pudo obtener la información del registro."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select OwnerID From PaperworkOwners Where (OwnerID=" & oRequest("RecordID").Item & ") And (LevelID=" & (CInt(oRequest("LevelID").Item) - 1) & ") And (OwnerID>-1)", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sFieldID = CStr(oRecordset.Fields("OwnerID").Value)
					sRecordID = CStr(oRecordset.Fields("OwnerID").Value)
				End If
				oRecordset.Close
			End If
		Case "PaperworkCatalogs"
		Case "Paperworks"
			sErrorDescription = "No se pudo obtener la información del registro."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaperworkID From Paperworks Where (PaperworkNumber='" & oRequest("RecordID").Item & "') And (StartDate>=" & CLng(oRequest("StartYear").Item & "0000") & ")", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sFieldID = CStr(oRecordset.Fields("PaperworkID").Value)
					sRecordID = CStr(oRecordset.Fields("PaperworkID").Value)
				End If
				oRecordset.Close
			End If
		Case "PositionName"
		Case "PositionsByType"
		Case "PositionsCatalogsLKP"
		Case "PositionsForEmployeeType"
		Case "PositionsGyS"
		Case "RecordsForGyS"
		Case "RiskLevel"
			sErrorDescription = "No se pudo obtener la información de la matriz de riesgos"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PositionID, ServiceID from Jobs where JobID = (Select JobId From Employees where EmployeeID = " & oRequest("EmployeeID").Item & ")", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
			If Not oRecordset.EOF Then
				lPositionID2 = oRecordset.Fields("PositionID").Value
				lServiceID2 = oRecordset.Fields("ServiceID").Value
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CenterTypeId From Services Where ServiceID = " & lServiceID2, "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
				lCenterTypeID = oRecordset.Fields("CenterTypeID").Value
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select BranchId From Positions Where PositionId = " & lPositionID2, "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
				lBranchId = oRecordset.Fields("BranchID").Value
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select RiskLevel From ProfessionalRiskMatrix Where (PositionID=" & lPositionID2 & ") And (ServiceID=" & lServiceID2 & ") And (CenterTypeID=" & lCenterTypeID & ") And (BranchID=" & lBranchID & ")", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
				lRiskLevel = oRecordset.Fields("RiskLevel").Value
			End If
		Case "ZoneForArea"
			sFieldID = -1
			sErrorDescription = "No se pudo obtener la información del registro."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Zones.ZoneID, Zones.ParentID, ZoneCode, ZoneName From Areas, Zones Where (Areas.ZoneID=Zones.ZoneID) And (AreaID=" & Replace(oRequest("RecordID").Item, "'", "") & ")", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sFieldID = CStr(oRecordset.Fields("ZoneID").Value)
					aZoneComponent(N_ID_ZONE) = CLng(oRecordset.Fields("ZoneID").Value)
					aZoneComponent(N_PARENT_ID_ZONE) = CLng(oRecordset.Fields("ParentID").Value)
					sRecordID = DisplayZonePathAsText(oRequest, oADODBConnection, aZoneComponent, sErrorDescription)
				End If
				oRecordset.Close
			End If
		Case "StartDateForConcept"
			sFieldID = -1
			sQuery = "Select StartDate, EndDate From EmployeesConceptsLKP Where (ConceptID = " & oRequest("RecordID").Item & ") And ((EndDate = 30000000) Or (EndDate > " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")) And (EmployeeID = " & oRequest("lEmployeeID").Item & ")"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sFieldID = 1
				End If
			End If
		Case "Zones_Level2"
		Case Else
			sRecordID = -1
			sErrorDescription = "No se pudo obtener la información del registro."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select " & sAction & " From " & oRequest("TableName").Item & " Where (" & oRequest("CodeField").Item & "='" & Replace(oRequest("RecordID").Item, "'", "") & "')", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then sRecordID = CLng(oRecordset.Fields(0).Value)
				oRecordset.Close
			End If
	End Select
End If
%>
<HTML>
	<HEAD>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/CheckFields.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/CommonLibrary.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/Events.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/Forms.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/HTMLLists.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/ImageLoader.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/PopupItem.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/RollOver.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/URLManipulation.js"></SCRIPT>
		<LINK REL="STYLESHEET" TYPE="text/css" HREF="Styles/SIAP.css" />
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<FORM NAME="SearchFrm" ID="SearchFrm"><FONT FACE="Arial" SIZE="2">
			<%If Len(oRequest) > 0 Then
				Select Case sAction
					Case "AreasForCompany"
						sErrorDescription = "No se pudo obtener la información del registro."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaID, AreaCode, AreaName From Areas Where (ParentID>-1) And (EndDate=30000000) And (Active=1) And (CompanyID=" & Replace(oRequest("RecordID").Item, "'", "") & ") Order By AreaCode", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
									Response.Write "RemoveAllItemsFromList(null, parent.window.document." & oRequest("TargetField").Item & ".AreaID);" & vbNewLine
									If lErrorNumber = 0 Then
										Do While Not oRecordset.EOF
											Response.Write "AddItemToList('" & CStr(oRecordset.Fields("AreaCode").Value) & ". " & CStr(oRecordset.Fields("AreaName").Value) & "', '" & CStr(oRecordset.Fields("AreaID").Value) & "', null, parent.window.document." & oRequest("TargetField").Item & ".AreaID);" & vbNewLine
											oRecordset.MoveNext
											If Err.number <> 0 Then Exit Do
										Loop
										oRecordset.Close
									End If
									Response.Write "parent.window.ShowAreaFields();" & vbNewLine
								Response.Write "//--></SCRIPT>" & vbNewLine
							End If
							oRecordset.Close
						End If
					Case "BankAccounts"
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							Response.Write "RemoveAllItemsFromList(null, parent.window.document." & oRequest("TargetField").Item & ");" & vbNewLine
							sErrorDescription = "No se pudo obtener la información del registro."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AccountID, BankName, AccountNumber From BankAccounts, Banks Where (BankAccounts.BankID=Banks.BankID) And (AccountID>-1) And (Banks.BankID=" & Replace(oRequest("RecordID").Item, "'", "") & ") And (BankAccounts.EmployeeID<0) And (BankAccounts.Active=1) And (Banks.Active=1) Order By BankName, AccountNumber", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									Do While Not oRecordset.EOF
										asTemp = Split(CStr(oRecordset.Fields("AccountNumber").Value), LIST_SEPARATOR)
										Response.Write "AddItemToList('" & CStr(oRecordset.Fields("BankName").Value) & ". " & asTemp(0) & "', '" & CStr(oRecordset.Fields("AccountID").Value) & "', null, parent.window.document." & oRequest("TargetField").Item & ");" & vbNewLine
										oRecordset.MoveNext
										If Err.number <> 0 Then Exit Do
									Loop
								Else
									Response.Write "AddItemToList('Ninguna', '-1', null, parent.window.document." & oRequest("TargetField").Item & ");" & vbNewLine
								End If
								oRecordset.Close
							Else
								Response.Write "AddItemToList('Ninguna', '-1', null, parent.window.document." & oRequest("TargetField").Item & ");" & vbNewLine
							End If
						Response.Write "//--></SCRIPT>" & vbNewLine
					Case "Budget_Level2"
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							sErrorDescription = "No se pudo obtener la información del registro."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select BudgetID, BudgetShortName, BudgetName From Budgets Where (BudgetID>-1) And (ParentID=" & Replace(oRequest("RecordID").Item, "'", "") & ") Order By BudgetShortName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									Response.Write "RemoveAllItemsFromList(null, parent.window.document." & oRequest("TargetField").Item & ".BudgetID2);" & vbNewLine
									Response.Write "AddItemToList('Todas', '', null, parent.window.document." & oRequest("TargetField").Item & ".BudgetID2);" & vbNewLine
									Do While Not oRecordset.EOF
										If StrComp(CStr(oRecordset.Fields("BudgetShortName").Value), CStr(oRecordset.Fields("BudgetName").Value), vbBinaryCompare) = 0 Then
											Response.Write "AddItemToList('" & CStr(oRecordset.Fields("BudgetName").Value) & "', '" & CStr(oRecordset.Fields("BudgetID").Value) & "', null, parent.window.document." & oRequest("TargetField").Item & ".BudgetID2);" & vbNewLine
										Else
											Response.Write "AddItemToList('" & CStr(oRecordset.Fields("BudgetShortName").Value) & ". " & CStr(oRecordset.Fields("BudgetName").Value) & "', '" & CStr(oRecordset.Fields("BudgetID").Value) & "', null, parent.window.document." & oRequest("TargetField").Item & ".BudgetID2);" & vbNewLine
										End If
										oRecordset.MoveNext
										If Err.number <> 0 Then Exit Do
									Loop
									Response.Write "RemoveAllItemsFromList(null, parent.window.document." & oRequest("TargetField").Item & ".BudgetID3);" & vbNewLine
									Response.Write "AddItemToList('Todas', '', null, parent.window.document." & oRequest("TargetField").Item & ".BudgetID3);" & vbNewLine
								End If
								oRecordset.Close
							End If
						Response.Write "//--></SCRIPT>" & vbNewLine
					Case "Budget_Level3"
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							sErrorDescription = "No se pudo obtener la información del registro."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select BudgetID, BudgetShortName, BudgetName From Budgets Where (BudgetID>-1) And (ParentID=" & Replace(oRequest("RecordID").Item, "'", "") & ") Order By BudgetShortName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									Response.Write "RemoveAllItemsFromList(null, parent.window.document." & oRequest("TargetField").Item & ".BudgetID3);" & vbNewLine
									Response.Write "AddItemToList('Todas', '', null, parent.window.document." & oRequest("TargetField").Item & ".BudgetID3);" & vbNewLine
									Do While Not oRecordset.EOF
										If StrComp(CStr(oRecordset.Fields("BudgetShortName").Value), CStr(oRecordset.Fields("BudgetName").Value), vbBinaryCompare) = 0 Then
											Response.Write "AddItemToList('" & CStr(oRecordset.Fields("BudgetName").Value) & "', '" & CStr(oRecordset.Fields("BudgetID").Value) & "', null, parent.window.document." & oRequest("TargetField").Item & ".BudgetID3);" & vbNewLine
										Else
											Response.Write "AddItemToList('" & CStr(oRecordset.Fields("BudgetShortName").Value) & ". " & CStr(oRecordset.Fields("BudgetName").Value) & "', '" & CStr(oRecordset.Fields("BudgetID").Value) & "', null, parent.window.document." & oRequest("TargetField").Item & ".BudgetID3);" & vbNewLine
										End If
										oRecordset.MoveNext
										If Err.number <> 0 Then Exit Do
									Loop
								End If
								oRecordset.Close
							End If
						Response.Write "//--></SCRIPT>" & vbNewLine
					Case "EmployeeAccessKey"
						If Len(sRecordID) = 0 Then
							Response.Write "&nbsp;&nbsp;&nbsp;<FONT COLOR=""#" & S_INSTRUCTIONS_FOR_GUI & """><B>Esta clave de acceso está disponible.</B></FONT>"
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								If Len(oRequest("TargetField").Item) > 0 Then
									Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ")" & vbNewLine
										Response.Write "parent.window.document." & oRequest("TargetField").Item & ".value = '1';" & vbNewLine
								End If
							Response.Write "//--></SCRIPT>" & vbNewLine
						Else
							Response.Write "&nbsp;&nbsp;&nbsp;<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>Esta clave de acceso ya está registrada en el sistema.</B></FONT>"
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								If Len(oRequest("TargetField").Item) > 0 Then
									Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ")" & vbNewLine
										Response.Write "parent.window.document." & oRequest("TargetField").Item & ".value = '0';" & vbNewLine
								End If
							Response.Write "//--></SCRIPT>" & vbNewLine
						End If
					Case "EmployeeConcept"
						If Len(sRecordID) = 0 Then
							Response.Write "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>El empleado no tiene registrado este concepto</B></FONT>"
						Else
							Response.Write sRecordID
						End If
					Case "EmployeeGender"
						If Len(sRecordID) = 0 Then
							Response.Write "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>El empleado no está registrado en el sistema</B></FONT>"
						ElseIf CInt(sRecordID) = 0 Then
							Response.Write "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>El empleado es de sexo femenino.</B></FONT>"
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								If Len(oRequest("TargetField").Item) > 0 Then
									Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ")" & vbNewLine
										Response.Write "parent.window.document." & oRequest("TargetField").Item & ".value = '0';" & vbNewLine
								End If
							Response.Write "//--></SCRIPT>" & vbNewLine
						ElseIf CInt(sRecordID) = 1 Then
							Response.Write "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>El empleado es de sexo masculino.</B></FONT>"
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								If Len(oRequest("TargetField").Item) > 0 Then
									Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ")" & vbNewLine
										Response.Write "parent.window.document." & oRequest("TargetField").Item & ".value = '1';" & vbNewLine
								End If
							Response.Write "//--></SCRIPT>" & vbNewLine
						End If
					Case "EmployeesGyS"
						Dim lAreaID
						Dim lServiceID
						Dim lPositionID
						Dim lLevelID
						Dim dWorkingHours
						sErrorDescription = "No se pudo obtener la información del empleado."
						If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) = 0 Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.*, Areas.AreaID, Areas.AreaCode, Areas.AreaName, Positions.PositionID, Positions.PositionShortName, Positions.PositionName, ServiceShortName, ServiceName, ShiftShortName, ShiftName, LevelShortName, PositionsSpecialJourneysLKP.RecordID From Employees, Jobs, Areas, Positions, Services, Shifts, Levels, PositionsSpecialJourneysLKP Where (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Jobs.PositionID=Positions.PositionID) And (Employees.ServiceID=Services.ServiceID) And (Employees.ShiftID=Shifts.ShiftID) And (Employees.LevelID=Levels.LevelID) And (PositionsSpecialJourneysLKP.PositionID=Positions.PositionID) And (PositionsSpecialJourneysLKP.LevelID=Employees.LevelID) And (PositionsSpecialJourneysLKP.WorkingHours=Employees.WorkingHours) And (PositionsSpecialJourneysLKP.ServiceID=Employees.ServiceID) And (PositionsSpecialJourneysLKP.CenterTypeID=Areas.CenterTypeID) And (PositionsSpecialJourneysLKP.IsActive" & oRequest("RecordType").Item & "=1) And (Employees.EmployeeNumber='" & Right(("000000" & Replace(oRequest("RecordID").Item, "'", "")), Len("000000")) & "') And (PositionsSpecialJourneysLKP.StartDate<=" & oRequest("RecordDate").Item & ") And (PositionsSpecialJourneysLKP.EndDate>=" & oRequest("RecordDate").Item & ") And (PositionsSpecialJourneysLKP.Active=1)", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
						Else
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.*, Areas.AreaID, Areas.AreaCode, Areas.AreaName, Positions.PositionID, Positions.PositionShortName, Positions.PositionName, ServiceShortName, ServiceName, ShiftShortName, ShiftName, LevelShortName, PositionsSpecialJourneysLKP.RecordID, Journeys.JourneyName From Employees, Jobs, Areas, Positions, Services, Shifts, Levels, PositionsSpecialJourneysLKP, Journeys Where (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Jobs.PositionID=Positions.PositionID) And (Employees.ServiceID=Services.ServiceID) And (Employees.ShiftID=Shifts.ShiftID) And (Employees.LevelID=Levels.LevelID) And (Journeys.JourneyID=Employees.JourneyID) And (PositionsSpecialJourneysLKP.PositionID=Positions.PositionID) And (PositionsSpecialJourneysLKP.LevelID=Employees.LevelID) And (PositionsSpecialJourneysLKP.WorkingHours=Employees.WorkingHours) And (PositionsSpecialJourneysLKP.ServiceID=Employees.ServiceID) And (PositionsSpecialJourneysLKP.CenterTypeID=Areas.CenterTypeID) And (PositionsSpecialJourneysLKP.IsActive" & oRequest("RecordType").Item & "=1) And (Employees.EmployeeNumber='" & Right(("000000" & Replace(oRequest("RecordID").Item, "'", "")), Len("000000")) & "') And (PositionsSpecialJourneysLKP.StartDate<=" & oRequest("RecordDate").Item & ") And (PositionsSpecialJourneysLKP.EndDate>=" & oRequest("RecordDate").Item & ") And (PositionsSpecialJourneysLKP.Active=1) And ((Employees.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (Jobs.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
						End If
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								sFieldID = CStr(oRecordset.Fields("EmployeeID").Value)
								sRecordID = ""
								Response.Write "<!-- " & CStr(oRecordset.Fields("RecordID").Value) & " -->" & vbNewLine
								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
									If Len(oRequest("TargetField").Item) > 0 Then
										Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ") {" & vbNewLine
											If Len(oRequest("Original").Item) = 0 Then
												sRecordID = sRecordID & "<B>Información del empleado:</B><BR /><BR />"
												Response.Write "parent.window.document." & oRequest("TargetField").Item & ".EmployeeID.value = '" & CStr(oRecordset.Fields("EmployeeID").Value) & "';" & vbNewLine
												Response.Write "parent.window.document." & oRequest("TargetField").Item & ".CheckEmployeeID.value = '1';" & vbNewLine
												Response.Write "parent.window.document." & oRequest("TargetField").Item & ".EmployeeNumber.value = '" & CStr(oRecordset.Fields("EmployeeNumber").Value) & "';" & vbNewLine
												sRecordID = sRecordID & "<B>Número del empleado: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & "<BR />"
												Response.Write "parent.window.document." & oRequest("TargetField").Item & ".EmployeeName.value = '" & CStr(oRecordset.Fields("EmployeeName").Value) & "';" & vbNewLine
												Response.Write "parent.window.document." & oRequest("TargetField").Item & ".EmployeeLastName.value = '" & CStr(oRecordset.Fields("EmployeeLastName").Value) & "';" & vbNewLine
												If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
													Response.Write "parent.window.document." & oRequest("TargetField").Item & ".EmployeeLastName2.value = '" & CStr(oRecordset.Fields("EmployeeLastName2").Value) & "';" & vbNewLine
												End If
												If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
													sRecordID = sRecordID & "<B>Nombre: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value)) & "<BR />"
												Else
													sRecordID = sRecordID & "<B>Nombre: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value)) & "<BR />"
												End If
												Response.Write "parent.window.document." & oRequest("TargetField").Item & ".RFC.value = '" & CStr(oRecordset.Fields("RFC").Value) & "';" & vbNewLine
												sRecordID = sRecordID & "<B>RFC: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)) & "<BR />"
												Response.Write "parent.window.document." & oRequest("TargetField").Item & ".CURP.value = '" & CStr(oRecordset.Fields("CURP").Value) & "';" & vbNewLine
											Else
												sRecordID = sRecordID & "<B>Información del empleado a suplir:</B><BR /><BR />"
												Response.Write "parent.window.document." & oRequest("TargetField").Item & ".OriginalEmployeeID.value = '" & CStr(oRecordset.Fields("EmployeeNumber").Value) & "';" & vbNewLine
												Response.Write "parent.window.document." & oRequest("TargetField").Item & ".CheckOriginalEmployeeID.value = '1';" & vbNewLine
											End If
											Response.Write "parent.window.document." & oRequest("TargetField").Item & ".AreaID.value = '" & CStr(oRecordset.Fields("AreaID").Value) & "';" & vbNewLine
											sRecordID = sRecordID & "<B>Adscripción: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value) & ". " & CStr(oRecordset.Fields("AreaName").Value)) & "<BR />"
											lAreaID = CLng(oRecordset.Fields("AreaID").Value)

											Response.Write "parent.window.document." & oRequest("TargetField").Item & ".ServiceID.value = '" & CStr(oRecordset.Fields("ServiceID").Value) & "';" & vbNewLine
											sRecordID = sRecordID & "<B>Servicio: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("ServiceShortName").Value) & ". " & CStr(oRecordset.Fields("ServiceName").Value)) & "<BR />"
											lServiceID = CLng(oRecordset.Fields("ServiceID").Value)

											Response.Write "parent.window.document." & oRequest("TargetField").Item & ".PositionID.value = '" & CStr(oRecordset.Fields("PositionID").Value) & "';" & vbNewLine
											sRecordID = sRecordID & "<B>Puesto: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value)) & "<BR />"
											lPositionID = CLng(oRecordset.Fields("PositionID").Value)

											Response.Write "parent.window.document." & oRequest("TargetField").Item & ".ShiftID.value = '" & CStr(oRecordset.Fields("ShiftID").Value) & "';" & vbNewLine
											sRecordID = sRecordID & "<B>Horario: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("ShiftShortName").Value) & ". " & CStr(oRecordset.Fields("ShiftName").Value)) & "<BR />"

											Response.Write "parent.window.document." & oRequest("TargetField").Item & ".LevelID.value = '" & CStr(oRecordset.Fields("LevelID").Value) & "';" & vbNewLine
											sRecordID = sRecordID & "<B>Nivel/subnivel: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value)) & "<BR />"
											lLevelID = CLng(oRecordset.Fields("LevelID").Value)

											Response.Write "parent.window.document." & oRequest("TargetField").Item & ".WorkingHours.value = '" & CStr(oRecordset.Fields("WorkingHours").Value) & "';" & vbNewLine
											Select Case CInt(oRequest("RecordType").Item)
												Case 2, 4 ' Suplencia
													If CLng(Replace(oRequest("EmployeeID").Item, "'", "")) < 800000 Then ' Interno
														Response.Write "parent.window.document." & oRequest("TargetField").Item & ".WorkedHours.value = '" & CStr(oRecordset.Fields("WorkingHours").Value) & "';" & vbNewLine
														Response.Write "parent.window.document." & oRequest("TargetField").Item & ".JourneyID.value = '" & CStr(oRecordset.Fields("JourneyID").Value) & "';" & vbNewLine
														sRecordID = sRecordID & "<B>Turno: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("JourneyID").Value) & ". " & CStr(oRecordset.Fields("ShiftName").Value)) & "<BR />"
													End If
											End Select
											sRecordID = sRecordID & "<B>Horas laboradas: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("WorkingHours").Value)) & "<BR />"
											dWorkingHours = CDbl(oRecordset.Fields("WorkingHours").Value)
											oRecordset.Close

											If (Len(oRequest("Original").Item) = 0) And (CInt(oRequest("RecordType").Item) <> 2) Then
												sErrorDescription = "No se pudo obtener la información del registro."
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct AreaID, AreaCode, AreaName From PositionsSpecialJourneysLKP, Areas Where (PositionsSpecialJourneysLKP.CenterTypeID=Areas.CenterTypeID) And (ServiceID=" & lServiceID & ") And (PositionID=" & lPositionID & ") And (LevelID=" & lLevelID & ") And (WorkingHours=" & dWorkingHours & ") And (IsActive" & CInt(oRequest("RecordType").Item) & "=1) And (PositionsSpecialJourneysLKP.StartDate<=" & oRequest("RecordDate").Item & ") And (PositionsSpecialJourneysLKP.EndDate>=" & oRequest("RecordDate").Item & ") And (Areas.StartDate<=" & oRequest("RecordDate").Item & ") And (Areas.EndDate>=" & oRequest("RecordDate").Item & ") And (PositionsSpecialJourneysLKP.Active=1) Order by AreaCode", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
												If lErrorNumber = 0 Then
													Response.Write "RemoveAllItemsFromList(null, parent.window.document." & oRequest("TargetField").Item & ".AreaID);" & vbNewLine
													Do While Not oRecordset.EOF
														Response.Write "AddItemToList('" & CStr(oRecordset.Fields("AreaCode").Value) & ". " & CStr(oRecordset.Fields("AreaName").Value) & "', '" & CStr(oRecordset.Fields("AreaID").Value) & "', null, parent.window.document." & oRequest("TargetField").Item & ".AreaID);" & vbNewLine
														oRecordset.MoveNext
														If Err.number <> 0 Then Exit Do
													Loop
													oRecordset.Close
													Response.Write "SelectItemByValue('" & lAreaID & "', false, parent.window.document." & oRequest("TargetField").Item & ".AreaID);" & vbNewLine
													If Len(oRequest("AreaID").Item) > 0 Then
														If CLng(oRequest("AreaID").Item) > -1 Then Response.Write "SelectItemByValue('" & oRequest("AreaID").Item & "', false, parent.window.document." & oRequest("TargetField").Item & ".AreaID);" & vbNewLine
													End If
												End If
											End If

											sErrorDescription = "No se pudo obtener la información del empleado."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select RiskLevels.RiskLevelID, RiskLevelName From EmployeesRisksLKP, RiskLevels Where (EmployeesRisksLKP.RiskLevel=RiskLevels.RiskLevelID) And (EmployeesRisksLKP.EmployeeID=" & sFieldID & ")", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
											If lErrorNumber = 0 Then
												If Not oRecordset.EOF Then
													Response.Write "parent.window.document." & oRequest("TargetField").Item & ".RiskLevelID.value = '" & CStr(oRecordset.Fields("RiskLevel").Value) & "';" & vbNewLine
													sRecordID = sRecordID & "<B>Riesgo laboral: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("RiskLevelName").Value)) & "<BR />"
												End If
											End If
											oRecordset.Close
										Response.Write "}" & vbNewLine
									End If
								Response.Write "//--></SCRIPT>" & vbNewLine
								Response.Write sRecordID
							Else
								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
									If Len(oRequest("TargetField").Item) > 0 Then
										Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ") {" & vbNewLine
											If Len(oRequest("Original").Item) = 0 Then
												Response.Write "parent.window.document." & oRequest("TargetField").Item & ".EmployeeID.value = '';" & vbNewLine
												Response.Write "parent.window.document." & oRequest("TargetField").Item & ".CheckEmployeeID.value = '';" & vbNewLine
											Else
												Response.Write "parent.window.document." & oRequest("TargetField").Item & ".OriginalEmployeeID.value = '';" & vbNewLine
												Response.Write "parent.window.document." & oRequest("TargetField").Item & ".CheckOriginalEmployeeID.value = '';" & vbNewLine
											End If
										Response.Write "}" & vbNewLine
									End If
								Response.Write "//--></SCRIPT>" & vbNewLine
								Select Case oRequest("RecordType").Item
									Case 1
										Response.Write "<B>El empleado no existe, no tiene permisos para acceder al centro de trabajo del usuario o las características de su puesto no permiten registrar guardias.</B>"
									Case 2
										Response.Write "<B>El empleado no existe, no tiene permisos para acceder al centro de trabajo del usuario o las características de su puesto no permiten registrar suplencias.</B>"
									Case 3
										Response.Write "<B>El empleado no existe, no tiene permisos para acceder al centro de trabajo del usuario o las características de su puesto no permiten registrar rezago quirúrgico.</B>"
									Case 4
										Response.Write "<B>El empleado no existe, no tiene permisos para acceder al centro de trabajo del usuario o las características de su puesto no permiten registrar PROVAC.</B>"
								End Select
							End If
						End If
					Case "EmployeesInfo"
						If Len(sRecordID) = 0 Then
							Response.Write "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>El empleado no está registrado en el sistema.</B></FONT>"
						Else
							Response.Write sRecordID
						End If
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							If Len(oRequest("TargetField").Item) > 0 Then
								Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ")" & vbNewLine
									Response.Write "parent.window.document." & oRequest("TargetField").Item & ".value = '" & sFieldID & "';" & vbNewLine
							End If
						Response.Write "//--></SCRIPT>" & vbNewLine
					Case "EmployeesNameFromNumber"
						If Len(sRecordID) = 0 Then
							Response.Write "&nbsp;&nbsp;&nbsp;<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>El número de empleado no existe.</B></FONT>"
						End If
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							If Len(oRequest("TargetField").Item) > 0 Then
								Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ")" & vbNewLine
									Response.Write "parent.window.document." & oRequest("TargetField").Item & ".value = '" & sRecordID & "';" & vbNewLine
							End If
						Response.Write "//--></SCRIPT>" & vbNewLine
					Case "ExternalGyS"
						sErrorDescription = "No se pudo obtener la información del empleado."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID From Employees Where (RFC='" & Replace(UCase(oRequest("RecordID").Item), "'", "") & "')", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If oRecordset.EOF Then
								sErrorDescription = "No se pudo obtener la información del empleado."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Top 1 * From EmployeesSpecialJourneys Where (RFC='" & Replace(UCase(oRequest("RecordID").Item), "'", "") & "') Order By RecordID Desc", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										If Len(Trim(oRequest("CURP").Item)) <> 0 Then
											If (StrComp(CStr(oRecordset.Fields("CURP").Value), oRequest("CURP").Item, vbBinaryCompare) <> 0) Then
												Response.Write "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>El RFC especificado tiene otra CURP registrada.</B></FONT>"
												Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
													If Len(oRequest("TargetField").Item) > 0 Then
														Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ") {" & vbNewLine
															Response.Write "parent.window.document." & oRequest("TargetField").Item & ".CheckEmployeeID.value = '';" & vbNewLine
														Response.Write "}" & vbNewLine
													End If
												Response.Write "//--></SCRIPT>" & vbNewLine
											Else
												Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
													If Len(oRequest("TargetField").Item) > 0 Then
														Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ") {" & vbNewLine
															Response.Write "parent.window.document." & oRequest("TargetField").Item & ".EmployeeID.value = '" & CStr(oRecordset.Fields("EmployeeID").Value) & "';" & vbNewLine
															Response.Write "parent.window.document." & oRequest("TargetField").Item & ".CheckEmployeeID.value = '1';" & vbNewLine
															Response.Write "parent.window.document." & oRequest("TargetField").Item & ".EmployeeNumber.value = '" & CStr(oRecordset.Fields("EmployeeNumber").Value) & "';" & vbNewLine
															Response.Write "parent.window.document." & oRequest("TargetField").Item & ".EmployeeName.value = '" & CStr(oRecordset.Fields("EmployeeName").Value) & "';" & vbNewLine
															Response.Write "parent.window.document." & oRequest("TargetField").Item & ".EmployeeLastName.value = '" & CStr(oRecordset.Fields("EmployeeLastName").Value) & "';" & vbNewLine
															If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
																Response.Write "parent.window.document." & oRequest("TargetField").Item & ".EmployeeLastName2.value = '" & CStr(oRecordset.Fields("EmployeeLastName2").Value) & "';" & vbNewLine
															End If
															Response.Write "parent.window.document." & oRequest("TargetField").Item & ".RFC.value = '" & CStr(oRecordset.Fields("RFC").Value) & "';" & vbNewLine
															Response.Write "parent.window.document." & oRequest("TargetField").Item & ".CURP.value = '" & CStr(oRecordset.Fields("CURP").Value) & "';" & vbNewLine
														Response.Write "}" & vbNewLine
													End If
												Response.Write "//--></SCRIPT>" & vbNewLine
											End If
										Else
											Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
												If Len(oRequest("TargetField").Item) > 0 Then
													Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ") {" & vbNewLine
														Response.Write "parent.window.document." & oRequest("TargetField").Item & ".EmployeeID.value = '" & CStr(oRecordset.Fields("EmployeeID").Value) & "';" & vbNewLine
														Response.Write "parent.window.document." & oRequest("TargetField").Item & ".CheckEmployeeID.value = '1';" & vbNewLine
														Response.Write "parent.window.document." & oRequest("TargetField").Item & ".EmployeeNumber.value = '" & CStr(oRecordset.Fields("EmployeeNumber").Value) & "';" & vbNewLine
														Response.Write "parent.window.document." & oRequest("TargetField").Item & ".EmployeeName.value = '" & CStr(oRecordset.Fields("EmployeeName").Value) & "';" & vbNewLine
														Response.Write "parent.window.document." & oRequest("TargetField").Item & ".EmployeeLastName.value = '" & CStr(oRecordset.Fields("EmployeeLastName").Value) & "';" & vbNewLine
														If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
															Response.Write "parent.window.document." & oRequest("TargetField").Item & ".EmployeeLastName2.value = '" & CStr(oRecordset.Fields("EmployeeLastName2").Value) & "';" & vbNewLine
														End If
														Response.Write "parent.window.document." & oRequest("TargetField").Item & ".RFC.value = '" & CStr(oRecordset.Fields("RFC").Value) & "';" & vbNewLine
														Response.Write "parent.window.document." & oRequest("TargetField").Item & ".CURP.value = '" & CStr(oRecordset.Fields("CURP").Value) & "';" & vbNewLine
													Response.Write "}" & vbNewLine
												End If
											Response.Write "//--></SCRIPT>" & vbNewLine
										End If
									Else
										Response.Write "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>El RFC o CURP no se encuentran registrados. Introduzca la información del empleado externo.</B></FONT>"
										Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
											If Len(oRequest("TargetField").Item) > 0 Then
												Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ") {" & vbNewLine
													Response.Write "parent.window.document." & oRequest("TargetField").Item & ".CheckEmployeeID.value = '1';" & vbNewLine
													Response.Write "parent.window.document." & oRequest("TargetField").Item & ".EmployeeName.value = '';" & vbNewLine
													Response.Write "parent.window.document." & oRequest("TargetField").Item & ".EmployeeLastName.value = '';" & vbNewLine
													Response.Write "parent.window.document." & oRequest("TargetField").Item & ".EmployeeLastName2.value = '';" & vbNewLine
													lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(EmployeeID) + 1 NextEmployeeID From EmployeesSpecialJourneys Where (EmployeeID >= 800000) And (EmployeeID < 900000)", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
													If lErrorNumber = 0 Then
														If IsNull(oRecordset.Fields("NextEmployeeID").Value) Or oRecordset.EOF Then
															Response.Write "parent.window.document." & oRequest("TargetField").Item & ".EmployeeID.value = '800000';" & vbNewLine
															Response.Write "parent.window.document." & oRequest("TargetField").Item & ".EmployeeNumber.value = '800000';" & vbNewLine
														Else
															Response.Write "parent.window.document." & oRequest("TargetField").Item & ".EmployeeID.value = '" & CStr(oRecordset.Fields("NextEmployeeID").Value) & "';" & vbNewLine
															Response.Write "parent.window.document." & oRequest("TargetField").Item & ".EmployeeNumber.value = '" & CStr(oRecordset.Fields("NextEmployeeID").Value) & "';" & vbNewLine
														End IF
													End If

													
												Response.Write "}" & vbNewLine
											End If
										Response.Write "//--></SCRIPT>" & vbNewLine
									End If
									oRecordset.Close
								End If
							Else
								Response.Write "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>El RFC o CURP indicado pertenece a un empleado interno.</B></FONT>"
								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
									If Len(oRequest("TargetField").Item) > 0 Then
										Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ") {" & vbNewLine
											Response.Write "parent.window.document." & oRequest("TargetField").Item & ".CheckEmployeeID.value = '';" & vbNewLine
										Response.Write "}" & vbNewLine
									End If
								Response.Write "//--></SCRIPT>" & vbNewLine
							End If
						End If
					Case "DocumentsForLicenses", "JobSwap"
						If Len(sRecordID) = 0 Then
							Response.Write "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>El empleado no está registrado en el sistema.</B></FONT>"
						Else
							Response.Write sRecordID
						End If
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							If Len(oRequest("TargetField").Item) > 0 Then
								Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ")" & vbNewLine
									Response.Write "parent.window.document." & oRequest("TargetField").Item & ".value = '" & sFieldID & "';" & vbNewLine
							End If
						Response.Write "//--></SCRIPT>" & vbNewLine
					Case "EmployeeNumber"
						If Len(sRecordID) = 0 Then
							Response.Write "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>El empleado no está registrado en el sistema.</B></FONT>"
						Else
							Response.Write sRecordID
						End If
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							If Len(oRequest("TargetField").Item) > 0 Then
								Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ")" & vbNewLine
									Response.Write "parent.window.document." & oRequest("TargetField").Item & ".value = '" & sFieldID & "';" & vbNewLine
							End If
						Response.Write "//--></SCRIPT>" & vbNewLine
					Case "EmployeeHeadNumber"
						If Len(sRecordID) = 0 Then
							Response.Write "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>El empleado no está registrado en el sistema o no es funcionario.</B></FONT>"
						Else
							Response.Write sRecordID
						End If
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							If Len(oRequest("TargetField").Item) > 0 Then
								Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ")" & vbNewLine
									Response.Write "parent.window.document." & oRequest("TargetField").Item & ".value = '" & sFieldID & "';" & vbNewLine
							End If
						Response.Write "//--></SCRIPT>" & vbNewLine
					Case "JobID", "JobNumber"
						If CLng(oRequest("lReasonID").Item) = 13 And (lErrorNumber <> 0) Then
							Response.Write "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><BR /><B>" & sErrorDescription & "</B></FONT>"
						Else
							If Len(sRecordID) = 0 Then
								Response.Write "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><BR /><B>La plaza no está registrado en el sistema.</B></FONT>"
							Else
								Response.Write sRecordID
							End If
						End If
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
						If lErrorNumber = 0 Then
							If Len(oRequest("TargetField").Item) > 0 Then
								If (Clng(oRequest("lReasonID").Item) = 13) And (len(oRequest("IsCombo").Item) = 0) Then
									lStartDate = oRequest("StartYear").Item & oRequest("StartMonth").Item & oRequest("StartDay").Item
									lEndDate = oRequest("EndYear").Item & oRequest("EndMonth").Item & oRequest("EndDay").Item
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "select StatusId, JobDate, EndDate From JobsHistoryList Where JobID = " & oRequest("RecordId").Item & " And (JobDate <= " & lStartDate & " And EndDate >= " & lEndDate & ")", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
									If Not oRecordset.EOF Then
										If (CInt(oRecordset.Fields("StatusID").Value) = 2) Or (CInt(oRecordset.Fields("StatusID").Value) = -1) Or (CInt(oRecordset.Fields("StatusID").Value) = 5) Or (CInt(oRecordset.Fields("StatusID").Value) = 4)  Or (CInt(oRecordset.Fields("StatusID").Value) = 7) Then
											Response.Write "AddItemToList('" & oRequest("RecordID").Item & "', '" & oRequest("RecordID").Item & "', null, parent.window.document." & oRequest("TargetField").Item & ");" & vbNewLine
											Response.Write "parent.window.document." & oRequest("TargetField").Item & ".value = '" & CLng(oRequest("RecordID").Item) & "';" & vbNewLine
										End If
									Else
										Response.Write "AddItemToList('" & oRequest("RecordID").Item & "', '" & oRequest("RecordID").Item & "', null, parent.window.document." & oRequest("TargetField").Item & ");" & vbNewLine
										Response.Write "parent.window.document." & oRequest("TargetField").Item & ".value = '" & CLng(oRequest("RecordID").Item) & "';" & vbNewLine
									End If
								Else
									Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ")" & vbNewLine
										Response.Write "parent.window.document." & oRequest("TargetField").Item & ".value = '" & CLng(oRequest("RecordId").Item) & "';" & vbNewLine
								End If
							End If
						End If
						Response.Write "//--></SCRIPT>" & vbNewLine
					Case "JobsInfo"
						If Len(sRecordID) = 0 Then
							Response.Write "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>La plaza no está registrada en el sistema.</B></FONT>"
						Else
							Response.Write sRecordID
						End If
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							If Len(oRequest("TargetField").Item) > 0 Then
								Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ") {" & vbNewLine
									Response.Write "parent.window.document." & oRequest("TargetField").Item & ".value = '" & CLng(sFieldID) & "';" & vbNewLine
									If Len(oRequest("SendPosition").Item) > 0 Then
										Response.Write "parent.window.document." & oRequest("TargetField").Item & ".form.PositionID.value = '" & lRecordID & "';" & vbNewLine
									End If
								Response.Write "}" & vbNewLine
							End If
						Response.Write "//--></SCRIPT>" & vbNewLine
					Case "PaperworkCatalogs"
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							Response.Write "function UpdateEstimatedDate(oList) {" & vbNewLine
								If StrComp(oRequest("StartDate").Item, "-1", vbBinaryCompare) <> 0 Then
									Response.Write "var asSubjectTypes = new Array("
										sErrorDescription = "No se pudo obtener la información del registro."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select SubjectTypeID, DaysForAttention From SubjectTypes Where (DaysForAttention>0) And (Active=1) Order By SubjectTypeID", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
										If lErrorNumber = 0 Then
											Do While Not oRecordset.EOF
												Response.Write "['" & CStr(oRecordset.Fields("SubjectTypeID").Value) & "', '" & CStr(oRecordset.Fields("DaysForAttention").Value) & "', '" & AddDaysToSerialDate(oRequest("StartDate").Item, CInt(oRecordset.Fields("DaysForAttention").Value)) & "']," & vbNewLine
												oRecordset.MoveNext
											Loop
										End If
									Response.Write "['-2', '0', '" & oRequest("StartDate").Item & "']);" & vbNewLine
								End If

								Response.Write "var oForm = parent.window.document." & oRequest("TargetField").Item & ";" & vbNewLine
								Response.Write "var sMonth = '';" & vbNewLine
								Response.Write "var sDay = '';" & vbNewLine

								Response.Write "if (oForm) {" & vbNewLine
									Response.Write "oForm.SubjectTypeID.value=oList.value;" & vbNewLine
									Response.Write "oForm.SubjectTypeName.value=GetSelectedText(oList);" & vbNewLine
									If StrComp(oRequest("StartDate").Item, "-1", vbBinaryCompare) <> 0 Then
										Response.Write "for (var i=0; i<asSubjectTypes.length; i++){" & vbNewLine
											Response.Write "if (asSubjectTypes[i][0] == oList.value) {" & vbNewLine
												Response.Write "SetDateCombos(asSubjectTypes[i][2].substr(0, 4), asSubjectTypes[i][2].substr(4, 2), asSubjectTypes[i][2].substr(6, 2), oForm.EstimatedDateYear, oForm.EstimatedDateMonth, oForm.EstimatedDateDay);" & vbNewLine
											Response.Write "}" & vbNewLine
										Response.Write "}" & vbNewLine
									End If
								Response.Write "}" & vbNewLine
							Response.Write "} // End of SelectSameItemsForOwners" & vbNewLine
						Response.Write "//--></SCRIPT>" & vbNewLine
						If Len(oRequest("SubjectTypeIDs").Item) > 0 Then
							sErrorDescription = "No se pudo obtener la información del registro."
							If IsNumeric(Replace(oRequest("RecordID").Item, "'", "")) Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From SubjectTypes Where (SubjectTypeID=" & Replace(oRequest("RecordID").Item, "'", "") & ") And (Active=1) Order By SubjectTypeName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
							Else
								If iConnectionType = ORACLE Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From SubjectTypes Where (UPPER(SubjectTypeName) Like UPPER('" & S_WILD_CHAR & Replace(Replace(oRequest("RecordID").Item, "'", ""), "'", S_WILD_CHAR) & S_WILD_CHAR & "')) And (Active=1) Order By SubjectTypeName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
								Else
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From SubjectTypes Where (SubjectTypeName Like '" & S_WILD_CHAR & Replace(Replace(oRequest("RecordID").Item, "'", ""), "'", S_WILD_CHAR) & S_WILD_CHAR & "') And (Active=1) Order By SubjectTypeName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
								End If
							End If
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									Response.Write "<SELECT NAME=""SubjectTypeID"" ID=""SubjectTypeIDCmb"" SIZE=""1"" onChange=""UpdateEstimatedDate(this)"">"
										Do While Not oRecordset.EOF
											Response.Write "<OPTION VALUE=""" & CStr(oRecordset.Fields("SubjectTypeID").Value) & """>" & CleanStringForHTML(CStr(oRecordset.Fields("SubjectTypeID").Value) & ". " & CStr(oRecordset.Fields("SubjectTypeName").Value)) & "</OPTION>"
											oRecordset.MoveNext
											If Err.number <> 0 Then Exit Do
										Loop
									Response.Write "</SELECT>"
									Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
										Response.Write "UpdateEstimatedDate(document.SearchFrm.SubjectTypeID);" & vbNewLine
									Response.Write "//--></SCRIPT>" & vbNewLine
								Else
									Response.Write "<FONT SIZE=""2"" COLOR=""#" & S_WARNING_FOR_GUI & """>Búsqueda vacía</FONT>"
								End If
							End If
						ElseIf Len(oRequest("SenderIDs").Item) > 0 Then
							sErrorDescription = "No se pudo obtener la información del registro."
							If IsNumeric(Replace(oRequest("RecordID").Item, "'", "")) Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From PaperworkSenders Where (SenderID=" & Replace(oRequest("RecordID").Item, "'", "") & ") Order By SenderName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
							Else
								If iConnectionType = ORACLE Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From PaperworkSenders Where (UPPER(SenderName) Like UPPER('" & S_WILD_CHAR & Replace(oRequest("RecordID").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')) Or (UPPER(PositionName) Like UPPER('" & S_WILD_CHAR & Replace(oRequest("RecordID").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')) Or (UPPER(EmployeeName) Like UPPER('" & S_WILD_CHAR & Replace(oRequest("RecordID").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')) Order By SenderName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
								Else
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From PaperworkSenders Where ((SenderName Like '" & S_WILD_CHAR & Replace(oRequest("RecordID").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "') Or (PositionName Like '" & S_WILD_CHAR & Replace(oRequest("RecordID").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "') Or (EmployeeName Like '" & S_WILD_CHAR & Replace(oRequest("RecordID").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')) Order By SenderName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
								End If
							End If
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									Response.Write "<SELECT NAME=""SenderID"" ID=""SnederIDCmb"" SIZE=""1"" onChange=""parent.window.document." & oRequest("TargetField").Item & ".SenderID.value=this.value; parent.window.document." & oRequest("TargetField").Item & ".SenderName.value=GetSelectedText(this);"">"
										Do While Not oRecordset.EOF
											Response.Write "<OPTION VALUE=""" & CStr(oRecordset.Fields("SenderID").Value) & """>" & CleanStringForHTML(CStr(oRecordset.Fields("SenderID").Value) & ". " & CStr(oRecordset.Fields("SenderName").Value) & ". " & CStr(oRecordset.Fields("EmployeeName").Value) & " (" & CStr(oRecordset.Fields("PositionName").Value) & ")") & "</OPTION>"
											oRecordset.MoveNext
											If Err.number <> 0 Then Exit Do
										Loop
									Response.Write "</SELECT>"
									Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
										Response.Write "parent.window.document." & oRequest("TargetField").Item & ".SenderID.value=document.SearchFrm.SenderID.value; parent.window.document." & oRequest("TargetField").Item & ".SenderName.value=GetSelectedText(document.SearchFrm.SenderID);" & vbNewLine
									Response.Write "//--></SCRIPT>" & vbNewLine
								Else
									Response.Write "<FONT SIZE=""2"" COLOR=""#" & S_WARNING_FOR_GUI & """>Búsqueda vacía</FONT>"
								End If
							End If
						ElseIf Len(oRequest("OwnerIDs").Item) > 0 Then
							sCondition = ""
							lErrorNumber = GetPaperworksOwnersForUser(sOwnerIDs, sErrorDescription)
							If InStr(1, sOwnerIDs & ",", ",-1,", vbBinaryCompare) = 0 Then sCondition = " And (OwnerID In (" & sOwnerIDs & "))"
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								sErrorDescription = "No se pudo obtener la información del registro."
								If IsNumeric(Replace(oRequest("RecordID").Item, "'", "")) Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From PaperworkOwners Where (OwnerID=" & Replace(oRequest("RecordID").Item, "'", "") & ") And (OwnerID>-1)" & sCondition & " Order By OwnerName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
								Else
									If iConnectionType = ORACLE Then
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From PaperworkOwners Where (UPPER(OwnerName) Like UPPER('" & S_WILD_CHAR & Replace(oRequest("RecordID").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')) And (OwnerID>-1)" & sCondition & " Order By OwnerName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
									Else
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From PaperworkOwners Where (OwnerName Like '" & S_WILD_CHAR & Replace(oRequest("RecordID").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "') And (OwnerID>-1)" & sCondition & " Order By OwnerName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
									End If
								End If
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										Response.Write "RemoveAllItemsFromList(null, parent.window.document." & oRequest("TargetField").Item & ".OwnerIDTemp);" & vbNewLine
										Do While Not oRecordset.EOF
											Response.Write "AddItemToList('" & CleanStringForHTML(CStr(oRecordset.Fields("OwnerID").Value) & ". " & CStr(oRecordset.Fields("OwnerName").Value) & ". Empleado: " & CStr(oRecordset.Fields("EmployeeID").Value)) & "', '" & CStr(oRecordset.Fields("OwnerID").Value) & "', null, parent.window.document." & oRequest("TargetField").Item & ".OwnerIDTemp);" & vbNewLine
											oRecordset.MoveNext
											If Err.number <> 0 Then Exit Do
										Loop
									End If
								End If
							Response.Write "//--></SCRIPT>" & vbNewLine
						End If
					Case "PositionName"
						Dim lNewJobID
						Dim lNewPositionID
						Dim lNewJobStatusID
						lStartDate = oRequest("EmployeeYear").Item & oRequest("EmployeeMonth").Item & oRequest("EmployeeDay").Item
						lEndDate = oRequest("Endyear").Item & oRequest("EndMonth").Item & oRequest("EndDay").Item
						sQuery = "Select JobId From Jobs Where JobId = " & oRequest("RecordID").Item
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								lNewJobID = oRecordset.Fields("JobID").Value
								sQuery = "Select JHL.PositionID, PositionShortName, PositionName, EmployeeID, JobDate, JHL.EndDate, JHL.StatusID From JobsHistoryList JHL, Positions Where (JHL.PositionID = Positions.PositionID) And (JobDate <= " & lStartDate & ") And (JHL.EndDate >= " & lEndDate & ") And (JHL.JobID = " & lNewJobID & ");"
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										lNewJobStatusID = CLng(oRecordset.Fields("StatusID"))
										lNewPositionID = CLng(oRecordset.Fields("PositionID").Value)
										If lNewPositionID <> -1 Then
											If lNewJobStatusID = 2 Or lNewJobStatusID = 4 Or lNewJobStatusID = 5 Then
												Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
												Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ")" & vbNewLine
												Response.Write "parent.window.document." & oRequest("TargetField").Item & ".value = '" & oRecordset.Fields("PositionID").Value & "';" & vbNewLine
												Response.Write "parent.window.document.HistoryListFrm.CheckJobID.value = '" & lNewJobID & "';" & vbNewLine
												Response.Write "//--></SCRIPT>" & vbNewLine
												Response.Write "<FONT COLOR=""#" & S_CONFIRMATION_FOR_GUI & """><B>" & oRecordset.Fields("PositionShortName").Value & " " & oRecordset.Fields("PositionName").Value & "</B></FONT>"
											Else
												sErrorDescription = "La plaza indicada no está vacante en el periodo indicado"
												lErrorNumber = -1
											End If
										Else
											sErrorDescription = "La plaza indicada no tiene puesto asociado en el periodo indicado"
											lErrorNumber = -1
										End If
									Else
										sErrorDescription = "La plaza indicada no está vacante para el periodo indicado"
										lErrorNumber = -1
									End If
								Else
									sErrorDescription = "No se pudo obtener la información del puesto"
									lErrorNumber = -1
								End If
							Else
								sErrorDescription = "La plaza indicada no existe"
								lErrorNumber = -1
							End If
						Else
							sErrorDescription = "No se pudo obtener la información de la plaza indicada"
							lErrorNumber = -1
						End If
						If lErrorNumber = -1 Then
							Response.Write "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>" & sErrorDescription & "</B></FONT>"
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								Response.Write "parent.window.document." & oRequest("TargetField").Item & ".value = '" & lErrorNumber & "';" & vbNewLine
							Response.Write "//--></SCRIPT>" & vbNewLine
						End If

					Case "PositionsByType"
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							sErrorDescription = "No se pudo obtener la información del registro."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PositionID, PositionShortName, PositionName, LevelID From Positions Where (Positions.PositionID>0) And (Positions.EndDate=30000000) And (Positions.CompanyID=1) And (Positions.PositionTypeID=" & oRequest("RecordID").Item & ") And (Positions.Active=1) Order By PositionShortName, PositionName, LevelID", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									Response.Write "RemoveAllItemsFromList(null, parent.window.document." & oRequest("TargetField").Item & ");" & vbNewLine
									Do While Not oRecordset.EOF
										Response.Write "AddItemToList('" & CStr(oRecordset.Fields("PositionShortName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value) & " (Nivel " & CStr(oRecordset.Fields("LevelID").Value) & ")', '" & CStr(oRecordset.Fields("PositionID").Value) & "', null, parent.window.document." & oRequest("TargetField").Item & ");" & vbNewLine
										oRecordset.MoveNext
										If Err.number <> 0 Then Exit Do
									Loop
								End If
								oRecordset.Close
							End If
						Response.Write "//--></SCRIPT>" & vbNewLine
					Case "PositionsCatalogsLKP"
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							sErrorDescription = "No se pudo obtener la información del registro."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeTypes.EmployeeTypeID, EmployeeTypeShortName, EmployeeTypeName, PositionTypes.PositionTypeID, PositionTypeShortName, PositionTypeName, Companies.CompanyID, CompanyShortName, CompanyName, ClassificationID, GroupGradeLevels.GroupGradeLevelID, GroupGradeLevelName, IntegrationID, Levels.LevelID, LevelName, WorkingHours From Positions, EmployeeTypes, PositionTypes, Companies, GroupGradeLevels, Levels Where (Positions.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Positions.PositionTypeID=PositionTypes.PositionTypeID) And (Positions.CompanyID=Companies.CompanyID) And (Positions.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (Positions.LevelID=Levels.LevelID) And (PositionID=" & Replace(oRequest("RecordID").Item, "'", "") & ")", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									sFieldID = CLng(oRecordset.Fields("EmployeeTypeID").Value)
									Response.Write "RemoveAllItemsFromList(null, parent.window.document." & oRequest("TargetField").Item & ".EmployeeTypeID);" & vbNewLine
									Response.Write "AddItemToList('" & CStr(oRecordset.Fields("EmployeeTypeShortName").Value) & ". " & CStr(oRecordset.Fields("EmployeeTypeName").Value) & "', '" & CStr(oRecordset.Fields("EmployeeTypeID").Value) & "', null, parent.window.document." & oRequest("TargetField").Item & ".EmployeeTypeID);" & vbNewLine
									'Response.Write "parent.window.document." & oRequest("TargetField").Item & ".PositionTypeID.value = '" & CStr(oRecordset.Fields("PositionTypeID").Value) & "';" & vbNewLine
									'Response.Write "parent.window.document." & oRequest("TargetField").Item & ".PositionTypeName.value = '" & CStr(oRecordset.Fields("PositionTypeShortName").Value) & ". " & CStr(oRecordset.Fields("PositionTypeName").Value) & "';" & vbNewLine
									Response.Write "RemoveAllItemsFromList(null, parent.window.document." & oRequest("TargetField").Item & ".CompanyID);" & vbNewLine
									Response.Write "AddItemToList('" & CStr(oRecordset.Fields("CompanyShortName").Value) & ". " & CStr(oRecordset.Fields("CompanyName").Value) & "', '" & CStr(oRecordset.Fields("CompanyID").Value) & "', null, parent.window.document." & oRequest("TargetField").Item & ".CompanyID);" & vbNewLine
									Response.Write "RemoveAllItemsFromList(null, parent.window.document." & oRequest("TargetField").Item & ".GroupGradeLevelID);" & vbNewLine
									Response.Write "AddItemToList('" & CStr(oRecordset.Fields("GroupGradeLevelName").Value) & "', '" & CStr(oRecordset.Fields("GroupGradeLevelID").Value) & "', null, parent.window.document." & oRequest("TargetField").Item & ".GroupGradeLevelID);" & vbNewLine
									Response.Write "RemoveAllItemsFromList(null, parent.window.document." & oRequest("TargetField").Item & ".LevelID);" & vbNewLine
									Response.Write "AddItemToList('" & CStr(oRecordset.Fields("LevelName").Value) & "', '" & CStr(oRecordset.Fields("LevelID").Value) & "', null, parent.window.document." & oRequest("TargetField").Item & ".LevelID);" & vbNewLine
									Response.Write "parent.window.document." & oRequest("TargetField").Item & ".ClassificationID.value = '" & CStr(oRecordset.Fields("ClassificationID").Value) & "';" & vbNewLine
									Response.Write "parent.window.document." & oRequest("TargetField").Item & ".IntegrationID.value = '" & CStr(oRecordset.Fields("IntegrationID").Value) & "';" & vbNewLine
									Response.Write "parent.window.document." & oRequest("TargetField").Item & ".WorkingHours.value = '" & CStr(oRecordset.Fields("WorkingHours").Value) & "';" & vbNewLine
								End If
								oRecordset.Close
							End If

							'Response.Write "RemoveAllItemsFromList(null, parent.window.document." & oRequest("TargetField").Item & ".CompanyID);" & vbNewLine
							'sErrorDescription = "No se pudo obtener la información del registro."
							'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CompanyID, CompanyShortName, CompanyName From PositionsCatalogsLKP, Companies Where (PositionsCatalogsLKP.RecordID=Companies.CompanyID) And (PositionID=" & Replace(oRequest("RecordID").Item, "'", "") & ") And (CatalogID=2) And (Companies.EndDate=30000000) Order By CompanyShortName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
							'If lErrorNumber = 0 Then
							'	If oRecordset.EOF Then
							'		oRecordset.Close
							'		sErrorDescription = "No se pudo obtener la información del registro."
							'		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CompanyID, CompanyShortName, CompanyName From Companies Where (Companies.EndDate=30000000) Order By CompanyShortName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
							'	End If
							'End If
							'If lErrorNumber = 0 Then
							'	Do While Not oRecordset.EOF
							'		Response.Write "AddItemToList('" & CStr(oRecordset.Fields("CompanyShortName").Value) & ". " & CStr(oRecordset.Fields("CompanyName").Value) & "', '" & CStr(oRecordset.Fields("CompanyID").Value) & "', null, parent.window.document." & oRequest("TargetField").Item & ".CompanyID);" & vbNewLine
							'		oRecordset.MoveNext
							'		If Err.number <> 0 Then Exit Do
							'	Loop
							'	oRecordset.Close
							'End If

							'Response.Write "RemoveAllItemsFromList(null, parent.window.document." & oRequest("TargetField").Item & ".GroupGradeLevelID);" & vbNewLine
							'If sFieldID = 1 Then
							'	sErrorDescription = "No se pudo obtener la información del registro."
							'	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select GroupGradeLevelID, GroupGradeLevelName From PositionsCatalogsLKP, GroupGradeLevels Where (PositionsCatalogsLKP.RecordID=GroupGradeLevels.GroupGradeLevelID) And (PositionID=" & Replace(oRequest("RecordID").Item, "'", "") & ") And (CatalogID=3) And (GroupGradeLevels.EndDate=30000000) Order By GroupGradeLevelName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
							'	If lErrorNumber = 0 Then
							'		If oRecordset.EOF Then
							'			oRecordset.Close
							'			sErrorDescription = "No se pudo obtener la información del registro."
							'			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select GroupGradeLevelID, GroupGradeLevelName From GroupGradeLevels Where (GroupGradeLevels.EndDate=30000000) Order By GroupGradeLevelName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
							'		End If
							'	End If
							'	If lErrorNumber = 0 Then
							'		Do While Not oRecordset.EOF
							'			Response.Write "AddItemToList('" & CStr(oRecordset.Fields("GroupGradeLevelName").Value) & "', '" & CStr(oRecordset.Fields("GroupGradeLevelID").Value) & "', null, parent.window.document." & oRequest("TargetField").Item & ".GroupGradeLevelID);" & vbNewLine
							'			oRecordset.MoveNext
							'			If Err.number <> 0 Then Exit Do
							'		Loop
							'		oRecordset.Close
							'	End If
							'Else
							'	Response.Write "AddItemToList('-1', '-1', null, parent.window.document." & oRequest("TargetField").Item & ".GroupGradeLevelID);" & vbNewLine
							'End If

							'Response.Write "RemoveAllItemsFromList(null, parent.window.document." & oRequest("TargetField").Item & ".JourneyID);" & vbNewLine
							'sErrorDescription = "No se pudo obtener la información del registro."
							'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JourneyID, JourneyShortName, JourneyName From PositionsCatalogsLKP, Journeys Where (PositionsCatalogsLKP.RecordID=Journeys.JourneyID) And (PositionID=" & Replace(oRequest("RecordID").Item, "'", "") & ") And (CatalogID=4) And (Journeys.EndDate=30000000) Order By JourneyShortName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
							'If lErrorNumber = 0 Then
							'	If oRecordset.EOF Then
							'		oRecordset.Close
							'		sErrorDescription = "No se pudo obtener la información del registro."
							'		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JourneyID, JourneyShortName, JourneyName From Journeys Where (Journeys.EndDate=30000000) Order By JourneyShortName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
							'	End If
							'End If
							'If lErrorNumber = 0 Then
							'	Do While Not oRecordset.EOF
							'		Response.Write "AddItemToList('" & CStr(oRecordset.Fields("JourneyShortName").Value) & ". " & CStr(oRecordset.Fields("JourneyName").Value) & "', '" & CStr(oRecordset.Fields("JourneyID").Value) & "', null, parent.window.document." & oRequest("TargetField").Item & ".JourneyID);" & vbNewLine
							'		oRecordset.MoveNext
							'		If Err.number <> 0 Then Exit Do
							'	Loop
							'	oRecordset.Close
							'End If

							'Response.Write "RemoveAllItemsFromList(null, parent.window.document." & oRequest("TargetField").Item & ".LevelID);" & vbNewLine
							'If sFieldID <> 1 Then
							'	sErrorDescription = "No se pudo obtener la información del registro."
							'	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select LevelID, LevelShortName From PositionsCatalogsLKP, Levels Where (PositionsCatalogsLKP.RecordID=Levels.LevelID) And (PositionID=" & Replace(oRequest("RecordID").Item, "'", "") & ") And (CatalogID=5) And (Levels.EndDate=30000000) Order By LevelShortName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
							'	If lErrorNumber = 0 Then
							'		If oRecordset.EOF Then
							'			oRecordset.Close
							'			sErrorDescription = "No se pudo obtener la información del registro."
							'			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select LevelID, LevelShortName  From Levels Where (Levels.EndDate=30000000) Order By LevelShortName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
							'		End If
							'	End If
							'	If lErrorNumber = 0 Then
							'		Do While Not oRecordset.EOF
							'			Response.Write "AddItemToList('" & CStr(oRecordset.Fields("LevelShortName").Value) & "', '" & CStr(oRecordset.Fields("LevelID").Value) & "', null, parent.window.document." & oRequest("TargetField").Item & ".LevelID);" & vbNewLine
							'			oRecordset.MoveNext
							'			If Err.number <> 0 Then Exit Do
							'		Loop
							'		oRecordset.Close
							'	End If
							'Else
							'	Response.Write "AddItemToList('-1', '-1', null, parent.window.document." & oRequest("TargetField").Item & ".LevelID);" & vbNewLine
							'End If
							Response.Write "parent.window.ShowJobFields();" & vbNewLine
						Response.Write "//--></SCRIPT>" & vbNewLine
					Case "PositionsForEmployeeType"
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							sErrorDescription = "No se pudo obtener la información del registro."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PositionID, PositionShortName, PositionName From Positions Where (EmployeeTypeID=" & Replace(oRequest("RecordID").Item, "'", "") & ") Order By PositionShortName, PositionName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									Response.Write "RemoveAllItemsFromList(null, parent.window.document." & oRequest("TargetField").Item & ");" & vbNewLine
									Do While Not oRecordset.EOF
										Response.Write "AddItemToList('" & CStr(oRecordset.Fields("PositionShortName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value) & "', '" & CStr(oRecordset.Fields("PositionID").Value) & "', null, parent.window.document." & oRequest("TargetField").Item & ");" & vbNewLine
										oRecordset.MoveNext
										If Err.number <> 0 Then Exit Do
									Loop
								End If
								oRecordset.Close
							End If
						Response.Write "//--></SCRIPT>" & vbNewLine
					Case "PositionsGyS"
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ") {" & vbNewLine
								sCondition = ""
								If Len(oRequest("PositionID").Item) > 0 Then
									If StrComp(oRequest("PositionID").Item, -1, vbBinaryCompare) <> 0 Then sCondition = sCondition & " And (PositionsSpecialJourneysLKP.PositionID=" & oRequest("PositionID").Item & ")"
								End If
								If Len(oRequest("AreaID").Item) > 0 Then
									If StrComp(oRequest("AreaID").Item, -1, vbBinaryCompare) <> 0 Then sCondition = sCondition & " And (Areas.AreaID=" & oRequest("AreaID").Item & ")"
								End If
								sCondition = sCondition & sAreaCondition
								If Len(oRequest("ServiceID").Item) > 0 Then
									If StrComp(oRequest("ServiceID").Item, -1, vbBinaryCompare) <> 0 Then sCondition = sCondition & " And (PositionsSpecialJourneysLKP.ServiceID=" & oRequest("ServiceID").Item & ")"
								End If
								If Len(oRequest("LevelID").Item) > 0 Then
									If StrComp(oRequest("LevelID").Item, -1, vbBinaryCompare) <> 0 Then sCondition = sCondition & " And (PositionsSpecialJourneysLKP.LevelID=" & oRequest("LevelID").Item & ")"
								End If
								If Len(oRequest("WorkingHours").Item) > 0 Then
									If StrComp(oRequest("WorkingHours").Item, -1, vbBinaryCompare) <> 0 Then sCondition = sCondition & " And (PositionsSpecialJourneysLKP.WorkingHours=" & oRequest("WorkingHours").Item & ")"
								End If

								If CInt(oRequest("RecordType").Item) <> 3 Then
									If (Len(oRequest("AreaID").Item) = 0) Or (StrComp(oRequest("AreaID").Item, -1, vbBinaryCompare) = 0) Then
										sErrorDescription = "No se pudo obtener la información del registro."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct AreaID, AreaCode, AreaName, Zones1.ZoneName As ZoneName1, Zones2.ZoneName As ZoneName2, Zones3.ZoneName As ZoneName3 From PositionsSpecialJourneysLKP, Areas, Zones As Zones1, Zones As Zones2, Zones As Zones3 Where (PositionsSpecialJourneysLKP.CenterTypeID=Areas.CenterTypeID) And (Areas.ZoneID=Zones3.ZoneID) And (Zones3.ParentID=Zones2.ZoneID) And (Zones2.ParentID=Zones1.ZoneID) " & sCondition & " And (IsActive" & CInt(oRequest("RecordType").Item) & "=1) And (PositionsSpecialJourneysLKP.StartDate<=" & oRequest("RecordDate").Item & ") And (PositionsSpecialJourneysLKP.EndDate>=" & oRequest("RecordDate").Item & ") And (Areas.StartDate<=" & oRequest("RecordDate").Item & ") And (Areas.EndDate>=" & oRequest("RecordDate").Item & ") And (PositionsSpecialJourneysLKP.Active=1) Order by AreaCode", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
										If lErrorNumber = 0 Then
											Response.Write "RemoveAllItemsFromList(null, parent.window.document." & oRequest("TargetField").Item & ".AreaID);" & vbNewLine
											Do While Not oRecordset.EOF
												Response.Write "AddItemToList('" & CStr(oRecordset.Fields("AreaCode").Value) & ". " & CStr(oRecordset.Fields("AreaName").Value) & " (" & CStr(oRecordset.Fields("ZoneName2").Value) & ", " & CStr(oRecordset.Fields("ZoneName1").Value) & ")', '" & CStr(oRecordset.Fields("AreaID").Value) & "', null, parent.window.document." & oRequest("TargetField").Item & ".AreaID);" & vbNewLine
												oRecordset.MoveNext
												If Err.number <> 0 Then Exit Do
											Loop
											oRecordset.Close
										End If
									End If
								End If

								If (Len(oRequest("ServiceID").Item) = 0) Or (StrComp(oRequest("ServiceID").Item, -1, vbBinaryCompare) = 0) Then
									sErrorDescription = "No se pudo obtener la información del registro."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct Services.ServiceID, ServiceShortName, ServiceName From PositionsSpecialJourneysLKP, Areas, Services Where (PositionsSpecialJourneysLKP.CenterTypeID=Areas.CenterTypeID) And (PositionsSpecialJourneysLKP.ServiceID=Services.ServiceID) " & sCondition & " And (IsActive" & CInt(oRequest("RecordType").Item) & "=1) And (PositionsSpecialJourneysLKP.StartDate<=" & oRequest("RecordDate").Item & ") And (PositionsSpecialJourneysLKP.EndDate>=" & oRequest("RecordDate").Item & ") And (Services.StartDate<=" & oRequest("RecordDate").Item & ") And (Services.EndDate>=" & oRequest("RecordDate").Item & ") And (PositionsSpecialJourneysLKP.Active=1) Order by ServiceShortName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
									If lErrorNumber = 0 Then
										Response.Write "RemoveAllItemsFromList(null, parent.window.document." & oRequest("TargetField").Item & ".ServiceID);" & vbNewLine
										Do While Not oRecordset.EOF
											Response.Write "AddItemToList('" & CStr(oRecordset.Fields("ServiceShortName").Value) & ". " & CStr(oRecordset.Fields("ServiceName").Value) & "', '" & CStr(oRecordset.Fields("ServiceID").Value) & "', null, parent.window.document." & oRequest("TargetField").Item & ".ServiceID);" & vbNewLine
											oRecordset.MoveNext
											If Err.number <> 0 Then Exit Do
										Loop
										oRecordset.Close
									End If
								End If

								If (Len(oRequest("LevelID").Item) = 0) Or (StrComp(oRequest("LevelID").Item, -1, vbBinaryCompare) = 0) Then
									sErrorDescription = "No se pudo obtener la información del registro."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct Levels.LevelID, LevelShortName From PositionsSpecialJourneysLKP, Areas, Levels Where (PositionsSpecialJourneysLKP.CenterTypeID=Areas.CenterTypeID) And (PositionsSpecialJourneysLKP.LevelID=Levels.LevelID) " & sCondition & " And (IsActive" & CInt(oRequest("RecordType").Item) & "=1) And (PositionsSpecialJourneysLKP.StartDate<=" & oRequest("RecordDate").Item & ") And (PositionsSpecialJourneysLKP.EndDate>=" & oRequest("RecordDate").Item & ") And (Levels.StartDate<=" & oRequest("RecordDate").Item & ") And (Levels.EndDate>=" & oRequest("RecordDate").Item & ") And (PositionsSpecialJourneysLKP.Active=1) Order by LevelShortName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
									If lErrorNumber = 0 Then
										Response.Write "RemoveAllItemsFromList(null, parent.window.document." & oRequest("TargetField").Item & ".LevelID);" & vbNewLine
										Do While Not oRecordset.EOF
											Response.Write "AddItemToList('" & CStr(oRecordset.Fields("LevelShortName").Value) & "', '" & CStr(oRecordset.Fields("LevelID").Value) & "', null, parent.window.document." & oRequest("TargetField").Item & ".LevelID);" & vbNewLine
											oRecordset.MoveNext
											If Err.number <> 0 Then Exit Do
										Loop
										oRecordset.Close
									End If
								End If

								sErrorDescription = "No se pudo obtener la información del registro."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct WorkingHours From PositionsSpecialJourneysLKP, Areas Where (PositionsSpecialJourneysLKP.CenterTypeID=Areas.CenterTypeID) " & sCondition & " And (PositionsSpecialJourneysLKP.StartDate<=" & oRequest("RecordDate").Item & ") And (PositionsSpecialJourneysLKP.EndDate>=" & oRequest("RecordDate").Item & ") And (IsActive" & CInt(oRequest("RecordType").Item) & "=1) And (PositionsSpecialJourneysLKP.Active=1) Order by WorkingHours", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									Response.Write "RemoveAllItemsFromList(null, parent.window.document." & oRequest("TargetField").Item & ".WorkingHours);" & vbNewLine
									Do While Not oRecordset.EOF
										Response.Write "AddItemToList('" & CStr(oRecordset.Fields("WorkingHours").Value) & "', '" & CStr(oRecordset.Fields("WorkingHours").Value) & "', null, parent.window.document." & oRequest("TargetField").Item & ".WorkingHours);" & vbNewLine
										oRecordset.MoveNext
										If Err.number <> 0 Then Exit Do
									Loop
									oRecordset.Close
								End If
							Response.Write "}" & vbNewLine
						Response.Write "//--></SCRIPT>" & vbNewLine
					Case "RiskLevel"
						If lErrorNumber <> 0 Then
							Response.Write "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><BR /><B>" & sErrorDescription & "</B></FONT>"
						Else
							If Not oRecordset.EOF Then
								If CInt(oRequest("RiskLevel").Item) = lRiskLevel Then
									Response.Write "<FONT COLOR=""#" & S_CONFIRMATION_FOR_GUI & """><BR /><B>El nivel de riesgo indicado se apega a la matriz de riesgos</B></FONT>"
								Else
									Response.Write "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><BR /><B>El nivel de riesgo indicado no se apega a la matriz de riesgos</B></FONT>"
								End If
							Else
								If CInt(oRequest("RecordID").Item) = 0 Then
									Response.Write "<FONT COLOR=""#" & S_CONFIRMATION_FOR_GUI & """><BR /><B>El nivel de riesgo indicado se apega a la matriz de riesgos</B></FONT>"
								Else
									Response.Write "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><BR /><B>El nivel de riesgo indicado no se apega a la matriz de riesgos</B></FONT>"
								End If
							End If
						End If
						oRecordset.Close			
					Case "RecordsForGyS"
						Dim lStartDateForRecord
						Dim lEndDateForRecord
						Dim lLimit
						lStartDate = CLng(oRequest("StartDate").Item)
						lEndDate = CLng(oRequest("EndDate").Item)
						If CInt(Mid(oRequest("PayrollDate").Item, Len("YYYYM"), Len("MM"))) > 3 Then
							lLimit = 300
						Else
							lLimit = 9100
						End If
						If lStartDate > CLng(oRequest("PayrollDate").Item) Then
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								Response.Write "alert('No se pueden agregar registros posteriores a la fecha de la quincena de aplicación.');" & vbNewLine
								Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ")" & vbNewLine
									Response.Write "parent.window.document." & oRequest("TargetField").Item & ".StartDateDay.focus();" & vbNewLine
							Response.Write "//--></SCRIPT>" & vbNewLine
						ElseIf lStartDate < (CLng(oRequest("PayrollDate").Item) - lLimit) Then
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								Response.Write "alert('No se pueden agregar registros anteriores a 3 meses.');" & vbNewLine
								Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ")" & vbNewLine
									Response.Write "parent.window.document." & oRequest("TargetField").Item & ".StartDateDay.focus();" & vbNewLine
							Response.Write "//--></SCRIPT>" & vbNewLine
						ElseIf lStartDate > lEndDate Then
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								Response.Write "alert('La fecha de inicio es posterior a la fecha final.');" & vbNewLine
								Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ")" & vbNewLine
									Response.Write "parent.window.document." & oRequest("TargetField").Item & ".StartDateDay.focus();" & vbNewLine
							Response.Write "//--></SCRIPT>" & vbNewLine
						Else
							If CInt(Right(lStartDate, Len("00"))) <= 15 Then
								lStartDate = CLng(Left(lStartDate, Len("000000")) & "01")
								lEndDate = CLng(Left(lStartDate, Len("000000")) & "15")
							Else
								lStartDate = CLng(Left(lStartDate, Len("000000")) & "16")
								If InStr(1, ",01,03,05,07,08,10,12,", "," & Mid(lStartDate, Len("00000"), Len("00")) & ",", vbBinaryCompare) > 0 Then
									lEndDate = CLng(Left(lStartDate, Len("000000")) & "31")
								ElseIf InStr(1, ",04,06,09,11,", "," & Mid(lStartDate, Len("00000"), Len("00")) & ",", vbBinaryCompare) > 0 Then
									lEndDate = CLng(Left(lStartDate, Len("000000")) & "30")
								Else
									If (CInt(Left(lStartDate, Len("0000"))) Mod 4) = 0 Then
										lEndDate = CLng(Left(lStartDate, Len("000000")) & "29")
									Else
										lEndDate = CLng(Left(lStartDate, Len("000000")) & "28")
									End If
								End If
							End If
							If CInt(oRequest("RecordType").Item) = 3 Then
								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
									Response.Write "var oTempForm01 = parent.window.document." & oRequest("TargetField").Item & ";" & vbNewLine
									Response.Write "if (oTempForm01) {" & vbNewLine
										Response.Write "oTempForm01.StartDateYear.value = '" & Left(lStartDate, Len("0000")) & "';" & vbNewLine
										Response.Write "oTempForm01.StartDateMonth.value = '" & Mid(lStartDate, Len("00000"), Len("00")) & "';" & vbNewLine
										Response.Write "oTempForm01.StartDateDay.value = '" & Right(lStartDate, Len("00")) & "';" & vbNewLine
										Response.Write "oTempForm01.EndDateYear.value = '" & Left(lEndDate, Len("0000")) & "';" & vbNewLine
										Response.Write "oTempForm01.EndDateMonth.value = '" & Mid(lEndDate, Len("00000"), Len("00")) & "';" & vbNewLine
										Response.Write "ChangeDaysListGivenTheMonthAndYear(oTempForm01.EndDateMonth.options[oTempForm01.EndDateMonth.selectedIndex].value, oTempForm01.EndDateYear.options[oTempForm01.EndDateYear.selectedIndex].value, oTempForm01.EndDateDay);" & vbNewLine
										Response.Write "window.setTimeout('oTempForm01.EndDateDay.value = \'" & Right(lEndDate, Len("00")) & "\'', 500);" & vbNewLine
										Response.Write "oTempForm01.EndDateDay.value = '" & Right(lEndDate, Len("00")) & "';" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "//--></SCRIPT>" & vbNewLine
							End If

							sRecordID = ""
							dCounter = Split("0,0,0", ",")
							dCounter(0) = 0
							dCounter(1) = 0
							dCounter(2) = 0
							If CInt(oRequest("RecordType").Item) = 3 Then sCondition = " And (SpecialJourneyID=425)"
							sErrorDescription = "No se pudo obtener la información del empleado."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesSpecialJourneys Where (RecordID<>" & oRequest("TheRecordID").Item & ") And (EmployeeID=" & Replace(oRequest("RecordID").Item, "'", "") & ") And (StartDate>=" & lStartDate & ") And (EndDate<=" & lEndDate & ") " & sCondition & " Order By StartDate", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select * From EmployeesSpecialJourneys Where (RecordID<>" & oRequest("TheRecordID").Item & ") And (EmployeeID=" & Replace(oRequest("RecordID").Item, "'", "") & ") And (StartDate>=" & lStartDate & ") And (EndDate<=" & lEndDate & ") " & sCondition & " Order By StartDate -->" & vbNewLine
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									sFieldID = CStr(oRecordset.Fields("EmployeeID").Value)
									bEmpty = False
									Do While Not oRecordset.EOF
										If (CLng(oRequest("StartDate").Item) <= CLng(oRecordset.Fields("StartDate").Value)) And (CLng(oRequest("EndDate").Item) >= CLng(oRecordset.Fields("StartDate").Value)) Then bEmpty = True
										If (CLng(oRequest("StartDate").Item) <= CLng(oRecordset.Fields("EndDate").Value)) And (CLng(oRequest("EndDate").Item) >= CLng(oRecordset.Fields("EndDate").Value)) Then bEmpty = True

										lStartDateForRecord = CLng(oRecordset.Fields("StartDate").Value)
										If lStartDateForRecord < lStartDate Then lStartDateForRecord = lStartDate
										lEndDateForRecord = CLng(oRecordset.Fields("EndDate").Value)
										If lEndDateForRecord > lEndDate Then lStartDateForRecord = lEndDate
										Select Case CInt(oRecordset.Fields("SpecialJourneyID").Value)
											Case 423 'Guardias
												dCounter(0) = dCounter(0) + (CDbl(oRecordset.Fields("WorkedHours").Value) * (Abs(DateDiff("d", GetDateFromSerialNumber(lStartDateForRecord), GetDateFromSerialNumber(lEndDateForRecord))) + 1))
											Case 424 'Suplencias
												dCounter(0) = dCounter(0) + (CDbl(oRecordset.Fields("WorkingHours").Value) * (Abs(DateDiff("d", GetDateFromSerialNumber(lStartDateForRecord), GetDateFromSerialNumber(lEndDateForRecord))) + 1))
											Case 425 'RQ
											Case 426 'PROVAC
												dCounter(0) = dCounter(0) + (CDbl(oRecordset.Fields("WorkingHours").Value) * (Abs(DateDiff("d", GetDateFromSerialNumber(lStartDateForRecord), GetDateFromSerialNumber(lEndDateForRecord))) + 1))
										End Select
										oRecordset.MoveNext
										If Err.number <> 0 Then Exit Do
									Loop
									oRecordset.Close
								End If
							End If

							lLimit = Split("0,0,0", ",")
							sErrorDescription = "No se pudo obtener la información del empleado."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From SpecialJourneysFactors Where (SpecialJourneysFactors.FactorID=" & oRequest("MovementID").Item & ")", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select * From SpecialJourneysFactors Where (SpecialJourneysFactors.FactorID=" & oRequest("MovementID").Item & ") -->" & vbNewLine
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									lLimit(0) = CDbl(oRecordset.Fields("MaxHoursPerDay").Value)
									lLimit(1) = CDbl(oRecordset.Fields("MaxDaysPerPayroll").Value)
									lLimit(2) = CDbl(oRecordset.Fields("MaxHoursPerPayroll").Value)
								End If
							End If

							sErrorDescription = "No se pudo obtener la información del empleado."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select WorkingHours, WorkedHours, StartDate, EndDate From EmployeesSpecialJourneys Where (RecordID<>" & oRequest("TheRecordID").Item & ") And (EmployeeID=" & Replace(oRequest("RecordID").Item, "'", "") & ") And (EmployeesSpecialJourneys.MovementID=" & oRequest("MovementID").Item & ") And (StartDate>=" & lStartDate & ") And (EndDate<=" & lEndDate & ") Order By StartDate", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
							Response.Write vbNewLine & "<!-- Query: Select WorkingHours, WorkedHours, StartDate, EndDate From EmployeesSpecialJourneys Where (RecordID<>" & oRequest("TheRecordID").Item & ") And (EmployeeID=" & Replace(oRequest("RecordID").Item, "'", "") & ") And (EmployeesSpecialJourneys.MovementID=" & oRequest("MovementID").Item & ") And (StartDate>=" & lStartDate & ") And (EndDate<=" & lEndDate & ") Order By StartDate -->" & vbNewLine
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									Do While Not oRecordset.EOF
										lStartDateForRecord = CLng(oRecordset.Fields("StartDate").Value)
										If lStartDateForRecord < lStartDate Then lStartDateForRecord = lStartDate
										lEndDateForRecord = CLng(oRecordset.Fields("EndDate").Value)
										If lEndDateForRecord > lEndDate Then lStartDateForRecord = lEndDate
										dCounter(1) = dCounter(1) + (Abs(DateDiff("d", GetDateFromSerialNumber(lStartDateForRecord), GetDateFromSerialNumber(lEndDateForRecord))) + 1)
										Select Case CInt(oRequest("RecordType").Item)
											Case 1 'Guardias
												dCounter(2) = dCounter(2) + (CDbl(oRecordset.Fields("WorkedHours").Value) * (Abs(DateDiff("d", GetDateFromSerialNumber(lStartDateForRecord), GetDateFromSerialNumber(lEndDateForRecord))) + 1))
											Case 2 'Suplencias
												dCounter(2) = dCounter(2) + (CDbl(oRecordset.Fields("WorkingHours").Value) * (Abs(DateDiff("d", GetDateFromSerialNumber(lStartDateForRecord), GetDateFromSerialNumber(lEndDateForRecord))) + 1))
											Case 3 'RQ
											Case 4 'PROVAC
										End Select
										oRecordset.MoveNext
										If Err.number <> 0 Then Exit Do
									Loop
									oRecordset.Close
								End If
							End If

							lStartDateForRecord = CLng(oRequest("StartDate").Item)
							lEndDateForRecord = CLng(oRequest("EndDate").Item)
							dCounter(1) = dCounter(1) + (Abs(DateDiff("d", GetDateFromSerialNumber(lStartDateForRecord), GetDateFromSerialNumber(lEndDateForRecord))) + 1)
							Select Case CInt(oRequest("RecordType").Item)
								Case 1 'Guardias
									dTotal = CDbl(oRequest("WorkedHours").Item) * (Abs(DateDiff("d", GetDateFromSerialNumber(lStartDateForRecord), GetDateFromSerialNumber(lEndDateForRecord))) + 1)
								Case 2 'Suplencias
									dTotal = CDbl(oRequest("WorkingHours").Item) * (Abs(DateDiff("d", GetDateFromSerialNumber(lStartDateForRecord), GetDateFromSerialNumber(lEndDateForRecord))) + 1)
								Case 3 'RQ
								Case 4 'PROVAC
									dTotal = CDbl(oRequest("WorkingHours").Item) * (Abs(DateDiff("d", GetDateFromSerialNumber(lStartDateForRecord), GetDateFromSerialNumber(lEndDateForRecord))) + 1)
							End Select
							dTotal = CDbl(oRequest("WorkedHours").Item)
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								If Len(oRequest("TargetField").Item) > 0 Then
									Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ") {" & vbNewLine
										If CInt(oRequest("RecordType").Item) = 3 Then
											Response.Write "parent.window.document." & oRequest("TargetField").Item & ".TempStartDate.value='" & lStartDate & "';" & vbNewLine
											Response.Write "parent.window.document." & oRequest("TargetField").Item & ".TempEndDate.value='" & lEndDate & "';" & vbNewLine
										Else
											Response.Write "parent.window.document." & oRequest("TargetField").Item & ".TempStartDate.value='" & lStartDateForRecord & "';" & vbNewLine
											Response.Write "parent.window.document." & oRequest("TargetField").Item & ".TempEndDate.value='" & lEndDateForRecord & "';" & vbNewLine
										End If
										If CInt(oRequest("RecordType").Item) = 3 Then
											If bEmpty Then
												sRecordID = "<BR /><IMG SRC=""Images/IcnErrorLevel2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" VSPACE=""3"" /><FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>Ya existen registros en las fechas indicadas.</B></FONT>"
												Response.Write "parent.window.document." & oRequest("TargetField").Item & ".ReportedHours.value='-2';" & vbNewLine
											Else
												sRecordID = "<BR />Información correcta, no existen conflictos con las fechas."
												Response.Write "parent.window.document." & oRequest("TargetField").Item & ".ReportedHours.value='1';" & vbNewLine
												bEmpty = True
											End If
										Else
											If False Then
											'If bEmpty Then
											'	sRecordID = "<BR /><IMG SRC=""Images/IcnErrorLevel2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" VSPACE=""3"" /><FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>Ya existen registros en las fechas indicadas.</B></FONT>"
											'	Response.Write "parent.window.document." & oRequest("TargetField").Item & ".ReportedHours.value='-2';" & vbNewLine
											'ElseIf (CDbl(lLimit(0)) > 0) And (CDbl(lLimit(0)) < CDbl(oRequest("WorkedHours").Item)) Then
											'	sRecordID = "<BR /><IMG SRC=""Images/IcnErrorLevel2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" VSPACE=""3"" /><FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>El número máximo de horas por día será excedido.</B></FONT>"
											'	Response.Write "parent.window.document." & oRequest("TargetField").Item & ".ReportedHours.value='-1';" & vbNewLine
											'	bEmpty = True
											'ElseIf (CDbl(lLimit(1)) > 0) And (CDbl(lLimit(1)) < dCounter(1)) Then
											'	sRecordID = "<BR /><IMG SRC=""Images/IcnErrorLevel2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" VSPACE=""3"" /><FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>El número máximo de días por quincena será excedido.</B></FONT>"
											'	Response.Write "parent.window.document." & oRequest("TargetField").Item & ".ReportedHours.value='-1';" & vbNewLine
											'	bEmpty = True
											'ElseIf (CDbl(lLimit(2)) > 0) And (CDbl(lLimit(2)) < (dCounter(2) + dTotal)) Then
											'	sRecordID = "<BR /><IMG SRC=""Images/IcnErrorLevel2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" VSPACE=""3"" /><FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>El número máximo de horas por quincena será excedido.</B></FONT>"
											'	Response.Write "parent.window.document." & oRequest("TargetField").Item & ".ReportedHours.value='-1';" & vbNewLine
											'	bEmpty = True
											'ElseIf (dCounter(0) + dTotal) > 265 Then
											'	sRecordID = "<BR /><IMG SRC=""Images/IcnErrorLevel2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" VSPACE=""3"" /><FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>El número máximo de horas por quincena para todos los registros será excedido.</B></FONT>"
											'	Response.Write "parent.window.document." & oRequest("TargetField").Item & ".ReportedHours.value='-1';" & vbNewLine
											'	bEmpty = True
											Else
												sRecordID = "<BR />Se adicionarán " & dTotal & " horas."
												'Response.Write "parent.window.document." & oRequest("TargetField").Item & ".ReportedHours.value='1';" & vbNewLine
												bEmpty = True
											End If
										End If

										dTotal = 0
										If bEmpty Then
											lRecordID = Replace(oRequest("RecordID").Item, "'", "")
											If Len(oRequest("OriginalEmployeeID").Item) > 0 Then
												If CLng(oRequest("OriginalEmployeeID").Item) > -1 Then lRecordID = oRequest("OriginalEmployeeID").Item
											End If
											sErrorDescription = "No se pudieron obtener los montos del Sueldo base."
											'If CLng(lRecordID) < 800000 Then
											If False Then
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.*, Areas.EconomicZoneID From EmployeesChangesLKP, EmployeesHistoryList, Areas Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesChangesLKP.EmployeeID=" & lRecordID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate<=" & lEndDate & ") Order By EmployeesHistoryList.EmployeeDate Desc", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
												Response.Write vbNewLine & "// Query: Select EmployeesHistoryList.*, Areas.EconomicZoneID From EmployeesChangesLKP, EmployeesHistoryList, Areas Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesChangesLKP.EmployeeID=" & lRecordID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate<=" & lEndDate & ") Order By EmployeesHistoryList.EmployeeDate Desc -->" & vbNewLine
											Else
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Positions.*, Areas.EconomicZoneID As EconomicZoneID_A From Positions, Areas, Zones, ZoneTypes Where (Areas.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (PositionID=" & oRequest("PositionID").Item & ") And (AreaID=" & oRequest("AreaID").Item & ")", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
												Response.Write vbNewLine & "// Query: Select Positions.*, (ZoneTypeID2-2) As EconomicZoneID From Positions, Areas, Zones, ZoneTypes Where (Areas.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (PositionID=" & oRequest("PositionID").Item & ") And (AreaID=" & oRequest("AreaID").Item & ") -->" & vbNewLine
											End If
											If lErrorNumber = 0 Then
												If Not oRecordset.EOF Then
													sCondition = "(EmployeeTypeID In (-1,<EMPLOYEE_TYPE_ID />)) And (ClassificationID In (-1,<CLASSIFICATION_ID />)) And (GroupGradeLevelID In (-1,<GROUP_GRADE_LEVEL_ID />)) And (IntegrationID In (-1,<INTEGRATION_ID />)) And (LevelID In (-1,<LEVEL_ID />)) And (EconomicZoneID In (0,<ECONOMIC_ZONE_ID />))"
													sCondition = Replace(sCondition, "<EMPLOYEE_TYPE_ID />", oRecordset.Fields("EmployeeTypeID").Value)
													sCondition = Replace(sCondition, "<CLASSIFICATION_ID />", oRecordset.Fields("ClassificationID").Value)
													sCondition = Replace(sCondition, "<GROUP_GRADE_LEVEL_ID />", oRecordset.Fields("GroupGradeLevelID").Value)
													sCondition = Replace(sCondition, "<INTEGRATION_ID />", oRecordset.Fields("IntegrationID").Value)
													sCondition = Replace(sCondition, "<LEVEL_ID />", oRecordset.Fields("LevelID").Value)
													sCondition = Replace(sCondition, "<ECONOMIC_ZONE_ID />", oRecordset.Fields("EconomicZoneID_A").Value)
													oRecordset.Close

													sErrorDescription = "No se pudieron obtener los montos del Sueldo base."
													lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptID, Max(ConceptAmount) As MaxConcept From ConceptsValues Where (ConceptID In (1,4,38,130)) And " & sCondition & " Group By ConceptID Order By ConceptID", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
													Response.Write vbNewLine & "// Query: Select ConceptID, Max(ConceptAmount) As MaxConcept From ConceptsValues Where (ConceptID In (1,4,38,130)) And " & sCondition & " Group By ConceptID Order By ConceptID -->" & vbNewLine
													If lErrorNumber = 0 Then
														dTotal = 0
														Do While Not oRecordset.EOF
															Select Case CLng(oRecordset.Fields("ConceptID").Value)
																Case 1
																	Select Case oRequest("RiskLevelID").Item
																		Case "1"
																			dTotal = dTotal + CDbl(oRecordset.Fields("MaxConcept").Value * 1.1)
																		Case "2"
																			dTotal = dTotal + CDbl(oRecordset.Fields("MaxConcept").Value * 1.2)
																		Case Else
																			dTotal = dTotal + CDbl(oRecordset.Fields("MaxConcept").Value)
																	End Select
																Case 4
																	dTotal = dTotal * ((100 + CDbl(oRecordset.Fields("MaxConcept").Value)) / 100 )
																Case 38
																	If CLng(Replace(oRequest("RecordID").Item, "'", "")) < 800000 Then
																		dTotal = dTotal + CDbl(oRecordset.Fields("MaxConcept").Value)
																	End If
																Case 130
																	
																	If CLng(Replace(oRequest("RecordID").Item, "'", "")) < 800000 Then
																		dTotal = dTotal + CDbl(oRecordset.Fields("MaxConcept").Value)
																	End If
															End Select
															oRecordset.MoveNext
															If Err.number <> 0 Then Exit Do
														Loop
														oRecordset.Close
													End If
												End If
											End If
											Select Case CInt(oRequest("RecordType").Item)
												Case 1 'Guardias
													If CLng(Replace(oRequest("RecordID").Item, "'", "")) < 800000 Then
														dTotal = (dTotal / 15) * 2 'Internos
													Else
														dTotal = (dTotal / 15) 'Externos
													End If
													dTotal = dTotal / CDbl(oRequest("WorkingHours").Item) * CDbl(oRequest("WorkedHours").Item)
												Case 2 'Suplencias
													If CLng(Replace(oRequest("RecordID").Item, "'", "")) < 800000 Then
														sErrorDescription = "No se pudieron obtener los montos del Sueldo base."
														'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select SpecialJourneyFactor1, SpecialJourneyFactor2 From EmployeesChangesLKP, EmployeesHistoryList, Journeys Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesChangesLKP.EmployeeID=" & lRecordID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lEndDate & ")", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select SpecialJourneyFactor1, SpecialJourneyFactor2 From Journeys Where (Journeys.JourneyID=" & CStr(oRequest("JourneyID").Item) & ")", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
														Response.Write vbNewLine & "// Query: Select SpecialJourneyFactor1, SpecialJourneyFactor2 From EmployeesChangesLKP, EmployeesHistoryList, Journeys Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesChangesLKP.EmployeeID=" & lRecordID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lEndDate & ") -->" & vbNewLine
														If CDbl(oRequest("WorkedHours").Item) >= CDbl(oRecordset.Fields("SpecialJourneyFactor2").Value) Then
															dTotal = (dTotal / 15) * CDbl(oRequest("WorkedHours").Item) * CDbl(oRecordset.Fields("SpecialJourneyFactor1").Value) 'Factor de jornada
														Else
															dTotal = (dTotal / 15) * CDbl(oRequest("WorkedHours").Item)
														End If
														oRecordset.Close
													Else
														dTotal = (dTotal / 15) 'Externos
														Select Case CInt(oRequest("JourneyID").Item)
															Case 1, 2, 11, 12
																If CDbl(oRequest("WorkedHours").Item) >= 5 Then
																	dTotal = dTotal * CDbl(oRequest("WorkedHours").Item) * 1.4
																Else
																	dTotal = dTotal * CDbl(oRequest("WorkedHours").Item)
																End If
															Case 3, 13
																If CDbl(oRequest("WorkedHours").Item) >= 5 Then
																	dTotal = dTotal * CDbl(oRequest("WorkedHours").Item) * 1.4 * 2
																Else
																	dTotal = dTotal * CDbl(oRequest("WorkedHours").Item) * 2
																End If
															Case 4,5, 14, 15
																If CDbl(oRequest("WorkedHours").Item) >= 5 Then
																	dTotal = dTotal * CDbl(oRequest("WorkedHours").Item) * 1.4 * 4
																Else
																	dTotal = dTotal * CDbl(oRequest("WorkedHours").Item) * 4
																End If
															Case 6, 16
																If CDbl(oRequest("WorkedHours").Item) >= 4 Then
																	dTotal = dTotal * CDbl(oRequest("WorkedHours").Item) * 1.4 * 2
																Else
																	dTotal = dTotal * CDbl(oRequest("WorkedHours").Item) * 2
																End If
															Case 7, 17
																If CDbl(oRequest("WorkedHours").Item) >= 3 Then
																	dTotal = dTotal * CDbl(oRequest("WorkedHours").Item) * 1.4 * 2
																Else
																	dTotal = dTotal * CDbl(oRequest("WorkedHours").Item) * 2
																End If
															Case 8, 18
																If CDbl(oRequest("WorkedHours").Item) >= 3 Then
																	dTotal = dTotal * CDbl(oRequest("WorkedHours").Item) * 1.4 * 2
																Else
																	dTotal = dTotal * CDbl(oRequest("WorkedHours").Item) * 2
																End If
														End Select
													End If
													'dCounter = (Abs(DateDiff("d", GetDateFromSerialNumber(lStartDateForRecord), GetDateFromSerialNumber(lEndDateForRecord))) + 1)
													'dTotal = dTotal * (dCounter + (Int(dCounter / CDbl(oRecordset.Fields("SpecialJourneyFactor2").Value)) * 2))
												Case 3 'RQ
												Case 4 'PROVAC
													If CLng(Replace(oRequest("RecordID").Item, "'", "")) < 800000 Then
														dTotal = (dTotal / 15) * 2 'Internos
													Else
														dTotal = (dTotal / 15) 'Externos
													End If
													dTotal = dTotal / CDbl(oRequest("WorkingHours").Item) * CDbl(oRequest("WorkedHours").Item) * (Abs(DateDiff("d", GetDateFromSerialNumber(lStartDateForRecord), GetDateFromSerialNumber(lEndDateForRecord))) + 1)
											End Select

											If CLng(oRequest("RecordID").Item) < 800000 Then
												lEmployeeTypeID = 0 'Personal Interno
											Else
												lEmployeeTypeID = 1 'Personal Externo
											End If
											'  VerifySpecialJourneyBudgetAmount(oADODBConnection, lAppliedDate, iAreaID, lEmployeeTypeID, lNewAmount)
											If VerifySpecialJourneyBudgetAmount(oADODBConnection, CLng(oRequest("PayrollDate").Item), CLng(oRequest("AreaID").Item), lEmployeeTypeID, dTotal) Then
												sRecordID = sRecordID & "<BR />El monto para cubrir las horas registradas " & FormatNumber(dTotal, 2, True, False, True) & " es cubierto por el presupuesto"
												'Response.Write "alert('No existe presupuesto');" & vbNewLine
												Response.Write "parent.window.document." & oRequest("TargetField").Item & ".ReportedHours.value='1';" & vbNewLine
											Else
												sRecordID = sRecordID & "<BR />El monto para cubrir las horas registradas " & FormatNumber(dTotal, 2, True, False, True) & " no es cubierto por el presupuesto"
												Response.Write "alert('No existe presupuesto');" & vbNewLine
												Response.Write "parent.window.document." & oRequest("TargetField").Item & ".ReportedHours.value='-3';" & vbNewLine
											End If
											If CInt(oRequest("RecordType").Item) <> 3 Then Response.Write "parent.window.document." & oRequest("TargetField").Item & ".ConceptAmount.value='" & FormatNumber(dTotal, 2, True, False, True) & "';" & vbNewLine
										End If
										'Response.Write "alert(" & dTotal & ");" & vbNewLine
									Response.Write "}" & vbNewLine
								End If
							Response.Write "//--></SCRIPT>" & vbNewLine
							Response.Write "<B>" & dCounter(0) & " horas reportadas</B> para todos los registros entre el " & DisplayDateFromSerialNumber(lStartDate, -1, -1, -1) & " y el " & DisplayDateFromSerialNumber(lEndDate, -1, -1, -1) & ".<BR />" & sRecordID
						End If
					Case "ZoneForArea"
						If Len(sRecordID) = 0 Then
							If sFieldID = -1 Then
								Response.Write "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>El centro de trabajo especificado no está registrado en el sistema.</B></FONT>"
							Else
								Response.Write "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>El centro de trabajo no tiene registrada su entidad.</B></FONT>"
							End If
						Else
							Response.Write sRecordID
						End If
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							If Len(oRequest("TargetField").Item) > 0 Then
								Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ")" & vbNewLine
									'Response.Write "parent.window.document." & oRequest("TargetField").Item & ".value = '" & DisplayDateFromSerialNumber(CLng(sField), -1, -1, -1) & "';" & vbNewLine
									Response.Write "parent.window.document." & oRequest("TargetField").Item & ".value = '" & sFieldID & "';" & vbNewLine
							End If
						Response.Write "//--></SCRIPT>" & vbNewLine
					Case "StartDateForConcept"
							If sFieldID = -1 Then
								Response.Write "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>El concepto indicado no está registrado.</B></FONT>"
							Else
								Response.Write DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
							End If
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							If Len(oRequest("TargetField").Item) > 0 Then
								Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ")" & vbNewLine
									Response.Write "parent.window.document." & oRequest("TargetField").Item & ".value = '" &  CLng(oRecordset.Fields("StartDate").Value) & "';" & vbNewLine
							End If
						Response.Write "//--></SCRIPT>" & vbNewLine
					Case "Zones_Level2"
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							sErrorDescription = "No se pudo obtener la información del registro."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ZoneID, ZoneCode, ZoneName From Zones Where (ZoneID>-1) And (ParentID=" & Replace(oRequest("RecordID").Item, "'", "") & ") Order By ZoneCode", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									Response.Write "RemoveAllItemsFromList(null, parent.window.document." & oRequest("TargetField").Item & ".LocationID);" & vbNewLine
									Response.Write "AddItemToList('Todos', '', null, parent.window.document." & oRequest("TargetField").Item & ".LocationID);" & vbNewLine
									Do While Not oRecordset.EOF
										Response.Write "AddItemToList('" & CStr(oRecordset.Fields("ZoneCode").Value) & ". " & CStr(oRecordset.Fields("ZoneName").Value) & "', '" & CStr(oRecordset.Fields("ZoneID").Value) & "', null, parent.window.document." & oRequest("TargetField").Item & ".LocationID);" & vbNewLine
										oRecordset.MoveNext
										If Err.number <> 0 Then Exit Do
									Loop
								End If
								oRecordset.Close
							End If
						Response.Write "//--></SCRIPT>" & vbNewLine
					Case Else
						If Len(sRecordID) = 0 Then
							Response.Write "&nbsp;&nbsp;&nbsp;<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>El registro no se encuentra en el sistema.</B></FONT>"
						Else
							Response.Write "&nbsp;&nbsp;&nbsp;<FONT COLOR=""#" & S_INSTRUCTIONS_FOR_GUI & """><B>Registro encontrado</B></FONT>"
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								If Len(oRequest("TargetField").Item) > 0 Then
									Response.Write "if (parent.window.document." & oRequest("TargetField").Item & ")" & vbNewLine
										Response.Write "parent.window.document." & oRequest("TargetField").Item & ".value = '" & sRecordID & "';" & vbNewLine
								End If
							Response.Write "//--></SCRIPT>" & vbNewLine
						End If
				End Select
			End If%>
		</FONT></FORM>
	</BODY>
</HTML>
<SCRIPT LANGUAGE="JavaScript"><!--
	//HidePopupItem('WaitSmallDiv', document.WaitSmallDiv)
//--></SCRIPT>
<%
Set oRecordset = Nothing
%>