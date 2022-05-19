<%
Function GetJobsURLValues(oRequest, iSelectedTab, bAction, sCondition)
'************************************************************
'Purpose: To initialize the global variables using the URL
'Inputs:  oRequest
'Outputs: iSelectedTab, bAction, sCondition
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetJobsURLValues"
	Dim oItem
	Dim aItem

	iSelectedTab = 1
	If Not IsEmpty(oRequest("Tab").Item) Then
		iSelectedTab = CInt(oRequest("Tab").Item)
	End If
	bAction = (Len(oRequest("Add").Item) > 0) Or (Len(oRequest("Modify").Item) > 0) Or (Len(oRequest("Remove").Item) > 0) Or (Len(oRequest("SetActive").Item) > 0)

	sCondition = ""
	If Len(oRequest("JobNumber").Item) > 0 Then
		'sCondition = sCondition & " And (JobNumber In ('" & Replace(Replace(oRequest("JobNumber").Item, ", ", ","), ",", "','") & "'))"
		sCondition = sCondition & " And (JobNumber Like ('" & S_WILD_CHAR & Replace(CLng(oRequest("JobNumber").Item), "´", "") & S_WILD_CHAR & "'))"
	End If
	If Len(oRequest("AreaID").Item) > 0 Then
		sCondition = sCondition & " And (Jobs.AreaID In (" & Replace(oRequest("AreaID").Item, ", ", ",") & "))"
	End If
	If Len(oRequest("PositionID").Item) > 0 Then
		sCondition = sCondition & " And (Jobs.PositionID In (" & Replace(oRequest("PositionID").Item, ", ", ",") & "))"
	End If
	If (InStr(1, oRequest, "StartStart", vbTextCompare) > 0) Or (InStr(1, oRequest, "EndStart", vbTextCompare) > 0) Then Call GetStartAndEndDatesFromURL("StartStart", "EndStart", "StartDate", False, sCondition)
	If (InStr(1, oRequest, "StartEnd", vbTextCompare) > 0) Or (InStr(1, oRequest, "EndEnd", vbTextCompare) > 0) Then Call GetStartAndEndDatesFromURL("StartEnd", "EndEnd", "EndDate", False, sCondition)

	GetJobsURLValues = Err.number
	Err.Clear
End Function

Function DoJobsAction(oRequest, oADODBConnection, sAction, sErrorDescription)
'************************************************************
'Purpose: To add, change or delete the information of the
'         specified component
'Inputs:  oRequest, oADODBConnection, sAction
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DoJobsAction"
	Dim oRecordset
	Dim sNames
	Dim lErrorNumber
	Dim sQuery
	Dim lJobDate
	Dim lJobDateOld
	Dim lEndDate
	Dim lEndDateOld
	Dim lZoneID
	Dim lAreaID
	Dim lPaymentCenterID
	Dim lJobTypeID
	Dim lShiftID
	Dim lClassifID
	Dim lGglID
	Dim lIntegrationID
	Dim lOccupationTypeID
	Dim lServiceID
	Dim lLevelID
	Dim lEmployeeId
	Dim lOwnerID
	Dim lStatusID

	If Len(oRequest("Add").Item) > 0 Then
		lErrorNumber = AddJob(oRequest, oADODBConnection, aJobComponent, True, sErrorDescription)
	ElseIf Len(oRequest("JobHistoryList").Item) > 0 Then 
		lJobDate = CLng(oRequest("JobYear").Item & oRequest("JobMonth").Item & oRequest("JobDay").Item)
		lEndDate = CLng(oRequest("EndYear").Item & oRequest("EndMonth").Item & oRequest("EndDay").Item)
		If CLng(lEndDate) = 0 Then lEndDate = 30000000
		sQuery = "Select * From JobsHistoryList Where (JobDate = " & lJobDate & ") And (EndDate = " & lEndDate & ") And (StatusID = " & oRequest("StatusID").Item & ") And (JobId =" & oRequest("JobIDH").Item & ")"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				lErrorNumber = -1
				sErrorDescription = "El registro indicado duplica la información del historial."
			Else
				sQuery = "Select ClassificationID, GroupGradeLevelID, IntegrationID, LevelID From Positions Where PositionID = " & CLng(oRequest("PositionID").Item)
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						lClassifID = oRecordset.Fields("ClassificationID").Value
						lGglID = oRecordset.Fields("GroupGradeLevelID").Value
						lIntegrationID = oRecordset.Fields("IntegrationID").Value
						lLevelID = oRecordset.Fields("LevelID").Value
					Else
						lErrorNumber = -1
						sErrorDescription = "No se encontrò la información complementaria del puesto"
					End If
				End If
				sQuery = "Select ZoneID, AreaID, PaymentCenterID, OccupationTypeID, JobTypeID, ShiftID, ServiceID from Jobs Where JobID = " & CLng(oRequest("JobIDH").Item)
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						lZoneID = oRecordset.Fields("ZoneID").Value
						lAreaID = oRecordset.Fields("AreaID").Value
						lPaymentCenterID = oRecordset.Fields("PaymentCenterID").Value
						lOccupationTypeID = oRecordset.Fields("OccupationTypeID").Value
						lJobTypeID = oRecordset.Fields("JobTypeID").Value
						lShiftID = oRecordset.Fields("ShiftID").Value
						lServiceID = oRecordset.Fields("ServiceID").Value
					Else
						lErrorNumber = -1
						sErrorDescription = "No se encontrò la información complementaria de la plaza"
					End If
				End If
				If lErrorNumber = 0 Then
					lEmployeeID = oRequest("EmployeeID").Item
					If Len(lEmployeeID) = 0 Then lEmployeeID = 0
					lOwnerID = oRequest("OwnerID").Item
					If Len(lOwnerID) = 0 Then lOwnerID = 0
					sQuery = "Insert Into JobsHistoryList (JobID, JobDate, EndDate, EmployeeID, OwnerID, CompanyID, ZoneID, AreaID, PaymentCenterID, PositionID, JobTypeID, ShiftID, JourneyID, ClassificationID, GroupGradeLevelID, IntegrationID, OccupationTypeID, ServiceID, LevelID, WorkingHours, StatusID, UserID, ModifyDate) Values (" & _
															oRequest("JobIDH").Item & "," & lJobDate & "," & lEndDate & "," & lEmployeeID & "," & lOwnerID & "," & oRequest("CompanyID").Item & "," & lZoneID & "," & lAreaID & "," & lPaymentCenterID & "," & oRequest("PositionID").Item & "," & _
															lJobTypeID & "," & lShiftID & "," & oRequest("JourneyID").Item & "," & lClassifID & "," & lGglID & "," & lIntegrationID & "," & lOccupationTypeID & "," & lServiceID & "," & lLevelID & "," & oRequest("WorkingHours").Item & "," & _
															oRequest("StatusID").Item & "," & aLoginComponent(N_USER_ID_LOGIN) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					If lErrorNumber = 0 Then
						sQuery = "Select Max(JobDate) MaxEndDate From JobsHistoryList Where JobId = " & oRequest("JobIDH").Item
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If CLng(oRecordset.Fields("MaxEndDate").Value) = lJobDate Then
							sQuery = "Update jobs Set CompanyID=" & oRequest("CompanyID").Item & ",ZoneID=" & lZoneID & ",AreaID=" & lAreaID & ",PaymentCenterID=" & lPaymentCenterID & ",PositionID=" & _
									 oRequest("PositionID").Item & ",JobTypeID=" & lJobTypeID & ",ShiftID=" & lShiftID & ",JourneyID=" & oRequest("JourneyID").Item & ",ClassificationID=" & lClassifID & _
									 ",GroupGradeLevelID=" & lGglID & ",IntegrationID=" & lIntegrationID & ",OccupationTypeID=" & lOccupationTypeID & ",ServiceID=" & lServiceID & ",LevelID=" & lLevelID & _
									 ",WorkingHours=" & oRequest("WorkingHours").Item & ",StatusID=" & oRequest("StatusID").Item & " Where JobID=" & oRequest("JobIDH").Item
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						End If
					End If
				End If
			End If
		End If
	ElseIf Len(oRequest("Modify").Item) > 0 Then
		If CLng(oRequest("JobIDH").Item) > 0 Then
			lJobDate = CLng(oRequest("JobYear").Item & oRequest("JobMonth").Item & oRequest("JobDay").Item)
			lEndDate = CLng(oRequest("EndYear").Item & oRequest("EndMonth").Item & oRequest("EndDay").Item)
			If CLng(lEndDate) = 0 Then lEndDate = 30000000
			lJobDateOld = oRequest("JobDateOld").Item
			lEndDateOld = oRequest("EndDateOld").Item
			lStatusID = oRequest("StatusID").Item
			sQuery = "Select LevelID, ClassificationID, IntegrationID, GroupGradeLevelID From Positions where PositionID = " & oRequest("PositionID").Item
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			lClassifID = oRecordset.Fields("ClassificationID").Value
			lGglID = oRecordset.Fields("GroupGradeLevelID").Value
			lIntegrationID = oRecordset.Fields("IntegrationID").Value
			lLevelID = oRecordset.Fields("LevelID").Value
			oRecordset.Close
			sQuery = "Update JobsHistoryList Set JobDate = " & lJobDate & ", EndDate = " & lEndDate & _
					", EmployeeID = " & oRequest("EmployeeID").Item & ", OwnerId = " & oRequest("OwnerID").Item & _
					", StatusID = " & oRequest("StatusID").Item & ", CompanyID = " & oRequest("CompanyID").Item & _
					", AreaID = " & oRequest("AreaID").Item & ", PaymentCenterID = " & oRequest("PaymentCenterID").Item & _
					", PositionID = " & oRequest("PositionID").Item & ", LevelID = " & lLevelID & _
					", ClassificationID = " & lClassifId & ", GroupGradeLevelID = " & lGglID & _
					", IntegrationID = " & lIntegrationID & ", JobTypeID = " & oRequest("JobTypeID").Item & _
					", OccupationTypeId = " & oRequest("OccupationTypeID").Item & ", ServiceID =" & oRequest("ServiceID").Item & _
					", ShiftID = " & oRequest("ShiftID").Item & ", JourneyID = " & oRequest("JourneyID").Item & _
					", WorkingHours = " & oRequest("WorkingHours").Item & _ 
					" Where (JobID=" & oRequest("JobIDH").Item & ") And (JobDate=" & lJobDateOld & ") And (EndDate = " & lEndDateOld & ")"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			If lErrorNumber = 0 Then
				sQuery = "Select Max(JobDate) MaxEndDate From JobsHistoryList Where JobId = " & oRequest("JobIDH").Item
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If CLng(oRecordset.Fields("MaxEndDate").Value) = lJobDate Then
					sQuery = "Update jobs Set CompanyID=" & oRequest("CompanyID").Item & ",AreaID=" & oRequest("AreaID").Item & ",PaymentCenterID=" & oRequest("PaymentCenterID").Item & ",PositionID=" & _
							 oRequest("PositionID").Item & ",JobTypeID=" & oRequest("JobTypeID").Item & ",ShiftID=" & oRequest("ShiftID").Item & ",JourneyID=" & oRequest("JourneyID").Item & ",ClassificationID=" & _
							 lClassifID & ",GroupGradeLevelID=" & lGglID & ",IntegrationID=" & lIntegrationID & _
							 ",OccupationTypeID=" & oRequest("OccupationTypeID").Item & ",ServiceID=" & oRequest("ServiceID").Item & _
							 ",LevelID=" & lLevelID & ",WorkingHours=" & oRequest("WorkingHours").Item & ",StatusID=" & oRequest("StatusID").Item & " Where JobID=" & oRequest("JobIDH").Item
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					sQuery = "Select StatusID From Jobs Where (JobId=" & oRequest("JobIDH").Item & ")"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					lStatusID = oRecordset.Fields("StatusID").Value
					If lStatusID = "1" Then
						sQuery = "Select EmployeeID From EmployeesHistoryList Where (JobId = " & oRequest("JobIDH").Item & ") And (EmployeeDate = " & lJobDateOld & ") And (EndDate = " & lEndDateOld & ")"
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						lEmployeeId = oRecordset.Fields("EmployeeID").Value
						sQuery = "Update Employees Set GroupGradeLevelId = " & lGglID & ", ClassificationID = " & lClassifId & ", IntegrationID = " & lIntegrationID & ", LevelID =  " & lLevelID & " Where (EmployeeID=" & lEmployeeID & ")"
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						sQuery = "Update EmployeesHistoryList Set PositionID = " & oRequest("PositionID").Item & ", GroupGradeLevelId = " & lGglID & ", ClassificationID = " & lClassifId & ", IntegrationID = " & lIntegrationID & ", LevelID =  " & lLevelID & " Where (JobId = " & oRequest("JobIDH").Item & ") And (EmployeeDate = " & lJobDateOld & ") And (EndDate = " & lEndDateOld & ")"
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
				End If
			End If
		Else
			lErrorNumber = ModifyJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
			If (lErrorNumber = 0) Then
				If (StrComp(oRequest("Modify").Item , "Aplicar Titularidad", vbBinaryCompare) = 0) Then
					aEmployeeComponent(N_ID_EMPLOYEE) = CLng(oRequest("OwnerID").Item)
					'lErrorNumber = GetNameFromTable(oADODBConnection, "EmployeeIDsFromJobs", aJobComponent(N_ID_JOB), "", "", aEmployeeComponent(N_ID_EMPLOYEE), sErrorDescription)
				End If
				If Len(aEmployeeComponent(N_ID_EMPLOYEE)) > 0 Then
					If Len(aEmployeeComponent(N_ID_EMPLOYEE)) > 6 then
						lErrorNumber = -1
						sErrorDescription = "El historial del empleado no pudo modificarse porque la plaza está ocupada por dos empleados activos"
					End If
					If (lErrorNumber = 0) And (aEmployeeComponent(N_ID_EMPLOYEE) > -1) Then
						lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						If lErrorNumber = 0 Then
							aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) = aJobComponent(N_SERVICE_ID_JOB)
							aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) = aJobComponent(N_LEVEL_ID_JOB)
							lErrorNumber = ModifyEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						End If
					End If
				End If
			End If
		End If
	ElseIf Len(oRequest("Remove").Item) > 0 Then
		If aJobComponent(N_ID_JOB) > -1 Then
			lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
			If lErrorNumber = 0 Then 
				If CLng(oRequest("JobIDH").Item) > 0 Then
					lJobDate = CLng(oRequest("JobYear").Item & oRequest("JobMonth").Item & oRequest("JobDay").Item)
					lEndDate = CLng(oRequest("EndYear").Item & oRequest("EndMonth").Item & oRequest("EndDay").Item)
					If lEndDate = 0 Then lEndDate = 30000000
					sQuery = "Delete From jobsHistoryList Where (JobID = " & aJobComponent(N_ID_JOB) & ") And (JobDate = " & lJobDate & ") And (EndDate = " & lEndDate & ")"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				Else
					lErrorNumber = RemoveJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
				End If
			End If
		End If
	ElseIf Len(oRequest("SetActive").Item) > 0 Then
		lErrorNumber = SetActiveForJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
	End If

	Set oRecordset = Nothing
	DoJobsAction = lErrorNumber
	Err.Clear
End Function

Function DisplayJobsSearchForm(oRequest, oADODBConnection, bFull, sErrorDescription)
'************************************************************
'Purpose: To display the search HTML form
'Inputs:  oRequest, oADODBConnection, bFull
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayJobsSearchForm"

	Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""Jobs.asp"" METHOD=""GET"">"
		If bFull Then Response.Write "<B>BÚSQUEDA DE PLAZAS</B><BR /><BR />"
		Response.Write "<TABLE BORDER=""0"" CELLPADING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número de la plaza:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""JobNumber"" ID=""JobNumberTxt"" SIZE=""6"" MAXLENGTH=""6"" CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			If bFull Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Áreas:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""AreaID"" ID=""AreaIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas, Areas As ParentAreas", "Areas.AreaID", "ParentAreas.AreaName, Areas.AreaShortName, Areas.AreaName", "(Areas.ParentID=ParentAreas.AreaID) And (Areas.ParentID>-1) And (Areas.Active=1) And (ParentAreas.Active=1)", "ParentAreas.AreaName, Areas.AreaShortName, Areas.AreaName", oRequest("AreaID").Item, "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Puesto:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""PositionID"" ID=""PositionIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todos</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Positions", "PositionID", "PositionShortName, PositionName", "(PositionID>-1) And (Active=1)", "PositionShortName, PositionName", oRequest("PositionID").Item, "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
			End If
			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""2"""
				If Not bFull Then Response.Write " ALIGN=""RIGHT"""
				Response.Write "><BR /><INPUT TYPE=""SUBMIT"" NAME=""DoSearch"" ID=""DoSearchBtn"" VALUE=""Buscar Plazas"" CLASS=""Buttons"" /></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	Response.Write "</FORM>"

	DisplayJobsSearchForm = Err.number
End Function

Function DisplayJobForms(oRequest, iSelectedTab, sErrorDescription)
'************************************************************
'Purpose: To display the forms for the given job
'Inputs:  oRequest, iSelectedTab
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayJobForms"
	Dim lErrorNumber
	Dim bForm

	Response.Write "<FONT FACE=""Arial"" SIZE=""2"">"
		Response.Write "<B>Número de plaza: </B>" & CleanStringForHTML(aJobComponent(N_ID_JOB)) & "<BR />"
		Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""960"" HEIGHT=""1"" /><BR /><BR />"
	Response.Write "</FONT>"
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
			Select Case iSelectedTab
				Case 2
					bForm = (StrComp(oRequest("Tab").Item, "2", vbBinaryCompare) = 0) And (Len(oRequest("JobDate").Item) > 0)
					Response.Write "<TD WIDTH=""40%"" VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><DIV NAME=""ReportHJDiv"" ID=""ReportHJDiv"""
					If bForm Then Response.Write " STYLE=""height: 350px; width:450px; overflow: auto;"""
					Response.Write ">"
					lErrorNumber = DisplayJobHistoryList(oRequest, oADODBConnection, False, True, aJobComponent, sErrorDescription)
					Response.Write "</DIV></TD>"
					If bForm Then
						lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
						Response.Write "<TD>&nbsp;</TD>"
						Response.Write "<TD BGCOLOR=""" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""1"" ><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
						Response.Write "<TD>&nbsp;</TD>"
						Response.Write "<TD WIDTH=""*"" VALIGN=""TOP"">"
							lErrorNumber = ShowJobHistoryListForm(oRequest, aJobComponent, sErrorDescription)
					End If
					Response.Write "</TD>"
				Case 3
					Response.Write "<TD WIDTH=""60%"" VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"
					lErrorNumber = DisplayJobsHistoryListTable(oRequest, oADODBConnection, False, aJobComponent, sErrorDescription)
					If aEmployeeComponent(N_ID_EMPLOYEE) > -1 Then
					'	Call DisplayEmployeeForm(oRequest, oADODBConnection, GetASPFileName(""), "", ",1,", -1, aEmployeeComponent, sErrorDescription)
					Else
						Call DisplayErrorMessage("Mensaje del sistema", "Esta plaza no está asignada a ningún empleado.")
						Response.Write "<BR /><BR />"
					End If
				Case Else
					If Len(oRequest.Item("ShowInfo")) > 0 Then
						lErrorNumber = DisplayJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
					Else
						lErrorNumber = DisplayJobForm(oRequest, oADODBConnection, GetASPFileName(""), aJobComponent, sErrorDescription)
					End If
			End Select
		If iSelectedTab <> 2 Then
			Response.Write "</FONT></TD>"
		End If
		Response.Write "</FONT></TD>"
	Response.Write "</TR></TABLE>"

	DisplayJobForms = lErrorNumber
	Err.Clear
End Function

Function DisplayJobHistoryList(oRequest, oADODBConnection, bForExport, bFull, aJobComponent, sErrorDescription)
'************************************************************
'Purpose: To display the history list for the job from the
'		  database in a table
'Inputs:  oRequest, oADODBConnection, bForExport, bFull, aJobComponent
'Outputs: aJobComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayJobHistoryList"
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
	Dim sQuery

	sErrorDescription = "No se pudo obtener la información de la plaza."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct JobsHistoryList.JobID, OwnerID, JobsHistoryList.EmployeeID, CompanyShortName, CompanyName, Areas.AreaCode, Areas.AreaName, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, PositionShortName, PositionName, JobTypeShortName, JobTypeName, ShiftShortName, ShiftName, JourneyShortName, JourneyName, JobsHistoryList.ClassificationID, GroupGradeLevelShortName, GroupGradeLevelName, JobsHistoryList.IntegrationID, OccupationTypeShortName, OccupationTypeName, ServiceShortName, ServiceName, LevelShortName, JobsHistoryList.WorkingHours, StatusName, JobsHistoryList.JobDate, JobsHistoryList.EndDate From JobsHistoryList, Companies, Zones, Areas, Areas As PaymentCenters, Positions, JobTypes, Shifts, Journeys, GroupGradeLevels, OccupationTypes, Services, Levels, StatusJobs Where (JobsHistoryList.CompanyID=Companies.CompanyID) And (JobsHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (JobsHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (JobsHistoryList.PositionID=Positions.PositionID) And (JobsHistoryList.JobTypeID=JobTypes.JobTypeID) And (JobsHistoryList.ShiftID=Shifts.ShiftID) And (JobsHistoryList.JourneyID=Journeys.JourneyID) And (JobsHistoryList.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (JobsHistoryList.OccupationTypeID=OccupationTypes.OccupationTypeID) And (JobsHistoryList.ServiceID=Services.ServiceID) And (JobsHistoryList.LevelID=Levels.LevelID) And (JobsHistoryList.StatusID=StatusJobs.StatusID) And (JobsHistoryList.JobID=" & aJobComponent(N_ID_JOB) & ") Order By JobsHistoryList.JobID, JobDate Desc", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><TABLE BORDER="""
				If bForExport Then
					Response.Write "1"
				Else
					Response.Write "0"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bFull Then
					If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_AdministracionDePlazas & ",", vbBinaryCompare) > 0) And (Not bForExport) And (Len(oRequest("ReasonID").Item) = 0) Then
						If InStr(request.Ite,"ReportID",vbBinaryCompare) = 0 Then
							asColumnsTitles = Split("Acciones,Fecha inicio,Fecha fin,Estatus,No. titularidad,Compañía,Adscripción,Centro de pago,Puesto,Nivel-subnivel,Clasificación,Grupo grado nivel,Integración,Tipo de plaza,Tipo de ocupación,Servicio,Turno,Horario,Jornada,", ",", -1, vbBinaryCompare)
							asCellWidths = Split("100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
							asCellAlignments = Split("CENTER,,,,,,,,,,,,CENTER,,CENTER,,,CENTER,CENTER", ",", -1, vbBinaryCompare)
						Else
							asColumnsTitles = Split("Fecha inicio,Fecha fin,Estatus,No. titularidad,Compañía,Adscripción,Centro de pago,Puesto,Nivel-subnivel,Clasificación,Grupo grado nivel,Integración,Tipo de plaza,Tipo de ocupación,Servicio,Turno,Horario,Jornada,", ",", -1, vbBinaryCompare)
							asCellWidths = Split("100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
							asCellAlignments = Split(",,,,,,,,,,,CENTER,,CENTER,,,CENTER,CENTER", ",", -1, vbBinaryCompare)
						End If
					Else
						asColumnsTitles = Split("Fecha inicio,Fecha fin,Estatus,No. titularidad,Compañía,Adscripción,Centro de pago,Puesto,Nivel-subnivel,Clasificación,Grupo grado nivel,Integración,Tipo de plaza,Tipo de ocupación,Servicio,Turno,Horario,Jornada,", ",", -1, vbBinaryCompare)
						asCellWidths = Split("100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
						asCellAlignments = Split(",,,,,,,,,,,CENTER,,CENTER,,,CENTER,CENTER", ",", -1, vbBinaryCompare)
					End If
				Else
					asColumnsTitles = Split("Fecha inicio,Fecha fin,Estatus,No. titularidad, Empleado", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,100,100,100,100", ",", -1, vbBinaryCompare)
					asCellAlignments = Split(",,,,", ",", -1, vbBinaryCompare)
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
					sFontBegin = ""
					sFontEnd = ""
					sBoldBegin = ""
					sBoldEnd = ""
					sRowContents = ""
					If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					If Len(request("ReportID").Item) = 0 Then
						If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_AdministracionDePlazas & ",", vbBinaryCompare) > 0) And (Not bForExport) And (Len(oRequest("ReasonID").Item) = 0) Then
							sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Jobs&Change=1&Tab=2&JobID=" & oRequest("JobID").Item & "" & "&JobDate=" & oRecordset.Fields("JobDate").Value & "" & "&EndDate=" & oRecordset.Fields("EndDate").Value & "" & "&EmployeeID=" & oRecordset.Fields("EmployeeID").Value & "" & "&OwnerID=" & oRecordset.Fields("OwnerID").Value & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
							sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Jobs&Delete=1&Tab=2&JobID=" & oRequest("JobID").Item & "" & "&JobDate=" & oRecordset.Fields("JobDate").Value & "" & "&EndDate=" & oRecordset.Fields("EndDate").Value & "" & "&EmployeeID=" & oRecordset.Fields("EmployeeID").Value & "" & "&OwnerID=" & oRecordset.Fields("OwnerID").Value & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
							sRowContents = sRowContents & TABLE_SEPARATOR
						End If
					End If
					If CLng(oRecordset.Fields("JobDate").Value) = 0 Then
						sRowContents = sRowContents & sFontBegin & sBoldBegin & "-" & sBoldEnd & sFontEnd
					Else
						sRowContents = sRowContents & sFontBegin & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("JobDate").Value), -1, -1, -1) & sBoldEnd & sFontEnd
					End If
					If (CLng(oRecordset.Fields("EndDate").Value) = 30000000) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & "Indefinida" & sBoldEnd & sFontEnd
					ElseIf (CLng(oRecordset.Fields("EndDate").Value) = 0) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & "---" & sBoldEnd & sFontEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1) & sBoldEnd & sFontEnd
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value))
					If bForExport Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & "=T(""" & Right("000000" & CStr(oRecordset.Fields("OwnerID").Value), Len("000000")) & """)" & sBoldEnd & sFontEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & Right("000000" & CStr(oRecordset.Fields("OwnerID").Value), Len("000000")) & sBoldEnd & sFontEnd
					End If
					If bFull Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("CompanyShortName").Value) & ". " & CStr(oRecordset.Fields("CompanyName").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value) & ". " & CStr(oRecordset.Fields("AreaName").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value) & ". " & CStr(oRecordset.Fields("PaymentCenterName").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value)) & sBoldEnd & sFontEnd
						If (CStr(oRecordset.Fields("ClassificationID").Value) <> "-1") Then
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ClassificationID").Value)) & sBoldEnd & sFontEnd
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & "Ninguna" & sBoldEnd & sFontEnd
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("GroupGradeLevelShortName").Value) & ". " & CStr(oRecordset.Fields("GroupGradeLevelName").Value)) & sBoldEnd & sFontEnd
						If (CStr(oRecordset.Fields("IntegrationID").Value) <> "-1") Then
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("IntegrationID").Value)) & sBoldEnd & sFontEnd
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & "Ninguna" & sBoldEnd & sFontEnd
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("JobTypeShortName").Value) & ". " & CStr(oRecordset.Fields("JobTypeName").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("OccupationTypeShortName").Value) & ". " & CStr(oRecordset.Fields("OccupationTypeName").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ServiceShortName").Value) & ". " & CStr(oRecordset.Fields("ServiceName").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("JourneyShortName").Value) & ". " & CStr(oRecordset.Fields("JourneyName").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ShiftShortName").Value) & ". " & CStr(oRecordset.Fields("ShiftName").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("WorkingHours").Value)) & sBoldEnd & sFontEnd
					End If

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
			Response.Write "</TABLE></DIV>" & vbNewLine
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayJobHistoryList = lErrorNumber
	Err.Clear
End Function

Function DisplayJobsTabs(oRequest, bError, sErrorDescription)
'************************************************************
'Purpose: To display the tabs for the jobs HTML forms
'Inputs:  oRequest, bError
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayJobsTabs"
	Dim asTitles
	Dim iIndex
	Dim sAction
	Dim lErrorNumber

	asTitles = Split(",Información de la plaza,Historial,Historial de ocupantes de la plaza", ",")
	If (Len(oRequest("New").Item) > 0) Or (bError And (Len(oRequest("Add").Item) > 0)) Then
		Response.Write "<TABLE BORDER=""0"" WIDTH=""98%"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
			Response.Write "<TD BGCOLOR=""#" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""5"" NAME=""TabContents1LfDiv"" ID=""TabContents1LfDiv""><IMG SRC=""Images/TbLf.gif"" WIDTH=""5"" HEIGHT=""21"" /></TD>"
			Response.Write "<TD BGCOLOR=""#" & S_MAIN_COLOR_FOR_GUI & """ BACKGROUND=""Images/TbBg.gif"" WIDTH=""130"" ALIGN=""CENTER"" NAME=""TabContents1Div"" ID=""TabContents1Div""><NOBR><FONT FACE=""Arial"" COLOR=""#" & S_MENU_LINK_FOR_GUI & """ SIZE=""2"" CLASS=""TabLink"">"
			Response.Write "<B>&nbsp;&nbsp;&nbsp;" & asTitles(1) & "&nbsp;&nbsp;&nbsp;</B></FONT></NOBR></TD>"
			Response.Write "<TD BGCOLOR=""#" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""5"" NAME=""TabContents1RgDiv"" ID=""TabContents1RgDiv""><IMG SRC=""Images/TbRg.gif"" WIDTH=""5"" HEIGHT=""21"" /></TD>"
			Response.Write "<TD BACKGROUND=""Images/TbBgDot.gif"" WIDTH=""*""><IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""21"" /></TD>"
		Response.Write "</TR></TABLE><BR />"
	Else
		sAction = "ShowInfo"
		If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then sAction = "Change"
		Response.Write "<TABLE BORDER=""0"" WIDTH=""98%"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
			For iIndex = 1 To UBound(asTitles)
				Response.Write "<TD BGCOLOR=""#"
					If iSelectedTab = iIndex Then
						Response.Write S_MAIN_COLOR_FOR_GUI
					Else
						Response.Write "CCCCCC"
					End If
				Response.Write """ WIDTH=""5"" NAME=""TabContents" & iIndex & "LfDiv"" ID=""TabContents" & iIndex & "LfDiv""><IMG SRC=""Images/TbLf.gif"" WIDTH=""5"" HEIGHT=""21"" /></TD>"
				Response.Write "<TD BGCOLOR=""#"
					If iSelectedTab = iIndex Then
						Response.Write S_MAIN_COLOR_FOR_GUI
					Else
						Response.Write "CCCCCC"
					End If
				Response.Write """ BACKGROUND=""Images/TbBg.gif"" WIDTH=""130"" ALIGN=""CENTER"" NAME=""TabContents" & iIndex & "Div"" ID=""TabContents" & iIndex & "Div""><NOBR><FONT FACE=""Arial"" SIZE=""2"">"
				Response.Write "<A HREF=""" & GetASPFileName("") & "?Action=Jobs&JobID=" & aJobComponent(N_ID_JOB) & "&" & sAction & "=1&Tab=" & iIndex & """ CLASS=""TabLink""><DIV NAME=""TabText" & iIndex & "Div"" ID=""TabText" & iIndex & "Div"" STYLE=""color: #"
					If iSelectedTab = iIndex Then
						Response.Write S_MENU_LINK_FOR_GUI
					Else
						Response.Write "000000"
					End If
				Response.Write ";""><B>&nbsp;&nbsp;&nbsp;" & asTitles(iIndex) & "&nbsp;&nbsp;&nbsp;</B></DIV></A></FONT></NOBR></TD>"
				Response.Write "<TD BGCOLOR=""#"
					If iSelectedTab = iIndex Then
						Response.Write S_MAIN_COLOR_FOR_GUI
					Else
						Response.Write "CCCCCC"
					End If
				Response.Write """ WIDTH=""5"" NAME=""TabContents" & iIndex & "RgDiv"" ID=""TabContents" & iIndex & "RgDiv""><IMG SRC=""Images/TbRg.gif"" WIDTH=""5"" HEIGHT=""21"" /></TD>"
			Next
			Response.Write "<TD BACKGROUND=""Images/TbBgDot.gif"" WIDTH=""*""><IMG SRC=""Images/Transparent.gif"" WIDTH=""21"" HEIGHT=""21"" /></TD>"
		Response.Write "</TR></TABLE>"
	End If

	DisplayJobsTabs = lErrorNumber
	Err.Clear
End Function

Function ShowJobHistoryListForm(oRequest, aJobComponent, sErrorDescription)
'************************************************************
'Purpose: To display the necesary fields form adding a new
'		  record into JobsHistoryList Table
'Inputs:  oRequest, aJobComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ShowJobHistoryListForm"
	Dim lErrorNumber
	Dim sNames
	Dim lStartDate
	Dim lEndDate
	Dim lEmployeeID
	Dim lOwnerID
	Dim sQuery
	Dim oRecordset
	Dim oRecordsetPosition

		If (Len(oRequest("Change").Item) > 0 Or Len (oRequest("Delete").Item)>0) Then
			lStartDate = oRequest("JobDate").Item
			lEndDate = oRequest("EndDate").Item
			lEmployeeID = oRequest("EmployeeID").Item
			lOwnerID = oRequest("OwnerID").Item
			sQuery = "Select * From JobsHistoryList Where (JobID = " & aJobComponent(N_ID_JOB) & ") And (JobDate=" & lStartDate & ") And (EndDate=" & lEndDate & ")"
		Else
			lStartDate = 0
			lEndDate = 0
			lEmployeeID = ""
			lOwnerID = ""
			sQuery = "Select * From Jobs Where JobID = " & aJobComponent(N_ID_JOB)
		End If
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write "<FORM NAME=""Action"" Value=""Jobs"" ID=""HistoryListFrm"" ACTION=""Jobs.asp"" METHOD=""GET"" >"
		If aEmployeeComponent(N_ID_JOB) <> -1 Then
			Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Información del historial</B></FONT>"
			Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número de la plaza:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""HIDDEN"" NAME=""JobIDH"" ID=""JobIDHTxt"" VALUE=""" & aJobComponent(N_ID_JOB) & """ CLASS=""TextFields"" /><FONT FACE=""Arial"" SIZE=""2"">" & aJobComponent(N_ID_JOB) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Empresa:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""CompanyID"" ID=""CompanyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "companies", "companyID", "CompanyShortName, CompanyName", "(CompanyID>-1) And (Active=1)", "companyShortName, companyName", oRecordset.Fields("CompanyID").Value, "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Centro de Trabajo:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""AreaID"" ID=""AreaIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(AreaID>-1) And (Active=1)", "AreaShortName, AreaName", oRecordset.Fields("AreaID").Value, "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Centro de Pago:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""PaymentCenterID"" ID=""PaymentCenterIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(AreaID>-1) And (Active=1)", "AreaShortName, AreaName", oRecordset.Fields("PaymentCenterID").Value, "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Puesto:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						Response.Write "<SELECT NAME=""PositionID"" ID=""PositionIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							sErrorDescription = "No se pudo obtener la información del registro."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Positions.PositionID, Positions.PositionShortName, Positions.PositionName, Positions.EmployeeTypeID, EmployeeTypeName, CompanyName, GroupGradeLevelShortName, GroupGradeLevelName, LevelName, ClassificationID, IntegrationID, WorkingHours From Positions, EmployeeTypes, Companies, GroupGradeLevels, Levels Where (Positions.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Positions.CompanyID=Companies.CompanyID) And (Positions.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (Positions.LevelID=Levels.LevelID) And (Positions.EndDate=30000000) And (EmployeeTypes.EndDate=30000000) And (Companies.EndDate=30000000) And (GroupGradeLevels.EndDate=30000000) And (Levels.EndDate=30000000) And (PositionID>-1) Order By Positions.PositionShortName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordsetPosition)
							If lErrorNumber = 0 Then
								Do While Not oRecordsetPosition.EOF
									Response.Write "<OPTION VALUE=""" & CStr(oRecordsetPosition.Fields("PositionID").Value) & """"
										If aJobComponent(N_POSITION_ID_JOB) = CLng(oRecordsetPosition.Fields("PositionID").Value) Then Response.Write " SELECTED=""1"""
									Response.Write ">" & CStr(oRecordsetPosition.Fields("PositionShortName").Value) & ". " & CStr(oRecordsetPosition.Fields("PositionName").Value) & " (Tabulador: " & CStr(oRecordsetPosition.Fields("EmployeeTypeName").Value) & ", Compañía: " & CStr(oRecordsetPosition.Fields("CompanyName").Value) & ", "
										If CLng(oRecordsetPosition.Fields("EmployeeTypeID").Value) = 1 Then
											Response.Write "GGN: " & CStr(oRecordsetPosition.Fields("GroupGradeLevelShortName").Value) & ", Clasificación:" & CStr(oRecordsetPosition.Fields("ClassificationID").Value) & ", Integración: " & CStr(oRecordsetPosition.Fields("IntegrationID").Value)
										Else
											Response.Write "Nivel: " & CStr(oRecordsetPosition.Fields("LevelName").Value)
										End If
									Response.Write ", Horas laboradas: " & CStr(oRecordsetPosition.Fields("WorkingHours").Value) & ")" & "</OPTION>"
									oRecordsetPosition.MoveNext
									If Err.number <> 0 Then Exit Do
								Loop
							End If
						Response.Write "</SELECT>"
				Response.Write "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de plaza:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""JobTypeID"" ID=""JobTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "JobTypes", "jobTypeID", "JobTypeName", "(Active=1)", "JobTypeName", oRecordset.Fields("JobTypeID").Value, "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de ocupación:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""OccupationTypeID"" ID=""OccupationTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "OccupationTypes", "OccupationTypeID", "OccupationTypeName", "(Active=1)", "OccupationTypeName", oRecordset.Fields("OccupationTypeID").Value, "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Servicio:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""ServiceID"" ID=""ServiceIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Services", "ServiceID", "ServiceShortName, ServiceName", "(Active=1)", "ServiceShortName, ServiceName", oRecordset.Fields("ServiceID").Value, "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Turno:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""JourneyID"" ID=""JourneyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Journeys", "JourneyID", "JourneyShortName, JourneyName", "(Active=1)", "JourneyShortName, JourneyName", oRecordset.Fields("JourneyID").Value, "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Horario:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""ShiftID"" ID=""ShiftIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Shifts", "ShiftID", "ShiftShortName, ShiftName", "(Active=1)", "ShiftShortName, ShiftName", oRecordset.Fields("ShiftID").Value, "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Jornada:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""WorkingHours"" ID=""WorkingHoursTxt"" VALUE=""" & oRecordset.Fields("WorkingHours").Value & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE>"
			Response.Write "<IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"

			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(lStartDate, "Job", Year(Date())-10, Year(Date())+1, True, False) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de término:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(lEndDate, "End", Year(Date())-10, Year(Date())+1, True, True) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del Empleado:&nbsp;</FONT></TD>"
					Response.Write "<TD>"
						Response.Write "<INPUT TYPE=""TEXT"" NAME=""EmployeeID"" ID=""EmployeeIDTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE=""" & lEmployeeID & """ CLASS=""TextFields"" />"
					Response.Write "</TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Titular:&nbsp;</FONT></TD>"
					Response.Write "<TD>"
						Response.Write "<INPUT TYPE=""TEXT"" NAME=""OwnerID"" ID=""OwnerIDTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE=""" & lOwnerID & """ CLASS=""TextFields"" />"
					Response.Write "</TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Estatus de la plaza:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><SELECT NAME=""StatusID"" ID=""StatusIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "StatusJobs", "StatusID", "StatusName", "(Active=1)", "StatusName", oRecordset.Fields("ServiceID").Value, "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE>"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobDateOld"" ID=""JobDateOldHdn"" VALUE=""" & oRequest("JobDate").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EndDateOld"" ID=""EndDateOldHdn"" VALUE=""" & oRequest("EndDate").Item & """ />"
			End If
			Response.Write "<BR /><BR />"
			If Len(oRequest("Delete").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveConceptWngDiv']); ConceptFrm.Remove.focus()"" />"
			ElseIf Len(oRequest("JobDate").Item) = 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""JobHistoryList"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			ElseIf StrComp(oRequest("JobDate").Item, "0", vbBinaryCompare) = 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""JobHistoryList"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			Else
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='Jobs.Asp?Action=Jobs&JobID=" & aJobComponent(N_ID_JOB) & "&Change=1&Tab=2'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveConceptWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
		Response.Write "</FORM>"
End Function
%>