<%
Function GetPositionsURLValues(oRequest, bShowForm, bAction, sCondition)
'************************************************************
'Purpose: To initialize the global variables using the URL
'Inputs:  oRequest
'Outputs: bShowForm, bAction, sCondition
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPositionsURLValues"
	Dim oItem
	Dim aItem

	bAction = (Len(oRequest("Add").Item) > 0) Or (Len(oRequest("Modify").Item) > 0) Or (Len(oRequest("Remove").Item) > 0) Or (Len(oRequest("SetActive").Item) > 0)
	bShowForm = (bAction And (Len(oRequest("Remove").Item) = 0) And (Len(oRequest("SetActive").Item) = 0)) Or (Len(oRequest("New").Item) > 0) Or (Len(oRequest("Change").Item) > 0) Or (Len(oRequest("Delete").Item) > 0)

	sCondition = ""
	If (Len(oRequest("PositionShortName").Item) > 0) And (StrComp(CStr(oRequest("PositionShortName").Item), "-2",vbBinaryCompare) <> 0) Then
		sCondition = sCondition & " And (PositionShortName In ('" & Replace(oRequest("PositionShortName").Item, ", ", ",") & "'))"
	End If
	If Len(oRequest("PositionName").Item) > 0 Then
		sCondition = sCondition & " And (PositionName In ('" & Replace(oRequest("PositionName").Item, ", ", ",") & "'))"
	End If

	GetPositionsURLValues = Err.number
	Err.Clear
End Function

Function DisplayPisitionsFilters(oRequest, sAction, sErrorDescription)
'************************************************************
'Purpose: To display the filter of the specified catalog.
'Inputs:  sAction
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPisitionsFilters"
	Dim bHasFilter
	Dim iIndex
	Dim lErrorNumber

	bHasFilter = False
	Response.Write "<FORM NAME=""FilterFrm"" ID=""FilterFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""GET"" /><FONT FACE=""Arial"" SIZE=""2"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & sAction & """ />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartPage"" ID=""StartPageHdn"" VALUE=""1"" />"
		Select Case sAction
			Case "Areas"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Seleccione los datos para filtrar los registros:&nbsp;</B>"
				Response.Write "<INPUT TYPE=""TEXT"" NAME=""FilterName"" ID=""FilterNameTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & oRequest("FilterName").Item & """ CLASS=""TextFields"" />"
				bHasFilter = True
			Case "FormFields"
			Case "Forms"
			Case "Profiles"
			Case "Users"
			Case "Zones"
			Case Else
				Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" />Mostrar los registros&nbsp;"
				Response.Write "<BR />"
				Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" />Del&nbsp;" & DisplayDateCombos(oRequest("StartForValueYear").Item, oRequest("StartForValueMonth").Item, oRequest("StartForValueDay").Item, "StartForValueYear", "StartForValueMonth", "StartForValueDay", N_FORM_START_YEAR, Year(Date()), True, True)
				'Response.Write "<BR />"
				Response.Write "&nbsp;al&nbsp;"
				Response.Write DisplayDateCombos(oRequest("EndForValueYear").Item, oRequest("EndForValueMonth").Item, oRequest("EndForValueDay").Item, "EndForValueYear", "EndForValueMonth", "EndForValueDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "<BR />"

				Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" />Mostrar los registros con clave:&nbsp;"
					Response.Write "<SELECT NAME=""PositionShortNameFilter"" ID=""PositionShortNameFilterCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE=""-1"">Todos</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Positions", "Distinct PositionShortName", "PositionShortName As PositionShortName2", "PositionShortName IS NOT NULL AND PositionShortName <> ' '", "PositionShortName", aPositionComponent(S_SHORT_NAME_POSITION), "", sErrorDescription)
					Response.Write "</SELECT>"
				Response.Write "<BR />"

				Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" />Mostrar los registros con Grupo, grado, nivel:&nbsp;"
					Response.Write "<SELECT NAME=""GroupGradeLevelIDFilter"" ID=""GroupGradeLevelIDFilterCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE=""-2"">Todos</OPTION>"
						Response.Write "<OPTION VALUE=""-1"">Ninguno</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "GroupGradeLevels", "GroupGradeLevelID", "GroupGradeLevelShortName", "(GroupGradeLevelID>-1) And (Active=1)", "GroupGradeLevelName", aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11), "", sErrorDescription)
					Response.Write "</SELECT>"
				Response.Write "<BR />"

                Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" />Mostrar los registros con Tabulador:&nbsp;"
					Response.Write "<SELECT NAME=""EmployeeTypeIDFilter"" ID=""EmployeeTypeIDFilterCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE=""-2"">Todos</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EmployeeTypes", "EmployeeTypeID", "EmployeeTypeName", "(EmployeeTypeID>=0) And (EmployeeTypeID<7) And (Active=1)", "EmployeeTypeName", aPositionComponent(N_EMPLOYEE_TYPE_ID_POSITION), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT>"
				Response.Write "<BR />"

				bHasFilter = True
		End Select
		If bHasFilter Then Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" /><INPUT TYPE=""SUBMIT"" NAME=""ApplyFilter"" ID=""ApplyFilterBtn"" VALUE=""  Filtrar  "" CLASS=""Buttons"" />"
		Response.Write "<BR /><BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""960"" HEIGHT=""1"" /><BR />"
	Response.Write "</FONT></FORM>"

	DisplayPisitionsFilters = lErrorNumber
	Err.Clear
End Function

Function DoPositionsAction(oRequest, oADODBConnection, sAction, sErrorDescription)
'************************************************************
'Purpose: To add, change or delete the information of the
'         specified component
'Inputs:  oRequest, oADODBConnection, sAction
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DoPositionsAction"
	Dim oRecordset
	Dim sNames
	Dim lErrorNumber

	If Len(oRequest("RemoveFile").Item) > 0 Then
		If FileExists(oRequest("FolderName").Item & "\" & oRequest("FileName").Item, sErrorDescription) Then
			lErrorNumber = DeleteFile(oRequest("FolderName").Item & "\" & oRequest("FileName").Item, sErrorDescription)
		End If
	ElseIf Len(oRequest("Add").Item) > 0 Then
		Select Case sAction
			Case "HistoryList"
			Case "Jobs"
				lErrorNumber = AddJob(oRequest, oADODBConnection, aJobComponent, True, sErrorDescription)
			Case "Positions"
				lErrorNumber = AddPosition(oRequest, oADODBConnection, aPositionComponent, sErrorDescription)
			Case Else
				'If (CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(6)) > 0) And (CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(6)) < Left(GetSerialNumberForDate(""), Len("00000000"))) Then aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11) = 0
				lErrorNumber = AddCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					lErrorNumber = ModifyPositionCatalogsLKP(oRequest, oADODBConnection, aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)), sErrorDescription)
				End If
		End Select
	ElseIf Len(oRequest("Modify").Item) > 0 Then
		Select Case sAction
			Case "HistoryList"
			Case "Jobs"
				lErrorNumber = ModifyJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
			Case "Positions"
				'lErrorNumber = ModifyPosition2(oRequest, oADODBConnection, aPositionComponent, sErrorDescription)
                lErrorNumber = ModifyPosition(oRequest, oADODBConnection, aPositionComponent, sErrorDescription)
			Case Else
				If (CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(6)) > 0) And (CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(6)) < Left(GetSerialNumberForDate(""), Len("00000000"))) Then aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11) = 0
				lErrorNumber = ModifyCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					lErrorNumber = ModifyPositionCatalogsLKP(oRequest, oADODBConnection, aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)), sErrorDescription)
				End If
		End Select
	ElseIf Len(oRequest("Remove").Item) > 0 Then
		If aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) > -1 Then
			lErrorNumber = GetCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
		End If
		Select Case sAction
			Case "HistoryList"
			Case "Jobs"
				lErrorNumber = RemoveJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
				If lErrorNumber = 0 Then aJobComponent(N_ID_JOB) = -1
			Case "Positions"
				lErrorNumber = RemovePosition(oRequest, oADODBConnection, aPositionComponent, sErrorDescription)
				aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) = -1
				aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11) = -2
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = ""
				sErrorDescription = "La información del puesto se ha borrado correctamente."
			Case Else
				lErrorNumber = RemoveCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					lErrorNumber = RemovePositionCatalogsLKP(oRequest, oADODBConnection, aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)), sErrorDescription)
				End If
		End Select
	ElseIf Len(oRequest("SetActive").Item) > 0 Then
		Select Case sAction
			Case "Jobs"
				lErrorNumber = SetActiveForJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
			Case Else
				lErrorNumber = SetActiveForCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
		End Select
	End If

	Set oRecordset = Nothing
	DoPositionsAction = lErrorNumber
	Err.Clear
End Function

Function DisplayPosition(oADODBConnection, lPositionID,lStartDate, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the search HTML form
'Inputs:  oADODBConnection, lPositionID, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPosition"
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información del puesto."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Positions.*, EmployeeTypeShortName, EmployeeTypeName, PositionTypeShortName, PositionTypeName From Positions, EmployeeTypes, PositionTypes Where (Positions.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Positions.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeeTypes.EndDate=30000000) And (PositionTypes.EndDate=30000000) And (PositionID=" & lPositionID & ") And (Positions.StartDate = "& lStartDate &")", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Información del puesto:</B></FONT><BR />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Clave:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringforHTML(CStr(oRecordset.Fields("PositionShortName").Value)) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Nombre corto:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringforHTML(CStr(oRecordset.Fields("PositionName").Value)) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Nombre largo:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringforHTML(CStr(oRecordset.Fields("PositionLongName").Value)) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Fecha de inicio:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplaydateFromSerialNumber(CStr(oRecordset.Fields("StartDate").Value), -1, -1, -1) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Fecha de término:&nbsp;</B></FONT></TD>"
					If CStr(oRecordset.Fields("EndDate").Value) = 30000000 Then
                        Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML("A la fecha") & "</FONT></TD>"
                    Else
                        Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplaydateFromSerialNumber(CStr(oRecordset.Fields("EndDate").Value), -1, -1, -1) & "</FONT></TD>"
                    End If
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Tipo de tabulador:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringforHTML(CStr(oRecordset.Fields("EmployeeTypeShortName").Value) & ". " & CStr(oRecordset.Fields("EmployeeTypeName").Value)) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Tipo de puesto:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringforHTML(CStr(oRecordset.Fields("PositionTypeShortName").Value) & ". " & CStr(oRecordset.Fields("PositionTypeName").Value)) & "</FONT></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE><BR />"
			oRecordset.Close

			Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Validaciones:</B></FONT><BR />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				sErrorDescription = "No se pudo obtener la información del puesto."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CompanyShortName, CompanyName From PositionsCatalogsLKP, Companies Where (PositionsCatalogsLKP.RecordID=Companies.CompanyID) And (PositionID=" & lPositionID & ") And (CatalogID=2) Order By CompanyShortName", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						Response.Write "<TR>"
							Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><B>Empresas:&nbsp;</B></FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
								Do While Not oRecordset.EOF
									Response.Write CleanStringforHTML(CStr(oRecordset.Fields("CompanyName").Value)) & "<BR />"
									oRecordset.MoveNext
									If Err.number <> 0 Then Exit Do
								Loop
							Response.Write "</FONT></TD>"
						Response.Write "</TR>"
					End If
				End If

				sErrorDescription = "No se pudo obtener la información del puesto."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CenterTypeShortName, CenterTypeName From PositionsCatalogsLKP, CenterTypes Where (PositionsCatalogsLKP.RecordID=CenterTypes.CenterTypeID) And (PositionID=" & lPositionID & ") And (CatalogID=1) Order By CenterTypeShortName", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						Response.Write "<TR>"
							Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><B>Tipos de centro de trabajo:&nbsp;</B></FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
								Do While Not oRecordset.EOF
									Response.Write CleanStringforHTML(CStr(oRecordset.Fields("CenterTypeShortName").Value) & ". " & CStr(oRecordset.Fields("CenterTypeName").Value)) & "<BR />"
									oRecordset.MoveNext
									If Err.number <> 0 Then Exit Do
								Loop
							Response.Write "</FONT></TD>"
						Response.Write "</TR>"
					End If
				End If

				sErrorDescription = "No se pudo obtener la información del puesto."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select GroupGradeLevelName From PositionsCatalogsLKP, GroupGradeLevels Where (PositionsCatalogsLKP.RecordID=GroupGradeLevels.GroupGradeLevelID) And (PositionID=" & lPositionID & ") And (CatalogID=3) Order By GroupGradeLevelName", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						Response.Write "<TR>"
							Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><B>Grupo, grado, nivel:&nbsp;</B></FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
								Do While Not oRecordset.EOF
									Response.Write CleanStringforHTML(CStr(oRecordset.Fields("GroupGradeLevelName").Value)) & "<BR />"
									oRecordset.MoveNext
									If Err.number <> 0 Then Exit Do
								Loop
							Response.Write "</FONT></TD>"
						Response.Write "</TR>"
					End If
				End If

				sErrorDescription = "No se pudo obtener la información del puesto."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JourneyShortName, JourneyName From PositionsCatalogsLKP, Journeys Where (PositionsCatalogsLKP.RecordID=Journeys.JourneyID) And (PositionID=" & lPositionID & ") And (CatalogID=4)", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						Response.Write "<TR>"
							Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><B>Turnos:&nbsp;</B></FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
								Do While Not oRecordset.EOF
									Response.Write CleanStringforHTML(CStr(oRecordset.Fields("JourneyShortName").Value) & ". " & CStr(oRecordset.Fields("JourneyName").Value)) & "<BR />"
									oRecordset.MoveNext
									If Err.number <> 0 Then Exit Do
								Loop
							Response.Write "</FONT></TD>"
						Response.Write "</TR>"
					End If
				End If

				sErrorDescription = "No se pudo obtener la información del puesto."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ShiftShortName, ShiftName From PositionsCatalogsLKP, Shifts Where (PositionsCatalogsLKP.RecordID=Shifts.ShiftID) And (PositionID=" & lPositionID & ") And (CatalogID=7) Order By ShiftShortName", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						Response.Write "<TR>"
							Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><B>Horarios:&nbsp;</B></FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
								Do While Not oRecordset.EOF
									Response.Write CleanStringforHTML(CStr(oRecordset.Fields("ShiftShortName").Value) & ". " & CStr(oRecordset.Fields("ShiftName").Value)) & "<BR />"
									oRecordset.MoveNext
									If Err.number <> 0 Then Exit Do
								Loop
							Response.Write "</FONT></TD>"
						Response.Write "</TR>"
					End If
				End If

				sErrorDescription = "No se pudo obtener la información del puesto."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ServiceShortName, ServiceName From PositionsCatalogsLKP, Services Where (PositionsCatalogsLKP.RecordID=Services.ServiceID) And (PositionID=" & lPositionID & ") And (CatalogID=6) Order By ServiceShortName", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						Response.Write "<TR>"
							Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><B>Servicios:&nbsp;</B></FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
								Do While Not oRecordset.EOF
									Response.Write CleanStringforHTML(CStr(oRecordset.Fields("ServiceShortName").Value) & ". " & CStr(oRecordset.Fields("ServiceName").Value)) & "<BR />"
									oRecordset.MoveNext
									If Err.number <> 0 Then Exit Do
								Loop
							Response.Write "</FONT></TD>"
						Response.Write "</TR>"
					End If
				End If

				sErrorDescription = "No se pudo obtener la información del puesto."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select LevelName From PositionsCatalogsLKP, Levels Where (PositionsCatalogsLKP.RecordID=Levels.LevelID) And (PositionID=" & lPositionID & ") And (CatalogID=5) Order By LevelName", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						Response.Write "<TR>"
							Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><B>Niveles:&nbsp;</B></FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
								Do While Not oRecordset.EOF
									Response.Write CleanStringforHTML(CStr(oRecordset.Fields("LevelName").Value)) & "<BR />"
									oRecordset.MoveNext
									If Err.number <> 0 Then Exit Do
								Loop
							Response.Write "</FONT></TD>"
						Response.Write "</TR>"
					End If
				End If

				sErrorDescription = "No se pudo obtener la información del puesto."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EconomicZoneName From PositionsCatalogsLKP, EconomicZones Where (PositionsCatalogsLKP.RecordID=EconomicZones.EconomicZoneID) And (PositionID=" & lPositionID & ") And (CatalogID=8) Order By EconomicZoneName", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						Response.Write "<TR>"
							Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><B>Zonas económicas:&nbsp;</B></FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
								Do While Not oRecordset.EOF
									Response.Write CleanStringforHTML(CStr(oRecordset.Fields("EconomicZoneName").Value)) & "<BR />"
									oRecordset.MoveNext
									If Err.number <> 0 Then Exit Do
								Loop
							Response.Write "</FONT></TD>"
						Response.Write "</TR>"
					End If
				End If
			Response.Write "</TABLE>"
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "El puesto especificado no está registrado en el sistema."
		End If
	End If

	Set oRecordset = Nothing
	DisplayPosition = lErrorNumber
	Err.Clear
End Function

Function DisplayPositionCompact(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
'************************************************************
'Purpose: To display the search HTML form
'Inputs:  oRequest, oADODBConnection, bFull
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPositionCompact"
	Dim oRecordset
	Dim sNames

	Response.Write "<TABLE WIDTH=""200"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		Response.Write "<TR BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2"" COLOR=""#" & S_MENU_LINK_FOR_GUI & """><B>" & CleanStringForHTML(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2)) & "</B></FONT></TD></TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Clave:&nbsp;</FONT></TD>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1)) & "</FONT></TD>"
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo tabulador:&nbsp;</FONT></TD>"
			Call GetNameFromTable(oADODBConnection, "EmployeeTypes", aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(7), "", "", sNames, sErrorDescription)
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo puesto:&nbsp;</FONT></TD>"
			Call GetNameFromTable(oADODBConnection, "PositionTypes", aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(8), "", "", sNames, sErrorDescription)
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
		Response.Write "</TR>"
		Response.Write "<TR><TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(3)) & "</FONT></TD></TR>"

		sErrorDescription = "No se pudo guardar la información del nuevo registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(JobID) As JobsCount From Jobs Where (PositionID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) & ")", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Plazas ocupadas:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						If Not IsNull(oRecordset.Fields("PositionCount")) Then
							Response.Write FormatNumber(CLng(oRecordset.Fields("JobsCount").Value), 0, True, False, True)
						Else
							Response.Write "0"
						End If
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
			End If
			oRecordset.Close
		End If

		sErrorDescription = "No se pudo guardar la información del nuevo registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(JobsInArea) As SumOfJobs From AreasPositionsLKP Where (PositionID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) & ")", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Plazas definidas:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						If Not IsNull(oRecordset.Fields("PositionCount")) Then
							Response.Write FormatNumber(CLng(oRecordset.Fields("SumOfJobs").Value), 0, True, False, True)
						Else
							Response.Write "0"
						End If
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
			End If
			oRecordset.Close
		End If
	Response.Write "</TABLE>"

	Set oRecordset = Nothing
	DisplayPositionCompact = lErrorNumber
	Err.Clear
End Function

Function DisplayPositionsCatalogsForm(oRequest, oADODBConnection, sAction, lPositionID, sErrorDescription)
'************************************************************
'Purpose: To display the Position and Catalogs Form
'Inputs:  oRequest, oADODBConnection, sAction, lPositionID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPositionsCatalogsForm"
	Dim oRecordset
	Dim sIDs
	Dim lErrorNumber

	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		Response.Write "function UpdateCatalogForm() {" & vbNewLine
			Response.Write "var oForm = document.CatalogFrm;" & vbNewLine

			Response.Write "if (oForm) {" & vbNewLine
				Response.Write "oForm.innerHTML += '<INPUT TYPE=""HIDDEN"" NAME=""CompanyIDs"" />';" & vbNewLine
				Response.Write "oForm.innerHTML += '<INPUT TYPE=""HIDDEN"" NAME=""CenterTypeIDs"" />';" & vbNewLine
				Response.Write "oForm.innerHTML += '<INPUT TYPE=""HIDDEN"" NAME=""GroupGradeLevelIDs"" />';" & vbNewLine
				Response.Write "oForm.innerHTML += '<INPUT TYPE=""HIDDEN"" NAME=""JourneyIDs"" />';" & vbNewLine
				Response.Write "oForm.innerHTML += '<INPUT TYPE=""HIDDEN"" NAME=""ShiftIDs"" />';" & vbNewLine
				Response.Write "oForm.innerHTML += '<INPUT TYPE=""HIDDEN"" NAME=""ServiceIDs"" />';" & vbNewLine
				Response.Write "oForm.innerHTML += '<INPUT TYPE=""HIDDEN"" NAME=""LevelIDs"" />';" & vbNewLine
				Response.Write "oForm.innerHTML += '<INPUT TYPE=""HIDDEN"" NAME=""EconomicZoneIDs"" />';" & vbNewLine
			Response.Write "}" & vbNewLine
		Response.Write "} // End of UpdateCatalogForm" & vbNewLine
		
		Response.Write "function GetCatalogsIDs() {" & vbNewLine
			Response.Write "var oSourceForm = document.CatalogsLKPFrm;" & vbNewLine
			Response.Write "var oTargetForm = document.CatalogFrm;" & vbNewLine

			Response.Write "if (oSourceForm && oTargetForm) {" & vbNewLine
				Response.Write "oTargetForm.CompanyIDs.value = GetCheckBoxSelection(oSourceForm.CompanyIDs);" & vbNewLine
				Response.Write "oTargetForm.CenterTypeIDs.value = GetCheckBoxSelection(oSourceForm.CenterTypeIDs);" & vbNewLine
				Response.Write "oTargetForm.GroupGradeLevelIDs.value = GetCheckBoxSelection(oSourceForm.GroupGradeLevelIDs);" & vbNewLine
				Response.Write "oTargetForm.JourneyIDs.value = GetCheckBoxSelection(oSourceForm.JourneyIDs);" & vbNewLine
				Response.Write "oTargetForm.ShiftIDs.value = GetCheckBoxSelection(oSourceForm.ShiftIDs);" & vbNewLine
				Response.Write "oTargetForm.ServiceIDs.value = GetCheckBoxSelection(oSourceForm.ServiceIDs);" & vbNewLine
				Response.Write "oTargetForm.LevelIDs.value = GetCheckBoxSelection(oSourceForm.LevelIDs);" & vbNewLine
				Response.Write "oTargetForm.EconomicZoneIDs.value = GetCheckBoxSelection(oSourceForm.EconomicZoneIDs);" & vbNewLine
			Response.Write "}" & vbNewLine
		Response.Write "} // End of GetCatalogsIDs" & vbNewLine
		
		Response.Write "UpdateCatalogForm();" & vbNewLine
	Response.Write "//--></SCRIPT>" & vbNewLine

	Response.Write "<FORM NAME=""CatalogsLKPFrm"" ID=""CatalogsLKPFrm"" onSubmit=""return false"">" & vbNewLine
		Response.Write "<B>Empresas:</B><BR />"
		Response.Write "<DIV NAME=""CompanyIDsDiv"" ID=""CompanyIDsDiv"" CLASS=""CheckboxList"" STYLE=""width: 620px; height: 102px""><FONT FACE=""Arial"" SIZE=""2"">"
			sIDs = ""
			sErrorDescription = "No se pudo guardar la información del puesto."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select RecordID From PositionsCatalogsLKP Where (PositionID=" & lPositionID & ") And (CatalogID=2)", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				sIDs = sIDs & CStr(oRecordset.Fields("RecordID").Value) & LIST_SEPARATOR
			End If
			Call GenerateCheckboxesFromQuery(oADODBConnection, "Companies", "CompanyID", "CompanyShortName, CompanyName", "(ParentID>-1) And (Active=1)", "CompanyShortName", sIDs, "CompanyIDs", sErrorDescription)
		Response.Write "</DIV><BR />"

		Response.Write "<B>Tipos de centro de trabajo:</B><BR />"
		Response.Write "<DIV NAME=""CenterTypeIDsDiv"" ID=""CenterTypeIDsDiv"" CLASS=""CheckboxList"" STYLE=""width: 620px; height: 102px""><FONT FACE=""Arial"" SIZE=""2"">"
			sIDs = ""
			sErrorDescription = "No se pudo guardar la información del puesto."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select RecordID From PositionsCatalogsLKP Where (PositionID=" & lPositionID & ") And (CatalogID=1)", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				sIDs = sIDs & CStr(oRecordset.Fields("RecordID").Value) & LIST_SEPARATOR
			End If
			Call GenerateCheckboxesFromQuery(oADODBConnection, "CenterTypes", "CenterTypeID", "CenterTypeShortName, CenterTypeName", "(Active=1)", "CenterTypeShortName, CenterTypeName", sIDs, "CenterTypeIDs", sErrorDescription)
		Response.Write "</DIV><BR />"

		Response.Write "<B>Grupo, grado, nivel:</B><BR />"
		Response.Write "<DIV NAME=""GroupGradeLevelIDsDiv"" ID=""GroupGradeLevelIDsDiv"" CLASS=""CheckboxList"" STYLE=""width: 620px; height: 102px""><FONT FACE=""Arial"" SIZE=""2"">"
			sIDs = ""
			sErrorDescription = "No se pudo guardar la información del puesto."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select RecordID From PositionsCatalogsLKP Where (PositionID=" & lPositionID & ") And (CatalogID=3)", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				sIDs = sIDs & CStr(oRecordset.Fields("RecordID").Value) & LIST_SEPARATOR
			End If
			Call GenerateCheckboxesFromQuery(oADODBConnection, "GroupGradeLevels", "GroupGradeLevelID", "GroupGradeLevelName", "(Active=1)", "GroupGradeLevelName", sIDs, "GroupGradeLevelIDs", sErrorDescription)
		Response.Write "</DIV><BR />"

		Response.Write "<B>Turnos:</B><BR />"
		Response.Write "<DIV NAME=""JourneyIDsDiv"" ID=""JourneyIDsDiv"" CLASS=""CheckboxList"" STYLE=""width: 620px; height: 102px""><FONT FACE=""Arial"" SIZE=""2"">"
			sIDs = ""
			sErrorDescription = "No se pudo guardar la información del puesto."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select RecordID From PositionsCatalogsLKP Where (PositionID=" & lPositionID & ") And (CatalogID=4)", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				sIDs = sIDs & CStr(oRecordset.Fields("RecordID").Value) & LIST_SEPARATOR
			End If
			Call GenerateCheckboxesFromQuery(oADODBConnection, "Journeys", "JourneyID", "JourneyShortName, JourneyName", "(Active=1)", "JourneyShortName, JourneyName", sIDs, "JourneyIDs", sErrorDescription)
		Response.Write "</DIV><BR />"

		Response.Write "<B>Horarios:</B><BR />"
		Response.Write "<DIV NAME=""ShiftIDsDiv"" ID=""ShiftIDsDiv"" CLASS=""CheckboxList"" STYLE=""width: 620px; height: 102px""><FONT FACE=""Arial"" SIZE=""2"">"
			sIDs = ""
			sErrorDescription = "No se pudo guardar la información del puesto."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select RecordID From PositionsCatalogsLKP Where (PositionID=" & lPositionID & ") And (CatalogID=7)", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				sIDs = sIDs & CStr(oRecordset.Fields("RecordID").Value) & LIST_SEPARATOR
			End If
			Call GenerateCheckboxesFromQuery(oADODBConnection, "Shifts", "ShiftID", "ShiftShortName, ShiftName", "(Active=1)", "ShiftShortName, ShiftName", sIDs, "ShiftIDs", sErrorDescription)
		Response.Write "</DIV><BR />"

		Response.Write "<B>Servicios:</B><BR />"
		Response.Write "<DIV NAME=""ServiceIDsDiv"" ID=""ServiceIDsDiv"" CLASS=""CheckboxList"" STYLE=""width: 620px; height: 102px""><FONT FACE=""Arial"" SIZE=""2"">"
			sIDs = ""
			sErrorDescription = "No se pudo guardar la información del puesto."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select RecordID From PositionsCatalogsLKP Where (PositionID=" & lPositionID & ") And (CatalogID=6)", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				sIDs = sIDs & CStr(oRecordset.Fields("RecordID").Value) & LIST_SEPARATOR
			End If
			Call GenerateCheckboxesFromQuery(oADODBConnection, "Services", "ServiceID", "ServiceShortName, ServiceName", "(Active=1)", "ServiceShortName, ServiceName", sIDs, "ServiceIDs", sErrorDescription)
		Response.Write "</DIV><BR />"

		Response.Write "<B>Niveles:</B><BR />"
		Response.Write "<DIV NAME=""LevelIDsDiv"" ID=""LevelIDsDiv"" CLASS=""CheckboxList"" STYLE=""width: 620px; height: 102px""><FONT FACE=""Arial"" SIZE=""2"">"
			sIDs = ""
			sErrorDescription = "No se pudo guardar la información del puesto."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select RecordID From PositionsCatalogsLKP Where (PositionID=" & lPositionID & ") And (CatalogID=5)", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				sIDs = sIDs & CStr(oRecordset.Fields("RecordID").Value) & LIST_SEPARATOR
			End If
			Call GenerateCheckboxesFromQuery(oADODBConnection, "Levels", "LevelID", "LevelName", "(Active=1)", "LevelName", sIDs, "LevelIDs", sErrorDescription)
		Response.Write "</DIV><BR />"


		Response.Write "<B>Zonas económicas:</B><BR />"
		Response.Write "<DIV NAME=""EconomicZoneIDsDiv"" ID=""EconomicZoneIDsDiv"" CLASS=""CheckboxList"" STYLE=""width: 620px; height: 42px""><FONT FACE=""Arial"" SIZE=""2"">"
			sIDs = ""
			sErrorDescription = "No se pudo guardar la información del puesto."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select RecordID From PositionsCatalogsLKP Where (PositionID=" & lPositionID & ") And (CatalogID=5)", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				sIDs = sIDs & CStr(oRecordset.Fields("RecordID").Value) & LIST_SEPARATOR
			End If
			Call GenerateCheckboxesFromQuery(oADODBConnection, "EconomicZones", "EconomicZoneID", "EconomicZoneName", "(EconomicZoneID>-1) And (Active=1)", "EconomicZoneName", sIDs, "EconomicZoneIDs", sErrorDescription)
		Response.Write "</DIV><BR />"
	Response.Write "</FORM>"

	Set oRecordset = Nothing
	DisplayPositionsCatalogsForm = lErrorNumber
	Err.Clear
End Function

Function DisplayPositionsSearchForm(oRequest, oADODBConnection, bFull, sErrorDescription)
'************************************************************
'Purpose: To display the search HTML form
'Inputs:  oRequest, oADODBConnection, bFull
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPositionsSearchForm"

	Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""Positions.asp"" METHOD=""GET"">"
		Response.Write "<TABLE BORDER=""0"" CELLPADING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Clave del puesto:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""PositionShortName"" ID=""PositionShortNameTxt"" SIZE=""10"" MAXLENGTH=""10"" CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre del área:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""PositionName"" ID=""PositionNameTxt"" SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			If bFull Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de nacimiento:&nbsp;</FONT></TD>"
					Response.Write "<TD>"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
						Response.Write DisplayDateCombos(CInt(oRequest("StartBirthYear").Item), CInt(oRequest("StartBirthMonth").Item), CInt(oRequest("StartBirthDay").Item), "StartBirthYear", "StartBirthMonth", "StartBirthDay", N_FORM_START_YEAR, Year(Date()), True, True)
						Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
						Response.Write DisplayDateCombos(CInt(oRequest("EndBirthYear").Item), CInt(oRequest("EndBirthMonth").Item), CInt(oRequest("EndBirthDay").Item), "EndBirthYear", "EndBirthMonth", "EndBirthDay", N_FORM_START_YEAR, Year(Date()), True, True)
					Response.Write "</TD>"
				Response.Write "</TR>"
			End If
			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""2"""
				If Not bFull Then Response.Write " ALIGN=""RIGHT"""
				Response.Write "><BR /><INPUT TYPE=""SUBMIT"" NAME=""DoSearch"" ID=""DoSearchBtn"" VALUE=""Buscar Puestos"" CLASS=""Buttons"" /></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	Response.Write "</FORM></TD>"

	DisplayPositionsSearchForm = Err.number
End Function

Function ModifyPositionCatalogsLKP(oRequest, oADODBConnection, lPositionID, sErrorDescription)
'************************************************************
'Purpose: To modify the PositionsCatalogsLKP table
'Inputs:  oRequest, oADODBConnection, lPositionID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyPositionCatalogsLKP"
	Dim sIDs
	Dim asIDs
	Dim iIndex
	Dim lErrorNumber

	sErrorDescription = "No se pudo guardar la información del puesto."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From PositionsCatalogsLKP Where (PositionID=" & lPositionID & ")", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	If lErrorNumber = 0 Then
		sIDs = ""
		sIDs = oRequest("CenterTypeIDs").Item
		If Len(sIDs) > 0 Then
			asIDs = Split(Replace(sIDs, "-760211" & LIST_SEPARATOR, ""), LIST_SEPARATOR)
			For iIndex = 0 To UBound(asIDs)
				sErrorDescription = "No se pudo guardar la información del puesto."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PositionsCatalogsLKP (PositionID, CatalogID, RecordID) Values (" & lPositionID & ", 1, " & asIDs(iIndex) & ")", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			Next
		End If

		sIDs = ""
		sIDs = oRequest("CompanyIDs").Item
		If Len(sIDs) > 0 Then
			asIDs = Split(Replace(sIDs, "-760211" & LIST_SEPARATOR, ""), LIST_SEPARATOR)
			For iIndex = 0 To UBound(asIDs)
				sErrorDescription = "No se pudo guardar la información del puesto."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PositionsCatalogsLKP (PositionID, CatalogID, RecordID) Values (" & lPositionID & ", 2, " & asIDs(iIndex) & ")", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			Next
		End If

		sIDs = ""
		sIDs = oRequest("GroupGradeLevelIDs").Item
		If Len(sIDs) > 0 Then
			asIDs = Split(Replace(sIDs, "-760211" & LIST_SEPARATOR, ""), LIST_SEPARATOR)
			For iIndex = 0 To UBound(asIDs)
				sErrorDescription = "No se pudo guardar la información del puesto."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PositionsCatalogsLKP (PositionID, CatalogID, RecordID) Values (" & lPositionID & ", 3, " & asIDs(iIndex) & ")", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			Next
		End If

		sIDs = ""
		sIDs = oRequest("JourneyIDs").Item
		If Len(sIDs) > 0 Then
			asIDs = Split(Replace(sIDs, "-760211" & LIST_SEPARATOR, ""), LIST_SEPARATOR)
			For iIndex = 0 To UBound(asIDs)
				sErrorDescription = "No se pudo guardar la información del puesto."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PositionsCatalogsLKP (PositionID, CatalogID, RecordID) Values (" & lPositionID & ", 4, " & asIDs(iIndex) & ")", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			Next
		End If

		sIDs = ""
		sIDs = oRequest("LevelIDs").Item
		If Len(sIDs) > 0 Then
			asIDs = Split(Replace(sIDs, "-760211" & LIST_SEPARATOR, ""), LIST_SEPARATOR)
			For iIndex = 0 To UBound(asIDs)
				sErrorDescription = "No se pudo guardar la información del puesto."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PositionsCatalogsLKP (PositionID, CatalogID, RecordID) Values (" & lPositionID & ", 5, " & asIDs(iIndex) & ")", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			Next
		End If

		sIDs = ""
		sIDs = oRequest("ServiceIDs").Item
		If Len(sIDs) > 0 Then
			asIDs = Split(Replace(sIDs, "-760211" & LIST_SEPARATOR, ""), LIST_SEPARATOR)
			For iIndex = 0 To UBound(asIDs)
				sErrorDescription = "No se pudo guardar la información del puesto."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PositionsCatalogsLKP (PositionID, CatalogID, RecordID) Values (" & lPositionID & ", 6, " & asIDs(iIndex) & ")", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			Next
		End If

		sIDs = ""
		sIDs = oRequest("ShiftIDs").Item
		If Len(sIDs) > 0 Then
			asIDs = Split(Replace(sIDs, "-760211" & LIST_SEPARATOR, ""), LIST_SEPARATOR)
			For iIndex = 0 To UBound(asIDs)
				sErrorDescription = "No se pudo guardar la información del puesto."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PositionsCatalogsLKP (PositionID, CatalogID, RecordID) Values (" & lPositionID & ", 7, " & asIDs(iIndex) & ")", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			Next
		End If

		sIDs = ""
		sIDs = oRequest("EconomicZoneIDs").Item
		If Len(sIDs) > 0 Then
			asIDs = Split(Replace(sIDs, "-760211" & LIST_SEPARATOR, ""), LIST_SEPARATOR)
			For iIndex = 0 To UBound(asIDs)
				sErrorDescription = "No se pudo guardar la información del puesto."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PositionsCatalogsLKP (PositionID, CatalogID, RecordID) Values (" & lPositionID & ", 8, " & asIDs(iIndex) & ")", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			Next
		End If
	End If

	ModifyPositionCatalogsLKP = lErrorNumber
	Err.Clear
End Function

Function RemovePositionCatalogsLKP(oRequest, oADODBConnection, lPositionID, sErrorDescription)
'************************************************************
'Purpose: To modify the PositionsCatalogsLKP table
'Inputs:  oRequest, oADODBConnection, lPositionID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemovePositionCatalogsLKP"
	Dim lErrorNumber

	sErrorDescription = "No se pudo eliminar la información del puesto."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From PositionsCatalogsLKP Where (PositionID=" & lPositionID & ")", "PositionsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

	RemovePositionCatalogsLKP = lErrorNumber
	Err.Clear
End Function
%>