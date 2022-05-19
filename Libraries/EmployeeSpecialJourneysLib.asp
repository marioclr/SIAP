<%
Function GetSpecialJourneysURLValues(oRequest, iSelectedTab, bAction, sCondition)
'************************************************************
'Purpose: To initialize the global variables using the URL
'Inputs:  oRequest
'Outputs: iSelectedTab, bAction, sCondition
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetSpecialJourneysURLValues"
	Dim oItem
	Dim aItem

	iSelectedTab = 1
	If Not IsEmpty(oRequest("Tab").Item) Then
		iSelectedTab = CInt(oRequest("Tab").Item)
	End If
	bAction = (Len(oRequest("Add").Item) > 0) Or (Len(oRequest("Modify").Item) > 0) Or (Len(oRequest("Remove").Item) > 0) Or (Len(oRequest("SetActive").Item) > 0) Or (Len(oRequest("Authorization").Item) > 0) Or (Len(oRequest("SetDeActive").Item) > 0)

	sCondition = ""

	GetSpecialJourneysURLValues = Err.number
	Err.Clear
End Function

Function DoSpecialJourneysAction(oRequest, oADODBConnection, sAction, iSpecialJourneyType, sErrorDescription)
'************************************************************
'Purpose: To add, change or delete the information of the
'         specified component
'Inputs:  oRequest, oADODBConnection, sAction
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DoSpecialJourneysAction"
	Dim oRecordset
	Dim sNames
	Dim lErrorNumber

	If Len(oRequest("RemoveFile").Item) > 0 Then
		If FileExists(oRequest("FolderName").Item & "\" & oRequest("FileName").Item, sErrorDescription) Then
			lErrorNumber = DeleteFile(oRequest("FolderName").Item & "\" & oRequest("FileName").Item, sErrorDescription)
		End If
	ElseIf Len(oRequest("Add").Item) > 0 Then
		Select Case sAction
			Case "ExternalSpecialJourney"
				lErrorNumber = AddExternalEmployee(oRequest, oADODBConnection, aSpecialJourneyComponent, sErrorDescription)
			Case "Jobs"
				lErrorNumber = AddJob(oRequest, oADODBConnection, aJobComponent, True, sErrorDescription)
			Case "SpecialJourney"
				'lErrorNumber = AddSpecialJourney(oRequest, oADODBConnection, aSpecialJourneyComponent, sErrorDescription)
				lErrorNumber = AddSpecialJourney(oRequest, oADODBConnection, iSpecialJourneyType, aSpecialJourneyComponent, sErrorDescription)
		End Select
	ElseIf Len(oRequest("Modify").Item) > 0 Then
		Select Case sAction
			Case "AreaPositions"
				If aAreaComponent(N_ID_AREA) > -1 Then
					lErrorNumber = ModifyAreaPositions(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
				End If
			Case "ExternalSpecialJourney"
				lErrorNumber = ModifyExternalEmployee(oRequest, oADODBConnection, aSpecialJourneyComponent, sErrorDescription)
			Case "Jobs"
				lErrorNumber = ModifyJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
			Case "SpecialJourney"
				lErrorNumber = ModifySpecialJourney(oRequest, oADODBConnection, aSpecialJourneyComponent, sErrorDescription)
		End Select
	ElseIf Len(oRequest("Remove").Item) > 0 Then
		If aAreaComponent(N_ID_AREA) > -1 Then
			lErrorNumber = GetArea(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
		End If
		Select Case sAction
			Case "HistoryList"
			Case "Jobs"
				lErrorNumber = RemoveJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
				If lErrorNumber = 0 Then aJobComponent(N_ID_JOB) = -1
			Case "ExternalSpecialJourney"
				lErrorNumber = RemoveExternalEmployee(oRequest, oADODBConnection, aSpecialJourneyComponent, sErrorDescription)
			Case "SpecialJourney"
				lErrorNumber = RemoveSpecialJourney(oRequest, oADODBConnection, aSpecialJourneyComponent, sErrorDescription)
		End Select
	ElseIf Len(oRequest("SetActive").Item) > 0 Then
		Select Case sAction
			Case "Jobs"
				lErrorNumber = SetActiveForJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
			Case "SpecialJourney"
				lErrorNumber = SetActiveForEmployeeSpecialJourney(oRequest, oADODBConnection, aSpecialJourneyComponent, 1, sErrorDescription)
		End Select
	ElseIf Len(oRequest("SetDeActive").Item) > 0 Then
		Select Case sAction
			Case "Jobs"
				lErrorNumber = SetActiveForJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
			Case "SpecialJourney"
				lErrorNumber = SetActiveForEmployeeSpecialJourney(oRequest, oADODBConnection, aSpecialJourneyComponent, 0, sErrorDescription)
		End Select
	ElseIf Len(oRequest("Authorization").Item) > 0 Then
		Select Case sAction
			Case "Jobs"
				lErrorNumber = SetActiveForJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
			Case "SpecialJourney"
				'lErrorNumber = SetActiveForEmployeeSpecialJourney(oRequest, oADODBConnection, aSpecialJourneyComponent, sErrorDescription)
				lErrorNumber = 0
		End Select
	End If

	Set oRecordset = Nothing
	DoSpecialJourneysAction = lErrorNumber
	Err.Clear
End Function

Function DisplaySpecialJourneysFilters(oRequest, sAction, sErrorDescription)
'************************************************************
'Purpose: To display the filter of the specified catalog.
'Inputs:  sAction
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplaySpecialJourneysFilters"
	Dim bHasFilter
	Dim iIndex
	Dim lErrorNumber
    Dim sParentCondition
    Dim sZoneFilter
    Dim sCodeFilter
    Dim sShortFilter
    Dim sPCTypeFilter

	bHasFilter = False
    sParentCondition = ""
    sZoneFilter = ""
    sCodeFilter = ""
    sShortFilter = ""
    sPCTypeFilter = ""
	Response.Write "<FORM NAME=""FilterFrm"" ID=""FilterFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""GET"" /><FONT FACE=""Arial"" SIZE=""2"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & sAction & """ />"
		Select Case sAction
			Case "Areas"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Mostrar áreas que contengan:&nbsp;</B>"
				Response.Write "<INPUT TYPE=""TEXT"" NAME=""FilterName"" ID=""FilterNameTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & oRequest("FilterName").Item & """ CLASS=""TextFields"" />"
				bHasFilter = True
			Case "FormFields"
			Case "Forms"
			Case "Profiles"
			Case "Users"
			Case "Zones"
			Case Else
                If CInt(oRequest("ParentID").Item) > 0 Then
                    sParentCondition = " And (ParentID=" & CStr(oRequest("ParentID").Item) & ")"
                End If
                If Len(oRequest("ApplyFilter").Item) > 0 Then
                    sZoneFilter = CInt(oRequest("ZoneID").Item)
                    sCodeFilter = CStr(oRequest("AreaCode").Item)
                    sShortFilter = CStr(oRequest("AreaShortName").Item)
                    sPCTypeFilter = CStr(oRequest("CenterTypeID").Item)
                End If

				Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" /><FONT FACE=""Arial"" SIZE=""2""><B>Mostrar centros de trabajo de la entidad:&nbsp;</B>"
					Response.Write "<SELECT NAME=""ZoneID"" ID=""ZoneIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						'Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Zones", "ZoneID", "ZoneCode, ZoneName", "(Active=1) And (ParentID=-1) And (ZoneID>-1)", "ZoneCode", asPath(2), "Ninguna;;;-1", sErrorDescription)
                        Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Zones", "ZoneID", "ZoneCode, ZoneName", "(Active=1) And (ParentID=-1) And (ZoneID>-1)", "ZoneCode", sZoneFilter, "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT>"
				Response.Write "<BR />"

				Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" /><B>Mostrar los centros de trabajo con clave:&nbsp;</B>"
					Response.Write "<SELECT NAME=""AreaCode"" ID=""AreaShortNameCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE=""-2"">Todos</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "Distinct AreaCode", "AreaCode As AreaCode2", "ParentID > 0 And AreaCode IS NOT NULL AND AreaCode <> ' '" & sParentCondition, "AreaCode", sCodeFilter, "", sErrorDescription)
					Response.Write "</SELECT>"
					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""5"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" /><B>ó con clave de 10 digitos:&nbsp;</B>"
					Response.Write "<SELECT NAME=""AreaShortName"" ID=""AreaShortNameCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE=""-2"">Todos</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "Distinct AreaShortName", "AreaShortName As AreaShortName2", "ParentID > 0 And AreaShortName IS NOT NULL AND AreaShortName <> ' '" & sParentCondition, "AreaShortName", sShortFilter, "", sErrorDescription)
					Response.Write "</SELECT>"
				Response.Write "<BR />"

				Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" /><B>Mostrar los centros de trabajo con tipo de centro de trabajo:&nbsp;</B>"
					Response.Write "<SELECT NAME=""CenterTypeID"" ID=""CenterTypeIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""UpdateSubtypes(this.value)"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "CenterTypes", "CenterTypeID", "CenterTypeShortName, CenterTypeName", "(Active=1)", "CenterTypeShortName", "", "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT>"
                Response.Write "<BR />"
				bHasFilter = True
		End Select
		If bHasFilter Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ApplyFilter"" ID=""ApplyFilterBtn"" VALUE=""  Filtrar  "" CLASS=""Buttons"" />"
		Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
		If Len(oRequest("ApplyFilter").Item) = 0 Then
			If Len(oRequest("ParentID").Item) > 0 Then
				If Len(oRequest("AreaID").Item) > 0 Then
					Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?ParentID=" & oRequest("ParentID").Item & "'"" />"
				Else
					Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?ParentID=" & aAreaComponent(N_PARENT_ID2_AREA) & "'"" />"
				End If
			Else
				Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?ParentID=-1'"" />"
			End If
		Else
			If aAreaComponent(N_ID_AREA) <> -1 Then
                sFilter = "ApplyFilter=1&ZoneID=" & CStr(oRequest("ZoneID").Item) & "&AreaCode=" & CStr(oRequest("AreaCode").Item) & "&AreaShortName=" & CStr(oRequest("AreaShortName").Item) & "&CenterTypeID=" & CStr(oRequest("CenterTypeID").Item)
				Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?" & sFilter & "'"" />"
			Else
				Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?ParentID=-1'"" />"
			End If
		End If
		Response.Write "</FONT></FORM>"

	DisplaySpecialJourneysFilters = lErrorNumber
	Err.Clear
End Function

Function DisplaySpecialJourneysSearchForm(oRequest, oADODBConnection, bFull, sErrorDescription)
'************************************************************
'Purpose: To display the search HTML form
'Inputs:  oRequest, oADODBConnection, bFull
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplaySpecialJourneysSearchForm"

	Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""Areas.asp"" METHOD=""GET"">"
		Response.Write "<TABLE BORDER=""0"" CELLPADING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Clave del área:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""AreaShortName"" ID=""AreaShortNameTxt"" CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre del área:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""AreaName"" ID=""AreaNameTxt"" CLASS=""TextFields"" /></TD>"
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
				Response.Write "><BR /><INPUT TYPE=""SUBMIT"" NAME=""DoSearch"" ID=""DoSearchBtn"" VALUE=""Buscar Áreas"" CLASS=""Buttons"" /></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	Response.Write "</FORM></TD>"

	DisplaySpecialJourneysSearchForm = Err.number
End Function

Function DisplaySpecialJourneysTabs(oRequest, bError, sErrorDescription)
'************************************************************
'Purpose: To display the tabs for the areas HTML forms
'Inputs:  oRequest, bError
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplaySpecialJourneysTabs"
	Dim sNames
	Dim asTitles
	Dim iIndex
	Dim sAction
	Dim lErrorNumber

	Call GetNameFromTable(oADODBConnection, "AreaLevelTypes", aAreaComponent(N_LEVEL_TYPE_ID_AREA), "", "", sNames, sErrorDescription)
	asTitles = Split(",Información del " & sNames & ",Puestos,Historial de cambios,Historial de puestos", ",")
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
				If Len(asTitles(iIndex)) > 0 Then
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
					Response.Write "<A HREF=""" & GetASPFileName("") & "?Action=Areas&AreaID=" & aAreaComponent(N_ID_AREA) & "&" & sAction & "=1&Tab=" & iIndex & """ CLASS=""TabLink""><DIV NAME=""TabText" & iIndex & "Div"" ID=""TabText" & iIndex & "Div"" STYLE=""color: #"
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
				End If
			Next
			Response.Write "<TD BACKGROUND=""Images/TbBgDot.gif"" WIDTH=""*""><IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""21"" /></TD>"
		Response.Write "</TR></TABLE>"
	End If

	DisplaySpecialJourneysTabs = lErrorNumber
	Err.Clear
End Function
%>