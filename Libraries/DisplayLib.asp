<%
Dim asPriorityNames
asPriorityNames = Split(",1 (alta),2 (medio alta),3 (media),4 (medio baja),5 (baja)", ",", -1, vbBinaryCompare)

Function DisplayContactInformation(oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the contact information
'Inputs:  oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayContactInformation"
	Dim lErrorNumber

	Response.Write "<B>" & CleanStringForHTML(GetAdminOption(aAdminOptionsComponent, CONTACT_NAME_OPTION)) & "</B><BR />"
	Response.Write "Teléfono: " & CleanStringForHTML(GetAdminOption(aAdminOptionsComponent, CONTACT_PHONE_OPTION)) & "<BR />"
	Response.Write "Correo electrónico: <A HREF=""mailto: " & CleanStringForHTML(GetAdminOption(aAdminOptionsComponent, CONTACT_EMAIL_OPTION)) & """>" & CleanStringForHTML(GetAdminOption(aAdminOptionsComponent, CONTACT_EMAIL_OPTION)) & "</A><BR />"

	DisplayContactInformation = lErrorNumber
	Err.Clear
End Function

Function DisplayIncrementalFetch(oRequest, iStart, iRows, oRecordset)
'************************************************************
'Purpose: To display the page numbers for the incremental page
'Inputs:  oRequest, iStart, iRows, oRecordset
'************************************************************
	Dim iEntries
	Dim iTempEntries
	Dim iTotalEntries
	Dim iPages
	Dim iIndex

	iEntries = 0
	iTempEntries = 0
	If Not IsNull(oRecordset) Then iEntries = GetNumberOfEntriesFromRecordset(oRecordset)
	If Not IsNumeric(iStart) Or (iStart = 0) Then iStart = 1
	iTotalEntries = iEntries + iTempEntries
	iPages = CInt(iTotalEntries / iRows + 0.49999)

	If iTotalEntries > iRows Then
		If iPages > 20 Then Response.Write "<FORM>"
			Response.Write "<SPAN CLASS=""IncrementalFetch""><FONT FACE=""Arial"" SIZE=""2""><B>"
			Response.Write "Renglones " & iStart & " - "
			If (iStart + iRows - 1) < iTotalEntries Then
				Response.Write (iStart + iRows - 1)
			Else
				Response.Write iTotalEntries
			End If
			Response.Write " de " & iTotalEntries & ":</B>&nbsp;&nbsp;&nbsp;"
				If iPages < 20 Then
					For iIndex = 1 To iPages
						Response.Write "<A HREF=""" & GetASPFileName("") & "?" & ReplaceValueInURLString(oRequest, "StartPage", (iIndex - 1) * iRows + 1) & """>"
							If iIndex = (Int(iStart / iRows) + 1) Then
								Response.Write "<B>" & iIndex & "</B>"
							Else
								Response.Write iIndex
							End If
						Response.Write "</A>"
						If iIndex < iPages Then Response.Write "&nbsp;&nbsp;&#183;&nbsp;&nbsp;"
					Next
				Else
					Response.Write "<SELECT CLASS=""Lists"" onChange=""window.location.href='" & GetASPFileName("") & "?" & ReplaceValueInURLString(oRequest, "StartPage", "") & "' + this.value"">"
						For iIndex = 1 To iPages
							Response.Write "<OPTION VALUE=""" & ((iIndex - 1) * iRows + 1) & """"
								If iIndex = (Int(iStart / iRows) + 1) Then Response.Write " SELECTED=""1"""
							Response.Write ">" & iIndex & "</OPTION>"
						Next
					Response.Write "</SELECT>"
				End If
			Response.Write "</FONT></SPAN><BR /><BR />"
		If iPages > 20 Then Response.Write "</FORM>"

		iIndex = 1
		If Not IsNull(oRecordset) Then
			If Not oRecordset.EOF Then
				For iIndex = iIndex To iStart - 1
					oRecordset.MoveNext
					If oRecordset.EOF Then Exit For
					If Err.number <> 0 Then Exit For
				Next
			End If
		End If
    Else
        'Response.Write "<FORM id=form1 name=form1>"
		Response.Write "<SPAN CLASS=""IncrementalFetch""><FONT FACE=""Arial"" SIZE=""2""><B>"
		Response.Write "Renglones " & iTotalEntries
        Response.Write "</FONT></SPAN><BR /><BR />"
	End If

	DisplayIncrementalFetch = Err.number
	Err.Clear
End Function

Function DisplayIncrementalFetchForSections(oRequest, iStart, iRows, iSectionType, oRecordset)
'************************************************************
'Purpose: To display the page numbers for the incremental page
'Inputs:  oRequest, iStart, iRows, oRecordset
'************************************************************
	Dim iEntries
	Dim iTempEntries
	Dim iTotalEntries
	Dim iPages
	Dim iIndex

	iEntries = 0
	iTempEntries = 0
	If Not IsNull(oRecordset) Then iEntries = GetNumberOfEntriesFromRecordset(oRecordset)
	If Not IsNumeric(iStart) Or (iStart = 0) Then iStart = 1
	iTotalEntries = iEntries + iTempEntries
	iPages = CDbl(iTotalEntries / iRows + 0.49999)

	If iTotalEntries > iRows Then
		If iPages > 20 Then Response.Write "<FORM id=form1 name=form1>"
			Response.Write "<SPAN CLASS=""IncrementalFetch""><FONT FACE=""Arial"" SIZE=""2""><B>"
			Response.Write "Renglones " & iStart & " - "
			If (iStart + iRows - 1) < iTotalEntries Then
				Response.Write (iStart + iRows - 1)
			Else
				Response.Write iTotalEntries
			End If
			Response.Write " de " & iTotalEntries & ":</B>&nbsp;&nbsp;&nbsp;"
				If iPages < 20 Then
					For iIndex = 1 To iPages
						If Len(oRequest("RowsType").Item) > 0 Then
							Response.Write "<A HREF=""" & GetASPFileName("") & "?" & ReplaceValueInURLString(oRequest, "StartPage", (iIndex - 1) * iRows + 1) & """>"
						Else
							Response.Write "<A HREF=""" & GetASPFileName("") & "?" & ReplaceValueInURLString(oRequest, "StartPage", (iIndex - 1) * iRows + 1) & "&RowsType=" & iSectionType & """>"
						End If
							If iIndex = (Int(iStart / iRows) + 1) Then
								Response.Write "<B>" & iIndex & "</B>"
							Else
								Response.Write iIndex
							End If
						Response.Write "</A>"
						If iIndex < iPages Then Response.Write "&nbsp;&nbsp;&#183;&nbsp;&nbsp;"
					Next
				Else
					If Len(oRequest("RowsType").Item) > 0 Then
						Response.Write "<SELECT CLASS=""Lists"" onChange=""window.location.href='" & GetASPFileName("") & "?" & ReplaceValueInURLString(oRequest, "StartPage", "") & "' + this.value"" id=select1 name=select1>"
					Else
						Response.Write "<SELECT CLASS=""Lists"" onChange=""window.location.href='" & GetASPFileName("") & "?" & ReplaceValueInURLString(oRequest, "StartPage", "") & "' + this.value + '&RowsType=" & iSectionType & "'"" id=select1 name=select1>"
					End If
						For iIndex = 1 To iPages
							Response.Write "<OPTION VALUE=""" & ((iIndex - 1) * iRows + 1) & """"
								If iIndex = (Int(iStart / iRows) + 1) Then Response.Write " SELECTED=""1"""
							Response.Write ">" & iIndex & "</OPTION>"
						Next
					Response.Write "</SELECT>"
				End If
			Response.Write "</FONT></SPAN><BR /><BR />"
		If iPages > 20 Then Response.Write "</FORM>"

		iIndex = 1
		If Not IsNull(oRecordset) Then
			If Not oRecordset.EOF Then
				For iIndex = iIndex To iStart - 1
					oRecordset.MoveNext
					If oRecordset.EOF Then Exit For
					If Err.number <> 0 Then Exit For
				Next
			End If
		End If
	End If

	DisplayIncrementalFetchForBank = Err.number
	Err.Clear
End Function

Function DisplayLikeCombo(sComboName)
'************************************************************
'Purpose: To display combo with the LIKE operator options
'Inputs:  sComboName
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayLikeCombo"

	Response.Write "<SELECT NAME=""" & sComboName & """ ID=""" & sComboName & "Cmb"" SIZE=""1"" CLASS=""Lists"">"
		Response.Write "<OPTION VALUE=""" & N_CONTENTS_LIKE & """>Que contenga</OPTION>"
		Response.Write "<OPTION VALUE=""" & N_DOES_NOT_CONTENT_LIKE & """>Que no contenga</OPTION>"
		Response.Write "<OPTION VALUE=""" & N_STARTS_LIKE & """>Que empiece con</OPTION>"
		Response.Write "<OPTION VALUE=""" & N_DOES_NOT_START_LIKE & """>Que no empiece con</OPTION>"
		Response.Write "<OPTION VALUE=""" & N_ENDS_LIKE & """>Que termine con</OPTION>"
		Response.Write "<OPTION VALUE=""" & N_DOES_NOT_END_LIKE & """>Que no termine con</OPTION>"
		Response.Write "<OPTION VALUE=""" & N_EQUAL_LIKE & """>Que sea igual a</OPTION>"
		Response.Write "<OPTION VALUE=""" & N_DIFFERENT_LIKE & """>Que sea diferente a</OPTION>"
	Response.Write "</SELECT>"

	DisplayLikeCombo = Err.number
	Err.Clear
End Function

Function DisplayLikeText(iOption)
'************************************************************
'Purpose: To display the LIKE operator
'Inputs:  iOption
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayLikeText"

	Select Case iOption
		Case N_DOES_NOT_CONTENT_LIKE
			DisplayLikeText = "Que no contenga:"
		Case N_STARTS_LIKE
			DisplayLikeText = "Que empiece con:"
		Case N_DOES_NOT_START_LIKE
			DisplayLikeText = "Que no empiece con:"
		Case N_ENDS_LIKE
			DisplayLikeText = "Que termine con:"
		Case N_DOES_NOT_END_LIKE
			DisplayLikeText = "Que no termine con:"
		Case N_EQUAL_LIKE
			DisplayLikeText = "Que sea igual a:"
		Case N_DIFFERENT_LIKE
			DisplayLikeText = "Que sea diferente a:"
		Case Else
			DisplayLikeText = "Que contenga:"
	End Select
	DisplayLikeText = DisplayLikeText & " "

	Err.Clear
End Function

Function DisplayWarningDiv(sDivName, sMessage)
'************************************************************
'Purpose: To display a warning inside a DIV tag
'Inputs:  sDivName, sMessage
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayWarningDiv"

	Response.Write "<DIV ID=""" & sDivName & """ STYLE=""display: none"">"
		Response.Write "<TABLE WIDTH=""320"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"" ALIGN=""CENTER""><TR><TD>"
			Response.Write "<TABLE BGCOLOR=""#" & S_WIDGET_FRAME_FOR_GUI & """ WIDTH=""320"" ALIGN=""CENTER"" BORDER=""0"" CELLPADDING=""1"" CELLSPACING=""0""><TR><TD>"
				Response.Write "<TABLE BGCOLOR=""#FFFFFF"" WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR><TD>"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""><BR />&nbsp;" & sMessage & "<BR /><BR /></FONT>"
					Response.Write "<CENTER><INPUT TYPE=""SUBMIT"" NAME=""Remove"" ID=""RemoveBtn"" VALUE=""     Sí     "" CLASS=""RedButtons"" />"
					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
					Response.Write "<INPUT TYPE=""BUTTON"" VALUE=""     No     "" CLASS=""Buttons"" onClick=""HideDisplay(document.all['" & sDivName & "'])"" /></CENTER><BR />"
				Response.Write "</TD></TR></TABLE>"
			Response.Write "</TD></TR></TABLE>"
		Response.Write "</TD></TR></TABLE>"
	Response.Write "</DIV>"

	DisplayWarningDiv = Err.number
	Err.Clear
End Function

Function DisplayYesNo(iValue, bDontReverse)
'************************************************************
'Purpose: To display Sí if the value is 1 and No if the value
'         is 0. The value can be reversed.
'Inputs:  iValue, bDontReverse
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayYesNo"
	Dim sValue

	If bDontReverse Then
		sValue = "Sí"
		If iValue = 0 Then sValue = "No"
	Else
		sValue = "No"
		If iValue = 0 Then sValue = "Sí"
	End If

	DisplayYesNo = sValue
	Err.Clear
End Function

Function GetColorForTreshold(dPercentage, sTresholdColor)
'************************************************************
'Purpose: To get a color from a range
'Inputs:  dPercentage
'Outputs: sTresholdColor
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetColorForTreshold"
	Dim sColor
	Dim asColor1
	Dim asColor2

	If CInt(dPercentage * 100) = CInt(GetAdminOption(aAdminOptionsComponent, RED_TRESHOLD_OPTION)) Then
		sColor = GetAdminOption(aAdminOptionsComponent, RED_COLOR_OPTION)
	ElseIf CInt(dPercentage * 100) = CInt(GetAdminOption(aAdminOptionsComponent, YELLOW_TRESHOLD_OPTION)) Then
		sColor = GetAdminOption(aAdminOptionsComponent, YELLOW_COLOR_OPTION)
	ElseIf CInt(dPercentage * 100) = CInt(GetAdminOption(aAdminOptionsComponent, GREEN_TRESHOLD_OPTION)) Then
		sColor = GetAdminOption(aAdminOptionsComponent, GREEN_COLOR_OPTION)
	ElseIf CInt(dPercentage * 100) < CInt(GetAdminOption(aAdminOptionsComponent, YELLOW_TRESHOLD_OPTION)) Then
		sColor = GetAdminOption(aAdminOptionsComponent, RED_COLOR_OPTION)
		asColor1 = Array(HexToLng(Left(sColor, 2)), HexToLng(Mid(sColor, 3, 2)), HexToLng(Right(sColor, 2)))
		sColor = GetAdminOption(aAdminOptionsComponent, YELLOW_COLOR_OPTION)
		asColor2 = Array(HexToLng(Left(sColor, 2)), HexToLng(Mid(sColor, 3, 2)), HexToLng(Right(sColor, 2)))
		sColor = Right(("00" & Hex(CInt(((CInt(asColor2(0)) - CInt(asColor1(0))) * dPercentage) + CInt(asColor1(0))))), Len("00")) & Right(("00" & Hex(CInt(((CInt(asColor2(1)) - CInt(asColor1(1))) * dPercentage) + CInt(asColor1(1))))), Len("00")) & Right(("00" & Hex(CInt(((CInt(asColor2(2)) - CInt(asColor1(2))) * dPercentage) + CInt(asColor1(2))))), Len("00"))
	Else
		sColor = GetAdminOption(aAdminOptionsComponent, YELLOW_COLOR_OPTION)
		asColor1 = Array(HexToLng(Left(sColor, 2)), HexToLng(Mid(sColor, 3, 2)), HexToLng(Right(sColor, 2)))
		sColor = GetAdminOption(aAdminOptionsComponent, GREEN_COLOR_OPTION)
		asColor2 = Array(HexToLng(Left(sColor, 2)), HexToLng(Mid(sColor, 3, 2)), HexToLng(Right(sColor, 2)))
		sColor = Right(("00" & Hex(CInt(((CInt(asColor2(0)) - CInt(asColor1(0))) * (dPercentage - (CInt(GetAdminOption(aAdminOptionsComponent, YELLOW_TRESHOLD_OPTION)) / 100))) + CInt(asColor1(0))))), Len("00")) & Right(("00" & Hex(CInt(((CInt(asColor2(1)) - CInt(asColor1(1))) * (dPercentage - (CInt(GetAdminOption(aAdminOptionsComponent, YELLOW_TRESHOLD_OPTION)) / 100))) + CInt(asColor1(1))))), Len("00")) & Right(("00" & Hex(CInt(((CInt(asColor2(2)) - CInt(asColor1(2))) * (dPercentage - (CInt(GetAdminOption(aAdminOptionsComponent, YELLOW_TRESHOLD_OPTION)) / 100))) + CInt(asColor1(2))))), Len("00"))
	End If

	sTresholdColor = sColor
	GetColorForTreshold = Err.number
	Err.Clear
End Function

Function GetPrivilegesNames(lPermissions, sTab, sNewLine)
'************************************************************
'Purpose: To display the name of the privileges
'Inputs:  lPermissions, sTab, sNewLine
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPrivilegesNames"
	Dim sPrivilegesNames

	sPrivilegesNames = ""
	If lPermissions And N_ADD_PERMISSIONS Then sPrivilegesNames = sPrivilegesNames & sTab & "Agregar registros" & sNewLine
	If lPermissions And N_MODIFY_PERMISSIONS Then sPrivilegesNames = sPrivilegesNames & sTab & "Modificar registros" & sNewLine
	If lPermissions And N_REMOVE_PERMISSIONS Then sPrivilegesNames = sPrivilegesNames & sTab & "Eliminar registros" & sNewLine
	If Not B_ISSSTE Then
		If lPermissions And N_BUDGET_PERMISSIONS Then sPrivilegesNames = sPrivilegesNames & sTab & "Administración de presupuestos" & sNewLine
		If lPermissions And N_AREAS_PERMISSIONS Then sPrivilegesNames = sPrivilegesNames & sTab & "Administración de áreas" & sNewLine
		If lPermissions And N_POSITIONS_PERMISSIONS Then sPrivilegesNames = sPrivilegesNames & sTab & "Administración de puestos" & sNewLine
		If lPermissions And N_JOBS_PERMISSIONS Then sPrivilegesNames = sPrivilegesNames & sTab & "Administración de plazas" & sNewLine
		If lPermissions And N_EMPLOYEES_PERMISSIONS Then sPrivilegesNames = sPrivilegesNames & sTab & "Administración de empleados" & sNewLine
		If lPermissions And N_EMPLOYEE_PAYROLL_PERMISSIONS Then sPrivilegesNames = sPrivilegesNames & sTab & "Administración de pagos a empleados" & sNewLine
		If lPermissions And N_SADE_PERMISSIONS Then sPrivilegesNames = sPrivilegesNames & sTab & "Administración de desarrollo humano" & sNewLine
		If lPermissions And N_PAYROLL_PERMISSIONS Then sPrivilegesNames = sPrivilegesNames & sTab & "Administración de la nómina" & sNewLine
		If lPermissions And N_PAYMENTS_PERMISSIONS Then sPrivilegesNames = sPrivilegesNames & sTab & "Administración de cheques" & sNewLine
		If lPermissions And N_REPORTS_PERMISSIONS Then sPrivilegesNames = sPrivilegesNames & sTab & "Ver reportes" & sNewLine
		If lPermissions And N_TOOLS_PERMISSIONS Then sPrivilegesNames = sPrivilegesNames & sTab & "Usar herramientas" & sNewLine
		If lPermissions And N_CATALOGS_PERMISSIONS Then sPrivilegesNames = sPrivilegesNames & sTab & "Administración de catálogos" & sNewLine
		If lPermissions And N_TACO_PERMISSIONS Then sPrivilegesNames = sPrivilegesNames & sTab & "Tablero de control" & sNewLine
	End If
	If lPermissions And N_DELETE_FILES_PERMISSIONS Then sPrivilegesNames = sPrivilegesNames & sTab & "Borrar archivos" & sNewLine
	If Len(sPrivilegesNames) = 0 Then sPrivilegesNames = sTab & "Ninguno" & sNewLine

	GetPrivilegesNames = Left(sPrivilegesNames, (Len(sPrivilegesNames) - Len(sNewLine)))
	Err.Clear
End Function

Function GetPercentageBar(dPercentage, sTresholdColor, sOnClick, sOnMouseOver, sOnMouseOut)
'************************************************************
'Purpose: To display an HTML table as a horizontal bar
'Inputs:  dPercentage, sTresholdColor, sOnClick, sOnMouseOver, sOnMouseOut
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPercentageBar"
	Dim sHTML

	sHTML = ""
	If Not bIsNetscape Then sHTML = sHTML & "<DIV onClick=""" & sOnClick & """ onMouseOver=""" & sOnMouseOver & """ onMouseOut=""" & sOnMouseOut & """>"
		sHTML = sHTML & "<TABLE WIDTH=""103"" BORDER=""0"" BGCOLOR=""#000000"" CELLPADDING=""0"" CELLSPACING=""1""><TR><TD>"
			sHTML = sHTML & "<TABLE WIDTH=""101"" BORDER=""0"" BGCOLOR=""#FFFFFF"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
				sHTML = sHTML & "<TD BGCOLOR=""#" & sTresholdColor & """ WIDTH=""" & ((dPercentage * 100) + 1) & """>"
					sHTML = sHTML & "<IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""10"">"
				sHTML = sHTML & "</TD>"
				If dPercentage < 1 Then
					sHTML = sHTML & "<TD WIDTH=""" & (99 - (dPercentage * 100)) & """>"
						sHTML = sHTML & "<IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""10"">"
					sHTML = sHTML & "</TD>"
				End If
			sHTML = sHTML & "</TR></TABLE>"
		sHTML = sHTML & "</TD></TR></TABLE>"
	If Not bIsNetscape Then sHTML = sHTML & "</DIV>"
	GetPercentageBar = sHTML
	Err.Clear
End Function
%>