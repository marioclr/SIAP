<%
Const N_ID_REPORT = 0
Const N_USER_ID_REPORT = 1
Const N_CONSTANT_ID_REPORT = 2
Const S_NAME_REPORT = 3
Const S_DESCRIPTION_REPORT = 4
Const S_PARAMETERS_REPORT = 5
Const S_URL_REPORT = 6
Const S_QUERY_CONDITION_REPORT = 7
Const B_CHECK_FOR_DUPLICATED_REPORT = 8
Const B_IS_DUPLICATED_REPORT = 9
Const B_COMPONENT_INITIALIZED_REPORT = 10

Const N_REPORT_COMPONENT_SIZE = 10

Dim aReportComponent()
Redim aReportComponent(N_REPORT_COMPONENT_SIZE)

Function InitializeReportComponent(oRequest, aReportComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Report Component
'         using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aReportComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeReportComponent"
	Redim Preserve aReportComponent(N_REPORT_COMPONENT_SIZE)

	If IsEmpty(aReportComponent(N_ID_REPORT)) Then
		If Len(oRequest("ReportID").Item) > 0 Then
			aReportComponent(N_ID_REPORT) = CLng(oRequest("ReportID").Item)
			If aReportComponent(N_ID_REPORT) >= 110000 Then aReportComponent(N_ID_REPORT) = -1
		Else
			aReportComponent(N_ID_REPORT) = -1
		End If
	End If

	If IsEmpty(aReportComponent(N_USER_ID_REPORT)) Then
		If Len(oRequest("UserID").Item) > 0 Then
			aReportComponent(N_USER_ID_REPORT) = CLng(oRequest("UserID").Item)
		Else
			aReportComponent(N_USER_ID_REPORT) = aLoginComponent(N_USER_ID_LOGIN)
		End If
	End If

	If IsEmpty(aReportComponent(S_NAME_REPORT)) Then
		If Len(oRequest("ReportName").Item) > 0 Then
			aReportComponent(S_NAME_REPORT) = oRequest("ReportName").Item
		Else
			aReportComponent(S_NAME_REPORT) = ""
		End If
	End If
	aReportComponent(S_NAME_REPORT) = Left(aReportComponent(S_NAME_REPORT), 100)

	If IsEmpty(aReportComponent(S_DESCRIPTION_REPORT)) Then
		If Len(oRequest("ReportDescription").Item) > 0 Then
			aReportComponent(S_DESCRIPTION_REPORT) = oRequest("ReportDescription").Item
		Else
			aReportComponent(S_DESCRIPTION_REPORT) = ""
		End If
	End If
	aReportComponent(S_DESCRIPTION_REPORT) = Left(aReportComponent(S_DESCRIPTION_REPORT), 255)

	If IsEmpty(aReportComponent(S_PARAMETERS_REPORT)) Then
		If Len(oRequest("ReportParameters").Item) > 0 Then
			aReportComponent(S_PARAMETERS_REPORT) = oRequest("ReportParameters").Item
		Else
			aReportComponent(S_PARAMETERS_REPORT) = RemoveParameterFromURLString(oRequest, "New")
		End If
	End If
	aReportComponent(S_PARAMETERS_REPORT) = Left(aReportComponent(S_PARAMETERS_REPORT), 12000)

	aReportComponent(N_CONSTANT_ID_REPORT) = CLng(GetParameterFromURLString(aReportComponent(S_PARAMETERS_REPORT), "ReportID"))

	aReportComponent(B_CHECK_FOR_DUPLICATED_REPORT) = True
	aReportComponent(B_IS_DUPLICATED_REPORT) = False

	aReportComponent(B_COMPONENT_INITIALIZED_REPORT) = True
	InitializeReportComponent = Err.number
	Err.Clear
End Function

Function AddReport(oRequest, oADODBConnection, aReportComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new report into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aReportComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddReport"
	Dim sParameters
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aReportComponent(B_COMPONENT_INITIALIZED_REPORT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeReportComponent(oRequest, aReportComponent)
	End If

	If (aReportComponent(N_ID_REPORT) = -1) Or (Len(oRequest("Add").Item) > 0) Then
		sErrorDescription = "No se pudo obtener un identificador para el nuevo reporte."
		lErrorNumber = GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, aReportComponent(N_ID_REPORT), sErrorDescription)
	End If

	If lErrorNumber = 0 Then
		If lErrorNumber = 0 Then
			If aReportComponent(B_IS_DUPLICATED_REPORT) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un reporte con el nombre " & aReportComponent(S_NAME_REPORT) & "."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ReportComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				If Not CheckReportInformationConsistency(aReportComponent, sErrorDescription) Then
					lErrorNumber = -1
				Else
					sParameters = Replace(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(aReportComponent(S_PARAMETERS_REPORT), "PolicyName"), "Saved"), "SavedReportID"), "'", "")
					sErrorDescription = "No se pudo guardar la información del nuevo registro."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & aReportComponent(N_ID_REPORT) & ", " & aReportComponent(N_USER_ID_REPORT) & ", " & aReportComponent(N_CONSTANT_ID_REPORT) & ", '" & Replace(aReportComponent(S_NAME_REPORT), "'", "") & "', '" & Replace(aReportComponent(S_DESCRIPTION_REPORT), "'", "´") & "', '" & Left(sParameters, 4000) & "', '" & Mid(sParameters, 4001, 4000) & "', '" & Mid(sParameters, 8001, 4000) & "')", "ReportComponent.asp", S_FUNCTION_NAME, 1310, sErrorDescription, Null)
				End If
			End If
		End If
	End If

	AddReport = lErrorNumber
	Err.Clear
End Function

Function GetReport(oRequest, oADODBConnection, aReportComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a report from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aReportComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetReport"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aReportComponent(B_COMPONENT_INITIALIZED_REPORT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeReportComponent(oRequest, aReportComponent)
	End If

	If aReportComponent(N_ID_REPORT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del reporte para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ReportComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del reporte."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Reports Where ReportID=" & aReportComponent(N_ID_REPORT), "ReportComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El reporte especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ReportComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aReportComponent(N_USER_ID_REPORT) = CLng(oRecordset.Fields("UserID").Value)
				aReportComponent(N_CONSTANT_ID_REPORT) = CLng(oRecordset.Fields("ConstantID").Value)
				aReportComponent(S_NAME_REPORT) = CStr(oRecordset.Fields("ReportName").Value)
				aReportComponent(S_DESCRIPTION_REPORT) = CStr(oRecordset.Fields("ReportDescription").Value)
				aReportComponent(S_PARAMETERS_REPORT) = CStr(oRecordset.Fields("ReportParameters1").Value)
				aReportComponent(S_PARAMETERS_REPORT) = aReportComponent(S_PARAMETERS_REPORT) & CStr(oRecordset.Fields("ReportParameters2").Value)
				aReportComponent(S_PARAMETERS_REPORT) = aReportComponent(S_PARAMETERS_REPORT) & CStr(oRecordset.Fields("ReportParameters3").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetReport = lErrorNumber
	Err.Clear
End Function

Function GetReports(oRequest, oADODBConnection, aReportComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the reports from
'		  the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aReportComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetReports"
	Dim sCondition
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aReportComponent(B_COMPONENT_INITIALIZED_REPORT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeReportComponent(oRequest, aReportComponent)
	End If

	sCondition = Trim(aReportComponent(S_QUERY_CONDITION_REPORT))
	If Len(sCondition) > 0 Then
		If InStr(1, sCondition, "And", vbBinaryCompare) <> 1 Then
			sCondition = " And " & sCondition
		End If
	End If
	sErrorDescription = "No se pudo obtener la información de los reportes."
	If (Len(oRequest("AllReports").Item) > 0) And (Len(oRequest("ReportToShow").Item) > 0) Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Reports.*, UserName, UserLastName From Reports, Users Where (Reports.UserID=Users.UserID) And (Reports.ConstantID=" & oRequest("ReportToShow").Item & ") And (ReportDescription<>'" & CATALOG_SEPARATOR & "') " & sCondition & " Order By ConstantID, ReportName, ReportID", "ReportComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Reports.*, UserName, UserLastName From Reports, Users Where (Reports.UserID=Users.UserID) And (Reports.UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ") And (ReportDescription<>'" & CATALOG_SEPARATOR & "') " & sCondition & " Order By ConstantID, ReportName, ReportID", "ReportComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	End If

	GetReports = lErrorNumber
	Err.Clear
End Function

Function ModifyReport(oRequest, oADODBConnection, aReportComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing report in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aReportComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyReport"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aReportComponent(B_COMPONENT_INITIALIZED_REPORT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeReportComponent(oRequest, aReportComponent)
	End If

	If aReportComponent(N_ID_REPORT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del reporte a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ReportComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If Not CheckReportInformationConsistency(aReportComponent, sErrorDescription) Then
			lErrorNumber = -1
		Else
			sErrorDescription = "No se pudo modificar la información del reporte."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Reports Set ReportName='" & Replace(aReportComponent(S_NAME_REPORT), "'", "") & "', ReportDescription='" & Replace(aReportComponent(S_DESCRIPTION_REPORT), "'", "") & "' Where (ReportID=" & aReportComponent(N_ID_REPORT) & ")", "ReportComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	ModifyReport = lErrorNumber
	Err.Clear
End Function

Function RemoveReport(oRequest, oADODBConnection, aReportComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a report from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aReportComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveReport"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aReportComponent(B_COMPONENT_INITIALIZED_REPORT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeReportComponent(oRequest, aReportComponent)
	End If

	If aReportComponent(N_ID_REPORT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el reporte a eliminar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ReportComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del reporte."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Reports Where (ReportID=" & aReportComponent(N_ID_REPORT) & ")", "ReportComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	RemoveReport = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfReport(aReportComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific report exists in the database
'Inputs:  aReportComponent
'Outputs: aReportComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfReport"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aReportComponent(B_COMPONENT_INITIALIZED_REPORT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeReportComponent(oRequest, aReportComponent)
	End If

	If Len(aReportComponent(S_NAME_REPORT)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el nombre del reporte para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ReportComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del reporte en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Reports Where (ReportName='" & Replace(aReportComponent(S_NAME_REPORT), "'", "") & "')", "ReportComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aReportComponent(B_IS_DUPLICATED_REPORT) = True
				aReportComponent(N_ID_REPORT) = CLng(oRecordset.Fields("ReportID").Value)
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfReport = lErrorNumber
	Err.Clear
End Function

Function CheckReportInformationConsistency(aReportComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the database
'Inputs:  aReportComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckReportInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aReportComponent(N_ID_REPORT)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del reporte no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aReportComponent(N_USER_ID_REPORT)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del usuario no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aReportComponent(N_CONSTANT_ID_REPORT)) Then aReportComponent(N_CONSTANT_ID_REPORT) = -1
	If Len(aReportComponent(S_NAME_REPORT)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nombre del reporte está vacío."
		bIsCorrect = False
	End If
	If Len(aReportComponent(S_DESCRIPTION_REPORT)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- La descripción del reporte está vacío."
		bIsCorrect = False
	End If
	If Len(aReportComponent(S_PARAMETERS_REPORT)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- Los parámetros del reporte están vacíos."
		bIsCorrect = False
	End If

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información del reporte contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ReportComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckReportInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function DisplayReportForm(oRequest, oADODBConnection, sAction, aReportComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a report from the
'		  database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, aReportComponent
'Outputs: aReportComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayReportForm"
	Dim lErrorNumber

	If aReportComponent(N_ID_REPORT) <> -1 Then
		lErrorNumber = GetReport(oRequest, oADODBConnection, aReportComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckReportFields(oForm) {" & vbNewLine
				If Len(oRequest("Delete").Item) = 0 Then
					Response.Write "if (oForm) {" & vbNewLine
						Response.Write "if (oForm.ReportName.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de introducir el nombre del reporte.');" & vbNewLine
							Response.Write "oForm.ReportName.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (oForm.ReportDescription.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de introducir la descripción del reporte.');" & vbNewLine
							Response.Write "oForm.ReportDescription.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (oForm.ReportParameters.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de introducir los parámetros del reporte.');" & vbNewLine
							Response.Write "oForm.ReportParameters.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "}" & vbNewLine
				End If
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckReportFields" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "<FORM NAME=""ReportFrm"" ID=""ReportFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckReportFields(this)"">"
			If Len(aReportComponent(S_URL_REPORT)) > 0 Then Call DisplayURLParametersAsHiddenValues(RemoveParameterFromURLString(aReportComponent(S_URL_REPORT), "ReportID"))
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""Reports"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReportID"" ID=""ReportIDHdn"" VALUE=""" & aReportComponent(N_ID_REPORT) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConstantID"" ID=""ConstantIDHdn"" VALUE=""" & aReportComponent(N_CONSTANT_ID_REPORT) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AllReports"" ID=""AllReportsHdn"" VALUE=""" & oRequest("AllReports").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReportToShow"" ID=""ReportToShowHdn"" VALUE=""" & oRequest("ReportToShow").Item & """ />"

			Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Nombre del reporte: </FONT>"
			Response.Write "<INPUT TYPE=""TEXT"" NAME=""ReportName"" ID=""ReportNameTxt"" SIZE=""37"" MAXLENGTH=""100"" VALUE=""" & aReportComponent(S_NAME_REPORT) & """ CLASS=""TextFields"" /><BR />"
			Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Descripción del reporte:</FONT><BR />"
			Response.Write "<TEXTAREA NAME=""ReportDescription"" ID=""ReportDescriptionTxtArea"" ROWS=""5"" COLS=""50"" MAXLENGTH=""255"" CLASS=""TextFields"">" & aReportComponent(S_DESCRIPTION_REPORT) & "</TEXTAREA><BR />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReportParameters"" ID=""ReportParametersHdn"" VALUE=""" & aReportComponent(S_PARAMETERS_REPORT) & """ /><BR />"
			Response.Write "<BR />"
			If Len(oRequest("New").Item) > 0 Then
				Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar como reporte nuevo"" CLASS=""Buttons"" />"
				If Len(oRequest("SavedReportID").Item) > 0 Then
					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
					Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar el reporte original"" CLASS=""Buttons"" onClick=""document.ReportFrm.ReportID.value='" & oRequest("SavedReportID").Item & "'"" />"
				End If
			ElseIf Len(oRequest("Change").Item) > 0 Then
				Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			ElseIf Len(oRequest("Delete").Item) > 0 Then
				Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveReportWngDiv']); ReportFrm.Remove.focus()"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.history.back()"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveReportWngDiv", "¿Está seguro que desea borrar el reporte de la &nbsp;base de datos?")
		Response.Write "</FORM>"
	End If

	DisplayReportForm = lErrorNumber
	Err.Clear
End Function

Function DisplayReportAsHiddenFields(oRequest, oADODBConnection, aReportComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a report using
'		  hidden form fields
'Inputs:  oRequest, oADODBConnection, aReportComponent
'Outputs: aReportComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayReportAsHiddenFields"

	If Len(aReportComponent(S_URL_REPORT)) > 0 Then Call DisplayURLParametersAsHiddenValues(aReportComponent(S_URL_REPORT))
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReportID"" ID=""ReportIDHdn"" VALUE=""" & aReportComponent(N_ID_REPORT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserID"" ID=""UserIDHdn"" VALUE=""" & aReportComponent(N_USER_ID_REPORT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConstantID"" ID=""ConstantIDHdn"" VALUE=""" & aReportComponent(N_CONSTANT_ID_REPORT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReportName"" ID=""ReportNameHdn"" VALUE=""" & aReportComponent(S_NAME_REPORT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReportDescription"" ID=""ReportDescriptionHdn"" VALUE=""" & aReportComponent(S_DESCRIPTION_REPORT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReportParameters"" ID=""ReportParametersHdn"" VALUE=""" & aReportComponent(S_PARAMETERS_REPORT) & """ />"

	DisplayReportAsHiddenFields = Err.number
	Err.Clear
End Function

Function DisplayReportsInThreeSmallColumns(oRequest, oADODBConnection, aReportComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a report using
'		  hidden form fields
'Inputs:  oRequest, oADODBConnection, aReportComponent
'Outputs: aReportComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayReportsInThreeSmallColumns"
	Dim oRecordset
	Dim aReportMenu()
	Dim sReportMenuData
	Dim lCurrentReport
	Dim iIndex
	Dim lErrorNumber

	lErrorNumber = GetReports(oRequest, oADODBConnection, aReportComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			lCurrentReport = -1
			iIndex = 0
			Do While Not oRecordset.EOF
				If lCurrentReport <> CLng(oRecordset.Fields("ConstantID").Value) Then
					iIndex = iIndex + 1
					lCurrentReport = CLng(oRecordset.Fields("ConstantID").Value)
				End If
				iIndex = iIndex + 1
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			ReDim aReportMenu(iIndex)
			lCurrentReport = -1
			iIndex = 0
			oRecordset.MoveFirst
			Do While Not oRecordset.EOF
				If lCurrentReport <> CLng(oRecordset.Fields("ConstantID").Value) Then
					sReportMenuData = "<TITLE />" & UCase(GetReportNameByConstant(CLng(oRecordset.Fields("ConstantID").Value)))
					aReportMenu(iIndex) = Split(sReportMenuData, LIST_SEPARATOR, -1, vbBinaryCompare)
					iIndex = iIndex + 1
					lCurrentReport = CLng(oRecordset.Fields("ConstantID").Value)
				End If
				sReportMenuData = CleanStringForHTML(CStr(oRecordset.Fields("ReportName").Value)) & LIST_SEPARATOR & _
								  "<B>Dueño: " & CleanStringForHTML(CStr(oRecordset.Fields("UserName").Value) & " " & CStr(oRecordset.Fields("UserLastName").Value)) & "</B><BR />" & _
								  CleanStringForHTML(CStr(oRecordset.Fields("ReportDescription").Value))
				If aLoginComponent(N_USER_ID_LOGIN) = CLng(oRecordset.Fields("UserID").Value) Then sReportMenuData = sReportMenuData & "<BR /><A HREF=""SavedReport.asp?Change=1&ReportID=" & CStr(oRecordset.Fields("ReportID").Value) & "&AllReports=" & oRequest("AllReports").Item & "&ReportToShow=" & oRequest("ReportToShow").Item & """><IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar Descripción"" BORDER=""0"" HSPACE=""3"" />Modificar</A><IMG SRC=""Images/Transparent.gif"" WIDTH=""40"" HEIGHT=""1"" /><A HREF=""SavedReport.asp?Delete=1&ReportID=" & CStr(oRecordset.Fields("ReportID").Value) & "&AllReports=" & oRequest("AllReports").Item & "&ReportToShow=" & oRequest("ReportToShow").Item & """><IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar"" BORDER=""0"" HSPACE=""3"" />Eliminar</A>"
				sReportMenuData = sReportMenuData & LIST_SEPARATOR & "Images/MnLeftArrows.gif" & LIST_SEPARATOR & _
								  "Reports.asp?Saved=1&SavedReportID=" & CStr(oRecordset.Fields("ReportID").Value) & "&" & CStr(oRecordset.Fields("ReportParameters1").Value) & "&AllReports=" & oRequest("AllReports").Item & "&ReportToShow=" & oRequest("ReportToShow").Item
				sReportMenuData = sReportMenuData & CStr(oRecordset.Fields("ReportParameters2").Value)
				sReportMenuData = sReportMenuData & CStr(oRecordset.Fields("ReportParameters3").Value)
				sReportMenuData = sReportMenuData & LIST_SEPARATOR
				Err.Clear
				If aLoginComponent(N_USER_ID_LOGIN) = CLng(oRecordset.Fields("UserID").Value) Then
					sReportMenuData = sReportMenuData & "-1"
				Else
					sReportMenuData = sReportMenuData & N_REPORTS_PERMISSIONS
				End If
				aReportMenu(iIndex) = Split(sReportMenuData, LIST_SEPARATOR, -1, vbBinaryCompare)
				iIndex = iIndex + 1
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
			aMenuComponent(A_ELEMENTS_MENU) = aReportMenu
			aMenuComponent(B_USE_DIV_MENU) = True
			Response.Write "<TABLE WIDTH=""900"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Call DisplayMenuInThreeSmallColumns(aMenuComponent)
			Response.Write "</TABLE>"
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen reportes registrados en la base de datos."
		End If
	End If

	Set oRecordset = Nothing
	DisplayReportsInThreeSmallColumns = lErrorNumber
	Err.Clear
End Function

%>