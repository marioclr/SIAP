<%
Const N_PROJECT_ID_HISTORY = 0
Const N_TASK_ID_HISTORY = 1
Const N_RECORD_ID_HISTORY = 2
Const N_DATE_HISTORY = 3
Const N_HOUR_HISTORY = 4
Const N_MINUTE_HISTORY = 5
Const N_USER_ID_HISTORY = 6
Const S_DESCRIPTION_HISTORY = 7
Const B_COMPONENT_INITIALIZED_HISTORY = 8

Const N_HISTORY_COMPONENT_SIZE = 8

Dim aHistoryComponent()
Redim aHistoryComponent(N_HISTORY_COMPONENT_SIZE)

Function InitializeHistoryComponent(oRequest, aHistoryComponent)
'************************************************************
'Purpose: To initialize the empty elements of the History Component
'         using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aHistoryComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeHistoryComponent"
	Redim Preserve aHistoryComponent(N_HISTORY_COMPONENT_SIZE)
	Dim oTempDate

	If IsEmpty(aHistoryComponent(N_PROJECT_ID_HISTORY)) Then
		If Len(oRequest("ProjectID").Item) > 0 Then
			aHistoryComponent(N_PROJECT_ID_HISTORY) = CLng(oRequest("ProjectID").Item)
		Else
			aHistoryComponent(N_PROJECT_ID_HISTORY) = -1
		End If
	End If

	If IsEmpty(aHistoryComponent(N_TASK_ID_HISTORY)) Then
		If Len(oRequest("TaskID").Item) > 0 Then
			aHistoryComponent(N_TASK_ID_HISTORY) = CLng(oRequest("TaskID").Item)
		Else
			aHistoryComponent(N_TASK_ID_HISTORY) = -1
		End If
	End If

	If IsEmpty(aHistoryComponent(N_RECORD_ID_HISTORY)) Then
		If Len(oRequest("RecordID").Item) > 0 Then
			aHistoryComponent(N_RECORD_ID_HISTORY) = CLng(oRequest("RecordID").Item)
		Else
			aHistoryComponent(N_RECORD_ID_HISTORY) = -1
		End If
	End If

	If IsEmpty(aHistoryComponent(N_DATE_HISTORY)) Then
		If Len(oRequest("HistoryYear").Item) > 0 Then
			aHistoryComponent(N_DATE_HISTORY) = CLng(oRequest("HistoryYear").Item & Right(("0" & oRequest("HistoryMonth").Item), Len("00")) & Right(("0" & oRequest("HistoryDay").Item)), Len("00"))
		ElseIf Len(oRequest("HistoryDate").Item) > 0 Then
			aHistoryComponent(N_DATE_HISTORY) = CLng(oRequest("HistoryDate").Item)
		Else
			aHistoryComponent(N_DATE_HISTORY) = Left(GetSerialNumberForDate(""), Len("00000000"))
		End If
	End If

	If IsEmpty(aHistoryComponent(N_HOUR_HISTORY)) Then
		If Len(oRequest("HistoryHour").Item) > 0 Then
			aHistoryComponent(N_HOUR_HISTORY) = CLng(oRequest("HistoryHour").Item)
		Else
			aHistoryComponent(N_HOUR_HISTORY) = Hour(Time())
		End If
	End If

	If IsEmpty(aHistoryComponent(N_MINUTE_HISTORY)) Then
		If Len(oRequest("HistoryMinute").Item) > 0 Then
			aHistoryComponent(N_MINUTE_HISTORY) = CLng(oRequest("HistoryMinute").Item)
		Else
			aHistoryComponent(N_MINUTE_HISTORY) = Minute(Time())
		End If
	End If

	If IsEmpty(aHistoryComponent(N_USER_ID_HISTORY)) Then
		If Len(oRequest("UserID").Item) > 0 Then
			aHistoryComponent(N_USER_ID_HISTORY) = CLng(oRequest("UserID").Item)
		Else
			aHistoryComponent(N_USER_ID_HISTORY) = aLoginComponent(N_USER_ID_LOGIN)
		End If
	End If

	If IsEmpty(aHistoryComponent(S_DESCRIPTION_HISTORY)) Then
		If Len(oRequest("HistoryDescription").Item) > 0 Then
			aHistoryComponent(S_DESCRIPTION_HISTORY) = oRequest("HistoryDescription").Item
		Else
			aHistoryComponent(S_DESCRIPTION_HISTORY) = ""
		End If
	End If
	aHistoryComponent(S_DESCRIPTION_HISTORY) = Left(aHistoryComponent(S_DESCRIPTION_HISTORY), 2000)

	aHistoryComponent(B_COMPONENT_INITIALIZED_HISTORY) = True
	InitializeHistoryComponent = Err.number
	Err.Clear
End Function

Function AddEventToHistoryList(oRequest, oADODBConnection, aHistoryComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new event on the history list database
'Inputs:  oRequest, oADODBConnection
'Outputs: aHistoryComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddEventToHistoryList"
	Dim iIsOpposition
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aHistoryComponent(B_COMPONENT_INITIALIZED_HISTORY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeHistoryComponent(oRequest, aHistoryComponent)
	End If

	If (aHistoryComponent(N_PROJECT_ID_HISTORY) = -1) Or (aHistoryComponent(N_TASK_ID_HISTORY) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador de la actividad para agregar el nuevo comentario."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "TaCoHistoryListComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If aHistoryComponent(N_RECORD_ID_HISTORY) = -1 Then
			sErrorDescription = "No se pudo obtener un identificador para el nuevo comentario."
			lErrorNumber = GetNewIDFromTable(oADODBConnection, "TasksHistoryList", "RecordID", "(ProjectID=" & aHistoryComponent(N_PROJECT_ID_HISTORY) & ") And (TaskID=" & aHistoryComponent(N_TASK_ID_HISTORY) & ")", 1, aHistoryComponent(N_RECORD_ID_HISTORY), sErrorDescription)
		End If
		If lErrorNumber = 0 Then
			If Not CheckHistoryListEventInformationConsistency(aHistoryComponent, sErrorDescription) Then
				lErrorNumber = -1
			Else
				sErrorDescription = "No se pudo guardar la información del nuevo comentario."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & TACO_PREFIX & "TasksHistoryList (ProjectID, TaskID, RecordID, HistoryDate, HistoryHour, HistoryMinute, UserID, bSystemMessage, HistoryDescription) Values (" & aHistoryComponent(N_PROJECT_ID_HISTORY) & ", " & aHistoryComponent(N_TASK_ID_HISTORY) & ", " & aHistoryComponent(N_RECORD_ID_HISTORY) & ", " & aHistoryComponent(N_DATE_HISTORY) & ", " & aHistoryComponent(N_HOUR_HISTORY) & ", " & aHistoryComponent(N_MINUTE_HISTORY) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", 0, '" & Replace(aHistoryComponent(S_DESCRIPTION_HISTORY), "'", "") & "')", "TaCoHistoryListComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
		End If
	End If

	Set oRecordset = Nothing
	AddEventToHistoryList = lErrorNumber
	Err.Clear
End Function

Function GetHistoryListEvent(oRequest, oADODBConnection, aHistoryComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the tasks for the
'         history list from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aHistoryComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetHistoryListEvent"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aHistoryComponent(B_COMPONENT_INITIALIZED_HISTORY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeHistoryComponent(oRequest, aHistoryComponent)
	End If

	If (aHistoryComponent(N_PROJECT_ID_HISTORY) = -1) Or (aHistoryComponent(N_TASK_ID_HISTORY) = -1) Or (aHistoryComponent(N_RECORD_ID_HISTORY) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del comentario para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "TaCoHistoryListComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del comentario."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From " & TACO_PREFIX & "TasksHistoryList Where (ProjectID=" & aHistoryComponent(N_PROJECT_ID_HISTORY) & ") And (TaskID=" & aHistoryComponent(N_TASK_ID_HISTORY) & ") And (RecordID=" & aHistoryComponent(N_RECORD_ID_HISTORY) & ")", "TaCoHistoryListComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = -1
				sErrorDescription = "El comentario especificado no se encuentra en el sistema."
			Else
				aHistoryComponent(N_DATE_HISTORY) = CLng(oRecordset.Fields("HistoryDate").Value)
				aHistoryComponent(N_HOUR_HISTORY) = CInt(oRecordset.Fields("HistoryHour").Value)
				aHistoryComponent(N_MINUTE_HISTORY) = CInt(oRecordset.Fields("HistoryMinute").Value)
				aHistoryComponent(N_USER_ID_HISTORY) = CLng(oRecordset.Fields("UserID").Value)
				aHistoryComponent(S_DESCRIPTION_HISTORY) = CStr(oRecordset.Fields("HistoryDescription").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetHistoryListEvent = lErrorNumber
	Err.Clear
End Function

Function GetHistoryList(oRequest, oADODBConnection, aHistoryComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the tasks for the
'         history list from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aHistoryComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetHistoryList"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aHistoryComponent(B_COMPONENT_INITIALIZED_HISTORY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeHistoryComponent(oRequest, aHistoryComponent)
	End If

	If (aHistoryComponent(N_PROJECT_ID_HISTORY) = -1) Or (aHistoryComponent(N_TASK_ID_HISTORY) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador de la actividad para obtener sus comentarios."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "TaCoHistoryListComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información de las actividades."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select " & TACO_PREFIX & "TasksHistoryList.*, UserName, UserLastName From " & TACO_PREFIX & "TasksHistoryList, Users Where (" & TACO_PREFIX & "TasksHistoryList.UserID=Users.UserID) And (" & TACO_PREFIX & "TasksHistoryList.ProjectID=" & aHistoryComponent(N_PROJECT_ID_HISTORY) & ") And (" & TACO_PREFIX & "TasksHistoryList.TaskID=" & aHistoryComponent(N_TASK_ID_HISTORY) & ") Order By HistoryDate, HistoryHour, HistoryMinute, RecordID", "TaCoHistoryListComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	End If

	GetHistoryList = lErrorNumber
	Err.Clear
End Function

Function ModifyEventOnHistoryList(oRequest, oADODBConnection, aHistoryComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an event on the history list database
'Inputs:  oRequest, oADODBConnection
'Outputs: aHistoryComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyEventOnHistoryList"
	Dim iIsOpposition
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aHistoryComponent(B_COMPONENT_INITIALIZED_HISTORY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeHistoryComponent(oRequest, aHistoryComponent)
	End If

	If (aHistoryComponent(N_PROJECT_ID_HISTORY) = -1) Or (aHistoryComponent(N_TASK_ID_HISTORY) = -1) Or (aHistoryComponent(N_RECORD_ID_HISTORY) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del comentario a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "TaCoHistoryListComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If lErrorNumber = 0 Then
			If Not CheckHistoryListEventInformationConsistency(aHistoryComponent, sErrorDescription) Then
				lErrorNumber = -1
			Else
				sErrorDescription = "No se pudo modificar la información del comentario."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update " & TACO_PREFIX & "TasksHistoryList Set UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ", HistoryDescription='" & Replace(aHistoryComponent(S_DESCRIPTION_HISTORY), "'", "") & "' Where (ProjectID=" & aHistoryComponent(N_PROJECT_ID_HISTORY) & ") And (TaskID=" & aHistoryComponent(N_TASK_ID_HISTORY) & ") And (RecordID=" & aHistoryComponent(N_RECORD_ID_HISTORY) & ")", "TaCoHistoryListComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
		End If
	End If

	ModifyEventOnHistoryList = lErrorNumber
	Err.Clear
End Function

Function RemoveHistoryList(oRequest, oADODBConnection, aHistoryComponent, sErrorDescription)
'************************************************************
'Purpose: To remove the trademark history list from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aHistoryComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveHistoryList"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aHistoryComponent(B_COMPONENT_INITIALIZED_HISTORY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeHistoryComponent(oRequest, aHistoryComponent)
	End If

	If (aHistoryComponent(N_PROJECT_ID_HISTORY) = -1) Or (aHistoryComponent(N_TASK_ID_HISTORY) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó la actividad para eliminar su historial."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "TaCoHistoryListComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar el historial de la actividad."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & TACO_PREFIX & "TasksHistoryList Where (ProjectID=" & aHistoryComponent(N_PROJECT_ID_HISTORY) & ") And (TaskID=" & aHistoryComponent(N_TASK_ID_HISTORY) & ")", "TaCoHistoryListComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	RemoveHistoryList = lErrorNumber
	Err.Clear
End Function

Function RemoveEventOnHistoryList(oRequest, oADODBConnection, aHistoryComponent, sErrorDescription)
'************************************************************
'Purpose: To remove an event on the history list database
'Inputs:  oRequest, oADODBConnection
'Outputs: aHistoryComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveEventOnHistoryList"
	Dim iIsOpposition
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aHistoryComponent(B_COMPONENT_INITIALIZED_HISTORY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeHistoryComponent(oRequest, aHistoryComponent)
	End If

	If (aHistoryComponent(N_PROJECT_ID_HISTORY) = -1) Or (aHistoryComponent(N_TASK_ID_HISTORY) = -1) Or (aHistoryComponent(N_RECORD_ID_HISTORY) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se pudo eliminar la información del comentario."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "TaCoHistoryListComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo eliminar la información del comentario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & TACO_PREFIX & "TasksHistoryList Where (ProjectID=" & aHistoryComponent(N_PROJECT_ID_HISTORY) & ") And (TaskID=" & aHistoryComponent(N_TASK_ID_HISTORY) & ") And (RecordID=" & aHistoryComponent(N_RECORD_ID_HISTORY) & ")", "TaCoHistoryListComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	RemoveEventOnHistoryList = lErrorNumber
	Err.Clear
End Function

Function CheckHistoryListEventInformationConsistency(aHistoryComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the database
'Inputs:  aHistoryComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckHistoryListEventInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aHistoryComponent(N_PROJECT_ID_HISTORY)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del proyecto no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aHistoryComponent(N_TASK_ID_HISTORY)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador de la actividad no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aHistoryComponent(N_RECORD_ID_HISTORY)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del comentario no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aHistoryComponent(N_DATE_HISTORY)) Then aHistoryComponent(N_DATE_HISTORY) = Left(GetSerialNumberForDate(""), Len("00000000"))
	If Not IsNumeric(aHistoryComponent(N_HOUR_HISTORY)) Then aHistoryComponent(N_HOUR_HISTORY) = Hour(Time())
	If Not IsNumeric(aHistoryComponent(N_MINUTE_HISTORY)) Then aHistoryComponent(N_MINUTE_HISTORY) = Minute(Time())
	If Not IsNumeric(aHistoryComponent(N_USER_ID_HISTORY)) Then aHistoryComponent(N_USER_ID_HISTORY) = aLoginComponent(N_USER_ID_LOGIN)
	If Len(aHistoryComponent(S_DESCRIPTION_HISTORY)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El comentario está vacío."
		bIsCorrect = False
	End If

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información del comentario contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "TaCoHistoryListComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckHistoryListEventInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function DisplayHistoryListForm(oRequest, oADODBConnection, sAction, aHistoryComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about an event from the
'		  database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, aHistoryComponent
'Outputs: aHistoryComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayHistoryListForm"
	Dim sNames
	Dim sTempNames
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aHistoryComponent(B_COMPONENT_INITIALIZED_HISTORY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeHistoryComponent(oRequest, aHistoryComponent)
	End If

	If (aHistoryComponent(N_PROJECT_ID_HISTORY) <> -1) And (aHistoryComponent(N_TASK_ID_HISTORY) <> -1) And (aHistoryComponent(N_RECORD_ID_HISTORY) <> -1) Then
		lErrorNumber = GetHistoryListEvent(oRequest, oADODBConnection, aHistoryComponent, sErrorDescription)
	End If
	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		Response.Write "function CheckHistoryListFields(oForm) {" & vbNewLine
			Response.Write "if (oForm) {" & vbNewLine
				Response.Write "if (oForm.HistoryDescription.value.length == 0) {" & vbNewLine
					Response.Write "alert('Favor de introducir el comentario.');" & vbNewLine
					Response.Write "oForm.HistoryDescription.focus();" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "}" & vbNewLine
			Response.Write "return true;" & vbNewLine
		Response.Write "} // End of CheckHistoryListFields" & vbNewLine
	Response.Write "//--></SCRIPT>" & vbNewLine
	Response.Write "<FORM NAME=""HistoryListFrm"" ID=""HistoryListFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckHistoryListFields(this)"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""HistoryList"" />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProjectID"" ID=""ProjectIDHdn"" VALUE=""" & aHistoryComponent(N_PROJECT_ID_HISTORY) & """ />"	
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaskID"" ID=""TaskIDHdn"" VALUE=""" & aHistoryComponent(N_TASK_ID_HISTORY) & """ />"	
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ParentID"" ID=""ParentIDHdn"" VALUE=""" & oRequest("ParentID").Item & """ />"	
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaskPath"" ID=""TaskPathHdn"" VALUE=""" & oRequest("TaskPath").Item & """ />"	
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RecordID"" ID=""RecordIDHdn"" VALUE=""" & aHistoryComponent(N_RECORD_ID_HISTORY) & """ />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Tab"" ID=""TabHdn"" VALUE=""4"" />"

		Response.Write "<FONT FACE=""Arial"" SIZE=""2"">"
			If aHistoryComponent(N_RECORD_ID_HISTORY) = -1 Then
				Response.Write "Este comentario será registrado el<BR />"
			Else
				Response.Write "Este comentario fue registrado el<BR />"
			End If

			If aHistoryComponent(N_DATE_HISTORY) = 0 Then aHistoryComponent(N_DATE_HISTORY) = Left(GetSerialNumberForDate(""), Len("00000000"))
			If aHistoryComponent(N_HOUR_HISTORY) = 0 Then aHistoryComponent(N_HOUR_HISTORY) = Hour(Time())
			If aHistoryComponent(N_MINUTE_HISTORY) = 0 Then aHistoryComponent(N_MINUTE_HISTORY) = Minute(Time())
			If aHistoryComponent(N_USER_ID_HISTORY) = -1 Then aHistoryComponent(N_USER_ID_HISTORY) = aLoginComponent(N_USER_ID_LOGIN)
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""HistoryDate"" ID=""HistoryDateHdn"" VALUE=""" & aHistoryComponent(N_DATE_HISTORY) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""HistoryHour"" ID=""HistoryHourHdn"" VALUE=""" & aHistoryComponent(N_HOUR_HISTORY) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""HistoryMinute"" ID=""HistoryMinuteHdn"" VALUE=""" & aHistoryComponent(N_MINUTE_HISTORY) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserID"" ID=""UserIDHdn"" VALUE=""" & aHistoryComponent(N_USER_ID_HISTORY) & """ />"
			Response.Write DisplayDateFromSerialNumber(aHistoryComponent(N_DATE_HISTORY), aHistoryComponent(N_HOUR_HISTORY), aHistoryComponent(N_MINUTE_HISTORY), -1) & "<BR /><BR />"

			Response.Write "Comentario:<BR />"
		Response.Write "</FONT>"
		Response.Write "<TEXTAREA NAME=""HistoryDescription"" ID=""HistoryDescriptionTxtArea"" ROWS=""10"" COLS=""30"" MAXLENGTH=""2000"" CLASS=""TextFields"">" & aHistoryComponent(S_DESCRIPTION_HISTORY) & "</TEXTAREA><BR />"
		Response.Write "<BR />"

		If aHistoryComponent(N_RECORD_ID_HISTORY) = -1 Then
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
		ElseIf Len(oRequest("Delete").Item) > 0 Then
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveCommentWngDiv']); document.HistoryListFrm.Remove.focus()"" />"
		Else
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
		End If
		Response.Write "<BR /><BR />"
		Call DisplayWarningDiv("RemoveCommentWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
	Response.Write "</FORM>"

	DisplayHistoryListForm = lErrorNumber
	Err.Clear
End Function

Function DisplayHistoryListAsHiddenFields(oRequest, oADODBConnection, aHistoryListComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about an  history list
'		  event using hidden form fields
'Inputs:  oRequest, oADODBConnection, aHistoryListComponent
'Outputs: aHistoryListComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayHistoryListAsHiddenFields"
	Dim bComponentInitialized

	bComponentInitialized = aHistoryListComponent(B_COMPONENT_INITIALIZED_HISTORY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeHistoryComponent(oRequest, aHistoryListComponent)
	End If

	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProjectID"" ID=""ProjectIDHdn"" VALUE=""" & aHistoryComponent(N_PROJECT_ID_HISTORY) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaskID"" ID=""TaskIDHdn"" VALUE=""" & aHistoryComponent(N_TASK_ID_HISTORY) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RecordID"" ID=""RecordIDHdn"" VALUE=""" & aHistoryComponent(N_RECORD_ID_HISTORY) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""HistoryDate"" ID=""HistoryDateHdn"" VALUE=""" & aHistoryComponent(N_DATE_HISTORY) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""HistoryHour"" ID=""HistoryHourHdn"" VALUE=""" & aHistoryComponent(N_HOUR_HISTORY) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""HistoryMinute"" ID=""HistoryMinuteHdn"" VALUE=""" & aHistoryComponent(N_MINUTE_HISTORY) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserID"" ID=""UserIDHdn"" VALUE=""" & aHistoryComponent(N_USER_ID_HISTORY) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""HistoryDescription"" ID=""HistoryDescriptionHdn"" VALUE=""" & aHistoryComponent(S_DESCRIPTION_HISTORY) & """ />"

	DisplayHistoryListAsHiddenFields = Err.number
	Err.Clear
End Function

Function DisplayHistoryList(oRequest, oADODBConnection, aHistoryComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about the history list
'		  from the database using HTML
'Inputs:  oRequest, oADODBConnection, aHistoryComponent
'Outputs: aHistoryComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayHistoryList"
	Dim sURL
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aHistoryComponent(B_COMPONENT_INITIALIZED_HISTORY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeHistoryComponent(oRequest, aHistoryComponent)
	End If

	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Historial de comentarios:</B><BR /><BR /></FONT>"

	Response.Write "<DIV CLASS=""HistoryList""><FONT FACE=""Arial"" SIZE=""2"">" & vbNewLine
		lErrorNumber = GetHistoryList(oRequest, oADODBConnection, aHistoryComponent, oRecordset, sErrorDescription)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				Do While Not oRecordset.EOF
					Response.Write "<B>" & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("HistoryDate").Value), CInt(oRecordset.Fields("HistoryHour").Value), CInt(oRecordset.Fields("HistoryMinute").Value), -1) & "</B><IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""1"" />"
					If CInt(oRecordset.Fields("bSystemMessage").Value) = 0 Then
						sURL = "ProjectID=" & aHistoryComponent(N_PROJECT_ID_HISTORY) & "&TaskID=" & aHistoryComponent(N_TASK_ID_HISTORY) & "&ParentID=" & oRequest("ParentID").Item & "&TaskPath=" & oRequest("TaskPath").Item & "&Tab=4&Action=HistoryList&"
						If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then
							Response.Write "<A HREF=""" & GetASPFileName("") & "?" & sURL & "RecordID=" & CStr(oRecordset.Fields("RecordID").Value) & "&Change=1&Tab=2""><IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" /></A>"
							Response.Write "<IMG SRC=""Images/Transaprent.gif"" WIDTH=""20"" HEIGHT=""1"" />"
						End If
						If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then
							Response.Write "<A HREF=""" & GetASPFileName("") & "?" & sURL & "RecordID=" & CStr(oRecordset.Fields("RecordID").Value) & "&Delete=1&Tab=2""><IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar"" BORDER=""0"" /></A>"
						End If
					End If
					Response.Write "<BR />"
					Response.Write CleanStringForHTML(CStr(oRecordset.Fields("UserName").Value) & " " & CStr(oRecordset.Fields("UserLastName").Value)) & ": "
					Response.Write CleanStringForHTML(CStr(oRecordset.Fields("HistoryDescription").Value)) & "<BR />"

					Response.Flush
					oRecordset.MoveNext
					If Not oRecordset.EOF Then Response.Write "<BR />"
					If Err.number <> 0 Then Exit Do
				Loop
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen comentarios registrados en la base de datos."
			End If
		End If
	Response.Write "</FONT></DIV>" & vbNewLine

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayHistoryList = lErrorNumber
	Err.Clear
End Function

Function DisplayHistoryListTxt(oRequest, oADODBConnection, bUseHTML, aHistoryComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about the history list
'		  from the database using text
'Inputs:  oRequest, oADODBConnection, bUseHTML, aHistoryComponent
'Outputs: aHistoryComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayHistoryListTxt"
	Dim sBoldBegin
	Dim sBoldEnd
	Dim sNewLine
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aHistoryComponent(B_COMPONENT_INITIALIZED_HISTORY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeHistoryComponent(oRequest, aHistoryComponent)
	End If

	lErrorNumber = GetHistoryList(oRequest, oADODBConnection, aHistoryComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If bUseHTML Then
			sBoldBegin = "<B>"
			sBoldEnd = "</B>"
			sNewLine = "<BR />"
		Else
			sNewLine = vbNewLine
		End If
		If bUseHTML Then Response.Write "<FONT FACE=""Arial"" SIZE=""2"">"
			If Not oRecordset.EOF Then
				Response.Write sBoldBegin & "Historial de comentarios:" & sBoldEnd & sNewLine & sNewLine
				Do While Not oRecordset.EOF
					Response.Write sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("HistoryDate").Value), CInt(oRecordset.Fields("HistoryHour").Value), CInt(oRecordset.Fields("HistoryMinute").Value), -1) & sBoldEnd & sNewLine
					Response.Write CStr(oRecordset.Fields("UserName").Value) & " " & CStr(oRecordset.Fields("UserLastName").Value) & ": "
					Response.Write CStr(oRecordset.Fields("HistoryDescription").Value) & sNewLine

					Response.Flush
					oRecordset.MoveNext
					If Not oRecordset.EOF Then
						Response.Write sNewLine
					End If
					If Err.number <> 0 Then Exit Do
				Loop
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen comentarios registrados en la base de datos."
			End If
		If bUseHTML Then Response.Write "</FONT>"
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayHistoryListTxt = lErrorNumber
	Err.Clear
End Function
%>