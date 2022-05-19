<%
Const L_ID_USER_OPTIONS = 0
Const A_OPTIONS = 1
Const B_COMPONENT_INITIALIZED_OPTIONS = 2

Const N_OPTIONS_COMPONENT_SIZE = 2

Const DEFAULT_OPTIONS = "1;;;Main.asp;;;1;;;1;;;100;;;EmployeeNumber;;;0;;;CheckNumber;;;0;;;2;;;0;;;0;;;0;;;0;;;0"
Const SHOW_PRINT_INFO_OPTION = 0
Const START_PAGE_OPTION = 1
Const TABLE_STYLE_OPTION = 2
Const EXPORT_FILTER_OPTION = 3
Const REPORT_ROWS_OPTION = 4
Const EMPLOYEE_ORDER_OPTION = 5
Const EMPLOYEE_SORT_OPTION = 6
Const PAYMENT_ORDER_OPTION = 7
Const PAYMENT_SORT_OPTION = 8
Const TRESHOLD_STYLE_OPTION = 9
Const FULL_PROJECT_OPTION = 10
Const CHECKS_LEFT_MARGIN1_OPTION = 11
Const CHECKS_TOP_MARGIN1_OPTION = 12
Const CHECKS_LEFT_MARGIN2_OPTION = 13
Const CHECKS_TOP_MARGIN2_OPTION = 14

Const DEFAULT_ADMIN_OPTIONS = ";;;5;;;vic_arjona@yahoo.com;;;30;;;Víctor Arjona;;;52.86.91.04;;;vic_arjona@yahoo.com;;;1;;;1;;;0;;;FF0000;;;FFFF00;;;00FF00;;;0;;;50;;;100;;;134.19;;;167.74;;;33.55;;;0.25"
Const LOGIN_FAILURES_OPTION = 1
Const SYSTEM_BLOCKED_RECIPIENTS_OPTION = 2
Const PASSWORDS_DAYS_OPTION = 3
Const CONTACT_NAME_OPTION = 4
Const CONTACT_PHONE_OPTION = 5
Const CONTACT_EMAIL_OPTION = 6
Const UPDATE_OPTION = 7
Const DELETE_OPTION = 8
Const INSERT_OPTION = 9
Const RED_COLOR_OPTION = 10
Const YELLOW_COLOR_OPTION = 11
Const GREEN_COLOR_OPTION = 12
Const RED_TRESHOLD_OPTION = 13
Const YELLOW_TRESHOLD_OPTION = 14
Const GREEN_TRESHOLD_OPTION = 15
Const FONAC_01_OPTION = 16
Const FONAC_02_OPTION = 17
Const FONAC_03_OPTION = 18
Const FONAC_04_OPTION = 19

Dim aDefaultAdminOptions
aDefaultAdminOptions = Split(DEFAULT_ADMIN_OPTIONS, LIST_SEPARATOR, -1, vbBinaryCompare)
Dim aDefaultOptions
aDefaultOptions = Split(DEFAULT_OPTIONS, LIST_SEPARATOR, -1, vbBinaryCompare)

Dim aAdminOptionsComponent()
Redim aAdminOptionsComponent(N_OPTIONS_COMPONENT_SIZE)
aAdminOptionsComponent(L_ID_USER_OPTIONS) = -2
If lErrorNumber = 0 Then
	lErrorNumber = GetAdminOptions(oRequest, oADODBConnection, aAdminOptionsComponent, sErrorDescription)
End If

Dim aOptionsComponent()
Redim aOptionsComponent(N_OPTIONS_COMPONENT_SIZE)

Dim sOptionsErrorDescription

Function InitializeOptionsComponent(oRequest, aOptionsComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Options Component
'         using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aOptionsComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeOptionsComponent"
	Redim Preserve aOptionsComponent(N_OPTIONS_COMPONENT_SIZE)

	If IsEmpty(aOptionsComponent(L_ID_USER_OPTIONS)) Then
		If Len(oRequest("UserID").Item) > 0 Then
			aOptionsComponent(L_ID_USER_OPTIONS) = CLng(oRequest("UserID").Item)
		ElseIf Len(oRequest("AccessKey").Item) > 0 Then
			Call GetNameFromTable(oADODBConnection, "UserAccessKey", "'" & oRequest("AccessKey").Item & "'", "", "", aOptionsComponent(L_ID_USER_OPTIONS), "")
			aOptionsComponent(L_ID_USER_OPTIONS) = CLng(aOptionsComponent(L_ID_USER_OPTIONS))
		Else
			aOptionsComponent(L_ID_USER_OPTIONS) = -2
		End If
	End If

	If IsEmpty(aOptionsComponent(A_OPTIONS)) Then
		If Len(oRequest("UserOptions").Item) > 0 Then
			aOptionsComponent(A_OPTIONS) = Split(oRequest("UserOptions").Item, LIST_SEPARATOR, -1, vbBinaryCompare)
		Else
			If aOptionsComponent(L_ID_USER_OPTIONS) <> -2 Then
				aOptionsComponent(A_OPTIONS) = aDefaultOptions
			Else
				aOptionsComponent(A_OPTIONS) = aDefaultAdminOptions
			End If
		End If
	End If

	aOptionsComponent(B_COMPONENT_INITIALIZED_OPTIONS) = True
	InitializeOptionsComponent = Err.number
	Err.Clear
End Function

Function UpdateOptionsComponent(oRequest, aOptionsComponent)
'************************************************************
'Purpose: To update the empty elements of the Options Component
'         using the URL parameters
'Inputs:  oRequest
'Outputs: aOptionsComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UpdateOptionsComponent"
	Dim aGroupsIDentifier

	If Len(oRequest("UserID").Item) > 0 Then
		aOptionsComponent(L_ID_USER_OPTIONS) = CLng(oRequest("UserID").Item)
	End If

	If Len(oRequest("UserOptions").Item) > 0 Then
		aOptionsComponent(A_OPTIONS) = Split(oRequest("UserOptions").Item, LIST_SEPARATOR, -1, vbBinaryCompare)
	End If

	UpdateOptionsComponent = Err.number
	Err.Clear
End Function

Function AddOptions(oRequest, oADODBConnection, aOptionsComponent, sErrorDescription)
'************************************************************
'Purpose: To add the user preferences into the database
'Inputs:  oRequest, oADODBConnection, aOptionsComponent
'Outputs: aOptionsComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddOptions"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aOptionsComponent(B_COMPONENT_INITIALIZED_OPTIONS)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeOptionsComponent(oRequest, aOptionsComponent)
	End If

	If aOptionsComponent(L_ID_USER_OPTIONS) = -2 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del usuario para agregar sus preferencias."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "OptionsComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If Not CheckOptionsInformationConsistency(aOptionsComponent, sErrorDescription) Then
			lErrorNumber = -1
		Else
			sOptionsErrorDescription = sErrorDescription
			sErrorDescription = "No se pudieron agregar las preferencias del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Preferences (UserID, UserPreferences) Values (" & aOptionsComponent(L_ID_USER_OPTIONS) & ", '" & Join(aOptionsComponent(A_OPTIONS), LIST_SEPARATOR) & "')", "OptionsComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	AddOptions = lErrorNumber
	Err.Clear
End Function

Function GetOption(aOptionsComponent, iOption)
'************************************************************
'Purpose: To get the specified user preference
'Inputs:  aOptionsComponent, iOption
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetOption"
	Dim sOptionValue
	Dim bComponentInitialized

	bComponentInitialized = aOptionsComponent(B_COMPONENT_INITIALIZED_OPTIONS)
	If (Not IsEmpty(bComponentInitialized)) And (bComponentInitialized) Then
		sOptionValue = ""
		If Not IsEmpty(aOptionsComponent(A_OPTIONS)(iOption)) Then
			sOptionValue = aOptionsComponent(A_OPTIONS)(iOption)
		Else
			sOptionValue = aDefaultOptions(iOption)
		End If
	Else
		sOptionValue = aDefaultOptions(iOption)
	End If
	If IsEmpty(sOptionValue) Then sOptionValue = aDefaultOptions(iOption)

	GetOption = sOptionValue
	Err.Clear
End Function

Function GetAdminOption(aAdminOptionsComponent, iOption)
'************************************************************
'Purpose: To get the specified administration option
'Inputs:  aAdminOptionsComponent, iOption
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetAdminOption"
	Dim sOptionValue
	Dim bComponentInitialized

	bComponentInitialized = aAdminOptionsComponent(B_COMPONENT_INITIALIZED_OPTIONS)
	If (Not IsEmpty(bComponentInitialized)) And (bComponentInitialized) Then
		sOptionValue = ""
		If Not IsEmpty(aAdminOptionsComponent(A_OPTIONS)(iOption)) Then
			sOptionValue = aAdminOptionsComponent(A_OPTIONS)(iOption)
		Else
			sOptionValue = aDefaultAdminOptions(iOption)
		End If
	Else
		sOptionValue = aDefaultAdminOptions(iOption)
	End If
	If IsEmpty(sOptionValue) Then sOptionValue = aDefaultAdminOptions(iOption)

	GetAdminOption = sOptionValue
	Err.Clear
End Function

Function GetOptions(oRequest, oADODBConnection, aOptionsComponent, sErrorDescription)
'************************************************************
'Purpose: To get the user preferences from the database
'Inputs:  oRequest, oADODBConnection, aOptionsComponent
'Outputs: aOptionsComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetOptions"
	Dim iBound
	Dim aTempOptions
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aOptionsComponent(B_COMPONENT_INITIALIZED_OPTIONS)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeOptionsComponent(oRequest, aOptionsComponent)
	End If

	If aOptionsComponent(L_ID_USER_OPTIONS) = -2 Then
'		lErrorNumber = -1
'		sErrorDescription = "No se especificó el identificador del usuario para obtener sus preferencias."
'		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "OptionsComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudieron obtener las preferencias del usuario."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select UserPreferences From Preferences Where (UserID=" & aOptionsComponent(L_ID_USER_OPTIONS) & ")", "OptionsComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				aOptionsComponent(A_OPTIONS) = aDefaultOptions
			Else
				aOptionsComponent(A_OPTIONS) = Split(CStr(oRecordset.Fields("UserPreferences").Value), LIST_SEPARATOR, -1, vbBinaryCompare)
				iBound = UBound(aDefaultOptions)
				If UBound(aDefaultOptions) <> UBound(aOptionsComponent(A_OPTIONS)) Then
					aTempOptions = aOptionsComponent(A_OPTIONS)
					Redim Preserve aTempOptions(iBound)
					aOptionsComponent(A_OPTIONS) = aTempOptions
				End If
				Call CheckOptionsInformationConsistency(aOptionsComponent, sErrorDescription)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetOptions = lErrorNumber
	Err.Clear
End Function

Function GetAdminOptions(oRequest, oADODBConnection, aAdminOptionsComponent, sErrorDescription)
'************************************************************
'Purpose: To get the administration options from the database
'Inputs:  oRequest, oADODBConnection, aAdminOptionsComponent
'Outputs: aAdminOptionsComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetAdminOptions"
	Dim iBound
	Dim aTempOptions
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAdminOptionsComponent(B_COMPONENT_INITIALIZED_OPTIONS)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeOptionsComponent(oRequest, aAdminOptionsComponent)
	End If

	sErrorDescription = "No se pudieron obtener las opciones del sistema."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select UserPreferences From Preferences Where (UserID=-2)", "OptionsComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			aAdminOptionsComponent(A_OPTIONS) = aDefaultAdminOptions
		Else
			aAdminOptionsComponent(A_OPTIONS) = Split(CStr(oRecordset.Fields("UserPreferences").Value), LIST_SEPARATOR, -1, vbBinaryCompare)
			iBound = UBound(aDefaultAdminOptions)
			If UBound(aDefaultAdminOptions) <> UBound(aAdminOptionsComponent(A_OPTIONS)) Then
				aTempOptions = aAdminOptionsComponent(A_OPTIONS)
				Redim Preserve aTempOptions(iBound)
				aAdminOptionsComponent(A_OPTIONS) = aTempOptions
			End If
			Call CheckAdminOptionsInformationConsistency(aAdminOptionsComponent, sErrorDescription)
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	GetAdminOptions = lErrorNumber
	Err.Clear
End Function

Function SetOption(aOptionsComponent, iOption, vValue, sErrorDescription)
'************************************************************
'Purpose: To set the specified user preference
'Inputs:  aOptionsComponent, iOption, vValue
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetOption"

	aOptionsComponent(A_OPTIONS)(iOption) = vValue
	Call CheckOptionsInformationConsistency(aOptionsComponent, sErrorDescription)

	SetOption = Err.number
	Err.Clear
End Function

Function SetOptions(oRequest, aOptionsComponent, sErrorDescription)
'************************************************************
'Purpose: To set the user preferences using the Request Object
'Inputs:  oRequest, aOptionsComponent
'Outputs: aOptionsComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetOptions"
	Dim oItem
	Dim sIndex
	Dim bComponentInitialized

	bComponentInitialized = aOptionsComponent(B_COMPONENT_INITIALIZED_OPTIONS)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeOptionsComponent(oRequest, aOptionsComponent)
	End If

	For Each oItem In oRequest
		If InStr(1, CStr(oItem), "P", vbBinaryCompare) = 1 Then
			sIndex = Right(CStr(oItem), (Len(CStr(oItem)) - Len("P")))
			If IsNumeric(sIndex) Then
				aOptionsComponent(A_OPTIONS)(CLng(sIndex)) = CStr(oRequest(oItem).Item)
			End If
		End If
		If Err.number <> 0 Then Exit For
	Next

	SetOptions = Err.number
	Err.Clear
End Function

Function ModifyOptions(oRequest, oADODBConnection, aOptionsComponent, sErrorDescription)
'************************************************************
'Purpose: To modify the user preferences
'Inputs:  oRequest, oADODBConnection, aOptionsComponent
'Outputs: aOptionsComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyOptions"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aOptionsComponent(B_COMPONENT_INITIALIZED_OPTIONS)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeOptionsComponent(oRequest, aOptionsComponent)
	End If

	If aOptionsComponent(L_ID_USER_OPTIONS) = -2 Then
'		lErrorNumber = -1
'		sErrorDescription = "No se especificó el identificador del usuario para modificar sus preferencias."
'		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "OptionsComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sOptionsErrorDescription = sErrorDescription
		If Not CheckOptionsInformationConsistency(aOptionsComponent, sErrorDescription) Then
			lErrorNumber = -1
		Else
			sOptionsErrorDescription = sErrorDescription
			sErrorDescription = "No se pudieron verificar la existencia de las preferencias del usuario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select UserPreferences From Preferences Where (UserID=" & aOptionsComponent(L_ID_USER_OPTIONS) & ")", "OptionsComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sErrorDescription = "No se pudieron modificar las preferencias del usuario."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Preferences Set UserPreferences='" & Join(aOptionsComponent(A_OPTIONS), LIST_SEPARATOR) & "' Where (UserID=" & aOptionsComponent(L_ID_USER_OPTIONS) & ")", "OptionsComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				Else
					lErrorNumber = AddOptions(oRequest, oADODBConnection, aOptionsComponent, sErrorDescription)
				End If
				oRecordset.Close
			End If
		End If
	End If

	Set oRecordset = Nothing
	ModifyOptions = lErrorNumber
	Err.Clear
End Function

Function ModifyAdminOptions(oRequest, oADODBConnection, aAdminOptionsComponent, sErrorDescription)
'************************************************************
'Purpose: To modify the administration options
'Inputs:  oRequest, oADODBConnection, aAdminOptionsComponent
'Outputs: aAdminOptionsComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyAdminOptions"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAdminOptionsComponent(B_COMPONENT_INITIALIZED_OPTIONS)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeOptionsComponent(oRequest, aAdminOptionsComponent)
	End If

	If Not CheckAdminOptionsInformationConsistency(aAdminOptionsComponent, sErrorDescription) Then
		lErrorNumber = -1
	Else
		sOptionsErrorDescription = sErrorDescription
		sErrorDescription = "No se pudieron modificar las opciones del sistema."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Preferences Set UserPreferences='" & Join(aAdminOptionsComponent(A_OPTIONS), LIST_SEPARATOR) & "' Where (UserID=-2)", "OptionsComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	ModifyAdminOptions = lErrorNumber
	Err.Clear
End Function

Function RemoveOptions(oRequest, oADODBConnection, aOptionsComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a note from the database
'Inputs:  oRequest, oADODBConnection, aOptionsComponent
'Outputs: aOptionsComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveOptions"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aOptionsComponent(B_COMPONENT_INITIALIZED_OPTIONS)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeOptionsComponent(oRequest, aOptionsComponent)
	End If

	If aOptionsComponent(L_ID_USER_OPTIONS) = -2 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del usuario para eliminar sus preferencias."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "OptionsComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudieron eliminar las preferencias del usuario."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Preferences Where (UserID=" & aOptionsComponent(L_ID_USER_OPTIONS) & ")", "OptionsComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	RemoveOptions = lErrorNumber
	Err.Clear
End Function

Function CheckAdminOptionsInformationConsistency(aAdminOptionsComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the database
'Inputs:  aAdminOptionsComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckAdminOptionsInformationConsistency"
	Dim iIndex
	Dim bCorrect
	Dim sIncorrectOptions

	sIncorrectOptions = ""
	If Not IsNumeric(aAdminOptionsComponent(L_ID_USER_OPTIONS)) Then aAdminOptionsComponent(L_ID_USER_OPTIONS) = -2

	For iIndex = 1 To UBound(aAdminOptionsComponent(A_OPTIONS))
		bCorrect = True
		Select Case iIndex
			Case LOGIN_FAILURES_OPTION
				bCorrect = CheckNumericRange(aAdminOptionsComponent(A_OPTIONS)(iIndex), 3, 100)
			Case SYSTEM_BLOCKED_RECIPIENTS_OPTION, CONTACT_NAME_OPTION, CONTACT_PHONE_OPTION, CONTACT_EMAIL_OPTION, RED_COLOR_OPTION, YELLOW_COLOR_OPTION, GREEN_COLOR_OPTION
			Case PASSWORDS_DAYS_OPTION
				bCorrect = CheckNumericRange(aAdminOptionsComponent(A_OPTIONS)(iIndex), 30, 365)
			Case UPDATE_OPTION, DELETE_OPTION, INSERT_OPTION
				bCorrect = CheckNumericRange(aAdminOptionsComponent(A_OPTIONS)(iIndex), 0, 1)
			Case RED_TRESHOLD_OPTION, YELLOW_TRESHOLD_OPTION, GREEN_TRESHOLD_OPTION
				bCorrect = CheckNumericRange(aAdminOptionsComponent(A_OPTIONS)(iIndex), 0, 100)
			Case FONAC_01_OPTION, FONAC_02_OPTION, FONAC_03_OPTION, FONAC_04_OPTION
				bCorrect = CheckNumericRange(aAdminOptionsComponent(A_OPTIONS)(iIndex), 0, 9999999999)
		End Select
		If Not bCorrect Then
			sIncorrectOptions = sIncorrectOptions & "<BR />&nbsp;" & iIndex & ": " & aAdminOptionsComponent(A_OPTIONS)(iIndex) & "(valor usado: " & aAdminOptionsComponent(A_OPTIONS)(iIndex)& ")"
			aAdminOptionsComponent(A_OPTIONS)(iIndex) = aDefaultAdminOptions(iIndex)
		End If
	Next

	If Len(sIncorrectOptions) > 0 Then
		sErrorDescription = "Había opciones del sistema con valores erróneos. Se utilizaron los valores originales para corregir el problema.<BR /><BR />Las opciones incorrectas son:" & sIncorrectOptions
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "OptionsComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckAdminOptionsInformationConsistency = True
	Err.Clear
End Function

Function CheckOptionsInformationConsistency(aOptionsComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the database
'Inputs:  aOptionsComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckOptionsInformationConsistency"
	Dim iIndex
	Dim bCorrect
	Dim bIsCorrect
	Dim sIncorrectOptions

	bIsCorrect = True
	sIncorrectOptions = ""
	If Not IsNumeric(aOptionsComponent(L_ID_USER_OPTIONS)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del usuario no es un valor numérico."
		bIsCorrect = False
	End If

	For iIndex = 1 To UBound(aOptionsComponent(A_OPTIONS))
		bCorrect = True
		Select Case iIndex
			Case SHOW_PRINT_INFO_OPTION, EXPORT_FILTER_OPTION, EMPLOYEE_SORT_OPTION, PAYMENT_SORT_OPTION, FULL_PROJECT_OPTION
				bCorrect = CheckNumericRange(aOptionsComponent(A_OPTIONS)(iIndex), 0, 1)
			Case START_PAGE_OPTION
				If InStr(1, aOptionsComponent(A_OPTIONS)(iIndex), ".asp", vbTextCompare) = 0 Then bCorrect = False
			Case TABLE_STYLE_OPTION
				bCorrect = CheckNumericRange(aOptionsComponent(A_OPTIONS)(iIndex), 1, 2)
			Case REPORT_ROWS_OPTION
				bCorrect = CheckNumericRange(aOptionsComponent(A_OPTIONS)(iIndex), 10, 200)
			Case EMPLOYEE_ORDER_OPTION, PAYMENT_ORDER_OPTION
			Case TRESHOLD_STYLE_OPTION
				bCorrect = CheckNumericRange(aOptionsComponent(A_OPTIONS)(iIndex), 1, 3)
			Case CHECKS_LEFT_MARGIN1_OPTION, CHECKS_TOP_MARGIN1_OPTION, CHECKS_LEFT_MARGIN2_OPTION, CHECKS_TOP_MARGIN2_OPTION
				bCorrect = CheckNumericRange(aOptionsComponent(A_OPTIONS)(iIndex), -20000, 20000)
		End Select
		If Not bCorrect Then
			sIncorrectOptions = sIncorrectOptions & "<BR />&nbsp;" & iIndex & ": " & aOptionsComponent(A_OPTIONS)(iIndex) & "(valor usado: " & aDefaultOptions(iIndex)& ")"
			aOptionsComponent(A_OPTIONS)(iIndex) = aDefaultOptions(iIndex)
		End If
	Next

	If Len(sIncorrectOptions) > 0 Then
		sErrorDescription = "Había preferencias con valores erróneos. Se utilizaron los valores originales para corregir el problema.<BR /><BR />Las opciones incorrectas son:" & sIncorrectOptions & ". " & Err.Description
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "OptionsComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckOptionsInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function CheckNumericRange(lValue, lMinValue, lMaxValue)
'************************************************************
'Purpose: To check the first input is a number and is in the
'		  given range
'Inputs:  lValue, lMinValue, lMaxValue
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckNumericRange"
	Dim lTempValue
	Dim bCorrect

	bCorrect = True
	If Len(lValue) = 0 Then
		bCorrect = False
	Else
		If Not IsNumeric(lValue) Then
			bCorrect = False
		Else
			lTempValue = CLng(lValue)
			If (lTempValue < lMinValue) Then bCorrect = False
			If (lMaxValue > lMinValue) Then
				If (lTempValue > lMaxValue) Then bCorrect = False
			End If
		End If
	End If
	CheckNumericRange = bCorrect
End Function
%>