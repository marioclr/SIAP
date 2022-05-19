<%
Const N_NOTHING_ALARM_ACTION = 0

Dim sSortColumnForAlarms
Dim bSortDescendingForAlarms
Dim iAction
If Len(oRequest("Action").Item) = 0 Then
	iAction = N_NOTHING_ALARM_ACTION
Else
	iAction = CInt(oRequest("Action").Item)
End If

Function SetSortOrderForAlarms(oRequest, oADODBConnection, sSortColumnForAlarms, bSortDescendingForAlarms, sErrorDescription)
'************************************************************
'Purpose: To get the sort order from the URL or from the user
'         preferences
'Inputs:  oRequest, oADODBConnection
'Outputs: sSortColumnForAlarms, bSortDescendingForAlarms, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetSortOrderForAlarms"
	Dim lErrorNumber

	If Len(oRequest("SortColumn").Item) > 0 Then
		sSortColumnForAlarms = oRequest("SortColumn").Item
		Select Case iAction
'			Case N_POLICIES_ALARM_ACTION
'				Call SetOption(aOptionsComponent, SORT_POLICIES_ALARM_OPTION, sSortColumnForAlarms, sErrorDescription)
'				If Len(oRequest("Desc").Item) > 0 Then bSortDescendingForAlarms = (StrComp(oRequest("Desc").Item, "0", vbBinaryCompare) <> 0)
'				Call SetOption(aOptionsComponent, ORDER_POLICIES_ALARM_OPTION, bSortDescendingForAlarms, sErrorDescription)
		End Select
'		Call ModifyOptions(oRequest, oADODBConnection, aOptionsComponent, sErrorDescription)
	Else
		Select Case iAction
'			Case N_POLICIES_ALARM_ACTION
'				sSortColumnForAlarms = GetOption(aOptionsComponent, SORT_POLICIES_ALARM_OPTION)
'				If IsEmpty(oRequest("Desc").Item) Then bSortDescendingForAlarms = GetOption(aOptionsComponent, ORDER_POLICIES_ALARM_OPTION)
'				If Len(oRequest("Desc").Item) = 0 Then bSortDescendingForAlarms = GetOption(aOptionsComponent, ORDER_POLICIES_ALARM_OPTION)
			Case Else
'				sSortColumnForAlarms = "ReportNumber"
				bSortDescendingForAlarms = (Len(oRequest("Desc").Item) > 0)
		End Select
	End If

	SetSortOrderForAlarms = lErrorNumber
	Err.Clear
End Function
%>