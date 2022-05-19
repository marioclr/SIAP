<%
Const N_PROJECT_ID_TASK = 0
Const N_ID_TASK = 1
Const N_PARENT_ID_TASK = 2
Const S_PATH_TASK = 3
Const S_NAME_TASK = 4
Const S_NUMBER_TASK = 5
Const N_LABEL_ID_TASK = 6
Const S_DESCRIPTION_TASK = 7
Const S_OBJECTIVE_TASK = 8
Const S_STRATEGY_TASK = 9
Const S_PURPOUSE_TASK = 10
Const S_INDICATOR_TASK = 11
Const S_MEASUREMENT_TASK = 12
Const S_FORMULA_TASK = 13
Const N_AGGREGATION_TYPE_TASK = 14
Const S_COMMENTS_TASK = 15
Const N_START_DATE_TASK = 16
Const N_END_DATE_TASK = 17
Const S_FILE_TASK = 18
Const S_PROJECT_FILE_TASK = 19
Const N_PUCO_SECTION_ID_TASK = 20
Const N_FORM_TASK = 21
Const S_REPORT_URL_TASK = 22
Const AN_PARENT_ID_TASK = 23
Const AD_MINIMUM_VALUE_TASK = 24
Const AD_AVERAGE_VALUE_TASK = 25
Const AD_MAXIMUM_VALUE_TASK = 26
Const AD_TARGET_VALUE_TASK = 27
Const AN_FIELD_TYPE_TASK = 28
Const AD_PERCENTAGE_TASK = 29
Const AD_REQUIRED_TASK = 30
Const AL_VARIABLES_TASK = 31
Const AD_VARIABLES_MINIMUM_VALUES_TASK = 32
Const AD_VARIABLES_AVERAGE_VALUES_TASK = 33
Const AD_VARIABLES_MAXIMUM_VALUES_TASK = 34
Const AD_VARIABLES_TARGET_VALUES_TASK = 35
Const AN_VARIABLES_FIELD_TYPES_TASK = 36
Const AN_VARIABLES_RELEVANCE_TASK = 37
Const N_START_DATE_STATUS_TASK = 38
Const N_END_DATE_STATUS_TASK = 39
Const D_VALUE_STATUS_TASK = 40
Const D_PERCENTAGE_STATUS_TASK = 41
Const L_STATUS_ID_TASK = 42
Const S_AREAS_TASK = 43
Const S_USERS_TASK = 44
Const S_CATEGORIES_TASK = 45
Const B_HAS_CHILDREN_TASK = 46
Const N_EASY_TASK = 47
Const S_QUERY_CONDITION_TASK = 48
Const B_CHECK_FOR_DUPLICATED_TASK = 49
Const B_IS_DUPLICATED_TASK = 50
Const B_COMPONENT_INITIALIZED_TASK = 51

Const N_TASK_COMPONENT_SIZE = 51

Dim aTaskComponent()
Redim aTaskComponent(N_TASK_COMPONENT_SIZE)

Function InitializeTaskComponent(oRequest, aTaskComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Task Component
'         using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aTaskComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeTaskComponent"
	Dim oItem
	Dim iIndex
	Redim Preserve aTaskComponent(N_TASK_COMPONENT_SIZE)

	If IsEmpty(aTaskComponent(N_PROJECT_ID_TASK)) Then
		If Len(oRequest("ProjectID").Item) > 0 Then
			aTaskComponent(N_PROJECT_ID_TASK) = CLng(oRequest("ProjectID").Item)
		Else
			aTaskComponent(N_PROJECT_ID_TASK) = -1
		End If
	End If

	If IsEmpty(aTaskComponent(N_ID_TASK)) Then
		If Len(oRequest("TaskID").Item) > 0 Then
			aTaskComponent(N_ID_TASK) = CLng(oRequest("TaskID").Item)
		Else
			aTaskComponent(N_ID_TASK) = -1
		End If
	End If

	If IsEmpty(aTaskComponent(N_PARENT_ID_TASK)) Then
		If Len(oRequest("ParentID").Item) > 0 Then
			aTaskComponent(N_PARENT_ID_TASK) = CLng(oRequest("ParentID").Item)
		Else
			aTaskComponent(N_PARENT_ID_TASK) = -1
		End If
	End If

	If IsEmpty(aTaskComponent(S_PATH_TASK)) Then
		If Len(oRequest("TaskPath").Item) > 0 Then
			aTaskComponent(S_PATH_TASK) = oRequest("TaskPath").Item
		Else
			aTaskComponent(S_PATH_TASK) = aTaskComponent(N_ID_TASK)
		End If
	End If

	If IsEmpty(aTaskComponent(S_NAME_TASK)) Then
		If Len(oRequest("TaskName").Item) > 0 Then
			aTaskComponent(S_NAME_TASK) = oRequest("TaskName").Item
		Else
			aTaskComponent(S_NAME_TASK) = ""
		End If
	End If
	aTaskComponent(S_NAME_TASK) = Left(aTaskComponent(S_NAME_TASK), 255)

	If IsEmpty(aTaskComponent(S_NUMBER_TASK)) Then
		If Len(oRequest("TaskNumber").Item) > 0 Then
			aTaskComponent(S_NUMBER_TASK) = oRequest("TaskNumber").Item
		Else
			aTaskComponent(S_NUMBER_TASK) = ""
		End If
	End If
	aTaskComponent(S_NUMBER_TASK) = Left(aTaskComponent(S_NUMBER_TASK), 30)

	If IsEmpty(aTaskComponent(N_LABEL_ID_TASK)) Then
		If Len(oRequest("LabelID").Item) > 0 Then
			aTaskComponent(N_LABEL_ID_TASK) = CLng(oRequest("LabelID").Item)
		Else
			aTaskComponent(N_LABEL_ID_TASK) = -1
		End If
	End If

	If IsEmpty(aTaskComponent(S_DESCRIPTION_TASK)) Then
		If Len(oRequest("TaskDescription").Item) > 0 Then
			aTaskComponent(S_DESCRIPTION_TASK) = oRequest("TaskDescription").Item
		Else
			aTaskComponent(S_DESCRIPTION_TASK) = ""
		End If
	End If
	aTaskComponent(S_DESCRIPTION_TASK) = Left(aTaskComponent(S_DESCRIPTION_TASK), 4000)

	If IsEmpty(aTaskComponent(S_OBJECTIVE_TASK)) Then
		If Len(oRequest("TaskObjective").Item) > 0 Then
			aTaskComponent(S_OBJECTIVE_TASK) = oRequest("TaskObjective").Item
		Else
			aTaskComponent(S_OBJECTIVE_TASK) = ""
		End If
	End If
	aTaskComponent(S_OBJECTIVE_TASK) = Left(aTaskComponent(S_OBJECTIVE_TASK), 4000)

	If IsEmpty(aTaskComponent(S_STRATEGY_TASK)) Then
		If Len(oRequest("TaskStrategy").Item) > 0 Then
			aTaskComponent(S_STRATEGY_TASK) = oRequest("TaskStrategy").Item
		Else
			aTaskComponent(S_STRATEGY_TASK) = ""
		End If
	End If
	aTaskComponent(S_STRATEGY_TASK) = Left(aTaskComponent(S_STRATEGY_TASK), 4000)

	If IsEmpty(aTaskComponent(S_PURPOUSE_TASK)) Then
		If Len(oRequest("TaskPurpouse").Item) > 0 Then
			aTaskComponent(S_PURPOUSE_TASK) = oRequest("TaskPurpouse").Item
		Else
			aTaskComponent(S_PURPOUSE_TASK) = ""
		End If
	End If
	aTaskComponent(S_PURPOUSE_TASK) = Left(aTaskComponent(S_PURPOUSE_TASK), 4000)

	If IsEmpty(aTaskComponent(S_INDICATOR_TASK)) Then
		If Len(oRequest("TaskIndicator").Item) > 0 Then
			aTaskComponent(S_INDICATOR_TASK) = oRequest("TaskIndicator").Item
		Else
			aTaskComponent(S_INDICATOR_TASK) = ""
		End If
	End If
	aTaskComponent(S_INDICATOR_TASK) = Left(aTaskComponent(S_INDICATOR_TASK), 255)

	If IsEmpty(aTaskComponent(S_MEASUREMENT_TASK)) Then
		If Len(oRequest("TaskMeasurement").Item) > 0 Then
			aTaskComponent(S_MEASUREMENT_TASK) = oRequest("TaskMeasurement").Item
		Else
			aTaskComponent(S_MEASUREMENT_TASK) = ""
		End If
	End If
	aTaskComponent(S_MEASUREMENT_TASK) = Left(aTaskComponent(S_MEASUREMENT_TASK), 4000)

	If IsEmpty(aTaskComponent(S_FORMULA_TASK)) Then
		If Len(oRequest("TaskFormula").Item) > 0 Then
			aTaskComponent(S_FORMULA_TASK) = oRequest("TaskFormula").Item
		Else
			aTaskComponent(S_FORMULA_TASK) = ""
		End If
	End If
	aTaskComponent(S_FORMULA_TASK) = Left(aTaskComponent(S_FORMULA_TASK), 4000)

	If IsEmpty(aTaskComponent(N_AGGREGATION_TYPE_TASK)) Then
		If Len(oRequest("AggregationTypeID").Item) > 0 Then
			aTaskComponent(N_AGGREGATION_TYPE_TASK) = CLng(oRequest("AggregationTypeID").Item)
		Else
			aTaskComponent(N_AGGREGATION_TYPE_TASK) = 0
		End If
	End If

	If IsEmpty(aTaskComponent(S_COMMENTS_TASK)) Then
		If Len(oRequest("TaskComments").Item) > 0 Then
			aTaskComponent(S_COMMENTS_TASK) = oRequest("TaskComments").Item
		Else
			aTaskComponent(S_COMMENTS_TASK) = ""
		End If
	End If
	aTaskComponent(S_COMMENTS_TASK) = Left(aTaskComponent(S_COMMENTS_TASK), 4000)

	If IsEmpty(aTaskComponent(N_START_DATE_TASK)) Then
		If Len(oRequest("StartDate").Item) > 0 Then
			aTaskComponent(N_START_DATE_TASK) = CLng(oRequest("StartDate").Item)
		ElseIf Len(oRequest("StartYear").Item) > 0 Then
			aTaskComponent(N_START_DATE_TASK) = CLng((oRequest("StartYear").Item) & Right(("0" & CLng(oRequest("StartMonth").Item)), Len("00")) & Right(("0" & CLng(oRequest("StartDay").Item)), Len("00")))
		Else
			aTaskComponent(N_START_DATE_TASK) = 0
		End If
	End If

	If IsEmpty(aTaskComponent(N_END_DATE_TASK)) Then
		If Len(oRequest("EndDate").Item) > 0 Then
			aTaskComponent(N_END_DATE_TASK) = CLng(oRequest("EndDate").Item)
		ElseIf Len(oRequest("EndYear").Item) > 0 Then
			aTaskComponent(N_END_DATE_TASK) = CLng((oRequest("EndYear").Item) & Right(("0" & CLng(oRequest("EndMonth").Item)), Len("00")) & Right(("0" & CLng(oRequest("EndDay").Item)), Len("00")))
		Else
			aTaskComponent(N_END_DATE_TASK) = 0
		End If
	End If

	If IsEmpty(aTaskComponent(S_FILE_TASK)) Then
		If Len(oRequest("TaskFile").Item) > 0 Then
			aTaskComponent(S_FILE_TASK) = oRequest("TaskFile").Item
		Else
			aTaskComponent(S_FILE_TASK) = ""
		End If
	End If
	aTaskComponent(S_FILE_TASK) = Left(aTaskComponent(S_FILE_TASK), 255)

	If IsEmpty(aTaskComponent(N_PUCO_SECTION_ID_TASK)) Then
		If Len(oRequest("PuCoSectionID").Item) > 0 Then
			aTaskComponent(N_PUCO_SECTION_ID_TASK) = CLng(oRequest("PuCoSectionID").Item)
		Else
			aTaskComponent(N_PUCO_SECTION_ID_TASK) = -1
		End If
	End If

	If IsEmpty(aTaskComponent(N_FORM_TASK)) Then
		If Len(oRequest("FormID").Item) > 0 Then
			aTaskComponent(N_FORM_TASK) = CLng(oRequest("FormID").Item)
		Else
			aTaskComponent(N_FORM_TASK) = -1
		End If
	End If

	If IsEmpty(aTaskComponent(S_REPORT_URL_TASK)) Then
		If Len(oRequest("ReportURL").Item) > 0 Then
			aTaskComponent(S_REPORT_URL_TASK) = oRequest("ReportURL").Item
		Else
			aTaskComponent(S_REPORT_URL_TASK) = ""
		End If
	End If
	aTaskComponent(S_REPORT_URL_TASK) = Left(aTaskComponent(S_REPORT_URL_TASK), 255)

	If IsEmpty(aTaskComponent(AN_PARENT_ID_TASK)) Then
		aTaskComponent(AN_PARENT_ID_TASK) = ""
		If Len(oRequest("ParentID").Item) > 0 Then
			For Each oItem In oRequest("ParentID")
				aTaskComponent(AN_PARENT_ID_TASK) = aTaskComponent(AN_PARENT_ID_TASK) & oItem & ";"
			Next
			aTaskComponent(AN_PARENT_ID_TASK) = Left(aTaskComponent(AN_PARENT_ID_TASK), (Len(aTaskComponent(AN_PARENT_ID_TASK)) - Len(";")))
		ElseIf Len(oRequest("ParentIDs").Item) > 0 Then
			aTaskComponent(AN_PARENT_ID_TASK) = Replace(oRequest("ParentIDs").Item, " ", "", 1, -1, vbBinaryCompare)
		End If
	End If
	aTaskComponent(AN_PARENT_ID_TASK) = Split(aTaskComponent(AN_PARENT_ID_TASK), ";", -1, vbBinaryCompare)

	If IsEmpty(aTaskComponent(AD_MINIMUM_VALUE_TASK)) Then
		aTaskComponent(AD_MINIMUM_VALUE_TASK) = ""
		If Len(oRequest("TaskMinimumValue").Item) > 0 Then
			For Each oItem In oRequest("TaskMinimumValue")
				aTaskComponent(AD_MINIMUM_VALUE_TASK) = aTaskComponent(AD_MINIMUM_VALUE_TASK) & oItem & ";"
			Next
			aTaskComponent(AD_MINIMUM_VALUE_TASK) = Left(aTaskComponent(AD_MINIMUM_VALUE_TASK), (Len(aTaskComponent(AD_MINIMUM_VALUE_TASK)) - Len(";")))
		ElseIf Len(oRequest("TaskMinimumValues").Item) > 0 Then
			aTaskComponent(AD_MINIMUM_VALUE_TASK) = Replace(oRequest("TaskMinimumValues").Item, " ", "", 1, -1, vbBinaryCompare)
		Else
			aTaskComponent(AD_MINIMUM_VALUE_TASK) = 0
		End If
	End If
	aTaskComponent(AD_MINIMUM_VALUE_TASK) = Split(aTaskComponent(AD_MINIMUM_VALUE_TASK), ";", -1, vbBinaryCompare)
	If UBound(aTaskComponent(AD_MINIMUM_VALUE_TASK)) < UBound(aTaskComponent(AN_PARENT_ID_TASK)) Then aTaskComponent(AD_MINIMUM_VALUE_TASK) = Split(JoinLists(aTaskComponent(AD_MINIMUM_VALUE_TASK), BuildList("0", ",", UBound(aTaskComponent(AN_PARENT_ID_TASK)) + 1), ","), ",", -1, vbBinaryCompare)

	If IsEmpty(aTaskComponent(AD_AVERAGE_VALUE_TASK)) Then
		aTaskComponent(AD_AVERAGE_VALUE_TASK) = ""
		If Len(oRequest("TaskAverageValue").Item) > 0 Then
			For Each oItem In oRequest("TaskAverageValue")
				aTaskComponent(AD_AVERAGE_VALUE_TASK) = aTaskComponent(AD_AVERAGE_VALUE_TASK) & oItem & ";"
			Next
			aTaskComponent(AD_AVERAGE_VALUE_TASK) = Left(aTaskComponent(AD_AVERAGE_VALUE_TASK), (Len(aTaskComponent(AD_AVERAGE_VALUE_TASK)) - Len(";")))
		ElseIf Len(oRequest("TaskAverageValues").Item) > 0 Then
			aTaskComponent(AD_AVERAGE_VALUE_TASK) = Replace(oRequest("TaskAverageValues").Item, " ", "", 1, -1, vbBinaryCompare)
		Else
			aTaskComponent(AD_AVERAGE_VALUE_TASK) = 50
		End If
	End If
	aTaskComponent(AD_AVERAGE_VALUE_TASK) = Split(aTaskComponent(AD_AVERAGE_VALUE_TASK), ";", -1, vbBinaryCompare)
	If UBound(aTaskComponent(AD_AVERAGE_VALUE_TASK)) < UBound(aTaskComponent(AN_PARENT_ID_TASK)) Then aTaskComponent(AD_AVERAGE_VALUE_TASK) = Split(JoinLists(aTaskComponent(AD_AVERAGE_VALUE_TASK), BuildList("50", ",", UBound(aTaskComponent(AN_PARENT_ID_TASK)) + 1), ","), ",", -1, vbBinaryCompare)

	If IsEmpty(aTaskComponent(AD_MAXIMUM_VALUE_TASK)) Then
		aTaskComponent(AD_MAXIMUM_VALUE_TASK) = ""
		If Len(oRequest("TaskMaximumValue").Item) > 0 Then
			For Each oItem In oRequest("TaskMaximumValue")
				aTaskComponent(AD_MAXIMUM_VALUE_TASK) = aTaskComponent(AD_MAXIMUM_VALUE_TASK) & oItem & ";"
			Next
			aTaskComponent(AD_MAXIMUM_VALUE_TASK) = Left(aTaskComponent(AD_MAXIMUM_VALUE_TASK), (Len(aTaskComponent(AD_MAXIMUM_VALUE_TASK)) - Len(";")))
		ElseIf Len(oRequest("TaskMaximumValues").Item) > 0 Then
			aTaskComponent(AD_MAXIMUM_VALUE_TASK) = Replace(oRequest("TaskMaximumValues").Item, " ", "", 1, -1, vbBinaryCompare)
		Else
			aTaskComponent(AD_MAXIMUM_VALUE_TASK) = 100
		End If
	End If
	aTaskComponent(AD_MAXIMUM_VALUE_TASK) = Split(aTaskComponent(AD_MAXIMUM_VALUE_TASK), ";", -1, vbBinaryCompare)
	If UBound(aTaskComponent(AD_MAXIMUM_VALUE_TASK)) < UBound(aTaskComponent(AN_PARENT_ID_TASK)) Then aTaskComponent(AD_MAXIMUM_VALUE_TASK) = Split(JoinLists(aTaskComponent(AD_MAXIMUM_VALUE_TASK), BuildList("100", ",", UBound(aTaskComponent(AN_PARENT_ID_TASK)) + 1), ","), ",", -1, vbBinaryCompare)

	If IsEmpty(aTaskComponent(AD_TARGET_VALUE_TASK)) Then
		aTaskComponent(AD_TARGET_VALUE_TASK) = ""
		If Len(oRequest("TaskTargetValue").Item) > 0 Then
			For Each oItem In oRequest("TaskTargetValue")
				aTaskComponent(AD_TARGET_VALUE_TASK) = aTaskComponent(AD_TARGET_VALUE_TASK) & oItem & ";"
			Next
			aTaskComponent(AD_TARGET_VALUE_TASK) = Left(aTaskComponent(AD_TARGET_VALUE_TASK), (Len(aTaskComponent(AD_TARGET_VALUE_TASK)) - Len(";")))
		ElseIf Len(oRequest("TaskTargetValues").Item) > 0 Then
			aTaskComponent(AD_TARGET_VALUE_TASK) = Replace(oRequest("TaskTargetValues").Item, " ", "", 1, -1, vbBinaryCompare)
		Else
			aTaskComponent(AD_TARGET_VALUE_TASK) = 80
		End If
	End If
	aTaskComponent(AD_TARGET_VALUE_TASK) = Split(aTaskComponent(AD_TARGET_VALUE_TASK), ";", -1, vbBinaryCompare)
	If UBound(aTaskComponent(AD_TARGET_VALUE_TASK)) < UBound(aTaskComponent(AN_PARENT_ID_TASK)) Then aTaskComponent(AD_TARGET_VALUE_TASK) = Split(JoinLists(aTaskComponent(AD_TARGET_VALUE_TASK), BuildList("80", ",", UBound(aTaskComponent(AN_PARENT_ID_TASK)) + 1), ","), ",", -1, vbBinaryCompare)

	If IsEmpty(aTaskComponent(AN_FIELD_TYPE_TASK)) Then
		aTaskComponent(AN_FIELD_TYPE_TASK) = ""
		If Len(oRequest("TaskFieldTypeID").Item) > 0 Then
			For Each oItem In oRequest("TaskFieldTypeID")
				aTaskComponent(AN_FIELD_TYPE_TASK) = aTaskComponent(AN_FIELD_TYPE_TASK) & oItem & ";"
			Next
			aTaskComponent(AN_FIELD_TYPE_TASK) = Left(aTaskComponent(AN_FIELD_TYPE_TASK), (Len(aTaskComponent(AN_FIELD_TYPE_TASK)) - Len(";")))
		ElseIf Len(oRequest("TaskFieldTypeIDs").Item) > 0 Then
			aTaskComponent(AN_FIELD_TYPE_TASK) = Replace(oRequest("TaskFieldTypeIDs").Item, " ", "", 1, -1, vbBinaryCompare)
		Else
			aTaskComponent(AN_FIELD_TYPE_TASK) = 2
		End If
	End If
	aTaskComponent(AN_FIELD_TYPE_TASK) = Split(aTaskComponent(AN_FIELD_TYPE_TASK), ";", -1, vbBinaryCompare)
	If UBound(aTaskComponent(AN_FIELD_TYPE_TASK)) < UBound(aTaskComponent(AN_PARENT_ID_TASK)) Then aTaskComponent(AN_FIELD_TYPE_TASK) = Split(JoinLists(aTaskComponent(AN_FIELD_TYPE_TASK), BuildList("50", ",", UBound(aTaskComponent(AN_PARENT_ID_TASK)) + 1), ","), ",", -1, vbBinaryCompare)

	If IsEmpty(aTaskComponent(AD_PERCENTAGE_TASK)) Then
		aTaskComponent(AD_PERCENTAGE_TASK) = ""
		If Len(oRequest("TaskPercentage").Item) > 0 Then
			For Each oItem In oRequest("TaskPercentage")
				aTaskComponent(AD_PERCENTAGE_TASK) = aTaskComponent(AD_PERCENTAGE_TASK) & (CDbl(oItem) / 100) & ";"
			Next
			aTaskComponent(AD_PERCENTAGE_TASK) = Left(aTaskComponent(AD_PERCENTAGE_TASK), (Len(aTaskComponent(AD_PERCENTAGE_TASK)) - Len(";")))
		ElseIf Len(oRequest("TaskPercentages").Item) > 0 Then
			aTaskComponent(AD_PERCENTAGE_TASK) = Replace(oRequest("TaskPercentages").Item, " ", "", 1, -1, vbBinaryCompare)
		Else
			aTaskComponent(AD_PERCENTAGE_TASK) = 0
		End If
	End If
	aTaskComponent(AD_PERCENTAGE_TASK) = Split(aTaskComponent(AD_PERCENTAGE_TASK), ";", -1, vbBinaryCompare)
	If UBound(aTaskComponent(AD_PERCENTAGE_TASK)) < UBound(aTaskComponent(AN_PARENT_ID_TASK)) Then aTaskComponent(AD_PERCENTAGE_TASK) = Split(JoinLists(aTaskComponent(AD_PERCENTAGE_TASK), BuildList("50", ",", UBound(aTaskComponent(AN_PARENT_ID_TASK)) + 1), ","), ",", -1, vbBinaryCompare)

	If IsEmpty(aTaskComponent(AD_REQUIRED_TASK)) Then
		aTaskComponent(AD_REQUIRED_TASK) = ""
		If Len(oRequest("TaskRequired").Item) > 0 Then
			For Each oItem In oRequest("TaskRequired")
				aTaskComponent(AD_REQUIRED_TASK) = aTaskComponent(AD_REQUIRED_TASK) & oItem & ";"
			Next
			aTaskComponent(AD_REQUIRED_TASK) = Left(aTaskComponent(AD_REQUIRED_TASK), (Len(aTaskComponent(AD_REQUIRED_TASK)) - Len(";")))
		ElseIf Len(oRequest("TaskRequireds").Item) > 0 Then
			aTaskComponent(AD_REQUIRED_TASK) = Replace(oRequest("TaskRequireds").Item, " ", "", 1, -1, vbBinaryCompare)
		Else
			aTaskComponent(AD_REQUIRED_TASK) = 1
		End If
	End If
	aTaskComponent(AD_REQUIRED_TASK) = Split(aTaskComponent(AD_REQUIRED_TASK), ";", -1, vbBinaryCompare)
	If UBound(aTaskComponent(AD_REQUIRED_TASK)) < UBound(aTaskComponent(AN_PARENT_ID_TASK)) Then aTaskComponent(AD_REQUIRED_TASK) = Split(JoinLists(aTaskComponent(AD_REQUIRED_TASK), BuildList("50", ",", UBound(aTaskComponent(AN_PARENT_ID_TASK)) + 1), ","), ",", -1, vbBinaryCompare)

	If IsEmpty(aTaskComponent(AL_VARIABLES_TASK)) Then
		aTaskComponent(AL_VARIABLES_TASK) = ""
		For Each oItem In oRequest("VariableIDs")
			aTaskComponent(AL_VARIABLES_TASK) = aTaskComponent(AL_VARIABLES_TASK) & oItem & ";"
		Next
		If Len(aTaskComponent(AL_VARIABLES_TASK)) > 0 Then aTaskComponent(AL_VARIABLES_TASK) = Left(aTaskComponent(AL_VARIABLES_TASK), (Len(aTaskComponent(AL_VARIABLES_TASK)) - Len(";")))
	End If
	aTaskComponent(AL_VARIABLES_TASK) = Split(aTaskComponent(AL_VARIABLES_TASK), ";", -1, vbBinaryCompare)

	If IsEmpty(aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK)) Then
		aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK) = ""
		For Each oItem In oRequest("VariableMinimumValues")
			aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK) = aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK) & oItem & ";"
		Next
		If Len(aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK)) > 0 Then aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK) = Left(aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK), (Len(aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK)) - Len(";")))
	End If
	aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK) = Split(aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK), ";", -1, vbBinaryCompare)
	If UBound(aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK)) < UBound(aTaskComponent(AL_VARIABLES_TASK)) Then aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK) = Split(JoinLists(aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK), BuildList("0", ",", UBound(aTaskComponent(AL_VARIABLES_TASK)) + 1), ","), ",", -1, vbBinaryCompare)

	If IsEmpty(aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK)) Then
		aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK) = ""
		For Each oItem In oRequest("VariableMinimumValues")
			aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK) = aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK) & oItem & ";"
		Next
		If Len(aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK)) > 0 Then aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK) = Left(aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK), (Len(aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK)) - Len(";")))
	End If
	aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK) = Split(aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK), ";", -1, vbBinaryCompare)
	If UBound(aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK)) < UBound(aTaskComponent(AL_VARIABLES_TASK)) Then aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK) = Split(JoinLists(aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK), BuildList("50", ",", UBound(aTaskComponent(AL_VARIABLES_TASK)) + 1), ","), ",", -1, vbBinaryCompare)

	If IsEmpty(aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK)) Then
		aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK) = ""
		For Each oItem In oRequest("VariableMinimumValues")
			aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK) = aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK) & oItem & ";"
		Next
		If Len(aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK)) > 0 Then aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK) = Left(aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK), (Len(aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK)) - Len(";")))
	End If
	aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK) = Split(aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK), ";", -1, vbBinaryCompare)
	If UBound(aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK)) < UBound(aTaskComponent(AL_VARIABLES_TASK)) Then aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK) = Split(JoinLists(aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK), BuildList("100", ",", UBound(aTaskComponent(AL_VARIABLES_TASK)) + 1), ","), ",", -1, vbBinaryCompare)

	If IsEmpty(aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK)) Then
		aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK) = ""
		For Each oItem In oRequest("VariableMinimumValues")
			aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK) = aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK) & oItem & ";"
		Next
		If Len(aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK)) > 0 Then aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK) = Left(aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK), (Len(aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK)) - Len(";")))
	End If
	aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK) = Split(aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK), ";", -1, vbBinaryCompare)
	If UBound(aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK)) < UBound(aTaskComponent(AL_VARIABLES_TASK)) Then aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK) = Split(JoinLists(aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK), BuildList("80", ",", UBound(aTaskComponent(AL_VARIABLES_TASK)) + 1), ","), ",", -1, vbBinaryCompare)

	If IsEmpty(aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK)) Then
		aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK) = ""
		For Each oItem In oRequest("VariableMinimumValues")
			aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK) = aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK) & oItem & ";"
		Next
		If Len(aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK)) > 0 Then aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK) = Left(aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK), (Len(aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK)) - Len(";")))
	End If
	aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK) = Split(aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK), ";", -1, vbBinaryCompare)
	If UBound(aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK)) < UBound(aTaskComponent(AL_VARIABLES_TASK)) Then aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK) = Split(JoinLists(aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK), BuildList("2", ",", UBound(aTaskComponent(AL_VARIABLES_TASK)) + 1), ","), ",", -1, vbBinaryCompare)

	If IsEmpty(aTaskComponent(AN_VARIABLES_RELEVANCE_TASK)) Then
		aTaskComponent(AN_VARIABLES_RELEVANCE_TASK) = ""
		For Each oItem In oRequest("VariableRelevances")
			aTaskComponent(AN_VARIABLES_RELEVANCE_TASK) = aTaskComponent(AN_VARIABLES_RELEVANCE_TASK) & oItem & ";"
		Next
		If Len(aTaskComponent(AN_VARIABLES_RELEVANCE_TASK)) > 0 Then aTaskComponent(AN_VARIABLES_RELEVANCE_TASK) = Left(aTaskComponent(AN_VARIABLES_RELEVANCE_TASK), (Len(aTaskComponent(AN_VARIABLES_RELEVANCE_TASK)) - Len(";")))
	End If
	aTaskComponent(AN_VARIABLES_RELEVANCE_TASK) = Split(aTaskComponent(AN_VARIABLES_RELEVANCE_TASK), ";", -1, vbBinaryCompare)
	If UBound(aTaskComponent(AN_VARIABLES_RELEVANCE_TASK)) < UBound(aTaskComponent(AL_VARIABLES_TASK)) Then aTaskComponent(AN_VARIABLES_RELEVANCE_TASK) = Split(JoinLists(aTaskComponent(AN_VARIABLES_RELEVANCE_TASK), BuildList("1", ",", UBound(aTaskComponent(AL_VARIABLES_TASK)) + 1), ","), ",", -1, vbBinaryCompare)

	If IsEmpty(aTaskComponent(N_START_DATE_STATUS_TASK)) Then
		If Len(oRequest("TaskStatusStartDate").Item) > 0 Then
			aTaskComponent(N_START_DATE_STATUS_TASK) = CLng(oRequest("TaskStatusStartDate").Item)
		ElseIf Len(oRequest("TaskStatusStartYear").Item) > 0 Then
			aTaskComponent(N_START_DATE_STATUS_TASK) = Clng(oRequest("TaskStatusStartYear").Item & Right(("0" & oRequest("TaskStatusStartMonth").Item), Len("00")) & Right(("0" & oRequest("TaskStatusStartDay").Item), Len("00")))
		Else
			aTaskComponent(N_START_DATE_STATUS_TASK) = 0
		End If
	End If

	If IsEmpty(aTaskComponent(N_END_DATE_STATUS_TASK)) Then
		If Len(oRequest("TaskStatusEndDate").Item) > 0 Then
			aTaskComponent(N_END_DATE_STATUS_TASK) = CLng(oRequest("TaskStatusEndDate").Item)
		ElseIf Len(oRequest("TaskStatusEndYear").Item) > 0 Then
			aTaskComponent(N_END_DATE_STATUS_TASK) = Clng(oRequest("TaskStatusEndYear").Item & Right(("0" & oRequest("TaskStatusEndMonth").Item), Len("00")) & Right(("0" & oRequest("TaskStatusEndDay").Item), Len("00")))
		Else
			aTaskComponent(N_END_DATE_STATUS_TASK) = 0
		End If
	End If

	If IsEmpty(aTaskComponent(D_VALUE_STATUS_TASK)) Then
		If Len(oRequest("TaskStatusValue").Item) > 0 Then
			aTaskComponent(D_VALUE_STATUS_TASK) = CDbl(oRequest("TaskStatusValue").Item)
		Else
			aTaskComponent(D_VALUE_STATUS_TASK) = 0
			For iIndex = 0 To UBound(aTaskComponent(AN_PARENT_ID_TASK))
				If CLng(aTaskComponent(AN_PARENT_ID_TASK)(iIndex)) = aTaskComponent(N_PARENT_ID_TASK) Then
					aTaskComponent(D_VALUE_STATUS_TASK) = aTaskComponent(AD_MINIMUM_VALUE_TASK)(iIndex)
					Exit For
				End If
			Next
		End If
	End If

	If IsEmpty(aTaskComponent(D_PERCENTAGE_STATUS_TASK)) Then
		If Len(oRequest("TaskStatusPercentage").Item) > 0 Then
			aTaskComponent(D_PERCENTAGE_STATUS_TASK) = CDbl(oRequest("TaskStatusPercentage").Item) / 100
		Else
			aTaskComponent(D_PERCENTAGE_STATUS_TASK) = 0
		End If
	End If

	If IsEmpty(aTaskComponent(L_STATUS_ID_TASK)) Then
		If aTaskComponent(D_PERCENTAGE_STATUS_TASK) = 1 Then
			aTaskComponent(L_STATUS_ID_TASK) = 1
		ElseIf Len(oRequest("TaskStatusID").Item) > 0 Then
			aTaskComponent(L_STATUS_ID_TASK) = CLng(oRequest("TaskStatusID").Item)
		Else
			aTaskComponent(L_STATUS_ID_TASK) = 0
		End If
	End If

	If IsEmpty(aTaskComponent(S_AREAS_TASK)) Then
		aTaskComponent(S_AREAS_TASK) = ""
		If Len(oRequest("AreaID").Item) > 0 Then
			For Each oItem In oRequest("AreaID")
				aTaskComponent(S_AREAS_TASK) = aTaskComponent(S_AREAS_TASK) & oItem & ","
			Next
			aTaskComponent(S_AREAS_TASK) = Left(aTaskComponent(S_AREAS_TASK), (Len(aTaskComponent(S_AREAS_TASK)) - Len(",")))
		ElseIf Len(oRequest("AreasID").Item) > 0 Then
			aTaskComponent(S_AREAS_TASK) = Replace(oRequest("AreasID").Item, " ", "", 1, -1, vbBinaryCompare)
		End If
	End If

	If IsEmpty(aTaskComponent(S_USERS_TASK)) Then
		aTaskComponent(S_USERS_TASK) = ""
		If Len(oRequest("UserID").Item) > 0 Then
			For Each oItem In oRequest("UserID")
				aTaskComponent(S_USERS_TASK) = aTaskComponent(S_USERS_TASK) & oItem & ","
			Next
			aTaskComponent(S_USERS_TASK) = Left(aTaskComponent(S_USERS_TASK), (Len(aTaskComponent(S_USERS_TASK)) - Len(",")))
		ElseIf Len(oRequest("UsersID").Item) > 0 Then
			aTaskComponent(S_USERS_TASK) = Replace(oRequest("UsersID").Item, " ", "", 1, -1, vbBinaryCompare)
		End If
	End If

	If IsEmpty(aTaskComponent(S_CATEGORIES_TASK)) Then
		aTaskComponent(S_CATEGORIES_TASK) = ""
		If Len(oRequest("CategoryID").Item) > 0 Then
			For Each oItem In oRequest("CategoryID")
				aTaskComponent(S_CATEGORIES_TASK) = aTaskComponent(S_CATEGORIES_TASK) & oItem & ","
			Next
			aTaskComponent(S_CATEGORIES_TASK) = Left(aTaskComponent(S_CATEGORIES_TASK), (Len(aTaskComponent(S_CATEGORIES_TASK)) - Len(",")))
		ElseIf Len(oRequest("CategoriesID").Item) > 0 Then
			aTaskComponent(S_CATEGORIES_TASK) = Replace(oRequest("CategoriesID").Item, " ", "", 1, -1, vbBinaryCompare)
		End If
	End If

	aTaskComponent(B_HAS_CHILDREN_TASK) = False
	aTaskComponent(N_EASY_TASK) = 0
	aTaskComponent(S_QUERY_CONDITION_TASK) = ""
	aTaskComponent(B_CHECK_FOR_DUPLICATED_TASK) = True
	aTaskComponent(B_IS_DUPLICATED_TASK) = False

	aTaskComponent(B_COMPONENT_INITIALIZED_TASK) = True
	InitializeTaskComponent = Err.number
	Err.Clear
End Function

Function AddTask(oRequest, oADODBConnection, aTaskComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new user into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aTaskComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddTask"
	Dim iIndex
	Dim asTaskLKP
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aTaskComponent(B_COMPONENT_INITIALIZED_TASK)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeTaskComponent(oRequest, aTaskComponent)
	End If

	If aTaskComponent(N_ID_TASK) = -1 Then
		sErrorDescription = "No se pudo obtener un identificador para la nueva actividad."
		lErrorNumber = GetNewIDFromTable(oADODBConnection, TACO_PREFIX & "Tasks", "TaskID", "(ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ")", 1, aTaskComponent(N_ID_TASK), sErrorDescription)
	End If

	If lErrorNumber = 0 Then
		If aTaskComponent(B_CHECK_FOR_DUPLICATED_TASK) Then
			lErrorNumber = CheckExistencyOfTask(oADODBConnection, False, aTaskComponent, sErrorDescription)
			If aTaskComponent(B_IS_DUPLICATED_TASK) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "La clave " & aTaskComponent(S_NUMBER_TASK) & " ya está registrada en el sistema y fue asignado a la actividad " & aTaskComponent(S_NAME_TASK) & "."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "TaCoTaskComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			End If
		End If

		If lErrorNumber = 0 Then
			If Not CheckTaskInformationConsistency(True, aTaskComponent, sErrorDescription) Then
				lErrorNumber = -1
			Else
				sErrorDescription = "No se pudo guardar la información de la nueva actividad."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & TACO_PREFIX & "Tasks (ProjectID, TaskID, TaskName, TaskNumber, LabelID, TaskDescription, TaskObjective, TaskStrategy, TaskPurpouse, TaskIndicator, TaskMeasurement, TaskFormula, AggregationTypeID, TaskComments, StartDate, EndDate, TaskFile, PuCoSectionID, FormID, ReportURL) Values (" & aTaskComponent(N_PROJECT_ID_TASK) & ", " & aTaskComponent(N_ID_TASK) & ", '" & Replace(aTaskComponent(S_NAME_TASK), "'", "´")  & "', '" & Replace(aTaskComponent(S_NUMBER_TASK), "'", "´")  & "', " & aTaskComponent(N_LABEL_ID_TASK) & ", '" & Replace(aTaskComponent(S_DESCRIPTION_TASK), "'", "´")  & "', '" & Replace(aTaskComponent(S_OBJECTIVE_TASK), "'", "´")  & "', '" & Replace(aTaskComponent(S_STRATEGY_TASK), "'", "´")  & "', '" & Replace(aTaskComponent(S_PURPOUSE_TASK), "'", "´")  & "', '" & Replace(aTaskComponent(S_INDICATOR_TASK), "'", "´") & "', '" & Replace(aTaskComponent(S_MEASUREMENT_TASK), "'", "´")  & "', '" & Replace(aTaskComponent(S_FORMULA_TASK), "'", "´")& "', " & aTaskComponent(N_AGGREGATION_TYPE_TASK) & ", '" & Replace(aTaskComponent(S_COMMENTS_TASK), "'", "´")  & "', " & aTaskComponent(N_START_DATE_TASK) & ", " & aTaskComponent(N_END_DATE_TASK) & ", '" & Replace(aTaskComponent(S_FILE_TASK), "'", "")  & "', " & aTaskComponent(N_PUCO_SECTION_ID_TASK) & ", " & aTaskComponent(N_FORM_TASK) & ", '" & Replace(aTaskComponent(S_REPORT_URL_TASK), "'", "") & "')", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				If lErrorNumber = 0 Then
					For iIndex = 0 To UBound(aTaskComponent(AN_PARENT_ID_TASK))
						sErrorDescription = "No se pudo guardar la información de la nueva actividad."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & TACO_PREFIX & "TasksLKP (ProjectID, TaskID, ParentID, MinimumValue, AverageValue, MaximumValue, TargetValue, FieldTypeID, TaskPercentage, TaskRequired) Values (" & aTaskComponent(N_PROJECT_ID_TASK) & ", " & aTaskComponent(N_ID_TASK) & ", " & aTaskComponent(AN_PARENT_ID_TASK)(iIndex) & ", " & aTaskComponent(AD_MINIMUM_VALUE_TASK)(iIndex) & ", " & aTaskComponent(AD_AVERAGE_VALUE_TASK)(iIndex) & ", " & aTaskComponent(AD_MAXIMUM_VALUE_TASK)(iIndex) & ", " & aTaskComponent(AD_TARGET_VALUE_TASK)(iIndex) & ", " & aTaskComponent(AN_FIELD_TYPE_TASK)(iIndex) & ", " & aTaskComponent(AD_PERCENTAGE_TASK)(iIndex) & ", " & aTaskComponent(AD_REQUIRED_TASK)(iIndex) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						If lErrorNumber = 0 Then
							sErrorDescription = "No se pudo guardar la información de la nueva actividad."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & TACO_PREFIX & "TasksStatusLKP (ProjectID, TaskID, ParentID, StatusDate, StartDate, EndDate, TaskValue, TaskPercentage, TaskStatusID) Values (" & aTaskComponent(N_PROJECT_ID_TASK) & ", " & aTaskComponent(N_ID_TASK) & ", " & aTaskComponent(AN_PARENT_ID_TASK)(iIndex) & ", 0, 0, 0, " & aTaskComponent(D_VALUE_STATUS_TASK) & ", 0, 0)", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						End If
					Next
				End If

				If lErrorNumber = 0 Then
					sErrorDescription = "No se pudo guardar la información de la nueva actividad."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & TACO_PREFIX & "TasksVariablesLKP  Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					If lErrorNumber = 0 Then
						sErrorDescription = "No se pudo guardar la información de la nueva actividad."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & TACO_PREFIX & "TasksVariablesStatusLKP  Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ") And (StatusDate=0)", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
					If lErrorNumber = 0 Then
						For iIndex = 0 To UBound(aTaskComponent(AL_VARIABLES_TASK))
							sErrorDescription = "No se pudo guardar la información de la nueva actividad."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & TACO_PREFIX & "TasksVariablesLKP (ProjectID, TaskID, VariableID, MinimumValue, AverageValue, MaximumValue, TargetValue, FieldTypeID, Relevance) Values (" & aTaskComponent(N_PROJECT_ID_TASK) & ", " & aTaskComponent(N_ID_TASK) & ", " & aTaskComponent(AL_VARIABLES_TASK)(iIndex) & ", " & aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK)(iIndex) & ", " & aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK)(iIndex) & ", " & aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK)(iIndex) & ", " & aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK)(iIndex) & ", " & aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK)(iIndex) & ", " & aTaskComponent(AN_VARIABLES_RELEVANCE_TASK)(iIndex) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudo guardar la información de la nueva actividad."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & TACO_PREFIX & "TasksVariablesStatusLKP (ProjectID, TaskID, VariableID, StatusDate, TaskValue) Values (" & aTaskComponent(N_PROJECT_ID_TASK) & ", " & aTaskComponent(N_ID_TASK) & ", " & aTaskComponent(AL_VARIABLES_TASK)(iIndex) & ", 0, 0)", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
							End If
						Next
					End If
				End If

				If lErrorNumber = 0 Then
					asTaskLKP = Split(aTaskComponent(S_AREAS_TASK), ",", -1, vbBinaryCompare)
					For iIndex = 0 To UBound(asTaskLKP)
						sErrorDescription = "No se pudo guardar la información de la nueva actividad."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & TACO_PREFIX & "TaskAreasLKP (ProjectID, TaskID, AreaID) Values (" & aTaskComponent(N_PROJECT_ID_TASK) & ", " & aTaskComponent(N_ID_TASK) & ", " & asTaskLKP(iIndex) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Next
				End If
				If lErrorNumber = 0 Then
					asTaskLKP = Split(aTaskComponent(S_USERS_TASK), ",", -1, vbBinaryCompare)
					For iIndex = 0 To UBound(asTaskLKP)
						sErrorDescription = "No se pudo guardar la información de la nueva actividad."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & TACO_PREFIX & "TaskUsersLKP (ProjectID, TaskID, UserID) Values (" & aTaskComponent(N_PROJECT_ID_TASK) & ", " & aTaskComponent(N_ID_TASK) & ", " & asTaskLKP(iIndex) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Next
				End If
				If lErrorNumber = 0 Then
					asTaskLKP = Split(aTaskComponent(S_CATEGORIES_TASK), ",", -1, vbBinaryCompare)
					For iIndex = 0 To UBound(asTaskLKP)
						sErrorDescription = "No se pudo guardar la información de la nueva actividad."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & TACO_PREFIX & "TaskCategoriesLKP (ProjectID, TaskID, CategoryID) Values (" & aTaskComponent(N_PROJECT_ID_TASK) & ", " & aTaskComponent(N_ID_TASK) & ", " & asTaskLKP(iIndex) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Next
				End If
			End If
		End If
	End If

	AddTask = lErrorNumber
	Err.Clear
End Function

Function ImportTask(oRequest, oADODBConnection, aTaskComponent, sErrorDescription)
'************************************************************
'Purpose: To import the information from a task
'Inputs:  oRequest, oADODBConnection, 
'Outputs: aTaskComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ImportTask"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aTaskComponent(B_COMPONENT_INITIALIZED_TASK)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeTaskComponent(oRequest, aTaskComponent)
	End If

	If (aTaskComponent(N_ID_TASK) = -1) Or (aTaskComponent(N_PROJECT_ID_TASK) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó la actividad para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "TaCoTaskComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		lErrorNumber = GetTask(oRequest, oADODBConnection, aTaskComponent, sErrorDescription)
		If lErrorNumber = 0 Then
			aTaskComponent(N_ID_TASK) = -1
			lErrorNumber = AddTask(oRequest, oADODBConnection, aTaskComponent, sErrorDescription)
		End If
	End If

	ImportTask = lErrorNumber
	Err.Clear
End Function

Function GetTask(oRequest, oADODBConnection, aTaskComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a task from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aTaskComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetTask"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aTaskComponent(B_COMPONENT_INITIALIZED_TASK)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeTaskComponent(oRequest, aTaskComponent)
	End If

	If (aTaskComponent(N_ID_TASK) = -1) Or (aTaskComponent(N_PROJECT_ID_TASK) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó la actividad para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "TaCoTaskComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información de la actividad."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From " & TACO_PREFIX & "Tasks Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El actividad especificada no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "TaCoTaskComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aTaskComponent(N_ID_TASK) = CLng(oRecordset.Fields("TaskID").Value)
				aTaskComponent(N_PROJECT_ID_TASK) = CLng(oRecordset.Fields("ProjectID").Value)
				aTaskComponent(S_NAME_TASK) = CStr(oRecordset.Fields("TaskName").Value)
				aTaskComponent(S_NUMBER_TASK) = CStr(oRecordset.Fields("TaskNumber").Value)
				aTaskComponent(N_LABEL_ID_TASK) = CLng(oRecordset.Fields("LabelID").Value)
				aTaskComponent(S_DESCRIPTION_TASK) = CStr(oRecordset.Fields("TaskDescription").Value)
				aTaskComponent(S_OBJECTIVE_TASK) = CStr(oRecordset.Fields("TaskObjective").Value)
				aTaskComponent(S_STRATEGY_TASK) = CStr(oRecordset.Fields("TaskStrategy").Value)
				aTaskComponent(S_PURPOUSE_TASK) = CStr(oRecordset.Fields("TaskPurpouse").Value)
				aTaskComponent(S_INDICATOR_TASK) = CStr(oRecordset.Fields("TaskIndicator").Value)
				aTaskComponent(S_MEASUREMENT_TASK) = CStr(oRecordset.Fields("TaskMeasurement").Value)
				aTaskComponent(S_FORMULA_TASK) = CStr(oRecordset.Fields("TaskFormula").Value)
				aTaskComponent(N_AGGREGATION_TYPE_TASK) = CLng(oRecordset.Fields("AggregationTypeID").Value)
				aTaskComponent(S_COMMENTS_TASK) = CStr(oRecordset.Fields("TaskComments").Value)
				aTaskComponent(N_START_DATE_TASK) = CLng(oRecordset.Fields("StartDate").Value)
				aTaskComponent(N_END_DATE_TASK) = CLng(oRecordset.Fields("EndDate").Value)
				aTaskComponent(S_FILE_TASK) = CStr(oRecordset.Fields("TaskFile").Value)
				aTaskComponent(N_PUCO_SECTION_ID_TASK) = CLng(oRecordset.Fields("PuCoSectionID").Value)
				aTaskComponent(N_FORM_TASK) = CLng(oRecordset.Fields("FormID").Value)
				aTaskComponent(S_REPORT_URL_TASK) = CStr(oRecordset.Fields("ReportURL").Value)
				oRecordset.Close

'				Call GetNameFromTable(oADODBConnection, TACO_PREFIX & "ProjectFile", aTaskComponent(N_PROJECT_ID_TASK), "", "", aTaskComponent(S_PROJECT_FILE_TASK), "")

				sErrorDescription = "No se pudo obtener la información de la actividad."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From " & TACO_PREFIX & "TasksLKP Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ") Order By ParentID", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					aTaskComponent(AN_PARENT_ID_TASK) = ""
					aTaskComponent(AD_MINIMUM_VALUE_TASK) = ""
					aTaskComponent(AD_AVERAGE_VALUE_TASK) = ""
					aTaskComponent(AD_MAXIMUM_VALUE_TASK) = ""
					aTaskComponent(AD_TARGET_VALUE_TASK) = ""
					aTaskComponent(AN_FIELD_TYPE_TASK) = ""
					aTaskComponent(AD_PERCENTAGE_TASK) = ""
					aTaskComponent(AD_REQUIRED_TASK) = ""
					Do While Not oRecordset.EOF
						aTaskComponent(AN_PARENT_ID_TASK) = aTaskComponent(AN_PARENT_ID_TASK) & CStr(oRecordset.Fields("ParentID").Value) & ";"
						aTaskComponent(AD_MINIMUM_VALUE_TASK) = aTaskComponent(AD_MINIMUM_VALUE_TASK) & CStr(oRecordset.Fields("MinimumValue").Value) & ";"
						aTaskComponent(AD_AVERAGE_VALUE_TASK) = aTaskComponent(AD_AVERAGE_VALUE_TASK) & CStr(oRecordset.Fields("AverageValue").Value) & ";"
						aTaskComponent(AD_MAXIMUM_VALUE_TASK) = aTaskComponent(AD_MAXIMUM_VALUE_TASK) & CStr(oRecordset.Fields("MaximumValue").Value) & ";"
						aTaskComponent(AD_TARGET_VALUE_TASK) = aTaskComponent(AD_TARGET_VALUE_TASK) & CStr(oRecordset.Fields("TargetValue").Value) & ";"
						aTaskComponent(AN_FIELD_TYPE_TASK) = aTaskComponent(AN_FIELD_TYPE_TASK) & CStr(oRecordset.Fields("FieldTypeID").Value) & ";"
						aTaskComponent(AD_PERCENTAGE_TASK) = aTaskComponent(AD_PERCENTAGE_TASK) & CStr(oRecordset.Fields("TaskPercentage").Value) & ";"
						aTaskComponent(AD_REQUIRED_TASK) = aTaskComponent(AD_REQUIRED_TASK) & CStr(oRecordset.Fields("TaskRequired").Value) & ";"
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					If Len(aTaskComponent(AN_PARENT_ID_TASK)) > 0 Then
						aTaskComponent(AN_PARENT_ID_TASK) = Left(aTaskComponent(AN_PARENT_ID_TASK), (Len(aTaskComponent(AN_PARENT_ID_TASK)) - Len(";")))
					Else
						aTaskComponent(AN_PARENT_ID_TASK) = -1
					End If
					If Len(aTaskComponent(AD_MINIMUM_VALUE_TASK)) > 0 Then
						aTaskComponent(AD_MINIMUM_VALUE_TASK) = Left(aTaskComponent(AD_MINIMUM_VALUE_TASK), (Len(aTaskComponent(AD_MINIMUM_VALUE_TASK)) - Len(";")))
					Else
						aTaskComponent(AD_MINIMUM_VALUE_TASK) = 0
					End If
					If Len(aTaskComponent(AD_AVERAGE_VALUE_TASK)) > 0 Then
						aTaskComponent(AD_AVERAGE_VALUE_TASK) = Left(aTaskComponent(AD_AVERAGE_VALUE_TASK), (Len(aTaskComponent(AD_AVERAGE_VALUE_TASK)) - Len(";")))
					Else
						aTaskComponent(AD_AVERAGE_VALUE_TASK) = 50
					End If
					If Len(aTaskComponent(AD_MAXIMUM_VALUE_TASK)) > 0 Then
						aTaskComponent(AD_MAXIMUM_VALUE_TASK) = Left(aTaskComponent(AD_MAXIMUM_VALUE_TASK), (Len(aTaskComponent(AD_MAXIMUM_VALUE_TASK)) - Len(";")))
					Else
						aTaskComponent(AD_MAXIMUM_VALUE_TASK) = 100
					End If
					If Len(aTaskComponent(AD_TARGET_VALUE_TASK)) > 0 Then
						aTaskComponent(AD_TARGET_VALUE_TASK) = Left(aTaskComponent(AD_TARGET_VALUE_TASK), (Len(aTaskComponent(AD_TARGET_VALUE_TASK)) - Len(";")))
					Else
						aTaskComponent(AD_TARGET_VALUE_TASK) = 80
					End If
					If Len(aTaskComponent(AN_FIELD_TYPE_TASK)) > 0 Then
						aTaskComponent(AN_FIELD_TYPE_TASK) = Left(aTaskComponent(AN_FIELD_TYPE_TASK), (Len(aTaskComponent(AN_FIELD_TYPE_TASK)) - Len(";")))
					Else
						aTaskComponent(AN_FIELD_TYPE_TASK) = 2
					End If
					If Len(aTaskComponent(AD_PERCENTAGE_TASK)) > 0 Then
						aTaskComponent(AD_PERCENTAGE_TASK) = Left(aTaskComponent(AD_PERCENTAGE_TASK), (Len(aTaskComponent(AD_PERCENTAGE_TASK)) - Len(";")))
					Else
						aTaskComponent(AD_PERCENTAGE_TASK) = 0
					End If
					If Len(aTaskComponent(AD_REQUIRED_TASK)) > 0 Then
						aTaskComponent(AD_REQUIRED_TASK) = Left(aTaskComponent(AD_REQUIRED_TASK), (Len(aTaskComponent(AD_REQUIRED_TASK)) - Len(";")))
					Else
						aTaskComponent(AD_REQUIRED_TASK) = 1
					End If
					aTaskComponent(AN_PARENT_ID_TASK) = Split(aTaskComponent(AN_PARENT_ID_TASK), ";", -1, vbBinaryCompare)
					aTaskComponent(AD_MINIMUM_VALUE_TASK) = Split(aTaskComponent(AD_MINIMUM_VALUE_TASK), ";", -1, vbBinaryCompare)
					aTaskComponent(AD_AVERAGE_VALUE_TASK) = Split(aTaskComponent(AD_AVERAGE_VALUE_TASK), ";", -1, vbBinaryCompare)
					aTaskComponent(AD_MAXIMUM_VALUE_TASK) = Split(aTaskComponent(AD_MAXIMUM_VALUE_TASK), ";", -1, vbBinaryCompare)
					aTaskComponent(AD_TARGET_VALUE_TASK) = Split(aTaskComponent(AD_TARGET_VALUE_TASK), ";", -1, vbBinaryCompare)
					aTaskComponent(AN_FIELD_TYPE_TASK) = Split(aTaskComponent(AN_FIELD_TYPE_TASK), ";", -1, vbBinaryCompare)
					aTaskComponent(AD_PERCENTAGE_TASK) = Split(aTaskComponent(AD_PERCENTAGE_TASK), ";", -1, vbBinaryCompare)
					aTaskComponent(AD_REQUIRED_TASK) = Split(aTaskComponent(AD_REQUIRED_TASK), ";", -1, vbBinaryCompare)
					oRecordset.Close
				End If

				sErrorDescription = "No se pudo obtener la información de la actividad."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From " & TACO_PREFIX & "TasksStatusLKP Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ") And (ParentID=" & aTaskComponent(N_PARENT_ID_TASK) & ") Order By StatusDate Desc", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					aTaskComponent(N_START_DATE_STATUS_TASK) = CLng(oRecordset.Fields("StartDate").Value)
					aTaskComponent(N_END_DATE_STATUS_TASK) = CLng(oRecordset.Fields("EndDate").Value)
					aTaskComponent(D_VALUE_STATUS_TASK) = CDbl(oRecordset.Fields("TaskValue").Value)
					aTaskComponent(D_PERCENTAGE_STATUS_TASK) = CDbl(oRecordset.Fields("TaskPercentage").Value)
					aTaskComponent(L_STATUS_ID_TASK) = CLng(oRecordset.Fields("TaskStatusID").Value)
					oRecordset.Close
				End If

				sErrorDescription = "No se pudo obtener la información de la actividad."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select " & TACO_PREFIX & "TasksVariablesLKP.* From " & TACO_PREFIX & "TasksVariablesLKP, " & TACO_PREFIX & "Variables Where (" & TACO_PREFIX & "TasksVariablesLKP.VariableID=" & TACO_PREFIX & "Variables.VariableID) And (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ") Order By Relevance, VariableName", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					aTaskComponent(AL_VARIABLES_TASK) = ""
					aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK) = ""
					aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK) = ""
					aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK) = ""
					aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK) = ""
					aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK) = ""
					aTaskComponent(AN_VARIABLES_RELEVANCE_TASK) = ""
					Do While Not oRecordset.EOF
						aTaskComponent(AL_VARIABLES_TASK) = aTaskComponent(AL_VARIABLES_TASK) & CStr(oRecordset.Fields("VariableID").Value) & ";"
						aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK) = aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK) & CStr(oRecordset.Fields("MinimumValue").Value) & ";"
						aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK) = aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK) & CStr(oRecordset.Fields("AverageValue").Value) & ";"
						aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK) = aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK) & CStr(oRecordset.Fields("MaximumValue").Value) & ";"
						aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK) = aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK) & CStr(oRecordset.Fields("TargetValue").Value) & ";"
						aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK) = aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK) & CStr(oRecordset.Fields("FieldTypeID").Value) & ";"
						aTaskComponent(AN_VARIABLES_RELEVANCE_TASK) = aTaskComponent(AN_VARIABLES_RELEVANCE_TASK) & CStr(oRecordset.Fields("Relevance").Value) & ";"
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					If Len(aTaskComponent(AL_VARIABLES_TASK)) > 0 Then aTaskComponent(AL_VARIABLES_TASK) = Left(aTaskComponent(AL_VARIABLES_TASK), (Len(aTaskComponent(AL_VARIABLES_TASK)) - Len(";")))
					If Len(aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK)) > 0 Then aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK) = Left(aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK), (Len(aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK)) - Len(";")))
					If Len(aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK)) > 0 Then aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK) = Left(aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK), (Len(aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK)) - Len(";")))
					If Len(aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK)) > 0 Then aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK) = Left(aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK), (Len(aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK)) - Len(";")))
					If Len(aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK)) > 0 Then aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK) = Left(aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK), (Len(aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK)) - Len(";")))
					If Len(aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK)) > 0 Then aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK) = Left(aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK), (Len(aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK)) - Len(";")))
					If Len(aTaskComponent(AN_VARIABLES_RELEVANCE_TASK)) > 0 Then aTaskComponent(AN_VARIABLES_RELEVANCE_TASK) = Left(aTaskComponent(AN_VARIABLES_RELEVANCE_TASK), (Len(aTaskComponent(AN_VARIABLES_RELEVANCE_TASK)) - Len(";")))
					aTaskComponent(AL_VARIABLES_TASK) = Split(aTaskComponent(AL_VARIABLES_TASK), ";")
					aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK) = Split(aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK), ";")
					aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK) = Split(aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK), ";")
					aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK) = Split(aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK), ";")
					aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK) = Split(aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK), ";")
					aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK) = Split(aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK), ";")
					aTaskComponent(AN_VARIABLES_RELEVANCE_TASK) = Split(aTaskComponent(AN_VARIABLES_RELEVANCE_TASK), ";")
				End If

				sErrorDescription = "No se pudo obtener la información de la actividad."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaID From " & TACO_PREFIX & "TaskAreasLKP Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ") Order By AreaID", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					aTaskComponent(S_AREAS_TASK) = ""
					Do While Not oRecordset.EOF
						aTaskComponent(S_AREAS_TASK) = aTaskComponent(S_AREAS_TASK) & CStr(oRecordset.Fields("AreaID").Value) & ","
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
				If Len(aTaskComponent(S_AREAS_TASK)) > 0 Then aTaskComponent(S_AREAS_TASK) = Left(aTaskComponent(S_AREAS_TASK), (Len(aTaskComponent(S_AREAS_TASK)) - Len(",")))

				sErrorDescription = "No se pudo obtener la información de la actividad."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select UserID From " & TACO_PREFIX & "TaskUsersLKP Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ") Order By UserID", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					aTaskComponent(S_USERS_TASK) = ""
					Do While Not oRecordset.EOF
						aTaskComponent(S_USERS_TASK) = aTaskComponent(S_USERS_TASK) & CStr(oRecordset.Fields("UserID").Value) & ","
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
				If Len(aTaskComponent(S_USERS_TASK)) > 0 Then aTaskComponent(S_USERS_TASK) = Left(aTaskComponent(S_USERS_TASK), (Len(aTaskComponent(S_USERS_TASK)) - Len(",")))

				sErrorDescription = "No se pudo obtener la información de la actividad."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CategoryID From " & TACO_PREFIX & "TaskCategoriesLKP Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ") Order By CategoryID", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					aTaskComponent(S_CATEGORIES_TASK) = ""
					Do While Not oRecordset.EOF
						aTaskComponent(S_CATEGORIES_TASK) = aTaskComponent(S_CATEGORIES_TASK) & CStr(oRecordset.Fields("CategoryID").Value) & ","
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
				If Len(aTaskComponent(S_CATEGORIES_TASK)) > 0 Then aTaskComponent(S_CATEGORIES_TASK) = Left(aTaskComponent(S_CATEGORIES_TASK), (Len(aTaskComponent(S_CATEGORIES_TASK)) - Len(",")))

				sErrorDescription = "No se pudo obtener la información de la actividad."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select TaskID From " & TACO_PREFIX & "TasksLKP Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (ParentID=" & aTaskComponent(N_ID_TASK) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					aTaskComponent(B_HAS_CHILDREN_TASK) = (Not oRecordset.EOF)
					oRecordset.Close
				End If
'				Call GetNameFromTable(oADODBConnection, TACO_PREFIX & "ProjectEasy", aTaskComponent(N_PROJECT_ID_TASK), "", "", aTaskComponent(N_EASY_TASK), sErrorDescription)
			End If
		End If
	End If

	Set oRecordset = Nothing
	GetTask = lErrorNumber
	Err.Clear
End Function

Function GetTasks(oRequest, oADODBConnection, aTaskComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the tasks from
'		  the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aTaskComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetTasks"
	Dim sCondition
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aTaskComponent(B_COMPONENT_INITIALIZED_TASK)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeTaskComponent(oRequest, aTaskComponent)
	End If

	sErrorDescription = "No se pudo obtener la información de las actividades."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select " & TACO_PREFIX & "Tasks.TaskID, ParentID, TaskName, TaskNumber, TaskPercentage, LabelName From " & TACO_PREFIX & "Tasks, " & TACO_PREFIX & "TasksLKP, " & TACO_PREFIX & "Labels Where (" & TACO_PREFIX & "Tasks.ProjectID=" & TACO_PREFIX & "TasksLKP.ProjectID) And (" & TACO_PREFIX & "Tasks.TaskID=" & TACO_PREFIX & "TasksLKP.TaskID) And (" & TACO_PREFIX & "Tasks.LabelID=" & TACO_PREFIX & "Labels.LabelID) And (" & TACO_PREFIX & "Tasks.ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (ParentID=" & aTaskComponent(N_PARENT_ID_TASK) & ") " & sCondition & " Order By TaskNumber, TaskName", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	GetTasks = lErrorNumber
	Err.Clear
End Function

Function ModifyTask(oRequest, oADODBConnection, aTaskComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing task in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aTaskComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyTask"
	Dim iIndex
	Dim jIndex
	Dim asTaskLKP
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aTaskComponent(B_COMPONENT_INITIALIZED_TASK)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeTaskComponent(oRequest, aTaskComponent)
	End If

	If (aTaskComponent(N_ID_TASK) = -1) Or (aTaskComponent(N_PROJECT_ID_TASK) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó la actividad a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "TaCoTaskComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If Not CheckTaskInformationConsistency(True, aTaskComponent, sErrorDescription) Then
			lErrorNumber = -1
		Else
			sErrorDescription = "No se pudo modificar la información de la actividad."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update " & TACO_PREFIX & "Tasks Set TaskName='" & Replace(aTaskComponent(S_NAME_TASK), "'", "´") & "', TaskNumber='" & Replace(aTaskComponent(S_NUMBER_TASK), "'", "´") & "', LabelID=" & aTaskComponent(N_LABEL_ID_TASK) & ", TaskDescription='" & Replace(aTaskComponent(S_DESCRIPTION_TASK), "'", "´") & "', TaskObjective='" & Replace(aTaskComponent(S_OBJECTIVE_TASK), "'", "´") & "', TaskStrategy='" & Replace(aTaskComponent(S_STRATEGY_TASK), "'", "´") & "', TaskPurpouse='" & Replace(aTaskComponent(S_PURPOUSE_TASK), "'", "´") & "', TaskIndicator='" & Replace(aTaskComponent(S_INDICATOR_TASK), "'", "´") & "', TaskMeasurement='" & Replace(aTaskComponent(S_MEASUREMENT_TASK), "'", "´") & "', TaskFormula='" & Replace(aTaskComponent(S_FORMULA_TASK), "'", "´") & "', AggregationTypeID=" & aTaskComponent(N_AGGREGATION_TYPE_TASK) & ", TaskComments='" & Replace(aTaskComponent(S_COMMENTS_TASK), "'", "´") & "', StartDate=" & aTaskComponent(N_START_DATE_TASK) & ", EndDate=" & aTaskComponent(N_END_DATE_TASK) & ", TaskFile='" & Replace(aTaskComponent(S_FILE_TASK), "'", "") & "', PuCoSectionID=" & aTaskComponent(N_PUCO_SECTION_ID_TASK) & ", FormID=" & aTaskComponent(N_FORM_TASK) & ", ReportURL='" & aTaskComponent(S_REPORT_URL_TASK) & "' Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			If lErrorNumber = 0 Then
				For iIndex = 0 To UBound(aTaskComponent(AN_PARENT_ID_TASK))
					If CLng(aTaskComponent(AN_PARENT_ID_TASK)(iIndex)) = aTaskComponent(N_PARENT_ID_TASK) Then
						sErrorDescription = "No se pudo modificar la información de la actividad."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update " & TACO_PREFIX & "TasksLKP Set MinimumValue=" & aTaskComponent(AD_MINIMUM_VALUE_TASK)(iIndex) & ", AverageValue=" & aTaskComponent(AD_AVERAGE_VALUE_TASK)(iIndex) & ", MaximumValue=" & aTaskComponent(AD_MAXIMUM_VALUE_TASK)(iIndex) & ", TargetValue=" & aTaskComponent(AD_TARGET_VALUE_TASK)(iIndex) & ", FieldTypeID=" & aTaskComponent(AN_FIELD_TYPE_TASK)(iIndex) & ", TaskPercentage=" & aTaskComponent(AD_PERCENTAGE_TASK)(iIndex) & ", TaskRequired=" & aTaskComponent(AD_REQUIRED_TASK)(iIndex) & " Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ") And (ParentID=" & aTaskComponent(N_PARENT_ID_TASK) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						Exit For
					End If
				Next
			End If

			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudo guardar la información de la nueva actividad."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & TACO_PREFIX & "TasksVariablesLKP  Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				If lErrorNumber = 0 Then
					sErrorDescription = "No se pudo guardar la información de la nueva actividad."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & TACO_PREFIX & "TasksVariablesStatusLKP  Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ") And (StatusDate=0)", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
				If lErrorNumber = 0 Then
					For iIndex = 0 To UBound(aTaskComponent(AL_VARIABLES_TASK))
						sErrorDescription = "No se pudo guardar la información de la nueva actividad."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & TACO_PREFIX & "TasksVariablesLKP (ProjectID, TaskID, VariableID, MinimumValue, AverageValue, MaximumValue, TargetValue, FieldTypeID, Relevance) Values (" & aTaskComponent(N_PROJECT_ID_TASK) & ", " & aTaskComponent(N_ID_TASK) & ", " & aTaskComponent(AL_VARIABLES_TASK)(iIndex) & ", " & aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK)(iIndex) & ", " & aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK)(iIndex) & ", " & aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK)(iIndex) & ", " & aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK)(iIndex) & ", " & aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK)(iIndex) & ", " & aTaskComponent(AN_VARIABLES_RELEVANCE_TASK)(iIndex) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						If lErrorNumber = 0 Then
							sErrorDescription = "No se pudo guardar la información de la nueva actividad."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & TACO_PREFIX & "TasksVariablesStatusLKP (ProjectID, TaskID, VariableID, StatusDate, TaskValue) Values (" & aTaskComponent(N_PROJECT_ID_TASK) & ", " & aTaskComponent(N_ID_TASK) & ", " & aTaskComponent(AL_VARIABLES_TASK)(iIndex) & ", 0, 0)", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						End If
					Next
				End If
			End If

			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudo modificar la información de la actividad."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & TACO_PREFIX & "TaskAreasLKP Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				If lErrorNumber = 0 Then
					asTaskLKP = Split(aTaskComponent(S_AREAS_TASK), ",", -1, vbBinaryCompare)
					For iIndex = 0 To UBound(asTaskLKP)
						sErrorDescription = "No se pudo guardar la información de la nueva actividad."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & TACO_PREFIX & "TaskAreasLKP (ProjectID, TaskID, AreaID) Values (" & aTaskComponent(N_PROJECT_ID_TASK) & ", " & aTaskComponent(N_ID_TASK) & ", " & asTaskLKP(iIndex) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Next
				End If
			End If
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudo modificar la información de la actividad."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & TACO_PREFIX & "TaskUsersLKP Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				If lErrorNumber = 0 Then
					asTaskLKP = Split(aTaskComponent(S_USERS_TASK), ",", -1, vbBinaryCompare)
					For iIndex = 0 To UBound(asTaskLKP)
						sErrorDescription = "No se pudo guardar la información de la nueva actividad."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & TACO_PREFIX & "TaskUsersLKP (ProjectID, TaskID, UserID) Values (" & aTaskComponent(N_PROJECT_ID_TASK) & ", " & aTaskComponent(N_ID_TASK) & ", " & asTaskLKP(iIndex) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Next
				End If
			End If
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudo modificar la información de la actividad."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & TACO_PREFIX & "TaskCategoriesLKP Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				If lErrorNumber = 0 Then
					asTaskLKP = Split(aTaskComponent(S_CATEGORIES_TASK), ",", -1, vbBinaryCompare)
					For iIndex = 0 To UBound(asTaskLKP)
						sErrorDescription = "No se pudo guardar la información de la nueva actividad."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & TACO_PREFIX & "TaskCategoriesLKP (ProjectID, TaskID, CategoryID) Values (" & aTaskComponent(N_PROJECT_ID_TASK) & ", " & aTaskComponent(N_ID_TASK) & ", " & asTaskLKP(iIndex) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Next
				End If
			End If
		End If
	End If

	ModifyTask = lErrorNumber
	Err.Clear
End Function

Function UpdateTaskStatus(oRequest, oADODBConnection, bFromChild, aTaskComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a task from the database
'Inputs:  oRequest, oADODBConnection, bFromChild
'Outputs: aTaskComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UpdateTaskStatus"
	Dim aParentTaskComponent()
	Dim iAggregationType
	Dim sSet
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aTaskComponent(B_COMPONENT_INITIALIZED_TASK)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeTaskComponent(oRequest, aTaskComponent)
	End If

	If (aTaskComponent(N_ID_TASK) = -1) Or (aTaskComponent(N_PROJECT_ID_TASK) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó la actividad a actualizar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "TaCoTaskComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If Not CheckTaskInformationConsistency(False, aTaskComponent, sErrorDescription) Then
			lErrorNumber = -1
		Else
			iAggregationType = 0
			sErrorDescription = "No se pudo obtener el avance de las actividades."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AggregationTypeID From " & TACO_PREFIX & "Tasks Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ")", "TaCoProjectsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then iAggregationType = CInt(oRecordset.Fields("AggregationTypeID").Value)
				oRecordset.Close
			End If

			If bFromChild Then
				aTaskComponent(D_PERCENTAGE_STATUS_TASK) = 0
				sErrorDescription = "No se pudieron actualizar las actividades padre."
				If iAggregationType = 1 Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(" & TACO_PREFIX & "TasksStatusLKP.TaskPercentage) As TotalPercentage From " & TACO_PREFIX & "TasksLKP, " & TACO_PREFIX & "TasksStatusLKP Where (" & TACO_PREFIX & "TasksLKP.ProjectID=" & TACO_PREFIX & "TasksStatusLKP.ProjectID) And (" & TACO_PREFIX & "TasksLKP.TaskID=" & TACO_PREFIX & "TasksStatusLKP.TaskID) And (" & TACO_PREFIX & "TasksLKP.ParentID=" & TACO_PREFIX & "TasksStatusLKP.ParentID) And (" & TACO_PREFIX & "TasksLKP.ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (" & TACO_PREFIX & "TasksLKP.ParentID=" & aTaskComponent(N_ID_TASK) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(" & TACO_PREFIX & "TasksLKP.TaskPercentage*" & TACO_PREFIX & "TasksStatusLKP.TaskPercentage) As TotalPercentage From " & TACO_PREFIX & "TasksLKP, " & TACO_PREFIX & "TasksStatusLKP Where (" & TACO_PREFIX & "TasksLKP.ProjectID=" & TACO_PREFIX & "TasksStatusLKP.ProjectID) And (" & TACO_PREFIX & "TasksLKP.TaskID=" & TACO_PREFIX & "TasksStatusLKP.TaskID) And (" & TACO_PREFIX & "TasksLKP.ParentID=" & TACO_PREFIX & "TasksStatusLKP.ParentID) And (" & TACO_PREFIX & "TasksLKP.ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (" & TACO_PREFIX & "TasksLKP.ParentID=" & aTaskComponent(N_ID_TASK) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				End If
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						aTaskComponent(D_PERCENTAGE_STATUS_TASK) = CDbl(oRecordset.Fields("TotalPercentage").Value)
						If aTaskComponent(D_PERCENTAGE_STATUS_TASK) = 1 Then sSet = ", TaskStatusID=1"
						oRecordset.Close
					End If
				End If
				sErrorDescription = "No se pudo actualizar la información de la actividad."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update " & TACO_PREFIX & "TasksStatusLKP Set TaskPercentage=" & aTaskComponent(D_PERCENTAGE_STATUS_TASK) & sSet & " Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			Else
				sErrorDescription = "No se pudo actualizar la información de la actividad."
				If Len(oRequest("FromParent").Item) > 0 Then
					If aTaskComponent(D_PERCENTAGE_STATUS_TASK) = 1 Then sSet = ", TaskStatusID=1"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update " & TACO_PREFIX & "TasksStatusLKP Set TaskPercentage=" & aTaskComponent(D_PERCENTAGE_STATUS_TASK) & sSet & " Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				Else
					If Len(oRequest("StatusForAllParents").Item) = 0 Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update " & TACO_PREFIX & "TasksStatusLKP Set StartDate=" & aTaskComponent(N_START_DATE_STATUS_TASK) & ", EndDate=" & aTaskComponent(N_END_DATE_STATUS_TASK) & ", TaskValue=" & aTaskComponent(D_VALUE_STATUS_TASK) & ", TaskPercentage=" & aTaskComponent(D_PERCENTAGE_STATUS_TASK) & ", TaskStatusID=" & aTaskComponent(L_STATUS_ID_TASK) & " Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ") And (ParentID=" & aTaskComponent(N_PARENT_ID_TASK) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Else
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update " & TACO_PREFIX & "TasksStatusLKP Set StartDate=" & aTaskComponent(N_START_DATE_STATUS_TASK) & ", EndDate=" & aTaskComponent(N_END_DATE_STATUS_TASK) & ", TaskValue=" & aTaskComponent(D_VALUE_STATUS_TASK) & ", TaskPercentage=" & aTaskComponent(D_PERCENTAGE_STATUS_TASK) & ", TaskStatusID=" & aTaskComponent(L_STATUS_ID_TASK) & " Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
				End If
				If (lErrorNumber = 0) And (Len(oRequest("ValueForAllParents").Item) > 0) Then
					sErrorDescription = "No se pudo actualizar la información de la actividad."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update " & TACO_PREFIX & "TasksStatusLKP Set TaskPercentage=" & aTaskComponent(D_PERCENTAGE_STATUS_TASK) & " Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudieron actualizar las actividades padre."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ParentID From " & TACO_PREFIX & "TasksLKP Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ") Order By ParentID", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						If CLng(oRecordset.Fields("ParentID").Value) > -1 Then
							Redim aParentTaskComponent(N_TASK_COMPONENT_SIZE)
							aParentTaskComponent(N_PROJECT_ID_TASK) = aTaskComponent(N_PROJECT_ID_TASK)
							aParentTaskComponent(N_ID_TASK) = CLng(oRecordset.Fields("ParentID").Value)
							aParentTaskComponent(B_COMPONENT_INITIALIZED_TASK) = True
							lErrorNumber = UpdateTaskStatus(oRequest, oADODBConnection, True, aParentTaskComponent, sErrorDescription)
						End If
						oRecordset.MoveNext
						If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
					Loop
					oRecordset.Close
				End If
			End If
		End If
	End If

	UpdateTaskStatus = lErrorNumber
	Err.Clear
End Function

Function RemoveTask(oRequest, oADODBConnection, aTaskComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a task from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aTaskComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveTask"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aTaskComponent(B_COMPONENT_INITIALIZED_TASK)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeTaskComponent(oRequest, aTaskComponent)
	End If

	If (aTaskComponent(N_ID_TASK) = -1) Or (aTaskComponent(N_PROJECT_ID_TASK) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el actividad a eliminar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "TaCoTaskComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del actividad."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & TACO_PREFIX & "Tasks Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo eliminar la información del actividad."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & TACO_PREFIX & "TasksLKP Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo eliminar la información del actividad."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & TACO_PREFIX & "TasksStatusLKP Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo eliminar la información del actividad."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & TACO_PREFIX & "TaskAreasLKP Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo eliminar la información del actividad."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & TACO_PREFIX & "TaskUsersLKP Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo eliminar la información del actividad."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & TACO_PREFIX & "TaskCategoriesLKP Where (ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (TaskID=" & aTaskComponent(N_ID_TASK) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	RemoveTask = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfTask(oADODBConnection, aTaskComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific task exists in the database
'Inputs:  oADODBConnection, aTaskComponent
'Outputs: aTaskComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfTask"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aTaskComponent(B_COMPONENT_INITIALIZED_TASK)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeTaskComponent(oRequest, aTaskComponent)
	End If

	If Len(aTaskComponent(S_NUMBER_TASK)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó la clave de la actividad para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "TaCoTaskComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del actividad en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select TaskName From " & TACO_PREFIX & "Tasks Where (TaskNumber='" & Replace(aTaskComponent(S_NUMBER_TASK), "'", "´") & "')", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aTaskComponent(B_IS_DUPLICATED_TASK) = True
				aTaskComponent(S_NAME_TASK) = CStr(oRecordset.Fields("TaskName").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	CheckExistencyOfTask = lErrorNumber
	Err.Clear
End Function

Function CheckTaskInformationConsistency(bFull, aTaskComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the database
'Inputs:  bFull, aTaskComponent
'Outputs: aTaskComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckTaskInformationConsistency"
	Dim iIndex
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aTaskComponent(N_PROJECT_ID_TASK)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del proyecto no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aTaskComponent(N_ID_TASK)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador de la actividad no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aTaskComponent(N_PARENT_ID_TASK)) Then aTaskComponent(N_PARENT_ID_TASK) = -1
	If bFull Then
		If Len(aTaskComponent(S_NAME_TASK)) = 0 Then
			sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nombre de la actividad está vacío."
			bIsCorrect = False
		End If
		If Len(aTaskComponent(S_NUMBER_TASK)) = 0 Then
			sErrorDescription = sErrorDescription & "<BR />&nbsp;- La clave de la actividad está vacía."
			bIsCorrect = False
		End If
		If Not IsNumeric(aTaskComponent(N_LABEL_ID_TASK)) Then aTaskComponent(N_LABEL_ID_TASK) = -1
	End If
	If Not IsNumeric(aTaskComponent(N_START_DATE_STATUS_TASK)) Then aTaskComponent(N_START_DATE_STATUS_TASK) = 0
	If Not IsNumeric(aTaskComponent(N_END_DATE_STATUS_TASK)) Then aTaskComponent(N_END_DATE_STATUS_TASK) = 0
	If Not IsNumeric(aTaskComponent(D_VALUE_STATUS_TASK)) Then
		aTaskComponent(D_VALUE_STATUS_TASK) = 0
		For iIndex = 0 To UBound(aTaskComponent(AN_PARENT_ID_TASK))
			If CLng(aTaskComponent(AN_PARENT_ID_TASK)(iIndex)) = aTaskComponent(N_PARENT_ID_TASK) Then
				aTaskComponent(D_VALUE_STATUS_TASK) = aTaskComponent(AD_MINIMUM_VALUE_TASK)(iIndex)
				Exit For
			End If
		Next
	End If
	If Not IsNumeric(aTaskComponent(D_PERCENTAGE_STATUS_TASK)) Then aTaskComponent(N_PARENT_ID_TASK) = 0
	If Not IsNumeric(aTaskComponent(L_STATUS_ID_TASK)) Then aTaskComponent(L_STATUS_ID_TASK) = 0

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información de la actividad contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "TaCoTaskComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckTaskInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function DisplayTask(oRequest, oADODBConnection, bDisplayFields, aTaskComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a task from the
'		  database using a HTML Form
'Inputs:  oRequest, oADODBConnection, bDisplayFields, aTaskComponent
'Outputs: aTaskComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayTask"
	Dim lFieldTypeID
	Dim sLabelName
	Dim sNames
	Dim sAreas
	Dim sUsers
	Dim sCategories
	Dim iIndex
	Dim lErrorNumber

	If (aTaskComponent(N_ID_TASK) <> -1) And (aTaskComponent(N_PROJECT_ID_TASK) <> -1) Then
		lErrorNumber = GetTask(oRequest, oADODBConnection, aTaskComponent, sErrorDescription)
	End If
	lFieldTypeID = 2
	For iIndex = 0 To UBound(aTaskComponent(AN_PARENT_ID_TASK))
		If CLng(aTaskComponent(AN_PARENT_ID_TASK)(iIndex)) = aTaskComponent(N_PARENT_ID_TASK) Then
			lFieldTypeID = aTaskComponent(AN_FIELD_TYPE_TASK)(iIndex)
			If aTaskComponent(N_PARENT_ID_TASK) = -1 Then
				Call GetNameFromTable(oADODBConnection, TACO_PREFIX & "Projects", aTaskComponent(N_PROJECT_ID_TASK), "", "", sNames, "")
				sNames = "el proyecto " & sNames
			Else
				Call GetNameFromTable(oADODBConnection, TACO_PREFIX & "TasksFullName", aTaskComponent(N_PROJECT_ID_TASK) & "," & aTaskComponent(N_PARENT_ID_TASK), "", "", sNames, "")
			End If
			Exit For
		End If
	Next
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function ToogleDiv(sDivName) {" & vbNewLine
				Response.Write "if (document.all[sDivName + 'Div']) {" & vbNewLine
					Response.Write "if (IsDisplayed(document.all[sDivName + 'Div'])) {" & vbNewLine
						Response.Write "HideDisplay(document.all[sDivName + 'Div']);" & vbNewLine
						Response.Write "ToogleImage(document.images[sDivName + 'Img'], 'Images\/BtnArrRight.gif', 'Images\/BtnArrRight.gif');" & vbNewLine
						Response.Write "document.images[sDivName + 'Img'].alt = 'Expandir';" & vbNewLine
					Response.Write "} else {" & vbNewLine
						Response.Write "ShowDisplay(document.all[sDivName + 'Div']);" & vbNewLine
						Response.Write "document.images[sDivName + 'Img'].alt = 'Colapsar';" & vbNewLine
						Response.Write "ToogleImage(document.images[sDivName + 'Img'], 'Images\/BtnArrDown.gif', 'Images\/BtnArrRight.gif');" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of ToogleDiv" & vbNewLine

			Response.Write "function CheckTaskValue(oForm) {" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					If lFieldTypeID = 2 Then
						Response.Write "if (! CheckFloatValue(oForm.TaskStatusValue, 'el valor de la actividad', N_BOTH_FLAG, N_CLOSED_FLAG, " & aTaskComponent(AD_MINIMUM_VALUE_TASK)(iIndex) & ", " & aTaskComponent(AD_MAXIMUM_VALUE_TASK)(iIndex) & "))" & vbNewLine
							Response.Write "return false;" & vbNewLine
					Else
						Response.Write "if (! CheckIntegerValue(oForm.TaskStatusValue, 'el valor de la actividad', N_BOTH_FLAG, N_CLOSED_FLAG, " & aTaskComponent(AD_MINIMUM_VALUE_TASK)(iIndex) & ", " & aTaskComponent(AD_MAXIMUM_VALUE_TASK)(iIndex) & "))" & vbNewLine
							Response.Write "return false;" & vbNewLine
					End If
					Response.Write "if (! CheckFloatValue(oForm.TaskStatusPercentage, 'el porcentaje', N_BOTH_FLAG, N_CLOSED_FLAG, 0, 100))" & vbNewLine
						Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine

				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckTaskValue" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine

		Response.Write "<FORM NAME=""TaskStatusFrm"" ID=""TaskStatusFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""POST"" onSubmit=""return CheckTaskValue(this);""><FONT FACE=""Arial"" SIZE=""2"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProjectID"" ID=""ProjectIDHdn"" VALUE=""" & aTaskComponent(N_PROJECT_ID_TASK) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaskID"" ID=""TaskIDHdn"" VALUE=""" & aTaskComponent(N_ID_TASK) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ParentID"" ID=""ParentIDHdn"" VALUE=""" & aTaskComponent(N_PARENT_ID_TASK) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaskPath"" ID=""TaskPathHdn"" VALUE=""" & aTaskComponent(S_PATH_TASK) & """ />"
			If (Len(aTaskComponent(S_DESCRIPTION_TASK) & aTaskComponent(S_OBJECTIVE_TASK) & aTaskComponent(S_STRATEGY_TASK) & aTaskComponent(S_PURPOUSE_TASK) & aTaskComponent(S_INDICATOR_TASK) & aTaskComponent(S_MEASUREMENT_TASK) & aTaskComponent(S_FORMULA_TASK) & aTaskComponent(S_COMMENTS_TASK)) > 0) Or (aTaskComponent(N_START_DATE_TASK) = 0) Or (aTaskComponent(N_END_DATE_TASK) = 0) Then
				Response.Write "<SPAN CLASS=""TitleBar""><A HREF=""javascript: ToogleDiv('TaskDescription')"" CLASS=""SpecialLink""><IMG SRC=""Images/BtnArrDown.gif"" WIDTH=""13"" HEIGHT=""13"" ALT=""Expandir"" BORDER=""0"" NAME=""TaskDescriptionImg"" HSPACE=""5"" ALIGN=""ABSMIDDLE"" /><FONT COLOR=""#FFFFFF""><B>Información descriptiva</B></FONT></A></SPAN><BR />"
				Response.Write "<DIV ID=""TaskDescriptionDiv"">"
					If Len(aTaskComponent(S_DESCRIPTION_TASK)) > 0 Then
						Response.Write "<B>Descripción: </B>" & CleanStringForHTML(aTaskComponent(S_DESCRIPTION_TASK)) & "<BR /><BR />"
					End If
					If Len(aTaskComponent(S_OBJECTIVE_TASK)) > 0 Then
						Response.Write "<B>Objetivo: </B>" & CleanStringForHTML(aTaskComponent(S_OBJECTIVE_TASK)) & "<BR /><BR />"
					End If
					If Len(aTaskComponent(S_STRATEGY_TASK)) > 0 Then
						Response.Write "<B>Estrategia: </B>" & CleanStringForHTML(aTaskComponent(S_STRATEGY_TASK)) & "<BR /><BR />"
					End If
					If Len(aTaskComponent(S_PURPOUSE_TASK)) > 0 Then
						Response.Write "<B>Propósito: </B>" & CleanStringForHTML(aTaskComponent(S_PURPOUSE_TASK)) & "<BR /><BR />"
					End If
					If Len(aTaskComponent(S_INDICATOR_TASK)) > 0 Then
						Response.Write "<B>Indicador: </B>" & CleanStringForHTML(aTaskComponent(S_INDICATOR_TASK)) & "<BR /><BR />"
					End If
					If Len(aTaskComponent(S_MEASUREMENT_TASK)) > 0 Then
						Response.Write "<B>Unidad de medida: </B>" & CleanStringForHTML(aTaskComponent(S_MEASUREMENT_TASK)) & "<BR /><BR />"
					End If
					If Len(aTaskComponent(S_FORMULA_TASK)) > 0 Then
						Response.Write "<B>Fórmula: </B>" & CleanStringForHTML(aTaskComponent(S_FORMULA_TASK)) & "<BR />"
						If aTaskComponent(B_HAS_CHILDREN_TASK) Then
							Call GetNameFromTable(oADODBConnection, TACO_PREFIX & "AggregationTypes", aTaskComponent(N_AGGREGATION_TYPE_TASK), "", "", sNames, "")
							Response.Write "<B>Este valor se calcula: </B>" & sNames & " de las actividades subordinadas.<BR />"
						End If
						Response.Write "<BR />"
					End If
					If Len(aTaskComponent(S_COMMENTS_TASK)) > 0 Then
						Response.Write "<B>Comentarios: </B>" & CleanStringForHTML(aTaskComponent(S_COMMENTS_TASK)) & "<BR /><BR />"
					End If
					If aTaskComponent(N_START_DATE_TASK) > 0 Then
						Response.Write "<B>Fecha programada de inicio: </B>" & DisplayDateFromSerialNumber(aTaskComponent(N_START_DATE_TASK), -1, -1, -1) & "<BR />"
					End If
					If aTaskComponent(N_END_DATE_TASK) > 0 Then
						Response.Write "<B>Fecha programada de término: </B>" & DisplayDateFromSerialNumber(aTaskComponent(N_END_DATE_TASK), -1, -1, -1) & "<BR />"
					End If
				Response.Write "</DIV>"

				Response.Write "<BR /><IMG SRC=""Images/DotBlack.gif"" WIDTH=""98%"" HEIGHT=""1"" /><BR /><BR />"
			End If

			Call GetNameFromTable(oADODBConnection, TACO_PREFIX & "LabelsFullNameForTask", aTaskComponent(N_PROJECT_ID_TASK) & "," & aTaskComponent(N_ID_TASK), "", "", sNames, "")
			Response.Write "<SPAN CLASS=""TitleBar""><A HREF=""javascript: ToogleDiv('TaskParameters')"" CLASS=""SpecialLink""><IMG SRC=""Images/BtnArrRight.gif"" WIDTH=""13"" HEIGHT=""13"" ALT=""Expandir"" BORDER=""0"" NAME=""TaskParametersImg"" HSPACE=""5"" ALIGN=""ABSMIDDLE"" /><FONT COLOR=""#FFFFFF""><B>Avance en " & sNames & "</B></FONT></A></SPAN><BR />"
			Response.Write "<DIV ID=""TaskParametersDiv"" STYLE=""display: none"">"
				Response.Write "<TABLE BGCOLOR=""#" & S_LIGHT_BGCOLOR & """ BORDER=""0"" CELLPADDIN=""0"" CELLSPACING=""0"">"
					Response.Write "<TR>"
						Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fecha real de inicio: </FONT></TD>"
						Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aTaskComponent(N_START_DATE_STATUS_TASK), "TaskStatusStart", N_START_YEAR, (Year(Date()) + 10), True, True) & "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fecha real de término: </FONT></TD>"
						Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aTaskComponent(N_END_DATE_STATUS_TASK), "TaskStatusEnd", N_START_YEAR, (Year(Date()) + 10), True, True) & "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><B>Valor: </B></TD>"
						Response.Write "<TD VALIGN=""TOP"">"
							Response.Write "<INPUT TYPE=""TEXT"" NAME=""TaskStatusValue"" ID=""TaskStatusValueTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE=""" & aTaskComponent(D_VALUE_STATUS_TASK) & """ CLASS=""TextFields"" />"
							Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>&nbsp;[" & aTaskComponent(AD_MINIMUM_VALUE_TASK)(iIndex) & " - " & aTaskComponent(AD_MAXIMUM_VALUE_TASK)(iIndex) & "] (valor meta: " & aTaskComponent(AD_TARGET_VALUE_TASK)(iIndex) & " " & CleanStringForHTML(aTaskComponent(S_MEASUREMENT_TASK)) & ")</B></FONT>"
						Response.Write "</TD>"
					Response.Write "</TR>"
					If Len(aTaskComponent(S_FORMULA_TASK)) > 0 Then
						Response.Write "<TR>"
							Response.Write "<TD COLSPAN=""2"" VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><B>(" & CleanStringForHTML(aTaskComponent(S_FORMULA_TASK)) & ")</B></FONT></TD>"
						Response.Write "</TR>"
						If aTaskComponent(B_HAS_CHILDREN_TASK) Then
							Call GetNameFromTable(oADODBConnection, TACO_PREFIX & "AggregationTypes", aTaskComponent(N_AGGREGATION_TYPE_TASK), "", "", sNames, "")
							Response.Write "<TR>"
								Response.Write "<TD COLSPAN=""2"" VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Este valor se calcula: " & sNames & " de las actividades subordinadas.</FONT></TD>"
							Response.Write "</TR>"
						End If
					End If
					Response.Write "<TR>"
						Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Porcentaje de avance: </FONT></TD>"
						Response.Write "<TD VALIGN=""TOP"">"
							Response.Write "<INPUT TYPE=""TEXT"" NAME=""TaskStatusPercentage"" ID=""TaskStatusPercentageTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE=""" & FormatNumber((aTaskComponent(D_PERCENTAGE_STATUS_TASK) * 100), 2, True, False, True) & """ CLASS=""TextFields"" onChange=""if (parseFloat(this.value) == 100) {SelectItemByValue('1', false, document.TaskStatusFrm.TaskStatusID);} else {SelectItemByValue('0', false, document.TaskStatusFrm.TaskStatusID);}"" /><FONT FACE=""Arial"" SIZE=""2""> %</FONT>"
							If UBound(aTaskComponent(AN_PARENT_ID_TASK)) > 0 Then Response.Write "<INPUT TYPE=""CHECKBOX"" NAME=""ValueForAllParents"" ID=""ValueForAllParentsChk"" VALUE=""1"" CHECKED=""1"" /><FONT FACE=""Arial"" SIZE=""2""> Aplicar este porcentaje para todas las relaciones de esta actividad con sus padres</FONT>"
						Response.Write "</TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Estatus: </FONT></TD>"
						Response.Write "<TD VALIGN=""TOP"">"
							Response.Write "<SELECT NAME=""TaskStatusID"" ID=""TaskStatusIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""if (this.value == '1') {document.TaskStatusFrm.TaskStatusPercentage.value='100'}"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, TACO_PREFIX & "Status", "StatusID", "StatusName", "", "StatusName", aTaskComponent(L_STATUS_ID_TASK), "", sErrorDescription)
							Response.Write "</SELECT>"
						Response.Write "</FONT></TD>"
					Response.Write "</TR>"
				Response.Write "</TABLE><BR />"

				If UBound(aTaskComponent(AN_PARENT_ID_TASK)) > 0 Then Response.Write "<INPUT TYPE=""CHECKBOX"" NAME=""StatusForAllParents"" ID=""StatusForAllParentsChk"" VALUE=""1"" CHECKED=""0"" /><FONT FACE=""Arial"" SIZE=""2""> Aplicar estos valores para todas las relaciones de esta actividad con sus padres.</FONT><BR />"
				If aTaskComponent(N_EASY_TASK) = 0 Then Response.Write "<BR />"
				Call GetNameFromTable(oADODBConnection, TACO_PREFIX & "LabelsFullNameForTask", aTaskComponent(N_PROJECT_ID_TASK) & "," & aTaskComponent(N_PARENT_ID_TASK), "", "", sLabelName, "")
				Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""UpdateTaskStatus"" ID=""UpdateTaskStatusBtn"" VALUE=""Actualizar " & sLabelName & """ CLASS=""Buttons"" /><BR />"
			Response.Write "</DIV>"

			Call GetNameFromTable(oADODBConnection, TACO_PREFIX & "Areas", aTaskComponent(S_AREAS_TASK), "&nbsp;&nbsp;&nbsp;", "<BR />", sAreas, "")
			Call GetNameFromTable(oADODBConnection, "Users", aTaskComponent(S_USERS_TASK), "&nbsp;&nbsp;&nbsp;", "<BR />", sUsers, "")
			Call GetNameFromTable(oADODBConnection, TACO_PREFIX & "Categories", aTaskComponent(S_CATEGORIES_TASK), "&nbsp;&nbsp;&nbsp;", "<BR />", sCategories, "")
			If Len(sAreas & sUsers & sCategories) > 0 Then
				Response.Write "<BR /><IMG SRC=""Images/DotBlack.gif"" WIDTH=""98%"" HEIGHT=""1"" /><BR /><BR />"
				Response.Write "<SPAN CLASS=""TitleBar""><A HREF=""javascript: ToogleDiv('TaskLKP')"" CLASS=""SpecialLink""><IMG SRC=""Images/BtnArrRight.gif"" WIDTH=""13"" HEIGHT=""13"" ALT=""Expandir"" BORDER=""0"" NAME=""TaskLKPImg"" HSPACE=""5"" ALIGN=""ABSMIDDLE"" /><FONT COLOR=""#FFFFFF""><B>Relaciones</B></FONT></A></SPAN><BR />"
				Response.Write "<DIV ID=""TaskLKPDiv"" STYLE=""display: none"">"
					If Len(sAreas) > 0 Then
						Response.Write "<B>Áreas responsables de esta actividad:</B><BR />"
						Response.Write sAreas & "<BR /><BR />"
					End If

					If Len(sUsers) > 0 Then
						Response.Write "<B>Personas responsables de esta actividad:</B><BR />"
						Response.Write sUsers & "<BR /><BR />"
					End If

					If Len(sCategories) > 0 Then
						Response.Write "<B>Categorías de esta actividad:</B><BR />"
						Response.Write sCategories & "<BR /><BR />"
					End If
				Response.Write "</DIV>"

				Response.Write "<BR /><IMG SRC=""Images/DotBlack.gif"" WIDTH=""98%"" HEIGHT=""1"" /><BR /><BR />"
			End If
		Response.Write "</FONT></FORM>"
	End If

	DisplayTask = lErrorNumber
	Err.Clear
End Function

Function DisplayTaskParents(oRequest, oADODBConnection, aTaskComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information for the task's parents
'Inputs:  oRequest, oADODBConnection, aTaskComponent
'Outputs: aTaskComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayTaskParents"
	Dim iIndex
	Dim jIndex
	Dim sBoldBegin
	Dim sBoldEnd
	Dim sNames
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim asFieldTitles

	asFieldTitles = Split("Valor mínimo,Valor medio,Valor máximo,Valor meta,Tipo de valor,Porcentaje,Obligatorio", ",")
	Response.Write "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
		sRowContents = "&nbsp;"
		asCellAlignments = ""
		For iIndex = 0 To UBound(aTaskComponent(AN_PARENT_ID_TASK))
			If aTaskComponent(AN_PARENT_ID_TASK)(iIndex) = -1 Then
				Call GetNameFromTable(oADODBConnection, TACO_PREFIX & "Projects", aTaskComponent(N_PROJECT_ID_TASK), "", "", sNames, "")
				sNames = "el proyecto " & sNames
			Else
				Call GetNameFromTable(oADODBConnection, TACO_PREFIX & "Tasks", aTaskComponent(N_PROJECT_ID_TASK) & "," & aTaskComponent(AN_PARENT_ID_TASK)(iIndex), "", "", sNames, "")
			End If
			sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
				If aTaskComponent(N_PARENT_ID_TASK) <> CLng(aTaskComponent(AN_PARENT_ID_TASK)(iIndex)) Then sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=Tasks&ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & "&TaskID=" & aTaskComponent(N_ID_TASK) & "&ParentID=" & aTaskComponent(AN_PARENT_ID_TASK)(iIndex) & "&Change=1"""
			sRowContents = sRowContents & "><FONT COLOR=""#FFFFFF"">"
				If Len(sNames) > 20 Then
					sRowContents = sRowContents & "<SPAN TITLE=""" & sNames & """ COLS=""1"">" & Left(sNames, 20) & "...</SPAN>"
				Else
					sRowContents = sRowContents & sNames
				End If
			sRowContents = sRowContents & "</FONT></A>"
			asCellAlignments = asCellAlignments & ",RIGHT"
		Next
		asColumnsTitles = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
		If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
			lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
		Else
			lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
		End If
		asCellAlignments = Split(asCellAlignments, ",", -1, vbBinaryCompare)
		For jIndex = AD_MINIMUM_VALUE_TASK To AD_REQUIRED_TASK
			sRowContents = asFieldTitles(jIndex - AD_MINIMUM_VALUE_TASK)
			For iIndex = 0 To UBound(aTaskComponent(AN_PARENT_ID_TASK))
				sBoldBegin = ""
				sBoldEnd = ""
				If aTaskComponent(N_PARENT_ID_TASK) = CLng(aTaskComponent(AN_PARENT_ID_TASK)(iIndex)) Then
					sBoldBegin = "<B>"
					sBoldEnd = "</B>"
				End If
				Select Case jIndex
					Case AN_FIELD_TYPE_TASK
						Call GetNameFromTable(oADODBConnection, "FieldTypes", aTaskComponent(jIndex)(iIndex), "", "", sNames, "")
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(sNames) & sBoldEnd
					Case AD_PERCENTAGE_TASK
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & FormatNumber((CDbl(aTaskComponent(jIndex)(iIndex)) * 100), True, False, False) & "%" & sBoldEnd
					Case AD_REQUIRED_TASK
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & DisplayYesNo(aTaskComponent(jIndex)(iIndex), True) & sBoldEnd
					Case Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & aTaskComponent(jIndex)(iIndex) & sBoldEnd
				End Select
			Next
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
		Next
	Response.Write "</TABLE>" & vbNewLine

	DisplayTaskParents = lErrorNumber
	Err.Clear
End Function

Function DisplayTaskForm(oRequest, oADODBConnection, sAction, aTaskComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a task from the
'		  database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, aTaskComponent
'Outputs: aTaskComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayTaskForm"
	Dim lFieldTypeID
	Dim sNames
	Dim iIndex
	Dim lErrorNumber

	If (aTaskComponent(N_ID_TASK) <> -1) And (aTaskComponent(N_PROJECT_ID_TASK) <> -1) Then
		lErrorNumber = GetTask(oRequest, oADODBConnection, aTaskComponent, sErrorDescription)
	End If
	lFieldTypeID = 2
	For iIndex = 0 To UBound(aTaskComponent(AN_PARENT_ID_TASK))
		If CLng(aTaskComponent(AN_PARENT_ID_TASK)(iIndex)) = aTaskComponent(N_PARENT_ID_TASK) Then
			lFieldTypeID = aTaskComponent(AN_FIELD_TYPE_TASK)(iIndex)
			If aTaskComponent(N_PARENT_ID_TASK) = -1 Then
				Call GetNameFromTable(oADODBConnection, TACO_PREFIX & "Projects", aTaskComponent(N_PROJECT_ID_TASK), "", "", sNames, "")
				sNames = "el proyecto " & sNames
			Else
				Call GetNameFromTable(oADODBConnection, TACO_PREFIX & "TasksFullName", aTaskComponent(N_PROJECT_ID_TASK) & "," & aTaskComponent(N_PARENT_ID_TASK), "", "", sNames, "")
			End If
			Exit For
		End If
	Next
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckTaskFields(oForm) {" & vbNewLine
				If Len(oRequest("Delete").Item) = 0 Then
					Response.Write "if (oForm) {" & vbNewLine
						Response.Write "if (oForm.TaskName.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de introducir el nombre de la actividad.');" & vbNewLine
							Response.Write "oForm.TaskName.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (oForm.TaskNumber.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de introducir la clave de la actividad.');" & vbNewLine
							Response.Write "oForm.TaskNumber.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine

						Response.Write "if (oForm.TaskFieldTypeID.value == '2') {" & vbNewLine
							Response.Write "if (! CheckFloatValue(oForm.TaskMinimumValue, 'el valor mínimo', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "if (! CheckFloatValue(oForm.TaskAverageValue, 'el valor medio', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "if (! CheckFloatValue(oForm.TaskMaximumValue, 'el valor máximo', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "if (! CheckFloatValue(oForm.TaskTargetValue, 'el valor meta', N_BOTH_FLAG, N_CLOSED_FLAG, parseInt(oForm.TaskMinimumValue.value), parseInt(oForm.TaskMaximumValue.value)))" & vbNewLine
								Response.Write "return false;" & vbNewLine
						Response.Write "} else {" & vbNewLine
							Response.Write "if (! CheckIntegerValue(oForm.TaskMinimumValue, 'el valor mínimo', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "if (! CheckIntegerValue(oForm.TaskAverageValue, 'el valor medio', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "if (! CheckIntegerValue(oForm.TaskMaximumValue, 'el valor máximo', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "if (! CheckIntegerValue(oForm.TaskTargetValue, 'el valor meta', N_BOTH_FLAG, N_CLOSED_FLAG, parseInt(oForm.TaskMinimumValue.value), parseInt(oForm.TaskMaximumValue.value)))" & vbNewLine
								Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine

						Response.Write "if (! CheckFloatValue(oForm.TaskPercentage, 'el porcentaje', N_BOTH_FLAG, N_CLOSED_FLAG, 0, 100))" & vbNewLine
							Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine

'					Response.Write "SelectAllItemsFromList(oForm.VariableIDs);" & vbNewLine
'					Response.Write "SelectAllItemsFromList(oForm.VariableMinimumValues);" & vbNewLine
'					Response.Write "SelectAllItemsFromList(oForm.VariableAverageValues);" & vbNewLine
'					Response.Write "SelectAllItemsFromList(oForm.VariableMaximumValues);" & vbNewLine
'					Response.Write "SelectAllItemsFromList(oForm.VariableTargetValues);" & vbNewLine
'					Response.Write "SelectAllItemsFromList(oForm.VariableFieldTypeIDs);" & vbNewLine
'					Response.Write "SelectAllItemsFromList(oForm.VariableRelevances);" & vbNewLine
				End If
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckTaskFields" & vbNewLine

			Response.Write "function AddVariable() {" & vbNewLine
				Response.Write "oForm = document.TaskFrm;" & vbNewLine
				Response.Write "bCorrect = true;" & vbNewLine

				Response.Write "if (oForm) {" & vbNewLine
					Response.Write "if (! CheckFloatValue(oForm.VariableMinimumValue, 'el valor mínimo de la variable', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
						Response.Write "bCorrect = false;" & vbNewLine
					Response.Write "if (bCorrect && (! CheckFloatValue(oForm.VariableAverageValue, 'el valor medio de la variable', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0)))" & vbNewLine
						Response.Write "bCorrect = false;" & vbNewLine
					Response.Write "if (bCorrect && (! CheckFloatValue(oForm.VariableMaximumValue, 'el valor máximo de la variable', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0)))" & vbNewLine
						Response.Write "bCorrect = false;" & vbNewLine
					Response.Write "if (bCorrect && (! CheckFloatValue(oForm.VariableTargetValue, 'el valor meta de la variable', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0)))" & vbNewLine
						Response.Write "bCorrect = false;" & vbNewLine

					Response.Write "if (bCorrect) {" & vbNewLine
						
						Response.Write "UnselectAllItemsFromList(oForm.VariableIDs);" & vbNewLine
						Response.Write "SelectListItemByValue(oForm.VariableID.value, true, oForm.VariableIDs);" & vbNewLine
						Response.Write "SelectSameItemsForVariables(oForm.VariableIDs);" & vbNewLine
						Response.Write "RemoveVariables();" & vbNewLine

						Response.Write "AddItemToList(GetSelectedText(oForm.VariableID), oForm.VariableID.value, null, oForm.VariableIDs);" & vbNewLine
						Response.Write "AddItemToList(oForm.VariableMinimumValue.value, oForm.VariableMinimumValue.value, null, oForm.VariableMinimumValues);" & vbNewLine
						Response.Write "AddItemToList(oForm.VariableAverageValue.value, oForm.VariableAverageValue.value, null, oForm.VariableAverageValues);" & vbNewLine
						Response.Write "AddItemToList(oForm.VariableMaximumValue.value, oForm.VariableMaximumValue.value, null, oForm.VariableMaximumValues);" & vbNewLine
						Response.Write "AddItemToList(oForm.VariableTargetValue.value, oForm.VariableTargetValue.value, null, oForm.VariableTargetValues);" & vbNewLine
						Response.Write "AddItemToList(GetSelectedText(oForm.VariableFieldTypeID), oForm.VariableFieldTypeID.value, null, oForm.VariableFieldTypeIDs);" & vbNewLine
						Response.Write "AddItemToList(oForm.VariableRelevance.value, oForm.VariableRelevance.value, null, oForm.VariableRelevances);" & vbNewLine

						Response.Write "oForm.VariableIDs.size = oForm.VariableIDs.options.length;" & vbNewLine
						Response.Write "oForm.VariableMinimumValues.size = oForm.VariableMinimumValues.options.length;" & vbNewLine
						Response.Write "oForm.VariableAverageValues.size = oForm.VariableAverageValues.options.length;" & vbNewLine
						Response.Write "oForm.VariableMaximumValues.size = oForm.VariableMaximumValues.options.length;" & vbNewLine
						Response.Write "oForm.VariableTargetValues.size = oForm.VariableTargetValues.options.length;" & vbNewLine
						Response.Write "oForm.VariableFieldTypeIDs.size = oForm.VariableFieldTypeIDs.options.length;" & vbNewLine
						Response.Write "oForm.VariableRelevances.size = oForm.VariableRelevances.options.length;" & vbNewLine

						Response.Write "oForm.VariableMinimumValue.value = '';" & vbNewLine
						Response.Write "oForm.VariableAverageValue.value = '';" & vbNewLine
						Response.Write "oForm.VariableMaximumValue.value = '';" & vbNewLine
						Response.Write "oForm.VariableTargetValue.value = '';" & vbNewLine
						Response.Write "oForm.VariableID.focus();" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of AddVariable" & vbNewLine

			Response.Write "function RemoveVariables() {" & vbNewLine
				Response.Write "oForm = document.TaskFrm;" & vbNewLine

				Response.Write "if (oForm) {" & vbNewLine
					Response.Write "RemoveSelectedItemsFromList(null, oForm.VariableIDs);" & vbNewLine
					Response.Write "RemoveSelectedItemsFromList(null, oForm.VariableMinimumValues);" & vbNewLine
					Response.Write "RemoveSelectedItemsFromList(null, oForm.VariableAverageValues);" & vbNewLine
					Response.Write "RemoveSelectedItemsFromList(null, oForm.VariableMaximumValues);" & vbNewLine
					Response.Write "RemoveSelectedItemsFromList(null, oForm.VariableTargetValues);" & vbNewLine
					Response.Write "RemoveSelectedItemsFromList(null, oForm.VariableFieldTypeIDs);" & vbNewLine
					Response.Write "RemoveSelectedItemsFromList(null, oForm.VariableRelevances);" & vbNewLine

					Response.Write "oForm.VariableIDs.size = oForm.VariableIDs.options.length;" & vbNewLine
					Response.Write "oForm.VariableMinimumValues.size = oForm.VariableMinimumValues.options.length;" & vbNewLine
					Response.Write "oForm.VariableAverageValues.size = oForm.VariableAverageValues.options.length;" & vbNewLine
					Response.Write "oForm.VariableMaximumValues.size = oForm.VariableMaximumValues.options.length;" & vbNewLine
					Response.Write "oForm.VariableTargetValues.size = oForm.VariableTargetValues.options.length;" & vbNewLine
					Response.Write "oForm.VariableFieldTypeIDs.size = oForm.VariableFieldTypeIDs.options.length;" & vbNewLine
					Response.Write "oForm.VariableRelevances.size = oForm.VariableRelevances.options.length;" & vbNewLine

					Response.Write "oForm.VariableID.focus();" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of RemoveVariables" & vbNewLine

			Response.Write "function SelectSameItemsForVariables(oList) {" & vbNewLine
				Response.Write "oForm = document.TaskFrm;" & vbNewLine

				Response.Write "if (oForm) {" & vbNewLine
					Response.Write "SelectSameItems(oList, oForm.VariableIDs)" & vbNewLine
					Response.Write "SelectSameItems(oList, oForm.VariableMinimumValues)" & vbNewLine
					Response.Write "SelectSameItems(oList, oForm.VariableAverageValues)" & vbNewLine
					Response.Write "SelectSameItems(oList, oForm.VariableMaximumValues)" & vbNewLine
					Response.Write "SelectSameItems(oList, oForm.VariableTargetValues)" & vbNewLine
					Response.Write "SelectSameItems(oList, oForm.VariableFieldTypeIDs)" & vbNewLine
					Response.Write "SelectSameItems(oList, oForm.VariableRelevances)" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of SelectSameItemsForVariables" & vbNewLine

			Response.Write "function ToogleDiv(sDivName) {" & vbNewLine
				Response.Write "if (document.all[sDivName + 'Div']) {" & vbNewLine
					Response.Write "if (IsDisplayed(document.all[sDivName + 'Div'])) {" & vbNewLine
						Response.Write "HideDisplay(document.all[sDivName + 'Div']);" & vbNewLine
						Response.Write "ToogleImage(document.images[sDivName + 'Img'], 'Images\/BtnArrRight.gif', 'Images\/BtnArrRight.gif');" & vbNewLine
						Response.Write "document.images[sDivName + 'Img'].alt = 'Expandir';" & vbNewLine
					Response.Write "} else {" & vbNewLine
						Response.Write "ShowDisplay(document.all[sDivName + 'Div']);" & vbNewLine
						Response.Write "document.images[sDivName + 'Img'].alt = 'Colapsar';" & vbNewLine
						Response.Write "ToogleImage(document.images[sDivName + 'Img'], 'Images\/BtnArrDown.gif', 'Images\/BtnArrRight.gif');" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of ToogleDiv" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "<FORM NAME=""TaskFrm"" ID=""TaskFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckTaskFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""Tasks"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProjectID"" ID=""ProjectIDHdn"" VALUE=""" & aTaskComponent(N_PROJECT_ID_TASK) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaskID"" ID=""TaskIDHdn"" VALUE=""" & aTaskComponent(N_ID_TASK) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ParentID"" ID=""ParentIDHdn"" VALUE=""" & aTaskComponent(N_PARENT_ID_TASK) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaskPath"" ID=""TaskPathHdn"" VALUE=""" & aTaskComponent(S_PATH_TASK) & """ />"

			Response.Write "<TABLE BORDER=""0"" CELLPADDIN=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Nombre: </FONT></TD>"
					Response.Write "<TD VALIGN=""TOP""><INPUT TYPE=""TEXT"" NAME=""TaskName"" ID=""TaskNameTxt"" SIZE=""30"" MAXLENGTH=""255"" VALUE=""" & aTaskComponent(S_NAME_TASK) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Clave: </FONT></TD>"
					Response.Write "<TD VALIGN=""TOP""><INPUT TYPE=""TEXT"" NAME=""TaskNumber"" ID=""TaskNumberTxt"" SIZE=""30"" MAXLENGTH=""30"" VALUE=""" & aTaskComponent(S_NUMBER_TASK) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Etiqueta: </FONT></TD>"
					Response.Write "<TD VALIGN=""TOP"">"
						Response.Write "<SELECT NAME=""LabelID"" ID=""LabelIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, TACO_PREFIX & "Labels", "LabelID", "LabelName", "(Active=1)", "LabelName", aTaskComponent(N_LABEL_ID_TASK), "", sErrorDescription)
						Response.Write "</SELECT>"
					Response.Write "</TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE><BR />"
			Response.Write "<FONT FACE=""Arial"" SIZE=""2"">"
				Response.Write "<SPAN CLASS=""TitleBar""><A HREF=""javascript: ToogleDiv('TaskDescription')"" CLASS=""SpecialLink""><IMG SRC=""Images/BtnArrDown.gif"" WIDTH=""13"" HEIGHT=""13"" ALT=""Expandir"" BORDER=""0"" NAME=""TaskDescriptionImg"" HSPACE=""5"" ALIGN=""ABSMIDDLE"" /><FONT COLOR=""#FFFFFF""><B>Información descriptiva</B></FONT></A></SPAN><BR />"
				Response.Write "<DIV ID=""TaskDescriptionDiv"">"
					Response.Write "Descripción:<BR />"
					Response.Write "<TEXTAREA NAME=""TaskDescription"" ID=""TaskDescriptionTxtArea"" ROWS=""5"" COLS=""50"" MAXLENGTH=""4000"" CLASS=""TextFields"">" & aTaskComponent(S_DESCRIPTION_TASK) & "</TEXTAREA><BR /><BR />"

					Response.Write "Objetivo:<BR />"
					Response.Write "<TEXTAREA NAME=""TaskObjective"" ID=""TaskObjectiveTxtArea"" ROWS=""5"" COLS=""50"" MAXLENGTH=""4000"" CLASS=""TextFields"">" & aTaskComponent(S_OBJECTIVE_TASK) & "</TEXTAREA><BR /><BR />"

'					Response.Write "Estrategia:<BR />"
'					Response.Write "<TEXTAREA NAME=""TaskStrategy"" ID=""TaskStrategyTxtArea"" ROWS=""5"" COLS=""50"" MAXLENGTH=""4000"" CLASS=""TextFields"">" & aTaskComponent(S_STRATEGY_TASK) & "</TEXTAREA><BR /><BR />"
'
'					Response.Write "Propósito:<BR />"
'					Response.Write "<TEXTAREA NAME=""TaskPurpouse"" ID=""TaskPurpouseTxtArea"" ROWS=""5"" COLS=""50"" MAXLENGTH=""4000"" CLASS=""TextFields"">" & aTaskComponent(S_PURPOUSE_TASK) & "</TEXTAREA><BR /><BR />"
'
'					Response.Write "Indicador:<BR />"
'					Response.Write "<TEXTAREA NAME=""TaskIndicator"" ID=""TaskIndicatorTxtArea"" ROWS=""5"" COLS=""50"" MAXLENGTH=""4000"" CLASS=""TextFields"">" & aTaskComponent(S_INDICATOR_TASK) & "</TEXTAREA><BR /><BR />"
'
'					Response.Write "Unidad de medida:&nbsp;"
'					Response.Write "<INPUT TYPE=""TEXT"" NAME=""TaskMeasurement"" ID=""TaskMeasurementTxt"" SIZE=""30"" MAXLENGTH=""255"" VALUE=""" & aTaskComponent(S_MEASUREMENT_TASK) & """ CLASS=""TextFields"" /><BR /><BR />"
'
					Response.Write "Productos y cantidad:<BR />"
					Response.Write "<TEXTAREA NAME=""TaskFormula"" ID=""TaskFormulaTxtArea"" ROWS=""5"" COLS=""50"" MAXLENGTH=""4000"" CLASS=""TextFields"">" & aTaskComponent(S_FORMULA_TASK) & "</TEXTAREA><BR /><BR />"
					Response.Write "Este valor se calcula:<BR />"
					Response.Write "<SELECT NAME=""AggregationTypeID"" ID=""AggregationTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, TACO_PREFIX & "AggregationTypes", "AggregationTypeID", "AggregationTypeName", "", "AggregationTypeID", aTaskComponent(N_AGGREGATION_TYPE_TASK), "", sErrorDescription)
					Response.Write "</SELECT>de las actividades subordinadas.<BR /><BR />"

					Response.Write "Insumos y Deptos. o Jefatura de servicios Proveedora:<BR />"
					Response.Write "<TEXTAREA NAME=""TaskComments"" ID=""TaskCommentsTxtArea"" ROWS=""5"" COLS=""50"" MAXLENGTH=""4000"" CLASS=""TextFields"">" & aTaskComponent(S_COMMENTS_TASK) & "</TEXTAREA><BR /><BR />"

					Response.Write "<TABLE BORDER=""0"" CELLPADDIN=""0"" CELLSPACING=""0"">"
						Response.Write "<TR>"
							Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fecha programada de inicio: </FONT></TD>"
							Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aTaskComponent(N_START_DATE_TASK), "Start", N_START_YEAR, (Year(Date()) + 10), True, True) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fecha programada de término: </FONT></TD>"
							Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aTaskComponent(N_END_DATE_TASK), "End", N_START_YEAR, (Year(Date()) + 10), True, True) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD ALIGN=""RIGHT"" COLSPAN=""2""><INPUT TYPE=""BUTTON"" VALUE=""Todo el año en curso"" CLASS=""Buttons"" onClick=""SendURLValuesToForm('StartYear=" & Year(Date()) & "&StartMonth=01&StartDay=01&EndYear=" & Year(Date()) & "&EndMonth=12&EndDay=31', document.TaskFrm)"" / id=1 name=1></TD>"
						Response.Write "</TR>"
'					Response.Write "</TABLE>"
'				Response.Write "</DIV><BR />"

'				Response.Write "<SPAN CLASS=""TitleBar""><A HREF=""javascript: ToogleDiv('TaskFiles')"" CLASS=""SpecialLink""><IMG SRC=""Images/BtnArrRight.gif"" WIDTH=""13"" HEIGHT=""13"" ALT=""Expandir"" BORDER=""0"" NAME=""TaskFilesImg"" HSPACE=""5"" ALIGN=""ABSMIDDLE"" /><FONT COLOR=""#FFFFFF""><B>Reportes, Formularios y mediateca</B></FONT></A></SPAN><BR />"
'				Response.Write "<DIV ID=""TaskFilesDiv"""
'					If (Len(oRequest("New").Item) = 0) And (Len(oRequest("Add").Item) = 0) Then Response.Write " STYLE=""display: none"""
'				Response.Write ">"
'					Response.Write "Archivo:<BR />"
'					Response.Write "<INPUT TYPE=""TEXT"" NAME=""TaskFile"" ID=""TaskFileTxt"" SIZE=""50"" MAXLENGTH=""255"" VALUE=""" & aTaskComponent(S_FILE_TASK) & """ CLASS=""TextFields"" /><BR /><BR />"
'
'					If B_PUCO Then
'						Response.Write "Sección de la mediateca:<BR />"
'						Response.Write "<SELECT NAME=""PuCoSectionID"" ID=""PuCoSectionIDCmb"" SIZE=""1"" CLASS=""Lists"">"
'							Response.Write "<OPTION VALUE=""-1"">Ninguna</OPTION>"
'							Response.Write GenerateListOptionsFromQuery(oPuCoADODBConnection, PUCO_PREFIX & "Sections", "SectionID", "SectionName", "(SectionID>-1)", "SectionName", aTaskComponent(N_PUCO_SECTION_ID_TASK), "", sErrorDescription)
'						Response.Write "</SELECT><BR /><BR />"
'					Else
'						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PuCoSectionID"" ID=""PuCoSectionIDHdn"" VALUE=""" & aTaskComponent(N_PUCO_SECTION_ID_TASK) & """ />"
'					End If
'
'					Response.Write "Formulario:<BR />"
'					Response.Write "<SELECT NAME=""FormID"" ID=""FormIDCmb"" SIZE=""1"" CLASS=""Lists"">"
'						Response.Write "<OPTION VALUE=""-1"">Ninguno</OPTION>"
'						Response.Write GenerateListOptionsFromQuery(oADODBConnection, TACO_PREFIX & "Forms", "FormID", "FormName", "", "FormName", aTaskComponent(N_FORM_TASK), "", sErrorDescription)
'					Response.Write "</SELECT><BR /><BR />"
'
'					Response.Write "Reporte:<BR />"
'					Response.Write "<INPUT TYPE=""TEXT"" NAME=""ReportURL"" ID=""ReportURLTxt"" SIZE=""50"" MAXLENGTH=""255"" VALUE=""" & aTaskComponent(S_REPORT_URL_TASK) & """ CLASS=""TextFields"" /><BR />"
'				Response.Write "</DIV><BR />"

Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaskMinimumValue"" ID=""TaskMinimumValueHdn"" VALUE=""0"" />"
Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaskAverageValue"" ID=""TaskAverageValueHdn"" VALUE=""50"" />"
Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaskMaximumValue"" ID=""TaskMaximumValueHdn"" VALUE=""100"" />"
Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaskFieldTypeID"" ID=""TaskFieldTypeIDHdn"" VALUE=""2"" />"
Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaskRequired"" ID=""TaskRequiredHdn"" VALUE=""1"" />"
'				Response.Write "<SPAN CLASS=""TitleBar""><A HREF=""javascript: ToogleDiv('TaskParameters')"" CLASS=""SpecialLink""><IMG SRC=""Images/BtnArrRight.gif"" WIDTH=""13"" HEIGHT=""13"" ALT=""Expandir"" BORDER=""0"" NAME=""TaskParametersImg"" HSPACE=""5"" ALIGN=""ABSMIDDLE"" /><FONT COLOR=""#FFFFFF""><B>Relación con " & CleanStringForHTML(LCase(sNames)) & "</B></FONT></A></SPAN><BR />"
'				Response.Write "<DIV ID=""TaskParametersDiv"">"
'					Response.Write "<TABLE BORDER=""0"" CELLPADDIN=""0"" CELLSPACING=""0""><TR>"
'						Response.Write "<TD VALIGN=""TOP""><TABLE WIDTH=""427"" BORDER=""0"" CELLPADDIN=""0"" CELLSPACING=""0"">"
'							Response.Write "<TR BGCOLOR=""#" & S_LIGHT_BGCOLOR & """>"
'								Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Valor mínimo (semáforo rojo): </FONT></TD>"
'								Response.Write "<TD VALIGN=""TOP"" COLSPAN=""2""><INPUT TYPE=""TEXT"" NAME=""TaskMinimumValue"" ID=""TaskMinimumValueTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE="""
'									Response.Write aTaskComponent(AD_MINIMUM_VALUE_TASK)(iIndex)
'								Response.Write """ CLASS=""TextFields"" /></TD>"
'							Response.Write "</TR>"
'							Response.Write "<TR BGCOLOR=""#" & S_LIGHT_BGCOLOR & """>"
'								Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Valor medio (semáforo amarillo): </FONT></TD>"
'								Response.Write "<TD VALIGN=""TOP"" COLSPAN=""2""><INPUT TYPE=""TEXT"" NAME=""TaskAverageValue"" ID=""TaskAverageValueTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE="""
'									Response.Write aTaskComponent(AD_AVERAGE_VALUE_TASK)(iIndex)
'								Response.Write """ CLASS=""TextFields"" /></TD>"
'							Response.Write "</TR>"
'							Response.Write "<TR BGCOLOR=""#" & S_LIGHT_BGCOLOR & """>"
'								Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Valor máximo (semáforo verde): </FONT></TD>"
'								Response.Write "<TD VALIGN=""TOP"" COLSPAN=""2""><INPUT TYPE=""TEXT"" NAME=""TaskMaximumValue"" ID=""TaskMaximumValueTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE="""
'									Response.Write aTaskComponent(AD_MAXIMUM_VALUE_TASK)(iIndex)
'								Response.Write """ CLASS=""TextFields"" /></TD>"
'							Response.Write "</TR>"
							Response.Write "<TR BGCOLOR=""#" & S_LIGHT_BGCOLOR & """>"
								Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Valor meta: </FONT></TD>"
								Response.Write "<TD VALIGN=""TOP"" COLSPAN=""2""><INPUT TYPE=""TEXT"" NAME=""TaskTargetValue"" ID=""TaskTargetValueTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE="""
									Response.Write aTaskComponent(AD_TARGET_VALUE_TASK)(iIndex)
								Response.Write """ CLASS=""TextFields"" /></TD>"
							Response.Write "</TR>"
'							Response.Write "<TR BGCOLOR=""#" & S_LIGHT_BGCOLOR & """>"
'								Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Tipo de valor: </FONT></TD>"
'								Response.Write "<TD VALIGN=""TOP"" COLSPAN=""2"">"
'									Response.Write "<SELECT NAME=""TaskFieldTypeID"" ID=""TaskFieldTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
'										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "FieldTypes", "FieldTypeID", "FieldTypeName", "(FieldTypeID In (2,4))", "FieldTypeName", lFieldTypeID, "", sErrorDescription)
'									Response.Write "</SELECT>"
'								Response.Write "</TD>"
'							Response.Write "</TR>"
							Response.Write "<TR>"
								If aTaskComponent(N_PARENT_ID_TASK) = -1 Then
									sNames = "el proyecto"
								Else
									Call GetNameFromTable(oADODBConnection, TACO_PREFIX & "LabelsFullNameForTask", aTaskComponent(N_PROJECT_ID_TASK) & "," & aTaskComponent(N_PARENT_ID_TASK), "", "", sNames, "")
								End If
								Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Porcentaje de participación en " & CleanStringForHTML(LCase(sNames)) & ": </FONT></TD>"
								Response.Write "<TD VALIGN=""TOP""><INPUT TYPE=""TEXT"" NAME=""TaskPercentage"" ID=""TaskPercentageTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE="""
									Response.Write FormatNumber((CDbl(aTaskComponent(AD_PERCENTAGE_TASK)(iIndex)) * 100), 2, True, False, False)
								Response.Write """ CLASS=""TextFields"" /><FONT FACE=""Arial"" SIZE=""2"">%</FONT></TD>"
							Response.Write "</TR>"
'							Response.Write "<TR>"
'								Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">¿Es obligatoria? </FONT></TD>"
'								Response.Write "<TD VALIGN=""TOP"" COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2"">"
'									Response.Write "<INPUT TYPE=""RADIO"" NAME=""TaskRequired"" ID=""TaskvRd"" VALUE=""1"""
'										If CInt(aTaskComponent(AD_REQUIRED_TASK)(iIndex)) = 1 Then Response.Write " CHECKED=""1"""
'									Response.Write " /> Sí&nbsp;&nbsp;&nbsp;"
'									Response.Write "<INPUT TYPE=""RADIO"" NAME=""TaskRequired"" ID=""TaskRequiredRd"" VALUE=""0"""
'										If CInt(aTaskComponent(AD_REQUIRED_TASK)(iIndex)) = 0 Then Response.Write " CHECKED=""1"""
'									Response.Write " /> No"
'								Response.Write "</FONT></TD>"
'							Response.Write "</TR>"
'						Response.Write "</TABLE></TD>"
'						If UBound(aTaskComponent(AN_PARENT_ID_TASK)) > 0 Then
'							Response.Write "<TD><A HREF=""javascript: ToggleDisplay(document.all['TaskParentsDiv']); ToogleImage(document.images['ExpandArrowImg'], 'Images\/ArrExpandLf.gif', 'Images\/ArrExpandRg.gif');""><IMG SRC=""Images/ArrExpandRg.gif"" WIDTH=""11"" HEIGHT=""40"" ALT="""" BORDER=""0"" NAME=""ExpandArrowImg"" /></A></TD>"
'							Response.Write "<TD NAME=""TaskParentsDiv"" ID=""TaskParentsDiv"" VALIGN=""TOP"" STYLE=""display: none"">"
'								lErrorNumber = DisplayTaskParents(oRequest, oADODBConnection, aTaskComponent, sErrorDescription)
'							Response.Write "</TD>"
'							Response.Write "<TD>&nbsp;&nbsp;</TD>"
'						End If
					Response.Write "</TR></TABLE>"
				Response.Write "</DIV><BR />"

'				Response.Write "<SPAN CLASS=""TitleBar""><A HREF=""javascript: ToogleDiv('Variables')"" CLASS=""SpecialLink""><IMG SRC=""Images/BtnArrRight.gif"" WIDTH=""13"" HEIGHT=""13"" ALT=""Expandir"" BORDER=""0"" NAME=""VariablesImg"" HSPACE=""5"" ALIGN=""ABSMIDDLE"" /><FONT COLOR=""#FFFFFF""><B>Variables</B></FONT></A></SPAN><BR />"
'				Response.Write "<DIV ID=""VariablesDiv"" STYLE=""display: none"">"
'					Call GetNameFromTable(oADODBConnection, TACO_PREFIX & "LabelsFullNameForTask", aTaskComponent(N_PROJECT_ID_TASK) & "," & aTaskComponent(N_ID_TASK), "", "", sNames, "")
'					If Len(sNames) = 0 Then sNames = "la actividad"
'					Response.Write "Variables que ponderan " & CleanStringForHTML(LCase(sNames)) & ":<BR />&nbsp;&nbsp;&nbsp;"
'					Response.Write "<TABLE BORDER=""0"" CELLPADDIN=""0"" CELLSPACING=""0""><TR>"
'						Response.Write "<TD VALIGN=""TOP""><TABLE BORDER=""0"" CELLPADDIN=""0"" CELLSPACING=""0"">"
'							Response.Write "<TR>"
'								Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Variable:&nbsp;</FONT></NOBR></TD>"
'								Response.Write "<TD><SELECT NAME=""VariableID"" ID=""VariableIDLst"" SIZE=""1"" CLASS=""Lists"">"
'									Response.Write GenerateListOptionsFromQuery(oADODBConnection, TACO_PREFIX & "Variables", "VariableID", "VariableName", "(Active=1)", "VariableName", "", "", sErrorDescription)
'								Response.Write "</SELECT></TD>"
'							Response.Write "</TR>"
'							Response.Write "<TR>"
'								Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Valor mínimo:&nbsp;</FONT></NOBR></TD>"
'								Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""VariableMinimumValue"" ID=""VariableMinimumValueTxt"" SIZE=""6"" MAXLENGTH=""6"""" CLASS=""TextFields"" /></TD>"
'							Response.Write "</TR>"
'							Response.Write "<TR>"
'								Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Valor medio:&nbsp;</FONT></NOBR></TD>"
'								Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""VariableAverageValue"" ID=""VariableAverageValueTxt"" SIZE=""6"" MAXLENGTH=""6"""" CLASS=""TextFields"" /></TD>"
'							Response.Write "</TR>"
'							Response.Write "<TR>"
'								Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Valor máximo:&nbsp;</FONT></NOBR></TD>"
'								Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""VariableMaximumValue"" ID=""VariableMaximumValueTxt"" SIZE=""6"" MAXLENGTH=""6"""" CLASS=""TextFields"" /></TD>"
'							Response.Write "</TR>"
'							Response.Write "<TR>"
'								Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Valor meta:&nbsp;</FONT></NOBR></TD>"
'								Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""VariableTargetValue"" ID=""VariableTargetValueTxt"" SIZE=""6"" MAXLENGTH=""6"""" CLASS=""TextFields"" /></TD>"
'							Response.Write "</TR>"
'							Response.Write "<TR>"
'								Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Tipo de valor:&nbsp;</FONT></NOBR></TD>"
'								Response.Write "<TD><SELECT NAME=""VariableFieldTypeID"" ID=""VariableFieldTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
'									Response.Write GenerateListOptionsFromQuery(oADODBConnection, TACO_PREFIX & "FieldTypes", "FieldTypeID", "FieldTypeName", "(FieldTypeID In (2,4))", "FieldTypeName", lFieldTypeID, "", sErrorDescription)
'								Response.Write "</SELECT></TD>"
'							Response.Write "</TR>"
'							Response.Write "<TR>"
'								Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Relevancia:&nbsp;</FONT></NOBR></TD>"
'								Response.Write "<TD><SELECT NAME=""VariableRelevance"" ID=""VariableRelevanceCmb"" SIZE=""1"" CLASS=""Lists"">"
'									For iIndex = 1 To 10
'										Response.Write "<OPTION VALUE=""" & iIndex & """>" & iIndex & "</OPTION>"
'									Next
'								Response.Write "</SELECT></TD>"
'							Response.Write "</TR>"
'						Response.Write "</TABLE></TD>"
'						Response.Write "<TD VALIGN=""TOP"">"
'							Response.Write "&nbsp;<A HREF=""javascript: AddVariable()""><IMG SRC=""Images/BtnCrclAdd.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Agregar"" BORDER=""0"" /></A>&nbsp;"
'							Response.Write "<BR /><BR /><BR /><BR /><BR /><BR />"
'							Response.Write "&nbsp;<A HREF=""javascript: RemoveVariables()""><IMG SRC=""Images/BtnCrclDelete.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Quitar"" BORDER=""0"" /></A>&nbsp;"
'						Response.Write "</TD>"
'						Response.Write "<TD VALIGN=""TOP""><TABLE BORDER=""0"" CELLPADDIN=""0"" CELLSPACING=""0"">"
'							Response.Write "<TR>"
'								Response.Write "<TD ALIGN=""CENTER""><NOBR><FONT FACE=""Arial"" SIZE=""2""><B>Variable&nbsp;</B></FONT></NOBR></TD>"
'								Response.Write "<TD ALIGN=""CENTER""><NOBR><FONT FACE=""Arial"" SIZE=""2""><B>Valor mínimo&nbsp;</B></FONT></NOBR></TD>"
'								Response.Write "<TD ALIGN=""CENTER""><NOBR><FONT FACE=""Arial"" SIZE=""2""><B>Valor medio&nbsp;</B></FONT></NOBR></TD>"
'								Response.Write "<TD ALIGN=""CENTER""><NOBR><FONT FACE=""Arial"" SIZE=""2""><B>Valor máximo&nbsp;</B></FONT></NOBR></TD>"
'								Response.Write "<TD ALIGN=""CENTER""><NOBR><FONT FACE=""Arial"" SIZE=""2""><B>Valor meta&nbsp;</B></FONT></NOBR></TD>"
'								Response.Write "<TD ALIGN=""CENTER""><NOBR><FONT FACE=""Arial"" SIZE=""2""><B>Tipo de valor&nbsp;</B></FONT></NOBR></TD>"
'								Response.Write "<TD ALIGN=""CENTER""><NOBR><FONT FACE=""Arial"" SIZE=""2""><B>Relevancia&nbsp;</B></FONT></NOBR></TD>"
'							Response.Write "</TR>"
'							Response.Write "<TR>"
'								iIndex = UBound(aTaskComponent(AL_VARIABLES_TASK)) + 1
'								If iIndex < 1 Then iIndex = 2
'								Response.Write "<TD VALIGN=""TOP""><SELECT NAME=""VariableIDs"" ID=""VariableIDsLst"" SIZE=""" & iIndex & """ MULTIPLE=""1"" CLASS=""Lists"" onChange=""SelectSameItemsForVariables(this);"">"
'									For iIndex = 0 To UBound(aTaskComponent(AL_VARIABLES_TASK))
'										Call GetNameFromTable(oADODBConnection, "Variables", aTaskComponent(AL_VARIABLES_TASK)(iIndex), "", "", sNames, sErrorDescription)
'										Response.Write "<OPTION VALUE=""" & aTaskComponent(AL_VARIABLES_TASK)(iIndex) & """>" & sNames & "</OPTION>"
'									Next
'								Response.Write "</SELECT>&nbsp;</TD>"
'								If iIndex = 0 Then iIndex = 2
'								Response.Write "<TD VALIGN=""TOP""><SELECT NAME=""VariableMinimumValues"" ID=""VariableMinimumValuesLst"" SIZE=""" & iIndex & """ MULTIPLE=""1"" CLASS=""Lists"" onChange=""SelectSameItemsForVariables(this);"">"
'									For iIndex = 0 To UBound(aTaskComponent(AL_VARIABLES_TASK))
'										Response.Write "<OPTION VALUE=""" & aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK)(iIndex) & """>" & aTaskComponent(AD_VARIABLES_MINIMUM_VALUES_TASK)(iIndex) & "</OPTION>"
'									Next
'								Response.Write "</SELECT>&nbsp;</TD>"
'								If iIndex = 0 Then iIndex = 2
'								Response.Write "<TD VALIGN=""TOP""><SELECT NAME=""VariableAverageValues"" ID=""VariableAverageValuesLst"" SIZE=""" & iIndex & """ MULTIPLE=""1"" CLASS=""Lists"" onChange=""SelectSameItemsForVariables(this);"">"
'									For iIndex = 0 To UBound(aTaskComponent(AL_VARIABLES_TASK))
'										Response.Write "<OPTION VALUE=""" & aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK)(iIndex) & """>" & aTaskComponent(AD_VARIABLES_AVERAGE_VALUES_TASK)(iIndex) & "</OPTION>"
'									Next
'								Response.Write "</SELECT>&nbsp;</TD>"
'								If iIndex = 0 Then iIndex = 2
'								Response.Write "<TD VALIGN=""TOP""><SELECT NAME=""VariableMaximumValues"" ID=""VariableMaximumValuesLst"" SIZE=""" & iIndex & """ MULTIPLE=""1"" CLASS=""Lists"" onChange=""SelectSameItemsForVariables(this);"">"
'									For iIndex = 0 To UBound(aTaskComponent(AL_VARIABLES_TASK))
'										Response.Write "<OPTION VALUE=""" & aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK)(iIndex) & """>" & aTaskComponent(AD_VARIABLES_MAXIMUM_VALUES_TASK)(iIndex) & "</OPTION>"
'									Next
'								Response.Write "</SELECT>&nbsp;</TD>"
'								If iIndex = 0 Then iIndex = 2
'								Response.Write "<TD VALIGN=""TOP""><SELECT NAME=""VariableTargetValues"" ID=""VariableTargetValuesLst"" SIZE=""" & iIndex & """ MULTIPLE=""1"" CLASS=""Lists"" onChange=""SelectSameItemsForVariables(this);"">"
'									For iIndex = 0 To UBound(aTaskComponent(AL_VARIABLES_TASK))
'										Response.Write "<OPTION VALUE=""" & aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK)(iIndex) & """>" & aTaskComponent(AD_VARIABLES_TARGET_VALUES_TASK)(iIndex) & "</OPTION>"
'									Next
'								Response.Write "</SELECT>&nbsp;</TD>"
'								If iIndex = 0 Then iIndex = 2
'								Response.Write "<TD VALIGN=""TOP""><SELECT NAME=""VariableFieldTypeIDs"" ID=""VariableFieldTypeIDsLst"" SIZE=""" & iIndex & """ MULTIPLE=""1"" CLASS=""Lists"" onChange=""SelectSameItemsForVariables(this);"">"
'									For iIndex = 0 To UBound(aTaskComponent(AL_VARIABLES_TASK))
'										Call GetNameFromTable(oADODBConnection, "FieldTypes", aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK)(iIndex), "", "", sNames, sErrorDescription)
'										Response.Write "<OPTION VALUE=""" & aTaskComponent(AN_VARIABLES_FIELD_TYPES_TASK)(iIndex) & """>" & sNames & "</OPTION>"
'									Next
'								Response.Write "</SELECT>&nbsp;</TD>"
'								If iIndex = 0 Then iIndex = 2
'								Response.Write "<TD VALIGN=""TOP""><SELECT NAME=""VariableRelevances"" ID=""VariableRelevancesLst"" SIZE=""" & iIndex & """ MULTIPLE=""1"" CLASS=""Lists"" onChange=""SelectSameItemsForVariables(this);"">"
'									For iIndex = 0 To UBound(aTaskComponent(AL_VARIABLES_TASK))
'										Response.Write "<OPTION VALUE=""" & aTaskComponent(AN_VARIABLES_RELEVANCE_TASK)(iIndex) & """>" & aTaskComponent(AN_VARIABLES_RELEVANCE_TASK)(iIndex) & "</OPTION>"
'									Next
'								Response.Write "</SELECT>&nbsp;</TD>"
'							Response.Write "</TR>"
'						Response.Write "</TABLE></TD>"
'					Response.Write "</TR></TABLE>"
'				Response.Write "</DIV><BR />"

'				Response.Write "<SPAN CLASS=""TitleBar""><A HREF=""javascript: ToogleDiv('TaskLKP')"" CLASS=""SpecialLink""><IMG SRC=""Images/BtnArrRight.gif"" WIDTH=""13"" HEIGHT=""13"" ALT=""Expandir"" BORDER=""0"" NAME=""TaskLKPImg"" HSPACE=""5"" ALIGN=""ABSMIDDLE"" /><FONT COLOR=""#FFFFFF""><B>Áreas, responsables y categorías</B></FONT></A></SPAN><BR />"
'				Response.Write "<DIV ID=""TaskLKPDiv"" STYLE=""display: none"">"
'					Response.Write "Áreas responsables de esta actividad:<BR />&nbsp;&nbsp;&nbsp;"
'					Response.Write "<SELECT NAME=""AreasID"" ID=""AreasIDLst"" SIZE=""8"" MULTIPLE=""1"" CLASS=""Lists"">"
'						Response.Write GenerateListOptionsFromQuery(oADODBConnection, TACO_PREFIX & "Areas, " & TACO_PREFIX & "Companies", "AreaID", "CompanyName, AreaName", "(" & TACO_PREFIX & "Areas.CompanyID=" & TACO_PREFIX & "Companies.CompanyID) And (" & TACO_PREFIX & "Areas.Active=1) And (" & TACO_PREFIX & "Companies.Active=1)", "CompanyName, AreaName", aTaskComponent(S_AREAS_TASK), "", sErrorDescription)
'					Response.Write "</SELECT><BR /><BR />"
'					Response.Write "Personas responsables de esta actividad:<BR />&nbsp;&nbsp;&nbsp;"
'					Response.Write "<SELECT NAME=""UsersID"" ID=""UsersIDLst"" SIZE=""8"" MULTIPLE=""1"" CLASS=""Lists"">"
'						Response.Write GenerateListOptionsFromQuery(oADODBConnection, TACO_PREFIX & "Users", "UserID", "UserLastName, UserName", "(UserID>=10)", "UserLastName, UserName", aTaskComponent(S_USERS_TASK), "", sErrorDescription)
'					Response.Write "</SELECT><BR /><BR />"
'					Response.Write "Categorías de esta actividad:<BR />&nbsp;&nbsp;&nbsp;"
'					Response.Write "<SELECT NAME=""CategoriesID"" ID=""CategoriesIDLst"" SIZE=""8"" MULTIPLE=""1"" CLASS=""Lists"">"
'						Response.Write GenerateListOptionsFromQuery(oADODBConnection, TACO_PREFIX & "Categories", "CategoryID", "CategoryName", "(Active=1)", "CategoryName", aTaskComponent(S_CATEGORIES_TASK), "", sErrorDescription)
'					Response.Write "</SELECT><BR />"
'				Response.Write "</DIV><BR />"
			Response.Write "</FONT>"

			If Len(oRequest("Change").Item) > 0 Then
				Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			ElseIf Len(oRequest("Delete").Item) > 0 Then
				Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveTaskWngDiv']); TaskFrm.Remove.focus()"" />"
			Else
				Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=Tasks'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveTaskWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
		Response.Write "</FORM>"
	End If

	DisplayTaskForm = lErrorNumber
	Err.Clear
End Function

Function DisplayTaskPath(oADODBConnection, lProjectID, sPath, bLastTaskSmall, sErrorDescription)
'************************************************************
'Purpose: To display the task path
'Inputs:  oADODBConnection, lProjectID, sPath, bLastTaskSmall
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayTaskPath"
	Dim sProjectName
	Dim sNames
	Dim alTaskID
	Dim iIndex
	Dim sURL
	Dim sTaskPath
	Dim oRecordset
	Dim lErrorNumber


	alTaskID = Split(sPath, ",", -1, vbBinaryCompare)
	sTaskPath = "-1"
	For iIndex = 0 To UBound(alTaskID)
		If CLng(alTaskID(iIndex)) = -1 Then
			Call GetNameFromTable(oADODBConnection, TACO_PREFIX & "Projects", lProjectID, "", "", sProjectName, "")
			If StrComp(GetASPFileName(""), "TaCo.asp", vbBinaryCompare) = 0 Then
				sURL = "Action=Tasks&ProjectID=" & lProjectID
			Else
				sURL = "View=" & oRequest("View").Item & "&ProjectID=" & lProjectID
			End If
			If iIndex = UBound(alTaskID) Then
				If Len(oRequest("AreaID").Item) > 0 Then
					Response.Write "<A HREF=""" & GetASPFileName("") & "?" & sURL & """>" & CleanStringForHTML(sProjectName) & "</A> > "
					Call GetNameFromTable(oADODBConnection, TACO_PREFIX & "Areas", CLng(oRequest("AreaID").Item), "", "", sNames, "")
					Response.Write "<B>" & CleanStringForHTML(sNames) & "</B>"
				ElseIf Len(oRequest("CategoryID").Item) > 0 Then
					Response.Write "<A HREF=""" & GetASPFileName("") & "?" & sURL & """>" & CleanStringForHTML(sProjectName) & "</A> > "
					Call GetNameFromTable(oADODBConnection, TACO_PREFIX & "Categories", CLng(oRequest("CategoryID").Item), "", "", sNames, "")
					Response.Write "<B>" & CleanStringForHTML(sNames) & "</B>"
				ElseIf Len(oRequest("UserID").Item) > 0 Then
					Response.Write "<A HREF=""" & GetASPFileName("") & "?" & sURL & """>" & CleanStringForHTML(sProjectName) & "</A> > "
					Call GetNameFromTable(oADODBConnection, "Users", CLng(oRequest("UserID").Item), "", "", sNames, "")
					Response.Write "<B>" & CleanStringForHTML(sNames) & "</B>"
				Else
					Response.Write "<B>" & CleanStringForHTML(sProjectName) & "</B>"
				End If
			Else
				Response.Write "<A HREF=""" & GetASPFileName("") & "?" & sURL & """>" & CleanStringForHTML(sProjectName) & "</A> > "
				If Len(oRequest("AreaID").Item) > 0 Then
					Call GetNameFromTable(oADODBConnection, TACO_PREFIX & "Areas", CLng(oRequest("AreaID").Item), "", "", sNames, "")
					Response.Write "<A HREF=""" & GetASPFileName("") & "?" & sURL & "&AreaID=" & oRequest("AreaID").Item & """>" & CleanStringForHTML(sNames) & "</A> > "
				ElseIf Len(oRequest("CategoryID").Item) > 0 Then
					Call GetNameFromTable(oADODBConnection, TACO_PREFIX & "Categories", CLng(oRequest("CategoryID").Item), "", "", sNames, "")
					Response.Write "<A HREF=""" & GetASPFileName("") & "?" & sURL & "&CategoryID=" & oRequest("CategoryID").Item & """>" & CleanStringForHTML(sNames) & "</A> > "
				ElseIf Len(oRequest("UserID").Item) > 0 Then
					Call GetNameFromTable(oADODBConnection, "Users", CLng(oRequest("UserID").Item), "", "", sNames, "")
					Response.Write "<A HREF=""" & GetASPFileName("") & "?" & sURL & "&UserID=" & oRequest("UserID").Item & """>" & CleanStringForHTML(sNames) & "</A> > "
				End If
			End If
		Else
			sTaskPath = sTaskPath & "," & alTaskID(iIndex)
			sErrorDescription = "No se pudo obtener la información de la actividad."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select TaskNumber, TaskName, LabelName From " & TACO_PREFIX & "Tasks, " & TACO_PREFIX & "Labels Where (" & TACO_PREFIX & "Tasks.LabelID=" & TACO_PREFIX & "Labels.LabelID) And (ProjectID=" & lProjectID & ") And (TaskID=" & alTaskID(iIndex) & ")", "TaCoTaskComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If iIndex = UBound(alTaskID) Then
					If bLastTaskSmall Then
						Response.Write "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("LabelName").Value)) & " "
							If StrComp(CStr(oRecordset.Fields("TaskNumber").Value), CStr(oRecordset.Fields("TaskName").Value), vbBinaryCompare) = 0 Then
								Response.Write CleanStringForHTML(CStr(oRecordset.Fields("TaskNumber").Value))
							Else
								Response.Write CleanStringForHTML(CStr(oRecordset.Fields("TaskNumber").Value) & " " & CStr(oRecordset.Fields("TaskName").Value))
							End If
						Response.Write "</B>"
					Else
						Response.Write "<BR /><BR /><FONT SIZE=""3""><B>" & CleanStringForHTML(CStr(oRecordset.Fields("LabelName").Value)) & " "
							If StrComp(CStr(oRecordset.Fields("TaskNumber").Value), CStr(oRecordset.Fields("TaskName").Value), vbBinaryCompare) = 0 Then
								Response.Write CleanStringForHTML(CStr(oRecordset.Fields("TaskNumber").Value))
							Else
								Response.Write CleanStringForHTML(CStr(oRecordset.Fields("TaskNumber").Value) & " " & CStr(oRecordset.Fields("TaskName").Value))
							End If
						Response.Write "</B></FONT>"
					End If
				Else
					If StrComp(GetASPFileName(""), "TaCo.asp", vbBinaryCompare) = 0 Then
						sURL = "Action=Tasks&ProjectID=" & lProjectID & "&ParentID=" & alTaskID(iIndex) & "&TaskPath=" & sTaskPath
					Else
						sURL = "View=" & oRequest("View").Item & "&ProjectID=" & lProjectID & "&TaskID=" & alTaskID(iIndex) & "&ParentID=" & alTaskID(iIndex-1) & "&TaskPath=" & sTaskPath & "&AreaID=" & oRequest("AreaID").Item & "&CategoryID=" & oRequest("CategoryID").Item & "&UserID=" & oRequest("UserID").Item
					End If
					Response.Write "<A HREF=""" & GetASPFileName("") & "?" & sURL & """>" & CleanStringForHTML(CStr(oRecordset.Fields("LabelName").Value)) & " "
						If StrComp(CStr(oRecordset.Fields("TaskNumber").Value), CStr(oRecordset.Fields("TaskName").Value), vbBinaryCompare) = 0 Then
							Response.Write CleanStringForHTML(CStr(oRecordset.Fields("TaskNumber").Value))
						Else
							Response.Write CleanStringForHTML(CStr(oRecordset.Fields("TaskNumber").Value) & " " & CStr(oRecordset.Fields("TaskName").Value))
						End If
					Response.Write "</A> > "
				End If
			End If
		End If
	Next
	Response.Write "<BR /><BR />"

	DisplayTaskPath = lErrorNumber
	Err.Clear
End Function

Function DisplayTasksTable(oRequest, oADODBConnection, lIDColumn, bUseLinks, aTaskComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the tasks from
'		  the database in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, aTaskComponent
'Outputs: aTaskComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayTasksTable"
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sBoldBegin
	Dim sBoldEnd
	Dim dPercentage
	Dim lErrorNumber

	lErrorNumber = GetTasks(oRequest, oADODBConnection, aTaskComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE WIDTH=""550"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bUseLinks And (((aLoginComponent(N_TASK_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_TASK_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
					asColumnsTitles = Split("&nbsp;,Clave," & CleanStringForHTML(CStr(oRecordset.Fields("LabelName").Value)) & ",%,Acciones", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,100,270,80,80", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("&nbsp;,Clave," & CleanStringForHTML(CStr(oRecordset.Fields("LabelName").Value)) & ",%", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,140,310,80", ",", -1, vbBinaryCompare)
				End If
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If

				asCellAlignments = Split(",,,RIGHT,RIGHT,CENTER", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("TaskID").Value), oRequest("TaskID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sRowContents = ""
					Select Case lIDColumn
						Case DISPLAY_RADIO_BUTTONS
							sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""TaskID"" ID=""TaskIDRd"" VALUE=""" & CStr(oRecordset.Fields("TaskID").Value) & """ />"
						Case DISPLAY_CHECKBOXES
							sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""TaskID"" ID=""TaskIDChk"" VALUE=""" & CStr(oRecordset.Fields("TaskID").Value) & """ />"
						Case Else
							sRowContents = sRowContents & "&nbsp;"
					End Select
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & "<A HREF=""" & GetASPFileName("") & "?Action=Tasks&ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & "&ParentID=" & CStr(oRecordset.Fields("TaskID").Value) & "&TaskPath=" & aTaskComponent(S_PATH_TASK) & "," & CStr(oRecordset.Fields("TaskID").Value) & """>" & CleanStringForHTML(CStr(oRecordset.Fields("TaskNumber").Value)) & "</A>" & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & "<A HREF=""" & GetASPFileName("") & "?Action=Tasks&ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & "&ParentID=" & CStr(oRecordset.Fields("TaskID").Value) & "&TaskPath=" & aTaskComponent(S_PATH_TASK) & "," & CStr(oRecordset.Fields("TaskID").Value) & """>" & CleanStringForHTML(CStr(oRecordset.Fields("TaskName").Value)) & "</A>" & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & FormatNumber((CDbl(oRecordset.Fields("TaskPercentage").Value) * 100), 2, True, False, True) & "%" & sBoldEnd
					dPercentage = dPercentage + (CDbl(oRecordset.Fields("TaskPercentage").Value) * 100)
					If bUseLinks And (((aLoginComponent(N_TASK_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_TASK_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
						sRowContents = sRowContents & TABLE_SEPARATOR
						If (aLoginComponent(N_TASK_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
							sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Tasks&ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & "&TaskID=" & CStr(oRecordset.Fields("TaskID").Value) & "&ParentID=" & CStr(oRecordset.Fields("ParentID").Value) & "&TaskPath=" & aTaskComponent(S_PATH_TASK) & "," & CStr(oRecordset.Fields("TaskID").Value) & "&Change=1"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
						End If

						If (aLoginComponent(N_TASK_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS Then
							sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Tasks&ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & "&TaskID=" & CStr(oRecordset.Fields("TaskID").Value) & "&ParentID=" & CStr(oRecordset.Fields("ParentID").Value) & "&TaskPath=" & aTaskComponent(S_PATH_TASK) & "," & CStr(oRecordset.Fields("TaskID").Value) & "&Delete=1"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>"
						End If
					End If

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				sRowContents = TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & "<B>" & FormatNumber(dPercentage, 2, True, False, True) & "%</B>"
				If bUseLinks And (((aLoginComponent(N_TASK_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_TASK_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then sRowContents = sRowContents  & TABLE_SEPARATOR
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen actividades registradas en la base de datos."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayTasksTable = lErrorNumber
	Err.Clear
End Function
%>