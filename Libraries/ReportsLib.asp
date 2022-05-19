<!-- #include file="ReportsQueriesLib.asp" -->
<%
Const LOGS_HISTORY_REPORTS = 700
Const AREAS_COUNT_REPORTS = 701
Const EMPLOYEES_COUNT_REPORTS = 702
Const JOBS_COUNT_REPORTS = 703
Const AREAS_LIST_REPORTS = 704
Const EMPLOYEES_LIST_REPORTS = 705
Const JOBS_LIST_REPORTS = 706
Const EMPLOYEE_HISTORY_LIST_REPORTS = 707
Const EMPLOYEE_FORM_HISTORY_LIST_REPORTS = 708
Const EMPLOYEE_PAYMENTS_HISTORY_LIST_REPORTS = 709
Const EMPLOYEE_PAYROLL_REPORTS = 710
Const SPECIAL_JOBS_LIST_REPORTS = 711
Const JOBS_LIST_BY_MODIFY_DATE = 712

Const ISSSTE_1001_REPORTS = 1001
Const ISSSTE_1002_REPORTS = 1002
Const ISSSTE_1003_REPORTS = 1003
Const ISSSTE_1004_REPORTS = 1004
Const ISSSTE_1005_REPORTS = 1005
Const ISSSTE_1006_REPORTS = 1006
Const ISSSTE_1007_REPORTS = 1007
Const ISSSTE_1008_REPORTS = 1008
Const ISSSTE_1009_REPORTS = 1009
Const ISSSTE_1010_REPORTS = 1010
Const ISSSTE_1011_REPORTS = 1011
Const ISSSTE_1012_REPORTS = 1012
Const ISSSTE_1013_REPORTS = 1013
Const ISSSTE_1014_REPORTS = 1014
Const ISSSTE_1015_REPORTS = 1015
Const ISSSTE_1016_REPORTS = 1016
Const ISSSTE_1017_REPORTS = 1017
Const ISSSTE_1018_REPORTS = 1018
Const ISSSTE_1019_REPORTS = 1019
Const ISSSTE_1020_REPORTS = 1020
Const ISSSTE_1021_REPORTS = 1021
Const ISSSTE_1022_REPORTS = 1022
Const ISSSTE_1023_REPORTS = 1023
Const ISSSTE_1024_REPORTS = 1024
Const ISSSTE_1025_REPORTS = 1025
Const ISSSTE_1026_REPORTS = 1026
Const ISSSTE_1027_REPORTS = 1027
Const ISSSTE_1028_REPORTS = 1028
Const ISSSTE_1029_REPORTS = 1029
Const ISSSTE_1030_REPORTS = 1030
Const ISSSTE_1031_REPORTS = 1031
Const ISSSTE_1032_REPORTS = 1032
Const ISSSTE_1033_REPORTS = 1033
Const ISSSTE_1034_REPORTS = 1034
Const ISSSTE_1035_REPORTS = 1035
Const ISSSTE_1100_REPORTS = 1100
Const ISSSTE_1101_REPORTS = 1101
Const ISSSTE_1102_REPORTS = 1102
Const ISSSTE_1103_REPORTS = 1103
Const ISSSTE_1104_REPORTS = 1104
Const ISSSTE_1105_REPORTS = 1105
Const ISSSTE_1106_REPORTS = 1106
Const ISSSTE_1107_REPORTS = 1107
Const ISSSTE_1108_REPORTS = 1108
Const ISSSTE_1109_REPORTS = 1109
Const ISSSTE_1110_REPORTS = 1110
Const ISSSTE_1111_REPORTS = 1111
Const ISSSTE_1112_REPORTS = 1112
Const ISSSTE_1113_REPORTS = 1113
Const ISSSTE_1114_REPORTS = 1114
Const ISSSTE_1115_REPORTS = 1115
Const ISSSTE_1116_REPORTS = 1116
Const ISSSTE_1117_REPORTS = 1117
Const ISSSTE_1118_REPORTS = 1118
Const ISSSTE_1119_REPORTS = 1119
Const ISSSTE_1120_REPORTS = 1120
Const ISSSTE_1151_REPORTS = 1151
Const ISSSTE_1152_REPORTS = 1152
Const ISSSTE_1153_REPORTS = 1153
Const ISSSTE_1154_REPORTS = 1154
Const ISSSTE_1155_REPORTS = 1155
Const ISSSTE_1157_REPORTS = 1157
Const ISSSTE_1200_REPORTS = 1200
Const ISSSTE_1201_REPORTS = 1201
Const ISSSTE_1202_REPORTS = 1202
Const ISSSTE_1203_REPORTS = 1203
Const ISSSTE_1204_REPORTS = 1204
Const ISSSTE_1205_REPORTS = 1205
Const ISSSTE_1206_REPORTS = 1206
Const ISSSTE_1207_REPORTS = 1207
Const ISSSTE_1208_REPORTS = 1208
Const ISSSTE_1209_REPORTS = 1209
Const ISSSTE_1210_REPORTS = 1210
Const ISSSTE_1211_REPORTS = 1211
Const ISSSTE_1221_REPORTS = 1221
Const ISSSTE_1222_REPORTS = 1222
Const ISSSTE_1223_REPORTS = 1223
Const ISSSTE_1224_REPORTS = 1224
Const ISSSTE_1225_REPORTS = 1225
Const ISSSTE_1311_REPORTS = 1311
Const ISSSTE_1334_REPORTS = 1334
Const ISSSTE_1335_REPORTS = 1335
Const ISSSTE_1336_REPORTS = 1336
Const ISSSTE_1337_REPORTS = 1337
Const ISSSTE_1338_REPORTS = 1338
Const ISSSTE_1339_REPORTS = 1339
Const ISSSTE_1340_REPORTS = 1340
Const ISSSTE_1354_REPORTS = 1354
Const ISSSTE_1356_REPORTS = 1356
Const ISSSTE_1364_REPORTS = 1364
Const ISSSTE_1365_REPORTS = 1365
Const ISSSTE_1367_REPORTS = 1367
Const ISSSTE_1369_REPORTS = 1369
Const ISSSTE_1371_REPORTS = 1371
Const ISSSTE_1372_REPORTS = 1372
Const ISSSTE_1373_REPORTS = 1373
Const ISSSTE_1374_REPORTS = 1374
Const ISSSTE_1400_REPORTS = 1400
Const ISSSTE_1401_REPORTS = 1401
Const ISSSTE_1402_REPORTS = 1402
Const ISSSTE_1403_REPORTS = 1403
Const ISSSTE_1404_REPORTS = 1404
Const ISSSTE_1411_REPORTS = 1411
Const ISSSTE_1412_REPORTS = 1412
Const ISSSTE_1413_REPORTS = 1413
Const ISSSTE_1414_REPORTS = 1414
Const ISSSTE_1415_REPORTS = 1415
Const ISSSTE_1416_REPORTS = 1416
Const ISSSTE_1417_REPORTS = 1417
Const ISSSTE_1420_REPORTS = 1420
Const ISSSTE_1421_REPORTS = 1421
Const ISSSTE_1422_REPORTS = 1422
Const ISSSTE_1423_REPORTS = 1423
Const ISSSTE_1424_REPORTS = 1424
Const ISSSTE_1425_REPORTS = 1425
Const ISSSTE_1426_REPORTS = 1426
Const ISSSTE_1427_REPORTS = 1427
Const ISSSTE_1428_REPORTS = 1428
Const ISSSTE_1429_REPORTS = 1429
Const ISSSTE_1430_REPORTS = 1430
Const ISSSTE_1431_REPORTS = 1431
Const ISSSTE_1432_REPORTS = 1432
Const ISSSTE_1433_REPORTS = 1433
Const ISSSTE_1434_REPORTS = 1434
Const ISSSTE_1435_REPORTS = 1435
Const ISSSTE_1470_REPORTS = 1470
Const ISSSTE_1471_REPORTS = 1471
Const ISSSTE_1472_REPORTS = 1472
Const ISSSTE_1473_REPORTS = 1473
Const ISSSTE_1474_REPORTS = 1474
Const ISSSTE_1475_REPORTS = 1475
Const ISSSTE_1476_REPORTS = 1476
Const ISSSTE_1477_REPORTS = 1477
Const ISSSTE_1478_REPORTS = 1478
Const ISSSTE_1490_REPORTS = 1490
Const ISSSTE_1491_REPORTS = 1491
Const ISSSTE_1492_REPORTS = 1492
Const ISSSTE_1493_REPORTS = 1493
Const ISSSTE_1494_REPORTS = 1494
Const ISSSTE_1495_REPORTS = 1495
Const ISSSTE_1496_REPORTS = 1496
Const ISSSTE_1497_REPORTS = 1497
Const ISSSTE_1498_REPORTS = 1498
Const ISSSTE_1499_REPORTS = 1499
Const ISSSTE_1503_REPORTS = 1503
Const ISSSTE_1504_REPORTS = 1504
Const ISSSTE_1561_REPORTS = 1561
Const ISSSTE_1562_REPORTS = 1562
Const ISSSTE_1563_REPORTS = 1563
Const ISSSTE_1571_REPORTS = 1571
Const ISSSTE_1581_REPORTS = 1581
Const ISSSTE_1582_REPORTS = 1582
Const ISSSTE_1583_REPORTS = 1583
Const ISSSTE_1584_REPORTS = 1584
Const ISSSTE_1600_REPORTS = 1600
Const ISSSTE_1602_REPORTS = 1602
Const ISSSTE_1603_REPORTS = 1603
Const ISSSTE_1604_REPORTS = 1604
Const ISSSTE_1605_REPORTS = 1605
Const ISSSTE_1606_REPORTS = 1606
Const ISSSTE_1607_REPORTS = 1607
Const ISSSTE_1608_REPORTS = 1608
Const ISSSTE_1609_REPORTS = 1609
Const ISSSTE_1610_REPORTS = 1610
Const ISSSTE_1611_REPORTS = 1611
Const ISSSTE_1612_REPORTS = 1612
Const ISSSTE_1613_REPORTS = 1613
Const ISSSTE_1701_REPORTS = 1701
Const ISSSTE_1702_REPORTS = 1702
Const ISSSTE_1703_REPORTS = 1703
Const ISSSTE_1704_REPORTS = 1704
Const ISSSTE_2420_REPORTS = 2420
Const ISSSTE_2421_REPORTS = 2421
Const ISSSTE_2422_REPORTS = 2422
Const ISSSTE_2423_REPORTS = 2423
Const ISSSTE_2426_REPORTS = 2426
Const ISSSTE_2427_REPORTS = 2427
Const ISSSTE_2428_REPORTS = 2428
Const ISSSTE_2429_REPORTS = 2429
Const ISSSTE_2430_REPORTS = 2430
Const ISSSTE_2431_REPORTS = 2431
Const ISSSTE_2432_REPORTS = 2432
Const ISSSTE_4701_REPORTS = 4701
Const ISSSTE_4702_REPORTS = 4702
Const ISSSTE_4703_REPORTS = 4703

Const L_ZIP_WARNING_FLAGS = -5
Const L_DONT_CLOSE_FILTER_DIV_FLAGS = -4
Const L_DONT_CLOSE_DIV_FLAGS = -3
Const L_NO_INSTRUCTIONS_FLAGS = -2
Const L_NO_DIV_FLAGS = -1
Const L_USER_FLAGS = 0
Const L_EMPLOYEE_NUMBER_FLAGS = 1
Const L_EMPLOYEE_NAME_FLAGS = 2
Const L_COMPANY_FLAGS = 3
Const L_EMPLOYEE_TYPE_FLAGS = 4
Const L_POSITION_TYPE_FLAGS = 5
Const L_CLASSIFICATION_FLAGS = 6
Const L_GROUP_GRADE_LEVEL_FLAGS = 7
Const L_INTEGRATION_FLAGS = 8
Const L_JOURNEY_FLAGS = 9
Const L_SHIFT_FLAGS = 10
Const L_LEVEL_FLAGS = 11
Const L_EMPLOYEE_STATUS_FLAGS = 12
Const L_PAYMENT_CENTER_FLAGS = 13
Const L_EMPLOYEE_EMAIL_FLAGS = 14
Const L_SOCIAL_SECURITY_NUMBER_FLAGS = 15
Const L_EMPLOYEE_BIRTH_FLAGS = 16
Const L_EMPLOYEE_COUNTRY_FLAGS = 17
Const L_EMPLOYEE_RFC_FLAGS = 18
Const L_EMPLOYEE_CURP_FLAGS = 19
Const L_EMPLOYEE_GENDER_FLAGS = 20
Const L_EMPLOYEE_ACTIVE_FLAGS = 21
Const L_EMPLOYEE_START_DATE_FLAGS = 22
Const L_JOB_NUMBER_FLAGS = 23
Const L_ZONE_FLAGS = 24
Const L_AREA_FLAGS = 25
Const L_POSITION_FLAGS = 26
Const L_JOB_TYPE_FLAGS = 27
Const L_OCCUPATION_TYPE_FLAGS = 28
Const L_JOB_START_DATE_FLAGS = 29
Const L_JOB_END_DATE_FLAGS = 30
Const L_JOB_STATUS_FLAGS = 31
Const L_JOB_ACTIVE_FLAGS = 32
Const L_AREA_CODE_FLAGS = 33
Const L_AREA_SHORT_NAME_FLAGS = 34
Const L_AREA_NAME_FLAGS = 35
Const L_AREA_TYPE_FLAGS = 36
Const L_CONFINE_TYPE_FLAGS = 37
Const L_CENTER_TYPE_FLAGS = 38
Const L_CENTER_SUBTYPE_FLAGS = 39
Const L_ATTENTION_LEVEL_FLAGS = 40
Const L_ECONOMIC_ZONE_FLAGS = 41
Const L_AREA_START_DATE_FLAGS = 42
Const L_AREA_END_DATE_FLAGS = 43
Const L_AREA_JOBS_FLAGS = 44
Const L_AREA_TOTAL_JOBS_FLAGS = 45
Const L_AREA_STATUS_FLAGS = 46
Const L_AREA_ACTIVE_FLAGS = 47
Const L_CONCEPT_ID_FLAGS = 48
Const L_TOTAL_PAYMENT_FLAGS = 49
Const L_BANK_FLAGS = 50
Const L_MEDICAL_AREAS_TYPES_FLAGS = 51
Const L_DOCUMENT_FOR_LICENSE_NUMBER_FLAGS = 52
Const L_DOCUMENT_REQUEST_NUMBER_FLAGS = 53
Const L_DOCUMENT_FOR_CANCEL_LICENSE_NUMBER_FLAGS = 54
Const L_GENERATING_AREAS_FLAGS = 55
Const L_CONCEPTS_VALUES_STATUS_FLAGS = 56
Const L_EMPLOYEE_REASON_ID_FLAGS = 57
Const L_ONE_COMPANY_FLAGS = 58
Const L_ONE_BANK_FLAGS = 59
Const L_ISSSTE_ONE_BANK_FLAGS = 60
Const L_EMPLOYEE_NUMBER7_FLAGS = 61
Const L_EMPLOYEE_NUMBER1_FLAGS = 62
Const L_EMPLOYEE_TYPE1_FLAGS = 63
Const L_ZONE_FOR_PAYMENT_CENTER_FLAGS = 64
Const L_MOVEMENT_TYPE = 65

Const L_PAYROLL_FLAGS = 100
Const L_MONTHS_FLAGS = 101
Const L_YEARS_FLAGS = 102
Const L_DOUBLE_MONTHS_FLAGS = 103
Const L_DATE_FLAGS = 104
Const L_STATES_FLAGS = 105
Const L_THIRD_CONCEPTS_FLAGS = 106
Const L_THIRD_CONCEPTS2_FLAGS = 107
Const L_CHECK_CONCEPTS_FLAGS = 108
Const L_CHECK_CONCEPTS_ALL_FLAGS = 109
Const L_CHECK_CONCEPTS_EMPLOYEES_FLAGS = 110
Const L_ONLY_CHECK_CONCEPTS_EMPLOYEES_FLAGS = 111
Const L_OPEN_PAYROLL_FLAGS = 130
Const L_CLOSED_PAYROLL_FLAGS = 131
Const L_PAYROLL1_FLAGS = 132
Const L_HAS_ALIMONY_FLAGS = 133
Const L_HAS_CREDITS_FLAGS = 134
Const L_CONCEPT_1_FLAGS = 135
Const L_PAYMENT_TYPE_FLAGS = 136
Const L_MEMORY_CONCEPT_ID_FLAGS = 137
Const L_CHECK_NUMBER_FLAGS = 138
Const L_ORDINARY_PAYROLL_FLAGS = 139
Const L_QUARTER_FLAGS = 140
Const L_CONCEPT_2_FLAGS = 141
Const L_CONCENTRATE_CONCEPTS_FLAGS = 142

Const L_PAPERWORK_NUMBER_FLAGS = 150
Const L_PAPERWORK_FOLIO_NUMBER_FLAGS = 151
Const L_PAPERWORK_START_DATE_FLAGS = 152
Const L_PAPERWORK_ESTIMATED_DATE_FLAGS = 153
Const L_PAPERWORK_END_DATE_FLAGS = 154
Const L_PAPERWORK_DOCUMENT_NUMBER_FLAGS = 155
Const L_PAPERWORK_TYPE_FLAGS = 156
Const L_PAPERWORK_OWNER_FLAGS = 157
Const L_PAPERWORK_STATUS_FLAGS = 158
Const L_PAPERWORK_PRIORITY_FLAGS = 159
Const L_PAPERWORK_OWNERS_FLAGS = 160
Const L_PAPERWORK_SUBJECT_TYPES = 161

Const L_STATE_TYPE_FLAGS = 170

Const L_COURSE_NAME_FLAGS = 175
Const L_COURSE_DIPLOMA_FLAGS = 176
Const L_COURSE_LOCATION_FLAGS = 177
Const L_COURSE_DURATION_FLAGS = 178
Const L_COURSE_PARTICIPANTS_FLAGS = 179
Const L_COURSE_DATES_FLAGS = 180
Const L_COURSE_GRADE_FLAGS = 181

Const L_BUDGET_AREA_FLAGS = 200
Const L_BUDGET_COMPANIES_FLAGS = 201
Const L_BUDGET_PROGRAM_DUTY_FLAGS = 202
Const L_BUDGET_FUND_FLAGS = 203
Const L_BUDGET_DUTY_FLAGS = 204
Const L_BUDGET_ACTIVE_DUTY_FLAGS = 205
Const L_BUDGET_SPECIFIC_DUTY_FLAGS = 206
Const L_BUDGET_PROGRAM_FLAGS = 207
Const L_BUDGET_REGION_FLAGS = 208
Const L_BUDGET_UR_FLAGS = 209
Const L_BUDGET_CT_FLAGS = 210
Const L_BUDGET_AUX_FLAGS = 211
Const L_BUDGET_LOCATION_FLAGS = 212
Const L_BUDGET_BUDGET1_FLAGS = 213
Const L_BUDGET_BUDGET2_FLAGS = 214
Const L_BUDGET_BUDGET3_FLAGS = 215
Const L_BUDGET_CONFINE_TYPE_FLAGS = 216
Const L_BUDGET_ACTIVITY1_FLAGS = 217
Const L_BUDGET_ACTIVITY2_FLAGS = 218
Const L_BUDGET_PROCESS_FLAGS = 219
Const L_BUDGET_YEAR_FLAGS = 220
Const L_BUDGET_MONTH_FLAGS = 221
Const L_BUDGET_ORIGINAL_POSITION_FLAGS = 222

Const L_CREDITS_TYPES_ID_FLAGS = 223
Const L_EMPLOYEE_BENEFICIARY_ID = 224
Const S_CREDITS_UPLOADED_FILE_NAME = 225
Const L_ABSENCE_ID_FLAGS = 226
Const L_ZONE_FLAGS_FOR_EMPLOYEES = 227
Const L_ABSENCE_ACTIVE_FLAGS = 228
Const L_ABSENCE_APPLIED_DATE_FLAGS = 229
Const L_CONCEPTS_APPLIED_DATE_FLAGS = 230
Const L_CREDITS_APPLIED_DATE_FLAGS = 231
Const L_ADJUSTMENT_APPLIED_DATE_FLAGS = 232
Const L_EXTRAHOURS_AND_SUNDAYS = 233
Const L_CONCEPT_ACTIVE_FLAGS = 234
Const L_BANK_ACCOUNTS_ACTIVE_FLAGS = 235
Const L_CREDITS_ACTIVE_FLAGS = 236
Const L_EMPLOYEE_CREDITOR_ID = 237
Const L_EMPLOYEE_SERVICES_SHEET_FLAGS = 238

Const L_CANCELL_PAYROLL_FLAGS = 239

Const L_LOG_DATE_FLAGS = 1000
Const L_REPORT_TITLE_FLAGS = 1001
Const L_REPORT_TYPE_FLAGS = 1002

Const L_AUDIT_TYPE_ID_FLAGS = 2000
Const L_AUDIT_OPERATION_TYPE_ID_FLAGS = 2001

Dim sFlags
Dim asTitles
Dim aReportTitle(100)

Const N_ID_REPORTS = 0
Const N_STEP_REPORTS = 1
Const B_READY_REPORTS = 2
Const B_HIDE_CONTINUE_REPORTS = 3
Const L_FLAGS_REPORTS = 4
Const B_COMPONENT_INITIALIZED_REPORTS = 5

Const N_REPORTS_COMPONENT_SIZE = 5

Dim aReportsComponent
Redim aReportsComponent(N_REPORTS_COMPONENT_SIZE)

Function InitializeReportsComponent(oRequest, aReportsComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Reports Component
'         using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aReportsComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeReportsComponent"
	Redim Preserve aReportsComponent(N_REPORTS_COMPONENT_SIZE)

	If IsEmpty(aReportsComponent(N_ID_REPORTS)) Then
		If Len(oRequest("ReportID").Item) > 0 Then
			aReportsComponent(N_ID_REPORTS) = CLng(oRequest("ReportID").Item)
		Else
			aReportsComponent(N_ID_REPORTS) = 0
		End If
	End If

	If IsEmpty(aReportsComponent(N_STEP_REPORTS)) Then
		If Len(oRequest("ReportStep").Item) > 0 Then
			aReportsComponent(N_STEP_REPORTS) = CInt(oRequest("ReportStep").Item)
		Else
			aReportsComponent(N_STEP_REPORTS) = 1
		End If
	End If

	If IsEmpty(aReportsComponent(B_READY_REPORTS)) Then
		aReportsComponent(B_READY_REPORTS) = (Len(oRequest("ReportReady").Item) > 0)
	End If

	aReportsComponent(B_HIDE_CONTINUE_REPORTS) = False

	If IsEmpty(aReportsComponent(L_FLAGS_REPORTS)) Then
		If Len(oRequest("ReportFlags").Item) > 0 Then
			aReportsComponent(L_FLAGS_REPORTS) = CLng(oRequest("ReportFlags").Item)
		Else
			aReportsComponent(L_FLAGS_REPORTS) = 0
		End If
	End If

	aReportsComponent(B_COMPONENT_INITIALIZED_REPORTS) = True
	InitializeReportsComponent = Err.number
	Err.Clear
End Function

Function BuildRowData(oRecordset, iFirstColumn, iLastColumn, sCurrentRecords)
'************************************************************
'Purpose: To build a table row using the given recordset
'Inputs:  oRecordset, iFirstColumn, iLastColumn
'Outputs: sCurrentRecords
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildRowData"
	Dim iIndex

	If iLastColumn = -1 Then iLastColumn = oRecordset.Fields.Count - 1
	sCurrentRecords = ""
	For iIndex = iFirstColumn To iLastColumn
		If IsNull(oRecordset.Fields(iIndex).Value) Then
			sCurrentRecords = sCurrentRecords & TABLE_SEPARATOR
		ElseIf InStr(1, oRecordset.Fields(iIndex).Name, "DATE", vbBinaryCompare) > 0 Then
			If (CLng(oRecordset.Fields(iIndex).Value) = 0) Or (CLng(oRecordset.Fields(iIndex).Value) = 30000000) Then
				sCurrentRecords = sCurrentRecords & "<CENTER>---</CENTER>" & TABLE_SEPARATOR
			Else
				sCurrentRecords = sCurrentRecords & DisplayDateFromSerialNumber(CStr(oRecordset.Fields(iIndex).Value), -1, -1, -1) & TABLE_SEPARATOR
			End If
		ElseIf InStr(1, oRecordset.Fields(iIndex).Name, "HOUR", vbBinaryCompare) > 0 Then
			sCurrentRecords = sCurrentRecords & CStr(oRecordset.Fields(iIndex).Value) & ":00 - " & CStr(oRecordset.Fields(iIndex).Value) & ":59 hrs" & TABLE_SEPARATOR
		ElseIf InStr(1, oRecordset.Fields(iIndex).Name, "MONTH", vbBinaryCompare) > 0 Then
			sCurrentRecords = sCurrentRecords & asMonthNames_es(CInt(oRecordset.Fields(iIndex).Value)) & TABLE_SEPARATOR
		ElseIf InStr(1, ",SHORTNAME,", "," & oRecordset.Fields(iIndex).Name & ",", vbBinaryCompare) > 0 Then
			If iIndex < iLastColumn Then
				If (InStr(1, oRecordset.Fields(iIndex).Name, "SHORTNAME", vbBinaryCompare) > 0) And (InStr(1, oRecordset.Fields(iIndex + 1).Name, "NAME", vbBinaryCompare) > 0) And (InStr(1, oRecordset.Fields(iIndex + 1).Name, "SHORTNAME", vbBinaryCompare) = 0) Then
					sCurrentRecords = sCurrentRecords & CleanStringForHTML(CStr(oRecordset.Fields(iIndex).Value))
					sCurrentRecords = sCurrentRecords & ". " & CleanStringForHTML(CStr(oRecordset.Fields(iIndex + 1).Value))
					sCurrentRecords = sCurrentRecords & TABLE_SEPARATOR
					iIndex = iIndex + 1
				End If
			End If
		ElseIf InStr(1, ",Active,", "," & oRecordset.Fields(iIndex).Name & ",", vbBinaryCompare) > 0 Then
			sCurrentRecords = sCurrentRecords & DisplayYesNo(CStr(oRecordset.Fields(iIndex).Value), True) & TABLE_SEPARATOR
		ElseIf StrComp(GetASPFileName(""), "Export.asp", vbTextCompare) = 0 Then
			If (InStr(1, CStr(oRecordset.Fields(iIndex).Value), vbNewLine, vbBinaryCompare) = 0) And (InStr(1, CStr(oRecordset.Fields(iIndex).Value), """", vbBinaryCompare) = 0) Then
				sCurrentRecords = sCurrentRecords & "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields(iIndex).Value)) & """)" & TABLE_SEPARATOR
			Else
				sCurrentRecords = sCurrentRecords & CleanStringForHTML(CStr(oRecordset.Fields(iIndex).Value)) & TABLE_SEPARATOR
			End If
		Else
			sCurrentRecords = sCurrentRecords & CleanStringForHTML(CStr(oRecordset.Fields(iIndex).Value)) & TABLE_SEPARATOR
		End If
	Next
	If Len(sCurrentRecords) > 0 Then sCurrentRecords = Left(sCurrentRecords, (Len(sCurrentRecords) - Len(TABLE_SEPARATOR)))

	BuildRowData = Err.number
	Err.Clear
End Function 

Function BuildTableTemplateHeader(oRequest, sFlags, sDataHeaderNames)
'************************************************************
'Purpose: To build the table headers using the URL or Flags
'Inputs:  oRequest, sFlags, sDataHeaderNames
'Outputs: A list with the headers
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildTableTemplateHeader"
	Dim sColumnsTitles
	Dim oItem
	Dim asFlags
	Dim iIndex

	sColumnsTitles = ""
	If Len(oRequest("Template").Item) > 0 Then
		For Each oItem In oRequest("Template")
			If Len(GetFlagName(CLng(oItem))) > 0 Then
				sColumnsTitles = sColumnsTitles & GetFlagName(CLng(oItem)) & ","
			End If
		Next
	Else
		asFlags = Split(sFlags, ",", -1, vbBinaryCompare)
		For iIndex = 0 To UBound(asFlags)
			If Len(GetFlagName(CLng(asFlags(iIndex)))) > 0 Then
				sColumnsTitles = sColumnsTitles & GetFlagName(CLng(asFlags(iIndex))) & ","
			End If
		Next
	End If
	sColumnsTitles = Trim(sColumnsTitles & sDataHeaderNames)
	If StrComp(Right(sColumnsTitles, Len(",")), ",", vbBinaryCompare) = 0 Then sColumnsTitles = Left(sColumnsTitles, (Len(sColumnsTitles) - Len(",")))

	BuildTableTemplateHeader = sColumnsTitles
	Err.Clear
End Function

Function DisplayFilterAsHidden(oRequest, sFlags, sErrorDescription)
'************************************************************
'Purpose: To display the filter information using the user's
'         selections
'Inputs:  oRequest, sFlags
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayFilterAsHidden"
	Dim sID
	Dim lErrorNumber

	sID = oRequest("UserID").Item
	bFromRequest = (Err.number = 0)
	Err.clear

	sFlags = "," & sFlags & ","
	If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_USER_FLAGS & ",")) > 0)) Then
		If bFromRequest Then
			sID = oRequest("UserID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "UserID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserID"" ID=""UserIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_EMPLOYEE_NUMBER_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_EMPLOYEE_NUMBER1_FLAGS & ",")) > 0)) Then
		If bFromRequest Then
			sID = Replace(Replace(Replace(oRequest("EmployeeNumbers").Item, " ", ""), vbNewLine, ","), ",,", ",")
			Do While (InStr(1, sID, ",,", vbBinaryCompare) > 0)
				sID = Replace(oRequest("EmployeeNumbers").Item, ",,", ",")
			Loop
			If Len(sID) = 0 Then sID = oRequest("EmployeeNumber").Item
			If Len(sID) = 0 Then sID = oRequest("EmployeeIDs").Item
		Else
			sID = Replace(Replace(Replace(GetParameterFromURLString(oRequest, "EmployeeNumbers"), " ", ""), vbNewLine, ","), ",,", ",")
			Do While (InStr(1, sID, ",,", vbBinaryCompare) > 0)
				sID = Replace(sID, ",,", ",")
			Loop
			If Len(sID) = 0 Then sID = GetParameterFromURLString(oRequest, "EmployeeNumber")
			If Len(sID) = 0 Then sID = GetParameterFromURLString(oRequest, "EmployeeIDs")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeNumbers"" ID=""EmployeeNumbersHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_NUMBER7_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("EmployeeTempNumber").Item
		Else
			sID = GetParameterFromURLString(oRequest, "EmployeeTempNumber")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeNumberTemp"" ID=""EmployeeNumberTempHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_NAME_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("EmployeeName").Item
		Else
			sID = GetParameterFromURLString(oRequest, "EmployeeName")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeName"" ID=""EmployeeNameHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_COMPANY_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_ONE_COMPANY_FLAGS & ",")) > 0)) Then
		If bFromRequest Then
			sID = oRequest("CompanyID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CompanyID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CompanyID"" ID=""CompanyIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_EMPLOYEE_TYPE_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_EMPLOYEE_TYPE1_FLAGS & ",")) > 0)) Then
		If bFromRequest Then
			sID = oRequest("EmployeeTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "EmployeeTypeID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_POSITION_TYPE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("PositionTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "PositionTypeID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PositionTypeID"" ID=""PositionTypeIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CLASSIFICATION_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("ClassificationID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "ClassificationID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ClassificationID"" ID=""ClassificationIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_GROUP_GRADE_LEVEL_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("GroupGradeLevelID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "GroupGradeLevelID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""GroupGradeLevelID"" ID=""GroupGradeLevelIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_INTEGRATION_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("IntegrationID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "IntegrationID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IntegrationID"" ID=""IntegrationIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_JOURNEY_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("JourneyID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "JourneyID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JourneyID"" ID=""JourneyIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_SHIFT_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("ShiftID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "ShiftID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ShiftID"" ID=""ShiftIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_LEVEL_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("LevelID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "LevelID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""LevelID"" ID=""LevelIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_STATUS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("EmployeeStatusID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "EmployeeStatusID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeStatusID"" ID=""EmployeeStatusIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAYMENT_CENTER_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("PaymentCenterID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "PaymentCenterID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PaymentCenterID"" ID=""PaymentCenterIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_EMAIL_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("EmployeeEmail").Item
		Else
			sID = GetParameterFromURLString(oRequest, "EmployeeEmail")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeEmail"" ID=""EmployeeEmailHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_SOCIAL_SECURITY_NUMBER_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("SocialSecurityNumber").Item
		Else
			sID = GetParameterFromURLString(oRequest, "SocialSecurityNumber")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SocialSecurityNumber"" ID=""SocialSecurityNumberHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_BIRTH_FLAGS & ",")) > 0) Then
		Call DisplaySerialDateAsHidden(GetSerialNumberFromURL("StartBirth"), "StartBirth")
		Call DisplaySerialDateAsHidden(GetSerialNumberFromURL("EndBirth"), "EndBirth")
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_COUNTRY_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("CountryID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CountryID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CountryID"" ID=""CountryIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_RFC_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("RFC").Item
		Else
			sID = GetParameterFromURLString(oRequest, "RFC")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RFC"" ID=""RFCHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_CURP_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("CURP").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CURP")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CURP"" ID=""CURPHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_GENDER_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("GenderID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "GenderID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""GenderID"" ID=""GenderIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_ACTIVE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("EmployeeActive").Item
		Else
			sID = GetParameterFromURLString(oRequest, "EmployeeActive")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeActive"" ID=""EmployeeActiveHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_START_DATE_FLAGS & ",")) > 0) Then
		Call DisplaySerialDateAsHidden(GetSerialNumberFromURL("StartEmployeeStart"), "StartEmployeeStart")
		Call DisplaySerialDateAsHidden(GetSerialNumberFromURL("EndEmployeeStart"), "EndEmployeeStart")
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_JOB_NUMBER_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("JobNumber").Item
		Else
			sID = GetParameterFromURLString(oRequest, "JobNumber")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobNumber"" ID=""JobNumberHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_ZONE_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_STATES_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_ZONE_FLAGS_FOR_EMPLOYEES & ",")) > 0)) Then
		If bFromRequest Then
			sID = oRequest("ZoneID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "ZoneID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ZoneID"" ID=""ZoneIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_ZONE_FOR_PAYMENT_CENTER_FLAGS & ",")) > 0)) Then
		If bFromRequest Then
			sID = oRequest("ZoneForPaymentCenterID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "ZoneForPaymentCenterID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ZoneForPaymentCenterID"" ID=""ZoneForPaymentCenterIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AreaID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AreaID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaID"" ID=""AreaIDHdn"" VALUE=""" & sID & """ />"
		If bFromRequest Then
			sID = oRequest("SubAreaID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "SubAreaID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SubAreaID"" ID=""SubAreaIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_POSITION_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("PositionID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "PositionID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PositionID"" ID=""PositionIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_JOB_TYPE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("JobTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "JobTypeID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobTypeID"" ID=""JobTypeIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_OCCUPATION_TYPE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("OccupationTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "OccupationTypeID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OccupationTypeID"" ID=""OccupationTypeIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_JOB_START_DATE_FLAGS & ",")) > 0) Then
		Call DisplaySerialDateAsHidden(GetSerialNumberFromURL("StartJobStart"), "StartJobStart")
		Call DisplaySerialDateAsHidden(GetSerialNumberFromURL("EndJobStart"), "EndJobStart")
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_JOB_END_DATE_FLAGS & ",")) > 0) Then
		Call DisplaySerialDateAsHidden(GetSerialNumberFromURL("StartJobEnd"), "StartJobEnd")
		Call DisplaySerialDateAsHidden(GetSerialNumberFromURL("EndJobEnd"), "EndJobEnd")
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_JOB_STATUS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("JobStatusID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "JobStatusID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobStatusID"" ID=""JobStatusIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_JOB_ACTIVE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("JobActive").Item
		Else
			sID = GetParameterFromURLString(oRequest, "JobActive")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobActive"" ID=""JobActiveHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_CODE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AreaCode").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AreaCode")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaCode"" ID=""AreaCodeHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_SHORT_NAME_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AreaShortName").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AreaShortName")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaShortName"" ID=""AreaShortNameHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_NAME_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AreaName").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AreaName")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaName"" ID=""AreaNameHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_TYPE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AreaTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AreaTypeID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaTypeID"" ID=""AreaTypeIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CONFINE_TYPE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("ConfineTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "ConfineTypeID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConfineTypeID"" ID=""ConfineTypeIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CENTER_TYPE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("CenterTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CenterTypeID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CenterTypeID"" ID=""CenterTypeIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CENTER_SUBTYPE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("CenterSubtypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CenterSubtypeID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CenterSubtypeID"" ID=""CenterSubtypeIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_ATTENTION_LEVEL_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AttentionLevelID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AttentionLevelID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AttentionLevelID"" ID=""AttentionLevelIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_ECONOMIC_ZONE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("EconomicZoneID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "EconomicZoneID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EconomicZoneID"" ID=""EconomicZoneIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_START_DATE_FLAGS & ",")) > 0) Then
		Call DisplaySerialDateAsHidden(GetSerialNumberFromURL("StartAreaStart"), "StartAreaStart")
		Call DisplaySerialDateAsHidden(GetSerialNumberFromURL("EndAreaStart"), "EndAreaStart")
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_END_DATE_FLAGS & ",")) > 0) Then
		Call DisplaySerialDateAsHidden(GetSerialNumberFromURL("StartAreaEnd"), "StartAreaEnd")
		Call DisplaySerialDateAsHidden(GetSerialNumberFromURL("EndAreaEnd"), "EndAreaEnd")
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_JOBS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("Jobs").Item
		Else
			sID = GetParameterFromURLString(oRequest, "Jobs")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Jobs"" ID=""JobsHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_TOTAL_JOBS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("TotalJobs").Item
		Else
			sID = GetParameterFromURLString(oRequest, "TotalJobs")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TotalJobs"" ID=""TotalJobsHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_STATUS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AreaStatusID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AreaStatusID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaStatusID"" ID=""AreaStatusIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_GENERATING_AREAS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("GeneratingAreaID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "GeneratingAreaID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""GeneratingAreaID"" ID=""GeneratingAreaIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CONCEPTS_VALUES_STATUS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("ConceptStatusID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "ConceptStatusID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptStatusID"" ID=""ConceptStatusIDIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_REASON_ID_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("ReasonID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "ReasonID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_ACTIVE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AreaActive").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AreaActive")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaActive"" ID=""AreaActiveHdn"" VALUE=""" & sID & """ />"
	End If
    If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_CONCEPT_ID_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_CONCEPT_1_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_THIRD_CONCEPTS_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_THIRD_CONCEPTS2_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_MEMORY_CONCEPT_ID_FLAGS & ",")) > 0)) Then
		If bFromRequest Then
			sID = oRequest("ConceptID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "ConceptID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptID"" ID=""ConceptIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_TOTAL_PAYMENT_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("TotalPaymentMin").Item
		Else
			sID = GetParameterFromURLString(oRequest, "TotalPaymentMin")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TotalPaymentMin"" ID=""TotalPaymentMinHdn"" VALUE=""" & sID & """ />"
		If bFromRequest Then
			sID = oRequest("TotalPaymentMax").Item
		Else
			sID = GetParameterFromURLString(oRequest, "TotalPaymentMax")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TotalPaymentMax"" ID=""TotalPaymentMaxHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_BANK_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_ONE_BANK_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_ISSSTE_ONE_BANK_FLAGS & ",")) > 0)) Then
		If bFromRequest Then
			sID = oRequest("BankID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BankID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BankID"" ID=""BankIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_MEDICAL_AREAS_TYPES_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("MedicalAreasTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "MedicalAreasTypeID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""MedicalAreasTypeID"" ID=""MedicalAreasTypeIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_DOCUMENT_FOR_LICENSE_NUMBER_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("DocumentForLicenseNumber").Item
		Else
			sID = GetParameterFromURLString(oRequest, "DocumentForLicenseNumber")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DocumentForLicenseNumber"" ID=""DocumentForLicenseNumberHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_DOCUMENT_REQUEST_NUMBER_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("RequestNumber").Item
		Else
			sID = GetParameterFromURLString(oRequest, "RequestNumber")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RequestNumber"" ID=""RequestNumberHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_DOCUMENT_FOR_CANCEL_LICENSE_NUMBER_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("DocumentForCancelLicenseNumber").Item
		Else
			sID = GetParameterFromURLString(oRequest, "DocumentForCancelLicenseNumber")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DocumentForCancelLicenseNumber"" ID=""DocumentForCancelLicenseNumberHdn"" VALUE=""" & sID & """ />"
	End If

	If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_PAYROLL_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_OPEN_PAYROLL_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_CLOSED_PAYROLL_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_PAYROLL1_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_CANCELL_PAYROLL_FLAGS & ",")) > 0)) Then
		If bFromRequest Then
			sID = oRequest("PayrollID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "PayrollID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PayrollID"" ID=""PayrollIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_MONTHS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("MonthID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "MonthID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""MonthID"" ID=""MonthIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_YEARS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("YearID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "YearID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""YearID"" ID=""YearIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_DOUBLE_MONTHS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("StartMonthID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "StartMonthID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartMonthID"" ID=""StartMonthIDHdn"" VALUE=""" & sID & """ />"
		If bFromRequest Then
			sID = oRequest("EndMonthID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "EndMonthID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EndMonthID"" ID=""EndMonthIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_DATE_FLAGS & ",")) > 0) Then
		lErrorNumber = GetDateRank(oRequest, "Start", "End", True, sDate)
		If lErrorNumber = 0 Then sFilter = sFilter & "<B>Periodo:</B><BR />&nbsp;&nbsp;&nbsp;" & sDate & "<BR />"
	End If
	If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_CHECK_CONCEPTS_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_CHECK_CONCEPTS_ALL_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_CHECK_CONCEPTS_EMPLOYEES_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_ONLY_CHECK_CONCEPTS_EMPLOYEES_FLAGS & ",")) > 0)) Then
		If bFromRequest Then
			sID = oRequest("CheckConceptID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CheckConceptID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CheckConceptID"" ID=""CheckConceptIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CHECK_NUMBER_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("CheckNumberMin").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CheckNumberMin")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CheckNumberMin"" ID=""CheckNumberMinHdn"" VALUE=""" & sID & """ />"
		If bFromRequest Then
			sID = oRequest("CheckNumberMax").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CheckNumberMax")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CheckNumberMax"" ID=""CheckNumberMaxHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_HAS_ALIMONY_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("HasAlimony").Item
		Else
			sID = GetParameterFromURLString(oRequest, "HasAlimony")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""HasAlimony"" ID=""HasAlimonyHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_HAS_CREDITS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("HasCredits").Item
		Else
			sID = GetParameterFromURLString(oRequest, "HasCredits")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""HasCredits"" ID=""HasCreditsHdn"" VALUE=""" & sID & """ />"
	End If

	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_NUMBER_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("PaperworkNumber").Item
		Else
			sID = GetParameterFromURLString(oRequest, "PaperworkNumber")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PaperworkNumber"" ID=""PaperworkNumberHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_FOLIO_NUMBER_FLAGS & ",")) > 0) Then
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FilterStartNumber"" ID=""FilterStartNumberHdn"" VALUE=""" & sID & """ />"
        Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FilterEndNumber"" ID=""FilterEndNumberHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_START_DATE_FLAGS & ",")) > 0) Then
		Call DisplaySerialDateAsHidden(GetSerialNumberFromURL("PaperworkStartStart"), "PaperworkStartStart")
		Call DisplaySerialDateAsHidden(GetSerialNumberFromURL("PaperworkStartEnd"), "PaperworkStartEnd")
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_ESTIMATED_DATE_FLAGS & ",")) > 0) Then
		Call DisplaySerialDateAsHidden(GetSerialNumberFromURL("PaperworkEstimatedStart"), "PaperworkEstimatedStart")
		Call DisplaySerialDateAsHidden(GetSerialNumberFromURL("PaperworkEstimatedEnd"), "PaperworkEstimatedEnd")
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_END_DATE_FLAGS & ",")) > 0) Then
		Call DisplaySerialDateAsHidden(GetSerialNumberFromURL("PaperworkEndStart"), "PaperworkEndStart")
		Call DisplaySerialDateAsHidden(GetSerialNumberFromURL("PaperworkEndEnd"), "PaperworkEndEnd")
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_DOCUMENT_NUMBER_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("PpwkDocumentNumber").Item
		Else
			sID = GetParameterFromURLString(oRequest, "PpwkDocumentNumber")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PpwkDocumentNumber"" ID=""PpwkDocumentNumberHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_TYPE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("PaperworkTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "PaperworkTypeID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PaperworkTypeID"" ID=""PaperworkTypeIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_OWNER_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("OwnerNumber").Item
		Else
			sID = GetParameterFromURLString(oRequest, "OwnerNumber")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OwnerNumber"" ID=""OwnerNumberHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_STATUS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("PaperworkStatusID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "PaperworkStatusID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PaperworkStatusID"" ID=""PaperworkStatusIDHdn"" VALUE=""" & sID & """ />"
	End If

    '***** TIPO ASUNTO
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_SUBJECT_TYPES & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("SubjectTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "SubjectTypeID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SubjectTypeID"" ID=""SubjectTypeIDHdn"" VALUE=""" & sID & """ />"
	End If

	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_PRIORITY_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("PriorityID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "PriorityID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PriorityID"" ID=""PriorityID"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_OWNERS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("OwnerIDs").Item
		Else
			sID = GetParameterFromURLString(oRequest, "OwnerIDs")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OwnerIDs"" ID=""OwnerIDsHdn"" VALUE=""" & sID & """ />"
	End If

	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_STATE_TYPE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("StateType").Item
		Else
			sID = GetParameterFromURLString(oRequest, "StateType")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StateType"" ID=""StateTypeHdn"" VALUE=""" & sID & """ />"
	End If

	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_AREA_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetAreaID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetAreaID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetAreaID"" ID=""BudgetAreaIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_COMPANIES_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetCompanyID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetCompanyID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetCompanyID"" ID=""BudgetCompanyIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_PROGRAM_DUTY_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("ProgramDutyID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "ProgramDutyID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProgramDutyID"" ID=""ProgramDutyIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_FUND_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetFundID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetFundID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetFundID"" ID=""BudgetFundIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_DUTY_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetDutyID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetDutyID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetDutyID"" ID=""BudgetDutyIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_ACTIVE_DUTY_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetActiveDutyID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetActiveDutyID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetActiveDutyID"" ID=""BudgetActiveDutyIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_SPECIFIC_DUTY_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetSpecificDutyID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetSpecificDutyID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetSpecificDutyID"" ID=""BudgetSpecificDutyIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_PROGRAM_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetProgramID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetProgramID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetProgramID"" ID=""BudgetProgramIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_REGION_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetRegionID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetRegionID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetRegionID"" ID=""BudgetRegionIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_UR_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetUR").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetUR")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetUR"" ID=""BudgetURHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_CT_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetCT").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetCT")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetCT"" ID=""BudgetCTHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_AUX_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetAUX").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetAUX")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetAUX"" ID=""BudgetAUXHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_LOCATION_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("LocationID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "LocationID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""LocationID"" ID=""LocationIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_BUDGET1_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetID1").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetID1")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetID1"" ID=""BudgetID1Hdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_BUDGET2_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetID2").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetID2")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetID2"" ID=""BudgetID2Hdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_BUDGET3_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetID3").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetID3")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetID3"" ID=""BudgetID3Hdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_CONFINE_TYPE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetConfineTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetConfineTypeID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetConfineTypeID"" ID=""BudgetConfineTypeIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_ACTIVITY1_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetActivityID1").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetActivityID1")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetActivityID1"" ID=""BudgetActivityID1Hdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_ACTIVITY2_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetActivityID2").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetActivityID2")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetActivityID2"" ID=""BudgetActivityID2Hdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_PROCESS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetProcessID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetProcessID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetProcessID"" ID=""BudgetProcessIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_YEAR_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetYear").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetYear")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetYear"" ID=""BudgetYearHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_MONTH_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetMonth").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetMonth")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetMonth"" ID=""BudgetMonthHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_ORIGINAL_POSITION_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetPositionID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetPositionID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetPositionID"" ID=""BudgetPositionIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CREDITS_TYPES_ID_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("CreditTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CreditTypeID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CreditTypeID"" ID=""CreditTypeIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_BENEFICIARY_ID & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BeneficiaryID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BeneficiaryID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BeneficiaryID"" ID=""BeneficiaryIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_CREDITOR_ID & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("CreditorID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CreditorID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CreditorID"" ID=""CreditorIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & S_CREDITS_UPLOADED_FILE_NAME & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("UploadedFileName").Item
		Else
			sID = GetParameterFromURLString(oRequest, "UploadedFileName")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UploadedFileName"" ID=""UploadedFileNameHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_ABSENCE_ID_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AbsenceID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AbsenceID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AbsenceID"" ID=""AbsenceIDHdn"" VALUE=""" & sID & """ />"
		If (sID = 35) Or (sID = 37) Or (sID = 38) Then
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PeriodVacationID"" ID=""PeriodVacationIDHdn"" VALUE=""" & oRequest("PeriodVacationID").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""YearID"" ID=""YearIDHdn"" VALUE=""" & oRequest("YearID").Item & """ />"
		End If
	End If

	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_LOG_DATE_FLAGS & ",")) > 0) Then
		Call DisplaySerialDateAsHidden(GetSerialNumberFromURL("StartLog"), "StartLog")
		Call DisplaySerialDateAsHidden(GetSerialNumberFromURL("EndLog"), "EndLog")
	End If

	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AUDIT_TYPE_ID_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AuditTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AuditTypeID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AuditTypeID"" ID=""AuditTypeIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AUDIT_OPERATION_TYPE_ID_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AuditOperationTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AuditOperationTypeID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AuditOperationTypeID"" ID=""AuditOperationTypeIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_ABSENCE_ACTIVE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("EmployeesAbsenceActive").Item
		Else
			sID = GetParameterFromURLString(oRequest, "EmployeesAbsenceActive")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeesAbsenceActive"" ID=""EmployeesAbsenceActiveHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_ABSENCE_APPLIED_DATE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AppliedDate").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AppliedDate")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AppliedDate"" ID=""AppliedDateHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CONCEPTS_APPLIED_DATE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("RegistrationDate").Item
		Else
			sID = GetParameterFromURLString(oRequest, "RegistrationDate")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RegistrationDate"" ID=""AppliedDateHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CREDITS_APPLIED_DATE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("CreditsAppliedDate").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CreditsAppliedDate")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CreditsAppliedDate"" ID=""CreditsAppliedDateHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_ADJUSTMENT_APPLIED_DATE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AdjustmentPayrollDate").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AdjustmentPayrollDate")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AdjustmentPayrollDate"" ID=""AdjustmentPayrollDateHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EXTRAHOURS_AND_SUNDAYS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AbsenceID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AbsenceID")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AbsenceID"" ID=""AbsenceIDHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CONCEPT_ACTIVE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("EmployeesConceptActive").Item
		Else
			sID = GetParameterFromURLString(oRequest, "EmployeesConceptActive")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeesConceptActive"" ID=""EmployeesConceptActiveHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BANK_ACCOUNTS_ACTIVE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BankAccountsActive").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BankAccountsActive")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BankAccountsActive"" ID=""BankAccountsActiveHdn"" VALUE=""" & sID & """ />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CREDITS_ACTIVE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("CreditsActive").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CreditsActive")
		End If
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CreditsActive"" ID=""CreditsActiveHdn"" VALUE=""" & sID & """ />"
	End If

	DisplayFilterAsHidden = lErrorNumber
	Err.Clear
End Function

Function DisplayFilterInformation(oRequest, sFlags, bForExport, sFilter, sErrorDescription)
'************************************************************
'Purpose: To display the filter information using the user's
'         selections
'Inputs:  oRequest, sFlags, bForExport
'Outputs: sFilter, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayFilterInformation"
	Dim bFromRequest
	Dim sNames
	Dim sID
	Dim sLike
	Dim sMin
	Dim sMax
	Dim sDate
    Dim sFolio
	Dim aTemp
	Dim iIndex
	Dim lErrorNumber

	sID = oRequest("UserID").Item
	bFromRequest = (Err.number = 0)
	Err.clear

	sFilter = ""
	sFlags = "," & sFlags & ","
	If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_PAYROLL_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_OPEN_PAYROLL_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_CLOSED_PAYROLL_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_PAYROLL1_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_CANCELL_PAYROLL_FLAGS & ",")) > 0)) Then
		If bFromRequest Then
			sID = oRequest("PayrollID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "PayrollID")
		End If
		sFilter = sFilter & "<B>Nómina:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Payrolls", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
			aReportTitle(L_PAYROLL_FLAGS) = Replace(sNames, "&nbsp;&nbsp;&nbsp;-", "")
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_MONTHS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("MonthID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "MonthID")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Mes:</B><BR />"
			sFilter = sFilter & asMonthNames_es(CInt(sID)) & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_YEARS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("YearID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "YearID")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Año:</B><BR />"
			sFilter = sFilter & sID & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_DOUBLE_MONTHS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("StartMonthID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "StartMonthID")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Meses:</B><BR />"
			sFilter = sFilter & "De " & asMonthNames_es(CInt(sID))
		End If

		If bFromRequest Then
			sID = oRequest("EndMonthID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "EndMonthID")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & " a " & asMonthNames_es(CInt(sID)) & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_USER_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("UserID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "UserID")
		End If
		sFilter = sFilter & "<B>Usuarios:</B><BR />"
		If Len(sID) > 0 Then
			lErrorNumber = GetNameFromTable(oADODBConnection, "Users", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, sErrorDescription)
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_EMPLOYEE_NUMBER_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_EMPLOYEE_NUMBER1_FLAGS & ",")) > 0)) Then
		If bFromRequest Then
			sID = oRequest("EmployeeIDs").Item
			If Len(sID) = 0 Then
				sID = Replace(Replace(Replace(oRequest("EmployeeNumbers").Item, " ", ""), vbNewLine, ","), ",,", ",")
				Do While (InStr(1, sID, ",,", vbBinaryCompare) > 0)
					sID = Replace(oRequest("EmployeeNumbers").Item, ",,", ",")
				Loop
			End If
			If Len(sID) = 0 Then
				sID = oRequest("EmployeeNumber").Item
				sID = Right(("000000" & sID), Len("000000"))
			End If
		Else
			sID = GetParameterFromURLString(oRequest, "EmployeeIDs")
			If Len(sID) = 0 Then
				sID = Replace(Replace(Replace(GetParameterFromURLString(oRequest, "EmployeeNumbers"), " ", ""), vbNewLine, ","), ",,", ",")
				Do While (InStr(1, sID, ",,", vbBinaryCompare) > 0)
					sID = Replace(sID, ",,", ",")
				Loop
			End If
			If Len(sID) = 0 Then
				sID = GetParameterFromURLString(oRequest, "EmployeeNumber")
				sID = Right(("000000" & sID), Len("000000"))
			End If
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Número de empleado:</B><BR />"
'			sFilter = sFilter & DisplayLikeText(CInt(sLike)) & sID & "<BR />"
			sFilter = sFilter & sID & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_NUMBER7_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("EmployeeTempIDs").Item
			If Len(sID) = 0 Then
				sID = oRequest("EmployeeNumberTemp").Item
				sID = Right(("0000000" & sID), Len("0000000"))
			End If
		Else
			sID = GetParameterFromURLString(oRequest, "EmployeeTempIDs")
			If Len(sID) = 0 Then
				sID = GetParameterFromURLString(oRequest, "EmployeeNumberTemp")
				sID = Right(("0000000" & sID), Len("0000000"))
			End If
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Número de empleado temporal:</B><BR />"
			sFilter = sFilter & sID & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_NAME_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("EmployeeName").Item
			sLike = oRequest("EmployeeNameLike").Item
		Else
			sID = GetParameterFromURLString(oRequest, "EmployeeName")
			sLike = GetParameterFromURLString(oRequest, "EmployeeNameLike")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Nombre del empleado:</B><BR />"
			sFilter = sFilter & DisplayLikeText(CInt(sLike)) & sID & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_JOB_NUMBER_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("JobIDs").Item
			If Len(sID) = 0 Then
				sID = oRequest("JobNumber").Item
				sID = Right(("000000" & sID), Len("000000"))
			End If
			sLike = oRequest("JobNumberLike").Item
		Else
			sID = GetParameterFromURLString(oRequest, "JobIDs")
			If Len(sID) = 0 Then
				sID = GetParameterFromURLString(oRequest, "JobNumber")
				sID = Right(("000000" & sID), Len("000000"))
			End If
			sLike = GetParameterFromURLString(oRequest, "JobNumberLike")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Número de plaza:</B><BR />"
			sFilter = sFilter & DisplayLikeText(CInt(sLike)) & sID & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_COMPANY_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_ONE_COMPANY_FLAGS & ",")) > 0)) Then
		If bFromRequest Then
			sID = oRequest("CompanyID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CompanyID")
		End If
		sFilter = sFilter & "<B>Empresas:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Companies", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
			aReportTitle(L_COMPANY_FLAGS) = Replace(sNames, "&nbsp;&nbsp;&nbsp;-", "")
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
			aReportTitle(L_COMPANY_FLAGS) = "ISSSTE"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AreaID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AreaID")
		End If
		sFilter = sFilter & "<B>Áreas:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Areas", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
			aReportTitle(L_AREA_FLAGS) = Replace(sNames, "&nbsp;&nbsp;&nbsp;-", "")
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
		End If
		If bFromRequest Then
			sID = oRequest("SubAreaID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "SubAreaID")
		End If
		sFilter = sFilter & "<B>Centros de trabajo:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "SubAreas", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_EMPLOYEE_TYPE_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_EMPLOYEE_TYPE1_FLAGS & ",")) > 0)) Then
		If bFromRequest Then
			sID = oRequest("EmployeeTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "EmployeeTypeID")
		End If
		sFilter = sFilter & "<B>Tipos de tabulador:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "EmployeeTypes", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
			aReportTitle(L_EMPLOYEE_TYPE_FLAGS) = Replace(sNames, "&nbsp;&nbsp;&nbsp;-", "")
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_POSITION_TYPE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("PositionTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "PositionTypeID")
		End If
		sFilter = sFilter & "<B>Tipos de puesto:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "PositionTypes", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
			aReportTitle(L_POSITION_TYPE_FLAGS) = Replace(sNames, "&nbsp;&nbsp;&nbsp;-", "")
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CLASSIFICATION_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("ClassificationID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "ClassificationID")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Clasificación:</B><BR />"
			sFilter = sFilter & sID & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_GROUP_GRADE_LEVEL_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("GroupGradeLevelID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "GroupGradeLevelID")
		End If
		sFilter = sFilter & "<B>Grupo, grado, nivel:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "GroupGradeLevels", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_INTEGRATION_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("IntegrationID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "IntegrationID")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Integración:</B><BR />"
			sFilter = sFilter & sID & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_JOURNEY_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("JourneyID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "JourneyID")
		End If
		sFilter = sFilter & "<B>Turnos:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Journeys", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_SHIFT_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("ShiftID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "ShiftID")
		End If
		sFilter = sFilter & "<B>Horarios:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Shifts", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_LEVEL_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("LevelID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "LevelID")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Niveles:</B><BR />"
			sFilter = sFilter & sID & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_STATUS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("EmployeeStatusID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "EmployeeStatusID")
		End If
		sFilter = sFilter & "<B>Estatus de los empleados:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "StatusEmployees", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAYMENT_CENTER_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("PaymentCenterID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "PaymentCenterID")
		End If
		sFilter = sFilter & "<B>Centros de pago:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Areas", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
			aReportTitle(L_PAYMENT_CENTER_FLAGS) = Replace(sNames, "&nbsp;&nbsp;&nbsp;-", "")
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_EMAIL_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("EmployeeEmail").Item
			sLike = oRequest("EmployeeEmailLike").Item
		Else
			sID = GetParameterFromURLString(oRequest, "EmployeeEmail")
			sLike = GetParameterFromURLString(oRequest, "EmployeeEmailLike")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Correo electrónico:</B><BR />"
			sFilter = sFilter & DisplayLikeText(CInt(sLike)) & sID & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_SOCIAL_SECURITY_NUMBER_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("SocialSecurityNumber").Item
			sLike = oRequest("SocialSecurityNumberLike").Item
		Else
			sID = GetParameterFromURLString(oRequest, "SocialSecurityNumber")
			sLike = GetParameterFromURLString(oRequest, "SocialSecurityNumberLike")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Número de seguro social:</B><BR />"
			sFilter = sFilter & DisplayLikeText(CInt(sLike)) & sID & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_BIRTH_FLAGS & ",")) > 0) Then
		lErrorNumber = GetDateRank(oRequest, "StartBirth", "EndBirth", True, sDate)
		If lErrorNumber = 0 Then sFilter = sFilter & "<B>Fechas de nacimiento:</B><BR />&nbsp;&nbsp;&nbsp;" & sDate & "<BR />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_COUNTRY_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("CountryID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CountryID")
		End If
		sFilter = sFilter & "<B>Países:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Countries", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_RFC_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("RFC").Item
			sLike = oRequest("RFCLike").Item
		Else
			sID = GetParameterFromURLString(oRequest, "RFC")
			sLike = GetParameterFromURLString(oRequest, "RFCLike")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>RFC:</B><BR />"
			sFilter = sFilter & DisplayLikeText(CInt(sLike)) & sID & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_CURP_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("CURP").Item
			sLike = oRequest("CURPLike").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CURP")
			sLike = GetParameterFromURLString(oRequest, "CURPLike")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>CURP:</B><BR />"
			sFilter = sFilter & DisplayLikeText(CInt(sLike)) & sID & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_GENDER_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("GenderID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "GenderID")
		End If
		sFilter = sFilter & "<B>Sexo:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Genders", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_ACTIVE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("EmployeeActive").Item
		Else
			sID = GetParameterFromURLString(oRequest, "EmployeeActive")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>¿El empleado está activo?:</B><BR />"
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-" & DisplayYesNo(sID, True) & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_START_DATE_FLAGS & ",")) > 0) Then
		lErrorNumber = GetDateRank(oRequest, "StartEmployeeStart", "EndEmployeeStart", True, sDate)
		If lErrorNumber = 0 Then sFilter = sFilter & "<B>Fecha de ingreso al Instituto:</B><BR />&nbsp;&nbsp;&nbsp;" & sDate & "<BR />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_GENERATING_AREAS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("GeneratingAreaID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "GeneratingAreaID")
		End If
		sFilter = sFilter & "<B>Areas generadoras:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "GeneratingAreas", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_ZONE_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_STATES_FLAGS & ",")) > 0)) Then
		If bFromRequest Then
			sID = oRequest("ZoneID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "ZoneID")
		End If
		If InStr(1, ",1006,1490,", "," & oRequest("ReportID").Item & ",", vbBinaryCompare) > 0 Then
			sFilter = sFilter & "<B>Entidades federativas (centro de trabajo):</B><BR />"
		Else
			sFilter = sFilter & "<B>Entidades federativas (centro de pago):</B><BR />"
		End If
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Zones", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
			If InStr(1, "," & Replace(sID, " ", "") & ",", ",38,", vbbinaryCompare) > 0 Then sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-HOSP. REG. PDTE. JUAREZ OAXACA, OAX.<BR />"
			aReportTitle(L_ZONE_FLAGS) = Replace(sNames, "&nbsp;&nbsp;&nbsp;-", "")
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
			aReportTitle(L_ZONE_FLAGS) = "FORANEO Y LOCAL"
		End If
	End If
	If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_ZONE_FOR_PAYMENT_CENTER_FLAGS & ",")) > 0)) Then
		If bFromRequest Then
			sID = oRequest("ZoneForPaymentCenterID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "ZoneForPaymentCenterID")
		End If
		sFilter = sFilter & "<B>Entidades federativas (centro de pago):</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Zones", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
			If InStr(1, "," & Replace(sID, " ", "") & ",", ",38,", vbbinaryCompare) > 0 Then sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-HOSP. REG. PDTE. JUAREZ OAXACA, OAX.<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_POSITION_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("PositionID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "PositionID")
		End If
		sFilter = sFilter & "<B>Puestos:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Positions", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
			aReportTitle(L_POSITION_FLAGS) = Replace(sNames, "&nbsp;&nbsp;&nbsp;-", "")
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_JOB_TYPE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("JobTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "JobTypeID")
		End If
		sFilter = sFilter & "<B>Tipos de plaza:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "JobTypes", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
			aReportTitle(L_JOB_TYPE_FLAGS) = Replace(sNames, "&nbsp;&nbsp;&nbsp;-", "")
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_OCCUPATION_TYPE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("OccupationTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "OccupationTypeID")
		End If
		sFilter = sFilter & "<B>Tipos de ocupación:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "OccupationTypes", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
			aReportTitle(L_OCCUPATION_TYPE_FLAGS) = Replace(sNames, "&nbsp;&nbsp;&nbsp;-", "")
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_JOB_START_DATE_FLAGS & ",")) > 0) Then
		lErrorNumber = GetDateRank(oRequest, "StartJobStart", "EndJobStart", True, sDate)
		If lErrorNumber = 0 Then sFilter = sFilter & "<B>Fecha de inicio de la plaza:</B><BR />&nbsp;&nbsp;&nbsp;" & sDate & "<BR />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_JOB_END_DATE_FLAGS & ",")) > 0) Then
		lErrorNumber = GetDateRank(oRequest, "StartJobEnd", "EndJobEnd", True, sDate)
		If lErrorNumber = 0 Then sFilter = sFilter & "<B>Fecha de término de la plaza:</B><BR />&nbsp;&nbsp;&nbsp;" & sDate & "<BR />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_JOB_STATUS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("JobStatusID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "JobStatusID")
		End If
		sFilter = sFilter & "<B>Estatus de plaza:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "StatusJobs", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_JOB_ACTIVE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("JobActive").Item
		Else
			sID = GetParameterFromURLString(oRequest, "JobActive")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>¿La plaza está activa?:</B><BR />"
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-" & DisplayYesNo(sID, True) & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_CODE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AreaCode").Item
			sLike = oRequest("AreaCodeLike").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AreaCode")
			sLike = GetParameterFromURLString(oRequest, "AreaCodeLike")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Código del centro de trabajo:</B><BR />"
			sFilter = sFilter & DisplayLikeText(CInt(sLike)) & sID & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_SHORT_NAME_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AreaShortName").Item
			sLike = oRequest("AreaShortNameLike").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AreaShortName")
			sLike = GetParameterFromURLString(oRequest, "AreaShortNameLike")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Clave del centro de trabajo:</B><BR />"
			sFilter = sFilter & DisplayLikeText(CInt(sLike)) & sID & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_NAME_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AreaName").Item
			sLike = oRequest("AreaNameLike").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AreaName")
			sLike = GetParameterFromURLString(oRequest, "AreaNameLike")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Nombre del centro de trabajo:</B><BR />"
			sFilter = sFilter & DisplayLikeText(CInt(sLike)) & sID & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_TYPE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AreaTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AreaTypeID")
		End If
		sFilter = sFilter & "<B>Tipo de área:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "AreaTypes", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CONFINE_TYPE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("ConfineTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "ConfineTypeID")
		End If
		sFilter = sFilter & "<B>Tipo de ámbito para las áreas:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "ConfineTypes", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CENTER_TYPE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("CenterTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CenterTypeID")
		End If
		sFilter = sFilter & "<B>Tipo de centro de trabajo:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "CenterTypes", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CENTER_SUBTYPE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("CenterSubtypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CenterSubtypeID")
		End If
		sFilter = sFilter & "<B>Subtipo de centro de trabajo:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "CenterSubtypes", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_ATTENTION_LEVEL_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AttentionLevelID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AttentionLevelID")
		End If
		sFilter = sFilter & "<B>Nivel de atención:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "AttentionLevels", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_ECONOMIC_ZONE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("EconomicZoneID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "EconomicZoneID")
		End If
		sFilter = sFilter & "<B>Zona económica:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "EconomicZones", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_START_DATE_FLAGS & ",")) > 0) Then
		lErrorNumber = GetDateRank(oRequest, "StartAreaStart", "EndAreaStart", True, sDate)
		If lErrorNumber = 0 Then sFilter = sFilter & "<B>Fecha de inicio del centro de trabajo:</B><BR />&nbsp;&nbsp;&nbsp;" & sDate & "<BR />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_END_DATE_FLAGS & ",")) > 0) Then
		lErrorNumber = GetDateRank(oRequest, "StartAreaEnd", "EndAreaEnd", True, sDate)
		If lErrorNumber = 0 Then sFilter = sFilter & "<B>Fecha de término del centro de trabajo:</B><BR />&nbsp;&nbsp;&nbsp;" & sDate & "<BR />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_JOBS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("Jobs").Item
		Else
			sID = GetParameterFromURLString(oRequest, "Jobs")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Plazas:</B><BR />"
			sFilter = sFilter & sID & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_TOTAL_JOBS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("TotalJobs").Item
		Else
			sID = GetParameterFromURLString(oRequest, "TotalJobs")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Total de plazas:</B><BR />"
			sFilter = sFilter & sID & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_STATUS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AreaStatusID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AreaStatusID")
		End If
		sFilter = sFilter & "<B>Estatus del centro de trabajo:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "StatusAreas", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CONCEPTS_VALUES_STATUS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("ConceptStatusID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "ConceptStatusID")
		End If
		sFilter = sFilter & "<B>Estatus del tabulador:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "StatusConceptsValues", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
			aReportTitle(L_CONCEPTS_VALUES_STATUS_FLAGS) = sID
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
			aReportTitle(L_CONCEPTS_VALUES_STATUS_FLAGS) = 1
		End If
	End If

    If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CONCEPT_2_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("ConceptID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "ConceptID")
		End If
		sFilter = sFilter & "<B>Conceptos de Pago:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Concepts", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
			aReportTitle(L_CONCEPT_2_FLAGS) = sID
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
			aReportTitle(L_CONCEPT_2_FLAGS) = 1
		End If
	End If
   
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_REASON_ID_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("ReasonID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "ReasonID")
		End If
		sFilter = sFilter & "<B>Tipo de movimiento:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Reasons", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
			aReportTitle(L_EMPLOYEE_REASON_ID_FLAGS) = Replace(sNames, "&nbsp;&nbsp;&nbsp;-", "")
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_ACTIVE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AreaActive").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AreaActive")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>¿El centro de trabajo está activo?:</B><BR />"
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-" & DisplayYesNo(sID, True) & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_CONCEPT_ID_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_CONCEPT_1_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_THIRD_CONCEPTS_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_THIRD_CONCEPTS2_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_MEMORY_CONCEPT_ID_FLAGS & ",")) > 0)) Then
		If bFromRequest Then
			sID = oRequest("ConceptID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "ConceptID")
		End If
		sFilter = sFilter & "<B>Concepto de pago:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Concepts", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			If InStr(1, sID, "6364", vbBinaryCompare) > 0 Then
				If Len(sNames) = 0 Then
					sNames = sNames & "&nbsp;&nbsp;&nbsp;-63 Y 64 SEGURO METLIFE"
				Else
					sNames = sNames & "<BR />&nbsp;&nbsp;&nbsp;-63 Y 64 SEGURO METLIFE"
				End If
			End If
			sFilter = sFilter & sNames & "<BR />"
			aReportTitle(L_CONCEPT_ID_FLAGS) = Replace(sNames, "&nbsp;&nbsp;&nbsp;-", "")
			Call GetNameFromTable(oADODBConnection, "ShortConcepts", sID, "", "<BR />", sNames, "")
			aReportTitle(L_CONCEPT_ID_FLAGS) = sNames & ";" & aReportTitle(L_CONCEPT_ID_FLAGS) 
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_TOTAL_PAYMENT_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("TotalPaymentMin").Item
		Else
			sID = GetParameterFromURLString(oRequest, "TotalPaymentMin")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Líquidos mayores o iguales a: </B>" & sID & "<BR />"
		End If
		If bFromRequest Then
			sID = oRequest("TotalPaymentMax").Item
		Else
			sID = GetParameterFromURLString(oRequest, "TotalPaymentMax")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Líquidos menores o iguales a: </B>" & sID & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_BANK_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_ONE_BANK_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_ISSSTE_ONE_BANK_FLAGS & ",")) > 0)) Then
		If bFromRequest Then
			sID = oRequest("BankID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BankID")
		End If
		sFilter = sFilter & "<B>Banco:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Banks", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
			aReportTitle(L_BANK_FLAGS) = Replace(sNames, "&nbsp;&nbsp;&nbsp;-", "")
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_MEDICAL_AREAS_TYPES_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("MedicalAreasTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "MedicalAreasTypeID")
		End If
		sFilter = sFilter & "<B>Tipo reporte:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "MedicalAreasTypes", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
			aReportTitle(L_MEDICAL_AREAS_TYPES_FLAGS) = Replace(sNames, "&nbsp;&nbsp;&nbsp;-", "")
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_DOCUMENT_FOR_LICENSE_NUMBER_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("DocumentForLicenseNumber").Item
			sLike = oRequest("DocumentForLicenseNumberLike").Item
		Else
			sID = GetParameterFromURLString(oRequest, "DocumentForLicenseNumber")
			sLike = GetParameterFromURLString(oRequest, "DocumentForLicenseNumberLike")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>No. de folio:</B><BR />"
			sFilter = sFilter & DisplayLikeText(CInt(sLike)) & sID & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_DOCUMENT_REQUEST_NUMBER_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("RequestNumber").Item
			sLike = oRequest("RequestNumberLike").Item
		Else
			sID = GetParameterFromURLString(oRequest, "RequestNumber")
			sLike = GetParameterFromURLString(oRequest, "RequestNumberLike")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>No. de solicitud:</B><BR />"
			sFilter = sFilter & DisplayLikeText(CInt(sLike)) & sID & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_DOCUMENT_FOR_CANCEL_LICENSE_NUMBER_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("DocumentForCancelLicenseNumber").Item
			sLike = oRequest("DocumentForCancelLicenseNumberLike").Item
		Else
			sID = GetParameterFromURLString(oRequest, "DocumentForCancelLicenseNumber")
			sLike = GetParameterFromURLString(oRequest, "DocumentForCancelLicenseNumberLike")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>No. de oficio de cancelación:</B><BR />"
			sFilter = sFilter & DisplayLikeText(CInt(sLike)) & sID & "<BR />"
		End If
	End If

	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_DATE_FLAGS & ",")) > 0) Then
		lErrorNumber = GetDateRank(oRequest, "Start", "End", True, sDate)
		If lErrorNumber = 0 Then sFilter = sFilter & "<B>Periodo:</B><BR />&nbsp;&nbsp;&nbsp;" & sDate & "<BR />"
	End If
	If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_CHECK_CONCEPTS_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_CHECK_CONCEPTS_ALL_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_CHECK_CONCEPTS_EMPLOYEES_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_ONLY_CHECK_CONCEPTS_EMPLOYEES_FLAGS & ",")) > 0)) Then
		If bFromRequest Then
			sID = oRequest("CheckConceptID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CheckConceptID")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Pagos de:</B><BR />"
			Select Case CLng(sID)
				Case -1
					sFilter = sFilter & "Empleados con cheque<BR />"
				Case -2
					sFilter = sFilter & "Empleados con depósito<BR />"
				Case 0
					sFilter = sFilter & "Empleados con cheque y depósitos<BR />"
				Case 11
					sFilter = sFilter & "Honorarios<BR />"
				Case 69
					sFilter = sFilter & "Pensión alimenticia<BR />"
				Case 155
					sFilter = sFilter & "Acreedores<BR />"
			End Select
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CHECK_NUMBER_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("CheckNumberMin").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CheckNumberMin")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Folios mayores o iguales a: </B>" & sID & "<BR />"
		End If
		If bFromRequest Then
			sID = oRequest("CheckNumberMax").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CheckNumberMax")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Folios menores o iguales a: </B>" & sID & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_HAS_ALIMONY_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("HasAlimony").Item
		Else
			sID = GetParameterFromURLString(oRequest, "HasAlimony")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Empleados con pensión alimenticia:</B><BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_HAS_CREDITS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("HasCredits").Item
		Else
			sID = GetParameterFromURLString(oRequest, "HasCredits")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Empleados con productos de terceros:</B><BR />"
		End If
	End If

	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_NUMBER_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("PaperworkNumber").Item
'			sLike = oRequest("PaperworkNumberLike").Item
		Else
			sID = GetParameterFromURLString(oRequest, "PaperworkNumber")
'			sLike = GetParameterFromURLString(oRequest, "PaperworkNumberLike")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Número de trámite:</B><BR />"
			sFilter = sFilter & sID & "<BR />"
'			sFilter = sFilter & DisplayLikeText(CInt(sLike)) & sID & "<BR />"
		End If
	End If

	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_FOLIO_NUMBER_FLAGS & ",")) > 0) Then
		lErrorNumber = GetFolioRank(oRequest, bFromRequest, sFolio)
		If lErrorNumber = 0 Then sFilter = sFilter & "<B>Número de folio:</B><BR />&nbsp;&nbsp;&nbsp;" & sFolio & "<BR />"
    End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_START_DATE_FLAGS & ",")) > 0) Then
		lErrorNumber = GetDateRank(oRequest, "PaperworkStartStart", "PaperworkStartEnd", True, sDate)
		If lErrorNumber = 0 Then sFilter = sFilter & "<B>Fecha de recepción:</B><BR />&nbsp;&nbsp;&nbsp;" & sDate & "<BR />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_ESTIMATED_DATE_FLAGS & ",")) > 0) Then
		lErrorNumber = GetDateRank(oRequest, "PaperworkEstimatedStart", "PaperworkEstimatedEnd", True, sDate)
		If lErrorNumber = 0 Then sFilter = sFilter & "<B>Fecha de límite de respuesta:</B><BR />&nbsp;&nbsp;&nbsp;" & sDate & "<BR />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_END_DATE_FLAGS & ",")) > 0) Then
		lErrorNumber = GetDateRank(oRequest, "PaperworkEndStart", "PaperworkEndEnd", True, sDate)
		If lErrorNumber = 0 Then sFilter = sFilter & "<B>Fecha de atención:</B><BR />&nbsp;&nbsp;&nbsp;" & sDate & "<BR />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_DOCUMENT_NUMBER_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("PpwkDocumentNumber").Item
			sLike = oRequest("PpwkDocumentNumberLike").Item
		Else
			sID = GetParameterFromURLString(oRequest, "PpwkDocumentNumber")
			sLike = GetParameterFromURLString(oRequest, "PpwkDocumentNumberLike")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Número de documento:</B><BR />"
			sFilter = sFilter & DisplayLikeText(CInt(sLike)) & sID & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_TYPE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("PaperworkTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "PaperworkTypeID")
		End If
		sFilter = sFilter & "<B>Tipo de trámite:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "PaperworkTypes", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_OWNER_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("OwnerNumber").Item
			sLike = oRequest("OwnerNumberLike").Item
		Else
			sID = GetParameterFromURLString(oRequest, "OwnerNumber")
			sLike = GetParameterFromURLString(oRequest, "OwnerNumberLike")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Número de empleado del responsable:</B><BR />"
			sFilter = sFilter & DisplayLikeText(CInt(sLike)) & sID & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_STATUS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("PaperworkStatusID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "PaperworkStatusID")
		End If
		sFilter = sFilter & "<B>Estatus de trámite:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "StatusPaperworks", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_SUBJECT_TYPES & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("SubjectTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "SubjectTypeID")
		End If
		sFilter = sFilter & "<B>Tipo de asunto:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "SubjectTypes", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_PRIORITY_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("PriorityID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "PriorityID")
		End If
		sFilter = sFilter & "<B>Prioridad:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Priorities", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_OWNERS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("OwnerIDs").Item
		Else
			sID = GetParameterFromURLString(oRequest, "OwnerIDs")
		End If
		sFilter = sFilter & "<B>Responsables:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "PaperworkOwners", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If

	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_STATE_TYPE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("StateType").Item
		Else
			sID = GetParameterFromURLString(oRequest, "StateType")
		End If
		Select Case sID
			Case "0"
				sFilter = sFilter & "<B>Foráneo</B><BR />"
			Case "1"
				sFilter = sFilter & "<B>Local</B><BR />"
			Case Else
				sFilter = sFilter & "<B>Local y foráneo</B><BR />"
		End Select
	End If

	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_COURSE_NAME_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("CourseID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CourseID")
		End If
		sFilter = sFilter & "<B>Curso:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "SADE_Curso", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_COURSE_DIPLOMA_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("ProfileID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "ProfileID")
		End If
		sFilter = sFilter & "<B>Diplomado:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "SADE_Perfiles", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If

	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_AREA_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetAreaID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetAreaID")
		End If
		sFilter = sFilter & "<B>Área:</B><BR />"
		If Len(sID) > 0 Then
			sFilter = sFilter & sID & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_COMPANIES_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetCompanyID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetCompanyID")
		End If
		sFilter = sFilter & "<B>Empresas:</B><BR />"
		If Len(sID) > 0 Then
			If InStr(1, sID, "-1", vbBinaryCompare) > 0 Then sFilter = sFilter & "ISSSTE ASEGURADOR<BR />"
			If InStr(1, sID, "170", vbBinaryCompare) > 0 Then sFilter = sFilter & "PENSIONISSSTE<BR />"
			If InStr(1, sID, "500", vbBinaryCompare) > 0 Then sFilter = sFilter & "SUPERISSSTE<BR />"
			If InStr(1, sID, "700", vbBinaryCompare) > 0 Then sFilter = sFilter & "FOVISSSTE<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_PROGRAM_DUTY_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("ProgramDutyID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "ProgramDutyID")
		End If
		sFilter = sFilter & "<B>Programa presupuestario:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "BudgetsProgramDuties", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_FUND_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetFundID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetFundID")
		End If
		sFilter = sFilter & "<B>Fondo:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "BudgetsFunds", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_DUTY_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetDutyID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetDutyID")
		End If
		sFilter = sFilter & "<B>Función:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "BudgetsDuties", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_ACTIVE_DUTY_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetActiveDutyID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetActiveDutyID")
		End If
		sFilter = sFilter & "<B>Subfunción activa:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "BudgetsActiveDuties", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_SPECIFIC_DUTY_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetSpecificDutyID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetSpecificDutyID")
		End If
		sFilter = sFilter & "<B>Subfunción específica:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "BudgetsSpecificDuties", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_PROGRAM_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetProgramID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetProgramID")
		End If
		sFilter = sFilter & "<B>Programa:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "BudgetsPrograms", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_REGION_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetRegionID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetRegionID")
		End If
		sFilter = sFilter & "<B>Región:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Zones", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_UR_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetUR").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetUR")
		End If
		sFilter = sFilter & "<B>UR:</B><BR />"
		If Len(sID) > 0 Then
			sFilter = sFilter & sID & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_CT_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetCT").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetCT")
		End If
		sFilter = sFilter & "<B>CT:</B><BR />"
		If Len(sID) > 0 Then
			sFilter = sFilter & sID & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_AUX_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetAUX").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetAUX")
		End If
		sFilter = sFilter & "<B>AUX:</B><BR />"
		If Len(sID) > 0 Then
			sFilter = sFilter & sID & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_LOCATION_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("LocationID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "LocationID")
		End If
		sFilter = sFilter & "<B>Municipio:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Zones", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_BUDGET1_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetID1").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetID1")
		End If
		sFilter = sFilter & "<B>Partida:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Budgets", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_BUDGET2_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetID2").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetID2")
		End If
		sFilter = sFilter & "<B>Subpartida:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Budgets", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_BUDGET3_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetID3").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetID3")
		End If
		sFilter = sFilter & "<B>Tipo de pago:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Budgets", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_CONFINE_TYPE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetConfineTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetConfineTypeID")
		End If
		sFilter = sFilter & "<B>Ámbito:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "BudgetsConfineTypes", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_ACTIVITY1_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetActivityID1").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetActivityID1")
		End If
		sFilter = sFilter & "<B>Actividad institucional:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "BudgetsActivities1", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_ACTIVITY2_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetActivityID2").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetActivityID2")
		End If
		sFilter = sFilter & "<B>Actividad presupuestaria:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "BudgetsActivities2", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_PROCESS_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetProcessID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetProcessID")
		End If
		sFilter = sFilter & "<B>Proceso:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "BudgetsProcesses", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_YEAR_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetYear").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetYear")
		End If
		sFilter = sFilter & "<B>Año:</B><BR />"
		If Len(sID) > 0 Then
			sFilter = sFilter & sID & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_MONTH_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetMonth").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetMonth")
		End If
		sFilter = sFilter & "<B>Mes:</B><BR />"
		If Len(sID) > 0 Then
			sFilter = sFilter & asMonthNames_es(CInt(sID)) & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_ORIGINAL_POSITION_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BudgetPositionID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BudgetPositionID")
		End If
		sFilter = sFilter & "<B>Puestos:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "BudgetsPositions", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CREDITS_TYPES_ID_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("CreditTypeID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CreditTypeID")
		End If
		sFilter = sFilter & "<B>Tipos de crédito:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "CreditTypes", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_BENEFICIARY_ID & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BeneficiaryID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BeneficiaryID")
		End If
		sFilter = sFilter & "<B>Beneficiarios de pensión alimenticia:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "BeneficiaryID", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_CREDITOR_ID & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("CreditorID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CreditorID")
		End If
		sFilter = sFilter & "<B>Acreedores:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "CreditorID", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & S_CREDITS_UPLOADED_FILE_NAME & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("UploadedFileName").Item
		Else
			sID = GetParameterFromURLString(oRequest, "UploadedFileName")
		End If
		sFilter = sFilter & "<B>Registros cargados desde archivos de terceros:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "CreditsFiles", "'" & sID & "'", "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_ABSENCE_ID_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AbsenceID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AbsenceID")
		End If
		sFilter = sFilter & "<B>Tipos de incidencias:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Absences", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
			If (sID = 35) Or (sID = 37) Or (sID = 38) Then
				sNames = "&nbsp;&nbsp;&nbsp;Año " & oRequest("YearID").Item & " Periodo " & oRequest("PeriodVacationID").Item
				sFilter = sFilter & sNames & "<BR />"
			End If
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_ABSENCE_ACTIVE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("EmployeesAbsenceActive").Item
		Else
			sID = GetParameterFromURLString(oRequest, "EmployeesAbsenceActive")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>¿La incidencia está aplicada?:</B><BR />"
			Select Case sID
				Case 0
					sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-NO<BR />"
				Case 1
					sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-SI<BR />"
				Case Else
					sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Cancelada<BR />"
			End Select
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_ABSENCE_APPLIED_DATE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AppliedDate").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AppliedDate")
		End If
		sFilter = sFilter & "<B>Quincena de aplicación de la incidencia:</B><BR />"
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>Quincena de aplicación de la incidencia</B><BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CONCEPTS_APPLIED_DATE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("RegistrationDate").Item
		Else
			sID = GetParameterFromURLString(oRequest, "RegistrationDate")
		End If
		sFilter = sFilter & "<B>Quincena de aplicación:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Payrolls", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CREDITS_APPLIED_DATE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("CreditsAppliedDate").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CreditsAppliedDate")
		End If
		sFilter = sFilter & "<B>Quincena de aplicación del crédito:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Payrolls", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_ADJUSTMENT_APPLIED_DATE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("CreditsAppliedDate").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CreditsAppliedDate")
		End If
		sFilter = sFilter & "<B>Quincena de aplicación del crédito:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Payrolls", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todas<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EXTRAHOURS_AND_SUNDAYS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("AbsenceID").Item
		Else
			sID = GetParameterFromURLString(oRequest, "AbsenceID")
		End If
		sFilter = sFilter & "<B>Tipo de concepto:</B><BR />"
		If Len(sID) > 0 Then
			Call GetNameFromTable(oADODBConnection, "Absences", sID, "&nbsp;&nbsp;&nbsp;-", "<BR />", sNames, "")
			sFilter = sFilter & sNames & "<BR />"
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Ambos<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CONCEPT_ACTIVE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("EmployeesConceptActive").Item
		Else
			sID = GetParameterFromURLString(oRequest, "EmployeesConceptActive")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>¿El concepto está aplicado?:</B><BR />"
			'sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-" & DisplayYesNo(sID, True) & "<BR />"
			Select Case CInt(sID)
				Case 0
					sNames = "En proceso"
				Case 1
					sNames = "Activo"
				Case 2
					sNames = "Cancelado"
			End Select
			sFilter = sFilter & sNames & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_LOG_DATE_FLAGS & ",")) > 0) Then
		lErrorNumber = GetDateRank(oRequest, "StartLog", "EndLog", True, sDate)
		If lErrorNumber = 0 Then sFilter = sFilter & "<B>Fecha de entrada al sistema:</B><BR />&nbsp;&nbsp;&nbsp;" & sDate & "<BR />"
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BANK_ACCOUNTS_ACTIVE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("BankAccountsActive").Item
		Else
			sID = GetParameterFromURLString(oRequest, "BankAccountsActive")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>¿La cuenta bancaria está aplicada?:</B><BR />"
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-" & DisplayYesNo(sID, True) & "<BR />"
		End If
	End If
	If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CREDITS_ACTIVE_FLAGS & ",")) > 0) Then
		If bFromRequest Then
			sID = oRequest("CreditsActive").Item
		Else
			sID = GetParameterFromURLString(oRequest, "CreditsActive")
		End If
		If Len(sID) > 0 Then
			sFilter = sFilter & "<B>¿El crédito está aplicado?:</B><BR />"
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-" & DisplayYesNo(sID, True) & "<BR />"
		End If
	End If

	If (lErrorNumber = 0) And (aReportsComponent(N_ID_REPORTS) = ISSSTE_1203_REPORTS) Then
		lErrorNumber = GetDateRank(oRequest, "DocumentStart", "DocumentEnd", True, sDate)
		If lErrorNumber = 0 Then sFilter = sFilter & "<B>Fecha de registro de la solicitud de la(s) hoja(s) única(s) de servicio:</B><BR />&nbsp;&nbsp;&nbsp;" & sDate
	End If

	If InStr(1, oRequest, "ShowFilter=False", vbBinaryCompare) = 0 Then
		If bForExport Then
			Call DisplayErrorMessageInPlainText("", "<B>Este reporte incluye lo siguiente:</B><BR />" & sFilter, "<BR />")
			If InStr(1, sFlags, ("," & L_DONT_CLOSE_FILTER_DIV_FLAGS & ",")) = 0 Then Response.Write "<BR />"
		Else
			If StrComp(GetASPFileName(""), "Payroll.asp", vbBinaryCompare) = 0 Then
				Response.Write "<DIV CLASS=""ReportFilter"">"
					Call DisplayErrorMessageInPlainText("", "<B>La prenómina fue calculada con el siguiente filtro:</B><BR /><BR />" & sFilter, "<BR />")
				Response.Write "</DIV>"
			Else
				Response.Write "<DIV CLASS=""ReportFilter"" STYLE=""height: 200px;"">"
					Call DisplayErrorMessageInPlainText("", "<B>Este reporte incluye lo siguiente:</B><BR /><BR />" & sFilter, "<BR />")
				If InStr(1, sFlags, ("," & L_DONT_CLOSE_FILTER_DIV_FLAGS & ",")) = 0 Then Response.Write "</DIV>"
			End If
		End If
		If InStr(1, sFlags, ("," & L_DONT_CLOSE_FILTER_DIV_FLAGS & ",")) = 0 Then Response.Write "<BR />"
	End If
	

	DisplayFilterInformation = lErrorNumber
	Err.Clear
End Function

Function GetFolioRank(oRequest, bFromRequest, sFilter)
	Dim sFolioStart
    Dim sFolioEnd

	If bFromRequest Then
		If (Len(oRequest("FilterStartNumber").Item) > 0 And Len(oRequest("FilterEndNumber").Item) > 0) Then
			sFilter = sFilter & "Entre " & CLng(oRequest("FilterStartNumber").Item) & " y el " & CLng(oRequest("FilterEndNumber").Item)
		ElseIf (Len(oRequest("FilterStartNumber").Item) > 0 And Len(oRequest("FilterEndNumber").Item) = 0) Then
			sFilter = sFilter & "Mayores a " & CLng(oRequest("FilterStartNumber").Item)
		ElseIf (Len(oRequest("FilterStartNumber").Item) = 0 And Len(oRequest("FilterEndNumber").Item) > 0) Then
			sFilter = sFilter & "Menores a " & CLng(oRequest("FilterEndNumber").Item)
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos"
		End If
	Else
		sFolioStart = GetParameterFromURLString(oRequest, "FilterStartNumber")
		sFolioEnd = GetParameterFromURLString(oRequest, "FilterEndNumber")
		If (Len(sFolioStart) > 0 And Len(sFolioEnd) > 0) Then
			sFilter = sFilter & "Entre " & CLng(sFolioStart) & " y el " & CLng(sFolioEnd)
		ElseIf (Len(sFolioStart) > 0 And Len(sFolioEnd) = 0) Then
			sFilter = sFilter & "Mayores a " & CLng(sFolioStart)
		ElseIf (Len(sFolioStart) = 0 And Len(sFolioEnd) > 0) Then
			sFilter = sFilter & "Menores a " & CLng(sFolioEnd)
		Else
			sFilter = sFilter & "&nbsp;&nbsp;&nbsp;-Todos"
		End If
    End If
    GetFolioRank = lErrorNumber
    Err.Clear
End Function

Function DisplayReportEditionDivsString(sDivIDs)
'************************************************************
'Purpose: To display the JavaScript code used by the report
'         edition feature
'Inputs:  sDivIDs
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayReportEditionDivsString"

	If Len(sDivIDs) > 0 Then
		Response.Write "<DONT_EXPORT><SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "if (sEditElements == '')" & vbNewLine
				Response.Write "sEditElements='" & Left(sDivIDs, (Len(sDivIDs) - Len(","))) & "';" & vbNewLine
			Response.Write "else" & vbNewLine
				Response.Write "sEditElements+='," & Left(sDivIDs, (Len(sDivIDs) - Len(","))) & "';" & vbNewLine
		Response.Write "//--></SCRIPT></DONT_EXPORT>" & vbNewLine
	End If
	DisplayReportEditionDivsString = Err.number
	Err.Clear
End Function

Function DisplayReportEditionLink(bAddNewLine, sDivIDs)
'************************************************************
'Purpose: To display the Report Edition Div and the link to
'         call the HTML Composer
'Inputs:  bAddNewLine, sDivIDs
'Outputs: sDivIDs
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayReportEditionLink"
	Dim sDivID

	sDivID = GetSerialNumberForDate("") & GenerateRandomNumbersSecuence(10) & "Div"
	sDivIDs = sDivIDs & sDivID & ","
	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><DIV ID=""" & sDivID & """ STYLE=""display: none""><BR />"
		Response.Write "<DIV ID=""Edit" & sDivID & """></DIV>"
		Response.Write "<DONT_EXPORT><FONT FACE=""Arial"" SIZE=""2""><A EDITOR=""1"" HREF=""javascript: OpenNewWindow('HTMLComposer.asp?TargetName=Edit" & sDivID & "', '', 'HTMLComposer', 510, 300, 'no', 'no')""><IMG SRC=""Images/IcnNote.gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=""0"" ALT=""Agregar comentarios"" />Agregar comentarios</A></FONT></DONT_EXPORT>"
		If bAddNewLine Then Response.Write "<DONT_EXPORT><BR /><BR /></DONT_EXPORT>"
	Response.Write "</DIV></FONT>"
	

	DisplayReportEditionLink = Err.number
	Err.Clear
End Function

Function DisplayReportFilter(sFlags, sErrorDescription)
'************************************************************
'Purpose: To display the HTML Form to filter the report
'Inputs:  sFlags
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayReportFilter"
	Dim iCounter
	Dim iIndex
	Dim oItem
	Dim sCondition
	Dim lErrorNumber

	sCondition = ""
	If StrComp(SERVER_NAME_FOR_LICENSE, "CASTOR", vbBinaryCompare) = 0 Then sCondition = " Top 55 "

	sFlags = "," & sFlags & ","
	iCounter = 1
	If InStr(1, sFlags, ("," & L_ZIP_WARNING_FLAGS & ",")) > 0 Then Response.Write "<DIV NAME=""ReportFilterDiv"" ID=""ReportFilterDiv"">"
		If InStr(1, sFlags, ("," & L_NO_INSTRUCTIONS_FLAGS & ",")) = 0 Then Call DisplayInstructionsMessage("Filtro", "Para filtrar la información del reporte, seleccione aquellos registros que acotarán la información. Puede seleccionar varios campos a la vez utilizando las teclas Shift y Control.<BR /><BR />Cuando el filtro esté listo, presione el botón para continuar.")
		Response.Write "<BR />"
		Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""DisplayOptionsIFrame"" FRAMEBORDER=""0"" WIDTH=""320"" HEIGHT=""0""></IFRAME><BR />"
		If InStr(1, sFlags, ("," & L_NO_DIV_FLAGS & ",")) = 0 Then Response.Write "<DIV CLASS=""ReportFilter"">"
			If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_PAYROLL_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_OPEN_PAYROLL_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_CLOSED_PAYROLL_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_PAYROLL1_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_ORDINARY_PAYROLL_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_CANCELL_PAYROLL_FLAGS & ",")) > 0)) Then
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "var aPayrolls = new Array("
						Call GenerateJavaScriptArrayFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollTypeID", "Length(PayrollID)=8 And (IsClosed<>1)", "PayrollID", sErrorDescription)
					Response.Write "['-1', '-1']);" & vbNewLine

					Response.Write "function DisplayPayrollFilters(sPayrollID) {" & vbNewLine
						If StrComp(GetASPFileName(""), "Payroll.asp", vbBinaryCompare) = 0 Then
							Response.Write "HideDisplay(document.all['ConceptsDiv']);" & vbNewLine
							Response.Write "HideDisplay(document.all['PeriodsDiv']);" & vbNewLine
							Response.Write "HideDisplay(document.all['DeleteDiv']);" & vbNewLine
							Response.Write "for (var i=0; i<aPayrolls.length; i++) {" & vbNewLine
								Response.Write "if (aPayrolls[i][0] == sPayrollID) {" & vbNewLine
									Response.Write "switch (aPayrolls[i][1]) {" & vbNewLine
										Response.Write "case '3':" & vbNewLine
											Response.Write "ShowDisplay(document.all['ConceptsDiv']);" & vbNewLine
											Response.Write "break;" & vbNewLine
										Response.Write "case '4':" & vbNewLine
											Response.Write "ShowDisplay(document.all['PeriodsDiv']);" & vbNewLine
											Response.Write "ShowDisplay(document.all['DeleteDiv']);" & vbNewLine
											If (Len(oRequest("StartPayrollYear").Item) = 0) And (Len(oRequest("StartPayrollMonth").Item) = 0) And (Len(oRequest("StartPayrollDay").Item) = 0)Then
												Response.Write "document.ReportFrm.StartPayrollYear.value = sPayrollID.substr(0, 4);" & vbNewLine
												Response.Write "document.ReportFrm.StartPayrollMonth.value = sPayrollID.substr(4, 2);" & vbNewLine
												Response.Write "if (parseInt(sPayrollID.substr(6, 2)) <= 15) {" & vbNewLine
													Response.Write "document.ReportFrm.StartPayrollDay.value = '01';" & vbNewLine
												Response.Write "} else {" & vbNewLine
													Response.Write "document.ReportFrm.StartPayrollDay.value = '15';" & vbNewLine
												Response.Write "}" & vbNewLine
												Response.Write "document.ReportFrm.EndPayrollYear.value = sPayrollID.substr(0, 4);" & vbNewLine
												Response.Write "document.ReportFrm.EndPayrollMonth.value = sPayrollID.substr(4, 2);" & vbNewLine
												Response.Write "document.ReportFrm.EndPayrollDay.value = sPayrollID.substr(6, 2);" & vbNewLine
												Response.Write "break;" & vbNewLine
											End If
									Response.Write "}" & vbNewLine
									Response.Write "break;" & vbNewLine
								Response.Write "}" & vbNewLine
							Response.Write "}" & vbNewLine
						End If
					Response.Write "} // End of DisplayPayrollFilters" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine

				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Nómina:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""PayrollID"" ID=""PayrollIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""DisplayPayrollFilters(this.value)"">"
					If(oRequest("ReportID").Item)="1403"Then
                        Response.Write "<option value=0 >  </option> "
                    End If
                    If InStr(1, sFlags, ("," & L_PAYROLL_FLAGS & ",")) > 0 Then
						'Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID<>0)", "PayrollID Desc", "", "", sErrorDescription)
                        Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName",  "Where (PayrollTypeID<>0) And (PayrollDate>20131231) ", "PayrollID Desc", "", "", sErrorDescription)
					ElseIf InStr(1, sFlags, ("," & L_OPEN_PAYROLL_FLAGS & ",")) > 0 Then
						'Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID<>0) And (IsClosed<>1)", "PayrollID Desc", "", "", sErrorDescription)
                        Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1)", "PayrollID Desc", "", "", sErrorDescription)
					ElseIf InStr(1, sFlags, ("," & L_CLOSED_PAYROLL_FLAGS & ",")) > 0 Then
						'Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID<>0) And (IsClosed=1)", "PayrollID Desc", "", "", sErrorDescription)
                        Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed=1)", "PayrollID Desc", "", "", sErrorDescription)
					ElseIf InStr(1, sFlags, ("," & L_PAYROLL1_FLAGS & ",")) > 0 Then
						'Response.Write "<OPTION VALUE=""-1"">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID<>0) And (payrollDate>20131231)", "PayrollID Desc", "", "", sErrorDescription)
					ElseIf InStr(1, sFlags, ("," & L_ORDINARY_PAYROLL_FLAGS & ",")) > 0 Then
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID=1)", "PayrollID Desc", "", "", sErrorDescription)
					ElseIf InStr(1, sFlags, ("," & L_CANCELL_PAYROLL_FLAGS & ",")) > 0 Then
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID=0) And (LENGTH(PayrollID) = 9)", "PayrollID Desc", "", "", sErrorDescription)
                    End If
				Response.Write "</SELECT><BR /><BR />"
				If (StrComp(GetASPFileName(""), "Payroll.asp", vbBinaryCompare) = 0) And (StrComp(oRequest("Action").Item, "ModifyPayroll", vbBinaryCompare) = 0) Then
					Response.Write "<DIV NAME=""ConceptsDiv"" ID=""ConceptsDiv"" STYLE=""display: none"">"
						Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Seleccione los conceptos a calcular:<BR /></FONT>"
						Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""PayrollConceptID"" ID=""PayrollConceptIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
							Response.Write "<OPTION VALUE="""">Todos</OPTION>"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(EndDate=30000000)", "ConceptShortName, ConceptName", "", "", sErrorDescription)
						Response.Write "</SELECT><BR /><BR />"
					Response.Write "</DIV>"

					Response.Write "<DIV NAME=""PeriodsDiv"" ID=""PeriodsDiv"" STYLE=""display: none"">"
						Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Calcular retroactivos desde </FONT>"
						Response.Write DisplayDateCombos(CInt(oRequest("StartPayrollYear").Item), CInt(oRequest("StartPayrollMonth").Item), CInt(oRequest("StartPayrollDay").Item), "StartPayrollYear", "StartPayrollMonth", "StartPayrollDay", Year(Date()), Year(Date()), True, False)
						Response.Write "<FONT FACE=""Arial"" SIZE=""2""> hasta </FONT>"
						Response.Write DisplayDateCombos(CInt(oRequest("EndPayrollYear").Item), CInt(oRequest("EndPayrollMonth").Item), CInt(oRequest("EndPayrollDay").Item), "EndPayrollYear", "EndPayrollMonth", "EndPayrollDay", Year(Date()), Year(Date()), True, False)
						Response.Write "<BR /><BR />"
					Response.Write "</DIV>"
					Response.Write "<DIV NAME=""DeleteDiv"" ID=""DeleteDiv"" STYLE=""display: none"">"
						Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Eliminar registros de la nómina: </FONT>"
						Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""DeleteFromPayroll"" ID=""DeleteFromPayrollCmb"" SIZE=""1 CLASS=""Lists"">"
							Response.Write "<OPTION VALUE="""">Sólo los registros correspondientes a los periodos marcados</OPTION>"
							Response.Write "<OPTION VALUE=""1"""
								If Len(oRequest("DeleteFromPayroll").Item) > 0 Then Response.Write " SELECTED=""1"""
							Response.Write ">Todos los registros</OPTION>"
						Response.Write "</SELECT><BR /><BR />"
					Response.Write "</DIV>"
					Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
						Response.Write "DisplayPayrollFilters(document.ReportFrm.PayrollID.value);" & vbNewLine
					Response.Write "//--></SCRIPT>" & vbNewLine
				End If
				iCounter = iCounter + 1
			End If
            'If(oRequest("ReportID").Item)="1403"Then
            '    Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
			'	Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Mes:<BR /></FONT>"
             '   Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			'	Response.Write DisplayDateCombosUsingSerial(30000000, "End", N_START_YEAR, Year(Date()), True, True)
                'Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""MonthID"" ID=""MonthCmb"" SIZE=""1"" CLASS=""Lists"">"
				'	Response.Write "<option value=01 > Enero </option> "
			'	Response.Write "<BR /><BR />"
            '    iCounter = iCounter + 1

           ' End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_MONTHS_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Mes:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""MonthID"" ID=""MonthIDCmb"" SIZE=""1"" CLASS=""Lists"">"
                    If(oRequest("ReportID").Item)="1403"Then
                        Response.Write "<option value=0 >  </option> "
                    End If
					For iIndex = 1 To UBound(asMonthNames_es)
						Response.Write "<OPTION VALUE=""" & iIndex & """>" & asMonthNames_es(iIndex) & "</OPTION>"
					Next
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If

            If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_QUARTER_FLAGS & ",")) > 0) Then
            	Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Bimestre:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""QuarterID"" ID=""QuarterIDCmb"" SIZE=""1"" CLASS=""Lists"">"
                    If(oRequest("ReportID").Item)="1403"Then
                        Response.Write "<option value=0 >  </option> "
                        Response.Write "<option value=1 >1</option> "
                        Response.Write "<option value=2 >2</option> "
                        Response.Write "<option value=3 >3</option> "
                        Response.Write "<option value=4 >4</option> "
                        Response.Write "<option value=5 >5</option> "
                        Response.Write "<option value=6 >6</option> "
                    End If
					'For iIndex = 1 To UBound(asMonthNames_es)
					'	Response.Write "<OPTION VALUE=""" & iIndex & """>" & asMonthNames_es(iIndex) & "</OPTION>"
					'Next
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_YEARS_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Año:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""YearID"" ID=""YearIDCmb"" SIZE=""1"" CLASS=""Lists"">"
					For iIndex = Year(Date()) To 2008 Step -1
						Response.Write "<OPTION VALUE=""" & iIndex & """>" & iIndex & "</OPTION>"
					Next
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_DOUBLE_MONTHS_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Meses:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">De </FONT><SELECT NAME=""StartMonthID"" ID=""StartMonthIDCmb"" SIZE=""1"" CLASS=""Lists"">"
					For iIndex = 1 To UBound(asMonthNames_es)
						Response.Write "<OPTION VALUE=""" & iIndex & """>" & asMonthNames_es(iIndex) & "</OPTION>"
					Next
				Response.Write "</SELECT>"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""> a </FONT><SELECT NAME=""EndMonthID"" ID=""EndMonthIDCmb"" SIZE=""1"" CLASS=""Lists"">"
					For iIndex = 1 To UBound(asMonthNames_es)
						Response.Write "<OPTION VALUE=""" & iIndex & """>" & asMonthNames_es(iIndex) & "</OPTION>"
					Next
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_CONCEPT_ID_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_CONCEPT_1_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_CONCEPT_2_FLAGS & ",")) > 0)  Or (InStr(1, sFlags, ("," & L_THIRD_CONCEPTS_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_THIRD_CONCEPTS2_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_MEMORY_CONCEPT_ID_FLAGS & ",")) > 0)) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Conceptos de pago:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
					If InStr(1, sFlags, ("," & L_CONCEPT_ID_FLAGS & ",")) > 0 Then
						Response.Write "<SELECT NAME=""ConceptID"" ID=""ConceptIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
							Response.Write "<OPTION VALUE="""">Todos</OPTION>"
							If aReportsComponent(N_ID_REPORTS) = 1108 Then
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID IN(" & EMPLOYEES_CONCEPTS & ")) And (EndDate=30000000)", "ConceptShortName, ConceptName", "", "", sErrorDescription)
							Else
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(EndDate=30000000)", "ConceptShortName, ConceptName", "", "", sErrorDescription)
							End If
					ElseIf InStr(1, sFlags, ("," & L_CONCEPT_1_FLAGS & ",")) > 0 Then
						Response.Write "<SELECT NAME=""ConceptID"" ID=""ConceptIDLst"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(EndDate=30000000)", "ConceptShortName, ConceptName", "", "", sErrorDescription)
                    ElseIf InStr(1, sFlags, ("," & L_CONCEPT_2_FLAGS & ",")) > 0 Then
						Response.Write "<SELECT NAME=""ConceptID"" ID=""ConceptIDLst"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write "<OPTION VALUE="""">Todos</OPTION>"
                            Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID In (-2,-1))", "ConceptShortName, ConceptName", "", "", sErrorDescription)
					ElseIf InStr(1, sFlags, ("," & L_THIRD_CONCEPTS_FLAGS & ",")) > 0 Then
						Response.Write "<SELECT NAME=""ConceptID"" ID=""ConceptIDLst"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "((ConceptID In (63,121,146)) Or (ConceptID In (Select CreditTypeID As ConceptID From CreditTypes Where (CreditTypeID<>86)))) And (EndDate=30000000)", "ConceptShortName", oRequest("ThirdConceptID").Item, "", sErrorDescription)
					ElseIf InStr(1, sFlags, ("," & L_THIRD_CONCEPTS2_FLAGS & ",")) > 0 Then
						Response.Write "<SELECT NAME=""ConceptID"" ID=""ConceptIDLst"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID In (57,58,59,126,64,83)) And (EndDate=30000000)", "ConceptShortName", oRequest("ThirdConceptID").Item, "", sErrorDescription)
							Response.Write "<OPTION VALUE=""6364"">63 Y 64 SEGURO METLIFE</OPTION>"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID In (80)) And (EndDate=30000000)", "ConceptShortName", oRequest("ThirdConceptID").Item, "", sErrorDescription)
					ElseIf InStr(1, sFlags, ("," & L_MEMORY_CONCEPT_ID_FLAGS & ",")) > 0 Then
						Response.Write "<SELECT NAME=""ConceptID"" ID=""ConceptIDLst"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write "<OPTION VALUE=""101"">CS. Sindicato independiente</OPTION>"
							Response.Write "<OPTION VALUE=""56,76,77"">54. Cuotas sindicales</OPTION>"
					End If
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If

            If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CONCENTRATE_CONCEPTS_FLAGS & ",")) > 0) Then 
                Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Tipo de Concentrado:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                Response.Write "<SELECT NAME=""ConcentrateConceptID"" ID=""ConcentrateConceptIDLst"" SIZE=""1"" CLASS=""Lists"">"
                                    Response.Write "<OPTION VALUE=""0"">Cancelados</OPTION>"
                                    Response.Write "<OPTION VALUE=""1"">Circulante</OPTION>"
                Response.Write "</SELECT><BR /><BR />"
            End If 
            
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_USER_FLAGS & ",")) > 0) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> 7) Then
				sErrorDescription = "No se pudo obtener la lista de responsables."
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Usuarios:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""UserID"" ID=""UserIDLst"" SIZE=""6"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Users", "UserID", "UserLastName, UserName", "(UserID>=10)", "UserLastName, UserName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_EMPLOYEE_NUMBER_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_EMPLOYEE_NUMBER1_FLAGS & ",")) > 0)) Then
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Select Case aReportsComponent(N_ID_REPORTS)
						Case ISSSTE_1203_REPORTS
							Response.Write "function VerifyEmployeeNumber() {" & vbNewLine
								Response.Write "var oForm = document.ReportFrm;" & vbNewLine
								Response.Write "if (oForm.EmployeeNumbers.value == '') {" & vbNewLine
									Response.Write "alert('Favor de introducir el número de empleado.');" & vbNewLine
									Response.Write "oForm.EmployeeNumbers.focus();" & vbNewLine
									Response.Write "return false;" & vbNewLine
								Response.Write "}" & vbNewLine

								Response.Write "if ((((parseInt('1' + oForm.DocumentStartDay.value) - 100) + ('1' + parseInt(oForm.DocumentStartMonth.value) - 100) + parseInt(oForm.DocumentStartYear.value)) > 0 ) && (((parseInt('1' + oForm.DocumentEndDay.value) - 100) + ('1' + parseInt(oForm.DocumentEndMonth.value) - 100) + parseInt(oForm.DocumentEndYear.value)) > 0 ) ) {" & vbNewLine
									Response.Write "if ((parseInt('1' + oForm.DocumentEndDay.value) - 100) * (parseInt('1' + oForm.DocumentEndMonth.value) - 100) * parseInt(oForm.DocumentEndYear.value) > 0 ) {" & vbNewLine
										Response.Write "if (((parseInt('1' + oForm.DocumentDay.value) - 100) + ((parseInt('1' + oForm.DocumentMonth.value) -100) * 100) + parseInt(oForm.DocumentYear.value) * 10000) > ((parseInt('1' + oForm.DocumentEndDay.value) - 100) + ((parseInt('1' + oForm.DocumentEndMonth.value) - 100) * 100) + parseInt(oForm.DocumentEndYear.value) * 10000)) {" & vbNewLine
											Response.Write "alert('Favor de verificar la vigencia del movimiento');" & vbNewLine
											Response.Write "oForm.DocumentStartDay.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "else" & vbNewLine
									Response.Write "{" & vbNewLine
										Response.Write "alert('Favor de verificar la vigencia del movimiento');" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "}" & vbNewLine
								Response.Write "else" & vbNewLine
								Response.Write "{" & vbNewLine
									Response.Write "alert('Favor de introducir las fehcas en que registro la solicitud de la hoja única del empleado');" & vbNewLine
									Response.Write "return false;" & vbNewLine
								Response.Write "}" & vbNewLine

								Response.Write "return true;" & vbNewLine
							Response.Write "} // End of VerifyEmployeeNumber" & vbNewLine
					End Select
					Response.Write "function AddEmployeeIDToSearchList() {" & vbNewLine
						Response.Write "var oForm = document.ReportFrm;" & vbNewLine
						Response.Write "if (oForm.EmployeeNumber.value != '') {" & vbNewLine
							Response.Write "oForm.EmployeeNumber.value = '000000' + oForm.EmployeeNumber.value;" & vbNewLine
							Response.Write "AddItemToList(oForm.EmployeeNumber.value.substr(oForm.EmployeeNumber.value.length - 6), oForm.EmployeeNumber.value.substr(oForm.EmployeeNumber.value.length - 6), null, oForm.EmployeeIDs)" & vbNewLine
							Response.Write "SelectAllItemsFromList(oForm.EmployeeIDs);" & vbNewLine
							Response.Write "oForm.EmployeeNumber.value = '';" & vbNewLine
							Response.Write "oForm.EmployeeNumber.focus();" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "} // End of AddEmployeeIDToSearchList" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
				If InStr(1, sFlags, ("," & L_EMPLOYEE_NUMBER1_FLAGS & ",")) > 0 Then
					Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Número de empleado:<BR /></FONT>"
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
					Response.Write "&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" SIZE=""10"" MAXLENGTH=""6"" VALUE=""" & oRequest("EmployeeNumber").Item & """ CLASS=""TextFields"" />"
					Response.Write "<BR /><BR />"
				Else
					Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
						Response.Write "<TD VALIGN=""TOP"">"
							Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
							Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Número de empleado:<BR /></FONT>"
							Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
							Response.Write "&nbsp;&nbsp;<TEXTAREA NAME=""EmployeeNumbers"" ID=""EmployeeNumbersTxtArea"" ROWS=""7"" COLS=""50"" MAXLENGTH=""2000"" CLASS=""TextFields"">" & oRequest("EmployeeNumbers").Item & "</TEXTAREA>"
'							Call DisplayLikeCombo("EmployeeNumberLike")
'							Response.Write "&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" SIZE=""10"" MAXLENGTH=""6"" VALUE=""" & oRequest("EmployeeNumber").Item & """ CLASS=""TextFields"" />"
'							Response.Write "&nbsp;&nbsp;<A HREF=""javascript: AddEmployeeIDToSearchList();""><IMG SRC=""Images/BtnCrclAdd.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Agregar"" BORDER=""0"" /></A>&nbsp;&nbsp;<BR />"
						Response.Write "</TD>"
						Response.Write "<TD>&nbsp;&nbsp;&nbsp;</TD>"
						Response.Write "<TD VALIGN=""TOP"">"
							Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Utilizar un archivo de texto:<BR /></FONT>"
							Response.Write "<IFRAME SRC=""BrowserFile.asp?Action=Filter&UserID=" & aLoginComponent(N_USER_ID_LOGIN) & """ NAME=""UploadInfoIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""96""></IFRAME>"
'							Response.Write "<BR /><SELECT NAME=""EmployeeIDs"" ID=""EmployeeIDsCmb"" SIZE=""6"" MULTIPLE=""1"" CLASS=""Lists"" STYLE=""width: 100px;"">"
'								If Len(oRequest("EmployeeIDs").Item) > 0 Then
'									For Each oItem In oRequest("EmployeeIDs")
'										Response.Write "<OPTION VALUE=""" & oItem & """ SELECTED=""1"">" & oItem & "</OPTION>"
'									Next
'								End If
'							Response.Write "</SELECT>"
'							Response.Write "&nbsp;<A HREF=""javascript: RemoveSelectedItemsFromList(null, document.ReportFrm.EmployeeIDs); SelectAllItemsFromList(document.ReportFrm.EmployeeIDs);""><IMG SRC=""Images/BtnCrclDelete.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Quitar"" BORDER=""0""></A><BR />"
						Response.Write "</TD>"
					Response.Write "</TR></TABLE>"
					Response.Write "<BR /><BR />"
				End If
				iCounter = iCounter + 1
				'Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				'	Response.Write "document.ReportFrm.EmployeeNumberLike.value=" & N_ENDS_LIKE & ";" & vbNewLine
				'Response.Write "//--></SCRIPT>" & vbNewLine
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_NUMBER7_FLAGS & ",")) > 0) Then
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "function AddEmployeeTempIDToSearchList() {" & vbNewLine
						Response.Write "var oForm = document.ReportFrm;" & vbNewLine
						Response.Write "if (oForm.EmployeeNumberTemp.value != '') {" & vbNewLine
							Response.Write "oForm.EmployeeNumberTemp.value = '0000000' + oForm.EmployeeNumberTemp.value;" & vbNewLine
							Response.Write "AddItemToList(oForm.EmployeeNumberTemp.value.substr(oForm.EmployeeNumberTemp.value.length - 7), oForm.EmployeeNumberTemp.value.substr(oForm.EmployeeNumberTemp.value.length - 7), null, oForm.EmployeeTempIDs)" & vbNewLine
							Response.Write "SelectAllItemsFromList(oForm.EmployeeTempIDs);" & vbNewLine
							Response.Write "oForm.EmployeeNumberTemp.value = '';" & vbNewLine
							Response.Write "oForm.EmployeeNumberTemp.focus();" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "} // End of AddEmployeeTempIDToSearchList" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
				Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
					Response.Write "<TD VALIGN=""TOP"">"
						Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Número de empleado temporal:<BR /></FONT>"
						Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
						Response.Write "&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""EmployeeNumberTemp"" ID=""EmployeeNumberTxt"" SIZE=""10"" MAXLENGTH=""7"" VALUE=""" & oRequest("EmployeeTempNumber").Item & """ CLASS=""TextFields"" />"
						Response.Write "&nbsp;&nbsp;<A HREF=""javascript: AddEmployeeTempIDToSearchList();""><IMG SRC=""Images/BtnCrclAdd.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Agregar"" BORDER=""0"" /></A>&nbsp;&nbsp;<BR />"
					Response.Write "</TD>"
					Response.Write "<TD VALIGN=""TOP""><BR />"
						Response.Write "<SELECT NAME=""EmployeeTempIDs"" ID=""EmployeeTempIDsCmb"" SIZE=""6"" MULTIPLE=""1"" CLASS=""Lists"" STYLE=""width: 100px;"">"
							If Len(oRequest("EmployeeTempIDs").Item) > 0 Then
								For Each oItem In oRequest("EmployeeTempIDs")
									Response.Write "<OPTION VALUE=""" & oItem & """ SELECTED=""1"">" & oItem & "</OPTION>"
								Next
							End If
						Response.Write "</SELECT>"
						Response.Write "&nbsp;<A HREF=""javascript: RemoveSelectedItemsFromList(null, document.ReportFrm.EmployeeTempIDs); SelectAllItemsFromList(document.ReportFrm.EmployeeTempIDs);""><IMG SRC=""Images/BtnCrclDelete.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Quitar"" BORDER=""0""></A><BR />"
					Response.Write "</TD>"
				Response.Write "</TR></TABLE><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_NAME_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Nombre del empleado:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				Call DisplayLikeCombo("EmployeeNameLike")
				Response.Write "&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""EmployeeName"" ID=""EmployeeNameTxt"" SIZE=""30"" MAXLENGTH=""255"" VALUE=""" & oRequest("EmployeeName").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_JOB_NUMBER_FLAGS & ",")) > 0) Then
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "function AddJobIDToSearchList() {" & vbNewLine
						Response.Write "var oForm = document.ReportFrm;" & vbNewLine
						Response.Write "if (oForm.JobNumber.value != '') {" & vbNewLine
							Response.Write "oForm.JobNumber.value = '000000' + oForm.JobNumber.value;" & vbNewLine
							Response.Write "AddItemToList(oForm.JobNumber.value.substr(oForm.JobNumber.value.length - 6), oForm.JobNumber.value.substr(oForm.JobNumber.value.length - 6), null, oForm.JobIDs)" & vbNewLine
							Response.Write "SelectAllItemsFromList(oForm.JobIDs);" & vbNewLine
							Response.Write "oForm.JobNumber.value = '';" & vbNewLine
							Response.Write "oForm.JobNumber.focus();" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "} // End of AddDisasterIDToSearchList" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
				Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
					Response.Write "<TD VALIGN=""TOP"">"
						Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Número de plaza:<BR /></FONT>"
						Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
						'Call DisplayLikeCombo("JobNumberLike")
						Response.Write "&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""JobNumber"" ID=""JobNumberTxt"" SIZE=""10"" MAXLENGTH=""6"" VALUE=""" & oRequest("JobNumber").Item & """ CLASS=""TextFields"" />"
						Response.Write "&nbsp;&nbsp;<A HREF=""javascript: AddJobIDToSearchList();""><IMG SRC=""Images/BtnCrclAdd.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Agregar"" BORDER=""0"" /></A>&nbsp;&nbsp;<BR />"
					Response.Write "</TD>"
					Response.Write "<TD VALIGN=""TOP""><BR />"
						Response.Write "<SELECT NAME=""JobIDs"" ID=""JobIDsCmb"" SIZE=""6"" MULTIPLE=""1"" CLASS=""Lists"" STYLE=""width: 100px;"">"
							If Len(oRequest("JobIDs").Item) > 0 Then
								For Each oItem In oRequest("JobIDs")
									Response.Write "<OPTION VALUE=""" & oItem & """ SELECTED=""1"">" & oItem & "</OPTION>"
								Next
							End If
						Response.Write "</SELECT>"
						Response.Write "&nbsp;<A HREF=""javascript: RemoveSelectedItemsFromList(null, document.ReportFrm.JobIDs); SelectAllItemsFromList(document.ReportFrm.JobIDs);""><IMG SRC=""Images/BtnCrclDelete.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Quitar"" BORDER=""0""></A><BR />"
					Response.Write "</TD>"
				Response.Write "</TR></TABLE><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_COMPANY_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_ONE_COMPANY_FLAGS & ",")) > 0)) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Empresas:<BR /></FONT>"
				If InStr(1, sFlags, ("," & L_COMPANY_FLAGS & ",")) > 0 Then
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""CompanyID"" ID=""CompanyIDLst"" SIZE=""6"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
				ElseIf InStr(1, sFlags, ("," & L_ONE_COMPANY_FLAGS & ",")) > 0 Then
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""CompanyID"" ID=""CompanyIDLst"" SIZE=""1"" CLASS=""Lists"">"
				End If
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Companies", "CompanyID", "CompanyShortName, CompanyName", "(ParentID>-1) And (EndDate=30000000)", "CompanyShortName, CompanyName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_REPORT_TYPE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Tipo de reporte:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""RADIO"" NAME=""ForWorkingCenter"" ID=""ForWorkingCenterRd"" VALUE="""""
					If Len(oRequest("ForWorkingCenter").Item) = 0 Then Response.Write " CHECKED=""1"""
				Response.Write " />&nbsp;Por centro de pago<BR />"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""RADIO"" NAME=""ForWorkingCenter"" ID=""ForWorkingCenterRd"" VALUE=""1"""
					If Len(oRequest("ForWorkingCenter").Item) > 0 Then Response.Write " CHECKED=""1"""
				Response.Write " />&nbsp;Presupuestal<BR /><BR />"
			End If

			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Áreas:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""AreaID"" ID=""AreaIDLst"" SIZE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {document.HierarchyMenuIFrame.location.href = 'HierarchyMenu.asp';} else {document.HierarchyMenuIFrame.location.href = 'HierarchyMenu.asp?Action=SubAreas&TargetField=' + this.form.name + '.SubAreaID&AreaID=' + this.value;}"">"
					If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) = 0 Then
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(ParentID=-1) And (AreaID>-1) And (EndDate=30000000)", "AreaCode, AreaName", oRequest("AreaID").Item, "", sErrorDescription)
					Else
						Response.Write "<OPTION VALUE=""-1"">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(ParentID=-1) And (AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) And (EndDate=30000000)", "AreaCode, AreaName", oRequest("AreaID").Item, "", sErrorDescription)
					End If
				Response.Write "</SELECT><BR />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SubAreaID"" ID=""SubAreaIDHdn"" VALUE=""" & oRequest("SubAreaID").Item & """ />"
				Response.Write "<IFRAME SRC=""HierarchyMenu.asp"
					If Len(oRequest("SubAreaID").Item) > 0 Then
						Response.Write "?Action=SubAreas&TargetField=ReportFrm.SubAreaID&AreaID=" & oRequest("AreaID").Item & "&SubAreaID=" & oRequest("SubAreaID").Item
					ElseIf Len(oRequest("AreaID").Item) > 0 Then
						Response.Write "?Action=SubAreas&TargetField=ReportFrm.SubAreaID&AreaID=" & oRequest("AreaID").Item
					Else
						'Response.Write "?Action=SubAreas&TargetField=ReportFrm.SubAreaID&AreaID=1&SubAreaID=" & oRequest("SubAreaID").Item
					End If
				Response.Write """ NAME=""HierarchyMenuIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""105""></IFRAME><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_EMPLOYEE_TYPE_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_EMPLOYEE_TYPE1_FLAGS & ",")) > 0)) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Tipos de tabulador:<BR /></FONT>"
				If InStr(1, sFlags, ("," & L_EMPLOYEE_TYPE1_FLAGS & ",")) > 0 Then
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDLst"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EmployeeTypes", "EmployeeTypeID", "EmployeeTypeName", "((EmployeeTypeID>=0) And (EmployeeTypeID<=7) And (EndDate=30000000))", "EmployeeTypeName", "", "", sErrorDescription)
					Response.Write "</SELECT><BR />"
				Else
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDLst"" SIZE=""10"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					If InStr(1, sFlags, ("," & L_CONCEPTS_VALUES_STATUS_FLAGS & ",")) = 0 Then Response.Write "<OPTION VALUE="""">Todos</OPTION>"
                    If StrComp(oRequest("Action").Item, "ModifyPayroll", vbBinaryCompare) = 0 Or StrComp(oRequest("Action").Item, "CalculatePayroll", vbBinaryCompare) = 0 Then
					    Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EmployeeTypes", "EmployeeTypeID", "EmployeeTypeName", "((EmployeeTypeID>=0) And (EmployeeTypeID<=7) Or (EmployeeTypeID IN (11,12))) And (EndDate=30000000)", "EmployeeTypeName", "", "", sErrorDescription)
                    Else
                        Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EmployeeTypes", "EmployeeTypeID", "EmployeeTypeName", "((EmployeeTypeID>=0) And (EmployeeTypeID<=7) Or (EmployeeTypeID IN (11,12))) And (EndDate=30000000)", "EmployeeTypeName", "", "", sErrorDescription)
                    End If
					Response.Write "</SELECT><BR />"
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""RADIO"" NAME=""EmployeeTypeTemp"" ID=""EmployeeTypeTempRd"" VALUE=""1"" onClick=""UnselectAllItemsFromList(document.ReportFrm.EmployeeTypeID); SelectItemByValue('1', false, document.ReportFrm.EmployeeTypeID)"" />Funcionarios&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""RADIO"" NAME=""EmployeeTypeTemp"" ID=""EmployeeTypeTempRd"" VALUE=""0"" onClick=""SelectAllItemsFromList(document.ReportFrm.EmployeeTypeID); UnSelectItemByValue('', false, document.ReportFrm.EmployeeTypeID); UnSelectItemByValue('1', false, document.ReportFrm.EmployeeTypeID); UnSelectItemByValue('7', false, document.ReportFrm.EmployeeTypeID);UnSelectItemByValue('12', false, document.ReportFrm.EmployeeTypeID);"" />Operativos"
				End If
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_POSITION_TYPE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Tipos de puesto:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""PositionTypeID"" ID=""PositionTypeIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "PositionTypes", "PositionTypeID", "PositionTypeName", "(EndDate=30000000)", "PositionTypeName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CLASSIFICATION_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Clasificación:&nbsp;</FONT>"
				Response.Write "<INPUT TYPE=""TEXT"" NAME=""ClassificationID"" ID=""ClassificationIDTxt"" SIZE=""2"" MAXLENGTH=""2"" VALUE=""" & oRequest("ClassificationID").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_GROUP_GRADE_LEVEL_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Grupo, grado, nivel:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""GroupGradeLevelID"" ID=""GroupGradeLevelIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "GroupGradeLevels", "GroupGradeLevelID", "GroupGradeLevelShortName", "(EndDate=30000000)", "GroupGradeLevelShortName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_INTEGRATION_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Integración:&nbsp;</FONT>"
				Response.Write "<INPUT TYPE=""TEXT"" NAME=""IntegrationID"" ID=""IntegrationIDTxt"" SIZE=""2"" MAXLENGTH=""2"" VALUE=""" & oRequest("IntegrationID").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_JOURNEY_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Turnos:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""JourneyID"" ID=""JourneyIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todas</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Journeys", "JourneyID", "JourneyShortName, JourneyName", "(EndDate=30000000)", "JourneyShortName, JourneyName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_SHIFT_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Horarios:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""ShiftID"" ID=""ShiftIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todas</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Shifts", "ShiftID", "ShiftShortName, ShiftName", "(EndDate=30000000)", "ShiftShortName, ShiftName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_LEVEL_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Nivel:&nbsp;</FONT>"
				Response.Write "<INPUT TYPE=""TEXT"" NAME=""LevelID"" ID=""LevelIDTxt"" SIZE=""3"" MAXLENGTH=""3"" VALUE=""" & oRequest("LevelID").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR />"
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_STATUS_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Estatus de los empleados:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""EmployeeStatusID"" ID=""EmployeeStatusIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "StatusEmployees", "StatusID", "StatusName", "", "StatusName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAYMENT_CENTER_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Centros de pago:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""PaymentCenterID"" ID=""PaymentCenterIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) = 0 Then
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", sCondition & "AreaID", "AreaCode, AreaName", "(ParentID>-1) And (EndDate=30000000)", "AreaCode, AreaName", "", "", sErrorDescription)
					Else
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(ParentID>-1) And (AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) And (EndDate=30000000)", "AreaCode, AreaName", "", "", sErrorDescription)
					End If
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_EMAIL_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Correo electrónico:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				Call DisplayLikeCombo("EmployeeEmailLike")
				Response.Write "&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""EmployeeEmail"" ID=""EmployeeEmailTxt"" SIZE=""30"" MAXLENGTH=""255"" VALUE=""" & oRequest("EmployeeEmail").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_SOCIAL_SECURITY_NUMBER_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Número de seguro social:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				Call DisplayLikeCombo("SocialSecurityNumberLike")
				Response.Write "&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""SocialSecurityNumber"" ID=""SocialSecurityNumberTxt"" SIZE=""30"" MAXLENGTH=""255"" VALUE=""" & oRequest("SocialSecurityNumber").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_BIRTH_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fechas de nacimiento:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
				Response.Write DisplayDateCombos(CInt(oRequest("StartBirthYear").Item), CInt(oRequest("StartBirthMonth").Item), CInt(oRequest("StartBirthDay").Item), "StartBirthYear", "StartBirthMonth", "StartBirthDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
				Response.Write DisplayDateCombos(CInt(oRequest("EndBirthYear").Item), CInt(oRequest("EndBirthMonth").Item), CInt(oRequest("EndBirthDay").Item), "EndBirthYear", "EndBirthMonth", "EndBirthDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_COUNTRY_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Países:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""CountryID"" ID=""CountryIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Countries", "CountryID", "CountryName", "", "CountryName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_RFC_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">RFC:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				Call DisplayLikeCombo("RFCLike")
				Response.Write "&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""RFC"" ID=""RFCTxt"" SIZE=""30"" MAXLENGTH=""255"" VALUE=""" & oRequest("RFC").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_CURP_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">CURP:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				Call DisplayLikeCombo("CURPLike")
				Response.Write "&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""CURP"" ID=""CURPTxt"" SIZE=""30"" MAXLENGTH=""255"" VALUE=""" & oRequest("CURP").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_GENDER_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Sexo:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""GenderID"" ID=""GenderIDLst"" SIZE=""3"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Genders", "GenderID", "GenderName", "", "GenderName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_ACTIVE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">¿El empleado está activo?<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""EmployeeActive"" ID=""EmployeeActiveLst"" SIZE=""3"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write "<OPTION VALUE=""0"">No</OPTION>"
					Response.Write "<OPTION VALUE=""1"">Sí</OPTION>"
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_START_DATE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fecha de ingreso al Instituto:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
				Response.Write DisplayDateCombos(CInt(oRequest("StartEmployeeStartYear").Item), CInt(oRequest("StartEmployeeStartMonth").Item), CInt(oRequest("StartEmployeeStartDay").Item), "StartEmployeeStartYear", "StartEmployeeStartMonth", "StartEmployeeStartDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
				Response.Write DisplayDateCombos(CInt(oRequest("EndEmployeeStartYear").Item), CInt(oRequest("EndEmployeeStartMonth").Item), CInt(oRequest("EndEmployeeStartDay").Item), "EndEmployeeStartYear", "EndEmployeeStartMonth", "EndEmployeeStartDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_GENERATING_AREAS_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Área generadora:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""GeneratingAreaID"" ID=""GeneratingAreaIDCmb"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(ParentID=-1) And (AreaID>-1) And (EndDate=30000000)", "AreaCode, AreaName", oRequest("AreaID").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_ZONE_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_STATES_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_ZONE_FLAGS_FOR_EMPLOYEES & ",")) > 0)) Then
				If (InStr(1, sFlags, ("," & L_ZONE_FLAGS_FOR_EMPLOYEES & ",")) = 0) Then
					Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
					If InStr(1, ",1006,1490,", "," & oRequest("ReportID").Item & ",", vbBinaryCompare) > 0 Then
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Entidades federativas (centro de trabajo):<BR /></FONT>"
					Else
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Entidades federativas (centro de pago):<BR /></FONT>"
					End If
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""ZoneID"" ID=""ZoneIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Zones", "ZoneID", "ZoneName", "(ParentID=-1) And (ZoneID>-1) And (EndDate=30000000)", "ZoneName", "", "", sErrorDescription)
						If InStr(1, sFlags, ("," & L_STATES_FLAGS & ",")) > 0 Then
							Response.Write "<OPTION VALUE=""38"">HOSP. REG. PDTE. JUAREZ OAXACA, OAX.</OPTION>"
						End If
					Response.Write "</SELECT><BR />"
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""RADIO"" NAME=""ZoneTypeTemp"" ID=""ZoneTypeTempRd"" VALUE=""1"" onClick=""UnselectAllItemsFromList(document.ReportFrm.ZoneID); SelectItemByValue('9', false, document.ReportFrm.ZoneID)"" />Local&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""RADIO"" NAME=""ZoneTypeTemp"" ID=""ZoneTypeTempRd"" VALUE=""0"" onClick=""SelectAllItemsFromList(document.ReportFrm.ZoneID); UnSelectItemByValue('', false, document.ReportFrm.ZoneID); UnSelectItemByValue('9', false, document.ReportFrm.ZoneID)"" />Estatales"
					Response.Write "<BR /><BR />"
					iCounter = iCounter + 1
				Else
					Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Entidades federativas:<BR /></FONT>"
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""ZoneID"" ID=""ZoneIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Zones", "ZoneID", "ZoneName", "(ParentID=-1) And (ZoneID>-1) And (EndDate=30000000)", "ZoneName", "", "", sErrorDescription)
					Response.Write "</SELECT><BR />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ZoneForEmployees"" ID=""ZoneForEmployeesHdn"" VALUE=""1"" />"
					iCounter = iCounter + 1
				End If
			End If
			If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_ZONE_FOR_PAYMENT_CENTER_FLAGS & ",")) > 0)) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Entidades federativas (centro de pago):<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""ZoneForPaymentCenterID"" ID=""ZoneForPaymentCenterIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todas</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Zones", "ZoneID", "ZoneName", "(ParentID=-1) And (ZoneID>-1) And (EndDate=30000000)", "ZoneName", "", "", sErrorDescription)
					If InStr(1, sFlags, ("," & L_STATES_FLAGS & ",")) > 0 Then
						Response.Write "<OPTION VALUE=""38"">HOSP. REG. PDTE. JUAREZ OAXACA, OAX.</OPTION>"
					End If
				Response.Write "</SELECT><BR />"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""RADIO"" NAME=""ZoneTypeTemp"" ID=""ZoneTypeTempRd"" VALUE=""1"" onClick=""UnselectAllItemsFromList(document.ReportFrm.ZoneForPaymentCenterID); SelectItemByValue('9', false, document.ReportFrm.ZoneForPaymentCenterID)"" />Local&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""RADIO"" NAME=""ZoneTypeTemp"" ID=""ZoneTypeTempRd"" VALUE=""0"" onClick=""SelectAllItemsFromList(document.ReportFrm.ZoneForPaymentCenterID); UnSelectItemByValue('', false, document.ReportFrm.ZoneForPaymentCenterID); UnSelectItemByValue('9', false, document.ReportFrm.ZoneForPaymentCenterID)"" />Estatales"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_POSITION_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Puestos:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""PositionID"" ID=""PositionIDIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Positions", "PositionID", "PositionShortName, PositionName", "(EndDate=30000000)", "PositionShortName, PositionName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_JOB_TYPE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Tipos de plaza:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""JobTypeID"" ID=""JobTypeIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "JobTypes", "JobTypeID", "JobTypeName", "", "JobTypeName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_OCCUPATION_TYPE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Tipos de ocupación:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""OccupationTypeID"" ID=""OccupationTypeIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "OccupationTypes", "OccupationTypeID", "OccupationTypeName", "", "OccupationTypeName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_JOB_START_DATE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fechas de inicio de las plazas:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
				Response.Write DisplayDateCombos(CInt(oRequest("StartJobStartYear").Item), CInt(oRequest("StartJobStartMonth").Item), CInt(oRequest("StartJobStartDay").Item), "StartJobStartYear", "StartJobStartMonth", "StartJobStartDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
				Response.Write DisplayDateCombos(CInt(oRequest("EndJobStartYear").Item), CInt(oRequest("EndJobStartMonth").Item), CInt(oRequest("EndJobStartDay").Item), "EndJobStartYear", "EndJobStartMonth", "EndJobStartDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_JOB_END_DATE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fechas de término de las plazas:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
				Response.Write DisplayDateCombos(CInt(oRequest("StartJobEndYear").Item), CInt(oRequest("StartJobEndMonth").Item), CInt(oRequest("StartJobEndDay").Item), "StartJobEndYear", "StartJobEndMonth", "StartJobEndDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
				Response.Write DisplayDateCombos(CInt(oRequest("EndJobEndYear").Item), CInt(oRequest("EndJobEndMonth").Item), CInt(oRequest("EndJobEndDay").Item), "EndJobEndYear", "EndJobEndMonth", "EndJobEndDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_JOB_STATUS_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Estatus de la plaza:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""JobStatusID"" ID=""JobStatusIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "StatusJobs", "StatusID", "StatusName", "", "StatusName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_JOB_ACTIVE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">¿La plaza está activa?<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""JobActive"" ID=""JobActiveLst"" SIZE=""3"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write "<OPTION VALUE=""0"">No</OPTION>"
					Response.Write "<OPTION VALUE=""1"">Sí</OPTION>"
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_CODE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Código del centro de trabajo:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				Call DisplayLikeCombo("AreaCodeLike")
				Response.Write "&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""AreaCode"" ID=""AreaCodeTxt"" SIZE=""30"" MAXLENGTH=""255"" VALUE=""" & oRequest("AreaCode").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_SHORT_NAME_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Clave del centro de trabajo:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				Call DisplayLikeCombo("AreaShortNameLike")
				Response.Write "&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""AreaShortName"" ID=""AreaShortNameTxt"" SIZE=""30"" MAXLENGTH=""255"" VALUE=""" & oRequest("AreaShortName").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_NAME_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Nombre del centro de trabajo:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				Call DisplayLikeCombo("AreaNameLike")
				Response.Write "&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""AreaName"" ID=""AreaNameTxt"" SIZE=""30"" MAXLENGTH=""255"" VALUE=""" & oRequest("AreaName").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_TYPE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Tipo de área:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""AreaTypeID"" ID=""AreaTypeIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todas</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "AreaTypes", "AreaTypeID", "AreaTypeShortName, AreaTypeName", "(EndDate=30000000)", "AreaTypeShortName, AreaTypeName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CONFINE_TYPE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Ámbito para las áreas:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""ConfineTypeID"" ID=""ConfineTypeIDLst"" SIZE=""4"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "ConfineTypes", "ConfineTypeID", "ConfineTypeShortName, ConfineTypeName", "(EndDate=30000000)", "ConfineTypeShortName, ConfineTypeName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CENTER_TYPE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Tipo de centro de trabajo:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""CenterTypeID"" ID=""CenterTypeIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "CenterTypes", "CenterTypeID", "CenterTypeShortName, CenterTypeName", "(EndDate=30000000)", "CenterTypeShortName, CenterTypeName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CENTER_SUBTYPE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Subtipo de centro de trabajo:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""CenterSubtypeID"" ID=""CenterSubtypeIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "CenterSubtypes", "CenterSubtypeID", "CenterSubtypeShortName, CenterSubtypeName", "(EndDate=30000000)", "CenterSubtypeShortName, CenterSubtypeName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_ATTENTION_LEVEL_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Nivel de atención:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""AttentionLevelID"" ID=""AttentionLevelIDLst"" SIZE=""6"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "AttentionLevels", "AttentionLevelID", "AttentionLevelShortName, AttentionLevelName", "(EndDate=30000000)", "AttentionLevelShortName, AttentionLevelName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_ECONOMIC_ZONE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Zona económica:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""EconomicZoneID"" ID=""EconomicZoneIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todas</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EconomicZones", "EconomicZoneID", "EconomicZoneName", "", "EconomicZoneName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_START_DATE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fechas de inicio de los centros de trabajo:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
				Response.Write DisplayDateCombos(CInt(oRequest("StartAreaStartYear").Item), CInt(oRequest("StartAreaStartMonth").Item), CInt(oRequest("StartAreaStartDay").Item), "StartAreaStartYear", "StartAreaStartMonth", "StartAreaStartDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
				Response.Write DisplayDateCombos(CInt(oRequest("EndAreaStartYear").Item), CInt(oRequest("EndAreaStartMonth").Item), CInt(oRequest("EndAreaStartDay").Item), "EndAreaStartYear", "EndAreaStartMonth", "EndAreaStartDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_END_DATE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fechas de término de los centros de trabajo:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
				Response.Write DisplayDateCombos(CInt(oRequest("StartAreaEndYear").Item), CInt(oRequest("StartAreaEndMonth").Item), CInt(oRequest("StartAreaEndDay").Item), "StartAreaEndYear", "StartAreaEndMonth", "StartAreaEndDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
				Response.Write DisplayDateCombos(CInt(oRequest("EndAreaEndYear").Item), CInt(oRequest("EndAreaEndMonth").Item), CInt(oRequest("EndAreaEndDay").Item), "EndAreaEndYear", "EndAreaEndMonth", "EndAreaEndDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_JOBS_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Plazas:&nbsp;</FONT>"
				Response.Write "<INPUT TYPE=""TEXT"" NAME=""Jobs"" ID=""JobsTxt"" SIZE=""4"" MAXLENGTH=""4"" VALUE=""" & oRequest("Jobs").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_TOTAL_JOBS_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Total de plazas:&nbsp;</FONT>"
				Response.Write "<INPUT TYPE=""TEXT"" NAME=""TotalJobs"" ID=""TotalJobsTxt"" SIZE=""4"" MAXLENGTH=""4"" VALUE=""" & oRequest("TotalJobs").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_STATUS_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Estatus del centro de trabajo:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""AreaStatusID"" ID=""AreaStatusIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "StatusAreas", "StatusID", "StatusName", "", "StatusName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CONCEPTS_VALUES_STATUS_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Estatus del tabulador:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""ConceptStatusID"" ID=""ConceptStatusIDLst"" SIZE=""4"" CLASS=""Lists"">"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "StatusConceptsValues", "StatusID", "StatusName", "", "StatusID", "1", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_REASON_ID_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Tipos de movimiento:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""ReasonID"" ID=""ReasonIDLst"" SIZE=""5"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Reasons", "ReasonID", "ReasonName", "ReasonID>=0", "ReasonName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AREA_ACTIVE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">¿El centro de trabajo está activo?<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""AreaActive"" ID=""AreaActiveLst"" SIZE=""3"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write "<OPTION VALUE=""0"">No</OPTION>"
					Response.Write "<OPTION VALUE=""1"">Sí</OPTION>"
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_TOTAL_PAYMENT_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Líquidos entre:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""TotalPaymentMin"" ID=""TotalPaymentMinTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & oRequest("TotalPaymentMin").Item & """ CLASS=""TextFields"" />"
				Response.Write "&nbsp;&nbsp;y&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""TotalPaymentMax"" ID=""TotalPaymentMaxTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & oRequest("TotalPaymentMax").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_BANK_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_ONE_BANK_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_ISSSTE_ONE_BANK_FLAGS & ",")) > 0)) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Bancos:<BR /></FONT>"
				If InStr(1, sFlags, ("," & L_BANK_FLAGS & ",")) > 0 Then
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""BankID"" ID=""BankIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
						Response.Write "<OPTION VALUE="""">Todos</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Banks", "BankID", "BankName", "(Banks.BankID>-1)", "BankName", "", "", sErrorDescription)
				ElseIf InStr(1, sFlags, ("," & L_ONE_BANK_FLAGS & ",")) > 0 Then
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""BankID"" ID=""BankIDLst"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Banks", "BankID", "BankName", "(Banks.BankID>-1)", "BankName", "", "", sErrorDescription)
				ElseIf InStr(1, sFlags, ("," & L_ISSSTE_ONE_BANK_FLAGS & ",")) > 0 Then
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""BankID"" ID=""BankIDLst"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Banks, BankAccounts", "Distinct Banks.BankID", "BankName", "(Banks.BankID=BankAccounts.BankID) And (Banks.BankID>-1) And (Banks.Active=1) And (BankAccounts.Active=1)", "BankName", "", "", sErrorDescription)
				End If
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_MEDICAL_AREAS_TYPES_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Tipo de Reporte:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""MedicalAreasTypeID"" ID=""MedicalAreasTypeIDIDLst"" SIZE=""3"" CLASS=""Lists"">"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "MedicalAreasTypes", "MedicalAreasTypeID", "MedicalAreasTypeName", "", "MedicalAreasTypeName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_DOCUMENT_FOR_LICENSE_NUMBER_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Número del oficio:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				Call DisplayLikeCombo("DocumentForLicenseNumberLike")
				Response.Write "&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""DocumentForLicenseNumber"" ID=""DocumentForLicenseNumberTxt"" SIZE=""25"" MAXLENGTH=""25"" VALUE=""" & oRequest("DocumentForLicenseNumber").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "document.ReportFrm.DocumentForLicenseNumberLike.value=" & N_ENDS_LIKE & ";" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_DOCUMENT_REQUEST_NUMBER_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Número del solicitud:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				Call DisplayLikeCombo("RequestNumberLike")
				Response.Write "&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""RequestNumber"" ID=""RequestNumberTxt"" SIZE=""25"" MAXLENGTH=""25"" VALUE=""" & oRequest("RequestNumber").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "document.ReportFrm.RequestNumberLike.value=" & N_ENDS_LIKE & ";" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_DOCUMENT_FOR_CANCEL_LICENSE_NUMBER_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Número del oficio de cancelación:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				Call DisplayLikeCombo("DocumentForCancelLicenseNumberLike")
				Response.Write "&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""DocumentForCancelLicenseNumber"" ID=""DocumentForCancelLicenseNumberTxt"" SIZE=""25"" MAXLENGTH=""25"" VALUE=""" & oRequest("DocumentForCancelLicenseNumber").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "document.ReportFrm.DocumentForCancelLicenseNumberLike.value=" & N_ENDS_LIKE & ";" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
			End If

			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_DATE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Periodo:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
				Response.Write DisplayDateCombos(CInt(oRequest("StartYear").Item), CInt(oRequest("StartMonth").Item), CInt(oRequest("StartDay").Item), "StartYear", "StartMonth", "StartDay", N_START_YEAR, Year(Date()), True, True)
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
				Response.Write DisplayDateCombos(CInt(oRequest("EndYear").Item), CInt(oRequest("EndMonth").Item), CInt(oRequest("EndDay").Item), "EndYear", "EndMonth", "EndDay", N_START_YEAR, Year(Date()) + 1, True, True)
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And ((InStr(1, sFlags, ("," & L_CHECK_CONCEPTS_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_CHECK_CONCEPTS_ALL_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_CHECK_CONCEPTS_EMPLOYEES_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_ONLY_CHECK_CONCEPTS_EMPLOYEES_FLAGS & ",")) > 0)) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Pagos de:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""CheckConceptID"" ID=""CheckConceptIDCmb"" SIZE=""1"" CLASS=""Lists"">"
					If InStr(1, sFlags, ("," & L_CHECK_CONCEPTS_ALL_FLAGS & ",")) > 0 Then
						Response.Write "<OPTION VALUE="""">Empleados con cheque y depósito</OPTION>"
					End If
					If InStr(1, sFlags, ("," & L_CHECK_CONCEPTS_EMPLOYEES_FLAGS & ",")) > 0 Then
						Response.Write "<OPTION VALUE="""">Empleados con cheque y depósito</OPTION>"
						Response.Write "<OPTION VALUE=""-2"">Empleados con depósito</OPTION>"
						Response.Write "<OPTION VALUE=""-1"">Empleados con cheque</OPTION>"
					ElseIf InStr(1, sFlags, ("," & L_ONLY_CHECK_CONCEPTS_EMPLOYEES_FLAGS & ",")) > 0 Then
						Response.Write "<OPTION VALUE=""-1"">Empleados con cheque</OPTION>"
						'Response.Write "<OPTION VALUE=""11"">Honorarios</OPTION>"
						Response.Write "<OPTION VALUE=""69"">Pensión alimenticia</OPTION>"
						Response.Write "<OPTION VALUE=""155"">Acreedores</OPTION>"
					Else
						If oRequest("ReportID").Item = "1400" Then Response.Write "<OPTION VALUE="""">Empleados con cheque y depósito</OPTION>"
						Response.Write "<OPTION VALUE=""-2"">Empleados con depósito</OPTION>"
						Response.Write "<OPTION VALUE=""-1"">Empleados con cheque</OPTION>"
						'Response.Write "<OPTION VALUE=""11"">Honorarios</OPTION>"
						Response.Write "<OPTION VALUE=""69"">Pensión alimenticia</OPTION>"
						Response.Write "<OPTION VALUE=""155"">Acreedores</OPTION>"
					End If
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
            If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAYMENT_TYPE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Tipo de nómina:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""PayrollTypeID"" ID=""PayrollTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
                    Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "PayrollTypes", "PayrollTypeID", "PayrollTypeName", "", "PayrollTypeName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CHECK_NUMBER_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Folios entre:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""CheckNumberMin"" ID=""CheckNumberMinTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & oRequest("CheckNumberMin").Item & """ CLASS=""TextFields"" />"
				Response.Write "&nbsp;&nbsp;y&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""CheckNumberMax"" ID=""CheckNumberMaxTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & oRequest("CheckNumberMax").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_HAS_ALIMONY_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<INPUT TYPE=""CHECKBOX"" NAME=""HasAlimony"" ID=""HasAlimonyChk"" VALUE=""1"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Mostrar únicamente empleados con pensión alimenticia.</FONT>"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_HAS_CREDITS_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<INPUT TYPE=""CHECKBOX"" NAME=""HasCredits"" ID=""HasCreditsChk"" VALUE=""1"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Mostrar únicamente empleados con productos de terceros.</FONT>"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If

			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_NUMBER_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Número de trámite:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				'Call DisplayLikeCombo("PaperworkNumberLike")
				'Response.Write "&nbsp;&nbsp;"
				Response.Write "<INPUT TYPE=""TEXT"" NAME=""PaperworkNumber"" ID=""PaperworkNumberTxt"" SIZE=""30"" MAXLENGTH=""255"" VALUE=""" & oRequest("PaperworkNumber").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_FOLIO_NUMBER_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Número de Folio:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Entre <INPUT TYPE=""TEXT"" NAME=""FilterStartNumber"" ID=""FilterStartNumberTxt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & oRequest("FilterStartNumber").Item & """ CLASS=""TextFields"" /> y <INPUT TYPE=""TEXT"" NAME=""FilterEndNumber"" ID=""FilterEndNumberTxt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & oRequest("FilterEndNumber").Item & """ CLASS=""TextFields"" /></FONT>"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_START_DATE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fechas de recepción:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
				Response.Write DisplayDateCombos(CInt(oRequest("PaperworkStartStartYear").Item), CInt(oRequest("PaperworkStartStartMonth").Item), CInt(oRequest("PaperworkStartStartDay").Item), "PaperworkStartStartYear", "PaperworkStartStartMonth", "PaperworkStartStartDay", N_START_YEAR, Year(Date()), True, True)
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
				Response.Write DisplayDateCombos(CInt(oRequest("PaperworkStartEndYear").Item), CInt(oRequest("PaperworkStartEndMonth").Item), CInt(oRequest("PaperworkStartEndDay").Item), "PaperworkStartEndYear", "PaperworkStartEndMonth", "PaperworkStartEndDay", N_START_YEAR, Year(Date()), True, True)
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_ESTIMATED_DATE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fechas de límite de respuesta:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
				Response.Write DisplayDateCombos(CInt(oRequest("PaperworkEstimatedStartYear").Item), CInt(oRequest("PaperworkEstimatedStartMonth").Item), CInt(oRequest("PaperworkEstimatedStartDay").Item), "PaperworkEstimatedStartYear", "PaperworkEstimatedStartMonth", "PaperworkEstimatedStartDay", N_START_YEAR, Year(Date()) + 1, True, True)
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
				Response.Write DisplayDateCombos(CInt(oRequest("PaperworkEstimatedEndYear").Item), CInt(oRequest("PaperworkEstimatedEndMonth").Item), CInt(oRequest("PaperworkEstimatedEndDay").Item), "PaperworkEstimatedEndYear", "PaperworkEstimatedEndMonth", "PaperworkEstimatedEndDay", N_START_YEAR, Year(Date()) + 1, True, True)
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_END_DATE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fechas de atención:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
				Response.Write DisplayDateCombos(CInt(oRequest("PaperworkEndStartYear").Item), CInt(oRequest("PaperworkEndStartMonth").Item), CInt(oRequest("PaperworkEndStartDay").Item), "PaperworkEndStartYear", "PaperworkEndStartMonth", "PaperworkEndStartDay", N_START_YEAR, Year(Date()), True, True)
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
				Response.Write DisplayDateCombos(CInt(oRequest("PaperworkEndEndYear").Item), CInt(oRequest("PaperworkEndEndMonth").Item), CInt(oRequest("PaperworkEndEndDay").Item), "PaperworkEndEndYear", "PaperworkEndEndMonth", "PaperworkEndEndDay", N_START_YEAR, Year(Date()), True, True)
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_DOCUMENT_NUMBER_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Número de documento:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				Call DisplayLikeCombo("PpwkDocumentNumberLike")
				Response.Write "&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""PpwkDocumentNumber"" ID=""PpwkDocumentNumberTxt"" SIZE=""30"" MAXLENGTH=""255"" VALUE=""" & oRequest("PpwkDocumentNumber").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_OWNERS_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Responsables:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""OwnerIDs"" ID=""OwnerIDsLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Dim sOwnerIDs
					Call GetPaperworksOwnersForUser(sOwnerIDs, "")
					If InStr(1, sOwnerIDs, "-1", vbBinaryCompare) = 0 Then
						sOwnerIDs = " And (OwnerID In (" & sOwnerIDs & "))"
					Else
						sOwnerIDs = ""
						Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					End If
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "PaperworkOwners", "OwnerID", "OwnerID As OwnerID2, OwnerName, EmployeeID", "(OwnerID>-1)" & sOwnerIDs, "OwnerID", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_TYPE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Tipo de trámite:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""PaperworkTypeID"" ID=""PaperworkTypeIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "PaperworkTypes", "PaperworkTypeID", "PaperworkTypeName", "", "PaperworkTypeName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_OWNER_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Número de empleado del responsable:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				Call DisplayLikeCombo("OwnerNumberLike")
				Response.Write "&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""OwnerNumber"" ID=""OwnerNumberTxt"" SIZE=""30"" MAXLENGTH=""255"" VALUE=""" & oRequest("OwnerNumber").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "document.ReportFrm.OwnerNumberLike.value=" & N_ENDS_LIKE & ";" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_STATUS_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Estatus del trámite:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""PaperworkStatusID"" ID=""PaperworkStatusIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "StatusPaperworks", "StatusID", "StatusName", "", "StatusName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If


			' ******** TIPO DE ASUNTO
            If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_SUBJECT_TYPES & ",")) > 0) Then
                Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
                Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Tipo de asunto:&nbsp;</FONT>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SubjectTypeID"" ID=""SubjectTypeIDHdn"" />"
				Response.Write "<INPUT TYPE=""TEXT"" NAME=""SubjectTypeName"" ID=""SubjectTypeNameTxt"" SIZE=""100"" VALUE="""" />"
				Response.Write "<A HREF=""javascript: SearchRecord(document.ReportFrm.SubjectTypeName.value, 'PaperworkCatalogs&SubjectTypeIDs=1&StartDate=-1', 'SearchSubjectTypesIFrame', 'ReportFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Buscar tipos de asunto"" BORDER=""0"" ALIGN=""ABSMIDDLE"" /></A><BR />"
				Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchSubjectTypesIFrame"" FRAMEBORDER=""0"" WIDTH=""650"" HEIGHT=""26""></IFRAME>"
                Response.Write "<BR /><BR />"
                iCounter = iCounter + 1
            End If
            ' ********  *************

            If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_PAPERWORK_PRIORITY_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Prioridad:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""PriorityID"" ID=""PriorityIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Priorities", "PriorityID", "PriorityName", "", "PriorityName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If

			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_STATE_TYPE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Reporte:<BR /></FONT>"
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""RADIO"" NAME=""StateType"" ID=""StateTypeRd"" VALUE="""" CHECKED=""1"" />Todos<BR />"
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""RADIO"" NAME=""StateType"" ID=""StateTypeRd"" VALUE=""0"" />Foráneo<BR />"
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""RADIO"" NAME=""StateType"" ID=""StateTypeRd"" VALUE=""1"" />Local<BR />"
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If

			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_COURSE_NAME_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Curso:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""CourseID"" ID=""CourseIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "SADE_Curso", "ID_Curso", "Nombre_Curso", "", "Nombre_Curso", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_COURSE_DIPLOMA_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Diplomado:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""ProfileID"" ID=""ProfileIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "SADE_Perfiles", "ID_Perfil", "Nombre_Perfil", "(ID_Padre=0)", "Nombre_Perfil", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If

			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_AREA_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Área:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""BudgetAreaID"" ID=""BudgetAreaIDTxt"" SIZE=""30"" MAXLENGTH=""255"" VALUE=""" & oRequest("BudgetAreaID").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_COMPANIES_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Compañía:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""BudgetCompanyID"" ID=""BudgetCompanyIDLst"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todas</OPTION>"
					Response.Write "<OPTION VALUE=""-1"">ISSSTE Asegurador</OPTION>"
					Response.Write "<OPTION VALUE=""170"">PENSIONISSSTE</OPTION>"
					Response.Write "<OPTION VALUE=""500"">SUPERISSSTE</OPTION>"
					Response.Write "<OPTION VALUE=""700"">FOVISSSTE</OPTION>"
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_PROGRAM_DUTY_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Programa presupuestario:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""BudgetProgramDutyID"" ID=""BudgetProgramDutyIDLst"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsProgramDuties", "ProgramDutyID", "ProgramDutyShortName, ProgramDutyName", "(ProgramDutyID>-1)", "ProgramDutyShortName", oRequest("BudgetProgramDutyID").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_FUND_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fondo:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""BudgetFundID"" ID=""BudgetFundIDLst"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsFunds", "FundID", "FundShortName, FundName", "(FundID>-1)", "FundShortName", oRequest("BudgetFundID").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_DUTY_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Función:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""BudgetDutyID"" ID=""BudgetDutyIDLst"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsDuties", "DutyID", "DutyShortName, DutyName", "(DutyID>-1)", "DutyShortName", oRequest("BudgetDutyID").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_ACTIVE_DUTY_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Subfunción activa:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""BudgetActiveDutyID"" ID=""BudgetActiveDutyIDLst"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsActiveDuties", "ActiveDutyID", "ActiveDutyShortName, ActiveDutyName", "(ActiveDutyID>-1)", "ActiveDutyShortName", oRequest("BudgetActiveDutyID").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_SPECIFIC_DUTY_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Subfunción específica:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""BudgetSpecificDutyID"" ID=""BudgetSpecificDutyIDLst"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsSpecificDuties", "SpecificDutyID", "SpecificDutyShortName, SpecificDutyName", "(SpecificDutyID>-1)", "SpecificDutyShortName", oRequest("BudgetSpecificDutyID").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_PROGRAM_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Programa:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""BudgetProgramID"" ID=""BudgetProgramIDLst"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsPrograms", "ProgramID", "ProgramShortName, ProgramName", "(ProgramID>-1)", "ProgramShortName", oRequest("BudgetProgramID").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_REGION_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Región:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""BudgetRegionID"" ID=""BudgetRegionIDLst"" SIZE=""1"" CLASS=""Lists"""
					If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_LOCATION_FLAGS & ",")) > 0) Then Response.Write " onChange=""if (this.value != '') {SearchRecord(this.value, 'Zones_Level2', 'DisplayOptionsIFrame', 'ReportFrm');}"""
				Response.Write ">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Zones", "ZoneID", "ZoneCode, ZoneName", "(ParentID=-1) And (ZoneID>-1)", "ZoneCode", oRequest("BudgetRegionID").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_UR_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">UR:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""BudgetUR"" ID=""BudgetURTxt"" SIZE=""3"" MAXLENGTH=""3"" VALUE=""" & oRequest("BudgetUR").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_CT_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">CT:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""BudgetCT"" ID=""BudgetCTTxt"" SIZE=""3"" MAXLENGTH=""3"" VALUE=""" & oRequest("BudgetCT").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_AUX_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">AUX:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""BudgetAUX"" ID=""BudgetAUXTxt"" SIZE=""2"" MAXLENGTH=""2"" VALUE=""" & oRequest("BudgetAUX").Item & """ CLASS=""TextFields"" />"
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_LOCATION_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Municipio:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""LocationID"" ID=""LocationIDLst"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					If Len(oRequest("BudgetRegionID").Item) > 0 Then Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Zones", "ZoneID", "ZoneCode, ZoneName", "(ParentID=" & oRequest("BudgetRegionID").Item & ") And (ZoneID>-1)", "ZoneCode", oRequest("LocationID").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_BUDGET1_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Partida:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""BudgetID1"" ID=""BudgetID1Lst"" SIZE=""1"" CLASS=""Lists"""
					If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_BUDGET2_FLAGS & ",")) > 0) Then Response.Write " onChange=""if (this.value != '') {SearchRecord(this.value, 'Budget_Level2', 'DisplayOptionsIFrame', 'ReportFrm');}"""
				Response.Write ">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Budgets", "BudgetID", "BudgetShortName", "(ParentID=-1) And (BudgetID>-1)", "BudgetShortName", oRequest("BudgetID1").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_BUDGET2_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Subpartida:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""BudgetID2"" ID=""BudgetID2Lst"" SIZE=""1"" CLASS=""Lists"""
					If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_BUDGET3_FLAGS & ",")) > 0) Then Response.Write " onChange=""if (this.value != '') {SearchRecord(this.value, 'Budget_Level3', 'DisplayOptionsIFrame', 'ReportFrm');}"""
				Response.Write ">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					If Len(oRequest("BudgetID1").Item) > 0 Then Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Budgets", "BudgetID", "BudgetShortName", "(ParentID=" & oRequest("BudgetID1").Item & ") And (BudgetID>-1)", "BudgetShortName", oRequest("BudgetID2").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_BUDGET3_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Tipo de pago:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""BudgetID3"" ID=""BudgetID3Lst"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					If Len(oRequest("BudgetID2").Item) > 0 Then Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Budgets", "BudgetID", "BudgetShortName, BudgetName", "(ParentID=" & oRequest("BudgetID2").Item & ") And (BudgetID>-1)", "BudgetShortName", oRequest("BudgetID3").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_CONFINE_TYPE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Ámbito:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""BudgetConfineTypeID"" ID=""BudgetConfineTypeIDLst"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsConfineTypes", "ConfineTypeID", "ConfineTypeShortName, ConfineTypeName", "(ConfineTypeID>-1)", "ConfineTypeShortName", oRequest("BudgetConfineTypeID").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_ACTIVITY1_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Actividad institucional:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""BudgetActivityID1"" ID=""BudgetActivityID1Lst"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsActivities1", "ActivityID", "ActivityShortName, ActivityName", "(ActivityID>-1)", "ActivityShortName", oRequest("BudgetActivityID1").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_ACTIVITY2_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Actividad presupuestaria:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""BudgetActivityID2"" ID=""BudgetActivityID2Lst"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsActivities2", "ActivityID", "ActivityShortName, ActivityName", "(ActivityID>-1)", "ActivityShortName", oRequest("BudgetActivityID2").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_PROCESS_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Proceso:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""BudgetProcessID"" ID=""BudgetProcessIDLst"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsProcesses", "ProcessID", "ProcessShortName, ProcessName", "(ProcessID>-1)", "ProcessShortName", oRequest("BudgetProcessID").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_YEAR_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Año:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""BudgetYear"" ID=""BudgetYearLst"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					For iIndex = N_PAYROLL_START_YEAR To Year(Date())
						Response.Write "<OPTION VALUE=""" & iIndex & """>" & iIndex & "</OPTION>"
					Next
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_MONTH_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Mes:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""BudgetMonth"" ID=""BudgetMonthLst"" SIZE=""13"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"" >"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					For iIndex = 1 To 12
						Response.Write "<OPTION VALUE=""" & iIndex & """>" & asMonthNames_es(iIndex) & "</OPTION>"
					Next
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BUDGET_ORIGINAL_POSITION_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Puestos:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""BudgetPositionID"" ID=""BudgetPositionIDIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsPositions", "PositionID", "PositionShortName, PositionName", "(EndDate=30000000)", "PositionShortName, PositionName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If

			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CREDITS_TYPES_ID_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Tipo de crédito:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""CreditTypeID"" ID=""CreditTypeIDLst"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "CreditTypes", "CreditTypeID", "CreditTypeShortName, CreditTypeName", "", "CreditTypeID", oRequest("CreditTypeID").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_BENEFICIARY_ID & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Beneficiaria de pensión alimenticia:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""BeneficiaryID"" ID=""BeneficiaryIDLst"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todas</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EmployeesBeneficiariesLKP", "BeneficiaryID", "BeneficiaryName, BeneficiaryLastName, BeneficiaryLastName2", "", "EmployeesBeneficiariesLKP", oRequest("BeneficiaryID").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_CREDITOR_ID & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Acreedor:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""CreditorID"" ID=""CreditorIDLst"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EmployeesCreditorsLKP", "CreditorID", "CreditorName, CreditorLastName, CreditorLastName2", "", "EmployeesCreditorsLKP", oRequest("CreditoryID").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & S_CREDITS_UPLOADED_FILE_NAME & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Archivo de carga de terceros:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""UploadedFileName"" ID=""UploadedFileNameLst"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Credits", "Distinct UploadedFileName", "UploadedFileName As UploadedFileName2", "(UploadedFileName Is Not Null) AND (UploadedFileName <> ' ') And (UploadedFileName <> '-')", "UploadedFileName", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_ABSENCE_ID_FLAGS & ",")) > 0) Then

				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "function ClearVacationPeriod() {" & vbNewLine
					Response.Write "if (document.ReportFrm.AbsenceID.value == 35 || document.ReportFrm.AbsenceID.value == 37 || document.ReportFrm.AbsenceID.value == 38) {"  & vbNewLine
						Response.Write "if (document.ReportFrm.AbsenceID.value == 35) {"  & vbNewLine
							Response.Write "RemoveAllItemsFromList(null, document.ReportFrm.PeriodVacationID);" & vbNewLine
							Response.Write "AddItemToList(1, 1, null, document.ReportFrm.PeriodVacationID);" & vbNewLine
							Response.Write "AddItemToList(2, 2, null, document.ReportFrm.PeriodVacationID);" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (document.ReportFrm.AbsenceID.value == 37) {"  & vbNewLine
							Response.Write "RemoveAllItemsFromList(null, document.ReportFrm.PeriodVacationID);" & vbNewLine
							Response.Write "AddItemToList(1, 1, null, document.ReportFrm.PeriodVacationID);" & vbNewLine
							Response.Write "AddItemToList(2, 2, null, document.ReportFrm.PeriodVacationID);" & vbNewLine
							Response.Write "AddItemToList(3, 3, null, document.ReportFrm.PeriodVacationID);" & vbNewLine
							Response.Write "AddItemToList(4, 4, null, document.ReportFrm.PeriodVacationID);" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (document.ReportFrm.AbsenceID.value == 38) {"  & vbNewLine
							Response.Write "RemoveAllItemsFromList(null, document.ReportFrm.PeriodVacationID);" & vbNewLine
							Response.Write "AddItemToList(1, 1, null, document.ReportFrm.PeriodVacationID);" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "ShowDisplay(document.all['EmployeeVacationDiv'])" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "else{" & vbNewLine
						Response.Write "HideDisplay(document.all['EmployeeVacationDiv'])" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "} // End of ClearVacationPeriod" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine

				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Tipo de incidencia:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""AbsenceID"" ID=""AbsenceIDLst"" SIZE=""1"" CLASS=""Lists"" onChange=""ClearVacationPeriod()"">"
				'Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""AbsenceID"" ID=""AbsenceIDLst"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Absences", "AbsenceID", "AbsenceShortName, AbsenceName", "", "AbsenceID", oRequest("AbsenceID").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"

				Response.Write "<DIV NAME=""EmployeeVacationDiv"" ID=""EmployeeVacationDiv"" STYLE=""display: none"">"
						Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Periodo:</FONT>"
						Response.Write "&nbsp;&nbsp;<SELECT NAME=""PeriodVacationID"" ID=""PeriodVacationIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write "<OPTION VALUE=""1"">1</OPTION>"
							Response.Write "<OPTION VALUE=""2"">2</OPTION>"
						Response.Write "</SELECT>&nbsp;&nbsp;"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Año:</FONT>"
						Response.Write "&nbsp;&nbsp;<SELECT NAME=""YearID"" ID=""YearIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							For iIndex = (Year(Date()) - 2) To Year(Date()) + 2
								Response.Write "<OPTION VALUE=""" & iIndex & """>" & iIndex & "</OPTION>"
							Next
						Response.Write "</SELECT><BR />"
					Response.Write "<BR />"
				Response.Write "</DIV>"

				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "HideDisplay(document.all['EmployeeVacationDiv'])" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_ABSENCE_ACTIVE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">¿La incidencia está aplicada?<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""EmployeesAbsenceActive"" ID=""EmployeesAbsenceActiveLst"" SIZE=""4"" MULTIPLE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write "<OPTION VALUE=""0"">No</OPTION>"
					Response.Write "<OPTION VALUE=""1"">Sí</OPTION>"
					Response.Write "<OPTION VALUE=""2"">Cancelada</OPTION>"
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_ABSENCE_APPLIED_DATE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Quincena de aplicación de la incidencia:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""AppliedDate"" ID=""AppliedDateCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID<>0)", "PayrollID Desc", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CONCEPTS_APPLIED_DATE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Quincena de aplicación:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""RegistrationDate"" ID=""RegistrationDateCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID<>0)", "PayrollID Desc", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CREDITS_APPLIED_DATE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Quincena de aplicación del crédito:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""CreditsAppliedDate"" ID=""CreditsAppliedDateCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID<>0)", "PayrollID Desc", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_ADJUSTMENT_APPLIED_DATE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Quincena de aplicación del reclamo:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""AdjustmentPayrollDate"" ID=""AdjustmentPayrollDateCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID<>0)", "PayrollID Desc", "", "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EXTRAHOURS_AND_SUNDAYS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Tipo de concepto:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""AbsenceID"" ID=""AbsenceIDLst"" SIZE=""1"" CLASS=""Lists"">"
					'Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Absences", "AbsenceID", "AbsenceShortName, AbsenceName", "(AbsenceID=201) Or (AbsenceID=202)", "AbsenceID", oRequest("AbsenceID").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CONCEPT_ACTIVE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">¿El concepto está aplicado?<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""EmployeesConceptActive"" ID=""EmployeesConceptActiveLst"" SIZE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE=""0"">En proceso</OPTION>"
					Response.Write "<OPTION VALUE=""1"" SELECTED=""1"">Aplicado</OPTION>"
					Response.Write "<OPTION VALUE=""2"">Cancelado</OPTION>"
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_BANK_ACCOUNTS_ACTIVE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">¿La cuenta bancaria está aplicada?<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""BankAccountsActive"" ID=""BankAccountsActiveLst"" SIZE=""3"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todas</OPTION>"
					Response.Write "<OPTION VALUE=""0"">No</OPTION>"
					Response.Write "<OPTION VALUE=""1"">Sí</OPTION>"
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_CREDITS_ACTIVE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">¿El crédito está aplicado?<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""CreditsActive"" ID=""CreditsActiveLst"" SIZE=""3"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write "<OPTION VALUE=""0"">No</OPTION>"
					Response.Write "<OPTION VALUE=""1"">Sí</OPTION>"
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_EMPLOYEE_SERVICES_SHEET_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Seleccione el tipo de la hoja única de servicios<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""ServicesSheetTypeID"" ID=""ServicesSheetTypeIDLst"" SIZE=""1"" CLASS=""Lists"" >"
					Response.Write "<OPTION VALUE=""A"">Sencilla</OPTION>"
					Response.Write "<OPTION VALUE=""B"">Normal</OPTION>"
					Response.Write "<OPTION VALUE=""C"">Completa</OPTION>"
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_LOG_DATE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fechas de entrada al sistema:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
				Response.Write DisplayDateCombos(CInt(oRequest("StartLogYear").Item), CInt(oRequest("StartLogMonth").Item), CInt(oRequest("StartLogDay").Item), "StartLogYear", "StartLogMonth", "StartLogDay", N_START_YEAR, Year(Date()), True, True)
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
				Response.Write DisplayDateCombos(CInt(oRequest("EndLogYear").Item), CInt(oRequest("EndLogMonth").Item), CInt(oRequest("EndLogDay").Item), "EndLogYear", "EndLogMonth", "EndLogDay", N_START_YEAR, Year(Date()), True, True)
				Response.Write "<BR /><BR />"
				iCounter = iCounter + 1
			End If

			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AUDIT_TYPE_ID_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Tipo de auditoria:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""AuditTypeID"" ID=""AuditTypeIDLst"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "AuditTypes", "AuditTypeID", "AuditTypeShortName, AuditTypeName", "", "AuditTypeID", oRequest("AuditTypeID").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_AUDIT_OPERATION_TYPE_ID_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Tipo de operación de auditoria:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""AuditOperationTypeID"" ID=""AuditOperationTypeIDLst"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "AuditOperationTypes", "AuditOperationTypeID", "AuditOperationShortName, AuditOperationName", "", "AuditOperationTypeID", oRequest("AuditOperationTypeID").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If

			If (lErrorNumber = 0) And (InStr(1, sFlags, ("," & L_REPORT_TITLE_FLAGS & ",")) > 0) Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Título del reporte<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""ReportTitle"" ID=""ReportTitleCmb"" SIZE=""1"" CLASS=""Lists"">"
					For iIndex = 0 To UBound(asTitles)
						Response.Write "<OPTION VALUE=""" & iIndex & """>" & asTitles(iIndex) & "</OPTION>"
					Next
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If
			If InStr(1, sFlags, ("," & L_MOVEMENT_TYPE & ",")) > 0 Then
				Response.Write "<IMG SRC=""Images/Crcl" & iCounter & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Tipo de reporte:<BR /></FONT>"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""ReportType"" ID=""ReportTypeLst"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE=""0"">Todos</OPTION>"
					Response.Write "<OPTION VALUE=""2"">Altas</OPTION>"
					Response.Write "<OPTION VALUE=""3"">Bajas</OPTION>"
					Response.Write "<OPTION VALUE=""4"">Cambio de datos personales</OPTION>"
					Response.Write "<OPTION VALUE=""5"">Cambios de sueldo</OPTION>"
					Response.Write "<OPTION VALUE=""6"">Cambios de días</OPTION>"
					Response.Write "<OPTION VALUE=""7"">Cambios de aportaciones</OPTION>"
				Response.Write "</SELECT><BR /><BR />"
				iCounter = iCounter + 1
			End If

		If (InStr(1, sFlags, ("," & L_NO_DIV_FLAGS & ",")) = 0) And (InStr(1, sFlags, ("," & L_DONT_CLOSE_DIV_FLAGS & ",")) = 0) Then Response.Write "</DIV>"
		Response.Write "<SCRIPT LANGUAGE=""JavaSCript""><!--" & vbNewLine
			Response.Write "SendURLValuesToForm(unescape('" & RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(oRequest, "Template"), "ReportID"), "ReportsStep") & "'.replace(/\+/gi, ' ')), document.ReportFrm);" & vbNewLine
			Response.Write "if (document.ReportFrm.DisasterID) {" & vbNewLine
				For Each oItem In oRequest("DisasterID")
					Response.Write "AddItemToList('" & oItem & "', '" & oItem & "', null, document.ReportFrm.DisasterID);" & vbNewLine
				Next
			Response.Write "}" & vbNewLine
			If (StrComp(GetASPFileName(""), "Payroll.asp", vbBinaryCompare) = 0) And ((InStr(1, sFlags, ("," & L_PAYROLL_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_OPEN_PAYROLL_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_CLOSED_PAYROLL_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_PAYROLL1_FLAGS & ",")) > 0) Or (InStr(1, sFlags, ("," & L_CANCELL_PAYROLL_FLAGS & ",")) > 0)) Then
				Response.Write "DisplayPayrollFilters(document.ReportFrm.PayrollID.value);" & vbNewLine
			End If
		Response.Write "//--></SCRIPT>" & vbNewLine

		If (InStr(1, sFlags, ("," & L_ZIP_WARNING_FLAGS & ",")) > 0) And B_USE_SMTP Then
			Response.Write "<BR /><BR />"
			Call DisplayInstructionsMessage("", "<IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""20"" /><B>Si la generación del reporte excediera los 5 minutos</B>, el sistema le enviará un correo electrónico cuando el reporte éste listo.")
		End If
	If InStr(1, sFlags, ("," & L_ZIP_WARNING_FLAGS & ",")) > 0 Then
		Response.Write "</DIV>"

		lErrorNumber = DisplaySavedZIPReports(oRequest, oADODBConnection, aReportsComponent(N_ID_REPORTS), sErrorDescription)
		If lErrorNumber = 0 Then
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "HideDisplay(document.all['ReportFilterDiv']);" & vbNewLine
				Response.Write "window.setTimeout(""HideDisplay(document.all['ContinueSpn'])"", 250);" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
		ElseIf lErrorNumber = L_ERR_NO_RECORDS Then
			lErrorNumber = 0
			sErrorDescription = ""
		End If
	End If
	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		If Len(oRequest("EmployeeIDs").Item) > 0 Then Response.Write "SelectAllItemsFromList(document.ReportFrm.EmployeeIDs);" & vbNewLine
		If Len(oRequest("JobIDs").Item) > 0 Then Response.Write "SelectAllItemsFromList(document.ReportFrm.JobIDs);" & vbNewLine
	Response.Write "//--></SCRIPT>" & vbNewLine

	Set oRecordset = Nothing
	DisplayReportFilter = lErrorNumber
	Err.Clear
End Function

Function DisplayReportTableTemplate(oRequest, sFlags, sErrorDescription)
'************************************************************
'Purpose: To display the items to be included in the report
'Inputs:  oRequest, sFlags
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayReportTableTemplate"
	Dim iColSpan
	Dim iSpan
	Dim iPos
	Dim iIndex
	Dim lErrorNumber
	Dim sColumnsTitles
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments

	iColSpan = 0
	sColumnsTitles = BuildTableTemplateHeader(oRequest, sFlags, "Datos")
	Response.Write "<TABLE WIDTH=""700"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
		asColumnsTitles = Split(sColumnsTitles, ",", -1, vbBinaryCompare)
		If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
			lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
		Else
			lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
		End If
		For iIndex = 0 To UBound(asColumnsTitles)
			If InStr(1, asColumnsTitles(iIndex), "<SPAN", vbBinaryCompare) = 1 Then
				iPos = InStr(1, asColumnsTitles(iIndex), "COLS=""", vbBinaryCompare) + Len("COLS=""")
				iSpan = CInt(Mid(asColumnsTitles(iIndex), iPos, (InStr(iPos, asColumnsTitles(iIndex), """", vbBinaryCompare) - iPos)))
				iColSpan = iColSpan + iSpan
			Else
				iColSpan = iColSpan + 1
			End If
		Next
		sRowContents = BuildList("Datos", TABLE_SEPARATOR, iColSpan)
		asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
		lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
	Response.Write "</TABLE>" & vbNewLine

	DisplayReportTableTemplate = lErrorNumber
	Err.Clear
End Function

Function DisplayReportTemplateForm(sFlags, sErrorDescription)
'************************************************************
'Purpose: To display the items to be included in the report
'Inputs:  sFlags
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayReportTemplateForm"
	Dim aFlags
	Dim iIndex
	Dim oItem

	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		Response.Write "function SendFlagsToTemplateIFrame(oList) {" & vbNewLine
			Response.Write "var sURL = 'Template.asp?';" & vbNewLine
			Response.Write "for (var i=0; i<oList.options.length; i++) {" & vbNewLine
				Response.Write "sURL += 'Template=' + oList.options[i].value + '&';" & vbNewLine
			Response.Write "}" & vbNewLine
			Response.Write "document.TemplateIFrame.location.href = sURL;" & vbNewLine
		Response.Write "}" & vbNewLine
	Response.Write "//--></SCRIPT>" & vbNewLine
	Call DisplayInstructionsMessage("Plantilla", "Seleccione los campos que desea incluir en el reporte. Puede seleccionar varios campos a la vez utilizando las teclas Shift y Control. Una vez que los campos han sido seleccionados, presiona el botón para agregar los campos a la lista de la derecha.<BR /><BR />Para cambiar el orden en que los campos seleccionados aparecerán en el reporte, utilice los botones para subir y bajar. Cuando la plantilla esté lista, presione el botón para continuar.")
	Response.Write "<BR />" & vbNewLine
	Response.Write "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0""><TR>" & vbNewLine
		Response.Write "<TD>" & vbNewLine
			Response.Write "<SELECT NAME=""FlagItems"" ID=""FlagItemLst"" SIZE=""7"" MULTIPLE=""1"" CLASS=""Lists"" STYLE=""width: 240px;"">"
				aFlags = Split(sFlags, ",", -1, vbBinaryCompare)
				For iIndex = 0 To UBound(aFlags)
					Select Case CInt(aFlags(iIndex))
						Case L_USER_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Responsables</OPTION>"
						Case L_EMPLOYEE_NUMBER_FLAGS, L_EMPLOYEE_NUMBER1_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Número de empleado</OPTION>"
						Case L_EMPLOYEE_NUMBER7_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Número de empleado temporal</OPTION>"
						Case L_EMPLOYEE_NAME_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Nombre del empleado</OPTION>"
						Case L_COMPANY_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Empresa</OPTION>"
						Case L_EMPLOYEE_TYPE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Tipos de tabulador</OPTION>"
						Case L_POSITION_TYPE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Tipos de puesto</OPTION>"
						Case L_CLASSIFICATION_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Clasificación</OPTION>"
						Case L_GROUP_GRADE_LEVEL_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Grupo, grado, nivel</OPTION>"
						Case L_INTEGRATION_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Integración</OPTION>"
						Case L_JOURNEY_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Turno</OPTION>"
						Case L_SHIFT_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Horario</OPTION>"
						Case L_LEVEL_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Nivel</OPTION>"
						Case L_EMPLOYEE_STATUS_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Estatus del empleado</OPTION>"
						Case L_PAYMENT_CENTER_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Centro de pago</OPTION>"
						Case L_EMPLOYEE_EMAIL_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Correo electrónico</OPTION>"
						Case L_SOCIAL_SECURITY_NUMBER_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Número de seguro social</OPTION>"
						Case L_EMPLOYEE_BIRTH_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Fecha de nacimiento</OPTION>"
						Case L_EMPLOYEE_COUNTRY_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>País</OPTION>"
						Case L_EMPLOYEE_RFC_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>RFC</OPTION>"
						Case L_EMPLOYEE_CURP_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>CURP</OPTION>"
						Case L_EMPLOYEE_GENDER_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Sexo</OPTION>"
						Case L_EMPLOYEE_ACTIVE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>¿Empleado activo?</OPTION>"
						Case L_EMPLOYEE_START_DATE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Fecha de ingreso al Instituto</OPTION>"
						Case L_JOB_NUMBER_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Número de plaza</OPTION>"
						Case L_ZONE_FLAGS, L_STATES_FLAGS, L_ZONE_FLAGS_FOR_EMPLOYEES, L_ZONE_FOR_PAYMENT_CENTER_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Entidad federativa</OPTION>"
						Case L_AREA_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Área</OPTION>"
						Case L_POSITION_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Puesto</OPTION>"
						Case L_JOB_TYPE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Tipo de plaza</OPTION>"
						Case L_OCCUPATION_TYPE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Tipo de ocupación</OPTION>"
						Case L_JOB_START_DATE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Inicio de la plaza</OPTION>"
						Case L_JOB_END_DATE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Término de la plaza</OPTION>"
						Case L_JOB_STATUS_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Estatus de la plaza</OPTION>"
						Case L_JOB_ACTIVE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>¿Plaza activa?</OPTION>"
						Case L_AREA_CODE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Código del centro de trabajo</OPTION>"
						Case L_AREA_SHORT_NAME_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Clave del centro de trabajo</OPTION>"
						Case L_AREA_NAME_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Nombre del centro de trabajo</OPTION>"
						Case L_AREA_TYPE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Tipo del área</OPTION>"
						Case L_CONFINE_TYPE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Ámbito para el área</OPTION>"
						Case L_CENTER_TYPE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Tipo de centro de trabajo</OPTION>"
						Case L_CENTER_SUBTYPE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Subtipo de centro de trabajo</OPTION>"
						Case L_ATTENTION_LEVEL_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Nivel de atención</OPTION>"
						Case L_ECONOMIC_ZONE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Zona económica</OPTION>"
						Case L_AREA_START_DATE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Fecha de inicio del centro de trabajo</OPTION>"
						Case L_AREA_END_DATE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Fecha de término del centro de trabajo</OPTION>"
						Case L_AREA_JOBS_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Plazas</OPTION>"
						Case L_AREA_TOTAL_JOBS_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Total de plazas</OPTION>"
						Case L_AREA_STATUS_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Estatus del centro de trabajo</OPTION>"
						Case L_GENERATING_AREAS_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Áreas generadoras</OPTION>"
						Case L_CONCEPTS_VALUES_STATUS_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Estatus del tabulador</OPTION>"
						Case L_EMPLOYEE_REASON_ID_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Tipo de movimiento</OPTION>"
						Case L_AREA_ACTIVE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>¿Centro de trabajo activo?</OPTION>"
						Case L_CONCEPT_ID_FLAGS, L_CONCEPT_1_FLAGS, L_THIRD_CONCEPTS_FLAGS, L_THIRD_CONCEPTS2_FLAGS, L_MEMORY_CONCEPT_ID_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Concepto de pago</OPTION>"
						Case L_TOTAL_PAYMENT_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Líquido</OPTION>"	
						Case L_BANK_FLAGS, L_ONE_BANK_FLAGS, L_ISSSTE_ONE_BANK_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Banco</OPTION>"
						Case L_MEDICAL_AREAS_TYPES_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Tipo de reporte</OPTION>"
						Case L_DOCUMENT_FOR_LICENSE_NUMBER_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>No. de oficio</OPTION>"
						Case L_DOCUMENT_REQUEST_NUMBER_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>No. de solicitud</OPTION>"
						Case L_DOCUMENT_FOR_CANCEL_LICENSE_NUMBER_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>No. de oficio de cancelación</OPTION>"

						Case L_PAYROLL_FLAGS, L_OPEN_PAYROLL_FLAGS, L_CLOSED_PAYROLL_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Nómina</OPTION>"
						Case L_MONTHS_FLAGS, L_DOUBLE_MONTHS_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Mes</OPTION>"
						Case L_YEARS_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Año</OPTION>"
						Case L_DATE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Periodo</OPTION>"

						Case L_PAPERWORK_NUMBER_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>No. de trámite</OPTION>"
						Case L_PAPERWORK_FOLIO_NUMBER_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>No. de folio</OPTION>"
						Case L_PAPERWORK_START_DATE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Fecha de recepción</OPTION>"
						Case L_PAPERWORK_ESTIMATED_DATE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Fecha límite de respuesta</OPTION>"
						Case L_PAPERWORK_END_DATE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Fecha de atención</OPTION>"
						Case L_PAPERWORK_DOCUMENT_NUMBER_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>No. de documento</OPTION>"
						Case L_PAPERWORK_TYPE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Tipo de trámite</OPTION>"
						Case L_PAPERWORK_OWNER_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Responsable</OPTION>"
						Case L_PAPERWORK_STATUS_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Estatus del trámite</OPTION>"
                        Case L_PAPERWORK_SUBJECT_TYPES
                            Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Tipo de asunto</OPTION>"
						Case L_PAPERWORK_PRIORITY_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Prioridad</OPTION>"
						Case L_PAPERWORK_OWNERS_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Responsables</OPTION>"

						Case L_COURSE_NAME_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Curso</OPTION>"
						Case L_COURSE_DIPLOMA_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Diplomado</OPTION>"
						Case L_COURSE_LOCATION_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Ubicación</OPTION>"
						Case L_COURSE_DURATION_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Duración</OPTION>"
						Case L_COURSE_PARTICIPANTS_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>No. de participantes</OPTION>"
						Case L_COURSE_DATES_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Fechas del curso</OPTION>"
						Case L_COURSE_GRADE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Calificaciones</OPTION>"

						Case L_BUDGET_AREA_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Áreas</OPTION>"
						Case L_BUDGET_COMPANIES_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Compañías</OPTION>"
						Case L_BUDGET_PROGRAM_DUTY_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Programa presupuestario</OPTION>"
						Case L_BUDGET_FUND_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Fondo</OPTION>"
						Case L_BUDGET_DUTY_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Función</OPTION>"
						Case L_BUDGET_ACTIVE_DUTY_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Subfunción activa</OPTION>"
						Case L_BUDGET_SPECIFIC_DUTY_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Subfunción específica</OPTION>"
						Case L_BUDGET_PROGRAM_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Programa</OPTION>"
						Case L_BUDGET_REGION_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Región</OPTION>"
						Case L_BUDGET_UR_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>UR</OPTION>"
						Case L_BUDGET_CT_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>CT</OPTION>"
						Case L_BUDGET_AUX_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>AUX</OPTION>"
						Case L_BUDGET_LOCATION_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Municipio</OPTION>"
						Case L_BUDGET_BUDGET1_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Partida</OPTION>"
						Case L_BUDGET_BUDGET2_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Subpartida</OPTION>"
						Case L_BUDGET_BUDGET3_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Tipo de pago</OPTION>"
						Case L_BUDGET_CONFINE_TYPE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Ámbito</OPTION>"
						Case L_BUDGET_ACTIVITY1_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Actividad institucional</OPTION>"
						Case L_BUDGET_ACTIVITY2_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Actividad presupuestaria</OPTION>"
						Case L_BUDGET_PROCESS_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Proceso</OPTION>"
						Case L_BUDGET_YEAR_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Año</OPTION>"
						Case L_BUDGET_MONTH_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Mes</OPTION>"
						Case L_BUDGET_ORIGINAL_POSITION_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Puesto</OPTION>"
						Case L_CREDITS_TYPES_ID_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Tipo de crédito</OPTION>"
						Case L_EMPLOYEE_BENEFICIARY_ID
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Beneficiaria de pensión alimenticia</OPTION>"
						Case L_EMPLOYEE_CREDITOR_ID
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Acreedor</OPTION>"
						Case S_CREDITS_UPLOADED_FILE_NAME
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Archivo de carga de tercero</OPTION>"
						Case L_ABSENCE_ID_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Tipo de incidencia</OPTION>"
						Case L_ABSENCE_ACTIVE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>¿Incidencia activa?</OPTION>"
						Case L_ABSENCE_APPLIED_DATE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>Quincena de aplicación de la incidencia</OPTION>"
						Case L_CONCEPT_ACTIVE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>¿Concepto activa?</OPTION>"
						Case L_BANK_ACCOUNTS_ACTIVE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>¿Cuenta bancaria activa?</OPTION>"
						Case L_CREDITS_ACTIVE_FLAGS
							Response.Write "<OPTION VALUE=""" & aFlags(iIndex) & """>¿Crédito activo?</OPTION>"
					End Select
				Next
			Response.Write "</SELECT>" & vbNewLine
		Response.Write "</TD>" & vbNewLine
		Response.Write "<TD>&nbsp;</TD>" & vbNewLine
		Response.Write "<TD>" & vbNewLine
			Response.Write "<A HREF=""javascript: DoNothing()"" onClick=""MoveItemsBetweenLists(['',''], document.ReportFrm.FlagItems, document.ReportFrm.Template); SendFlagsToTemplateIFrame(document.ReportFrm.Template)""><IMG SRC=""Images/BtnCrclAdd.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Agregar"" BORDER=""0""></A><BR /><BR />"
			Response.Write "<A HREF=""javascript: DoNothing()"" onClick=""MoveItemsBetweenLists(['',''], document.ReportFrm.Template, document.ReportFrm.FlagItems); SendFlagsToTemplateIFrame(document.ReportFrm.Template)""><IMG SRC=""Images/BtnCrclAddLeft.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Quitar"" BORDER=""0""></A>"
		Response.Write "</TD>" & vbNewLine
		Response.Write "<TD>&nbsp;</TD>" & vbNewLine
		Response.Write "<TD>" & vbNewLine
			Response.Write "<SELECT NAME=""Template"" ID=""TemplateLst"" SIZE=""7"" MULTIPLE=""1"" CLASS=""Lists"" STYLE=""width: 240px;"">"
			Response.Write "</SELECT>"
		Response.Write "</TD>" & vbNewLine
		Response.Write "<TD>&nbsp;</TD>" & vbNewLine
		Response.Write "<TD>" & vbNewLine
			Response.Write "<A HREF=""javascript: DoNothing()"" onClick=""MoveListItemUp(document.ReportFrm.Template); SendFlagsToTemplateIFrame(document.ReportFrm.Template)""><IMG SRC=""Images/BtnCrclAddUp.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Subir"" BORDER=""0""></A><BR /><BR />"
			Response.Write "<A HREF=""javascript: DoNothing()"" onClick=""MoveListItemDown(document.ReportFrm.Template); SendFlagsToTemplateIFrame(document.ReportFrm.Template)""><IMG SRC=""Images/BtnCrclAddDown.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Bajar"" BORDER=""0""></A>"
		Response.Write "</TD>" & vbNewLine
	Response.Write "</TR></TABLE><BR />" & vbNewLine
	Response.Write "<IFRAME SRC=""Template.asp"" NAME=""TemplateIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""112""></IFRAME>" & vbNewLine
	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		For Each oItem In oRequest("Template")
			Response.Write "SelectListItemByValue('" & oItem & "', true, document.ReportFrm.FlagItems);" & vbNewLine
			Response.Write "MoveItemsBetweenLists(['',''], document.ReportFrm.FlagItems, document.ReportFrm.Template);" & vbNewLine
			Response.Write "SendFlagsToTemplateIFrame(document.ReportFrm.Template);" & vbNewLine
		Next
	Response.Write "//--></SCRIPT>" & vbNewLine
	Call DisplayURLParametersAsHiddenValues(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(oRequest, "ReportType"), "ReportStep"), "Template"), "ReportID"))

	DisplayReportTemplateForm = Err.number
	Err.Clear
End Function

Function DisplaySavedZIPReports(oRequest, oADODBConnection, lReportID, sErrorDescription)
'************************************************************
'Purpose: To verify if the given report has been run before and
'         to display the information for the run reports
'Inputs:  oRequest, oADODBConnection, lReportID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplaySavedZIPReports"
	Dim sFilter
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información de los reportes guardados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Reports Where (UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ") And (ConstantID=" & lReportID & ") Order By ReportName", "ReportsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<DIV NAME=""ZIPReportDiv"" ID=""ZIPReportDiv"">"
				Select Case lReportID
					Case ISSSTE_1203_REPORTS
					Case Else
						Response.Write "<IMG SRC=""Images/Crcl1.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" /> <B><A HREF=""javascript: HideDisplay(document.all['ZIPReportDiv']); ShowDisplay(document.all['ReportFilterDiv']); ShowDisplay(document.all['ContinueSpn']);"">Genere nuevo reporte</A></B><BR /><BR />"
						Response.Write "<IMG SRC=""Images/Crcl2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" /> <B>Seleccione uno de los reportes generados con anterioridad:</B><BR /><BR />"
				End Select
				Response.Write "<TABLE WIDTH=""800"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
					asColumnsTitles = Split("Fecha,Filtro utilizado,Archivo", ",", -1, vbBinaryCompare)
					asCellWidths = Split("200,500,100", ",", -1, vbBinaryCompare)
					asCellAlignments = Split(",,CENTER", ",", -1, vbBinaryCompare)
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If

					Do While Not oRecordset.EOF
						sRowContents = DisplayDateFromSerialNumber(CLng(Left(CStr(oRecordset.Fields("ReportName").Value), Len("00000000"))), -1, -1, -1)
						sFilter = ""
						lErrorNumber = DisplayFilterInformation(CStr(oRecordset.Fields("ReportParameters1").Value) & "&ShowFilter=False", CStr(oRecordset.Fields("ReportParameters2").Value), False, sFilter, sErrorDescription)
						sRowContents = sRowContents & TABLE_SEPARATOR & sFilter
						sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & lReportID & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & CStr(oRecordset.Fields("ReportName").Value) & ".zip""><IMG SRC=""Images/IcnFileZIP.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Ver el reporte"" BORDER=""0"" /></A>"
						sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;&nbsp;<A HREF=""javascript: OpenNewWindow('Remove.asp?Action=Reports&ConstantID=" & lReportID & "&UserID=" & aLoginComponent(N_USER_ID_LOGIN) & "&ReportName=" & CStr(oRecordset.Fields("ReportName").Value) & "', null, 'RemoveWnd', 320, 240, 'no', 'yes')""><IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar el reporte"" BORDER=""0"" /></A>"
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)

						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
				Response.Write "</TABLE>" & vbNewLine
			Response.Write "</DIV>"
		Else
			lErrorNumber = L_ERR_NO_RECORDS
		End If
	End If

	Set oRecordset = Nothing
	DisplaySavedZIPReports = lErrorNumber
	Err.Clear
End Function

Function GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
'************************************************************
'Purpose: To count the disasters using the filter information
'Inputs:  oRequest
'Outputs: sCondition, lPayrollID, lForPayrollID
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetConditionFromURL"
	Dim sTemp
	Dim asTemp
	Dim sTableName
	Dim oItem
	Dim iIndex
	Dim jIndex

	If Len(oRequest("UserID").Item) > 0 Then
		sCondition = sCondition & " And (Users.UserID In (" & oRequest("UserID").Item & "))"
	End If
	If (Len(Replace(oRequest("EmployeeIDs").Item, " ", "")) > 0) And (Len(Replace(oRequest("EmployeeTempIDs").Item, " ", "")) > 0)  Then
		sCondition = sCondition & " And ((Employees.EmployeeID In (" & Replace(oRequest("EmployeeIDs").Item, " ", "") & ")) Or (Employees.EmployeeID In (" & Replace(oRequest("EmployeeTempIDs").Item, " ", "") & ")))"
	ElseIf Len(Replace(oRequest("EmployeeIDs").Item, " ", "")) > 0 Then
		sCondition = sCondition & " And (Employees.EmployeeID In (" & Replace(oRequest("EmployeeIDs").Item, " ", "") & "))"
	ElseIf Len(Replace(oRequest("EmployeeTempIDs").Item, " ", "")) > 0 Then
			sCondition = sCondition & " And (Employees.EmployeeID In (" & Replace(oRequest("EmployeeTempIDs").Item, " ", "") & "))"
	ElseIf Len(oRequest("EmployeeNumbers").Item) > 0 Then
		If InStr(1, oRequest("EmployeeNumbers").Item, "Emp_", vbBinaryCompare) = 0 Then
			sTemp = oRequest("EmployeeNumbers").Item
		Else
			sTemp = GetFileContents(Server.MapPath("Uploaded Files\Filters\" & oRequest("EmployeeNumbers").Item), sErrorDescription)
		End If
		sTemp = Replace(Replace(Replace(sTemp, " ", ""), vbNewLine, ","), ",,", ",")
		Do While (InStr(1, sTemp, ",,", vbBinaryCompare) > 0)
			sTemp = Replace(sTemp, ",,", ",")
		Loop
		If iConnectionType = ORACLE Then
			asTemp = Split(sTemp, ",")
			If UBound(asTemp) > 990 Then
				If (Len(oRequest("CalculatePayroll").Item) <> 0) Or (Len(oRequest("ModifyCalculate").Item) <> 0) Then
					sTableName = "sTmpPayroll" & Left(CStr(aLoginComponent(N_USER_ID_LOGIN)),Len("00000")) & Left(GetSerialNumberForDate(""), Len("00000000"))
				ElseIf Len(oRequest("ReportID").Item) <> 0 Then
					sTableName = "sTmp" & oRequest("ReportID").Item & Left(CStr(aLoginComponent(N_USER_ID_LOGIN)),Len("00000")) & Left(GetSerialNumberForDate(""), Len("00000000"))
				Else
					sTableName = "sTmpEmp" & Left(CStr(aLoginComponent(N_USER_ID_LOGIN)),Len("00000")) & Left(GetSerialNumberForDate(""), Len("00000000"))
				End If
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Create Table " & sTableName & " (EmployeeID int NOT NULL)", "ReportsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				If Err.number = 0 Then
					For iIndex = 0 To UBound(asTemp)
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & sTableName & " (EmployeeID) VALUES (" & asTemp(iIndex) & ")", "ReportsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Next
				End If
				sCondition = sCondition & " And (Employees.EmployeeID = (Select Distinct EmployeeID From " & sTableName & " Where (Employees.EmployeeID = " & sTableName & ".EmployeeID)))"
			Else
				sCondition = sCondition & " And (Employees.EmployeeID In (" & sTemp & "))"
			End If
		Else
			sCondition = sCondition & " And (Employees.EmployeeID In (" & sTemp & "))"
		End If
	ElseIf Len(oRequest("EmployeeNumber").Item) > 0 Then
		sCondition = sCondition & " And (Employees.EmployeeNumber = '" & Replace(Right(("000000" & oRequest("EmployeeNumber").Item), Len("000000")), "'", S_WILD_CHAR) & "')"
	ElseIf False Then
		Select Case CInt(oRequest("EmployeeNumberLike").Item)
			Case N_DOES_NOT_CONTENT_LIKE
				sCondition = sCondition & " And (Employees.EmployeeNumber Not Like '" & S_WILD_CHAR & Replace(oRequest("EmployeeNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_STARTS_LIKE
				sCondition = sCondition & " And (Employees.EmployeeNumber Like '" & Replace(oRequest("EmployeeNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_DOES_NOT_START_LIKE
				sCondition = sCondition & " And (Employees.EmployeeNumber Not Like '" & Replace(oRequest("EmployeeNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_ENDS_LIKE
				sCondition = sCondition & " And (Employees.EmployeeNumber Like '" & S_WILD_CHAR & Replace(oRequest("EmployeeNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_DOES_NOT_END_LIKE
				sCondition = sCondition & " And (Employees.EmployeeNumber Not Like '" & S_WILD_CHAR & Replace(oRequest("EmployeeNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_EQUAL_LIKE
				sCondition = sCondition & " And (Employees.EmployeeNumber = '" & Replace(oRequest("EmployeeNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_DIFFERENT_LIKE
				sCondition = sCondition & " And (Employees.EmployeeNumber <> '" & Replace(oRequest("EmployeeNumber").Item, "'", S_WILD_CHAR) & "')"
			Case Else
				sCondition = sCondition & " And (Employees.EmployeeNumber Like '" & S_WILD_CHAR & Replace(oRequest("EmployeeNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
		End Select
	End If
	If Len(oRequest("EmployeeName").Item) > 0 Then
		Select Case CInt(oRequest("EmployeeNameLike").Item)
			Case N_DOES_NOT_CONTENT_LIKE
				sCondition = sCondition & " And ((EmployeeName Not Like '" & S_WILD_CHAR & Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "') Or (EmployeeLastName Not Like '" & S_WILD_CHAR & Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "') Or (EmployeeLastName2 Not Like '" & S_WILD_CHAR & Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "'))"
			Case N_STARTS_LIKE
				sCondition = sCondition & " And ((EmployeeName Like '" & Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "') Or (EmployeeLastName Like '" & Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "') Or (EmployeeLastName2 Like '" & Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "'))"
			Case N_DOES_NOT_START_LIKE
				sCondition = sCondition & " And ((EmployeeName Not Like '" & Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "') Or (EmployeeLastName Not Like '" & Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "') Or (EmployeeLastName2 Not Like '" & Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "'))"
			Case N_ENDS_LIKE
				sCondition = sCondition & " And ((EmployeeName Like '" & S_WILD_CHAR & Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR) & "') Or (EmployeeLastName Like '" & S_WILD_CHAR & Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR) & "') Or (EmployeeLastName2 Like '" & S_WILD_CHAR & Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR) & "'))"
			Case N_DOES_NOT_END_LIKE
				sCondition = sCondition & " And ((EmployeeName Not Like '" & S_WILD_CHAR & Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR) & "') Or (EmployeeLastName Not Like '" & S_WILD_CHAR & Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR) & "') Or (EmployeeLastName2 Not Like '" & S_WILD_CHAR & Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR) & "'))"
			Case N_EQUAL_LIKE
				sCondition = sCondition & " And ((EmployeeName = '" & Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR) & "') Or (EmployeeLastName = '" & Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR) & "') Or (EmployeeLastName2 = '" & Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR) & "'))"
			Case N_DIFFERENT_LIKE
				sCondition = sCondition & " And ((EmployeeName <> '" & Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR) & "') Or (EmployeeLastName <> '" & Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR) & "') Or (EmployeeLastName2 <> '" & Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR) & "'))"
			Case Else
				sCondition = sCondition & " And ((EmployeeName Like '" & S_WILD_CHAR & Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "') Or (EmployeeLastName Like '" & S_WILD_CHAR & Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "') Or (EmployeeLastName2 Like '" & S_WILD_CHAR & Replace(oRequest("EmployeeName").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "'))"
		End Select
	End If
	If Len(oRequest("AbsenceID").Item) > 0 Then
		sCondition = sCondition & " And (EmployeesAbsencesLKP.AbsenceID=" & oRequest("AbsenceID").Item & ")"
		If (CInt(oRequest("AbsenceID").Item) = 35) Or (CInt(oRequest("AbsenceID").Item) = 37) Or (CInt(oRequest("AbsenceID").Item) = 38) Then
			sCondition = sCondition & " And (EmployeesAbsencesLKP.VacationPeriod=" & CStr(oRequest("YearID").Item) & CStr(oRequest("PeriodVacationID").Item) & ")"
		End If
	End If
	If Len(oRequest("AdjustmentPayrollDate").Item) > 0 Then
		sCondition = sCondition & " And (EmployeesAdjustmentsLKP.PayrollDate =" & oRequest("AdjustmentPayrollDate").Item & ")"
	End If
	If Len(oRequest("AppliedDate").Item) > 0 Then
		sCondition = sCondition & " And ((EmployeesAbsencesLKP.AppliedDate=" & oRequest("AppliedDate").Item & ") Or (EmployeesAbsencesLKP.AppliedRemoveDate=" & oRequest("AppliedDate").Item & "))"
	End If
	If Len(oRequest("CompanyID").Item) > 0 Then
		sCondition = sCondition & " And (Companies.CompanyID In (" & oRequest("CompanyID").Item & "))"
	End If
	If Len(oRequest("DocumentDate").Item) > 0 Then
		sCondition = sCondition & " And (EmployeesDocs.DocumentDate=" & oRequest("DocumentDate").Item & ")"
	End If
	If Len(oRequest("ReasonID").Item) > 0 Then
		sCondition = sCondition & " And (Reasons.ReasonID In (" & oRequest("ReasonID").Item & "))"
	End If
	If Len(oRequest("EmployeeTypeID").Item) > 0 Then
        If StrComp(oRequest("Action").Item, "ModifyPayroll", vbBinaryCompare) = 0 Or StrComp(oRequest("Action").Item, "CalculatePayroll", vbBinaryCompare) = 0 Then
            If StrComp(oRequest("EmployeeTypeID").Item, "11", vbBinaryCompare) = 0 Then
                sCondition = sCondition & " And (EmployeesHistoryList.JobID=Jobs.JobID) And (JobTypeID=3)"
            Else
                sCondition = sCondition & " And (EmployeeTypes.EmployeeTypeID In (" & oRequest("EmployeeTypeID").Item & "))"
            End If
        Else
		    sCondition = sCondition & " And (EmployeeTypes.EmployeeTypeID In (" & oRequest("EmployeeTypeID").Item & "))"
        End If
	End If
	If Len(oRequest("PositionTypeID").Item) > 0 Then
		sCondition = sCondition & " And (PositionTypes.PositionTypeID In (" & oRequest("PositionTypeID").Item & "))"
	End If
	If Len(oRequest("ClassificationID").Item) > 0 Then
		sCondition = sCondition & " And (Employees.ClassificationID In (" & oRequest("ClassificationID").Item & "))"
	End If
	If Len(oRequest("CreditsAppliedDate").Item) > 0 Then
		sCondition = sCondition & " And (Credits.StartDate=" & oRequest("CreditsAppliedDate").Item & ")"
	End If
	If Len(oRequest("GroupGradeLevelID").Item) > 0 Then
		sCondition = sCondition & " And (Employees.GroupGradeLevelID In (" & oRequest("GroupGradeLevelID").Item & "))"
	End If
	If Len(oRequest("IntegrationID").Item) > 0 Then
		sCondition = sCondition & " And (Employees.IntegrationID In (" & oRequest("IntegrationID").Item & "))"
	End If
	If Len(oRequest("JourneyID").Item) > 0 Then
		sCondition = sCondition & " And (Journeys.JourneyID In (" & oRequest("JourneyID").Item & "))"
	End If
	If Len(oRequest("ShiftID").Item) > 0 Then
		sCondition = sCondition & " And (Shifts.ShiftID In (" & oRequest("ShiftID").Item & "))"
	End If
	If Len(oRequest("LevelID").Item) > 0 Then
		sCondition = sCondition & " And (Levels.LevelID In (" & oRequest("LevelID").Item & "))"
	End If
	If Len(oRequest("EmployeeStatusID").Item) > 0 Then
		sCondition = sCondition & " And (StatusEmployees.StatusID In (" & oRequest("EmployeeStatusID").Item & "))"
	End If
	If Len(oRequest("RegistrationDate").Item) > 0 Then
		Select Case aReportsComponent(N_ID_REPORTS)
			Case ISSSTE_1108_REPORTS
				sCondition = sCondition & " And (EmployeesConceptsLKP.RegistrationDate=" & oRequest("RegistrationDate").Item & ")"
			Case ISSSTE_1223_REPORTS, ISSSTE_1224_REPORTS
				sCondition = sCondition & " And (EmployeesBeneficiariesLKP.StartDate=" & oRequest("RegistrationDate").Item & ")"
			Case ISSSTE_2431_REPORTS, ISSSTE_2431_REPORTS, ISSSTE_2432_REPORTS, ISSSTE_2432_REPORTS
				sCondition = sCondition & " And (BankAccounts.StartDate=" & oRequest("RegistrationDate").Item & ")"
		End Select
	End If
	If Len(oRequest("PaymentCenterID").Item) > 0 Then
		sCondition = sCondition & " And (PaymentCenters.AreaID In (" & oRequest("PaymentCenterID").Item & "))"
	ElseIf StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
		If InStr(1, ",1003,1006,1007,1027,1401,1474,1475,1476,1490,1494,", "," & oRequest("ReportID").Item & ",") = 0 Then
			If (InStr(1, oRequest, "PaymentCenterID", vbBinaryCompare) > 0) Then sCondition = sCondition & " And (PaymentCenters.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))"
		End If
	End If
	If Len(oRequest("EmployeeEmail").Item) > 0 Then
		Select Case CInt(oRequest("EmployeeEmailLike").Item)
			Case N_DOES_NOT_CONTENT_LIKE
				sCondition = sCondition & " And (EmployeeEmail Not Like '" & S_WILD_CHAR & Replace(oRequest("EmployeeEmail").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_STARTS_LIKE
				sCondition = sCondition & " And (EmployeeEmail Like '" & Replace(oRequest("EmployeeEmail").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_DOES_NOT_START_LIKE
				sCondition = sCondition & " And (EmployeeEmail Not Like '" & Replace(oRequest("EmployeeEmail").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_ENDS_LIKE
				sCondition = sCondition & " And (EmployeeEmail Like '" & S_WILD_CHAR & Replace(oRequest("EmployeeEmail").Item, "'", S_WILD_CHAR) & "')"
			Case N_DOES_NOT_END_LIKE
				sCondition = sCondition & " And (EmployeeEmail Not Like '" & S_WILD_CHAR & Replace(oRequest("EmployeeEmail").Item, "'", S_WILD_CHAR) & "')"
			Case N_EQUAL_LIKE
				sCondition = sCondition & " And (EmployeeEmail = '" & Replace(oRequest("EmployeeEmail").Item, "'", S_WILD_CHAR) & "')"
			Case N_DIFFERENT_LIKE
				sCondition = sCondition & " And (EmployeeEmail <> '" & Replace(oRequest("EmployeeEmail").Item, "'", S_WILD_CHAR) & "')"
			Case Else
				sCondition = sCondition & " And (EmployeeEmail Like '" & S_WILD_CHAR & Replace(oRequest("EmployeeEmail").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
		End Select
	End If
	If Len(oRequest("SocialSecurityNumber").Item) > 0 Then
		Select Case CInt(oRequest("SocialSecurityNumberLike").Item)
			Case N_DOES_NOT_CONTENT_LIKE
				sCondition = sCondition & " And (SocialSecurityNumber Not Like '" & S_WILD_CHAR & Replace(oRequest("SocialSecurityNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_STARTS_LIKE
				sCondition = sCondition & " And (SocialSecurityNumber Like '" & Replace(oRequest("SocialSecurityNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_DOES_NOT_START_LIKE
				sCondition = sCondition & " And (SocialSecurityNumber Not Like '" & Replace(oRequest("SocialSecurityNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_ENDS_LIKE
				sCondition = sCondition & " And (SocialSecurityNumber Like '" & S_WILD_CHAR & Replace(oRequest("SocialSecurityNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_DOES_NOT_END_LIKE
				sCondition = sCondition & " And (SocialSecurityNumber Not Like '" & S_WILD_CHAR & Replace(oRequest("SocialSecurityNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_EQUAL_LIKE
				sCondition = sCondition & " And (SocialSecurityNumber = '" & Replace(oRequest("SocialSecurityNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_DIFFERENT_LIKE
				sCondition = sCondition & " And (SocialSecurityNumber <> '" & Replace(oRequest("SocialSecurityNumber").Item, "'", S_WILD_CHAR) & "')"
			Case Else
				sCondition = sCondition & " And (SocialSecurityNumber Like '" & S_WILD_CHAR & Replace(oRequest("SocialSecurityNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
		End Select
	End If
	If (CInt(oRequest("StartBirthYear").Item) > 0) And (CInt(oRequest("StartBirthMonth").Item) > 0) And (CInt(oRequest("StartBirthDay").Item) > 0) And (CInt(oRequest("EndBirthYear").Item) > 0) And (CInt(oRequest("EndBirthMonth").Item) > 0) And (CInt(oRequest("EndBirthDay").Item) > 0) Then Call GetStartAndEndDatesFromURL("StartBirth", "EndBirth", "BirthDate", True, sCondition)
	If Len(oRequest("CountryID").Item) > 0 Then
		sCondition = sCondition & " And (Countries.CountryID In (" & oRequest("CountryID").Item & "))"
	End If
	If Len(oRequest("RFC").Item) > 0 Then
		Select Case CInt(oRequest("RFCLike").Item)
			Case N_DOES_NOT_CONTENT_LIKE
				sCondition = sCondition & " And (RFC Not Like '" & S_WILD_CHAR & Replace(oRequest("RFC").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_STARTS_LIKE
				sCondition = sCondition & " And (RFC Like '" & Replace(oRequest("RFC").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_DOES_NOT_START_LIKE
				sCondition = sCondition & " And (RFC Not Like '" & Replace(oRequest("RFC").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_ENDS_LIKE
				sCondition = sCondition & " And (RFC Like '" & S_WILD_CHAR & Replace(oRequest("RFC").Item, "'", S_WILD_CHAR) & "')"
			Case N_DOES_NOT_END_LIKE
				sCondition = sCondition & " And (RFC Not Like '" & S_WILD_CHAR & Replace(oRequest("RFC").Item, "'", S_WILD_CHAR) & "')"
			Case N_EQUAL_LIKE
				sCondition = sCondition & " And (RFC = '" & Replace(oRequest("RFC").Item, "'", S_WILD_CHAR) & "')"
			Case N_DIFFERENT_LIKE
				sCondition = sCondition & " And (RFC <> '" & Replace(oRequest("RFC").Item, "'", S_WILD_CHAR) & "')"
			Case Else
				sCondition = sCondition & " And (RFC Like '" & S_WILD_CHAR & Replace(oRequest("RFC").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
		End Select
	End If
	If Len(oRequest("CURP").Item) > 0 Then
		Select Case CInt(oRequest("CURPLike").Item)
			Case N_DOES_NOT_CONTENT_LIKE
				sCondition = sCondition & " And (CURP Not Like '" & S_WILD_CHAR & Replace(oRequest("CURP").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_STARTS_LIKE
				sCondition = sCondition & " And (CURP Like '" & Replace(oRequest("CURP").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_DOES_NOT_START_LIKE
				sCondition = sCondition & " And (CURP Not Like '" & Replace(oRequest("CURP").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_ENDS_LIKE
				sCondition = sCondition & " And (CURP Like '" & S_WILD_CHAR & Replace(oRequest("CURP").Item, "'", S_WILD_CHAR) & "')"
			Case N_DOES_NOT_END_LIKE
				sCondition = sCondition & " And (CURP Not Like '" & S_WILD_CHAR & Replace(oRequest("CURP").Item, "'", S_WILD_CHAR) & "')"
			Case N_EQUAL_LIKE
				sCondition = sCondition & " And (CURP = '" & Replace(oRequest("CURP").Item, "'", S_WILD_CHAR) & "')"
			Case N_DIFFERENT_LIKE
				sCondition = sCondition & " And (CURP <> '" & Replace(oRequest("CURP").Item, "'", S_WILD_CHAR) & "')"
			Case Else
				sCondition = sCondition & " And (CURP Like '" & S_WILD_CHAR & Replace(oRequest("CURP").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
		End Select
	End If
	If Len(oRequest("GenderID").Item) > 0 Then
		sCondition = sCondition & " And (Genders.GenderID In (" & oRequest("GenderID").Item & "))"
	End If
	If Len(oRequest("EmployeeActive").Item) > 0 Then
		sCondition = sCondition & " And (Employees.Active In (" & oRequest("EmployeeActive").Item & "))"
	End If
	If ((CInt(oRequest("StartEmployeeStartYear").Item) > 0) And (CInt(oRequest("StartEmployeeStartMonth").Item) > 0) And (CInt(oRequest("StartEmployeeStartDay").Item) > 0)) Or ((CInt(oRequest("EndEmployeeStartYear").Item) > 0) And (CInt(oRequest("EndEmployeeStartMonth").Item) > 0) And (CInt(oRequest("EndEmployeeStartDay").Item) > 0)) Then Call GetStartAndEndDatesFromURL("StartEmployeeStart", "EndEmployeeStart", "Employees.StartDate", True, sCondition)
	If Len(Replace(oRequest("JobIDs").Item, " ", "")) > 0 Then
		sCondition = sCondition & " And (Jobs.JobID In (" & Replace(oRequest("JobIDs").Item, " ", "") & "))"
	ElseIf Len(oRequest("JobNumber").Item) > 0 Then
		sCondition = sCondition & " And (Jobs.JobNumber = '" & Replace(Right(("000000" & oRequest("JobNumber").Item), Len("000000")), "'", S_WILD_CHAR) & "')"
	ElseIf False Then
		Select Case CInt(oRequest("JobNumberLike").Item)
			Case N_DOES_NOT_CONTENT_LIKE
				sCondition = sCondition & " And (JobNumber Not Like '" & S_WILD_CHAR & Replace(oRequest("JobNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_STARTS_LIKE
				sCondition = sCondition & " And (JobNumber Like '" & Replace(oRequest("JobNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_DOES_NOT_START_LIKE
				sCondition = sCondition & " And (JobNumber Not Like '" & Replace(oRequest("JobNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_ENDS_LIKE
				sCondition = sCondition & " And (JobNumber Like '" & S_WILD_CHAR & Replace(oRequest("JobNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_DOES_NOT_END_LIKE
				sCondition = sCondition & " And (JobNumber Not Like '" & S_WILD_CHAR & Replace(oRequest("JobNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_EQUAL_LIKE
				sCondition = sCondition & " And (JobNumber = '" & Replace(oRequest("JobNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_DIFFERENT_LIKE
				sCondition = sCondition & " And (JobNumber <> '" & Replace(oRequest("JobNumber").Item, "'", S_WILD_CHAR) & "')"
			Case Else
				sCondition = sCondition & " And (JobNumber Like '" & S_WILD_CHAR & Replace(oRequest("JobNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
		End Select
	End If
	If InStr(1, ",1490,1490,", "," & oRequest("ReportID").Item & ",") = 0 Then
		If Len(oRequest("ZoneID").Item) > 0 Then
			If Len(oRequest("ZoneForEmployees").Item) > 0 Then
				If InStr(1, oRequest("ZoneID").Item, ",", vbbinaryCompare) = 0 Then
					sCondition = sCondition & " And (Areas.ZoneID In (" & Replace(oRequest("ZoneID").Item, " ", "") & "))"
				End If
			Else
				If InStr(1, oRequest("ZoneID").Item, ",", vbbinaryCompare) > 0 Then
					If InStr(1, "," & Replace(oRequest("ZoneID").Item, " ", "") & ",", ",38,", vbbinaryCompare) > 0 Then
						sCondition = sCondition & " And ((ParentZones.ZoneID In (" & Replace(oRequest("ZoneID").Item, " ", "") & ")) Or (Areas.AreaPath Like '" & S_WILD_CHAR & ",38," & S_WILD_CHAR & "'))"
					Else
						sCondition = sCondition & " And (ParentZones.ZoneID In (" & Replace(oRequest("ZoneID").Item, " ", "") & "))"
					End If
				ElseIf InStr(1, "," & Replace(oRequest("ZoneID").Item, " ", "") & ",", ",38,", vbbinaryCompare) > 0 Then
					sCondition = sCondition & " And (Areas.AreaPath Like '" & S_WILD_CHAR & ",38," & S_WILD_CHAR & "')"
				Else
					sCondition = sCondition & " And (Zones.ZonePath Like '" & S_WILD_CHAR & "," & oRequest("ZoneID").Item & "," & S_WILD_CHAR & "')"
				End If
			End If
		End If
	End If
	If InStr(1, oRequest("ZoneForPaymentCenterID").Item, ",", vbbinaryCompare) > 0 Then
		If InStr(1, "," & Replace(oRequest("ZoneForPaymentCenterID").Item, " ", "") & ",", ",38,", vbbinaryCompare) > 0 Then
			sCondition = sCondition & " And ((ParentZonesForPaymentCenter.ZoneID In (" & Replace(oRequest("ZoneForPaymentCenterID").Item, " ", "") & ")) Or (PaymentCenters.AreaPath Like '" & S_WILD_CHAR & ",38," & S_WILD_CHAR & "'))"
		Else
			sCondition = sCondition & " And (ParentZonesForPaymentCenter.ZoneID In (" & Replace(oRequest("ZoneForPaymentCenterID").Item, " ", "") & "))"
		End If
	ElseIf InStr(1, "," & Replace(oRequest("ZoneForPaymentCenterID").Item, " ", "") & ",", ",38,", vbbinaryCompare) > 0 Then
		sCondition = sCondition & " And (PaymentCenters.AreaPath Like '" & S_WILD_CHAR & ",38," & S_WILD_CHAR & "')"
	ElseIf Len(oRequest("ZoneForPaymentCenterID").Item) > 0 Then
		sCondition = sCondition & " And (ZonesForPaymentCenter.ZonePath Like '" & S_WILD_CHAR & "," & oRequest("ZoneForPaymentCenterID").Item & "," & S_WILD_CHAR & "')"
	End If
	
	If Len(oRequest("SubAreaID").Item) > 0 Then
		sCondition = sCondition & " And (Areas.AreaID In (" & oRequest("SubAreaID").Item & "))"
'	ElseIf InStr(1, oRequest("AreaID").Item, ",", vbBinaryCompare) > 0 Then
'		sCondition = sCondition & " And (Areas.AreaID In (" & oRequest("AreaID").Item & "))"
	ElseIf Len(oRequest("AreaID").Item) > 0 Then
		sCondition = sCondition & " And (Areas.AreaPath Like '" & S_WILD_CHAR & "," & oRequest("AreaID").Item & "," & S_WILD_CHAR & "')"
	End If
	If InStr(1, ",1003,1006,1007,1027,1401,1475,1476,1490,1494,", "," & oRequest("ReportID").Item & ",") = 0 Then
		If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
			If (InStr(1, oRequest, "SubAreaID", vbBinaryCompare) > 0) Or (InStr(1, oRequest, "AreaID", vbBinaryCompare) > 0) Then sCondition = sCondition & " And (Areas.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))"
		End If
	End If
	If Len(oRequest("PositionID").Item) > 0 Then
		sCondition = sCondition & " And (Positions.PositionID In (" & oRequest("PositionID").Item & "))"
	End If
	If Len(oRequest("JobTypeID").Item) > 0 Then
		sCondition = sCondition & " And (JobTypes.JobTypeID In (" & oRequest("JobTypeID").Item & "))"
	End If
	If Len(oRequest("OccupationTypeID").Item) > 0 Then
		sCondition = sCondition & " And (OccupationTypes.OccupationTypeID In (" & oRequest("OccupationTypeID").Item & "))"
	End If
	If Len(oRequest("GeneratingAreaID").Item) > 0 Then
		sCondition = sCondition & " And (Areas.AreaPath Like '" & S_WILD_CHAR & "," & oRequest("GeneratingAreaID").Item & "," & S_WILD_CHAR & "')"
	End If
	If (CInt(oRequest("StartJobStartYear").Item) > 0) And (CInt(oRequest("StartJobStartMonth").Item) > 0) And (CInt(oRequest("StartJobStartDay").Item) > 0) And (CInt(oRequest("EndJobStartYear").Item) > 0) And (CInt(oRequest("EndJobStartMonth").Item) > 0) And (CInt(oRequest("EndJobStartDay").Item) > 0) Then Call GetStartAndEndDatesFromURL("StartJobStart", "EndJobStart", "Jobs.StartDate", True, sCondition)
	If (CInt(oRequest("StartJobEndYear").Item) > 0) And (CInt(oRequest("StartJobEndMonth").Item) > 0) And (CInt(oRequest("StartJobEndDay").Item) > 0) And (CInt(oRequest("EndJobEndYear").Item) > 0) And (CInt(oRequest("EndJobEndMonth").Item) > 0) And (CInt(oRequest("EndJobEndDay").Item) > 0) Then Call GetStartAndEndDatesFromURL("StartJobEnd", "EndJobEnd", "Jobs.EndDate", False, sCondition)
	If Len(oRequest("JobStatusID").Item) > 0 Then
		sCondition = sCondition & " And (StatusJobs.StatusID In (" & oRequest("JobStatusID").Item & "))"
	End If
	If Len(oRequest("JobActive").Item) > 0 Then
		sCondition = sCondition & " And (Jobs.Active In (" & oRequest("JobActive").Item & "))"
	End If
	If Len(oRequest("AreaCode").Item) > 0 Then
		Select Case CInt(oRequest("AreaCodeLike").Item)
			Case N_DOES_NOT_CONTENT_LIKE
				sCondition = sCondition & " And (AreaCode Not Like '" & S_WILD_CHAR & Replace(oRequest("AreaCode").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_STARTS_LIKE
				sCondition = sCondition & " And (AreaCode Like '" & Replace(oRequest("AreaCode").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_DOES_NOT_START_LIKE
				sCondition = sCondition & " And (AreaCode Not Like '" & Replace(oRequest("AreaCode").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_ENDS_LIKE
				sCondition = sCondition & " And (AreaCode Like '" & S_WILD_CHAR & Replace(oRequest("AreaCode").Item, "'", S_WILD_CHAR) & "')"
			Case N_DOES_NOT_END_LIKE
				sCondition = sCondition & " And (AreaCode Not Like '" & S_WILD_CHAR & Replace(oRequest("AreaCode").Item, "'", S_WILD_CHAR) & "')"
			Case N_EQUAL_LIKE
				sCondition = sCondition & " And (AreaCode = '" & Replace(oRequest("AreaCode").Item, "'", S_WILD_CHAR) & "')"
			Case N_DIFFERENT_LIKE
				sCondition = sCondition & " And (AreaCode <> '" & Replace(oRequest("AreaCode").Item, "'", S_WILD_CHAR) & "')"
			Case Else
				sCondition = sCondition & " And (AreaCode Like '" & S_WILD_CHAR & Replace(oRequest("AreaCode").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
		End Select
	End If
	If Len(oRequest("AreaShortName").Item) > 0 Then
		Select Case CInt(oRequest("AreaShortNameLike").Item)
			Case N_DOES_NOT_CONTENT_LIKE
				sCondition = sCondition & " And (AreaShortName Not Like '" & S_WILD_CHAR & Replace(oRequest("AreaShortName").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_STARTS_LIKE
				sCondition = sCondition & " And (AreaShortName Like '" & Replace(oRequest("AreaShortName").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_DOES_NOT_START_LIKE
				sCondition = sCondition & " And (AreaShortName Not Like '" & Replace(oRequest("AreaShortName").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_ENDS_LIKE
				sCondition = sCondition & " And (AreaShortName Like '" & S_WILD_CHAR & Replace(oRequest("AreaShortName").Item, "'", S_WILD_CHAR) & "')"
			Case N_DOES_NOT_END_LIKE
				sCondition = sCondition & " And (AreaShortName Not Like '" & S_WILD_CHAR & Replace(oRequest("AreaShortName").Item, "'", S_WILD_CHAR) & "')"
			Case N_EQUAL_LIKE
				sCondition = sCondition & " And (AreaShortName = '" & Replace(oRequest("AreaShortName").Item, "'", S_WILD_CHAR) & "')"
			Case N_DIFFERENT_LIKE
				sCondition = sCondition & " And (AreaShortName <> '" & Replace(oRequest("AreaShortName").Item, "'", S_WILD_CHAR) & "')"
			Case Else
				sCondition = sCondition & " And (AreaShortName Like '" & S_WILD_CHAR & Replace(oRequest("AreaShortName").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
		End Select
	End If
	If Len(oRequest("AreaName").Item) > 0 Then
		Select Case CInt(oRequest("AreaNameLike").Item)
			Case N_DOES_NOT_CONTENT_LIKE
				sCondition = sCondition & " And (AreaName Not Like '" & S_WILD_CHAR & Replace(oRequest("AreaName").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_STARTS_LIKE
				sCondition = sCondition & " And (AreaName Like '" & Replace(oRequest("AreaName").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_DOES_NOT_START_LIKE
				sCondition = sCondition & " And (AreaName Not Like '" & Replace(oRequest("AreaName").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_ENDS_LIKE
				sCondition = sCondition & " And (AreaName Like '" & S_WILD_CHAR & Replace(oRequest("AreaName").Item, "'", S_WILD_CHAR) & "')"
			Case N_DOES_NOT_END_LIKE
				sCondition = sCondition & " And (AreaName Not Like '" & S_WILD_CHAR & Replace(oRequest("AreaName").Item, "'", S_WILD_CHAR) & "')"
			Case N_EQUAL_LIKE
				sCondition = sCondition & " And (AreaName = '" & Replace(oRequest("AreaName").Item, "'", S_WILD_CHAR) & "')"
			Case N_DIFFERENT_LIKE
				sCondition = sCondition & " And (AreaName <> '" & Replace(oRequest("AreaName").Item, "'", S_WILD_CHAR) & "')"
			Case Else
				sCondition = sCondition & " And (AreaName Like '" & S_WILD_CHAR & Replace(oRequest("AreaName").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
		End Select
	End If
	If Len(oRequest("AreaTypeID").Item) > 0 Then
		sCondition = sCondition & " And (AreaTypes.AreaTypeID In (" & oRequest("AreaTypeID").Item & "))"
	End If
	If Len(oRequest("ConfineTypeID").Item) > 0 Then
		sCondition = sCondition & " And (ConfineTypes.ConfineTypeID In (" & oRequest("ConfineTypeID").Item & "))"
	End If
	If Len(oRequest("CenterTypeID").Item) > 0 Then
		sCondition = sCondition & " And (CenterTypes.CenterTypeID In (" & oRequest("CenterTypeID").Item & "))"
	End If
	If Len(oRequest("CenterSubtypeID").Item) > 0 Then
		sCondition = sCondition & " And (CenterSubtypes.CenterSubtypeID In (" & oRequest("CenterSubtypeID").Item & "))"
	End If
	If Len(oRequest("AttentionLevelID").Item) > 0 Then
		sCondition = sCondition & " And (AttentionLevels.AttentionLevelID In (" & oRequest("AttentionLevelID").Item & "))"
	End If
	If Len(oRequest("EconomicZoneID").Item) > 0 Then
		sCondition = sCondition & " And (EconomicZones.EconomicZoneID In (" & oRequest("EconomicZoneID").Item & "))"
	End If
	If (CInt(oRequest("StartAreaStartYear").Item) > 0) And (CInt(oRequest("StartAreaStartMonth").Item) > 0) And (CInt(oRequest("StartAreaStartDay").Item) > 0) And (CInt(oRequest("EndAreaStartYear").Item) > 0) And (CInt(oRequest("EndAreaStartMonth").Item) > 0) And (CInt(oRequest("EndAreaStartDay").Item) > 0) Then Call GetStartAndEndDatesFromURL("StartAreaStart", "EndAreaStart", "Areas.StartDate", True, sCondition)
	If (CInt(oRequest("StartAreaEndYear").Item) > 0) And (CInt(oRequest("StartAreaEndMonth").Item) > 0) And (CInt(oRequest("StartAreaEndDay").Item) > 0) And (CInt(oRequest("EndAreaEndYear").Item) > 0) And (CInt(oRequest("EndAreaEndMonth").Item) > 0) And (CInt(oRequest("EndAreaEndDay").Item) > 0) Then Call GetStartAndEndDatesFromURL("StartAreaEnd", "EndAreaEnd", "Areas.EndDate", False, sCondition)
	If Len(oRequest("Jobs").Item) > 0 Then
		sCondition = sCondition & " And (Jobs=" & oRequest("Jobs").Item & ")"
	End If
	If Len(oRequest("TotalJobs").Item) > 0 Then
		sCondition = sCondition & " And (TotalJobs=" & oRequest("TotalJobs").Item & ")"
	End If
	If Len(oRequest("AreaStatusID").Item) > 0 Then
		sCondition = sCondition & " And (StatusAreas.StatusID In (" & oRequest("AreaStatusID").Item & "))"
	End If
	If Len(oRequest("AreaActive").Item) > 0 Then
		sCondition = sCondition & " And (Areas.Active In (" & oRequest("AreaActive").Item & "))"
	End If
	If Len(oRequest("ConceptID").Item) > 0 Then
		sCondition = sCondition & " And (Concepts.ConceptID In (" & oRequest("ConceptID").Item & "))"
	End If
	If StrComp(oRequest("ReportID").Item, "1003", vbBinaryCompare) <> 0 Then
		If Len(oRequest("TotalPaymentMin").Item) > 0 Then
			sCondition = sCondition & " And ((Percepciones.ConceptAmount - Deducciones.ConceptAmount) >= " & Replace(oRequest("TotalPaymentMin").Item, NUMERIC_SEPARATOR, "") & ")"
		End If
		If Len(oRequest("TotalPaymentMax").Item) > 0 Then
			sCondition = sCondition & " And ((Percepciones.ConceptAmount - Deducciones.ConceptAmount) <= " & Replace(oRequest("TotalPaymentMax").Item, NUMERIC_SEPARATOR, "") & ")"
		End If
	Else
		If (Len(oRequest("TotalPaymentMin").Item) > 0) Or (Len(oRequest("TotalPaymentMax").Item) > 0) Then
			sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID In (Select EmployeeID From Payroll_YYYYMMDD Where (ConceptID=0)"
				If Len(oRequest("TotalPaymentMin").Item) > 0 Then sCondition = sCondition & " And (ConceptAmount>=" & Replace(oRequest("TotalPaymentMin").Item, NUMERIC_SEPARATOR, "") & ")"
				If Len(oRequest("TotalPaymentMax").Item) > 0 Then sCondition = sCondition & " And (ConceptAmount<=" & Replace(oRequest("TotalPaymentMax").Item, NUMERIC_SEPARATOR, "") & ")"
			sCondition = sCondition & "))"
		End If
	End If
	If Len(oRequest("BankID").Item) > 0 Then
		If (StrComp(oRequest("ReportID").Item, "1470", vbBinaryCompare) = 0) Then
			If (StrComp(oRequest("CheckConceptID").Item, "155", vbBinaryCompare) <> 0) And ((StrComp(oRequest("CheckConceptID").Item, "69", vbBinaryCompare) <> 0)) Then
				sCondition = sCondition & " And (Banks.BankID In (" & oRequest("BankID").Item & "))"
			End If
		Else
			If (StrComp(oRequest("ReportID").Item, "1475", vbBinaryCompare) = 0) Then
				If (StrComp(oRequest("CheckConceptID").Item, "155", vbBinaryCompare) <> 0) And ((StrComp(oRequest("CheckConceptID").Item, "69", vbBinaryCompare) <> 0)) Then
					sCondition = sCondition & " And (Banks.BankID In (" & oRequest("BankID").Item & "))"
				End If
			Else
				sCondition = sCondition & " And (Banks.BankID In (" & oRequest("BankID").Item & "))"
			End If
		End If
	End If
	'If Len(oRequest("MedicalAreasTypeID").Item) > 0 Then
	'	sCondition = sCondition & " And (MedicalAreas.MedicalAreasTypeID = " & oRequest("MedicalAreasTypeID").Item & ")"
	'End If
	If Len(oRequest("DocumentForLicenseNumber").Item) > 0 Then
		Select Case CInt(oRequest("DocumentForLicenseNumberLike").Item)
			Case N_DOES_NOT_CONTENT_LIKE
				sCondition = sCondition & " And (DocumentsForLicenses.DocumentForLicenseNumber Not Like '" & S_WILD_CHAR & Replace(oRequest("DocumentForLicenseNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_STARTS_LIKE
				sCondition = sCondition & " And (DocumentsForLicenses.DocumentForLicenseNumber Like '" & Replace(oRequest("DocumentForLicenseNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_DOES_NOT_START_LIKE
				sCondition = sCondition & " And (DocumentsForLicenses.DocumentForLicenseNumber Not Like '" & Replace(oRequest("DocumentForLicenseNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_ENDS_LIKE
				sCondition = sCondition & " And (DocumentsForLicenses.DocumentForLicenseNumber Like '" & S_WILD_CHAR & Replace(oRequest("DocumentForLicenseNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_DOES_NOT_END_LIKE
				sCondition = sCondition & " And (DocumentsForLicenses.DocumentForLicenseNumber Not Like '" & S_WILD_CHAR & Replace(oRequest("DocumentForLicenseNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_EQUAL_LIKE
				sCondition = sCondition & " And (DocumentsForLicenses.DocumentForLicenseNumber = '" & Replace(oRequest("DocumentForLicenseNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_DIFFERENT_LIKE
				sCondition = sCondition & " And (DocumentsForLicenses.DocumentForLicenseNumber <> '" & Replace(oRequest("DocumentForLicenseNumber").Item, "'", S_WILD_CHAR) & "')"
			Case Else
				sCondition = sCondition & " And (DocumentsForLicenses.DocumentForLicenseNumber Like '" & S_WILD_CHAR & Replace(oRequest("DocumentForLicenseNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
		End Select
	End If
	If (Len(oRequest("StartDate").Item) > 0) Or (Len(oRequest("StartYear").Item) > 0) Or (Len(oRequest("StartMonth").Item) > 0) Or (Len(oRequest("StartDay").Item) > 0) Then Call GetStartAndEndDatesFromURL("Start", "End", "XXXDate", True, sCondition)
	If Len(oRequest("RequestNumber").Item) > 0 Then
		Select Case CInt(oRequest("RequestNumberLike").Item)
			Case N_DOES_NOT_CONTENT_LIKE
				sCondition = sCondition & " And (DocumentsForLicenses.RequestNumber Not Like '" & S_WILD_CHAR & Replace(oRequest("RequestNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_STARTS_LIKE
				sCondition = sCondition & " And (DocumentsForLicenses.RequestNumber Like '" & Replace(oRequest("RequestNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_DOES_NOT_START_LIKE
				sCondition = sCondition & " And (DocumentsForLicenses.RequestNumber Not Like '" & Replace(oRequest("RequestNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_ENDS_LIKE
				sCondition = sCondition & " And (DocumentsForLicenses.RequestNumber Like '" & S_WILD_CHAR & Replace(oRequest("RequestNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_DOES_NOT_END_LIKE
				sCondition = sCondition & " And (DocumentsForLicenses.RequestNumber Not Like '" & S_WILD_CHAR & Replace(oRequest("RequestNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_EQUAL_LIKE
				sCondition = sCondition & " And (DocumentsForLicenses.RequestNumber = '" & Replace(oRequest("RequestNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_DIFFERENT_LIKE
				sCondition = sCondition & " And (DocumentsForLicenses.RequestNumber <> '" & Replace(oRequest("RequestNumber").Item, "'", S_WILD_CHAR) & "')"
			Case Else
				sCondition = sCondition & " And (DocumentsForLicenses.RequestNumber Like '" & S_WILD_CHAR & Replace(oRequest("RequestNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
		End Select
	End If
	If Len(oRequest("DocumentForCancelLicenseNumber").Item) > 0 Then
		Select Case CInt(oRequest("DocumentForCancelLicenseNumberLike").Item)
			Case N_DOES_NOT_CONTENT_LIKE
				sCondition = sCondition & " And (DocumentsForLicenses.DocumentForCancelLicenseNumber Not Like '" & S_WILD_CHAR & Replace(oRequest("DocumentForCancelLicenseNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_STARTS_LIKE
				sCondition = sCondition & " And (DocumentsForLicenses.DocumentForCancelLicenseNumber Like '" & Replace(oRequest("DocumentForCancelLicenseNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_DOES_NOT_START_LIKE
				sCondition = sCondition & " And (DocumentsForLicenses.DocumentForCancelLicenseNumber Not Like '" & Replace(oRequest("DocumentForCancelLicenseNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_ENDS_LIKE
				sCondition = sCondition & " And (DocumentsForLicenses.DocumentForCancelLicenseNumber Like '" & S_WILD_CHAR & Replace(oRequest("DocumentForCancelLicenseNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_DOES_NOT_END_LIKE
				sCondition = sCondition & " And (DocumentsForLicenses.DocumentForCancelLicenseNumber Not Like '" & S_WILD_CHAR & Replace(oRequest("DocumentForCancelLicenseNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_EQUAL_LIKE
				sCondition = sCondition & " And (DocumentsForLicenses.DocumentForCancelLicenseNumber = '" & Replace(oRequest("DocumentForCancelLicenseNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_DIFFERENT_LIKE
				sCondition = sCondition & " And (DocumentsForLicenses.DocumentForCancelLicenseNumber <> '" & Replace(oRequest("DocumentForCancelLicenseNumber").Item, "'", S_WILD_CHAR) & "')"
			Case Else
				sCondition = sCondition & " And (DocumentsForLicenses.DocumentForCancelLicenseNumber Like '" & S_WILD_CHAR & Replace(oRequest("DocumentForCancelLicenseNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
		End Select
	End If

	If Len(oRequest("PaperworkNumber").Item) > 0 Then
		asTemp = Split(Replace(Replace(oRequest("PaperworkNumber").Item, "'", ""), " ", ""), ",")
		For iIndex = 0 To UBound(asTemp)
			If InStr(1, asTemp(iIndex), "-", vbBinaryCompare) > 0 Then
				asTemp(iIndex) = Split(asTemp(iIndex), "-")
				sTemp = ""
				For jIndex = CLng(asTemp(iIndex)(0)) To CLng(asTemp(iIndex)(1))
					sTemp = sTemp & jIndex & ","
				Next
				asTemp(iIndex) = sTemp & "-2"
			End If
		Next
		sCondition = sCondition & " And (PaperworkNumber In (" & Join(asTemp, ",") & "))"
	ElseIf False Then
		Select Case CInt(oRequest("PaperworkNumberLike").Item)
			Case N_DOES_NOT_CONTENT_LIKE
				sCondition = sCondition & " And (PaperworkNumber Not Like '" & S_WILD_CHAR & Replace(oRequest("PaperworkNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_STARTS_LIKE
				sCondition = sCondition & " And (PaperworkNumber Like '" & Replace(oRequest("PaperworkNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_DOES_NOT_START_LIKE
				sCondition = sCondition & " And (PaperworkNumber Not Like '" & Replace(oRequest("PaperworkNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_ENDS_LIKE
				sCondition = sCondition & " And (PaperworkNumber Like '" & S_WILD_CHAR & Replace(oRequest("PaperworkNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_DOES_NOT_END_LIKE
				sCondition = sCondition & " And (PaperworkNumber Not Like '" & S_WILD_CHAR & Replace(oRequest("PaperworkNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_EQUAL_LIKE
				sCondition = sCondition & " And (PaperworkNumber = '" & Replace(oRequest("PaperworkNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_DIFFERENT_LIKE
				sCondition = sCondition & " And (PaperworkNumber <> '" & Replace(oRequest("PaperworkNumber").Item, "'", S_WILD_CHAR) & "')"
			Case Else
				sCondition = sCondition & " And (PaperworkNumber Like '" & S_WILD_CHAR & Replace(oRequest("PaperworkNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
		End Select
	End If
	If Len(oRequest("FilterStartNumber").Item) > 0 Then
		sCondition = sCondition & " And (PaperworkNumber>=" & Replace(oRequest("FilterStartNumber").Item, "´", "") & ")"
	End If
	If Len(oRequest("FilterEndNumber").Item) > 0 Then
		sCondition = sCondition & " And (PaperworkNumber<=" & Replace(oRequest("FilterEndNumber").Item, "´", "") & ")"
	End If
	If ((CInt(oRequest("PaperworkStartStartYear").Item) > 0) And (CInt(oRequest("PaperworkStartStartMonth").Item) > 0) And (CInt(oRequest("PaperworkStartStartDay").Item) > 0)) Or ((CInt(oRequest("PaperworkStartEndYear").Item) > 0) And (CInt(oRequest("PaperworkStartEndMonth").Item) > 0) And (CInt(oRequest("PaperworkStartEndDay").Item) > 0)) Then Call GetStartAndEndDatesFromURL("PaperworkStartStart", "PaperworkStartEnd", "Paperworks.StartDate", True, sCondition)
	If ((CInt(oRequest("PaperworkEstimatedStartYear").Item) > 0) And (CInt(oRequest("PaperworkEstimatedStartMonth").Item) > 0) And (CInt(oRequest("PaperworkEstimatedStartDay").Item) > 0)) Or ((CInt(oRequest("PaperworkEstimatedEndYear").Item) > 0) And (CInt(oRequest("PaperworkEstimatedEndMonth").Item) > 0) And (CInt(oRequest("PaperworkEstimatedEndDay").Item) > 0)) Then Call GetStartAndEndDatesFromURL("PaperworkEstimatedStart", "PaperworkEstimatedEnd", "Paperworks.EstimatedDate", True, sCondition)
	If ((CInt(oRequest("PaperworkEndStartYear").Item) > 0) And (CInt(oRequest("PaperworkEndStartMonth").Item) > 0) And (CInt(oRequest("PaperworkEndStartDay").Item) > 0)) Or ((CInt(oRequest("PaperworkEndEndYear").Item) > 0) And (CInt(oRequest("PaperworkEndEndMonth").Item) > 0) And (CInt(oRequest("PaperworkEndEndDay").Item) > 0)) Then Call GetStartAndEndDatesFromURL("PaperworkEndStart", "PaperworkEndEnd", "Paperworks.EndDate", True, sCondition)
	If Len(oRequest("PpwkDocumentNumber").Item) > 0 Then
		Select Case CInt(oRequest("PpwkDocumentNumberLike").Item)
			Case N_DOES_NOT_CONTENT_LIKE
				sCondition = sCondition & " And (DocumentNumber Not Like '" & S_WILD_CHAR & Replace(oRequest("PpwkDocumentNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_STARTS_LIKE
				sCondition = sCondition & " And (DocumentNumber Like '" & Replace(oRequest("PpwkDocumentNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_DOES_NOT_START_LIKE
				sCondition = sCondition & " And (DocumentNumber Not Like '" & Replace(oRequest("PpwkDocumentNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_ENDS_LIKE
				sCondition = sCondition & " And (DocumentNumber Like '" & S_WILD_CHAR & Replace(oRequest("PpwkDocumentNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_DOES_NOT_END_LIKE
				sCondition = sCondition & " And (DocumentNumber Not Like '" & S_WILD_CHAR & Replace(oRequest("PpwkDocumentNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_EQUAL_LIKE
				sCondition = sCondition & " And (DocumentNumber = '" & Replace(oRequest("PpwkDocumentNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_DIFFERENT_LIKE
				sCondition = sCondition & " And (DocumentNumber <> '" & Replace(oRequest("PpwkDocumentNumber").Item, "'", S_WILD_CHAR) & "')"
			Case Else
				sCondition = sCondition & " And (DocumentNumber Like '" & S_WILD_CHAR & Replace(oRequest("PpwkDocumentNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
		End Select
	End If
	If Len(oRequest("PaperworkTypeID").Item) > 0 Then
		sCondition = sCondition & " And (PaperworkTypes.PaperworkTypeID In (" & oRequest("PaperworkTypeID").Item & "))"
	End If
	If Len(oRequest("OwnerNumber").Item) > 0 Then
		Select Case CInt(oRequest("OwnerNumberLike").Item)
			Case N_DOES_NOT_CONTENT_LIKE
				sCondition = sCondition & " And (PaperworkOwners.EmployeeID Not Like '" & S_WILD_CHAR & Replace(oRequest("OwnerNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_STARTS_LIKE
				sCondition = sCondition & " And (PaperworkOwners.EmployeeID Like '" & Replace(oRequest("OwnerNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_DOES_NOT_START_LIKE
				sCondition = sCondition & " And (PaperworkOwners.EmployeeID Not Like '" & Replace(oRequest("OwnerNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
			Case N_ENDS_LIKE
				sCondition = sCondition & " And (PaperworkOwners.EmployeeID Like '" & S_WILD_CHAR & Replace(oRequest("OwnerNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_DOES_NOT_END_LIKE
				sCondition = sCondition & " And (PaperworkOwners.EmployeeID Not Like '" & S_WILD_CHAR & Replace(oRequest("OwnerNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_EQUAL_LIKE
				sCondition = sCondition & " And (PaperworkOwners.EmployeeID = '" & Replace(oRequest("OwnerNumber").Item, "'", S_WILD_CHAR) & "')"
			Case N_DIFFERENT_LIKE
				sCondition = sCondition & " And (PaperworkOwners.EmployeeID <> '" & Replace(oRequest("OwnerNumber").Item, "'", S_WILD_CHAR) & "')"
			Case Else
				sCondition = sCondition & " And (PaperworkOwners.EmployeeID Like '" & S_WILD_CHAR & Replace(oRequest("OwnerNumber").Item, "'", S_WILD_CHAR) & S_WILD_CHAR & "')"
		End Select
	End If
	If Len(oRequest("OwnerIDs").Item) > 0 Then
		sCondition = sCondition & " And (PaperworkOwnersLKP.OwnerID In (" & oRequest("OwnerIDs").Item & "))"
	End If
	If Len(oRequest("PaperworkStatusID").Item) > 0 Then
		sCondition = sCondition & " And (StatusPaperworks.StatusID In (" & oRequest("PaperworkStatusID").Item & "))"
	End If
	If Len(oRequest("SubjectTypeID").Item) > 0 Then
		sCondition = sCondition & " And (Paperworks.SubjectTypeID In (" & Replace(oRequest("SubjectTypeID").Item, ", ", ",") & "))"
	End If
	If Len(oRequest("PriorityID").Item) > 0 Then
		sCondition = sCondition & " And (Paperworks.PriorityID In (" & oRequest("PriorityID").Item & "))"
	End If

	If Len(oRequest("CourseID").Item) > 0 Then
		sCondition = sCondition & " And (SADE_Curso.ID_Curso In (" & oRequest("CourseID").Item & "))"
	End If
	If Len(oRequest("ProfileID").Item) > 0 Then
		sCondition = sCondition & " And (SADE_Curso.MostrarEvaluaciones In (" & oRequest("ProfileID").Item & "))"
	End If

	If Len(oRequest("BudgetAreaID").Item) > 0 Then
		sCondition = sCondition & " And (BudgetsMoney.AreaID In (" & oRequest("BudgetAreaID").Item & "))"
	End If
	If Len(oRequest("BudgetCompanyID").Item) > 0 Then
		If StrComp(oRequest("BudgetCompanyID").Item, "-1", vbBinaryCompare) = 0 Then
			sCondition = sCondition & " And (BudgetsMoney.AreaID Not In (170,500,700))"
		Else
			sCondition = sCondition & " And (BudgetsMoney.AreaID In (" & oRequest("BudgetCompanyID").Item & "))"
		End If
	End If
	If Len(oRequest("BudgetFundID").Item) > 0 Then
		sCondition = sCondition & " And (BudgetsMoney.FundID In (" & oRequest("BudgetFundID").Item & "))"
	End If
	If Len(oRequest("BudgetDutyID").Item) > 0 Then
		sCondition = sCondition & " And (BudgetsMoney.DutyID In (" & oRequest("BudgetDutyID").Item & "))"
	End If
	If Len(oRequest("BudgetActiveDutyID").Item) > 0 Then
		sCondition = sCondition & " And (BudgetsMoney.ActiveDutyID In (" & oRequest("BudgetActiveDutyID").Item & "))"
	End If
	If Len(oRequest("BudgetSpecificDutyID").Item) > 0 Then
		sCondition = sCondition & " And (BudgetsMoney.SpecificDutyID In (" & oRequest("BudgetSpecificDutyID").Item & "))"
	End If
	If Len(oRequest("BudgetProgramID").Item) > 0 Then
		sCondition = sCondition & " And (BudgetsMoney.ProgramID In (" & oRequest("BudgetProgramID").Item & "))"
	End If
	If Len(oRequest("BudgetRegionID").Item) > 0 Then
		sCondition = sCondition & " And (BudgetsMoney.RegionID In (" & oRequest("BudgetRegionID").Item & "))"
	End If
	If Len(oRequest("BudgetUR").Item) > 0 Then
		sCondition = sCondition & " And (BudgetsMoney.BudgetUR In (" & oRequest("BudgetUR").Item & "))"
	End If
	If Len(oRequest("BudgetCT").Item) > 0 Then
		sCondition = sCondition & " And (BudgetsMoney.BudgetCT In (" & oRequest("BudgetCT").Item & "))"
	End If
	If Len(oRequest("BudgetAUX").Item) > 0 Then
		sCondition = sCondition & " And (BudgetsMoney.BudgetAUX In (" & oRequest("BudgetAUX").Item & "))"
	End If
	If Len(oRequest("LocationID").Item) > 0 Then
		sCondition = sCondition & " And (BudgetsMoney.LocationID In (" & oRequest("LocationID").Item & "))"
	End If
	If Len(oRequest("BudgetBudgetID1").Item) > 0 Then
		sCondition = sCondition & " And (BudgetsMoney.BudgetID1 In (" & oRequest("BudgetBudgetID1").Item & "))"
	End If
	If Len(oRequest("BudgetBudgetID2").Item) > 0 Then
		sCondition = sCondition & " And (BudgetsMoney.BudgetID2 In (" & oRequest("BudgetBudgetID2").Item & "))"
	End If
	If Len(oRequest("BudgetBudgetID3").Item) > 0 Then
		sCondition = sCondition & " And (BudgetsMoney.BudgetID3 In (" & oRequest("BudgetBudgetID3").Item & "))"
	End If
	If Len(oRequest("BudgetConfineTypeID").Item) > 0 Then
		sCondition = sCondition & " And (BudgetsMoney.ConfineTypeID In (" & oRequest("BudgetConfineTypeID").Item & "))"
	End If
	If Len(oRequest("BudgetActivityID1").Item) > 0 Then
		sCondition = sCondition & " And (BudgetsMoney.ActivityID1 In (" & oRequest("BudgetActivityID1").Item & "))"
	End If
	If Len(oRequest("BudgetActivityID2").Item) > 0 Then
		sCondition = sCondition & " And (BudgetsMoney.ActivityID2 In (" & oRequest("BudgetActivityID2").Item & "))"
	End If
	If Len(oRequest("BudgetProcessID").Item) > 0 Then
		sCondition = sCondition & " And (BudgetsMoney.ProcessID In (" & oRequest("BudgetProcessID").Item & "))"
	End If
	If Len(oRequest("BudgetYear").Item) > 0 Then
		sCondition = sCondition & " And (BudgetsMoney.BudgetYear In (" & oRequest("BudgetYear").Item & "))"
	End If
	If Len(oRequest("BudgetMonth").Item) > 0 Then
		sCondition = sCondition & " And (BudgetsMoney.BudgetMonth In (" & oRequest("BudgetMonth").Item & "))"
	End If
	If Len(oRequest("BudgetPositionID").Item) > 0 Then
		sCondition = sCondition & " And (BudgetsPositions.PositionID In (" & oRequest("BudgetPositionID").Item & "))"
	End If
	If Len(oRequest("CreditTypeID").Item) > 0 Then
		sCondition = sCondition & " And (Credits.CreditTypeID In (" & oRequest("CreditTypeID").Item & "))"
	End If
	If Len(oRequest("UploadedFileName").Item) > 0 Then
		Select Case aReportsComponent(N_ID_REPORTS)
			Case ISSSTE_1222_REPORTS
				sCondition = sCondition & " And (UploadThirdCreditsRejected.UploadedFileName In ('" & oRequest("UploadedFileName").Item & "'))"
			Case Else
				sCondition = sCondition & " And (Credits.Active=0) And (Credits.UploadedFileName In ('" & oRequest("UploadedFileName").Item & "'))"
		End Select
	End If

	If (CInt(oRequest("StartLogYear").Item) > 0) And (CInt(oRequest("StartLogMonth").Item) > 0) And (CInt(oRequest("StartLogDay").Item) > 0) And (CInt(oRequest("EndLogYear").Item) > 0) And (CInt(oRequest("EndLogMonth").Item) > 0) And (CInt(oRequest("EndLogDay").Item) > 0) Then Call GetStartAndEndDatesFromURL("StartLog", "EndLog", "LogDate", True, sCondition)
	If Len(oRequest("PayrollID").Item) > 0 Then
		lPayrollID = CLng(oRequest("PayrollID").Item)
		If lPayrollID > -1 Then Call GetNameFromTable(oADODBConnection, "ForPayrollID", lPayrollID, "", "", lForPayrollID, sErrorDescription)
	End If
	If Len(oRequest("CheckConceptID").Item) > 0 Then
		Select Case CLng(oRequest("CheckConceptID").Item)
			Case -2
				'sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID<600000) And (BankAccounts.AccountNumber<>'.')"
                sCondition = sCondition & " And (BankAccounts.AccountNumber<>'.')"
			Case -1
				'sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID<600000) And (BankAccounts.AccountNumber='.')"
                sCondition = sCondition & " And (BankAccounts.AccountNumber='.')"
			Case 0
				'sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID<600000)"
                sCondition = sCondition & " "
			Case 11
				'sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID>=600000) And (EmployeesHistoryList.EmployeeID<700000)"
			Case 69
				'sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID>=700000) And (EmployeesHistoryList.EmployeeID<800000)"
			Case 155
				'sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID>=1000000) And (EmployeesHistoryList.EmployeeID<1100000)"
			Case Else
				sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID<600000)"
		End Select
	End If
	If Len(oRequest("CheckNumberMin").Item) > 0 Then
		Select Case iConnectionType
			Case ACCESS, ACCESS_DSN
				sCondition = sCondition & " And (CDbl(Payments.CheckNumber)>=" & oRequest("CheckNumberMin").Item & ")"
			Case ORACLE
				sCondition = sCondition & " And (to_number(Payments.CheckNumber)>=" & oRequest("CheckNumberMin").Item & ")"
			Case Else
				sCondition = sCondition & " And (Cast(Payments.CheckNumber As float)>=" & oRequest("CheckNumberMin").Item & ")"
		End Select
	End If
	If Len(oRequest("CheckNumberMax").Item) > 0 Then
		Select Case iConnectionType
			Case ACCESS, ACCESS_DSN
				sCondition = sCondition & " And (CDbl(Payments.CheckNumber)<=" & oRequest("CheckNumberMax").Item & ")"
			Case ORACLE
				sCondition = sCondition & " And (to_number(Payments.CheckNumber)<=" & oRequest("CheckNumberMax").Item & ")"
			Case Else
				sCondition = sCondition & " And (Cast(Payments.CheckNumber As float)<=" & oRequest("CheckNumberMax").Item & ")"
		End Select
	End If
	If Len(oRequest("HasAlimony").Item) > 0 Then
		sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID In (Select EmployeeID From Payroll_YYYYMMDD Where (ConceptID=70) And (ConceptAmount>0)))"
	End If
	If Len(oRequest("HasCredits").Item) > 0 Then
		sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID In (Select EmployeeID From Payroll_YYYYMMDD Where (ConceptID In (57,58,59,126,64,83)) And (ConceptAmount>0)))"
	End If
	If Len(oRequest("EmployeesAbsenceActive").Item) > 0 Then
		If (CInt(oRequest("EmployeesAbsenceActive").Item) = 0) Or (CInt(oRequest("EmployeesAbsenceActive").Item) = 1) Then
			sCondition = sCondition & " And (EmployeesAbsencesLKP.Active In (" & oRequest("EmployeesAbsenceActive").Item & "))"
		Else
			sCondition = sCondition & " And (EmployeesAbsencesLKP.Active<0)"
		End If
	End If
	If Len(oRequest("EmployeesConceptActive").Item) > 0 Then
		If CInt(oRequest("EmployeesConceptActive").Item) = 2 Then
			sCondition = sCondition & " And (EmployeesConceptsLKP.Active In (" & oRequest("EmployeesConceptActive").Item & ")) And (EmployeesConceptsLKP.EndDate=0)"
		Else
			sCondition = sCondition & " And (EmployeesConceptsLKP.Active In (" & oRequest("EmployeesConceptActive").Item & "))"
		End If
	End If
	If Len(oRequest("BankAccountsActive").Item) > 0 Then
		sCondition = sCondition & " And (BankAccounts.Active=" & CInt(oRequest("BankAccountsActive").Item) & ")"
	End If

	GetConditionFromURL = Err.number
	Err.Clear
End Function

Function GetDBFieldsNames(oRequest, sFlags, sCondition, sFieldNames, sTableNames, sJoinCondition, sSortFields)
'************************************************************
'Purpose: To get the fields for the query from the database
'Inputs:  oRequest, sFlags, sCondition
'Outputs: sFieldNames, sTableNames, sJoinCondition, sSortFields
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetDBFieldsNames"
	Dim oItem

	sFieldNames = ""
	sTableNames = " "
	sJoinCondition = ""
	If Len(oRequest("Template").Item) > 0 Then
		For Each oItem In oRequest("Template")
			Call GetFlagFieldName(CLng(oItem), sFieldNames, sTableNames, sJoinCondition, sSortFields)
		Next
	Else
		Call GetFlagFieldName(sFlags, sFieldNames, sTableNames, sJoinCondition, sSortFields)
	End If

	If ((InStr(1, sCondition, "=Areas.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(Areas.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " Areas,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "Areas, "
		If (InStr(1, sJoinCondition, "=Areas.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(Areas.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Jobs.AreaID=Areas.AreaID) And "
	End If
	If ((InStr(1, sCondition, "=AreaTypes.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(AreaTypes.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " AreaTypes,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "AreaTypes, "
		If (InStr(1, sJoinCondition, "=AreaTypes.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(AreaTypes.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Areas.AreaTypeID=AreaTypes.AreaTypeID) And "
	End If
	If ((InStr(1, sCondition, "=CenterSubtypes.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(CenterSubtypes.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " CenterSubtypes,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "CenterSubtypes, "
		If (InStr(1, sJoinCondition, "=CenterSubtypes.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(CenterSubtypes.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Areas.CenterSubtypesID=CenterSubtypes.CenterSubtypesID) And "
	End If
	If ((InStr(1, sCondition, "=CenterTypes.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(CenterTypes.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " CenterTypes,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "CenterTypes, "
		If (InStr(1, sJoinCondition, "=CenterTypes.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(CenterTypes.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Areas.CenterTypeID=CenterTypes.CenterTypeID) And "
	End If
	If ((InStr(1, sCondition, "=Companies.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(Companies.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " Companies,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "Companies, "
		If (InStr(1, sJoinCondition, "=Companies.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(Companies.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Employees.CompanyID=Companies.CompanyID) And "
	End If
	If ((InStr(1, sCondition, "=ConfineTypes.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(ConfineTypes.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " ConfineTypes,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "ConfineTypes, "
		If (InStr(1, sJoinCondition, "=ConfineTypes.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(ConfineTypes.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Areas.ConfineTypeID=ConfineTypes.ConfineTypeID) And "
	End If
	If ((InStr(1, sCondition, "=Countries.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(Countries.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " Countries,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "Countries, "
		If (InStr(1, sJoinCondition, "=Countries.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(Countries.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Employees.CountryID=Countries.CountryID) And "
	End If
	If ((InStr(1, sCondition, "=EconomicZones.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(EconomicZones.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " EconomicZones,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "EconomicZones, "
		If (InStr(1, sJoinCondition, "=EconomicZones.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(EconomicZones.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Areas.EconomicZoneID=EconomicZones.EconomicZoneID) And "
	End If
	If ((InStr(1, sCondition, "=Employees.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(Employees.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " Employees,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "Employees, "
	End If
	If ((InStr(1, sCondition, "=EmployeeTypes.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(EmployeeTypes.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " EmployeeTypes,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "EmployeeTypes, "
		If (InStr(1, sJoinCondition, "=EmployeeTypes.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(EmployeeTypes.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Employees.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And "
	End If
	If ((InStr(1, sCondition, "=Genders.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(Genders.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " Genders,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "Genders, "
		If (InStr(1, sJoinCondition, "=Genders.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(Genders.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Employees.GenderID=Genders.GenderID) And "
	End If
	If ((InStr(1, sCondition, "=GroupGradeLevels.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(GroupGradeLevels.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " GroupGradeLevels,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "GroupGradeLevels, "
		If (InStr(1, sJoinCondition, "=GroupGradeLevels.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(GroupGradeLevels.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Employees.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And "
	End If
	If ((InStr(1, sCondition, "=Jobs.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(Jobs.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " Jobs,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "Jobs, "
	End If
	If ((InStr(1, sCondition, "=JobTypes.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(JobTypes.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " JobTypes,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "JobTypes, "
		If (InStr(1, sJoinCondition, "=JobTypes.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(JobTypes.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Jobs.JobTypeID=JobTypes.JobTypeID) And "
	End If
	If ((InStr(1, sCondition, "=Journeys.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(Journeys.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " Journeys,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "Journeys, "
		If (InStr(1, sJoinCondition, "=Journeys.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(Journeys.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Employees.JourneyID=Journeys.JourneyID) And "
	End If
	If ((InStr(1, sCondition, "=Levels.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(Levels.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " Levels,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "Levels, "
		If (InStr(1, sJoinCondition, "=Levels.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(Levels.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Employees.LevelID=Levels.LevelID) And "
	End If
	If ((InStr(1, sCondition, "=Shifts.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(Shifts.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " Shifts,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "Shifts, "
		If (InStr(1, sJoinCondition, "=Shifts.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(Shifts.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Employees.ShiftID=Shifts.ShiftID) And "
	End If
	If ((InStr(1, sCondition, "=OccupationTypes.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(OccupationTypes.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " OccupationTypes,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "OccupationTypes, "
		If (InStr(1, sJoinCondition, "=OccupationTypes.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(OccupationTypes.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Jobs.OccupationTypeID=OccupationTypes.OccupationTypeID) And "
	End If
	If ((InStr(1, sCondition, "=PaymentCenters.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(PaymentCenters.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " Areas As PaymentCenters,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "Areas As PaymentCenters, "
		If (InStr(1, sJoinCondition, "=PaymentCenters.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(PaymentCenters.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Areas.PaymentCenterID=PaymentCenters.AreaID) And "
	End If
	If ((InStr(1, sCondition, "=Positions.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(Positions.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " Positions,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "Positions, "
		If (InStr(1, sJoinCondition, "=Positions.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(Positions.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Jobs.PositionID=Positions.PositionID) And "
	End If
	If ((InStr(1, sCondition, "=PositionTypes.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(PositionTypes.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " PositionTypes,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "PositionTypes, "
		If (InStr(1, sJoinCondition, "=PositionTypes.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(PositionTypes.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Positions.PositionTypeID=PositionTypes.PositionTypeID) And "
	End If
	If ((InStr(1, sCondition, "=StatusAreas.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(StatusAreas.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " StatusAreas,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "StatusAreas, "
		If (InStr(1, sJoinCondition, "=StatusAreas.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(StatusAreas.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Areas.StatusID=StatusAreas.StatusID) And "
	End If
	If ((InStr(1, sCondition, "=StatusEmployees.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(StatusEmployees.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " StatusEmployees,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "StatusEmployees, "
		If (InStr(1, sJoinCondition, "=StatusEmployees.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(StatusEmployees.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Employees.StatusID=StatusEmployees.StatusID) And "
	End If
	If ((InStr(1, sCondition, "=StatusJobs.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(StatusJobs.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " StatusJobs,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "StatusJobs, "
		If (InStr(1, sJoinCondition, "=StatusJobs.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(StatusJobs.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Jobs.StatusID=StatusJobs.StatusID) And "
	End If
	If ((InStr(1, sCondition, "=SystemLogs.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(SystemLogs.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " SystemLogs,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "SystemLogs, "
	End If
	If ((InStr(1, sCondition, "=Users.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(Users.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " Users,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "Users, "
	End If
	If ((InStr(1, sCondition, "=Zones.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(Zones.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " Zones,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "Zones, "
		If (InStr(1, sJoinCondition, "=Zones.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(Zones.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Areas.ZoneID=Zones.ZoneID) And "
	End If
	If ((InStr(1, sCondition, "=PaperworkTypes.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(PaperworkTypes.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " PaperworkTypes,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "PaperworkTypes, "
		If (InStr(1, sJoinCondition, "=PaperworkTypes.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(PaperworkTypes.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Paperworks.PaperworkTypeID=PaperworkTypes.PaperworkTypeID) And "
	End If
	If ((InStr(1, sCondition, "=StatusPaperworks.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(StatusPaperworks.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " StatusPaperworks,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "StatusPaperworks, "
		If (InStr(1, sJoinCondition, "=StatusPaperworks.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(StatusPaperworks.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Paperworks.StatusID=StatusPaperworks.StatusID) And "
	End If
	If ((InStr(1, sCondition, "=Owners.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(Owners.", vbBinaryCompare) > 0)) Then
		If (InStr(1, sTableNames, " Owners,", vbBinaryCompare) = 0) Then sTableNames = sTableNames & "Employees As Owners, "
		If (InStr(1, sJoinCondition, "=Owners.", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(Owners.", vbBinaryCompare) = 0) Then sJoinCondition = sJoinCondition & "(Paperworks.OwnerID1=Owners.EmployeeID) And "
	End If

	If Len(sFieldNames) > 0 Then sFieldNames = ", " & Left(sFieldNames, (Len(sFieldNames) - Len(", ")))
	If Len(sTableNames) > 0 Then sTableNames = Left(sTableNames, (Len(sTableNames) - Len(", ")))
	If Len(sJoinCondition) > 0 Then sJoinCondition = " And " & Left(sJoinCondition, (Len(sJoinCondition) - Len(" And ")))
	If Len(sSortFields) > 0 Then sSortFields = Left(sSortFields, (Len(sSortFields) - Len(", ")))

	GetDBFieldsNames = Err.number
	Err.Clear
End Function

Function GetFlagFieldName(sFlag, sFieldNames, sTableNames, sJoinCondition, sSortFields)
'************************************************************
'Purpose: To get the name of a flag given its id
'Inputs:  sFlag
'Outputs: sFieldNames, sTableNames, sJoinCondition, sSortFields
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetFlagFieldName"

	Select Case sFlag
		Case L_USER_FLAGS
			sFieldNames = sFieldNames & "UserLastName, UserName, "
			If InStr(1, sTableNames, " Users,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Users, "
			sJoinCondition = sJoinCondition & "(Users.UserID>=10) And "
			sSortFields = sSortFields & "UserLastName, UserName, "
		Case L_EMPLOYEE_NUMBER_FLAGS, L_EMPLOYEE_NUMBER1_FLAGS
			sFieldNames = sFieldNames & "EmployeeNumber, "
			If InStr(1, sTableNames, " Employees,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Employees, "
			sSortFields = sSortFields & "EmployeeNumber, "
		Case L_EMPLOYEE_NAME_FLAGS
			sFieldNames = sFieldNames & "EmployeeLastName, EmployeeLastName2, EmployeeName, "
			If InStr(1, sTableNames, " Employees,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Employees, "
			sSortFields = sSortFields & "EmployeeLastName, EmployeeLastName2, EmployeeName, "
		Case L_COMPANY_FLAGS
			sFieldNames = sFieldNames & "CompanyShortName, CompanyName, "
			If InStr(1, sTableNames, " Companies,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Companies, "
			sJoinCondition = sJoinCondition & "(Employees.CompanyID=Companies.CompanyID) And "
			sSortFields = sSortFields & "CompanyShortName, CompanyName, "
		Case L_EMPLOYEE_TYPE_FLAGS
			sFieldNames = sFieldNames & "EmployeeTypeShortName, EmployeeTypeName, "
			If InStr(1, sTableNames, " EmployeeTypes,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "EmployeeTypes, "
			sJoinCondition = sJoinCondition & "(Employees.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And "
			sSortFields = sSortFields & "EmployeeTypeShortName, EmployeeTypeName, "
            If StrComp(oRequest("Action").Item, "ModifyPayroll", vbBinaryCompare) = 0 Or StrComp(oRequest("Action").Item, "CalculatePayroll", vbBinaryCompare) = 0 Then
                If InStr(1, sTableNames, " Jobs,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Jobs, "
			    sJoinCondition = sJoinCondition & "(EmployeesHistoryList.JobID=Jobs.JobID) And (JobTypeID=3) "
            End If
		Case L_POSITION_TYPE_FLAGS
			sFieldNames = sFieldNames & "PositionTypeShortName, PositionTypeName, "
			If InStr(1, sTableNames, " PositionTypes,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "PositionTypes, "
			sJoinCondition = sJoinCondition & "(Employees.PositionTypeID=PositionTypes.PositionTypeID) And "
			sSortFields = sSortFields & "PositionTypeShortName, PositionTypeName, "
		Case L_CLASSIFICATION_FLAGS
			sFieldNames = sFieldNames & "Employees.ClassificationID, "
			If InStr(1, sTableNames, " Employees,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Employees, "
			sSortFields = sSortFields & "Employees.ClassificationID, "
		Case L_GROUP_GRADE_LEVEL_FLAGS
			sFieldNames = sFieldNames & "GroupGradeLevelName, "
			If InStr(1, sTableNames, " GroupGradeLevels,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "GroupGradeLevels, "
			sJoinCondition = sJoinCondition & "(Employees.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And "
			sSortFields = sSortFields & "GroupGradeLevelName, "
		Case L_INTEGRATION_FLAGS
			sFieldNames = sFieldNames & "Employees.IntegrationID, "
			If InStr(1, sTableNames, " Employees,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Employees, "
			sSortFields = sSortFields & "Employees.IntegrationID, "
		Case L_JOURNEY_FLAGS
			sFieldNames = sFieldNames & "JourneyShortName, JourneyName, "
			If InStr(1, sTableNames, " Journeys,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Journeys, "
			sJoinCondition = sJoinCondition & "(Employees.JourneyID=Journeys.JourneyID) And "
			sSortFields = sSortFields & "JourneyShortName, JourneyName, "
		Case L_SHIFT_FLAGS
			sFieldNames = sFieldNames & "ShiftShortName, ShiftName, "
			If InStr(1, sTableNames, " Shifts,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Shifts, "
			sJoinCondition = sJoinCondition & "(Employees.ShiftID=Shifts.ShiftID) And "
			sSortFields = sSortFields & "ShiftShortName, ShiftName, "
		Case L_LEVEL_FLAGS
			sFieldNames = sFieldNames & "Employees.LevelID, "
			If InStr(1, sTableNames, " Levels,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Levels, "
			sJoinCondition = sJoinCondition & "(Employees.LevelID=Levels.LevelID) And "
			sSortFields = sSortFields & "Employees.LevelID, "
		Case L_EMPLOYEE_STATUS_FLAGS
			sFieldNames = sFieldNames & "StatusEmployees.StatusName, "
			If InStr(1, sTableNames, " StatusEmployees,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "StatusEmployees, "
			sJoinCondition = sJoinCondition & "(Employees.StatusID=StatusEmployees.StatusID) And "
			sSortFields = sSortFields & "StatusEmployees.StatusName, "
		Case L_PAYMENT_CENTER_FLAGS
			sFieldNames = sFieldNames & "PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, "
			If InStr(1, sTableNames, " Areas As PaymentCenters,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Areas As PaymentCenters, "
			sJoinCondition = sJoinCondition & "(Employees.PaymentCenterID=PaymentCenters.AreaID) And "
			sSortFields = sSortFields & "PaymentCenters.AreaCode, PaymentCenters.AreaName, "
		Case L_EMPLOYEE_EMAIL_FLAGS
			sFieldNames = sFieldNames & "EmployeeEmail, "
			If InStr(1, sTableNames, " Employees,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Employees, "
			sSortFields = sSortFields & "EmployeeEmail, "
		Case L_SOCIAL_SECURITY_NUMBER_FLAGS
			sFieldNames = sFieldNames & "SocialSecurityNumber, "
			If InStr(1, sTableNames, " Employees,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Employees, "
			sSortFields = sSortFields & "SocialSecurityNumber, "
		Case L_EMPLOYEE_BIRTH_FLAGS
			sFieldNames = sFieldNames & "BirthDate, "
			If InStr(1, sTableNames, " Employees,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Employees, "
			sSortFields = sSortFields & "BirthDate, "
		Case L_EMPLOYEE_COUNTRY_FLAGS
			sFieldNames = sFieldNames & "CountryName, "
			If InStr(1, sTableNames, " Countries,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Countries, "
			sJoinCondition = sJoinCondition & "(Employees.CountryID=Countries.CountryID) And "
			sSortFields = sSortFields & "CountryName, "
		Case L_EMPLOYEE_RFC_FLAGS
			sFieldNames = sFieldNames & "RFC, "
			If InStr(1, sTableNames, " Employees,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Employees, "
			sSortFields = sSortFields & "RFC, "
		Case L_EMPLOYEE_CURP_FLAGS
			sFieldNames = sFieldNames & "CURP, "
			If InStr(1, sTableNames, " Employees,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Employees, "
			sSortFields = sSortFields & "CURP, "
		Case L_EMPLOYEE_GENDER_FLAGS
			sFieldNames = sFieldNames & "GenderName, "
			If InStr(1, sTableNames, " Genders,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Genders, "
			sJoinCondition = sJoinCondition & "(Employees.GenderID=Genders.GenderID) And "
			sSortFields = sSortFields & "GenderName, "
		Case L_EMPLOYEE_ACTIVE_FLAGS
			sFieldNames = sFieldNames & "Employees.Active, "
			If InStr(1, sTableNames, " Employees,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Employees, "
			sSortFields = sSortFields & "Employees.Active, "
		Case L_EMPLOYEE_START_DATE_FLAGS
			sFieldNames = sFieldNames & "Employees.StartDate, "
			If InStr(1, sTableNames, " Employees,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Employees, "
			sSortFields = sSortFields & "Employees.StartDate, "
		Case L_JOB_NUMBER_FLAGS
			sFieldNames = sFieldNames & "JobNumber, "
			If InStr(1, sTableNames, " Jobs,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Jobs, "
			sSortFields = sSortFields & "JobNumber, "
		Case L_ZONE_FLAGS, L_STATES_FLAGS
			sFieldNames = sFieldNames & "ZoneCode As ZoneShortName, ZoneName, "
			If InStr(1, sTableNames, " Zones,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Zones, "
			sJoinCondition = sJoinCondition & "(Areas.ZoneID=Zones.ZoneID) And "
			sSortFields = sSortFields & "ZoneCode, ZoneName, "
		Case L_ZONE_FOR_PAYMENT_CENTER_FLAGS
			sFieldNames = sFieldNames & "ZonesForPaymentCenter.ZoneCode As ZoneForPaymentShortName, ZonesForPaymentCenter.ZoneName As ZoneForPaymentName, "
			If InStr(1, sTableNames, " ZonesForPaymentCenter,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "ZonesForPaymentCenter, "
			sJoinCondition = sJoinCondition & "(PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And "
			sSortFields = sSortFields & "ZonesForPaymentCenter.ZoneCode, ZonesForPaymentCenter.ZoneName, "
		Case L_ZONE_FLAGS_FOR_EMPLOYEES
			sFieldNames = sFieldNames & "ZoneCode, ZoneName, "
			If InStr(1, sTableNames, " Zones,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Zones, "
			sJoinCondition = sJoinCondition & "(Employees.ZoneID=Zones.ZoneID) And "
			sSortFields = sSortFields & "ZoneCode, ZoneName, "
		Case L_AREA_FLAGS
			sFieldNames = sFieldNames & "Areas.AreaCode As AreaShortName, Areas.AreaName, "
			If InStr(1, sTableNames, " Areas,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Areas, "
			sJoinCondition = sJoinCondition & "(Jobs.AreaID=Areas.AreaID) And "
			sSortFields = sSortFields & "Areas.AreaCode, Areas.AreaName, "
		Case L_POSITION_FLAGS
			sFieldNames = sFieldNames & "PositionShortName, PositionName, "
			If InStr(1, sTableNames, " Positions,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Positions, "
			sJoinCondition = sJoinCondition & "(Jobs.PositionID=Positions.PositionID) And "
			sSortFields = sSortFields & "PositionShortName, PositionName, "
		Case L_JOB_TYPE_FLAGS
			sFieldNames = sFieldNames & "JobTypeName, "
			If InStr(1, sTableNames, " JobTypes,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "JobTypes, "
			sJoinCondition = sJoinCondition & "(Jobs.JobTypeID=JobTypes.JobTypeID) And "
			sSortFields = sSortFields & "JobTypeName, "
		Case L_OCCUPATION_TYPE_FLAGS
			sFieldNames = sFieldNames & "OccupationTypeShortName, OccupationTypeName, "
			If InStr(1, sTableNames, " OccupationTypes,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "OccupationTypes, "
			sJoinCondition = sJoinCondition & "(Jobs.OccupationTypeID=OccupationTypes.OccupationTypeID) And "
			sSortFields = sSortFields & "OccupationTypeShortName, OccupationTypeName, "
		Case L_JOB_START_DATE_FLAGS
			sFieldNames = sFieldNames & "Jobs.StartDate, "
			If InStr(1, sTableNames, " Jobs,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Jobs, "
			sSortFields = sSortFields & "Jobs.StartDate, "
		Case L_JOB_END_DATE_FLAGS
			sFieldNames = sFieldNames & "Jobs.EndDate, "
			If InStr(1, sTableNames, " Jobs,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Jobs, "
			sSortFields = sSortFields & "Jobs.EndDate, "
		Case L_JOB_STATUS_FLAGS
			sFieldNames = sFieldNames & "StatusJobs.StatusName, "
			If InStr(1, sTableNames, " StatusJobs,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "StatusJobs, "
			sJoinCondition = sJoinCondition & "(Jobs.JobTypeID=JobTypes.JobTypeID) And "
			sSortFields = sSortFields & "StatusJobs.StatusName, "
		Case L_JOB_ACTIVE_FLAGS
			sFieldNames = sFieldNames & "Jobs.Active, "
			If InStr(1, sTableNames, " Jobs,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Jobs, "
			sSortFields = sSortFields & "Jobs.Active, "
		Case L_AREA_CODE_FLAGS
			sFieldNames = sFieldNames & "AreaCode, "
			If InStr(1, sTableNames, " Areas,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Areas, "
			sSortFields = sSortFields & "AreaCode, "
		Case L_AREA_SHORT_NAME_FLAGS
			sFieldNames = sFieldNames & "AreaShortName, "
			If InStr(1, sTableNames, " Areas,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Areas, "
			sSortFields = sSortFields & "AreaShortName, "
		Case L_AREA_NAME_FLAGS
			sFieldNames = sFieldNames & "AreaName, "
			If InStr(1, sTableNames, " Areas,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Areas, "
			sSortFields = sSortFields & "AreaName, "
		Case L_AREA_TYPE_FLAGS
			sFieldNames = sFieldNames & "AreaTypeName, "
			If InStr(1, sTableNames, " AreaTypes,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "AreaTypes, "
			sJoinCondition = sJoinCondition & "(Areas.AreaTypeID=AreaTypes.AreaTypeID) And "
			sSortFields = sSortFields & "AreaTypeName, "
		Case L_CONFINE_TYPE_FLAGS
			sFieldNames = sFieldNames & "ConfineTypeName, "
			If InStr(1, sTableNames, " ConfineTypes,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "ConfineTypes, "
			sJoinCondition = sJoinCondition & "(Areas.ConfineTypeID=ConfineTypes.ConfineTypeID) And "
			sSortFields = sSortFields & "ConfineTypeName, "
		Case L_CENTER_TYPE_FLAGS
			sFieldNames = sFieldNames & "CenterTypeName, "
			If InStr(1, sTableNames, " CenterTypes,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "CenterTypes, "
			sJoinCondition = sJoinCondition & "(Areas.CenterTypeID=CenterTypes.CenterTypeID) And "
			sSortFields = sSortFields & "CenterTypeName, "
		Case L_CENTER_SUBTYPE_FLAGS
			sFieldNames = sFieldNames & "CenterSubtypeName, "
			If InStr(1, sTableNames, " CenterSubtypes,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "CenterSubtypes, "
			sJoinCondition = sJoinCondition & "(Areas.CenterSubtypeID=CenterSubtypes.CenterSubtypeID) And "
			sSortFields = sSortFields & "CenterSubtypeName, "
		Case L_ATTENTION_LEVEL_FLAGS
			sFieldNames = sFieldNames & "AttentionLevelShortName, AttentionLevelName, "
			If InStr(1, sTableNames, " AttentionLevels,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "AttentionLevels, "
			sJoinCondition = sJoinCondition & "(Areas.AttentionLevelID=AttentionLevels.AttentionLevelID) And "
			sSortFields = sSortFields & "AttentionLevelShortName, AttentionLevelName, "
		Case L_ECONOMIC_ZONE_FLAGS
			sFieldNames = sFieldNames & "EconomicZoneName, "
			If InStr(1, sTableNames, " EconomicZones,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "EconomicZones, "
			sJoinCondition = sJoinCondition & "(Areas.EconomicZoneID=EconomicZones.EconomicZoneID) And "
			sSortFields = sSortFields & "EconomicZoneName, "
		Case L_AREA_START_DATE_FLAGS
			sFieldNames = sFieldNames & "Areas.StartDate, "
			If InStr(1, sTableNames, " Areas,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Areas, "
			sSortFields = sSortFields & "Areas.StartDate, "
		Case L_AREA_END_DATE_FLAGS
			sFieldNames = sFieldNames & "Areas.EndDate, "
			If InStr(1, sTableNames, " Areas,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Areas, "
			sSortFields = sSortFields & "Areas.EndDate, "
		Case L_AREA_JOBS_FLAGS
			sFieldNames = sFieldNames & "Jobs, "
			If InStr(1, sTableNames, " Areas,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Areas, "
			sSortFields = sSortFields & "Jobs, "
		Case L_AREA_TOTAL_JOBS_FLAGS
			sFieldNames = sFieldNames & "TotalJobs, "
			If InStr(1, sTableNames, " Areas,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Areas, "
			sSortFields = sSortFields & "TotalJobs, "
		Case L_AREA_STATUS_FLAGS
			sFieldNames = sFieldNames & "StatusAreas.StatusName, "
			If InStr(1, sTableNames, " StatusAreas,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "StatusAreas, "
			sJoinCondition = sJoinCondition & "(Areas.StatusID=StatusAreas.StatusID) And "
			sSortFields = sSortFields & "StatusAreas.StatusName, "
		Case L_AREA_ACTIVE_FLAGS
			sFieldNames = sFieldNames & "Areas.Active, "
			If InStr(1, sTableNames, " Areas,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Areas, "
			sSortFields = sSortFields & "Areas.Active, "
		Case L_CONCEPT_ID_FLAGS, L_CONCEPT_1_FLAGS, L_THIRD_CONCEPTS_FLAGS, L_THIRD_CONCEPTS2_FLAGS, L_MEMORY_CONCEPT_ID_FLAGS
			sFieldNames = sFieldNames & "Concepts.ConceptShortName, Concepts.ConceptName, "
			sSortFields = sSortFields & "Concepts.ConceptShortName, Concepts.ConceptName, "
		Case L_TOTAL_PAYMENT_FLAGS
		Case L_BANK_FLAGS, L_ONE_BANK_FLAGS, L_ISSSTE_ONE_BANK_FLAGS
			sFieldNames = sFieldNames & "BankName, "
			If InStr(1, sTableNames, " Banks,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Banks, "
			sJoinCondition = sJoinCondition & "(BankAccounts.BankID=Banks.BankID) And "
			sSortFields = sSortFields & "BankName, "
		Case L_MEDICAL_AREAS_TYPES_FLAGS

		Case L_PAYROLL_FLAGS, L_OPEN_PAYROLL_FLAGS, L_CLOSED_PAYROLL_FLAGS
		Case L_MONTHS_FLAGS, L_DOUBLE_MONTHS_FLAGS
		Case L_YEARS_FLAGS
		Case L_DATE_FLAGS
			sFieldNames = sFieldNames & "XXXDate, "
			sSortFields = sSortFields & "XXXDate, "

		Case L_PAPERWORK_NUMBER_FLAGS
			sFieldNames = sFieldNames & "PaperworkNumber, "
			sSortFields = sSortFields & "PaperworkNumber, "
		Case L_PAPERWORK_FOLIO_NUMBER_FLAGS
			sFieldNames = sFieldNames & "PaperworkNumber, "
			sSortFields = sSortFields & "PaperworkNumber, "
		Case L_PAPERWORK_START_DATE_FLAGS
			sFieldNames = sFieldNames & "StartDate, "
			sSortFields = sSortFields & "StartDate, "
		Case L_PAPERWORK_ESTIMATED_DATE_FLAGS
			sFieldNames = sFieldNames & "EstimatedDate, "
			sSortFields = sSortFields & "EstimatedDate, "
		Case L_PAPERWORK_END_DATE_FLAGS
			sFieldNames = sFieldNames & "EndDate, "
			sSortFields = sSortFields & "EndDate, "
		Case L_PAPERWORK_DOCUMENT_NUMBER_FLAGS
			sFieldNames = sFieldNames & "DocumentNumber, "
			sSortFields = sSortFields & "DocumentNumber, "
		Case L_PAPERWORK_TYPE_FLAGS
			sFieldNames = sFieldNames & "PaperworkTypeName, "
			If InStr(1, sTableNames, " PaperworkTypes,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "PaperworkTypes, "
			sSortFields = sSortFields & "PaperworkTypeName, "
		Case L_PAPERWORK_OWNER_FLAGS
			sFieldNames = sFieldNames & "Owners.EmployeeLastName As OwnerLastName, Owners.EmployeeLastName As OwnerLastName2, Owners.EmployeeName As OwnerName, "
			If InStr(1, sTableNames, " StatusPaperworks,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "StatusPaperworks, "
			sSortFields = sSortFields & "Owners.EmployeeLastName, Owners.EmployeeLastName, Owners.EmployeeName, "
		Case L_PAPERWORK_STATUS_FLAGS
			sFieldNames = sFieldNames & "StatusPaperworks.StatusName, "
			If InStr(1, sTableNames, " StatusPaperworks,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "StatusPaperworks, "
			sSortFields = sSortFields & "StatusPaperworks.StatusName, "
        Case L_PAPERWORK_SUBJECT_TYPES
			sFieldNames = sFieldNames & "SubjectTypes.SubjectTypeName, "
			If InStr(1, sTableNames, " SubjectTypes,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "SubjectTypes, "
			sSortFields = sSortFields & "SubjectTypes.SubjectTypeName, "
		Case L_PAPERWORK_PRIORITY_FLAGS
			sFieldNames = sFieldNames & "Priorities.PriorityName, "
			If InStr(1, sTableNames, " Priorities,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Priorities, "
			sJoinCondition = sJoinCondition & "(Paperworks.PriorityID=Priorities.PriorityID) And "
			sSortFields = sSortFields & "Priorities.PriorityName, "
		Case L_PAPERWORK_OWNERS_FLAGS
			sFieldNames = sFieldNames & "PaperworkOwners.OwnerID As PaperworkOwnerID, PaperworkOwners.OwnerName As PaperworkOwnerName, PaperworkOwners.EmployeeID As PaperworkEmployeeID, "
			If InStr(1, sTableNames, " PaperworkOwnersLKP,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "PaperworkOwnersLKP, "
			If InStr(1, sTableNames, " PaperworkOwners,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "PaperworkOwners, "
			sJoinCondition = sJoinCondition & "(Paperworks.PaperworkID=PaperworkOwnersLKP.PaperworkID) And (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) "
			sSortFields = sSortFields & "PaperworkOwners.OwnerID, PaperworkOwners.OwnerName, PaperworkOwners.EmployeeID, "
		Case L_COURSE_NAME_FLAGS
			sFieldNames = sFieldNames & "SADE_Curso.Nombre_Curso, "
			If InStr(1, sTableNames, " SADE_Curso,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "SADE_Curso, "
			sSortFields = sSortFields & "SADE_Curso.Nombre_Curso, "
		Case L_COURSE_DIPLOMA_FLAGS
			sFieldNames = sFieldNames & "SADE_Perfiles.Nombre_Perfil, "
			If InStr(1, sTableNames, " SADE_Perfiles,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "SADE_Perfiles, "
			sJoinCondition = sJoinCondition & "(SADE_Curso.MostrarEvaluaciones=SADE_Perfiles.ID_Perfil) And "
			sSortFields = sSortFields & "SADE_Perfiles.Nombre_Perfil, "
		Case L_COURSE_LOCATION_FLAGS
			sFieldNames = sFieldNames & "SADE_Curso.Descripcion, "
			If InStr(1, sTableNames, " SADE_Curso,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "SADE_Curso, "
			sSortFields = sSortFields & "SADE_Curso.Descripcion, "
		Case L_COURSE_DURATION_FLAGS
			sFieldNames = sFieldNames & "SADE_Curso.TiempoEstimado, "
			If InStr(1, sTableNames, " SADE_Curso,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "SADE_Curso, "
			sSortFields = sSortFields & "SADE_Curso.TiempoEstimado, "
		Case L_COURSE_PARTICIPANTS_FLAGS
			sFieldNames = sFieldNames & "SADE_Curso.Participantes_Minimo, SADE_Curso.Participantes_Maximo, "
			If InStr(1, sTableNames, " SADE_Curso,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "SADE_Curso, "
			sSortFields = sSortFields & "SADE_Curso.Participantes_Minimo, SADE_Curso.Participantes_Maximo, "
		Case L_COURSE_DATES_FLAGS
			sFieldNames = sFieldNames & "SADE_CursosGruposLKP.Fecha_Inicio As CourseStartDate, SADE_CursosGruposLKP.Fecha_Final As CourseEndDate, "
			If InStr(1, sTableNames, " SADE_CursosGruposLKP,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "SADE_CursosGruposLKP, "
			sJoinCondition = sJoinCondition & "(SADE_Curso.ID_Curso=SADE_CursosGruposLKP.ID_Curso) And "
			sSortFields = sSortFields & "SADE_CursosGruposLKP.Fecha_Inicio, SADE_CursosGruposLKP.Fecha_Final, "
		Case L_COURSE_GRADE_FLAGS
			sFieldNames = sFieldNames & "SADE_CursosEmpleadosLKP.Calificacion, "
			If InStr(1, sTableNames, " SADE_CursosEmpleadosLKP,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "SADE_CursosEmpleadosLKP, "
			sJoinCondition = sJoinCondition & "(SADE_Curso.ID_Curso=SADE_CursosEmpleadosLKP.ID_Curso) And "
			sSortFields = sSortFields & "SADE_CursosEmpleadosLKP.Calificacion, "

		Case L_BUDGET_AREA_FLAGS
			sFieldNames = sFieldNames & "BudgetsMoney.AreaID, "
			sSortFields = sSortFields & "BudgetsMoney.AreaID, "
		Case L_BUDGET_PROGRAM_DUTY_FLAGS
			sFieldNames = sFieldNames & "BudgetsProgramDuties.ProgramDutyShortName, BudgetsProgramDuties.ProgramDutyName, "
			sTableNames = sTableNames & "BudgetsProgramDuties, "
			sJoinCondition = sJoinCondition & "(BudgetsMoney.ProgramDutyID=BudgetsProgramDuties.ProgramDutyID) And "
			sSortFields = sSortFields & "BudgetsProgramDuties.ProgramDutyShortName, BudgetsProgramDuties.ProgramDutyName, "
		Case L_BUDGET_FUND_FLAGS
			sFieldNames = sFieldNames & "BudgetsFunds.FundShortName, BudgetsFunds.FundName, "
			sTableNames = sTableNames & "BudgetsFunds, "
			sJoinCondition = sJoinCondition & "(BudgetsMoney.FundID=BudgetsFunds.FundID) And "
			sSortFields = sSortFields & "BudgetsFunds.FundShortName, BudgetsFunds.FundName, "
		Case L_BUDGET_DUTY_FLAGS
			sFieldNames = sFieldNames & "BudgetsDuties.DutyShortName, BudgetsDuties.DutyName, "
			sTableNames = sTableNames & "BudgetsDuties, "
			sJoinCondition = sJoinCondition & "(BudgetsMoney.DutyID=BudgetsDuties.DutyID) And "
			sSortFields = sSortFields & "BudgetsDuties.DutyShortName, BudgetsDuties.DutyName, "
		Case L_BUDGET_ACTIVE_DUTY_FLAGS
			sFieldNames = sFieldNames & "BudgetsActiveDuties.ActiveDutyShortName, BudgetsActiveDuties.ActiveDutyName, "
			sTableNames = sTableNames & "BudgetsActiveDuties, "
			sJoinCondition = sJoinCondition & "(BudgetsMoney.ActiveDutyID=BudgetsActiveDuties.ActiveDutyID) And "
			sSortFields = sSortFields & "BudgetsActiveDuties.ActiveDutyShortName, BudgetsActiveDuties.ActiveDutyName, "
		Case L_BUDGET_SPECIFIC_DUTY_FLAGS
			sFieldNames = sFieldNames & "BudgetsSpecificDuties.SpecificDutyShortName, BudgetsSpecificDuties.SpecificDutyName, "
			sTableNames = sTableNames & "BudgetsSpecificDuties, "
			sJoinCondition = sJoinCondition & "(BudgetsMoney.SpecificDutyID=BudgetsSpecificDuties.SpecificDutyID) And "
			sSortFields = sSortFields & "BudgetsSpecificDuties.SpecificDutyShortName, BudgetsSpecificDuties.SpecificDutyName, "
		Case L_BUDGET_PROGRAM_FLAGS
			sFieldNames = sFieldNames & "BudgetsPrograms.ProgramShortName, BudgetsPrograms.ProgramName, "
			sTableNames = sTableNames & "BudgetsPrograms, "
			sJoinCondition = sJoinCondition & "(BudgetsMoney.ProgramID=BudgetsPrograms.ProgramID) And "
			sSortFields = sSortFields & "BudgetsPrograms.ProgramShortName, BudgetsPrograms.ProgramName, "
		Case L_BUDGET_REGION_FLAGS
			sFieldNames = sFieldNames & "Zones1.ZoneCode As Zone1ShortName, Zones1.ZoneName As Zone1Name, "
			sTableNames = sTableNames & "Zones As Zones1, "
			sJoinCondition = sJoinCondition & "(BudgetsMoney.RegionID=Zones1.ZoneID) And "
			sSortFields = sSortFields & "Zones1.ZoneCode, Zones1.ZoneName, "
		Case L_BUDGET_UR_FLAGS
			sFieldNames = sFieldNames & "BudgetsMoney.BudgetUR, "
			sSortFields = sSortFields & "BudgetsMoney.BudgetUR, "
		Case L_BUDGET_CT_FLAGS
			sFieldNames = sFieldNames & "BudgetsMoney.BudgetCT, "
			sSortFields = sSortFields & "BudgetsMoney.BudgetCT, "
		Case L_BUDGET_AUX_FLAGS
			sFieldNames = sFieldNames & "BudgetsMoney.BudgetAUX, "
			sSortFields = sSortFields & "BudgetsMoney.BudgetAUX, "
		Case L_BUDGET_LOCATION_FLAGS
			sFieldNames = sFieldNames & "Zones2.ZoneCode As Zone2ShortName, Zones2.ZoneName As Zone2Name, "
			sTableNames = sTableNames & "Zones As Zones2, "
			sJoinCondition = sJoinCondition & "(BudgetsMoney.LocationID=Zones2.ZoneID) And "
			sSortFields = sSortFields & "Zones2.ZoneCode, Zones2.ZoneName, "
		Case L_BUDGET_BUDGET1_FLAGS
			sFieldNames = sFieldNames & "Budgets1.BudgetName As BudgetName1, "
			sTableNames = sTableNames & "Budgets As Budgets1, "
			sJoinCondition = sJoinCondition & "(BudgetsMoney.BudgetID1=Budgets1.BudgetID) And "
			sSortFields = sSortFields & "Budgets1.BudgetName, "
		Case L_BUDGET_BUDGET2_FLAGS
			sFieldNames = sFieldNames & "Budgets2.BudgetName As BudgetName2, "
			sTableNames = sTableNames & "Budgets As Budgets2, "
			sJoinCondition = sJoinCondition & "(BudgetsMoney.BudgetID2=Budgets2.BudgetID) And "
			sSortFields = sSortFields & "Budgets2.BudgetName, "
		Case L_BUDGET_BUDGET3_FLAGS
			sFieldNames = sFieldNames & "Budgets3.BudgetShortName As Budget3ShortName, Budgets3.BudgetName As Budget3Name, "
			sTableNames = sTableNames & "Budgets As Budgets3, "
			sJoinCondition = sJoinCondition & "(BudgetsMoney.BudgetID3=Budgets3.BudgetID) And "
			sSortFields = sSortFields & "Budgets3.BudgetShortName, Budgets3.BudgetName, "
		Case L_BUDGET_CONFINE_TYPE_FLAGS
			sFieldNames = sFieldNames & "BudgetsConfineTypes.ConfineTypeShortName, BudgetsConfineTypes.ConfineTypeName, "
			sTableNames = sTableNames & "BudgetsConfineTypes, "
			sJoinCondition = sJoinCondition & "(BudgetsMoney.ConfineTypeID=BudgetsConfineTypes.ConfineTypeID) And "
			sSortFields = sSortFields & "BudgetsConfineTypes.ConfineTypeShortName, BudgetsConfineTypes.ConfineTypeName, "
		Case L_BUDGET_ACTIVITY1_FLAGS
			sFieldNames = sFieldNames & "BudgetsActivities1.ActivityShortName As Activity1ShortName, BudgetsActivities1.ActivityName As Activity1Name, "
			sTableNames = sTableNames & "BudgetsActivities1, "
			sJoinCondition = sJoinCondition & "(BudgetsMoney.ActivityID1=BudgetsActivities1.ActivityID) And "
			sSortFields = sSortFields & "BudgetsActivities1.ActivityShortName, BudgetsActivities1.ActivityName, "
		Case L_BUDGET_ACTIVITY2_FLAGS
			sFieldNames = sFieldNames & "BudgetsActivities2.ActivityShortName As Activity2ShortName, BudgetsActivities2.ActivityName As Activity2Name, "
			sTableNames = sTableNames & "BudgetsActivities2, "
			sJoinCondition = sJoinCondition & "(BudgetsMoney.ActivityID2=BudgetsActivities2.ActivityID) And "
			sSortFields = sSortFields & "BudgetsActivities2.ActivityShortName, BudgetsActivities2.ActivityName, "
		Case L_BUDGET_PROCESS_FLAGS
			sFieldNames = sFieldNames & "BudgetsProcesses.ProcessShortName, BudgetsProcesses.ProcessName, "
			sTableNames = sTableNames & "BudgetsProcesses, "
			sJoinCondition = sJoinCondition & "(BudgetsMoney.ProcessID=BudgetsProcesses.ProcessID) And "
			sSortFields = sSortFields & "BudgetsProcesses.ProcessShortName, BudgetsProcesses.ProcessName, "
		Case L_BUDGET_YEAR_FLAGS
			sFieldNames = sFieldNames & "BudgetsMoney.BudgetYear, "
			sSortFields = sSortFields & "BudgetsMoney.BudgetYear, "
		Case L_BUDGET_MONTH_FLAGS
			sFieldNames = sFieldNames & "BudgetsMoney.BudgetMonth, "
			sSortFields = sSortFields & "BudgetsMoney.BudgetMonth, "
		Case L_BUDGET_ORIGINAL_POSITION_FLAGS
			sFieldNames = sFieldNames & "PositionShortName, PositionName, "
			If InStr(1, sTableNames, " BudgetsPositions,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "BudgetsPositions, "
			sSortFields = sSortFields & "PositionShortName, PositionName, "

		Case L_LOG_DATE_FLAGS
			sFieldNames = sFieldNames & "LogDate, "
			If InStr(1, sTableNames, " SystemLogs,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "SystemLogs, "
			sJoinCondition = sJoinCondition & "(Users.UserID=SystemLogs.UserID) And "
			sSortFields = sSortFields & "LogDate, "
		Case L_CREDITS_TYPES_ID_FLAGS
			sFieldNames = sFieldNames & "CreditTypeShortName, CreditTypeName, "
			If InStr(1, sTableNames, " CreditTypes,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "CreditTypes, "
			sJoinCondition = sJoinCondition & "(Credits.CreditTypes=CreditTypes.CreditTypeID) And "
			sSortFields = sSortFields & "CreditTypeShortName, CreditTypeName, "
		Case L_EMPLOYEE_BENEFICIARY_ID
			sFieldNames = sFieldNames & "BeneficiaryName, BeneficiaryLastName,BeneficiaryLastName2, "
			If InStr(1, sTableNames, " EmployeesBeneficiariesLKP,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "EmployeesBeneficiariesLKP, "
			sJoinCondition = sJoinCondition & "(Employees.EmployeeID=EmployeesBeneficiariesLKP.EmployeeID) And "
			sSortFields = sSortFields & "BeneficiaryID, "
		Case L_EMPLOYEE_CREDITOR_ID
			sFieldNames = sFieldNames & "CreditorName, CreditorLastName,CreditorLastName2, "
			If InStr(1, sTableNames, " EmployeesCreditorsLKP,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "EmployeesCreditorsLKP, "
			sJoinCondition = sJoinCondition & "(Employees.EmployeeID=EmployeesCreditorsLKP.EmployeeID) And "
			sSortFields = sSortFields & "CreditorID, "
		Case S_CREDITS_UPLOADED_FILE_NAME
			sFieldNames = sFieldNames & "UploadedFileName, "
			sSortFields = sSortFields & "UploadedFileName, "
		Case L_ABSENCE_ID_FLAGS
			sFieldNames = sFieldNames & "AbsenceTypeName, "
			If InStr(1, sTableNames, " AbsenceTypes,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "AbsenceTypes, "
			sJoinCondition = sJoinCondition & "(Absences.AbsenceTypeID=AbsenceTypes.AbsenceTypeID) And "
			sSortFields = sSortFields & "AbsenceTypeName, "
		Case L_ABSENCE_ACTIVE_FLAGS
			sFieldNames = sFieldNames & "EmployeesAbsencesLKP.Active, "
			If InStr(1, sTableNames, " EmployeesAbsencesLKP,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "EmployeesAbsencesLKP, "
			sSortFields = sSortFields & "EmployeesAbsencesLKP.Active, "
		Case L_CONCEPT_ACTIVE_FLAGS
			sFieldNames = sFieldNames & "EmployeesConceptsLKP.Active, "
			If InStr(1, sTableNames, " EmployeesConceptsLKP,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "EmployeesConceptsLKP, "
			sSortFields = sSortFields & "EmployeesConceptsLKP.Active, "
		Case L_BANK_ACCOUNTS_ACTIVE_FLAGS
			sFieldNames = sFieldNames & "BankAccounts.Active, "
			If InStr(1, sTableNames, " BankAccounts,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "BankAccounts, "
			sSortFields = sSortFields & "BankAccounts.Active, "
		Case L_CREDITS_ACTIVE_FLAGS
			sFieldNames = sFieldNames & "Credits.Active, "
			If InStr(1, sTableNames, " Credits,", vbBinaryCompare) = 0 Then sTableNames = sTableNames & "Credits, "
			sSortFields = sSortFields & "Credits.Active, "
	End Select

	GetFlagFieldName = Err.number
	Err.Clear
End Function

Function GetFlagName(sFlag)
'************************************************************
'Purpose: To get the name of a flag given its id
'Inputs:  sFlag
'Outputs: A string with the name of the flag
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetFlagName"

	Select Case sFlag
		Case L_USER_FLAGS
			GetFlagName = "Apellido del responsable,Nombre"
		Case L_EMPLOYEE_NUMBER_FLAGS, L_EMPLOYEE_NUMBER1_FLAGS
			GetFlagName = "Número de empleado"
		Case L_EMPLOYEE_NUMBER7_FLAGS
			GetFlagName = "Número de empleado temporal"
		Case L_EMPLOYEE_NAME_FLAGS
			GetFlagName = "Apellido paterno,Apellido materno,Nombre"
		Case L_COMPANY_FLAGS
			GetFlagName = "<SPAN COLS=""2"" />Empresa"
		Case L_EMPLOYEE_TYPE_FLAGS
			GetFlagName = "<SPAN COLS=""2"" />Tipo de tabulador"
		Case L_POSITION_TYPE_FLAGS
			GetFlagName = "<SPAN COLS=""2"" />Tipo de puesto"
		Case L_CLASSIFICATION_FLAGS
			GetFlagName = "Clasificación"
		Case L_GROUP_GRADE_LEVEL_FLAGS
			GetFlagName = "Grupo. grado. nivel"
		Case L_INTEGRATION_FLAGS
			GetFlagName = "Integración"
		Case L_JOURNEY_FLAGS
			GetFlagName = "<SPAN COLS=""2"" />Turno"
		Case L_SHIFT_FLAGS
			GetFlagName = "<SPAN COLS=""2"" />Horario"
		Case L_LEVEL_FLAGS
			GetFlagName = "Nivel"
		Case L_EMPLOYEE_STATUS_FLAGS
			GetFlagName = "Estatus del empleado"
		Case L_PAYMENT_CENTER_FLAGS
			GetFlagName = "<SPAN COLS=""2"" />Centro de pago"
		Case L_EMPLOYEE_EMAIL_FLAGS
			GetFlagName = "Correo electrónico"
		Case L_SOCIAL_SECURITY_NUMBER_FLAGS
			GetFlagName = "Número de seguro social"
		Case L_EMPLOYEE_BIRTH_FLAGS
			GetFlagName = "Fecha de nacimiento"
		Case L_EMPLOYEE_COUNTRY_FLAGS
			GetFlagName = "País"
		Case L_EMPLOYEE_RFC_FLAGS
			GetFlagName = "RFC"
		Case L_EMPLOYEE_CURP_FLAGS
			GetFlagName = "CURP"
		Case L_EMPLOYEE_GENDER_FLAGS
			GetFlagName = "Sexo"
		Case L_EMPLOYEE_ACTIVE_FLAGS
			GetFlagName = "¿Empleado activo?"
		Case L_EMPLOYEE_START_DATE_FLAGS
			GetFlagName = "Fecha de ingreso al Instituto"
		Case L_JOB_NUMBER_FLAGS
			GetFlagName = "Número de plaza"
		Case L_ZONE_FLAGS, L_STATES_FLAGS, L_ZONE_FLAGS_FOR_EMPLOYEES, L_ZONE_FOR_PAYMENT_CENTER_FLAGS
			GetFlagName = "<SPAN COLS=""2"" />Entidad federativa"
		Case L_AREA_FLAGS
			GetFlagName = "<SPAN COLS=""2"" />Área"
		Case L_POSITION_FLAGS
			GetFlagName = "<SPAN COLS=""2"" />Puesto"
		Case L_JOB_TYPE_FLAGS
			GetFlagName = "Tipo de plaza"
		Case L_OCCUPATION_TYPE_FLAGS
			GetFlagName = "<SPAN COLS=""2"" />Tipo de ocupación"
		Case L_JOB_START_DATE_FLAGS
			GetFlagName = "Inicio de la plaza"
		Case L_JOB_END_DATE_FLAGS
			GetFlagName = "Término de la plaza"
		Case L_JOB_STATUS_FLAGS
			GetFlagName = "Estatus de la plaza"
		Case L_JOB_ACTIVE_FLAGS
			GetFlagName = "¿Plaza activa?"
		Case L_AREA_CODE_FLAGS
			GetFlagName = "Código del centro de trabajo"
		Case L_AREA_SHORT_NAME_FLAGS
			GetFlagName = "Clave del centro de trabajo"
		Case L_AREA_NAME_FLAGS
			GetFlagName = "Nombre del centro de trabajo"
		Case L_AREA_TYPE_FLAGS
			GetFlagName = "Tipo de área"
		Case L_CONFINE_TYPE_FLAGS
			GetFlagName = "Ámbito para el área"
		Case L_CENTER_TYPE_FLAGS
			GetFlagName = "Tipo del centro de trabajo"
		Case L_CENTER_SUBTYPE_FLAGS
			GetFlagName = "Subtipo del centro de trabajo"
		Case L_ATTENTION_LEVEL_FLAGS
			GetFlagName = "<SPAN COLS=""2"" />Nivel de atención"
		Case L_ECONOMIC_ZONE_FLAGS
			GetFlagName = "Zona económica"
		Case L_AREA_START_DATE_FLAGS
			GetFlagName = "Fecha de inicio del centro de trabajo"
		Case L_AREA_END_DATE_FLAGS
			GetFlagName = "Fecha de término del centro de trabajo"
		Case L_AREA_JOBS_FLAGS
			GetFlagName = "Plazas"
		Case L_AREA_TOTAL_JOBS_FLAGS
			GetFlagName = "Total de plazas"
		Case L_AREA_STATUS_FLAGS
			GetFlagName = "Estatus del centro de trabajo"
		Case L_CONCEPTS_VALUES_STATUS_FLAGS
			GetFlagName = "Estatus del tabulador"
		Case L_EMPLOYEE_REASON_ID_FLAGS
			GetFlagName = "Tipo de movimiento"
		Case L_AREA_ACTIVE_FLAGS
			GetFlagName = "¿Centro de trabajo activo?"
		Case L_CONCEPT_ID_FLAGS, L_CONCEPT_1_FLAGS, L_THIRD_CONCEPTS_FLAGS, L_THIRD_CONCEPTS2_FLAGS, L_MEMORY_CONCEPT_ID_FLAGS
			GetFlagName = "Concepto de pago"
		Case L_TOTAL_PAYMENT_FLAGS
			GetFlagName = "Líquido"
		Case L_BANK_FLAGS, L_ONE_BANK_FLAGS, L_ISSSTE_ONE_BANK_FLAGS
			GetFlagName = "Bancos"
		Case L_MEDICAL_AREAS_TYPES_FLAGS
			GetFlagName = "Reporte UNIMED"
		Case L_DOCUMENT_FOR_LICENSE_NUMBER_FLAGS
			GetFlagName = "No. de oficio"
		Case L_DOCUMENT_REQUEST_NUMBER_FLAGS
			GetFlagName = "No. de solicitud"
		Case L_DOCUMENT_FOR_CANCEL_LICENSE_NUMBER_FLAGS
			GetFlagName = "No. de oficio de cancelación"

		Case L_PAYROLL_FLAGS, L_OPEN_PAYROLL_FLAGS, L_CLOSED_PAYROLL_FLAGS
			GetFlagName = "Nómina"
		Case L_MONTHS_FLAGS, L_DOUBLE_MONTHS_FLAGS
			GetFlagName = "Mes"
		Case L_YEARS_FLAGS
			GetFlagName = "Año"
		Case L_DATE_FLAGS
			GetFlagName = "Periodo"

		Case L_PAPERWORK_NUMBER_FLAGS
			GetFlagName = "No. del trámite"
		Case L_PAPERWORK_FOLIO_NUMBER_FLAGS
			GetFlagName = "No. del folio"
		Case L_PAPERWORK_START_DATE_FLAGS
			GetFlagName = "Fecha de recepción"
		Case L_PAPERWORK_ESTIMATED_DATE_FLAGS
			GetFlagName = "Fecha límite de respuesta"
		Case L_PAPERWORK_END_DATE_FLAGS
			GetFlagName = "Fecha de atención"
		Case L_PAPERWORK_DOCUMENT_NUMBER_FLAGS
			GetFlagName = "No. de documento"
		Case L_PAPERWORK_TYPE_FLAGS
			GetFlagName = "Tipo de trámite"
		Case L_PAPERWORK_OWNER_FLAGS
			GetFlagName = "Apellido paterno del responsable,Apellido materno del responsable,Nombre del responsable"
		Case L_PAPERWORK_STATUS_FLAGS
			GetFlagName = "Estatus del trámite"
        Case L_PAPERWORK_SUBJECT_TYPES
            GetFlagName = "Tipo de asunto"
		Case L_PAPERWORK_PRIORITY_FLAGS
			GetFlagName = "Prioridad"
		Case L_PAPERWORK_OWNERS_FLAGS
			GetFlagName = "Clave responsable,Responsable,No. empleado"

		Case L_COURSE_NAME_FLAGS
			GetFlagName = "Curso"
		Case L_COURSE_DIPLOMA_FLAGS
			GetFlagName = "Diplomado"
		Case L_COURSE_LOCATION_FLAGS
			GetFlagName = "Ubicación"
		Case L_COURSE_DURATION_FLAGS
			GetFlagName = "Duración"
		Case L_COURSE_PARTICIPANTS_FLAGS
			GetFlagName = "No. mínimo,No. máximo"
		Case L_COURSE_DATES_FLAGS
			GetFlagName = "Fecha de inicio,Fecha final"
		Case L_COURSE_GRADE_FLAGS
			GetFlagName = "Calificación"

		Case L_BUDGET_AREA_FLAGS
			GetFlagName = "Área"
		Case L_BUDGET_COMPANIES_FLAGS
			GetFlagName = "Compañía"
		Case L_BUDGET_PROGRAM_DUTY_FLAGS
			GetFlagName = "Programa presupuestario"
		Case L_BUDGET_FUND_FLAGS
			GetFlagName = "Fondo"
		Case L_BUDGET_DUTY_FLAGS
			GetFlagName = "Función"
		Case L_BUDGET_ACTIVE_DUTY_FLAGS
			GetFlagName = "Subfunción activa"
		Case L_BUDGET_SPECIFIC_DUTY_FLAGS
			GetFlagName = "Subfunción específica"
		Case L_BUDGET_PROGRAM_FLAGS
			GetFlagName = "Programa"
		Case L_BUDGET_REGION_FLAGS
			GetFlagName = "<SPAN COLS=""2"" />Región"
		Case L_BUDGET_UR_FLAGS
			GetFlagName = "UR"
		Case L_BUDGET_CT_FLAGS
			GetFlagName = "CT"
		Case L_BUDGET_AUX_FLAGS
			GetFlagName = "AUX"
		Case L_BUDGET_LOCATION_FLAGS
			GetFlagName = "<SPAN COLS=""2"" />Municipio"
		Case L_BUDGET_BUDGET1_FLAGS
			GetFlagName = "Partida"
		Case L_BUDGET_BUDGET2_FLAGS
			GetFlagName = "Subpartida"
		Case L_BUDGET_BUDGET3_FLAGS
			GetFlagName = "<SPAN COLS=""2"" />Tipo de pago"
		Case L_BUDGET_CONFINE_TYPE_FLAGS
			GetFlagName = "Ámbito"
		Case L_BUDGET_ACTIVITY1_FLAGS
			GetFlagName = "Actividad institucional"
		Case L_BUDGET_ACTIVITY2_FLAGS
			GetFlagName = "Actividad presupuestaria"
		Case L_BUDGET_PROCESS_FLAGS
			GetFlagName = "Proceso"
		Case L_BUDGET_YEAR_FLAGS
			GetFlagName = "Año"
		Case L_BUDGET_MONTH_FLAGS
			GetFlagName = "Mes"
		Case L_BUDGET_ORIGINAL_POSITION_FLAGS
			GetFlagName = "<SPAN COLS=""2"" />Puesto"
		Case L_CREDITS_TYPES_ID_FLAGS
			GetFlagName = "Tipo de crédito"
		Case L_EMPLOYEE_BENEFICIARY_ID
			GetFlagName = "Beneficiaria de pensión alimenticia"
		Case L_EMPLOYEE_CREDITOR_ID
			GetFlagName = "Acreedor"
		Case L_LOG_DATE_FLAGS
			GetFlagName = "Fecha de entrada al sistema"
		Case L_ABSENCE_ACTIVE_FLAGS
			GetFlagName = "¿Incidencia activa?"
		Case L_ABSENCE_APPLIED_DATE_FLAGS
			GetFlagName = "Quincena de aplicación de la incidencia"
		Case L_CONCEPT_APPLIED_DATE_FLAGS
			GetFlagName = "Quincena de aplicación de la prestación"
		Case L_CONCEPT_ACTIVE_FLAGS
			GetFlagName = "¿Concepto activo?"
		Case L_BANK_ACCOUNTS_ACTIVE_FLAGS
			GetFlagName = "¿Cuenta bancaria activa?"
		Case L_CREDITS_ACTIVE_FLAGS
			GetFlagName = "¿Crédito activo?"
		Case Else
			GetFlagName = ""
	End Select

	Err.Clear
End Function

Function GetReportNameByConstant(lReportNumber)
'************************************************************
'Purpose: To get the name of the report given its constant
'Inputs:  lReportNumber
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetReportNameByConstant"

	Select Case CLng(lReportNumber)
		Case LOGS_HISTORY_REPORTS
			GetReportNameByConstant = "Historial de entradas al sistema"
		Case AREAS_COUNT_REPORTS
			GetReportNameByConstant = "Conteo de centros de trabajo"
		Case EMPLOYEES_COUNT_REPORTS
			GetReportNameByConstant = "Conteo de empleados"
		Case JOBS_COUNT_REPORTS
			GetReportNameByConstant = "Conteo de plazas"
		Case AREAS_LIST_REPORTS
			GetReportNameByConstant = "Información de los centros de trabajo"
		Case EMPLOYEES_LIST_REPORTS
			GetReportNameByConstant = "Plantilla de personal"
		Case JOBS_LIST_BY_MODIFY_DATE
			GetReportNameByConstant = "Plazas Creadas o Modificadas"
		Case JOBS_LIST_REPORTS, SPECIAL_JOBS_LIST_REPORTS
			If StrComp(oRequest("JobStatusID").Item, "1", vbBinaryCompare) = 0 Then
				GetReportNameByConstant = "Plazas ocupadas"
			ElseIf StrComp(oRequest("JobStatusID").Item, "2", vbBinaryCompare) = 0 Then
				GetReportNameByConstant = "Plazas vacantes"
			ElseIf StrComp(oRequest("JobStatusID").Item, "3", vbBinaryCompare) = 0 Then
				GetReportNameByConstant = "Plazas congeladas"
			ElseIf StrComp(oRequest("JobStatusID").Item, "4", vbBinaryCompare) = 0 Then
				GetReportNameByConstant = "Plazas con licencia"
			ElseIf Len(oRequest("MovedEmployees").Item) > 0 Then
				GetReportNameByConstant = "Plazas en interinatos"
			ElseIf Len(oRequest("JobsOwners").Item) > 0 Then
				GetReportNameByConstant = "Plazas reservadas"
			Else
				GetReportNameByConstant = "Información de las plazas"
			End If
		Case EMPLOYEE_HISTORY_LIST_REPORTS
			GetReportNameByConstant = "Historial del empleado"
		Case EMPLOYEE_FORM_HISTORY_LIST_REPORTS
			GetReportNameByConstant = "Historial de cambios del formato FM1"
		Case EMPLOYEE_PAYMENTS_HISTORY_LIST_REPORTS
			GetReportNameByConstant = "Historial de pagos"
		Case EMPLOYEE_PAYROLL_REPORTS
			GetReportNameByConstant = "Pagos por quincena"

		Case ISSSTE_1001_REPORTS
			GetReportNameByConstant = "Hoja de cifras"
		Case ISSSTE_1002_REPORTS
			GetReportNameByConstant = "Revisión de diferencias"
		Case ISSSTE_1003_REPORTS, ISSSTE_1470_REPORTS
			GetReportNameByConstant = "Listado de firmas"
		Case ISSSTE_1004_REPORTS
			GetReportNameByConstant = "Resumen por conceptos de nómina"
		Case ISSSTE_1005_REPORTS
			GetReportNameByConstant = "Hoja informativa de servicios de informática"
		Case ISSSTE_1006_REPORTS
			GetReportNameByConstant = "Concentrado de conceptos de la nómina ordinaria"
		Case ISSSTE_1007_REPORTS
			GetReportNameByConstant = "Remesa para cubrir la nómina"
		Case ISSSTE_1008_REPORTS
			GetReportNameByConstant = "Empleados por tipo de tabulador y por empresa"
		Case ISSSTE_1009_REPORTS
			GetReportNameByConstant = "Resumen mensual de nóminas"
		Case ISSSTE_1010_REPORTS
			GetReportNameByConstant = "Ramas médica, paramédica, de grupos afines y operativa, de enlace y de alto nivel de responsabilidad"
		Case ISSSTE_1011_REPORTS
			GetReportNameByConstant = "Pensión alimenticia de ramas médica, paramédica, de grupos afines y operativa, de enlace y de alto nivel de responsabilidad"
		Case ISSSTE_1012_REPORTS
			GetReportNameByConstant = "Diferencias totales por concepto"
		Case ISSSTE_1013_REPORTS
			GetReportNameByConstant = "Diferencias de empleados por unidad administrativa"
		Case ISSSTE_1014_REPORTS
			GetReportNameByConstant = "Altas por unidad administrativa"
		Case ISSSTE_1015_REPORTS
			GetReportNameByConstant = "Altas de empleados por unidad administrativa"
		Case ISSSTE_1016_REPORTS
			GetReportNameByConstant = "Bajas por unidad administrativa"
		Case ISSSTE_1017_REPORTS
			GetReportNameByConstant = "Bajas de empleados por unidad administrataiva"
		Case ISSSTE_1018_REPORTS
			GetReportNameByConstant = "Diferencias de sueldo por unidad administrativa"
		Case ISSSTE_1019_REPORTS
			GetReportNameByConstant = "Cambios de puesto por unidad administrativa"
		Case ISSSTE_1020_REPORTS
			GetReportNameByConstant = "Funcionarios con líquido mayor a la suma de sueldo base más compensación"
		Case ISSSTE_1021_REPORTS
			GetReportNameByConstant = "Totales por nómina"
		Case ISSSTE_1022_REPORTS
			GetReportNameByConstant = "Empleados con líquidos mayores"
		Case ISSSTE_1023_REPORTS
			GetReportNameByConstant = "Reporte de movimientos"
		Case ISSSTE_1024_REPORTS
			GetReportNameByConstant = "Funcionarios y operativos por concepto de pago y empresa"
		Case ISSSTE_1025_REPORTS
			GetReportNameByConstant = "Recibos de pago por empleado"
		Case ISSSTE_1026_REPORTS
			GatReportNameByConstant = "Reporte de incidencias del personal"
		Case ISSSTE_1027_REPORTS
			GatReportNameByConstant = "Listado de cheques"
		Case ISSSTE_1028_REPORTS
			GatReportNameByConstant = "Resumen de Nóminas al SAR"
		Case ISSSTE_1029_REPORTS
			GatReportNameByConstant = "Reporte de concentrado de incidencias"
		Case ISSSTE_1030_REPORTS
			GatReportNameByConstant = "Incidencias con horas extras y primas dominicales"
		Case ISSSTE_1031_REPORTS
			GatReportNameByConstant = "Reporte de altas y bajas del bimestre"
		Case ISSSTE_1032_REPORTS
			GatReportNameByConstant = "Reporte de dispersión por unidad administrativa"
		Case ISSSTE_1033_REPORTS
			GatReportNameByConstant = "Reporte de aportaciones"
		Case ISSSTE_1034_REPORTS
			GatReportNameByConstant = "Control y distribución de comprobantes de abono en cuenta de trabajadores"
		Case ISSSTE_1100_REPORTS
			GetReportNameByConstant = "Catálogo de trabajadores por centro de pago"
		Case ISSSTE_1101_REPORTS
			GetReportNameByConstant = "Aguinaldos"
		Case ISSSTE_1102_REPORTS
			GetReportNameByConstant = "Reporte de reclamos de pago por ajustes y deducciones"
		Case ISSSTE_1103_REPORTS
			GetReportNameByConstant = "Reporte de movimientos en trámite"
		Case ISSSTE_1104_REPORTS
			GetReportNameByConstant = "Reporte de movimientos por usuario"
		Case ISSSTE_1105_REPORTS
			GetReportNameByConstant = "Reporte de registro de movimientos"
		Case ISSSTE_1106_REPORTS
			GetReportNameByConstant = "Reporte de honorarios"
		Case ISSSTE_1107_REPORTS
			GetReportNameByConstant = "Nómina de personal correspondiente a conceptos"
		Case ISSSTE_1108_REPORTS
			GetReportNameByConstant = "Reporte de personal con conceptos"
		Case ISSSTE_1109_REPORTS
			GetReportNameByConstant = "Impresión del formato FM1"
		Case ISSSTE_1110_REPORTS
			GetReportNameByConstant = "Impresión de formato de honorarios"
		Case ISSSTE_1111_REPORTS
			GetReportNameByConstant = "Histórico de plazas"
		Case ISSSTE_1112_REPORTS
			GetReportNameByConstant = "Hoja única de servicio"
		Case ISSSTE_1113_REPORTS
			GetReportNameByConstant = "Reporte de validación del pago de aguinaldo"
		Case ISSSTE_1114_REPORTS, ISSSTE_1472_REPORTS
			GetReportNameByConstant = "Pagos cancelados"
		Case ISSSTE_1115_REPORTS
			GetReportNameByConstant = "Impresión de formato de baja honorarios"
		Case ISSSTE_1116_REPORTS, ISSSTE_1204_REPORTS, ISSSTE_1702_REPORTS
			GetReportNameByConstant = "Antigüedad para un empleado"
		Case ISSSTE_1117_REPORTS, ISSSTE_1205_REPORTS, ISSSTE_1703_REPORTS
			GetReportNameByConstant = "Reporte de Antigüedades"
		Case ISSSTE_1118_REPORTS, ISSSTE_1206_REPORTS, ISSSTE_1704_REPORTS
			GetReportNameByConstant = "Validación de nómina 1o de Octubre"
		Case ISSSTE_1119_REPORTS
			GetReportNameByConstant = "Reporte de empleados con derecho al concepto 41"
		Case ISSSTE_1151_REPORTS
			GetReportNameByConstant = "Acumulados anuales"
		Case ISSSTE_1152_REPORTS
			GetReportNameByConstant = "Constancia de percepciones y deducciones anuales"
		Case ISSSTE_1153_REPORTS
			GetReportNameByConstant = "Ajuste anual del impuesto sobre la renta"
		Case ISSSTE_1154_REPORTS
			GetReportNameByConstant = "Recálculo anual de impuestos"
		Case ISSSTE_1155_REPORTS
			GetReportNameByConstant = "Aplicación del ajuste anual del impuesto sobre la renta"
		Case ISSSTE_1157_REPORTS
			GetReportNameByConstant = "Declaración informativa múltiple (DIM)"
		Case ISSSTE_1200_REPORTS
			GetReportNameByConstant = "Funcionarios y operativos por concepto de pago y empresa"
		Case ISSSTE_1201_REPORTS
			GetReportNameByConstant = "Reporte de personal con conceptos"
		Case ISSSTE_1202_REPORTS
			GetReportNameByConstant = "Reporte de personal con créditos"
		Case ISSSTE_1203_REPORTS
			GetReportNameByConstant = "Hoja única de servicio"
		Case ISSSTE_1207_REPORTS
			GetReportNameByConstant = "Constancia de servicio activo"
		Case ISSSTE_1208_REPORTS
			GetReportNameByConstant = "Constancia de descuento"
		Case ISSSTE_1209_REPORTS
			GetReportNameByConstant = "Reporte de revisión de nóminas"
		Case ISSSTE_1210_REPORTS
			GetReportNameByConstant = "Reporte de horas extras y primas dominicales"
		Case ISSSTE_1211_REPORTS
			GetReportNameByConstant = "Calificación de médicos y enfermeras"
		Case ISSSTE_1221_REPORTS
			GetReportNameByConstant = "Reporte de registros cargados desde archivo de terceros"
		Case ISSSTE_1222_REPORTS
			GetReportNameByConstant = "Reporte de registros rechazados desde archivos de terceros"
		Case ISSSTE_1223_REPORTS
			GetReportNameByConstant = "Reporte de beneficiarios de pensiones alimenticias por empleado"
		Case ISSSTE_1224_REPORTS
			GetReportNameByConstant = "Reporte de empleados con pensiones alimenticias"
		Case ISSSTE_1225_REPORTS
			GetReportNameByConstant = "Generación de archivo Repcsi"
		Case ISSSTE_1311_REPORTS
			GetReportNameByConstant = "Plantilla de nómina"
		Case ISSSTE_1334_REPORTS
			GetReportNameByConstant = "Reporte UNIMED"
		Case ISSSTE_1335_REPORTS
			GetReportNameByConstant = "Catálogo de puestos y tabuladores de puestos"
		Case ISSSTE_1336_REPORTS
			GetReportNameByConstant = "Catálogo de centros de trabajo"
		Case ISSSTE_1337_REPORTS
			GetReportNameByConstant = "Catálogo de centros de pago"
		Case ISSSTE_1338_REPORTS
			GetReportNameByConstant = "Revisión de nóminas"
		Case ISSSTE_1339_REPORTS
			GetReportNameByConstant = "Archivo SICAD"
		Case ISSSTE_1340_REPORTS
			GetReportNameByConstant = "Archivo SICAD de cancelaciones"
		Case ISSSTE_1354_REPORTS
			GetReportNameByConstant = "Información registrada en la bolsa de trabajo"
		Case ISSSTE_1356_REPORTS
			GetReportNameByConstant = "Búsqueda de información de escalafón"
		Case ISSSTE_1364_REPORTS
			GetReportNameByConstant = "Reporte de desarrollo humano"
		Case ISSSTE_1365_REPORTS
			GetReportNameByConstant = "Seguimiento de cursos de capacitación"
		Case ISSSTE_1367_REPORTS
			GetReportNameByConstant = "Reporte de curriculum por empleado"
		Case ISSSTE_1369_REPORTS
			GetReportNameByConstant = "Registro de detección de necesidades"
		Case ISSSTE_1371_REPORTS
			GetReportNameByConstant = "Generación de archivo para registro de servidores públicos"
		Case ISSSTE_1372_REPORTS
			GetReportNameByConstant = "RUSP. Información básica"
		Case ISSSTE_1373_REPORTS
			GetReportNameByConstant = "RUSP. Bajas"
		Case ISSSTE_1374_REPORTS
			GetReportNameByConstant = "RUSP. Datos personales"
		Case ISSSTE_1400_REPORTS
			GetReportNameByConstant = "CLCs"
		Case ISSSTE_1401_REPORTS
			GetReportNameByConstant = "Archivo para carga del SPEP"
		Case ISSSTE_1402_REPORTS
			GetReportNameByConstant = "Archivo para carga del SPEP por centro de trabajo"
		Case ISSSTE_1403_REPORTS
			GetReportNameByConstant = "Resumen mensual de nóminas"
		Case ISSSTE_1404_REPORTS
			GetReportNameByConstant = "Fajillas por estado"
		Case ISSSTE_1411_REPORTS
			GetReportNameByConstant = "Cifras iniciales"
		Case ISSSTE_1412_REPORTS
			GetReportNameByConstant = "Archivo para carga del SPEP del FONAC"
		Case ISSSTE_1413_REPORTS
			GetReportNameByConstant = "Archivo para contabilidad"
		Case ISSSTE_1414_REPORTS
			GetReportNameByConstant = "Detalle de las aportaciones por quincena"
		Case ISSSTE_1415_REPORTS
			GetReportNameByConstant = "Reporte para el Fiduciario"
		Case ISSSTE_1416_REPORTS
			GetReportNameByConstant = "Respaldo de los empleados cotizantes por quincena"
		Case ISSSTE_1417_REPORTS
			GetReportNameByConstant = "Cifras para el pago de nómina"
		Case ISSSTE_1420_REPORTS, ISSSTE_2420_REPORTS
			GetReportNameByConstant = "Reporte de personal interno"
		Case ISSSTE_1421_REPORTS, ISSSTE_2421_REPORTS
			GetReportNameByConstant = "Reporte de los conceptos 15 (guardias), 31 (suplencias) y C5 (guardias PROVAC) de personal interno"
		Case ISSSTE_1422_REPORTS, ISSSTE_2422_REPORTS
			GetReportNameByConstant = "Reporte de captura de personal externo"
		Case ISSSTE_1423_REPORTS, ISSSTE_2423_REPORTS
			GetReportNameByConstant = "Reporte de validación de personal externo"
		Case ISSSTE_1424_REPORTS
			If Len(oRequest("External").Item) = 0 Then
				GetReportNameByConstant = "Consolidado de personal interno"
			Else
				GetReportNameByConstant = "Consolidado de los prestadores de servicio"
			End If
		Case ISSSTE_1425_REPORTS
			If Len(oRequest("External").Item) = 0 Then
				GetReportNameByConstant = "Validación de captura de personal interno"
			Else
				GetReportNameByConstant = "Validación de captura de prestadores de servicios"
			End If
		Case ISSSTE_1426_REPORTS, ISSSTE_2426_REPORTS
			GetReportNameByConstant = "Volantes"
		Case ISSSTE_1427_REPORTS, ISSSTE_2427_REPORTS
			GetReportNameByConstant = "Listado de firmas"
		Case ISSSTE_1428_REPORTS, ISSSTE_2428_REPORTS
			GetReportNameByConstant = "Reporte de totales"
		Case ISSSTE_1429_REPORTS, ISSSTE_2429_REPORTS
			GetReportNameByConstant = "Reporte concentrado por quincena"
		Case ISSSTE_1430_REPORTS, ISSSTE_2430_REPORTS
			GetReportNameByConstant = "Reporte estadístico de causas"
		Case ISSSTE_1431_REPORTS, ISSSTE_2431_REPORTS
			GetReportNameByConstant = "Reporte de cuentas bancarias"
		Case ISSSTE_1432_REPORTS, ISSSTE_2432_REPORTS
			GetReportNameByConstant = "Listado de actualización de cuentas bancarias"
		Case ISSSTE_1433_REPORTS
			GetReportNameByConstant = "Reporte de estímulos"
		Case ISSSTE_1434_REPORTS
			GetReportNameByConstant = "Reporte de incidencias"
		Case ISSSTE_1435_REPORTS
			GetReportNameByConstant = "Reporte de tabuladores de pago"
		Case ISSSTE_1471_REPORTS
			GetReportNameByConstant = "Recibo de distribución y recepción de cheques"
		Case ISSSTE_1473_REPORTS
			GetReportNameByConstant = "Bloqueos aplicados"
		Case ISSSTE_1474_REPORTS
			GetReportNameByConstant = "Archivo de depósitos bancarios"
		Case ISSSTE_1475_REPORTS
			GetReportNameByConstant = "Archivo de liberación de cheques"
		Case ISSSTE_1476_REPORTS
			GetReportNameByConstant = "Recibo por pago de honorarios"
		Case ISSSTE_1477_REPORTS
			GetReportNameByConstant = "Impuesto sobre la renta"
		Case ISSSTE_1478_REPORTS
			GetReportNameByConstant = "Cálculo del impuesto sobre nóminas"
		Case ISSSTE_1490_REPORTS
			GetReportNameByConstant = "Reporte de cifras"
		Case ISSSTE_1491_REPORTS
			GetReportNameByConstant = "Salida de archivos de terceros"
		Case ISSSTE_1492_REPORTS
			GetReportNameByConstant = "Reportes de terceros"
		Case ISSSTE_1493_REPORTS
			GetReportNameByConstant = "Archivo para carga del SPEP por concepto"
		Case ISSSTE_1494_REPORTS
			GetReportNameByConstant = "Memoria de cálculo para el entero de cuotas sindicales"
		Case ISSSTE_1499_REPORTS
			GetReportNameByConstant = "Constancia de percepciones y deducciones"
		Case ISSSTE_1503_REPORTS
			GetReportNameByConstant = "Costeo de plazas"
		Case ISSSTE_1504_REPORTS
			GetReportNameByConstant = "Consulta de presupuesto"
		Case ISSSTE_1561_REPORTS
			GetReportNameByConstant = "Proyecto de presupuesto"
		Case ISSSTE_1562_REPORTS
			GetReportNameByConstant = "Archivo de carga para el SPEP"
		Case ISSSTE_1563_REPORTS
			GetReportNameByConstant = "Formato único de movimientos presupuestales"
		Case ISSSTE_1571_REPORTS
			GetReportNameByConstant = "Registro de un costeo como original"
		Case ISSSTE_1581_REPORTS
			GetReportNameByConstant = "Personal ocupado por rama de actividad"
		Case ISSSTE_1582_REPORTS
			GetReportNameByConstant = "Personal ocupado y pago de sueldos y salarios en la administración pública federal"
		Case ISSSTE_1583_REPORTS
			GetReportNameByConstant = "Prestaciones a favor de los servidores públicos del ISSSTE"
		Case ISSSTE_1584_REPORTS
			GetReportNameByConstant = "Trabajadores cotizantes al régimen del ISSSTE"
		Case ISSSTE_1600_REPORTS
			GetReportNameByConstant = "Generación de volantes"
		Case ISSSTE_1602_REPORTS
			GetReportNameByConstant = "Impresión de guías"
		Case ISSSTE_1603_REPORTS
			GetReportNameByConstant = "Reporte de estatus de documentos"
		Case ISSSTE_1604_REPORTS
			GetReportNameByConstant = "Reportes de asuntos defasados/resueltos"
		Case ISSSTE_1605_REPORTS
			GetReportNameByConstant = "Generación de oficios"
		Case ISSSTE_1606_REPORTS
			GetReportNameByConstant = "Reporte de los empleados con licencia sindical"
		Case ISSSTE_1607_REPORTS
			GetReportNameByConstant = "Concentrado de control de correspondencia"
		Case ISSSTE_1610_REPORTS
			GetReportNameByConstant = "Número de asuntos recibidos por destinatario"
		Case ISSSTE_1611_REPORTS
			GetReportNameByConstant = "Número de asuntos recibidos por rangos de fecha"
		Case ISSSTE_1612_REPORTS
			GetReportNameByConstant = "Número de asuntos recibidos por destinatario y rangos de fecha"
		Case ISSSTE_1613_REPORTS
			GetReportNameByConstant = "Asuntos pendientes de descargo"
		Case ISSSTE_1701_REPORTS
			GetReportNameByConstant = "Consulta de presupuesto"
	End Select

	Err.Clear
End Function

Function GetReportPathByConstant(lSectionID, lReportNumber)
'************************************************************
'Purpose: To get the name of the report given its constant
'Inputs:  lSectionID, lReportNumber
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetReportPathByConstant"
	Dim sPath1
	Dim sPath4
	Dim sPath42
	Dim sPath47

	If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
		sPath1 = "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=71"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=713"">Reportes</A> > "
		sPath4 = "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=733"">Reportes</A> > "
		sPath42 = "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > "
		sPath47 = "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=47"">Cheques y depósitos</A> > "
    ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
        sPath1 = "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=34"">Reportes</A> > "
	Else
		sPath1 = "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=19"">Reportes</A> > "
		sPath4 = "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=49"">Reportes</A> > "
		sPath42 = "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > "
		sPath47 = "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=47"">Cheques y depósitos</A> > "
	End If
	Select Case CLng(lReportNumber)
		Case LOGS_HISTORY_REPORTS
			GetReportPathByConstant = "<B>Historial de entradas al sistema</B>"
		Case AREAS_COUNT_REPORTS
			GetReportPathByConstant = "<B>Conteo de centros de trabajo</B>"
		Case EMPLOYEES_COUNT_REPORTS
			If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=19"">Reportes</A> > <B>Conteo de empleados</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 2 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=24"">Reportes</A> > <B>Conteo de empleados</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=34"">Reportes</A> > <B>Conteo de empleados</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=49"">Reportes</A> > <B>Conteo de empleados</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=71"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=71"">Reportes</A> > <B>Conteo de empleados</B>"
			Else
				GetReportPathByConstant = "<B>Conteo de empleados</B>"
			End If
		Case JOBS_COUNT_REPORTS
			GetReportPathByConstant = "<B>Conteo de plazas</B>"
		Case AREAS_LIST_REPORTS
			GetReportPathByConstant = "<B>Información de los centros de trabajo</B>"
		Case EMPLOYEES_LIST_REPORTS
			If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=19"">Reportes</A> > <B>Información de los empleados</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 2 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=24"">Reportes</A> > <B>Información de los empleados</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=34"">Reportes</A> > <B>Información de los empleados</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=49"">Reportes</A> > <B>Información de los empleados</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=71"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=71"">Reportes</A> > <B>Información de los empleados</B>"
			Else
				GetReportPathByConstant = "<B>Información de los empleados</B>"
			End If
		Case JOBS_LIST_REPORTS, SPECIAL_JOBS_LIST_REPORTS
			If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=71"">Personal</A> > <B> <A HREF=""Jobs.asp""> Administración de plazas</A></B> >"
            ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
                GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <B> <A HREF=""Jobs.asp""> Administración de plazas</A></B> >"
			Else
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <B> <A HREF=""Jobs.asp""> Administración de plazas</A></B> >"
			End If
			If StrComp(oRequest("JobStatusID").Item, "1", vbBinaryCompare) = 0 Then
				GetReportPathByConstant = GetReportPathByConstant & "<B>Plazas ocupadas</B>"
			ElseIf StrComp(oRequest("JobStatusID").Item, "2", vbBinaryCompare) = 0 Then
				GetReportPathByConstant = GetReportPathByConstant & "<B>Plazas vacantes</B>"
			ElseIf StrComp(oRequest("JobStatusID").Item, "3", vbBinaryCompare) = 0 Then
				GetReportPathByConstant = GetReportPathByConstant & "<B>Plazas congeladas</B>"
			ElseIf StrComp(oRequest("JobStatusID").Item, "4", vbBinaryCompare) = 0 Then
				GetReportPathByConstant = GetReportPathByConstant & "<B>Plazas con licencia</B>"
			ElseIf Len(oRequest("MovedEmployees").Item) > 0 Then
				GetReportPathByConstant = GetReportPathByConstant & "<B>Plazas en interinatos</B>"
			ElseIf Len(oRequest("JobsOwners").Item) > 0 Then
				GetReportPathByConstant = GetReportPathByConstant & "<B>Plazas reservadas</B>"
			Else
				GetReportPathByConstant = GetReportPathByConstant & "<B>Reporte de plazas por estatus</B>"
			End If
		Case EMPLOYEE_HISTORY_LIST_REPORTS
			GetReportPathByConstant = "<B>Historial del empleado</B>"
		Case EMPLOYEE_FORM_HISTORY_LIST_REPORTS
			GetReportPathByConstant = "<B>Historial de cambios del formato FM1</B>"
		Case EMPLOYEE_PAYMENTS_HISTORY_LIST_REPORTS
			GetReportPathByConstant = "<B>Historial de pagos</B>"
		Case EMPLOYEE_PAYROLL_REPORTS
			GetReportPathByConstant = "<B>Pagos por quincena</B>"
		Case JOBS_LIST_BY_MODIFY_DATE 
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Jobs.asp"">Plazas</A> > <B>Plazas Creadas o modificadas</B>"
		Case ISSSTE_1001_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Hoja de cifras</B>"
		Case ISSSTE_1002_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Revisión de diferencias</B>"
		Case ISSSTE_1003_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Listado de firmas</B>"
		Case ISSSTE_1004_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Resumen por conceptos de nómina</B>"
		Case ISSSTE_1005_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Hoja informativa de servicios de informática</B>"
		Case ISSSTE_1006_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Concentrado de conceptos de la nómina ordinaria</B>"
		Case ISSSTE_1007_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Remesa para cubrir la nómina</B>"
		Case ISSSTE_1008_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Empleados por tipo de tabulador y por empresa</B>"
		Case ISSSTE_1009_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Resumen mensual de nóminas</B>"
		Case ISSSTE_1010_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Ramas médica, paramédica, de grupos afines y operativa, de enlace y de alto nivel de responsabilidad</B>"
		Case ISSSTE_1011_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Pensión alimenticia de ramas médica, paramédica, de grupos afines y operativa, de enlace y de alto nivel de responsabilidad</B>"
		Case ISSSTE_1012_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Diferencias totales por concepto</B>"
		Case ISSSTE_1013_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Diferencias de empleados por unidad administrativa</B>"
		Case ISSSTE_1014_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Altas por unidad administrativa</B>"
		Case ISSSTE_1015_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Altas de empleados por unidad administrativa</B>"
		Case ISSSTE_1016_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Bajas por unidad administrativa</B>"
		Case ISSSTE_1017_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Bajas de empleados por unidad administrataiva</B>"
		Case ISSSTE_1018_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Diferencias de sueldo por unidad administrativa</B>"
		Case ISSSTE_1019_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Cambios de puesto por unidad administrativa</B>"
		Case ISSSTE_1020_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Funcionarios con líquido mayor a la suma de sueldo base más compensación</B>"
		Case ISSSTE_1021_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Totales por nómina</B>"
		Case ISSSTE_1022_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Empleados con líquidos mayores</B>"
		Case ISSSTE_1023_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Reporte de movimientos</B>"
		Case ISSSTE_1024_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Funcionarios y operativos por concepto de pago y empresa</B>"
		Case ISSSTE_1025_REPORTS
			GetReportPathByConstant = "<B>Recibos de pago por empleado</B>"
		Case ISSSTE_1026_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Incidencias de los empleados</B>"
		Case ISSSTE_1027_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Listado de cheques</B>"
		Case ISSSTE_1028_REPORTS, ISSSTE_1031_REPORTS, ISSSTE_1032_REPORTS, ISSSTE_1033_REPORTS, ISSSTE_1034_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Resumen de Nóminas al SAR</B>"
		Case ISSSTE_1029_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Concentrado de incidencias</B>"
		Case ISSSTE_1030_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Incidencias con horas extras y primas dominicales</B>"
		Case ISSSTE_1100_REPORTS
			GetReportPathByConstant = "<B>Catálogo de trabajadores por centro de pago</B>"
		Case ISSSTE_1101_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <B>Aguinaldos</B>"
		Case ISSSTE_1102_REPORTS
			'GetReportPathByConstant = sPath1 & "<B>Reporte de reclamos de pago por ajustes y deducciones</B>"
			If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
				GetReportPathByConstant = sPath1 & "<B>Reporte de reclamos de pago por ajustes y deducciones</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 2 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=24"">Reportes</A> > <B>Reporte de reclamos de pago por ajustes y deducciones</B>"
			End If
		Case ISSSTE_1103_REPORTS
			GetReportPathByConstant = sPath1 & "<B>Reporte de movimientos en trámite</B>"
		Case ISSSTE_1104_REPORTS
			GetReportPathByConstant = sPath1 & "<B>Reporte de movimientos por usuario</B>"
		Case ISSSTE_1105_REPORTS
			GetReportPathByConstant = sPath1 & "<B>Reporte de registro de movimientos</B>"
		Case ISSSTE_1106_REPORTS
			GetReportPathByConstant = sPath1 & "<B>Reporte de honorarios</B>"
		Case ISSSTE_1107_REPORTS
			GetReportPathByConstant = sPath1 & "<B>Nómina de personal correspondiente a conceptos</B>"
		Case ISSSTE_1108_REPORTS
			If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=19"">Reportes</A> > <B>Conceptos registrados a los empleados</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 2 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=24"">Reportes</A> > <B>Conceptos registrados a los empleados</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=49"">Reportes</A> > <B>Conceptos registrados a los empleados</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=72"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=723"">Reportes</A> > <B>Conceptos registrados a los empleados</B>"
			End If
		Case ISSSTE_1109_REPORTS
			GetReportPathByConstant = sPath1 & "<B>Impresión del formato FM1</B>"
		Case ISSSTE_1110_REPORTS
			GetReportPathByConstant = sPath1 & "<B>Impresión de formato de honorarios</B>"
		Case ISSSTE_1111_REPORTS
			GetReportPathByConstant = sPath1 & "<B>Histórico de plazas</B>"
		Case ISSSTE_1112_REPORTS
			GetReportPathByConstant = sPath1 & "<B>Hoja única de servicio</B>"
		Case ISSSTE_1113_REPORTS
			GetReportPathByConstant = sPath1 & "<B>Reporte de validación del pago de aguinaldo</B>"
		Case ISSSTE_1114_REPORTS
			GetReportPathByConstant = sPath1 & "<B>Pagos cancelados</B>"
		Case ISSSTE_1115_REPORTS
			GetReportPathByConstant = sPath1 & "<B>Impresión de formato de baja de honorarios</B>"
		Case ISSSTE_1116_REPORTS
			GetReportPathByConstant = sPath1 & "<B>Antigüedad para un empleado</B>"
		Case ISSSTE_1117_REPORTS
			GetReportPathByConstant = sPath1 & "<B>Reporte de Antigüedades</B>"
		Case ISSSTE_1118_REPORTS
			GetReportPathByConstant = sPath1 & "<B>Validación de nómina 1o de Octubre</B>"
		Case ISSSTE_1119_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=24"">Reportes</A> > <B>Reporte de empleados con derecho al concepto 41</B>"
		Case ISSSTE_1151_REPORTS
			'GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=15"">Acumulados anuales</A> > <B>Acumulados anuales</B>"
			If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=15"">Acumulados anuales</A> > <B>Reporte de Acumulados anuales</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 2 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=291"">Acumulados anuales</A> > <B>Reporte de Acumulados anuales</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=72"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=723"">Acumulados anuales</A> > <B>Reporte de Acumulados anuales</B>"
			End If
		Case ISSSTE_1152_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=15"">Acumulados anuales</A> > <B>Constancia de percepciones y deducciones anuales</B>"
		Case ISSSTE_1153_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=15"">Acumulados anuales</A> > <B>Ajuste anual del impuesto sobre la renta</B>"
		Case ISSSTE_1154_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=15"">Acumulados anuales</A> > <B>Recálculo anual de impuestos</B>"
		Case ISSSTE_1155_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=15"">Acumulados anuales</A> > <B>Aplicación del ajuste anual del impuesto sobre la renta</B>"
		Case ISSSTE_1157_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=15"">Acumulados anuales</A> > <B>Declaración informativa múltiple (DIM)</B>"
		Case ISSSTE_1200_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=24"">Reportes</A> > <B>Funcionarios y operativos por concepto de pago y empresa</B>"
		Case ISSSTE_1201_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=24"">Reportes</A> > <B>Reporte de personal con conceptos</B>"
		Case ISSSTE_1202_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=24"">Reportes</A> > <B>Reporte de personal con créditos</B>"
		Case ISSSTE_1203_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=25"">Certificaciones y archivo</A> > <B>Hoja única de servicio</B>"
		Case ISSSTE_1204_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=25"">Certificaciones y archivo</A> > <B>Antigüedad para un empleado</B>"
		Case ISSSTE_1205_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=25"">Certificaciones y archivo</A> > <B>Reporte de antigüedades</B>"
		Case ISSSTE_1206_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=25"">Certificaciones y archivo</A> > <B>Validación de nómina 1o de Octubre</B>"
		Case ISSSTE_1207_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=25"">Certificaciones y archivo</A> > <B>Constancia de servicio activo</B>"
		Case ISSSTE_1208_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=25"">Certificaciones y archivo</A> > <B>Constancia de descuento</B>"
		Case ISSSTE_1209_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=24"">Reportes</A> > <B>Reporte de revisión de nóminas</B>"
		Case ISSSTE_1210_REPORTS
			If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=19"">Reportes</A> > <B>Reporte de horas extras y primas dominicales</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 2 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=24"">Reportes</A> > <B>Reporte de horas extras y primas dominicales</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=72"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=723"">Reportes</A> > <B>Reporte de horas extras y primas dominicales</B>"
			End If
		Case ISSSTE_1211_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=21"">Terceros institucionales</A> > <B>Calificación de médicos y enfermeras</B>"
		Case ISSSTE_1221_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=21"">Terceros institucionales</A> > <B>Reporte de registros cargados desde archivo de terceros</B>"
		Case ISSSTE_1222_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=21"">Terceros institucionales</A> > <B>Reporte de registros rechazados desde archivos de terceros</B>"
		Case ISSSTE_1223_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=23"">Pensión alimenticia</A> > <B>Reporte de pensiones alimenticias por empleado</B>"
		Case ISSSTE_1224_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=23"">Pensión alimenticia</A> > <B>Reporte de empleados con pensiones alimenticias</B>"
		Case ISSSTE_1225_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=21"">Terceros institucionales</A> > <B>Generación de archivo Repcsi</B>"
		Case ISSSTE_1311_REPORTS
			If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=19"">Reportes</A> > <B>Plantilla de nómina</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 2 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=24"">Reportes</A> > <B>Plantilla de nómina</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=34"">Reportes</A> > <B>Plantilla de nómina</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=49"">Reportes</A> > <B>Plantilla de nómina</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=71"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=71"">Reportes</A> > <B>Plantilla de nómina</B>"
			Else
				GetReportPathByConstant = "<B>Plantilla de nómina</B>"
			End If
		Case ISSSTE_1334_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=34"">Reportes</A> > <B>Generación del reporte de Unidades Médicas</B>"
		Case ISSSTE_1335_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=34"">Reportes</A> > <B>Catálogo de puestos y tabuladores de puestos</B>"
		Case ISSSTE_1336_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=34"">Reportes</A> > <B>Catálogo de centros de trabajo</B>"
		Case ISSSTE_1337_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=34"">Reportes</A> > <B>Catálogo de centros de pago</B>"
		Case ISSSTE_1338_REPORTS
			If CInt(Request.Cookies("SIAP_SectionID")) = 7  Then
                GetReportPathByConstant = sPath42 & "<B>Revisión de Nóminas</B>"
            Else
                GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=49"">Reportes</A> > <B>Revisión de nóminas</B>"
		    End If
        Case ISSSTE_1339_REPORTS        
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=49"">Reportes</A> > <B>Archivo SICAD</B>"
        Case ISSSTE_1340_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=49"">Reportes</A> > <B>Archivo SICAD de cancelaciones</B>"
		Case ISSSTE_1354_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=28"">Registro de bolsa de trabajo y escalafón</A> > <B>Información registrada en la bolsa de trabajo</B>"
		Case ISSSTE_1356_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=28"">Registro de bolsa de trabajo y escalafón</A> > <B>Búsqueda de información de escalafón</B>"
		Case ISSSTE_1364_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=36"">Desarrollo Humano</A> > <B>Reportes</B>"
		Case ISSSTE_1365_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=36"">Desarrollo Humano</A> > <B>Seguimiento de cursos de capacitación</B>"
		Case ISSSTE_1367_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=36"">Desarrollo Humano</A> > <B>Reporte de curriculum por empleado</B>"
		Case ISSSTE_1369_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=36"">Desarrollo humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=369"">Registro de detección de necesidades</A> > <B>Búsqueda de información</B>"
		Case ISSSTE_1371_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=34"">Reportes</A> > <B>Generación de archivo para registro de servidores públicos</B>"
		Case ISSSTE_1372_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=34"">Reportes</A> > <B>RUSP. Información básica</B>"
		Case ISSSTE_1373_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=34"">Reportes</A> > <B>RUSP. Bajas</B>"
		Case ISSSTE_1374_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=34"">Reportes</A> > <B>RUSP. Datos personales</B>"
		Case ISSSTE_1400_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Catalogs.asp"">Catálogos</A> > <B>CLCs</B>"
		Case ISSSTE_1401_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Archivo para carga del SPEP</B>"
		Case ISSSTE_1402_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Archivo para carga del SPEP por centro de trabajo</B>"
		Case ISSSTE_1403_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Catalogs.asp"">Catálogos</A> > <B>Resumen mensual de nóminas</B>"
		Case ISSSTE_1404_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Fajillas por estado</B>"
		Case ISSSTE_1411_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Cifras iniciales</B>"
		Case ISSSTE_1412_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Archivo para carga del SPEP del FONAC</B>"
		Case ISSSTE_1413_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Archivo para contabilidad</B>"
		Case ISSSTE_1414_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Detalle de las aportaciones por quincena</B>"
		Case ISSSTE_1415_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Reporte para el Fiduciario</B>"
		Case ISSSTE_1416_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Respaldo de los empleados cotizantes por quincena</B>"
		Case ISSSTE_1417_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Cifras para el pago de nómina</B>"
		Case ISSSTE_1420_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Reporte de personal interno</B>"
		Case ISSSTE_1421_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Reporte de los conceptos 15 (guardias), 31 (suplencias) y C5 (guardias PROVAC) de personal interno</B>"
		Case ISSSTE_1422_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Reporte de captura de personal externo</B>"
		Case ISSSTE_1423_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Reporte de validación de personal externo</B>"
		Case ISSSTE_1424_REPORTS
			If Len(oRequest("External").Item) = 0 Then
				GetReportPathByConstant = sPath4 & "<B>Consolidado de personal interno</B>"
			Else
				GetReportPathByConstant = sPath4 & "<B>Consolidado de los prestadores de servicio</B>"
			End If
		Case ISSSTE_1425_REPORTS
			If Len(oRequest("External").Item) = 0 Then
				GetReportPathByConstant = sPath4 & "<B>Validación de captura de personal interno</B>"
			Else
				GetReportPathByConstant = sPath4 & "<B>Validación de captura de prestadores de servicios</B>"
			End If
		Case ISSSTE_1426_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Volantes</B>"
		Case ISSSTE_1427_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Listado de firmas</B>"
		Case ISSSTE_1428_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Reporte de totales</B>"
		Case ISSSTE_1429_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Reporte concentrado por quincena</B>"
		Case ISSSTE_1430_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Reporte estadístico de causas</B>"
		Case ISSSTE_1431_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Reporte de cuentas bancarias</B>"
		Case ISSSTE_1432_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Listado de actualización de cuentas bancarias</B>"
		Case ISSSTE_1433_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Reporte de estímulos</B>"
		Case ISSSTE_1434_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Reporte de incidencias</B>"
		Case ISSSTE_1435_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Reporte de tabuladores de pago</B>"
		Case ISSSTE_1490_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Reporte de cifras</B>"
		Case ISSSTE_1491_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Salida de archivos de terceros</B>"
		Case ISSSTE_1492_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Reportes de terceros</B>"
		Case ISSSTE_1493_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Archivo para carga del SPEP por concepto</B>"
		Case ISSSTE_1494_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Memoria de cálculo para el entero de cuotas sindicales</B>"
		Case ISSSTE_1470_REPORTS
			GetReportPathByConstant = sPath47 & "<B>Listado de firmas</B>"
		Case ISSSTE_1471_REPORTS
			GetReportPathByConstant = sPath47 & "<B>Recibo de distribución y recepción de cheques</B>"
		Case ISSSTE_1472_REPORTS
			GetReportPathByConstant = sPath47 & "<B>Pagos cancelados</B>"
		Case ISSSTE_1473_REPORTS
			GetReportPathByConstant = sPath47 & "<B>Bloqueos aplicados</B>"
		Case ISSSTE_1474_REPORTS
			GetReportPathByConstant = sPath47 & "<B>Archivo de depósitos bancarios</B>"
		Case ISSSTE_1475_REPORTS
			GetReportPathByConstant = sPath47 & "<B>Archivo de liberación de cheques</B>"
		Case ISSSTE_1476_REPORTS
			GetReportPathByConstant = sPath47 & "<B>Recibo por pago de honorarios</B>"
		Case ISSSTE_1477_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Impuesto sobre la renta</B>"
		Case ISSSTE_1478_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Cálculo del impuesto sobre nóminas</B>"
		Case ISSSTE_1499_REPORTS
			GetReportPathByConstant = sPath4 & "<B>Constancia de percepciones y deducciones</B>"
		Case ISSSTE_1503_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=5"">Presupuesto</A> > <B>Costeo de plazas</B>"
		Case ISSSTE_1504_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=5"">Presupuesto</A> > <B>Consulta de presupuesto</B>"
		Case ISSSTE_1561_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=5"">Presupuesto</A> > <A HREF=""Main_ISSSTE.asp?SectionID=56"">Reportes sobre el costeo de plazas</A> > <B>Proyecto de presupuesto</B>"
		Case ISSSTE_1562_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=5"">Presupuesto</A> > <A HREF=""Main_ISSSTE.asp?SectionID=56"">Reportes sobre el costeo de plazas</A> > <B>Archivo de carga para el SPEP</B>"
		Case ISSSTE_1563_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=5"">Presupuesto</A> > <A HREF=""Main_ISSSTE.asp?SectionID=56"">Reportes sobre el costeo de plazas</A> > <B>Formato único de movimientos presupuestales</B>"
		Case ISSSTE_1571_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=5"">Presupuesto</A> > <B>Registro de un costeo como original</B>"
		Case ISSSTE_1581_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=5"">Presupuesto</A> > <A HREF=""Main_ISSSTE.asp?SectionID=58"">Reportes</A> > <B>Personal ocupado por rama de actividad</B>"
		Case ISSSTE_1582_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=5"">Presupuesto</A> > <A HREF=""Main_ISSSTE.asp?SectionID=58"">Reportes</A> > <B>Personal ocupado y pago de sueldos y salarios en la administración pública federal</B>"
		Case ISSSTE_1583_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=5"">Presupuesto</A> > <A HREF=""Main_ISSSTE.asp?SectionID=58"">Reportes</A> > <B>Prestaciones a favor de los servidores públicos del ISSSTE</B>"
		Case ISSSTE_1584_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=5"">Presupuesto</A> > <A HREF=""Main_ISSSTE.asp?SectionID=58"">Reportes</A> > <B>Trabajadores cotizantes al régimen del ISSSTE</B>"
		Case ISSSTE_1600_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Generación de volantes</B>"
		Case ISSSTE_1602_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Impresión de guías</B>"
		Case ISSSTE_1603_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Reporte de estatus de documentos</B>"
		Case ISSSTE_1604_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Reportes de asuntos defasados/resueltos</B>"
		Case ISSSTE_1605_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=6"">Departamento técnico</A> > <A HREF=""Main_ISSSTE.asp?SectionID=62"">Emisión de licencias por comisión sindical</A> > <B>Generación de oficios</B>"
		Case ISSSTE_1606_REPORTS
			If CInt(Request.Cookies("SIAP_SectionID")) = 2 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Reporte de los empleados con licencia sindical</B>"
            ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 8 Then
                GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=8"">Atención al personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Trámites al personal</A> > <B>Concentrado de control de correspondencia</B>"
			Else
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=6"">Departamento técnico</A> > <A HREF=""Main_ISSSTE.asp?SectionID=64"">Reportes</A> > <B>Reporte de los empleados con licencia sindical</B>"
			End If
		Case ISSSTE_1607_REPORTS
			If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Concentrado de control de correspondencia</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 2 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Concentrado de control de correspondencia</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Concentrado de control de correspondencia</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Concentrado de control de correspondencia</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 5 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=5"">Presupuesto</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Concentrado de control de correspondencia</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 6 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=6"">Departamento técnico</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Concentrado de control de correspondencia</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Concentrado de control de correspondencia</B>"
            ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 8 Then
                GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=8"">Atención al personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Trámites al personal</A> > <B>Concentrado de control de correspondencia</B>"
			Else
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Concentrado de control de correspondencia</B>"
			End If
		Case ISSSTE_1608_REPORTS
			If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Reporte de documentos</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 2 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Reporte de documentos</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Reporte de documentos</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Reporte de documentos</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 5 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=5"">Presupuesto</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Reporte de documentos</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 6 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=6"">Departamento técnico</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Reporte de documentos</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Reporte de documentos</B>"
            ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 8 Then
                GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=8"">Atención al personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Trámites al personal</A> > <B>Reporte de documentos</B>"
			Else
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Reporte de documentos</B>"
			End If
		Case ISSSTE_1609_REPORTS, ISSSTE_1619_REPORTS
			Dim sReportDescription
			If aReportsComponent(N_ID_REPORTS) = ISSSTE_1609_REPORTS Then
				sReportDescription = "Listas de Entrega"
			Else
				sReportDescription = "Reporte ESP"
			End If
			If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>" & sReportDescription & "</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 2 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>" & sReportDescription & "</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>" & sReportDescription & "</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>" & sReportDescription & "</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 5 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=5"">Presupuesto</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>" & sReportDescription & "</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 6 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=6"">Departamento técnico</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>" & sReportDescription & "</B>"
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>" & sReportDescription & "</B>"
            ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 8 Then
                GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=8"">Atención al personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Trámites al personal</A> > <B>" & sReportDescription & "</B>"
			Else
				GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>" & sReportDescription & "</B>"
			End If
		Case ISSSTE_1610_REPORTS
            If CInt(Request.Cookies("SIAP_SectionID")) = 8 Then
                GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=8"">Atención al personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Trámites al personal</A> > <B>Concentrado de control de correspondencia</B>"
			Else
			    GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Número de asuntos recibidos por destinatario</B>"
            End If
		Case ISSSTE_1611_REPORTS
            If CInt(Request.Cookies("SIAP_SectionID")) = 8 Then
                GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=8"">Atención al personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Trámites al personal</A> > <B>Concentrado de control de correspondencia</B>"
			Else
			    GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Número de asuntos recibidos por rangos de fecha</B>"
            End If
		Case ISSSTE_1612_REPORTS
            If CInt(Request.Cookies("SIAP_SectionID")) = 8 Then
                GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=8"">Atención al personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Trámites al personal</A> > <B>Concentrado de control de correspondencia</B>"
			Else
			    GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Número de asuntos recibidos por destinatario y rangos de fecha</B>"
            End If
		Case ISSSTE_1613_REPORTS
            If CInt(Request.Cookies("SIAP_SectionID")) = 8 Then
                GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=8"">Atención al personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Trámites al personal</A> > <B>Concentrado de control de correspondencia</B>"
			Else
			    GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > <B>Asuntos pendientes de descargo</B>"
            End If
		Case ISSSTE_1701_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=75"">Presupuesto</A> > <B>Consulta de presupuesto</B>"
		Case ISSSTE_1702_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=72"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=723"">Reportes</A> > <B>Antigüedad para un empleado</B>"
		Case ISSSTE_1703_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=72"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=723"">Reportes</A> > <B>Reporte de antigüedades</B>"
		Case ISSSTE_1704_REPORTS
			GetReportPathByConstant = "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=72"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=723"">Reportes</A> > <B>Validación de nómina 1o de Octubre</B>"
		Case ISSSTE_2420_REPORTS
			GetReportPathByConstant = sPath42 & "<B>Reporte de personal interno</B>"
		Case ISSSTE_2421_REPORTS
			GetReportPathByConstant = sPath42 & "<B>Reporte de los conceptos 15 (guardias), 31 (suplencias) y C5 (guardias PROVAC) de personal interno</B>"
		Case ISSSTE_2422_REPORTS
			GetReportPathByConstant = sPath42 & "<B>Reporte de captura de personal externo</B>"
		Case ISSSTE_2423_REPORTS
			GetReportPathByConstant = sPath42 & "<B>Reporte de validación de personal externo</B>"
		Case ISSSTE_2426_REPORTS
			GetReportPathByConstant = sPath42 & "<B>Volantes</B>"
		Case ISSSTE_2427_REPORTS
			GetReportPathByConstant = sPath42 & "<B>Listado de firmas</B>"
		Case ISSSTE_2428_REPORTS
			GetReportPathByConstant = sPath42 & "<B>Reporte de totales</B>"
		Case ISSSTE_2429_REPORTS
			GetReportPathByConstant = sPath42 & "<B>Reporte concentrado por quincena</B>"
		Case ISSSTE_2430_REPORTS
			GetReportPathByConstant = sPath42 & "<B>Reporte estadístico de causas</B>"
		Case ISSSTE_2431_REPORTS
			GetReportPathByConstant = sPath42 & "<B>Reporte de cuentas bancarias</B>"
		Case ISSSTE_2432_REPORTS
			GetReportPathByConstant = sPath42 & "<B>Listado de actualización de cuentas bancarias</B>"
		Case ISSSTE_4701_REPORTS
			GetReportPathByConstant = sPath47 & "<B>Listado de firmas de cancelaciones</B>"
		Case ISSSTE_4702_REPORTS
			GetReportPathByConstant = sPath47 & "<B>Concentrado de conceptos de cancelaciones</B>"
		Case ISSSTE_4703_REPORTS
			GetReportPathByConstant = sPath47 & "<B>Reporte de cifras de cancelaciones</B>"
	End Select

	Err.Clear
End Function

Function SendReportAlert(sFileName, lDate, sErrorDescription)
'************************************************************
'Purpose: To send an alert letting the user know the report
'         is ready
'Inputs:  sFileName, lDate
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SendReportAlert"
	Dim lErrorNumber

	ReDim aEmailComponent(N_EMAIL_COMPONENT_SIZE)
	aEmailComponent(S_TO_EMAIL) = aLoginComponent(S_USER_E_MAIL_LOGIN)
	aEmailComponent(S_FROM_EMAIL) = aLoginComponent(S_USER_E_MAIL_LOGIN)
	aEmailComponent(S_SUBJECT_EMAIL) = "SIAP. Su reporte está listo"
	aEmailComponent(S_BODY_EMAIL) = GetFileContents(Server.MapPath("Template_ReportReady.htm"), sErrorDescription)
	If Len(aEmailComponent(S_BODY_EMAIL)) > 0 Then
		aEmailComponent(S_BODY_EMAIL) = Replace(aEmailComponent(S_BODY_EMAIL), "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
		aEmailComponent(S_BODY_EMAIL) = Replace(aEmailComponent(S_BODY_EMAIL), "<USER_NAME />", CleanStringForHTML(aLoginComponent(S_USER_NAME_LOGIN) & " " & aLoginComponent(S_USER_LAST_NAME_LOGIN)))
		aEmailComponent(S_BODY_EMAIL) = Replace(aEmailComponent(S_BODY_EMAIL), "<REPORT_NAME />", GetReportNameByConstant(aReportsComponent(N_ID_REPORTS)))
		aEmailComponent(S_BODY_EMAIL) = Replace(aEmailComponent(S_BODY_EMAIL), "<REPORT_DATE />", DisplayDateFromSerialNumber(lDate, -1, -1, -1))
		aEmailComponent(S_BODY_EMAIL) = Replace(aEmailComponent(S_BODY_EMAIL), "<REPORT_URL />", sFileName)
		lErrorNumber = SendEmail(oRequest, aEmailComponent, sErrorDescription)
	End If

	SendReportAlert = lErrorNumber
	Err.Clear
End Function
%>