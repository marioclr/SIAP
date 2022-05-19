<%
Const N_USER_ID_CALENDAR = 0
Const N_GROUP_ID_CALENDAR = 1
Const N_YEAR_CALENDAR = 2
Const N_MONTH_CALENDAR = 3
Const N_DAY_CALENDAR = 4
Const N_HOUR_CALENDAR = 5
Const N_MINUTE_CALENDAR = 6
Const N_TASK_ID_CALENDAR = 7
Const S_MONTH_CALENDAR = 8
Const N_WEEK_CALENDAR = 9
Const S_DAY_CALENDAR = 10
Const N_FIRST_DAY_CALENDAR = 11
Const N_DAYS_CALENDAR = 12
Const B_GRAY_SUNDAY_CALENDAR = 13
Const N_SELECTED_DAY_CALENDAR = 14
Const S_MARKED_DAYS_CALENDAR = 15
Const S_SPECIAL_DAYS_CALENDAR = 16
Const S_TASK_COLOR_CALENDAR = 17
Const S_TASK_TITLE_CALENDAR = 18
Const S_TASK_DESCRIPTION_CALENDAR = 19
Const N_DURATION_CALENDAR = 20
Const N_TASK_TYPE_CALENDAR = 21
Const S_TASK_TYPE_CALENDAR = 22
Const B_TASK_DONE_CALENDAR = 23
Const N_TASK_COUNT_CALENDAR = 24
Const S_TARGET_PAGE_CALENDAR = 25
Const S_TARGET_FRAME_CALENDAR = 26
Const S_JAVASCRIPT_CALENDAR = 27
Const N_OLD_USER_ID_CALENDAR = 28
Const N_OLD_GROUP_ID_CALENDAR = 29
Const N_OLD_YEAR_CALENDAR = 30
Const N_OLD_MONTH_CALENDAR = 31
Const N_OLD_DAY_CALENDAR = 32
Const N_OLD_HOUR_CALENDAR = 33
Const N_OLD_MINUTE_CALENDAR = 34
Const N_OLD_TASK_ID_CALENDAR = 35
Const B_ONLY_SUNDAY_CALENDAR = 36
Const B_ONLY_PAYDAYS_CALENDAR = 37
Const B_ONLY_HOLIDAYS_CALENDAR = 38
Const B_COMPONENT_INITIALIZED_CALENDAR = 39

Const N_CALENDAR_COMPONENT_SIZE = 39

Const N_MAX_COLORS = 12

Dim asMonthNamesCalendar
asMonthNamesCalendar = Split("Enero,Febrero,Marzo,Abril,Mayo,Junio,Julio,Agosto,Septiembre,Octubre,Noviembre,Diciembre", ",", -1, vbBinaryCompare)
Dim asDayNamesCalendar
asDayNamesCalendar = Split("Domingo,Lunes,Martes,Miércoles,Jueves,Viernes,Sábado", ",", -1, vbBinaryCompare)
Dim asTaskTypesCalendar
asTaskTypesCalendar = Split("Tiempo Libre,Actividad,Junta,Cita,Comida,Clase,Revisión Médica,Reunión,Exposición,Taller", ",", -1, vbBinaryCompare)
Dim asTaskColorsCalendar
asTaskColorsCalendar = Split("000000,FF0000,000066,006600,666600,660066,006666,663300,660033,336600,006633,330066,003366", ",", -1, vbBinaryCompare)
Dim asTaskColorNamesCalendar
asTaskColorNamesCalendar = Split("Negro,Rojo,Azul Oscuro,Verde Oscuro,Mostaza,Morado claro,Azul cielo,Café,Guinda,Verde claro,Verde cielo,Morado oscuro,Azul", ",", -1, vbBinaryCompare)

Dim aCalendarComponent()
Redim aCalendarComponent(N_CALENDAR_COMPONENT_SIZE)

Function InitializeCalendarComponent(oRequest, aCalendarComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Calendar
'         Component using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aCalendarComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeCalendarComponent"
	Redim Preserve aCalendarComponent(N_CALENDAR_COMPONENT_SIZE)

	If IsEmpty(aCalendarComponent(N_USER_ID_CALENDAR)) Then
		If Len(oRequest("AllGroupActivity").Item) > 0 Then
			aCalendarComponent(N_USER_ID_CALENDAR) = -1
		ElseIf Len(oRequest("UserID").Item) > 0 Then
			aCalendarComponent(N_USER_ID_CALENDAR) = CLng(oRequest("UserID").Item)
		Else
			aCalendarComponent(N_USER_ID_CALENDAR) = aLoginComponent(N_USER_ID_LOGIN)
		End If
	End If

	If IsEmpty(aCalendarComponent(N_GROUP_ID_CALENDAR)) Then
		If Len(oRequest("AllGroupsActivity").Item) > 0 Then
			aCalendarComponent(N_GROUP_ID_CALENDAR) = -1
			aCalendarComponent(N_USER_ID_CALENDAR) = -1
		ElseIf Len(oRequest("GroupID").Item) > 0 Then
			aCalendarComponent(N_GROUP_ID_CALENDAR) = CLng(oRequest("GroupID").Item)
		Else
			aCalendarComponent(N_GROUP_ID_CALENDAR) = -1
		End If
	End If

	If IsEmpty(aCalendarComponent(N_YEAR_CALENDAR)) Then
		If Len(oRequest("Year").Item) > 0 Then
			aCalendarComponent(N_YEAR_CALENDAR) = CInt(oRequest("Year").Item)
		ElseIf Len(oRequest("Date").Item) > 0 Then
			aCalendarComponent(N_YEAR_CALENDAR) = Year(oRequest("Date").Item)
		ElseIf Len(oRequest("TimeStamp").Item) > 0 Then
			aCalendarComponent(N_YEAR_CALENDAR) = Year(oRequest("TimeStamp").Item)
		Else
			aCalendarComponent(N_YEAR_CALENDAR) = Year(Now())
		End If
	End If

	If IsEmpty(aCalendarComponent(N_MONTH_CALENDAR)) Then
		If Len(oRequest("Month").Item) > 0 Then
			aCalendarComponent(N_MONTH_CALENDAR) = CInt(oRequest("Month").Item)
			aCalendarComponent(N_MONTH_CALENDAR) = (aCalendarComponent(N_MONTH_CALENDAR) mod 12)
			If aCalendarComponent(N_MONTH_CALENDAR) = 0 Then
				aCalendarComponent(N_MONTH_CALENDAR) = 12
			ElseIf aCalendarComponent(N_MONTH_CALENDAR) < 1 Then
				aCalendarComponent(N_MONTH_CALENDAR) = 1
			End If
		ElseIf Len(oRequest("Date").Item) > 0 Then
			aCalendarComponent(N_MONTH_CALENDAR) = Month(oRequest("Date").Item)
		ElseIf Len(oRequest("TimeStamp").Item) > 0 Then
			aCalendarComponent(N_MONTH_CALENDAR) = Month(oRequest("TimeStamp").Item)
		Else
			aCalendarComponent(N_MONTH_CALENDAR) = Month(Now())
		End If
	End If
	If IsEmpty(aCalendarComponent(S_MONTH_CALENDAR)) And (aCalendarComponent(N_MONTH_CALENDAR) > 0) Then
		aCalendarComponent(S_MONTH_CALENDAR) = asMonthNamesCalendar(aCalendarComponent(N_MONTH_CALENDAR) - 1)
	End If

	If IsEmpty(aCalendarComponent(N_DAY_CALENDAR)) Then
		If Len(oRequest("Day").Item) > 0 Then
			aCalendarComponent(N_DAY_CALENDAR) = CInt(oRequest("Day").Item)
			If aCalendarComponent(N_DAY_CALENDAR) < 1 Then
				aCalendarComponent(N_DAY_CALENDAR) = 1
			End If
		ElseIf Len(oRequest("Date").Item) > 0 Then
			aCalendarComponent(N_DAY_CALENDAR) = Day(oRequest("Date").Item)
		ElseIf Len(oRequest("TimeStamp").Item) > 0 Then
			aCalendarComponent(N_DAY_CALENDAR) = Day(oRequest("TimeStamp").Item)
		Else
			aCalendarComponent(N_DAY_CALENDAR) = Day(Now())
		End If
	End If
	If IsEmpty(aCalendarComponent(S_DAY_CALENDAR)) Then
		aCalendarComponent(S_DAY_CALENDAR) = asDayNamesCalendar(Weekday(DateSerial(aCalendarComponent(N_YEAR_CALENDAR), aCalendarComponent(N_MONTH_CALENDAR), aCalendarComponent(N_DAY_CALENDAR)), vbSunday) - 1)
	End If

	If IsEmpty(aCalendarComponent(N_HOUR_CALENDAR)) Then
		If Len(oRequest("AllDayActivity").Item) > 0 Then
			aCalendarComponent(N_HOUR_CALENDAR) = 25
		ElseIf Len(oRequest("Hour").Item) > 0 Then
			aCalendarComponent(N_HOUR_CALENDAR) = CInt(oRequest("Hour").Item)
			If aCalendarComponent(N_HOUR_CALENDAR) < 0 Then
				aCalendarComponent(N_HOUR_CALENDAR) = 25
			ElseIf aCalendarComponent(N_HOUR_CALENDAR) > 25 Then
				aCalendarComponent(N_HOUR_CALENDAR) = 25
			End If
		ElseIf Len(oRequest("TimeStamp").Item) > 0 Then
			aCalendarComponent(N_HOUR_CALENDAR) = Hour(oRequest("TimeStamp").Item)
		Else
			aCalendarComponent(N_HOUR_CALENDAR) = 25
		End If
	End If

	If IsEmpty(aCalendarComponent(N_MINUTE_CALENDAR)) Then
		If Len(oRequest("AllDayActivity").Item) > 0 Then
			aCalendarComponent(N_MINUTE_CALENDAR) = 0
		ElseIf Len(oRequest("Minute").Item) > 0 Then
			aCalendarComponent(N_MINUTE_CALENDAR) = CInt(oRequest("Minute").Item)
			If aCalendarComponent(N_MINUTE_CALENDAR) < 0 Then
				aCalendarComponent(N_MINUTE_CALENDAR) = 0
			ElseIf aCalendarComponent(N_MINUTE_CALENDAR) > 60 Then
				aCalendarComponent(N_MINUTE_CALENDAR) = 59
			End If
		ElseIf Len(oRequest("TimeStamp").Item) > 0 Then
			aCalendarComponent(N_MINUTE_CALENDAR) = Minute(oRequest("TimeStamp").Item)
		Else
			aCalendarComponent(N_MINUTE_CALENDAR) = 0
		End If
	End If

	If IsEmpty(aCalendarComponent(N_TASK_ID_CALENDAR)) Then
		If Len(oRequest("TaskID").Item) > 0 Then
			aCalendarComponent(N_TASK_ID_CALENDAR) = CLng(oRequest("TaskID").Item)
		Else
			aCalendarComponent(N_TASK_ID_CALENDAR) = -1
		End If
	End If

	If IsEmpty(aCalendarComponent(N_WEEK_CALENDAR)) Then
		If Len(oRequest("Week").Item) > 0 Then
			aCalendarComponent(N_WEEK_CALENDAR) = CInt(oRequest("Week").Item)
			If aCalendarComponent(N_WEEK_CALENDAR) < 1 Then
				aCalendarComponent(N_WEEK_CALENDAR) = 1
			End If
		Else
			aCalendarComponent(N_WEEK_CALENDAR) = -1
		End If
	End If

	If IsEmpty(aCalendarComponent(B_GRAY_SUNDAY_CALENDAR)) Then
		aCalendarComponent(B_GRAY_SUNDAY_CALENDAR) = (Len(oRequest("DontGraySunday").Item) = 0)
	End If

	If IsEmpty(aCalendarComponent(N_FIRST_DAY_CALENDAR)) Then
		aCalendarComponent(N_FIRST_DAY_CALENDAR) = Weekday(DateSerial(aCalendarComponent(N_YEAR_CALENDAR), aCalendarComponent(N_MONTH_CALENDAR), 1), vbSunday)
	End If

	Call GetDaysInMonth(aCalendarComponent(N_YEAR_CALENDAR), aCalendarComponent(N_MONTH_CALENDAR), aCalendarComponent(N_DAYS_CALENDAR))

	If IsEmpty(aCalendarComponent(N_SELECTED_DAY_CALENDAR)) Then
		If Len(oRequest("SelectedDay").Item) > 0 Then
			aCalendarComponent(N_SELECTED_DAY_CALENDAR) = CLng(oRequest("SelectedDay").Item)
		Else
			aCalendarComponent(N_SELECTED_DAY_CALENDAR) = CLng(aCalendarComponent(N_YEAR_CALENDAR) & Right(("0" & aCalendarComponent(N_MONTH_CALENDAR)), Len("00")) & Right(("0" & aCalendarComponent(N_DAY_CALENDAR)), Len("00")))
		End If
	End If

	If IsEmpty(aCalendarComponent(S_MARKED_DAYS_CALENDAR)) Then
		If Len(oRequest("MarkedDays").Item) > 0 Then
			aCalendarComponent(S_MARKED_DAYS_CALENDAR) = oRequest("MarkedDays").Item
		Else
			aCalendarComponent(S_MARKED_DAYS_CALENDAR) = ""
		End If
	End If

	If IsEmpty(aCalendarComponent(S_SPECIAL_DAYS_CALENDAR)) Then
		If Len(oRequest("SpecialDays").Item) > 0 Then
			aCalendarComponent(S_SPECIAL_DAYS_CALENDAR) = oRequest("SpecialDays").Item
		Else
			aCalendarComponent(S_SPECIAL_DAYS_CALENDAR) = ""
		End If
	End If

	If IsEmpty(aCalendarComponent(S_TASK_COLOR_CALENDAR)) Then
		If Len(oRequest("TaskColor").Item) > 0 Then
			aCalendarComponent(S_TASK_COLOR_CALENDAR) = oRequest("TaskColor").Item
		Else
			aCalendarComponent(S_TASK_COLOR_CALENDAR) = "000000"
		End If
	End If
	aCalendarComponent(S_TASK_COLOR_CALENDAR) = Left(aCalendarComponent(S_TASK_COLOR_CALENDAR), 6)

	If IsEmpty(aCalendarComponent(S_TASK_TITLE_CALENDAR)) Then
		If Len(oRequest("TaskTitle").Item) > 0 Then
			aCalendarComponent(S_TASK_TITLE_CALENDAR) = oRequest("TaskTitle").Item
		Else
			aCalendarComponent(S_TASK_TITLE_CALENDAR) = ""
		End If
	End If
	aCalendarComponent(S_TASK_TITLE_CALENDAR) = Left(aCalendarComponent(S_TASK_TITLE_CALENDAR), 255)

	If IsEmpty(aCalendarComponent(S_TASK_DESCRIPTION_CALENDAR)) Then
		If Len(oRequest("TaskDescription").Item) > 0 Then
			aCalendarComponent(S_TASK_DESCRIPTION_CALENDAR) = oRequest("TaskDescription").Item
		Else
			aCalendarComponent(S_TASK_DESCRIPTION_CALENDAR) = ""
		End If
	End If
	aCalendarComponent(S_TASK_DESCRIPTION_CALENDAR) = Left(aCalendarComponent(S_TASK_DESCRIPTION_CALENDAR), 1000)

	If IsEmpty(aCalendarComponent(N_TASK_TYPE_CALENDAR)) Then
		If Len(oRequest("TypeID").Item) > 0 Then
			aCalendarComponent(N_TASK_TYPE_CALENDAR) = CInt(oRequest("TypeID").Item)
		Else
			aCalendarComponent(N_TASK_TYPE_CALENDAR) = 2
		End If
	End If
	aCalendarComponent(S_TASK_TYPE_CALENDAR) = asTaskTypesCalendar(aCalendarComponent(N_TASK_TYPE_CALENDAR))

	aCalendarComponent(B_TASK_DONE_CALENDAR) = CInt(oRequest("TaskDone").Item)
	
	If IsEmpty(aCalendarComponent(N_DURATION_CALENDAR)) Then
		If Len(oRequest("AllDayActivity").Item) > 0 Then
			aCalendarComponent(N_DURATION_CALENDAR) = 0
		ElseIf Len(oRequest("TaskDuration").Item) > 0 Then
			aCalendarComponent(N_DURATION_CALENDAR) = CInt(oRequest("TaskDuration").Item)
		Else
			aCalendarComponent(N_DURATION_CALENDAR) = 0
		End If
	End If

	If IsEmpty(aCalendarComponent(S_TARGET_PAGE_CALENDAR)) Then
		aCalendarComponent(S_TARGET_PAGE_CALENDAR) = GetASPFileName(Request.ServerVariables("PATH_INFO"))
	End If

	If IsEmpty(aCalendarComponent(S_TARGET_FRAME_CALENDAR)) Then
		aCalendarComponent(S_TARGET_FRAME_CALENDAR) = ""
	End If

	If IsEmpty(aCalendarComponent(S_JAVASCRIPT_CALENDAR)) Then
		aCalendarComponent(S_JAVASCRIPT_CALENDAR) = ""
	End If

	If IsEmpty(aCalendarComponent(N_OLD_USER_ID_CALENDAR)) Then
		If Len(oRequest("OldUserID").Item) > 0 Then
			aCalendarComponent(N_OLD_USER_ID_CALENDAR) = CLng(oRequest("OldUserID").Item)
		Else
			aCalendarComponent(N_OLD_USER_ID_CALENDAR) = aCalendarComponent(N_USER_ID_CALENDAR)
		End If
	End If

	If IsEmpty(aCalendarComponent(N_OLD_GROUP_ID_CALENDAR)) Then
		If Len(oRequest("OldGroupID").Item) > 0 Then
			aCalendarComponent(N_OLD_GROUP_ID_CALENDAR) = CLng(oRequest("OldGroupID").Item)
		Else
			aCalendarComponent(N_OLD_GROUP_ID_CALENDAR) = aCalendarComponent(N_GROUP_ID_CALENDAR)
		End If
	End If

	If IsEmpty(aCalendarComponent(N_OLD_YEAR_CALENDAR)) Then
		If Len(oRequest("OldYear").Item) > 0 Then
			aCalendarComponent(N_OLD_YEAR_CALENDAR) = CInt(oRequest("OldYear").Item)
		ElseIf Len(oRequest("OldDate").Item) > 0 Then
			aCalendarComponent(N_OLD_YEAR_CALENDAR) = Year(oRequest("OldDate").Item)
		ElseIf Len(oRequest("OldTimeStamp").Item) > 0 Then
			aCalendarComponent(N_OLD_YEAR_CALENDAR) = Year(oRequest("OldTimeStamp").Item)
		Else
			aCalendarComponent(N_OLD_YEAR_CALENDAR) = aCalendarComponent(N_YEAR_CALENDAR)
		End If
	End If

	If IsEmpty(aCalendarComponent(N_OLD_MONTH_CALENDAR)) Then
		If Len(oRequest("OldMonth").Item) > 0 Then
			aCalendarComponent(N_OLD_MONTH_CALENDAR) = CInt(oRequest("OldMonth").Item)
			aCalendarComponent(N_OLD_MONTH_CALENDAR) = (aCalendarComponent(N_OLD_MONTH_CALENDAR) mod 12)
			If aCalendarComponent(N_OLD_MONTH_CALENDAR) = 0 Then
				aCalendarComponent(N_OLD_MONTH_CALENDAR) = 12
			ElseIf aCalendarComponent(N_OLD_MONTH_CALENDAR) < 1 Then
				aCalendarComponent(N_OLD_MONTH_CALENDAR) = 1
			End If
		ElseIf Len(oRequest("OldDate").Item) > 0 Then
			aCalendarComponent(N_OLD_MONTH_CALENDAR) = Month(oRequest("OldDate").Item)
		ElseIf Len(oRequest("OldTimeStamp").Item) > 0 Then
			aCalendarComponent(N_OLD_MONTH_CALENDAR) = Month(oRequest("OldTimeStamp").Item)
		Else
			aCalendarComponent(N_OLD_MONTH_CALENDAR) = aCalendarComponent(N_MONTH_CALENDAR)
		End If
	End If
	If IsEmpty(aCalendarComponent(S_MONTH_CALENDAR)) And (aCalendarComponent(N_OLD_MONTH_CALENDAR) > 0) Then
		aCalendarComponent(S_MONTH_CALENDAR) = asMonthNamesCalendar(aCalendarComponent(N_OLD_MONTH_CALENDAR) - 1)
	End If

	If IsEmpty(aCalendarComponent(N_OLD_DAY_CALENDAR)) Then
		If Len(oRequest("OldDay").Item) > 0 Then
			aCalendarComponent(N_OLD_DAY_CALENDAR) = CInt(oRequest("OldDay").Item)
			If aCalendarComponent(N_OLD_DAY_CALENDAR) < 1 Then
				aCalendarComponent(N_OLD_DAY_CALENDAR) = 1
			End If
		ElseIf Len(oRequest("OldDate").Item) > 0 Then
			aCalendarComponent(N_OLD_DAY_CALENDAR) = Day(oRequest("OldDate").Item)
		ElseIf Len(oRequest("OldTimeStamp").Item) > 0 Then
			aCalendarComponent(N_OLD_DAY_CALENDAR) = Day(oRequest("OldTimeStamp").Item)
		Else
			aCalendarComponent(N_OLD_DAY_CALENDAR) = aCalendarComponent(N_DAY_CALENDAR)
		End If
	End If

	If IsEmpty(aCalendarComponent(N_OLD_HOUR_CALENDAR)) Then
		If Len(oRequest("OldHour").Item) > 0 Then
			aCalendarComponent(N_OLD_HOUR_CALENDAR) = CInt(oRequest("OldHour").Item)
			If aCalendarComponent(N_OLD_HOUR_CALENDAR) < 0 Then
				aCalendarComponent(N_OLD_HOUR_CALENDAR) = -1
			ElseIf aCalendarComponent(N_OLD_HOUR_CALENDAR) > 25 Then
				aCalendarComponent(N_OLD_HOUR_CALENDAR) = -1
			End If
		ElseIf Len(oRequest("OldTime").Item) > 0 Then
			aCalendarComponent(N_OLD_HOUR_CALENDAR) = CDbl(oRequest("OldTime").Item)
			If aCalendarComponent(N_OLD_HOUR_CALENDAR) < 0 Then
				aCalendarComponent(N_OLD_HOUR_CALENDAR) = 0
			End If
		ElseIf Len(oRequest("OldTimeStamp").Item) > 0 Then
			aCalendarComponent(N_OLD_HOUR_CALENDAR) = Hour(oRequest("OldTimeStamp").Item)
		Else
			aCalendarComponent(N_OLD_HOUR_CALENDAR) = aCalendarComponent(N_HOUR_CALENDAR)
		End If
	End If

	If IsEmpty(aCalendarComponent(N_OLD_MINUTE_CALENDAR)) Then
		If Len(oRequest("OldMinute").Item) > 0 Then
			aCalendarComponent(N_OLD_MINUTE_CALENDAR) = CInt(oRequest("OldMinute").Item)
			If aCalendarComponent(N_OLD_MINUTE_CALENDAR) < 0 Then
				aCalendarComponent(N_OLD_MINUTE_CALENDAR) = 0
			ElseIf aCalendarComponent(N_OLD_MINUTE_CALENDAR) > 60 Then
				aCalendarComponent(N_OLD_MINUTE_CALENDAR) = 59
			End If
		ElseIf Len(oRequest("OldTime").Item) > 0 Then
			aCalendarComponent(N_OLD_MINUTE_CALENDAR) = CDbl(oRequest("OldTime").Item)
			If aCalendarComponent(N_OLD_MINUTE_CALENDAR) < 0 Then
				aCalendarComponent(N_OLD_MINUTE_CALENDAR) = 0
			End If
		ElseIf Len(oRequest("OldTimeStamp").Item) > 0 Then
			aCalendarComponent(N_OLD_MINUTE_CALENDAR) = Minute(oRequest("OldTimeStamp").Item)
		Else
			aCalendarComponent(N_OLD_MINUTE_CALENDAR) = aCalendarComponent(N_MINUTE_CALENDAR)
		End If
	End If

	If IsEmpty(aCalendarComponent(N_OLD_TASK_ID_CALENDAR)) Then
		If Len(oRequest("OldTaskID").Item) > 0 Then
			aCalendarComponent(N_OLD_TASK_ID_CALENDAR) = CLng(oRequest("OldTaskID").Item)
		Else
			aCalendarComponent(N_OLD_TASK_ID_CALENDAR) = aCalendarComponent(N_TASK_ID_CALENDAR)
		End If
	End If

	aCalendarComponent(B_ONLY_SUNDAY_CALENDAR) = (Len(oRequest("OnlySundays").Item) > 0)
	aCalendarComponent(B_ONLY_PAYDAYS_CALENDAR) = (Len(oRequest("OnlyPaydays").Item) > 0)
	aCalendarComponent(B_ONLY_HOLIDAYS_CALENDAR) = (Len(oRequest("OnlyHolidays").Item) > 0)
	aCalendarComponent(B_COMPONENT_INITIALIZED_CALENDAR) = True
	InitializeCalendarComponent = Err.number
	Err.Clear
End Function

Function InitializeMonth(aCalendarComponent)
'************************************************************
'Purpose: To initialize the information about the month
'Outputs: aCalendarComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeMonth"

	aCalendarComponent(S_MONTH_CALENDAR) = asMonthNamesCalendar(aCalendarComponent(N_MONTH_CALENDAR) - 1)
	aCalendarComponent(N_FIRST_DAY_CALENDAR) = Weekday(DateSerial(aCalendarComponent(N_YEAR_CALENDAR), aCalendarComponent(N_MONTH_CALENDAR), 1), vbSunday)

	If InStr(1, ",1,3,5,7,8,10,12,", ("," & aCalendarComponent(N_MONTH_CALENDAR) & ","), vbBinaryCompare) > 0 Then
		aCalendarComponent(N_DAYS_CALENDAR) = 31
	ElseIf InStr(1, ",4,6,9,11,", ("," & aCalendarComponent(N_MONTH_CALENDAR) & ","), vbBinaryCompare) > 0 Then
		aCalendarComponent(N_DAYS_CALENDAR) = 30
	ElseIf (aCalendarComponent(N_YEAR_CALENDAR) mod 4) = 0 Then
		aCalendarComponent(N_DAYS_CALENDAR) = 29
	Else
		aCalendarComponent(N_DAYS_CALENDAR) = 28
	End If

	InitializeMonth = Err.number
	Err.Clear
End Function

Function GetDaysInMonth(iYear, iMonth, iDays)
'************************************************************
'Purpose: To get the number of days for the given month
'Inputs:  iYear, iMonth
'Outputs: iDays
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetDaysInMonth"

	If InStr(1, ",1,3,5,7,8,10,12,", ("," & iMonth & ","), vbBinaryCompare) > 0 Then
		iDays = 31
	ElseIf InStr(1, ",4,6,9,11,", ("," & iMonth & ","), vbBinaryCompare) > 0 Then
		iDays = 30
	ElseIf (iYear mod 4) = 0 Then
		iDays = 29
	Else
		iDays = 28
	End If

	GetDaysInMonth = Err.number
	Err.Clear
End Function

Function GetHolidaysDates(oRequest, oADODBConnection, aCalendarComponent, sDates, sErrorDescription)
'************************************************************
'Purpose: To get the dates for all the absences for the
'         employee from the database
'Inputs:  oRequest, oADODBConnection, aCalendarComponent
'Outputs: sDates, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetHolidaysDates"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim lDaysInMonth

	bComponentInitialized = aCalendarComponent(B_COMPONENT_INITIALIZED_CALENDAR)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCalendarComponent(oRequest, aCalendarComponent)
	End If

	Call GetDaysInMonth(aCalendarComponent(N_YEAR_CALENDAR), aCalendarComponent(N_MONTH_CALENDAR), lDaysInMonth)
	sErrorDescription = "No se pudo obtener la información de los dìas festivos."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Holiday From Holidays Where (Holiday>=" & aCalendarComponent(N_YEAR_CALENDAR) & aCalendarComponent(N_MONTH_CALENDAR) & "00 And Holiday <= " & aCalendarComponent(N_YEAR_CALENDAR) & aCalendarComponent(N_MONTH_CALENDAR) & lDaysInMonth & ") Order By Holiday", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		sDates = ""
		Do While Not oRecordset.EOF
			sDates = sDates & CStr(oRecordset.Fields("Holiday").Value) & ","
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
	End If

	GetHolidaysDates = lErrorNumber
	Err.Clear
End Function

Function AddTask(oRequest, oADODBConnection, aCalendarComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new task into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aCalendarComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddTask"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCalendarComponent(B_COMPONENT_INITIALIZED_CALENDAR)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCalendarComponent(oRequest, aCalendarComponent)
	End If

	If aCalendarComponent(N_TASK_ID_CALENDAR) = -1 Then
		sErrorDescription = "No se pudo obtener un identificador para la nueva actividad."
		lErrorNumber = GetNewIDFromTable(oADODBConnection, "CalendarTasks", "TaskID", ("Where (TaskYear=" & aCalendarComponent(N_YEAR_CALENDAR) & ") And (TaskMonth= " & aCalendarComponent(N_MONTH_CALENDAR) & ") And (TaskDay=" & aCalendarComponent(N_DAY_CALENDAR) & ") And (TaskHour=" & aCalendarComponent(N_HOUR_CALENDAR) & ") And (TaskMinute=" & aCalendarComponent(N_MINUTE_CALENDAR) & ") And (UserID=" & aCalendarComponent(N_USER_ID_CALENDAR) & ") And (GroupID=" & aCalendarComponent(N_GROUP_ID_CALENDAR) & ")"), 1, aCalendarComponent(N_TASK_ID_CALENDAR), sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		lErrorNumber = CheckTaskExistency(oRequest, oADODBConnection, aCalendarComponent, sErrorDescription)
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo agregar la actividad al calendario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into CalendarTasks (UserID, GroupID, TaskYear, TaskMonth, TaskDay, TaskHour, TaskMinute, TaskID, TaskDate, TaskColor, TaskTitle, TaskDescription, TaskDuration, TypeID, TaskDone) Values (" & aCalendarComponent(N_USER_ID_CALENDAR) & ", " & aCalendarComponent(N_GROUP_ID_CALENDAR) & ", " & aCalendarComponent(N_YEAR_CALENDAR) & ", " & aCalendarComponent(N_MONTH_CALENDAR) & ", " & aCalendarComponent(N_DAY_CALENDAR) & ", " & aCalendarComponent(N_HOUR_CALENDAR) & ", " & aCalendarComponent(N_MINUTE_CALENDAR) & ", " & aCalendarComponent(N_TASK_ID_CALENDAR) & ", " & (aCalendarComponent(N_YEAR_CALENDAR) & Right(("0" & aCalendarComponent(N_MONTH_CALENDAR)), Len("00")) & Right(("0" & aCalendarComponent(N_DAY_CALENDAR)), Len("00"))) & ", '" & Replace(aCalendarComponent(S_TASK_COLOR_CALENDAR), "'", "") & "', '" & Replace(aCalendarComponent(S_TASK_TITLE_CALENDAR), "'", "") & "', '" & Replace(aCalendarComponent(S_TASK_DESCRIPTION_CALENDAR), "'", "") & "', " & aCalendarComponent(N_DURATION_CALENDAR) & ", " & aCalendarComponent(N_TASK_TYPE_CALENDAR) & ", " & aCalendarComponent(B_TASK_DONE_CALENDAR) & ")", "CalendarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	AddTask = lErrorNumber
	Err.Clear
End Function

Function GetTask(oRequest, oADODBConnection, aCalendarComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a task from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aCalendarComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetTask"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCalendarComponent(B_COMPONENT_INITIALIZED_CALENDAR)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCalendarComponent(oRequest, aCalendarComponent)
	End If

	sErrorDescription = "No se pudo obtener la actividad del calendario."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CalendarTasks.*, TypeName From CalendarTasks, CalendarTaskTypes Where (TaskYear=" & aCalendarComponent(N_OLD_YEAR_CALENDAR) & ") And (TaskMonth= " & aCalendarComponent(N_OLD_MONTH_CALENDAR) & ") And (TaskDay=" & aCalendarComponent(N_OLD_DAY_CALENDAR) & ") And (TaskHour=" & aCalendarComponent(N_OLD_HOUR_CALENDAR) & ") And (TaskMinute=" & aCalendarComponent(N_OLD_MINUTE_CALENDAR) & ") And (TaskID=" & aCalendarComponent(N_OLD_TASK_ID_CALENDAR) & ") And (UserID=" & aCalendarComponent(N_OLD_USER_ID_CALENDAR) & ") And (GroupID=" & aCalendarComponent(N_OLD_GROUP_ID_CALENDAR) & ") And (CalendarTasks.TypeID = CalendarTaskTypes.TypeID)", "CalendarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			lErrorNumber = -1
			sErrorDescription = "No existe ninguna actividad en el calendario para el " & DisplayDate(aCalendarComponent(N_YEAR_CALENDAR), aCalendarComponent(N_MONTH_CALENDAR), aCalendarComponent(N_DAY_CALENDAR), -1, -1, -1)
			If aCalendarComponent(N_HOUR_CALENDAR) < 25 Then
				sErrorDescription = sErrorDescription & " a las " & aCalendarComponent(N_HOUR_CALENDAR) & ":" & aCalendarComponent(N_MINUTE_CALENDAR) & " horas."
			Else
				sErrorDescription = sErrorDescription & "."
			End If
		Else
			aCalendarComponent(S_TASK_COLOR_CALENDAR) = CStr(oRecordset.Fields("TaskColor").Value)
			aCalendarComponent(S_TASK_TITLE_CALENDAR) = CStr(oRecordset.Fields("TaskTitle").Value)
			aCalendarComponent(S_TASK_DESCRIPTION_CALENDAR) = CStr(oRecordset.Fields("TaskDescription").Value)
			aCalendarComponent(N_DURATION_CALENDAR) = CInt(oRecordset.Fields("TaskDuration").Value)
			aCalendarComponent(N_TASK_TYPE_CALENDAR) = CLng(oRecordset.Fields("TypeID").Value)
			aCalendarComponent(S_TASK_TYPE_CALENDAR) = asTaskTypesCalendar(aCalendarComponent(N_TASK_TYPE_CALENDAR))
			aCalendarComponent(B_TASK_DONE_CALENDAR) = CInt(oRecordset.Fields("TaskDone").Value)
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	GetTask = lErrorNumber
	Err.Clear
End Function

Function GetMonth(oRequest, oADODBConnection, aCalendarComponent, sErrorDescription)
'************************************************************
'Purpose: To get the tasks registered in a month from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aCalendarComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetMonth"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCalendarComponent(B_COMPONENT_INITIALIZED_CALENDAR)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCalendarComponent(oRequest, aCalendarComponent)
	End If

	sErrorDescription = "No se pudieron obtener las actividades del calendario."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct TaskDay From CalendarTasks Where (TaskYear=" & aCalendarComponent(N_YEAR_CALENDAR) & ") And (TaskMonth= " & aCalendarComponent(N_MONTH_CALENDAR) & ") And ((UserID=" & aCalendarComponent(N_USER_ID_CALENDAR) & ") Or (UserID=-1 And GroupID=-1) Or (UserID=-1 And GroupID=" & aCalendarComponent(N_GROUP_ID_CALENDAR) & ")) Order By TaskDay", "CalendarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			aCalendarComponent(S_MARKED_DAYS_CALENDAR) = ","
			Do While Not oRecordset.EOF
				aCalendarComponent(S_MARKED_DAYS_CALENDAR) = aCalendarComponent(S_MARKED_DAYS_CALENDAR) & aCalendarComponent(N_YEAR_CALENDAR) & Right(("0" & aCalendarComponent(N_MONTH_CALENDAR)), Len("00")) & Right(("0" & oRecordset.Fields("TaskDay")), Len("00")) & ","
				oRecordset.MoveNext
				If Err.number <> 0 Then
					lErrorNumber = Err.number
					sErrorDescription = "Ocurrió un error al obtener las actividades registradas en el mes de " & asMonthNamesCalendar(aCalendarComponent(N_MONTH_CALENDAR)-1) & ".<BR />" & Err.description
					Exit Do
				End If
			Loop
			aCalendarComponent(S_MARKED_DAYS_CALENDAR) = Left(aCalendarComponent(S_MARKED_DAYS_CALENDAR), (Len(aCalendarComponent(S_MARKED_DAYS_CALENDAR)) - Len(",")))
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	GetMonth = lErrorNumber
	Err.Clear
End Function

Function OwnsTask(lUserID, lGroupID)
'************************************************************
'Purpose: Check if the user owns a specific task
'Inputs:  lUserID, lGroupID
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "OwnsTask"
	
'	OwnsTask = ((lUserID = aLoginComponent(N_USER_ID_LOGIN)) And (lGroupID = aLoginComponent(N_USER_GROUP_LOGIN))) Or _
'	(aLoginComponent(B_TUTOR_LOGIN) And (lGroupID = aLoginComponent(N_USER_GROUP_LOGIN))) Or _
'	aLoginComponent(B_ADMINISTRATOR_LOGIN)
	OwnsTask = (lUserID = aLoginComponent(N_USER_ID_LOGIN))
End Function

Function ModifyTask(oRequest, oADODBConnection, aCalendarComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing task in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aCalendarComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyTask"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCalendarComponent(B_COMPONENT_INITIALIZED_CALENDAR)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCalendarComponent(oRequest, aCalendarComponent)
	End If

	If OwnsTask(aCalendarComponent(N_USER_ID_CALENDAR), aCalendarComponent(N_GROUP_ID_CALENDAR)) Then
		sErrorDescription = "No se pudo modificar la actividad en el calendario."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update CalendarTasks Set UserID=" & aCalendarComponent(N_USER_ID_CALENDAR) & ", GroupID=" & aCalendarComponent(N_GROUP_ID_CALENDAR) & ", TaskYear=" & aCalendarComponent(N_YEAR_CALENDAR) & ", TaskMonth=" & aCalendarComponent(N_MONTH_CALENDAR) & ", TaskDay=" & aCalendarComponent(N_DAY_CALENDAR) & ", TaskHour=" & aCalendarComponent(N_HOUR_CALENDAR) & ", TaskMinute=" & aCalendarComponent(N_MINUTE_CALENDAR) & ", TaskID=" & aCalendarComponent(N_TASK_ID_CALENDAR) & ", TaskColor='" & Replace(aCalendarComponent(S_TASK_COLOR_CALENDAR), "'", "") & "', TaskTitle='" & Replace(aCalendarComponent(S_TASK_TITLE_CALENDAR), "'", "") & "', TaskDescription='" & Replace(aCalendarComponent(S_TASK_DESCRIPTION_CALENDAR), "'", "") & "', TaskDuration=" & aCalendarComponent(N_DURATION_CALENDAR) & ", TypeID=" & aCalendarComponent(N_TASK_TYPE_CALENDAR) & ", TaskDone=" & aCalendarComponent(B_TASK_DONE_CALENDAR) & " Where (TaskYear=" & aCalendarComponent(N_OLD_YEAR_CALENDAR) & ") And (TaskMonth=" & aCalendarComponent(N_OLD_MONTH_CALENDAR) & ") And (TaskDay=" & aCalendarComponent(N_OLD_DAY_CALENDAR) & ") And (TaskHour=" & aCalendarComponent(N_OLD_HOUR_CALENDAR) & ") And (TaskMinute=" & aCalendarComponent(N_OLD_MINUTE_CALENDAR) & ") And (TaskID=" & aCalendarComponent(N_OLD_TASK_ID_CALENDAR) & ") And (UserID=" & aCalendarComponent(N_OLD_USER_ID_CALENDAR) & ") And (GroupID=" & aCalendarComponent(N_OLD_GROUP_ID_CALENDAR) & ")", "CalendarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	ModifyTask = lErrorNumber
	Err.Clear
End Function

Function RemoveTask(oRequest, oADODBConnection, aCalendarComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a task from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aCalendarComponent, aCalendarComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveTask"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCalendarComponent(B_COMPONENT_INITIALIZED_CALENDAR)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCalendarComponent(oRequest, aCalendarComponent)
	End If

	If OwnsTask(aCalendarComponent(N_USER_ID_CALENDAR), aCalendarComponent(N_GROUP_ID_CALENDAR)) Then
		sErrorDescription = "No se pudo borrar la actividad del calendario."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From CalendarTasks Where (TaskYear=" & aCalendarComponent(N_YEAR_CALENDAR) & ") And (TaskMonth= " & aCalendarComponent(N_MONTH_CALENDAR) & ") And (TaskDay=" & aCalendarComponent(N_DAY_CALENDAR) & ") And (TaskHour=" & aCalendarComponent(N_HOUR_CALENDAR) & ") And (TaskMinute=" & aCalendarComponent(N_MINUTE_CALENDAR) & ") And (TaskID=" & aCalendarComponent(N_TASK_ID_CALENDAR) & ") And (UserID=" & aCalendarComponent(N_USER_ID_CALENDAR) & ") And (GroupID=" & aCalendarComponent(N_GROUP_ID_CALENDAR) & ")", "CalendarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	Else
		lErrorNumber = -1
		sErrorDescription = "Usted no cuenta con los permisos necesarios para borrar la actividad seleccionada."
	End If

	RemoveGroup = lErrorNumber
	Err.Clear
End Function

Function CheckTaskExistency(oRequest, oADODBConnection, aCalendarComponent, sErrorDescription)
'************************************************************
'Purpose: To check the existency of a task  in the database
'Inputs:  oRequest, oADODBConnection, aCalendarComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckTaskExistency"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCalendarComponent(B_COMPONENT_INITIALIZED_CALENDAR)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCalendarComponent(oRequest, aCalendarComponent)
	End If

	sErrorDescription = "No se pudo verificar la existencia de la actividad del calendario."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From CalendarTasks Where (TaskYear=" & aCalendarComponent(N_YEAR_CALENDAR) & ") And (TaskMonth= " & aCalendarComponent(N_MONTH_CALENDAR) & ") And (TaskDay=" & aCalendarComponent(N_DAY_CALENDAR) & ") And (TaskHour=" & aCalendarComponent(N_HOUR_CALENDAR) & ") And (TaskMinute=" & aCalendarComponent(N_MINUTE_CALENDAR) & ") And (TaskID=" & aCalendarComponent(N_TASK_ID_CALENDAR) & ") And (UserID=" & aCalendarComponent(N_USER_ID_CALENDAR) & ") And (GroupID=" & aCalendarComponent(N_GROUP_ID_CALENDAR) & ")", "CalendarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			lErrorNumber = -1
			sErrorDescription = "Ya existe una actividad en el calendario para el " & DisplayDate(aCalendarComponent(N_YEAR_CALENDAR), aCalendarComponent(N_MONTH_CALENDAR), aCalendarComponent(N_DAY_CALENDAR), aCalendarComponent(N_HOUR_CALENDAR), aCalendarComponent(N_MINUTE_CALENDAR), -1) & " horas."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	CheckTaskExistency = lErrorNumber
	Err.Clear
End Function

Function IsNewDateEqualsToOldDate(oRequest, aCalendarComponent, sErrorDescription)
'************************************************************
'Purpose: To check the existency of a task  in the database
'Inputs:  oRequest, aCalendarComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "IsNewDateEqualsToOldDate"
	Dim bComponentInitialized

	bComponentInitialized = aCalendarComponent(B_COMPONENT_INITIALIZED_CALENDAR)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCalendarComponent(oRequest, aCalendarComponent)
	End If

	IsNewDateEqualsToOldDate = (aCalendarComponent(N_USER_ID_CALENDAR) = aCalendarComponent(N_OLD_USER_ID_CALENDAR)) Or _
	(aCalendarComponent(N_GROUP_ID_CALENDAR) = aCalendarComponent(N_OLD_GROUP_ID_CALENDAR)) Or _
	(aCalendarComponent(N_YEAR_CALENDAR) = aCalendarComponent(N_OLD_YEAR_CALENDAR)) Or _
	(aCalendarComponent(N_MONTH_CALENDAR) = aCalendarComponent(N_OLD_MONTH_CALENDAR)) Or _
	(aCalendarComponent(N_DAY_CALENDAR) = aCalendarComponent(N_OLD_DAY_CALENDAR)) Or _
	(aCalendarComponent(N_HOUR_CALENDAR) = aCalendarComponent(N_OLD_HOUR_CALENDAR)) Or _
	(aCalendarComponent(N_MINUTE_CALENDAR) = aCalendarComponent(N_OLD_MINUTE_CALENDAR)) Or _
	(aCalendarComponent(N_TASK_ID_CALENDAR) = aCalendarComponent(N_OLD_TASK_ID_CALENDAR))

	Err.Clear
End Function

Function DateIsHoliday(oRequest, oADODBConnection, aCalendarComponent, iDay, sErrorDescription)
'************************************************************
'Purpose: To get the dates for all the absences for the
'         employee from the database
'Inputs:  oRequest, oADODBConnection, aCalendarComponent, iDay
'Outputs: sDates, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetHolidaysDates"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery

	bComponentInitialized = aCalendarComponent(B_COMPONENT_INITIALIZED_CALENDAR)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCalendarComponent(oRequest, aCalendarComponent)
	End If


	sQuery = "Select * From Holidays" & _
			 " Where (Holiday=" & aCalendarComponent(N_YEAR_CALENDAR) & Right(("0" & aCalendarComponent(N_MONTH_CALENDAR)), Len("00")) & Right(("0" & iDay), Len("00")) & ")"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			DateIsHoliday = True
		Else
			DateIsHoliday = False
		End If
	Else
		sErrorDescription = "Error al verificar si el día es festivo."
		DateIsHoliday = False
	End If

	Err.Clear
End Function

Function DisplayTask(oRequest, aCalendarComponent, sErrorDescription)
'************************************************************
'Purpose: To display a task using the information from the
'         component
'Inputs:  oRequest, aCalendarComponent
'Outputs: aCalendarComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayTask"
	Dim iIndex
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCalendarComponent(B_COMPONENT_INITIALIZED_CALENDAR)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCalendarComponent(oRequest, aCalendarComponent)
	End If

'	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserID"" ID=""UserIDHdn"" VALUE=""" & aCalendarComponent(N_USER_ID_CALENDAR) & """/>"
'	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""GroupID"" ID=""GroupIDHdn"" VALUE=""" & aCalendarComponent(N_GROUP_ID_CALENDAR) & """/>"
	Response.Write "<TABLE BGCOLOR=""#CCCCCC"" WIDTH=""300"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
		Response.Write "<TR>"
			Response.Write "<TD WIDTH=""16"" COLSPAN=""2""><IMG SRC=""Images/TopLeftCorner.gif"" WIDTH=""16"" HEIGHT=""16"" /></TD>"
			Response.Write "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"" COLOR=""#" & asTaskColorsCalendar(aCalendarComponent(S_TASK_COLOR_CALENDAR)) & """><NOBR>&nbsp;<B>" & CleanStringForHTML(aCalendarComponent(S_TASK_TITLE_CALENDAR)) & "</B>&nbsp;</NOBR></FONT></TD>"
			Response.Write "<TD WIDTH=""16"" ALIGN=""RIGHT"" COLSPAN=""2""><IMG SRC=""Images/TopRightCorner.gif"" WIDTH=""16"" HEIGHT=""16"" /></TD>"
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD WIDTH=""1""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
			Response.Write "<TD BGCOLOR=""#FFFFFF"" WIDTH=""15""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
			Response.Write "<TD BGCOLOR=""#FFFFFF""><FONT FACE=""Arial"" SIZE=""2"" COLOR=""#" & asTaskColorsCalendar(aCalendarComponent(S_TASK_COLOR_CALENDAR)) & """><BR />"
				Response.Write "<B>Tipo:&nbsp;</B>" & aCalendarComponent(S_TASK_TYPE_CALENDAR) & "<BR />"
				Response.Write "<B>Fecha:&nbsp;</B>" & DisplayDate(aCalendarComponent(N_YEAR_CALENDAR), aCalendarComponent(N_MONTH_CALENDAR), aCalendarComponent(N_DAY_CALENDAR), -1, -1, -1) & "<BR />"
				If aCalendarComponent(N_HOUR_CALENDAR) = 25 Then
					Response.Write "Actividad para todo el día<BR />"
				Else
					Response.Write "<B>Hora de Inicio:&nbsp;</B>" & aCalendarComponent(N_HOUR_CALENDAR) & ":" & Right(("0" & aCalendarComponent(N_MINUTE_CALENDAR)), Len("00")) & "<BR />"
					Response.Write "<B>Duración:&nbsp;</B>" & aCalendarComponent(N_DURATION_CALENDAR) & "<BR />"
				End If
				If aCalendarComponent(B_TASK_DONE_CALENDAR) Then
					Response.Write "Actividad realizada"
				End If
			Response.Write "&nbsp;<BR /></FONT>"
				Response.Write "<FORM><TEXTAREA NAME=""TaskDescription"" ID=""TaskDescriptionTxtArea"" ROWS=""4"" COLS=""30"" CLASS=""TextFields"" STYLE=""color: " &  asTaskColorsCalendar(aCalendarComponent(S_TASK_COLOR_CALENDAR)) & """>" & aCalendarComponent(S_TASK_DESCRIPTION_CALENDAR) & "</TEXTAREA></FORM>"
			Response.Write "</FONT></TD>"
			Response.Write "<TD BGCOLOR=""#FFFFFF"" WIDTH=""15""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
			Response.Write "<TD WIDTH=""1""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD WIDTH=""1"" COLSPAN=""5""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
		Response.Write "</TR>"
	Response.Write "</TABLE><BR />"

	DisplayTask = lErrorNumber
	Err.Clear
End Function

Function DisplayTaskAsForm(oRequest, sAction, aCalendarComponent, sErrorDescription)
'************************************************************
'Purpose: To display a task as a form using the information
'         from the component
'Inputs:  oRequest, sAction, aCalendarComponent
'Outputs: aCalendarComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayTaskAsForm"
	Dim iIndex

	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		Response.Write "function CheckTaskFields(oForm) {" & vbNewLine
			Response.Write "if (oForm) {" & vbNewLine
				Response.Write "if (IsDisplayed('RemoveWng')) {" & vbNewLine
					Response.Write "return true;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (oForm.TaskTitle.value.length == 0) {" & vbNewLine
					Response.Write "alert('Favor de introducir el título de la actividad.');" & vbNewLine
					Response.Write "oForm.TaskTitle.focus();" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "}" & vbNewLine
			Response.Write "return true;" & vbNewLine
		Response.Write "} // End of CheckTaskFields" & vbNewLine
	Response.Write "//--></SCRIPT>" & vbNewLine
	Response.Write "<FORM NAME=""TaskFrm"" ID=""TaskFrm"" ACTION=""" & sAction & """ METHOD=""POST""  onSubmit=""return CheckTaskFields(this)"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserID"" ID=""UserIDHdn"" VALUE=""" & aCalendarComponent(N_USER_ID_CALENDAR) & """/>"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""GroupID"" ID=""GroupIDHdn"" VALUE=""" & aCalendarComponent(N_GROUP_ID_CALENDAR) & """/>"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaskID"" ID=""TaskIDHdn"" VALUE=""" & aCalendarComponent(N_TASK_ID_CALENDAR) & """/>"

		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OldYear"" ID=""OldYearHdn"" VALUE=""" & aCalendarComponent(N_OLD_YEAR_CALENDAR) & """/>"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OldMonth"" ID=""OldMonthHdn"" VALUE=""" & aCalendarComponent(N_OLD_MONTH_CALENDAR) & """/>"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OldDay"" ID=""OldDayHdn"" VALUE=""" & aCalendarComponent(N_OLD_DAY_CALENDAR) & """/>"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OldHour"" ID=""OldHourHdn"" VALUE=""" & aCalendarComponent(N_OLD_HOUR_CALENDAR) & """/>"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OldMinute"" ID=""OldMinuteHdn"" VALUE=""" & aCalendarComponent(N_OLD_MINUTE_CALENDAR) & """/>"

		Response.Write "<TABLE WIDTH=""400"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD WIDTH=""100""><FONT FACE=""Arial"" SIZE=""2"">Título:&nbsp;</FONT></TD>"
				Response.Write "<TD WIDTH=""300""><INPUT TYPE=""TEXT"" NAME=""TaskTitle"" ID=""TaskTitleTxt"" SIZE=""34"" MAXLENGTH=""255"" VALUE=""" & Replace(aCalendarComponent(S_TASK_TITLE_CALENDAR), """", "&#34;") & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
If False Then
			Response.Write "<TR>"
				Response.Write "<TD WIDTH=""100""><FONT FACE=""Arial"" SIZE=""2"">Tipo:&nbsp;</FONT></TD>"
				Response.Write "<TD WIDTH=""300"">"
					Response.Write "<SELECT NAME=""TypeID"" ID=""TypeIDCmb"" CLASS=""Lists"">"
						For iIndex = 0 To UBound(asTaskTypesCalendar)
							Response.Write "<OPTION VALUE=""" & iIndex & """"
							If iIndex = aCalendarComponent(N_TASK_TYPE_CALENDAR) Then
								Response.Write " SELECTED=""1"""
							End If
							Response.Write ">" & asTaskTypesCalendar(iIndex) & "</OPTION>"
						Next
					Response.Write "</SELECT>"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">&nbsp;&nbsp;&nbsp;&nbsp;Color:&nbsp;</FONT>"
					Response.Write "<SELECT NAME=""TaskColor"" ID=""TaskColorCmb"" CLASS=""Lists"" onChange=""DisplayColor(this)"">"
						For iIndex = 0 To UBound(asTaskColorNamesCalendar)
							Response.Write "<OPTION VALUE=""" & asTaskColorsCalendar(iIndex) & """"
							If StrComp(aCalendarComponent(S_TASK_COLOR_CALENDAR), asTaskColorNamesCalendar(iIndex), vbBinaryCompare) = 0 Then
								Response.Write " SELECTED=""1"""
							End If
							Response.Write ">" & asTaskColorNamesCalendar(iIndex) & "</OPTION>"
						Next
					Response.Write "</SELECT>"
				Response.Write "</TD>"
			Response.Write "</TR>"
Else
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TypeID"" ID=""TypeIDHdn"" VALUE=""" & aCalendarComponent(N_TASK_TYPE_CALENDAR) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaskColor"" ID=""TaskColorHdn"" VALUE=""" & aCalendarComponent(S_TASK_COLOR_CALENDAR) & """ />"
End If
			Response.Write "<TR>"
				Response.Write "<TD WIDTH=""100""><FONT FACE=""Arial"" SIZE=""2"">Fecha:&nbsp;</FONT></TD>"
				Response.Write "<TD WIDTH=""300"">"
					Response.Write DisplayDateCombos(aCalendarComponent(N_YEAR_CALENDAR), aCalendarComponent(N_MONTH_CALENDAR), aCalendarComponent(N_DAY_CALENDAR), "Year", "Month", "Day", (Year(Date()) - 1), (Year(Date()) + 10), True, False)
				Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD WIDTH=""100""><FONT FACE=""Arial"" SIZE=""2"">Hora de Inicio:&nbsp;</FONT></TD>"
				Response.Write "<TD WIDTH=""300"">"
					Response.Write "<SELECT NAME=""Hour"" ID=""HourCmb"" CLASS=""Lists"""
						If aCalendarComponent(N_HOUR_CALENDAR) = 25 Then
							Response.Write "DISABLED=""1"" "
						End If
					Response.Write">"
						For iIndex = 0 to 23
							Response.Write "<OPTION VALUE=""" & iIndex & """"
								If aCalendarComponent(N_HOUR_CALENDAR) = iIndex Then
									Response.Write " SELECTED=""1"""
								End If
							Response.Write ">" & iIndex & "</OPTION>"
						Next
					Response.Write "</SELECT><FONT FACE=""Arial"" SIZE=""2""> : </FONT>"
					Response.Write "<SELECT NAME=""Minute"" ID=""MinuteCmb"" CLASS=""Lists"""
						If aCalendarComponent(N_HOUR_CALENDAR) = 25 Then
							Response.Write "DISABLED=""1"" "
						End If
					Response.Write ">"
						For iIndex = 0 to 59 Step 5
							Response.Write "<OPTION VALUE=""" & iIndex & """"
								If aCalendarComponent(N_MINUTE_CALENDAR) = iIndex Then
									Response.Write " SELECTED=""1"""
								End If
							Response.Write ">" & Right(("0" & iIndex), Len("00")) & "</OPTION>"
						Next
					Response.Write "</SELECT>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				Response.Write "<INPUT TYPE=""CHECKBOX"" NAME=""AllDayActivity"" ID=""AllDayActivityChk"" VALUE=""1"""
					If aCalendarComponent(N_HOUR_CALENDAR) = 25 Then
						Response.Write "CHECKED=""1"" "
					End If
				Response.Write " onClick=""this.form.Hour.disabled=this.checked; this.form.Minute.disabled=this.checked;"" />&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Actividad para todo el día</FONT></TD>"
			Response.Write "</TR>"
If False Then
			Response.Write "<TR>"
				Response.Write "<TD WIDTH=""100""><FONT FACE=""Arial"" SIZE=""2"">Duración:&nbsp;</FONT></TD>"
				Response.Write "<TD WIDTH=""300""><INPUT TYPE=""TEXT"" NAME=""TaskDuration"" SIZE=""6"" MAXLENGTH=""6"" VALUE=""" & aCalendarComponent(N_DURATION_CALENDAR) & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
Else
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaskDuration"" ID=""TaskDurationHdn"" VALUE=""" & aCalendarComponent(N_DURATION_CALENDAR) & """ />"
End If
			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2""><BR />Comentarios:<BR /></FONT><TEXTAREA NAME=""TaskDescription"" ID=""TaskDescriptionTxtArea"" ROWS=""4"" COLS=""45"" MAXLENGTH=""1000"" CLASS=""TextFields"">" & aCalendarComponent(S_TASK_DESCRIPTION_CALENDAR) & "</TEXTAREA></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2""><INPUT TYPE=""CHECKBOX"" NAME=""TaskDone"" ID=""TaskDoneChk"" VALUE=""1"""
				If aCalendarComponent(B_TASK_DONE_CALENDAR) Then
					Response.Write " CHECKED=""1"""
				End If
				Response.Write " /> Actividad realizada</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""2""><BR />"
					Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" STYLE=""display: inline"" />"
					Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" STYLE=""display: none"" />"
					Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" STYLE=""display: none"" onClick=""ShowDisplay(document.all['RemoveCountryWngDiv']); TaskFrm.Remove.focus()"" />"
					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
					Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE="" Cancel "" CLASS=""Buttons"" onClick=""window.document.body.focus(); HideDisplay(document.TaskFrm.Add); HideDisplay(document.TaskFrm.Modify); HideDisplay(document.TaskFrm.RemoveWng); HidePopupItem('TaskFormDiv', document.TaskFormDiv)"" />"
					Response.Write "<BR /><BR />"
					Call DisplayWarningDiv("RemoveCountryWngDiv", "¿Está seguro que desea borrar la actividad de la &nbsp;base de datos?")
				Response.Write "</TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
If False Then
		Response.Write "<DIV ID=""ColorTableDiv"" STYLE=""position: absolute; "
			If bIsNetscape Then
				Response.Write "top: 88px;left: 500px"
			Else
				Response.Write "top: 83px;left: 500px"
			End If
		Response.Write """><TABLE WIDTH=""20"" BGCOLOR=""#" & aCalendarComponent(S_TASK_COLOR_CALENDAR) & """ CELLPADDING=""0"" CELLSPACING=""0""><TR><TD><IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""20"" /></TD></TR></TABLE></DIV>"
End If
	Response.Write "</FORM>"

	DisplayTaskAsForm = Err.number
	Err.Clear
End Function

Function DisplayTaskIDAsHiddenFields(oRequest, aCalendarComponent, sErrorDescription)
'************************************************************
'Purpose: To display a task as a form using the information
'         from the component
'Inputs:  oRequest, aCalendarComponent
'Outputs: aCalendarComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayTaskIDAsHiddenFields"

	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserID"" ID=""UserIDHdn"" VALUE=""" & aCalendarComponent(N_USER_ID_CALENDAR) & """/>"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""GroupID"" ID=""GroupIDHdn"" VALUE=""" & aCalendarComponent(N_GROUP_ID_CALENDAR) & """/>"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Year"" ID=""YearHdn"" VALUE=""" & aCalendarComponent(N_YEAR_CALENDAR) & """/>"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Month"" ID=""MonthHdn"" VALUE=""" & aCalendarComponent(N_MONTH_CALENDAR) & """/>"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Day"" ID=""DayHdn"" VALUE=""" & aCalendarComponent(N_DAY_CALENDAR) & """/>"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Hour"" ID=""HourHdn"" VALUE=""" & aCalendarComponent(N_HOUR_CALENDAR) & """/>"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Minute"" ID=""MinuteHdn"" VALUE=""" & aCalendarComponent(N_MINUTE_CALENDAR) & """/>"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaskID"" ID=""TaskIDHdn"" VALUE=""" & aCalendarComponent(N_TASK_ID_CALENDAR) & """/>"

	DisplayTaskIDAsHiddenFields = Err.number
	Err.Clear
End Function

Function DisplayDay(oRequest, oADODBConnection, aCalendarComponent, sErrorDescription)
'************************************************************
'Purpose: To display a day using the information from the
'         component
'Inputs:  oRequest, oADODBConnection, bAddRadioButtons, aCalendarComponent
'Outputs: aCalendarComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayDay"
	Dim sTemp
	Dim sURL
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sBoldBegin
	Dim sBoldEnd
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCalendarComponent(B_COMPONENT_INITIALIZED_CALENDAR)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCalendarComponent(oRequest, aCalendarComponent)
	End If

	Response.Write "<TABLE WIDTH=""542"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"" ALIGN=""CENTER"">" & vbNewLine
		Response.Write "<TR>"
			Response.Write "<TD WIDTH=""1"" BGCOLOR=""#" & S_MAIN_COLOR_FOR_GUI & """ ROWSPAN=""3""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
			Response.Write "<TD WIDTH=""540"" BGCOLOR=""#" & S_MAIN_COLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
			Response.Write "<TD WIDTH=""1"" BGCOLOR=""#" & S_MAIN_COLOR_FOR_GUI & """ ROWSPAN=""3""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
		Response.Write "</TR>" & vbNewLine
		Response.Write "<TR>"
			Response.Write "<TD ALIGN=""CENTER"" VALIGN=""MIDDLE"">&nbsp;<FONT FACE=""Arial"" SIZE=""2"" COLOR=""#" & S_MAIN_COLOR_FOR_GUI & """><B>" & vbNewLine
				Response.Write aCalendarComponent(S_DAY_CALENDAR) & " " & DisplayDate(aCalendarComponent(N_YEAR_CALENDAR), aCalendarComponent(N_MONTH_CALENDAR), aCalendarComponent(N_DAY_CALENDAR), -1, -1, -1) & vbNewLine
			Response.Write "</B></FONT></TD>"
		Response.Write "</TR>" & vbNewLine
		Response.Write "<TR><TD BGCOLOR=""#" & S_MAIN_COLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR>" & vbNewLine
		Response.Write "<TR><TD><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""3"" /></TD></TR>" & vbNewLine
	Response.Write "</TABLE>" & vbNewLine

	sErrorDescription = "No se pudieron obtener las actividades."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From CalendarTasks Where (TaskYear=" & aCalendarComponent(N_YEAR_CALENDAR) & ") And (TaskMonth= " & aCalendarComponent(N_MONTH_CALENDAR) & ") And (TaskDay= " & aCalendarComponent(N_DAY_CALENDAR) & ") And ((UserID=" & aCalendarComponent(N_USER_ID_CALENDAR) & ") Or (UserID=-1 And GroupID=-1) Or (UserID=-1 And GroupID=" & aCalendarComponent(N_GROUP_ID_CALENDAR) & ")) Order By TaskHour, TaskMinute, TaskID", "CalendarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			Response.Write "<FONT FACE=""Arial"" SIZE=""2""><BR /><BR />No hay actividades para este día.<BR /><BR /><BR /></FONT>"
		Else
			Response.Write "<TABLE WIDTH=""470"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bUseLinks Then
					asColumnsTitles = Split("Hora,Actividad,Completa,Acciones", ",", -1, vbBinaryCompare)
					asCellWidths = Split("60,270,70,70", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("&nbsp;,Hora,Actividad,Completa", ",", -1, vbBinaryCompare)
					asCellWidths = Split("60,340,70", ",", -1, vbBinaryCompare)
				End If
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If

				asCellAlignments = Split("RIGHT,,CENTER,CENTER", ",", -1, vbBinaryCompare)
				aCalendarComponent(N_TASK_COUNT_CALENDAR) = 0
				sTemp = ""
				Do While Not oRecordset.EOF
					sRowContents = ""
					If StrComp(sTemp, (CStr(oRecordset.Fields("TaskHour").Value) & "." & CStr(oRecordset.Fields("TaskMinute").Value)), vbBinaryCompare) <> 0 Then
						sRowContents = sRowContents & "<FONT FACE=""Arial"" SIZE=""2"">"
							If CInt(oRecordset.Fields("TaskHour").Value) = 25 Then
								sRowContents = sRowContents & "--:--"
							Else
								sRowContents = sRowContents & CStr(oRecordset.Fields("TaskHour").Value) & ":" & Right(("0" & CStr(oRecordset.Fields("TaskMinute").Value)), Len("00"))
							End If
						sRowContents = sRowContents & "</FONT>"
					End If
					sTemp = CStr(oRecordset.Fields("TaskHour").Value) & "." & CStr(oRecordset.Fields("TaskMinute").Value)

					sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT FACE=""Arial"" SIZE=""2"" COLOR=""#" & CStr(oRecordset.Fields("TaskColor").Value) & """><A"
						If Len(aCalendarComponent(S_JAVASCRIPT_CALENDAR)) > 0 Then sRowContents = sRowContents & " HREF=""javascript: " & aCalendarComponent(S_JAVASCRIPT_CALENDAR) & "(" & CStr(oRecordset.Fields("UserID").Value) & ", " & CStr(oRecordset.Fields("GroupID").Value) & ", " & aCalendarComponent(N_YEAR_CALENDAR) & ", " & aCalendarComponent(N_MONTH_CALENDAR) & ", " & aCalendarComponent(N_DAY_CALENDAR) & ", " & CStr(oRecordset.Fields("TaskHour").Value) & ", " & CStr(oRecordset.Fields("TaskMinute").Value) & ", " & CStr(oRecordset.Fields("TaskID").Value) & ")"" TARGET=""" & aCalendarComponent(S_TARGET_FRAME_CALENDAR) & """"
					sRowContents = sRowContents & "><FONT COLOR=""#" & CStr(oRecordset.Fields("TaskColor").Value) & """>" & CleanStringForHTML(CStr(oRecordset.Fields("TaskTitle").Value)) & "</FONT></A>"
					'sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;(" & asTaskTypesCalendar(CInt(oRecordset.Fields("TypeID").Value)) & ")"
					sRowContents = sRowContents & "</FONT>"
					sRowContents = sRowContents & TABLE_SEPARATOR & "<IMG SRC=""Images/"
					If CInt(oRecordset.Fields("TaskDone").Value) = 1 Then
						sRowContents = sRowContents & "BtnCheck"
					Else
						sRowContents = sRowContents & "Transparent"
					End If
					sRowContents = sRowContents & ".gif"" WIDTH=""10"" HEIGHT=""8"" />&nbsp;"
					sRowContents = sRowContents & TABLE_SEPARATOR
					If OwnsTask(CLng(oRecordset.Fields("UserID").Value), CLng(oRecordset.Fields("GroupID").Value)) Then
						sURL = "UserID=" & CStr(oRecordset.Fields("UserID").Value) & "&GroupID=" & CStr(oRecordset.Fields("GroupID").Value) & "&Year=" & CStr(oRecordset.Fields("TaskYear").Value) & "&Month=" & Right(("0" & CStr(oRecordset.Fields("TaskMonth").Value)), Len("00")) & "&Day=" & Right(("0" & CStr(oRecordset.Fields("TaskDay").Value)), Len("00")) & "&Hour=" & CStr(oRecordset.Fields("TaskHour").Value) & "&Minute=" & CStr(oRecordset.Fields("TaskMinute").Value) & "&TaskID=" & CStr(oRecordset.Fields("TaskID").Value) & "&TaskColor=" & CStr(oRecordset.Fields("TaskColor").Value) & "&TaskTitle=" & CStr(oRecordset.Fields("TaskTitle").Value) & "&TaskDescription=" & CStr(oRecordset.Fields("TaskDescription").Value) & "&TaskDuration=" & CStr(oRecordset.Fields("TaskDuration").Value) & "&TypeID=" & CStr(oRecordset.Fields("TypeID").Value) & "&TaskDone=" & CStr(oRecordset.Fields("TaskDone").Value) & "&OldYear=" & CStr(oRecordset.Fields("TaskYear").Value) & "&OldMonth=" & CStr(oRecordset.Fields("TaskMonth").Value) & "&OldDay=" & CStr(oRecordset.Fields("TaskDay").Value) & "&OldHour=" & CStr(oRecordset.Fields("TaskHour").Value) & "&OldMinute=" & CStr(oRecordset.Fields("TaskMinute").Value) & "&AllDayActivity="
						If CInt(oRecordset.Fields("TaskHour").Value) = 25 Then
							sURL = sURL & "1"
						Else
							sURL = sURL & "0"
						End If
						If True Then'aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then
							sRowContents = sRowContents & "<A HREF=""javascript: HideDisplay(document.TaskFrm.Add); ShowDisplay(document.TaskFrm.Modify); HideDisplay(document.TaskFrm.RemoveWng); SendURLValuesToForm('" & sURL & "', document.TaskFrm); document.TaskFrm.Hour.disabled=document.TaskFrm.AllDayActivity.checked; document.TaskFrm.Minute.disabled=document.TaskFrm.AllDayActivity.checked; ShowPopupItem('TaskFormDiv', document.TaskFormDiv); document.TaskFrm.TaskTitle.focus();"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
						End If
						If B_DELETE_CATALOGS Then
							sRowContents = sRowContents & "<A HREF=""javascript: HideDisplay(document.TaskFrm.Add); HideDisplay(document.TaskFrm.Modify); ShowDisplay(document.TaskFrm.RemoveWng); SendURLValuesToForm('" & sURL & "', document.TaskFrm); document.TaskFrm.Hour.disabled=document.TaskFrm.AllDayActivity.checked; document.TaskFrm.Minute.disabled=document.TaskFrm.AllDayActivity.checked; ShowPopupItem('TaskFormDiv', document.TaskFormDiv); document.TaskFrm.TaskTitle.focus();"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>"
						End If
					End If

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
					bEven = Not bEven
					aCalendarComponent(N_TASK_COUNT_CALENDAR) = aCalendarComponent(N_TASK_COUNT_CALENDAR) + 1
				Loop
			Response.Write "</TABLE>" & vbNewLine
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayDay = lErrorNumber
	Err.Clear
End Function

Function DisplayWeek(oRequest, oADODBConnection, aCalendarComponent, sErrorDescription)
'************************************************************
'Purpose: To display a week using the information from the
'         component
'Inputs:  oRequest, oADODBConnection, bAddRadioButtons, aCalendarComponent
'Outputs: aCalendarComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayWeek"
	Dim bEven
	Dim sBGColor
	Dim sTemp
	Dim iIndex
	Dim oTempDate
	Dim iDaysInMonth
	Dim iStartDay
	Dim iStartMonth
	Dim iStartYear
	Dim iEndDay
	Dim iEndMonth
	Dim iEndYear
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCalendarComponent(B_COMPONENT_INITIALIZED_CALENDAR)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCalendarComponent(oRequest, aCalendarComponent)
	End If

	Call GetDaysInMonth(iStartYear, iStartMonth, iDaysInMonth)
	oTempDate = DateSerial(aCalendarComponent(N_YEAR_CALENDAR), aCalendarComponent(N_MONTH_CALENDAR), aCalendarComponent(N_DAY_CALENDAR))
	oTempDate = DateAdd("d", -(Weekday(oTempDate, vbSunday) - 1), oTempDate)
	iStartDay = Day(oTempDate)
	iStartMonth = Month(oTempDate)
	iStartYear = Year(oTempDate)
	oTempDate = DateAdd("d", 6, oTempDate)
	iEndDay = Day(oTempDate)
	iEndMonth = Month(oTempDate)
	iEndYear = Year(oTempDate)
	oTempDate = DateAdd("d", -6, oTempDate)
	Response.Write "<TABLE WIDTH=""600"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
		Response.Write "<TR>"
			Response.Write "<TD WIDTH=""1"" BGCOLOR=""#000000"" ROWSPAN=""3""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
			Response.Write "<TD WIDTH=""598"" BGCOLOR=""#000000""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
			Response.Write "<TD WIDTH=""1"" BGCOLOR=""#000000"" ROWSPAN=""3""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
		Response.Write "</TR>" & vbNewLine
		Response.Write "<TR>"
			Response.Write "<TD ALIGN=""CENTER"" VALIGN=""MIDDLE"">&nbsp;<FONT FACE=""Arial"" SIZE=""2""><B>" & vbNewLine
				Response.Write "Semana del " & DisplayDate(iStartYear, iStartMonth, iStartDay, -1, -1, -1) & " al " & DisplayDate(iEndYear, iEndMonth, iEndDay, -1, -1, -1) & vbNewLine
			Response.Write "</B></FONT></TD>"
		Response.Write "</TR>" & vbNewLine
		Response.Write "<TR><TD BGCOLOR=""#000000""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR>" & vbNewLine
		Response.Write "<TR><TD><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""3"" /></TD></TR>" & vbNewLine
	Response.Write "</TABLE>" & vbNewLine
	Response.Write "<TABLE WIDTH=""600"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
		Response.Write "<TR>"
			For iIndex = 0 To 6
				iStartDay = Day(oTempDate)
				iStartMonth = Month(oTempDate)
				iStartYear = Year(oTempDate)
				Response.Write "<TD VALIGN=""TOP"">" & vbNewLine
					sErrorDescription = "No se pudieron obtener las actividades del calendario."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From CalendarTasks Where (TaskYear=" & iStartYear & ") And (TaskMonth= " & iStartMonth & ") And (TaskDay= " & iStartDay & ") And ((UserID=" & aCalendarComponent(N_USER_ID_CALENDAR) & ") Or (UserID=-1 And GroupID=-1) Or (UserID=-1 And GroupID=" & aCalendarComponent(N_GROUP_ID_CALENDAR) & ")) Order By TaskHour, TaskMinute, TaskID", "CalendarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						bEven = False
						aCalendarComponent(N_TASK_COUNT_CALENDAR) = 0
						sTemp = ""
						Response.Write "<TABLE WIDTH=""100%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
								Response.Write "<TR BGCOLOR=""#CCCCCC"">"
									Response.Write "<TD ALIGN=""CENTER"" COLSPAN=""3"">&nbsp;<FONT FACE=""Arial"" SIZE=""2""><B>" & asDayNamesCalendar(iIndex) & "</B></FONT></TD>"
								Response.Write "</TR>"
							Do While Not oRecordset.EOF
								If bEven Then
									sBGColor = "DDDDDD"
								Else
									sBGColor = "EEEEEE"
								End If
								Response.Write "<TR>"
									Response.Write "<TD BGCOLOR=""#" & sBGColor & """>&nbsp;"
										If StrComp(sTemp, (CStr(oRecordset.Fields("TaskHour").Value) & "." & CStr(oRecordset.Fields("TaskMinute").Value)), vbBinaryCompare) <> 0 Then
											Response.Write "<FONT FACE=""Arial"" SIZE=""2"">"
												If CInt(oRecordset.Fields("TaskHour").Value) = 25 Then
													Response.Write "--:--"
												Else
													Response.Write CStr(oRecordset.Fields("TaskHour").Value) & ":"
													Response.Write Right(("0" & CStr(oRecordset.Fields("TaskMinute").Value)), Len("00"))
												End If
											Response.Write "</FONT>"
										End If
										sTemp = CStr(oRecordset.Fields("TaskHour").Value) & "." & CStr(oRecordset.Fields("TaskMinute").Value)
									Response.Write "</TD>"
									Response.Write "<TD BGCOLOR=""#" & sBGColor & """>"
										sTemp = "?"
										If InStr(1, aCalendarComponent(S_TARGET_PAGE_CALENDAR), "?", vbBinaryCompare) > 0 Then sTemp = "&"
										If Len(aCalendarComponent(S_JAVASCRIPT_CALENDAR)) = 0 Then
											Response.Write "<A HREF=""" & aCalendarComponent(S_TARGET_PAGE_CALENDAR) & sTemp & "UserID=" & oRecordset.Fields("UserID") & "&GroupID=" & oRecordset.Fields("GroupID") & "&Year=" & oRecordset.Fields("TaskYear") & "&Month=" & oRecordset.Fields("TaskMonth") & "&Day=" & oRecordset.Fields("TaskDay") & "&Hour=" & oRecordset.Fields("TaskHour") & "&Minute=" & oRecordset.Fields("TaskMinute") & "&TaskID=" & oRecordset.Fields("TaskID") & """ TARGET=""" & aCalendarComponent(S_TARGET_FRAME_CALENDAR) & """>"
										Else
											Response.Write "<A HREF=""javascript: " & aCalendarComponent(S_JAVASCRIPT_CALENDAR) & "(" & oRecordset.Fields("UserID") & ", " & oRecordset.Fields("GroupID") & ", " & oRecordset.Fields("TaskYear") & ", " & oRecordset.Fields("TaskMonth") & ", " & oRecordset.Fields("TaskDay") & ", " & oRecordset.Fields("TaskHour") & ", " & oRecordset.Fields("TaskMinute") & ", " & oRecordset.Fields("TaskID") & ")"" TARGET=""" & aCalendarComponent(S_TARGET_FRAME_CALENDAR) & """>"
										End If
											Response.Write "<IMG SRC=""Images/IcnActivitySmall" & oRecordset.Fields("TypeID") & ".gif"" WIDTH=""20"" HEIGHT=""20"" ALT=""" & CleanStringForHTML(oRecordset.Fields("TaskTitle")) & " (" & asTaskTypesCalendar(CInt(oRecordset.Fields("TypeID"))) & ")"" BORDER=""0"" ALIGN=""LEFT"" />"
										Response.Write "</A>"
										If CInt(oRecordset.Fields("TaskDone").Value) = 1 Then
											Response.Write "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Actividad realizada"" />"
										End If
									Response.Write "</TD>"
									Response.Write "<TD WIDTH=""1"" BGCOLOR=""#" & CStr(oRecordset.Fields("TaskColor").Value) & """>&nbsp;&nbsp;</TD>"
								Response.Write "</TR>"
								oRecordset.MoveNext
								If Err.number <> 0 Then
									lErrorNumber = Err.number
									sErrorDescription = "Ocurrió un error al desplegar las actividades para el día."
									If Len(Err.description) > 0 Then
										sErrorDescription = sErrorDescription & "<BR />" & Err.description
									End If
									Exit Do
								End If
								bEven = Not bEven
								aCalendarComponent(N_TASK_COUNT_CALENDAR) = aCalendarComponent(N_TASK_COUNT_CALENDAR) + 1
							Loop
						Response.Write "</TABLE>"
					End If
				Response.Write "</TD>"
				Response.Write "<TD BCOLOR=""#000000"" WIDTH=""1""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
				oTempDate = DateAdd("d", 1, oTempDate)
			Next
		Response.Write "</TR>" & vbNewLine
	Response.Write "</TABLE>" & vbNewLine

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayWeek = lErrorNumber
	Err.Clear
End Function

Function DisplayMonth(oRequest, aCalendarComponent, sErrorDescription)
'************************************************************
'Purpose: To display a month using the information from the
'         component
'Inputs:  oRequest, bAddRadioButtons, aCalendarComponent
'Outputs: aCalendarComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayMonth"
	Dim sOpenSpecialTags
	Dim sCloseSpecialTags
	Dim iIndex
	Dim jIndex
	Dim sTemp
	Dim sYearMonth
	Dim sMarkedDay
	Dim bGrayFrame
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCalendarComponent(B_COMPONENT_INITIALIZED_CALENDAR)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCalendarComponent(oRequest, aCalendarComponent)
	End If

	sTemp = "?"
	If InStr(1, aCalendarComponent(S_TARGET_PAGE_CALENDAR), "?", vbBinaryCompare) > 0 Then sTemp = "&"
	Response.Write "<DIV NAME=""MonthsNamesDiv"" ID=""MonthsNamesDiv"" CLASS=""ClassPopupItem"" STYLE=""left: 13px; top: 2px; filter:alpha(opacity=90);""><TABLE BGCOLOR=""#" & S_MAIN_COLOR_FOR_GUI & """ BORDER=""0"" CELLSPACING=""1"" CELLPADDING=""0""><TR><TD><TABLE BGCOLOR=""#FFFFFF"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""2"">" & vbNewLine
		For iIndex = 1 To 6
			Response.Write "<TR>" & vbNewLine
				Response.Write "<TD><FONT FACE=""Verdana"" SIZE=""1""><A HREF=""" & aCalendarComponent(S_TARGET_PAGE_CALENDAR) & sTemp & "Year=" & aCalendarComponent(N_YEAR_CALENDAR) & "&Month=" & iIndex & "&Day=1&FromArrow=1"" TARGET=""" & aCalendarComponent(S_TARGET_FRAME_CALENDAR) & """ CLASS=""SpecialLink"">" & asMonthNames_es(iIndex) & "&nbsp;" & aCalendarComponent(N_YEAR_CALENDAR) & "</A>&nbsp;&nbsp;</FONT></TD>" & vbNewLine
				Response.Write "<TD><FONT FACE=""Verdana"" SIZE=""1""><A HREF=""" & aCalendarComponent(S_TARGET_PAGE_CALENDAR) & sTemp & "Year=" & aCalendarComponent(N_YEAR_CALENDAR) & "&Month=" & iIndex + 6 & "&Day=1&FromArrow=1"" TARGET=""" & aCalendarComponent(S_TARGET_FRAME_CALENDAR) & """ CLASS=""SpecialLink"">" & asMonthNames_es(iIndex + 6) & "&nbsp;" & aCalendarComponent(N_YEAR_CALENDAR) & "</A></FONT></TD>" & vbNewLine
				If iIndex = 1 Then Response.Write "<TD VALIGN=""TOP"" ROWSPAN=""6""><A HREF=""javascript: HidePopupItem('MonthsNamesDiv', document.all['MonthsNamesDiv'])""><IMG SRC=""Images/BtnClose.gif"" WIDTH=""11"" HEIGHT=""10"" ALT=""Cerrar"" BORDER=""0""></A></TD>" & vbNewLine
			Response.Write "</TR>" & vbNewLine
		Next
	Response.Write "</TABLE></TD></TR></TABLE></DIV>" & vbNewLine
	Response.Write "<TABLE WIDTH=""140"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
		Response.Write "<TR>" & vbNewLine
			Response.Write "<TD BGCOLOR=""#CCCCCC"">&nbsp;</TD>" & vbNewLine
			Response.Write "<TD BGCOLOR=""#CCCCCC"" ALIGN=""CENTER"" VALIGN=""MIDDLE"" COLSPAN=""5""><A HREF=""javascript: ShowPopupItem('MonthsNamesDiv', document.all['MonthsNamesDiv'], false)"" CLASS=""CalendarClass""><FONT FACE=""Verdana"" SIZE=""1""><B>" & aCalendarComponent(S_MONTH_CALENDAR) & "&nbsp;&nbsp;" & aCalendarComponent(N_YEAR_CALENDAR) & "</B></FONT></A></TD>" & vbNewLine
			Response.Write "<TD BGCOLOR=""#CCCCCC"">&nbsp;</TD>" & vbNewLine
		Response.Write "</TR>" & vbNewLine
		Response.Write "<TR>" & vbNewLine
			For iIndex = 0 To 6
				Response.Write "<TD WIDTH=""20"" ALIGN=""RIGHT""><FONT FACE=""Verdana"" SIZE=""1"">" & Left(asDayNamesCalendar(iIndex), Len("D")) & "</FONT></TD>" & vbNewLine
			Next
		Response.Write "</TR>" & vbNewLine
		Response.Write "<TR BGCOLOR=""#000000""><TD COLSPAN=""7""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR>" & vbNewLine
		iIndex = 1
		sYearMonth = "," & aCalendarComponent(N_YEAR_CALENDAR) & Right(("0" & aCalendarComponent(N_MONTH_CALENDAR)), Len("00"))
		If aCalendarComponent(N_FIRST_DAY_CALENDAR) <> vbSunday Then
			Response.Write "<TR>" & vbNewLine
				For iIndex = 1 To aCalendarComponent(N_FIRST_DAY_CALENDAR) - 1
					Response.Write "<TD><FONT FACE=""Verdana"" SIZE=""1"">&nbsp;</FONT></TD>" & vbNewLine
				Next
				For iIndex = 1 To (8 - aCalendarComponent(N_FIRST_DAY_CALENDAR))
					sMarkedDay = sYearMonth & Right(("0" & iIndex), Len("00")) & ","
					Response.Write "<TD "
						bGrayFrame = False
						If aCalendarComponent(N_YEAR_CALENDAR) = Year(Date()) Then
							If aCalendarComponent(N_MONTH_CALENDAR) = Month(Date()) Then
								If iIndex = Day(Date()) Then
									If InStr(1, ("," & aCalendarComponent(N_SELECTED_DAY_CALENDAR) & ","), sMarkedDay, vbBinaryCompare) = 0 Then
										Response.Write "BACKGROUND=""Images/FrameRed.gif"" "
									Else
										Response.Write "BACKGROUND=""Images/FrameRedGray.gif"" "
									End If
								Else
									bGrayFrame = True
								End If
							Else
								bGrayFrame = True
							End If
						Else
							bGrayFrame = True
						End If
						If bGrayFrame And (InStr(1, ("," & aCalendarComponent(N_SELECTED_DAY_CALENDAR) & ","), sMarkedDay, vbBinaryCompare) > 0) Then
							Response.Write "BACKGROUND=""Images/FrameGray.gif"" "
						End If
					Response.Write "ALIGN=""RIGHT""><A ID=""" & aCalendarComponent(N_YEAR_CALENDAR) & Right(("0" & aCalendarComponent(N_MONTH_CALENDAR)), Len("00")) & Right(("0" & iIndex), Len("00")) & """"
						If aCalendarComponent(B_ONLY_HOLIDAYS_CALENDAR) Then
							If DateIsHoliday(oRequest, oADODBConnection, aCalendarComponent, iIndex, sErrorDescription) Then
								Response.Write " HREF=""" & aCalendarComponent(S_TARGET_PAGE_CALENDAR) & sTemp & "Year=" & aCalendarComponent(N_YEAR_CALENDAR) & "&Month=" & aCalendarComponent(N_MONTH_CALENDAR) & "&Day=" & iIndex & """"
							End If
						ElseIf (Not aCalendarComponent(B_ONLY_SUNDAY_CALENDAR)) And (Not aCalendarComponent(B_ONLY_PAYDAYS_CALENDAR)) Then 
							Response.Write " HREF=""" & aCalendarComponent(S_TARGET_PAGE_CALENDAR) & sTemp & "Year=" & aCalendarComponent(N_YEAR_CALENDAR) & "&Month=" & aCalendarComponent(N_MONTH_CALENDAR) & "&Day=" & iIndex & """"
						End If
						If InStr(1, oRequest("Action").Item, "Holidays") > 0 Then Response.Write " onClick=""return AddHolidayDescription(this.id);"""
					Response.Write " TARGET=""" & aCalendarComponent(S_TARGET_FRAME_CALENDAR) & """ CLASS=""CalendarClass""><FONT FACE=""Verdana"" SIZE=""1"""
						If DateIsHoliday(oRequest, oADODBConnection, aCalendarComponent, iIndex, sErrorDescription) Then
							Response.Write " COLOR=""#FF0000"""
						ElseIf (jIndex = 1) And aCalendarComponent(B_GRAY_SUNDAY_CALENDAR) Then
							Response.Write " COLOR=""#606060"""
						End If
					Response.Write ">"

						sOpenSpecialTags = ""
						sCloseSpecialTags = ""
						If aCalendarComponent(B_ONLY_HOLIDAYS_CALENDAR) Then
							If DateIsHoliday(oRequest, oADODBConnection, aCalendarComponent, iIndex, sErrorDescription) Then
								sOpenSpecialTags = "<B>"
								sCloseSpecialTags = "</B>"
							End If
						ElseIf InStr(1, ("," & aCalendarComponent(S_MARKED_DAYS_CALENDAR) & ","), sMarkedDay, vbBinaryCompare) > 0 Then
							sOpenSpecialTags = "<B>"
							sCloseSpecialTags = "</B>"
						End If
						If InStr(1, ("," & aCalendarComponent(S_SPECIAL_DAYS_CALENDAR) & ","), sMarkedDay, vbBinaryCompare) > 0 Then
							sOpenSpecialTags = sOpenSpecialTags & "<FONT COLOR=""#003399"">"
							sCloseSpecialTags = "</FONT>" & sCloseSpecialTags
						End If
						Response.Write sOpenSpecialTags & iIndex & sCloseSpecialTags
					Response.Write "</FONT></A></TD>" & vbNewLine
				Next
			Response.Write "</TR>" & vbNewLine
		End If

		Do While (iIndex <= aCalendarComponent(N_DAYS_CALENDAR))
			Response.Write "<TR>" & vbNewLine
				For jIndex = 1 To 7
					sMarkedDay = sYearMonth & Right(("0" & iIndex), Len("00")) & ","
					If iIndex <= aCalendarComponent(N_DAYS_CALENDAR) Then
						Response.Write "<TD "
							bGrayFrame = False
							If aCalendarComponent(N_YEAR_CALENDAR) = Year(Date()) Then
								If aCalendarComponent(N_MONTH_CALENDAR) = Month(Date()) Then
									If iIndex = Day(Date()) Then
										If InStr(1, ("," & aCalendarComponent(N_SELECTED_DAY_CALENDAR) & ","), sMarkedDay, vbBinaryCompare) = 0 Then
											Response.Write "BACKGROUND=""Images/FrameRed.gif"" "
										Else
											Response.Write "BACKGROUND=""Images/FrameRedGray.gif"" "
										End If
									Else
										bGrayFrame = True
									End If
								Else
									bGrayFrame = True
								End If
							Else
								bGrayFrame = True
							End If
							If bGrayFrame And (InStr(1, ("," & aCalendarComponent(N_SELECTED_DAY_CALENDAR) & ","), sMarkedDay, vbBinaryCompare) > 0) Then
								Response.Write "BACKGROUND=""Images/FrameGray.gif"" "
							End If
						Response.Write "ALIGN=""RIGHT""><A ID=""" & aCalendarComponent(N_YEAR_CALENDAR) & Right(("0" & aCalendarComponent(N_MONTH_CALENDAR)), Len("00")) & Right(("0" & iIndex), Len("00")) & """"
							If aCalendarComponent(B_ONLY_PAYDAYS_CALENDAR) Then
								If (iIndex = 15) Or (iIndex = aCalendarComponent(N_DAYS_CALENDAR)) Then Response.Write " HREF=""" & aCalendarComponent(S_TARGET_PAGE_CALENDAR) & sTemp & "Year=" & aCalendarComponent(N_YEAR_CALENDAR) & "&Month=" & aCalendarComponent(N_MONTH_CALENDAR) & "&Day=" & iIndex & """"
							ElseIf aCalendarComponent(B_ONLY_HOLIDAYS_CALENDAR) Then
								If DateIsHoliday(oRequest, oADODBConnection, aCalendarComponent, iIndex, sErrorDescription) Then
									Response.Write " HREF=""" & aCalendarComponent(S_TARGET_PAGE_CALENDAR) & sTemp & "Year=" & aCalendarComponent(N_YEAR_CALENDAR) & "&Month=" & aCalendarComponent(N_MONTH_CALENDAR) & "&Day=" & iIndex & """"
								End If
							ElseIf (Not aCalendarComponent(B_ONLY_SUNDAY_CALENDAR)) Or (jIndex = 1) Then
								Response.Write " HREF=""" & aCalendarComponent(S_TARGET_PAGE_CALENDAR) & sTemp & "Year=" & aCalendarComponent(N_YEAR_CALENDAR) & "&Month=" & aCalendarComponent(N_MONTH_CALENDAR) & "&Day=" & iIndex & """"
							End If
							If InStr(1, oRequest("Action").Item, "Holidays") > 0 Then Response.Write " onClick=""return AddHolidayDescription(this.id);"""
						Response.Write " TARGET=""" & aCalendarComponent(S_TARGET_FRAME_CALENDAR) & """ CLASS=""CalendarClass""><FONT FACE=""Verdana"" SIZE=""1"""
							If DateIsHoliday(oRequest, oADODBConnection, aCalendarComponent, iIndex, sErrorDescription) Then
								Response.Write "COLOR=""#FF0000"""
							ElseIf (jIndex = 1) And aCalendarComponent(B_GRAY_SUNDAY_CALENDAR) Then
								Response.Write "COLOR=""#606060"""
							End If
						Response.Write ">"
							sOpenSpecialTags = ""
							sCloseSpecialTags = ""
							If aCalendarComponent(B_ONLY_HOLIDAYS_CALENDAR) Then
								If DateIsHoliday(oRequest, oADODBConnection, aCalendarComponent, iIndex, sErrorDescription) Then
									sOpenSpecialTags = "<B>"
									sCloseSpecialTags = "</B>"
								End If
							ElseIf InStr(1, ("," & aCalendarComponent(S_MARKED_DAYS_CALENDAR) & ","), sMarkedDay, vbBinaryCompare) > 0 Then
								sOpenSpecialTags = "<B>"
								sCloseSpecialTags = "</B>"
							End If
							If InStr(1, ("," & aCalendarComponent(S_SPECIAL_DAYS_CALENDAR) & ","), sMarkedDay, vbBinaryCompare) > 0 Then
								sOpenSpecialTags = sOpenSpecialTags & "<FONT COLOR=""#003399"">"
								sCloseSpecialTags = "</FONT>" & sCloseSpecialTags
							End If
							Response.Write sOpenSpecialTags & iIndex & sCloseSpecialTags
						Response.Write "</FONT></A></TD>" & vbNewLine
						iIndex = iIndex + 1
					Else
						Exit Do
					End If
				Next
			Response.Write "</TR>" & vbNewLine
		Loop

		If jIndex < 7 Then
			For iIndex = jIndex To 7
				Response.Write "<TD><FONT FACE=""Verdana"" SIZE=""1"">&nbsp;</FONT></TD>" & vbNewLine
			Next
			Response.Write "</TR>" & vbNewLine
		End If
	Response.Write "</TABLE>" & vbNewLine

	DisplayMonth = lErrorNumber
	Err.Clear
End Function

Function DisplayBigMonth(oRequest, oADODBConnection, aCalendarComponent, sErrorDescription)
'************************************************************
'Purpose: To display a month using the information from the
'         component
'Inputs:  oRequest, oADODBConnection, bAddRadioButtons, aCalendarComponent
'Outputs: aCalendarComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayBigMonth"
	Dim sOpenSpecialTags
	Dim sCloseSpecialTags
	Dim iIndex
	Dim jIndex
	Dim sTemp
	Dim sYearMonth
	Dim sMarkedDay
	Dim bGrayFrame
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCalendarComponent(B_COMPONENT_INITIALIZED_CALENDAR)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCalendarComponent(oRequest, aCalendarComponent)
	End If

	sErrorDescription = "No se pudo obtener las actividades del mes de " & aCalendarComponent(S_MONTH_CALENDAR) & " de " & aCalendarComponent(N_YEAR_CALENDAR) & "."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From CalendarTasks Where (TaskYear=" & aCalendarComponent(N_YEAR_CALENDAR) & ") And (TaskMonth= " & aCalendarComponent(N_MONTH_CALENDAR) & ") And ((UserID=" & aCalendarComponent(N_USER_ID_CALENDAR) & ") Or (UserID=-1 And GroupID=-1) Or (UserID=-1 And GroupID=" & aCalendarComponent(N_GROUP_ID_CALENDAR) & ")) Order By TaskDay, TaskHour, TaskMinute, TaskID", "CalendarComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Response.Write "<TABLE WIDTH=""595"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			Response.Write "<TR>" & vbNewLine
				Response.Write "<TD BGCOLOR=""#CCCCCC"" ALIGN=""CENTER"" VALIGN=""MIDDLE"" COLSPAN=""15"">&nbsp;<FONT FACE=""Arial"" SIZE=""2""><B>" & aCalendarComponent(S_MONTH_CALENDAR) & "&nbsp;&nbsp;" & aCalendarComponent(N_YEAR_CALENDAR) & "</B></FONT>&nbsp;</TD>" & vbNewLine
			Response.Write "</TR>" & vbNewLine
			Response.Write "<TR BGCOLOR=""#000000""><TD COLSPAN=""15""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR>" & vbNewLine
			Response.Write "<TR>" & vbNewLine
				Response.Write "<TD BGCOLOR=""#000000"" WIDTH=""1""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>" & vbNewLine
				For iIndex = 0 To 6
					Response.Write "<TD WIDTH=""85"" ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">" & asDayNamesCalendar(iIndex) & "</FONT></TD>" & vbNewLine
					Response.Write "<TD BGCOLOR=""#000000"" WIDTH=""1""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>" & vbNewLine
				Next
			Response.Write "</TR>" & vbNewLine
			Response.Write "<TR BGCOLOR=""#000000""><TD COLSPAN=""15""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR>" & vbNewLine
			iIndex = 1
			sYearMonth = "," & aCalendarComponent(N_YEAR_CALENDAR) & Right(("0" & aCalendarComponent(N_MONTH_CALENDAR)), Len("00"))
			If aCalendarComponent(N_FIRST_DAY_CALENDAR) <> vbSunday Then
				Response.Write "<TR>" & vbNewLine
					Response.Write "<TD BGCOLOR=""#000000"" WIDTH=""1""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>" & vbNewLine
					For iIndex = 1 To aCalendarComponent(N_FIRST_DAY_CALENDAR) - 1
						Response.Write "<TD>&nbsp;</TD>" & vbNewLine
						Response.Write "<TD BGCOLOR=""#000000"" WIDTH=""1""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>" & vbNewLine
					Next
					For iIndex = 1 To (8 - aCalendarComponent(N_FIRST_DAY_CALENDAR))
						sMarkedDay = sYearMonth & Right(("0" & iIndex), Len("00")) & ","
						Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"
							sOpenSpecialTags = ""
							sCloseSpecialTags = ""
							If InStr(1, ("," & aCalendarComponent(S_MARKED_DAYS_CALENDAR) & ","), sMarkedDay, vbBinaryCompare) > 0 Then
								sOpenSpecialTags = "<B>"
								sCloseSpecialTags = "</B>"
							End If
							If (aCalendarComponent(N_YEAR_CALENDAR) = Year(Date())) And (aCalendarComponent(N_MONTH_CALENDAR) = Month(Date())) And (iIndex = Day(Date())) Then
								sOpenSpecialTags = sOpenSpecialTags & "<FONT COLOR=""#D20000"">"
								sCloseSpecialTags = "</FONT>" & sCloseSpecialTags
							ElseIf InStr(1, ("," & aCalendarComponent(S_SPECIAL_DAYS_CALENDAR) & ","), sMarkedDay, vbBinaryCompare) > 0 Then
								sOpenSpecialTags = sOpenSpecialTags & "<FONT COLOR=""#003399"">"
								sCloseSpecialTags = "</FONT>" & sCloseSpecialTags
							End If
							Response.Write sOpenSpecialTags & iIndex & sCloseSpecialTags & "<BR />"
							If Not oRecordset.EOF Then
								If CInt(oRecordset.Fields("TaskDay").Value) = iIndex Then
									Response.Write "<FONT SIZE=""1"">"
										Do While Not oRecordset.EOF
											If iIndex <> CInt(oRecordset.Fields("TaskDay").Value) Then Exit Do
											If CInt(oRecordset.Fields("TaskHour").Value) <> 25 Then
												Response.Write CStr(oRecordset.Fields("TaskHour").Value) & ":" & Right(("0" & CStr(oRecordset.Fields("TaskMinute").Value)), Len("00")) & "&nbsp;"
											Else
												Response.Write "--:--&nbsp;"
											End If
											sTemp = "?"
											If InStr(1, aCalendarComponent(S_TARGET_PAGE_CALENDAR), "?", vbBinaryCompare) > 0 Then sTemp = "&"
											If Len(aCalendarComponent(S_JAVASCRIPT_CALENDAR)) = 0 Then
												Response.Write "<A HREF=""" & aCalendarComponent(S_TARGET_PAGE_CALENDAR) & sTemp & "UserID=" & CStr(oRecordset.Fields("UserID").Value) & "&GroupID=" & CStr(oRecordset.Fields("GroupID").Value) & "&Year=" & aCalendarComponent(N_YEAR_CALENDAR) & "&Month=" & aCalendarComponent(N_MONTH_CALENDAR) & "&Day=" & iIndex & "&Hour=" & CStr(oRecordset.Fields("TaskHour").Value) & "&Minute=" & CStr(oRecordset.Fields("TaskMinute").Value) & "&TaskID=" & CStr(oRecordset.Fields("TaskID").Value) & """ TARGET=""" & aCalendarComponent(S_TARGET_FRAME_CALENDAR) & """>"
											Else
												Response.Write "<A HREF=""javascript: " & aCalendarComponent(S_JAVASCRIPT_CALENDAR) & "(" & CStr(oRecordset.Fields("UserID").Value) & ", " & CStr(oRecordset.Fields("GroupID").Value) & ", " & aCalendarComponent(N_YEAR_CALENDAR) & ", " & aCalendarComponent(N_MONTH_CALENDAR) & ", " & iIndex & ", " & CStr(oRecordset.Fields("TaskHour").Value) & ", " & CStr(oRecordset.Fields("TaskMinute").Value) & ", " & CStr(oRecordset.Fields("TaskID").Value) & ")"" TARGET=""" & aCalendarComponent(S_TARGET_FRAME_CALENDAR) & """>"
											End If
												Response.Write "<FONT COLOR=""#" & CStr(oRecordset.Fields("TaskColor").Value) & """>" & CleanStringForHTML(CStr(oRecordset.Fields("TaskTitle").Value)) & "</FONT>"
											Response.Write "</A><BR />"
											If Len(CStr(oRecordset.Fields("TaskDescription").Value)) > 0 Then Response.Write "<FONT COLOR=""#" & CStr(oRecordset.Fields("TaskColor").Value) & """>" & CleanStringForHTML(CStr(oRecordset.Fields("TaskDescription").Value)) & "<BR />"
											Response.Write "<BR />"
											oRecordset.MoveNext
											If Err.number <> 0 Then Exit Do
										Loop
									Response.Write "</FONT>"
								Else
									Response.Write "<BR /><BR /><BR /><BR /><BR />"
								End If
							Else
								Response.Write "<BR /><BR /><BR /><BR /><BR />"
							End If
						Response.Write "</FONT></TD>" & vbNewLine
						Response.Write "<TD BGCOLOR=""#000000"" WIDTH=""1""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>" & vbNewLine
					Next
				Response.Write "</TR>" & vbNewLine
				Response.Write "<TR BGCOLOR=""#000000""><TD COLSPAN=""15""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR>" & vbNewLine
			End If

			Do While (iIndex <= aCalendarComponent(N_DAYS_CALENDAR))
				Response.Write "<TR>" & vbNewLine
					Response.Write "<TD BGCOLOR=""#000000"" WIDTH=""1""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>" & vbNewLine
					For jIndex = 1 To 7
						sMarkedDay = sYearMonth & Right(("0" & iIndex), Len("00")) & ","
						If iIndex <= aCalendarComponent(N_DAYS_CALENDAR) Then
							Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"""
								If (jIndex = 1) And aCalendarComponent(B_GRAY_SUNDAY_CALENDAR) Then
									Response.Write "COLOR=""#606060"""
								End If
							Response.Write ">"
								sOpenSpecialTags = ""
								sCloseSpecialTags = ""
								If InStr(1, ("," & aCalendarComponent(S_MARKED_DAYS_CALENDAR) & ","), sMarkedDay, vbBinaryCompare) > 0 Then
									sOpenSpecialTags = "<B>"
									sCloseSpecialTags = "</B>"
								End If
								If (aCalendarComponent(N_YEAR_CALENDAR) = Year(Date())) And (aCalendarComponent(N_MONTH_CALENDAR) = Month(Date())) And (iIndex = Day(Date())) Then
									sOpenSpecialTags = sOpenSpecialTags & "<FONT COLOR=""#D20000"">"
									sCloseSpecialTags = "</FONT>" & sCloseSpecialTags
								ElseIf InStr(1, ("," & aCalendarComponent(S_SPECIAL_DAYS_CALENDAR) & ","), sMarkedDay, vbBinaryCompare) > 0 Then
									sOpenSpecialTags = sOpenSpecialTags & "<FONT COLOR=""#003399"">"
									sCloseSpecialTags = "</FONT>" & sCloseSpecialTags
								End If
								Response.Write sOpenSpecialTags & iIndex & sCloseSpecialTags & "<BR />"
								If Not oRecordset.EOF Then
									If iIndex = CInt(oRecordset.Fields("TaskDay").Value) Then
										Response.Write "<FONT SIZE=""1"">"
											Do While Not oRecordset.EOF
												If CInt(oRecordset.Fields("TaskHour").Value) <> 25 Then
													Response.Write CStr(oRecordset.Fields("TaskHour").Value) & ":" & Right(("0" & CStr(oRecordset.Fields("TaskMinute").Value)), Len("00")) & "&nbsp;"
												Else
													Response.Write "--:--&nbsp;"
												End If
												If Len(aCalendarComponent(S_JAVASCRIPT_CALENDAR)) = 0 Then
													Response.Write "<A HREF=""" & aCalendarComponent(S_TARGET_PAGE_CALENDAR) & sTemp & "UserID=" & CStr(oRecordset.Fields("UserID").Value) & "&GroupID=" & CStr(oRecordset.Fields("GroupID").Value) & "&Year=" & aCalendarComponent(N_YEAR_CALENDAR) & "&Month=" & aCalendarComponent(N_MONTH_CALENDAR) & "&Day=" & iIndex & "&Hour=" & CStr(oRecordset.Fields("TaskHour").Value) & "&Minute=" & CStr(oRecordset.Fields("TaskMinute").Value) & "&TaskID=" & CStr(oRecordset.Fields("TaskID").Value) & """ TARGET=""" & aCalendarComponent(S_TARGET_FRAME_CALENDAR) & """>"
												Else
													Response.Write "<A HREF=""javascript: " & aCalendarComponent(S_JAVASCRIPT_CALENDAR) & "(" & CStr(oRecordset.Fields("UserID").Value) & ", " & CStr(oRecordset.Fields("GroupID").Value) & ", " & aCalendarComponent(N_YEAR_CALENDAR) & ", " & aCalendarComponent(N_MONTH_CALENDAR) & ", " & iIndex & ", " & CStr(oRecordset.Fields("TaskHour").Value) & ", " & CStr(oRecordset.Fields("TaskMinute").Value) & ", " & CStr(oRecordset.Fields("TaskID").Value) & ")"" TARGET=""" & aCalendarComponent(S_TARGET_FRAME_CALENDAR) & """>"
												End If
													Response.Write "<FONT COLOR=""#" & CStr(oRecordset.Fields("TaskColor").Value) & """>" & CleanStringForHTML(CStr(oRecordset.Fields("TaskTitle").Value)) & "</FONT>"
												Response.Write "</A><BR />"
												If Len(CStr(oRecordset.Fields("TaskDescription").Value)) > 0 Then Response.Write "<FONT COLOR=""#" & CStr(oRecordset.Fields("TaskColor").Value) & """>" & CleanStringForHTML(CStr(oRecordset.Fields("TaskDescription").Value)) & "<BR />"
												Response.Write "<BR />"
												oRecordset.MoveNext
												If Err.number <> 0 Then Exit Do
												If iIndex = CInt(oRecordset.Fields("TaskDay").Value) Then Exit Do
											Loop
										Response.Write "</FONT>"
									Else
										Response.Write "<BR /><BR /><BR /><BR /><BR />"
									End If
								Else
									Response.Write "<BR /><BR /><BR /><BR /><BR />"
								End If
							Response.Write "</FONT></TD>" & vbNewLine
							iIndex = iIndex + 1
							Response.Write "<TD BGCOLOR=""#000000"" WIDTH=""1""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>" & vbNewLine
						Else
							Exit Do
						End If
					Next
				Response.Write "</TR>" & vbNewLine
				Response.Write "<TR BGCOLOR=""#000000""><TD COLSPAN=""15""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR>" & vbNewLine
			Loop

			If jIndex <= 7 Then
				For iIndex = jIndex To 7
					Response.Write "<TD>&nbsp;</TD>" & vbNewLine
					Response.Write "<TD BGCOLOR=""#000000"" WIDTH=""1""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>" & vbNewLine
				Next
				Response.Write "</TR>" & vbNewLine
				Response.Write "<TR BGCOLOR=""#000000""><TD COLSPAN=""15""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR>" & vbNewLine
			End If
		Response.Write "</TABLE>" & vbNewLine
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayBigMonth = lErrorNumber
	Err.Clear
End Function

Function DisplayYear(oRequest, aCalendarComponent, sErrorDescription)
'************************************************************
'Purpose: To display a year using the information from the
'         component
'Inputs:  oRequest, bAddRadioButtons, aCalendarComponent
'Outputs: aCalendarComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayYear"
	Dim sOpenSpecialTags
	Dim sCloseSpecialTags
	Dim iIndex
	Dim jIndex
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCalendarComponent(B_COMPONENT_INITIALIZED_CALENDAR)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCalendarComponent(oRequest, aCalendarComponent)
	End If

	Response.Write "<TABLE WIDTH=""600"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
		For iIndex = 0 To 2
			Response.Write "<TR>" & vbNewLine
				For jIndex = 1 To 4
					Response.Write "<TD WIDTH=""140"" VALIGN=""TOP"">" & vbNewLine
						aCalendarComponent(N_MONTH_CALENDAR) = (iIndex * 4) + jIndex
						Call InitializeMonth(aCalendarComponent)
						lErrorNumber = DisplayMonth(oRequest, aCalendarComponent, sErrorDescription)
					Response.Write "</TD>" & vbNewLine
					Response.Write "<TD WIDTH=""10"">&nbsp;</TD>" & vbNewLine
					If lErrorNumber <> 0 Then Exit For
				Next
			Response.Write "</TR>" & vbNewLine
			Response.Write "<TR><TD COLSPAN=""8""><FONT SIZE=""1"">&nbsp;</FONT></TD></TR>" & vbNewLine
			If lErrorNumber <> 0 Then Exit For
		Next
	Response.Write "</TABLE>" & vbNewLine

	DisplayYear = lErrorNumber
	Err.Clear
End Function
%>