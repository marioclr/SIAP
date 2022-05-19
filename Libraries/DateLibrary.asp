<%
Const S_MONTH_NAMES_EN = ",January,February,March,April,May,June,July,August,September,October,November,December"
Const S_MONTH_NAMES_ES = ",Enero,Febrero,Marzo,Abril,Mayo,Junio,Julio,Agosto,Septiembre,Octubre,Noviembre,Diciembre"
Dim asMonthNames_en
Dim asMonthNames_es
asMonthNames_en = Split(S_MONTH_NAMES_EN, ",", -1, vbBinaryCompare)
asMonthNames_es = Split(S_MONTH_NAMES_ES, ",", -1, vbBinaryCompare)

Const S_DAY_NAMES_EN = ",Sunday,Monday,Tuesday,Wednesday,Thursday,Friday,Saturday"
Const S_DAY_NAMES_ES = ",Domingo,Lunes,Martes,Miércoles,Jueves,Viernes,Sábado"
Dim asDayNames_en
Dim asDayNames_es
asDayNames_en = Split(S_DAY_NAMES_EN, ",", -1, vbBinaryCompare)
asDayNames_es = Split(S_DAY_NAMES_ES, ",", -1, vbBinaryCompare)

Function AddDaysToSerialDate(sSerialDate, iDays)
'************************************************************
'Purpose: To convert the serial date and add it the given days
'Inputs:  sSerialDate, iDays
'Outputs: A new serial date
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddDaysToSerialDate"
	Dim oDate

	If Len(sSerialDate) = 0 Then
		sSerialDate = Left(GetSerialNumberForDate(""), Len("00000000"))
	End If
	oDate = DateSerial(Left(sSerialDate, Len("0000")), Mid(sSerialDate, Len("00000"), Len("00")), Right(sSerialDate, Len("00")))
	oDate = DateAdd("d", iDays, oDate)

	AddDaysToSerialDate = CLng(Year(oDate) & Right(("0" & Month(oDate)), Len("00")) & Right(("0" & Day(oDate)), Len("00")))
	Err.Clear
End Function

Function AddMonthsToSerialDate(sSerialDate, iMonths)
'************************************************************
'Purpose: To convert the serial date and add it the given days
'Inputs:  sSerialDate, iDays
'Outputs: A new serial date
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddMonthsToSerialDate"
	Dim oDate

	If Len(sSerialDate) = 0 Then
		sSerialDate = Left(GetSerialNumberForDate(""), Len("00000000"))
	End If
	oDate = DateSerial(Left(sSerialDate, Len("0000")), Mid(sSerialDate, Len("00000"), Len("00")), Right(sSerialDate, Len("00")))
	oDate = DateAdd("m", iMonths, oDate)

	AddMonthsToSerialDate = CLng(Year(oDate) & Right(("0" & Month(oDate)), Len("00")) & Right(("0" & Day(oDate)), Len("00")))
	Err.Clear
End Function

Function AddDaysToSerialDateForVacations(oADODBConnection, sSerialDate, iDays, iJourneyType)
'************************************************************
'Purpose: To convert the serial date and add it the given days
'Inputs:  sSerialDate, iDays, iJourneyType
'Outputs: A new serial date
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddDaysToSerialDateForVacations"
	Dim oDate
	Dim iDay
	Dim lDateRevision
	Dim iHolidayDaysCount
	Dim iCount
	Dim sErrorDescription

	If Len(sSerialDate) = 0 Then
		sSerialDate = Left(GetSerialNumberForDate(""), Len("00000000"))
	End If

	For iDay = 0 To iDays
		lDateRevision = AddDaysToSerialDate(sSerialDate, iDay)
		oDate = DateSerial(Left(lDateRevision, Len("0000")), Mid(lDateRevision, Len("00000"), Len("00")), Right(lDateRevision, Len("00")))
		Select Case iJourneyType
			Case 1
				If ((Weekday(oDate) = vbSunday) Or (Weekday(oDate) = vbSaturday) Or ( IsHoliday(oADODBConnection, lDateRevision, sErrorDescription))) Then
					iHolidayDaysCount = iHolidayDaysCount + 1
				End If
			Case 21
				If ((Weekday(oDate) = vbTuesday) Or (Weekday(oDate) = vbThursday) Or (Weekday(oDate) = vbSaturday) Or (Weekday(oDate) = vbSunday)) Then
					iHolidayDaysCount = iHolidayDaysCount + 1
				End If
			Case 22
				If ((Weekday(oDate) = vbMonday) Or (Weekday(oDate) = vbWednesday) Or (Weekday(oDate) = vbFriday)) Then
					iHolidayDaysCount = iHolidayDaysCount + 1
				End If
			Case 23
				If ((Weekday(oDate) = vbSaturday) Or (Weekday(oDate) = vbSunday)) Then
					iHolidayDaysCount = iHolidayDaysCount + 1
				End If
			Case 31, 32
				If (((Weekday(oDate) = vbMonday) Or (Weekday(oDate) = vbTuesday) Or (Weekday(oDate) = vbWednesday) Or (Weekday(oDate) = vbThursday) Or (Weekday(oDate) = vbFriday)) And ( Not IsHoliday(oADODBConnection, lDateRevision, sErrorDescription))) Then
					iHolidayDaysCount = iHolidayDaysCount + 1
				End If
			Case 41
				If (((Weekday(oDate) = vbMonday) Or (Weekday(oDate) = vbTuesday) Or (Weekday(oDate) = vbWednesday) Or (Weekday(oDate) = vbThursday) Or (Weekday(oDate) = vbFriday) Or (Weekday(oDate) = vbSunday)) And (Not IsHoliday(oADODBConnection, lDateRevision, sErrorDescription))) Then
					iHolidayDaysCount = iHolidayDaysCount + 1
				End If
			Case 42
				If (((Weekday(oDate) = vbMonday) Or (Weekday(oDate) = vbTuesday) Or (Weekday(oDate) = vbWednesday) Or (Weekday(oDate) = vbThursday) Or (Weekday(oDate) = vbFriday) Or (Weekday(oDate) = vbSaturday)) And (Not IsHoliday(oADODBConnection, lDateRevision, sErrorDescription))) Then
					iHolidayDaysCount = iHolidayDaysCount + 1
				End If
		End Select
	Next

	iCount = 0
	Do While iCount < iHolidayDaysCount
		oDate = DateAdd("d", 1, oDate)
		Select Case iJourneyType
			Case 1
				If Not ((Weekday(oDate) = vbSunday) Or (Weekday(oDate) = vbSaturday) Or (IsHoliday(oADODBConnection, Left(GetSerialNumberForDate(oDate), Len("00000000")), sErrorDescription))) Then
					iCount = iCount + 1
				End If
			Case 21
				If Not ((Weekday(oDate) = vbTuesday) Or (Weekday(oDate) = vbThursday) Or (Weekday(oDate) = vbSaturday) Or (Weekday(oDate) = vbSunday)) Then
					iCount = iCount + 1
				End If
			Case 22
				If Not ((Weekday(oDate) = vbMonday) Or (Weekday(oDate) = vbWednesday) Or (Weekday(oDate) = vbFriday)) Then
					iCount = iCount + 1
				End If
			Case 23
				If Not ((Weekday(oDate) = vbSaturday) Or (Weekday(oDate) = vbSunday)) Then
					iCount = iCount + 1
				End If
			Case 31, 32
				If Not (((Weekday(oDate) = vbMonday) Or (Weekday(oDate) = vbTuesday) Or (Weekday(oDate) = vbWednesday) Or (Weekday(oDate) = vbThursday) Or (Weekday(oDate) = vbFriday)) And (Not IsHoliday(oADODBConnection, Left(GetSerialNumberForDate(oDate), Len("00000000")), sErrorDescription))) Then
					iCount = iCount + 1
				End If
			Case 41
				If Not (((Weekday(oDate) = vbMonday) Or (Weekday(oDate) = vbTuesday) Or (Weekday(oDate) = vbWednesday) Or (Weekday(oDate) = vbThursday) Or (Weekday(oDate) = vbFriday) Or (Weekday(oDate) = vbSunday)) And (Not IsHoliday(oADODBConnection, Left(GetSerialNumberForDate(oDate), Len("00000000")), sErrorDescription))) Then
					iCount = iCount + 1
				End If
			Case 42
				If Not (((Weekday(oDate) = vbMonday) Or (Weekday(oDate) = vbTuesday) Or (Weekday(oDate) = vbWednesday) Or (Weekday(oDate) = vbThursday) Or (Weekday(oDate) = vbFriday) Or (Weekday(oDate) = vbSaturday)) And (Not IsHoliday(oADODBConnection, Left(GetSerialNumberForDate(oDate), Len("00000000")), sErrorDescription))) Then
					iCount = iCount + 1
				End If
		End Select
	Loop

	AddDaysToSerialDateForVacations = CLng(Year(oDate) & Right(("0" & Month(oDate)), Len("00")) & Right(("0" & Day(oDate)), Len("00")))
	Err.Clear
End Function

Function AddYearsToSerialDate(sSerialDate, iYears)
'************************************************************
'Purpose: To convert the serial date and add it the given years
'Inputs:  sSerialDate, iDays
'Outputs: A new serial date
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddYearsToSerialDate"
	Dim oDate

	If Len(sSerialDate) = 0 Then
		sSerialDate = Left(GetSerialNumberForDate(""), Len("00000000"))
	End If
	oDate = DateSerial(Left(sSerialDate, Len("0000")), Mid(sSerialDate, Len("00000"), Len("00")), Right(sSerialDate, Len("00")))
	oDate = DateAdd("yyyy", iYears, oDate)

	AddYearsToSerialDate = CLng(Year(oDate) & Right(("0" & Month(oDate)), Len("00")) & Right(("0" & Day(oDate)), Len("00")))
	Err.Clear
End Function

Function CalculateAgeFromSerialNumber(lBirthDate, lEndDate)
'************************************************************
'Purpose: To get the age since the birth date to the end date
'Inputs:  lBirthDate, lEndDate
'Outputs: A number with the age
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CalculateAgeFromSerialNumber"
	Dim oDate
	Dim iAge

	If (Len(lEndDate) = 0) Or (CInt(lEndDate) = 0) Then
		lEndDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	End If

	If Len(lBirthDate) = 0 Then
		CalculateAgeFromSerialNumber = 0
	Else
		oDate = GetDateFromSerialNumber(lBirthDate)
		iAge = DateDiff("yyyy", oDate, GetDateFromSerialNumber(lEndDate))
		If (CInt(Right(lBirthDate, 4)) > CInt(Right(lEndDate, 4))) Then
			iAge = iAge - 1
		End If
		CalculateAgeFromSerialNumber = iAge
	End If

	Err.Clear
End Function

Function ConvertMinutesIntoDays(lValue)
'************************************************************
'Purpose: To calculate the number of minutes, hours and days
'         that the given number represents.
'Inputs:  lValue
'Outputs: A string representing the number of hours and days
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ConvertMinutesIntoDays"
	Dim nValue
	Dim dValue

	nValue = Int(lValue)
	dValue = lValue - nValue
	'ConvertMinutesIntoDays = Int(nValue/3600) & "hr " & Right("0" & (Int((nValue Mod 3600) / 60)), Len("00")) & "min " & Right("0" & ((nValue Mod 3600) Mod 60), Len("00")) & " seg"
	ConvertMinutesIntoDays = Int(nValue/1440) & " días " & Right("0" & Int((nValue Mod 1440) / 60), Len("00")) & " horas " & Right("0" & (Int((nValue Mod 1440) Mod 60)), Len("00")) & " minutos "
	If dValue > 0 Then
		ConvertMinutesIntoDays = ConvertMinutesIntoDays & CInt(dValue * 60) & " segundos"
	Else
		ConvertMinutesIntoDays = ConvertMinutesIntoDays & "00 segundos"
	End If
	Err.Clear
End Function

Function DisplayDate(iYear, iMonth, iDay, iHour, iMinute, iSecond)
'************************************************************
'Purpose: To display a date using the name of the months
'Inputs:  iYear, iMonth, iDay, iHour, iMinute, iSecond
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayDate"

	DisplayDate = ""
	If (iYear > 0) Or (iMonth > 0) Or (iDay > 0) Then
		If iDay > 0 Then DisplayDate = DisplayDate & iDay & " de "
		If iMonth > 0 Then DisplayDate = DisplayDate & asMonthNames_es(iMonth) & " de "
		If iYear > 0 Then DisplayDate = DisplayDate & iYear
		If Not IsNull(iHour) Then
			If (iHour > -1) And (iMinute > -1) Then
				DisplayDate = DisplayDate & " a las " & Right(("0" & iHour), Len("00")) & ":" & Right(("0" & iMinute), Len("00"))
				If Not IsNull(iSecond) Then
					If iSecond > -1 Then DisplayDate = DisplayDate & ":" & Right(("0" & iSecond), Len("00"))
				End If
			End If
		End If
	End If

	Err.Clear
End Function

Function DisplayDateAndTimeFromSerialNumber(sSerialDate, sSerialTime)
'************************************************************
'Purpose: To display a date using the name of the months
'Inputs:  sSerialDate, sSerialTime
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayDateAndTimeFromSerialNumber"
	Dim iYear
	Dim iMonth
	Dim iDay
	Dim iHour
	Dim iMinute
	Dim iSecond
	Dim sTempTime

	iYear = CLng(Left(sSerialDate, Len("1976")))
	If Err.number <> 0 Then
		iYear = 0
		Err.Clear()
	End If
	iMonth = CLng(Mid(sSerialDate, Len("19760"), Len("02")))
	If Err.number <> 0 Then
		iMonth = 0
		Err.Clear()
	End If
	iDay = CLng(Right(sSerialDate, Len("11")))
	If Err.number <> 0 Then
		iDay = 0
		Err.Clear()
	End If
	If Len(sSerialTime) > 0 Then
		sTempTime = sSerialTime
		If Len(sTempTime) < 5 Then sTempTime = sTempTime & "00"
		iHour = Left(sTempTime, Len("00"))
		iMinute = Mid(sTempTime, Len("000"), Len("00"))
		iSecond = Right(sTempTime, Len("00"))
	Else
		iHour = -1
		iMinute = -1
		iSecond = -1
	End If

	DisplayDateAndTimeFromSerialNumber = DisplayDate(iYear, iMonth, iDay, iHour, iMinute, iSecond)
	Err.Clear
End Function

Function DisplayDateCombos(iYear, iMonth, iDay, sYearFieldName, sMonthFieldName, sDayFieldName, iFirstYear, iLastYear, bUseJavaScriptValidation, bAddEmptyOptions)
'************************************************************
'Purpose: To display 3 combo lists containing the days, months
'         and years. The given date will be selected.
'Inputs:  iYear, iMonth, iDay, sYearFieldName, sMonthFieldName, sDayFieldName, iFirstYear, iLastYear, bUseJavaScriptValidation, bAddEmptyOptions
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayDateCombos"
	Dim iIndex

	iYear = CInt(iYear)
	iMonth = CInt(iMonth)
	iDay = CInt(iDay)
	If StrComp(GetASPFileName(""), "ErrorLog.asp", vbBinaryCompare) <> 0 Then
		DisplayDateCombos = "<A HREF=""javascript: OpenCalendarWindow(document.all['" & sYearFieldName & "'].form.name, '" & Replace(sYearFieldName, "Year", "") & "')"">"
			DisplayDateCombos = DisplayDateCombos & "<IMG SRC=""Images/IcnCalendar.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Seleccionar fecha"" BORDER=""0"" />"
		DisplayDateCombos = DisplayDateCombos & "</A>&nbsp;"
	End If
	DisplayDateCombos = DisplayDateCombos & "<NOBR><SELECT NAME=""" & sDayFieldName & """ ID=""" & sDayFieldName & "Cmb"" CLASS=""Lists"""
		If bAddEmptyOptions Then DisplayDateCombos = DisplayDateCombos & " onChange=""if (this.form." & sMonthFieldName & ".options[0].selected) {this.form." & sMonthFieldName & ".options[1].selected = true;}"""
	DisplayDateCombos = DisplayDateCombos & ">"
		If bAddEmptyOptions Then
			DisplayDateCombos = DisplayDateCombos & "<OPTION VALUE=""0""></OPTION>"
		End If
		For iIndex = 1 To 31
			DisplayDateCombos = DisplayDateCombos & "<OPTION VALUE="""
				DisplayDateCombos = DisplayDateCombos & Right(("0" & iIndex), Len("00"))
			DisplayDateCombos = DisplayDateCombos & """"
			If iDay > -1 Then
				If iDay = iIndex Then DisplayDateCombos = DisplayDateCombos & " SELECTED=""1"""
			End If
			DisplayDateCombos = DisplayDateCombos & ">" & iIndex & "</OPTION>"
		Next
	DisplayDateCombos = DisplayDateCombos & "</SELECT>&nbsp;-&nbsp;"
	DisplayDateCombos = DisplayDateCombos & "<SELECT NAME=""" & sMonthFieldName & """ ID=""" & sMonthFieldName & "Cmb"" CLASS=""Lists"""
		If bAddEmptyOptions And Not bUseJavaScriptValidation Then DisplayDateCombos = DisplayDateCombos & " onChange=""if (this.options[0].selected) {this.form." & sDayFieldName & ".options[0].selected = true;}"""
		If bUseJavaScriptValidation Then
			DisplayDateCombos = DisplayDateCombos & " onChange="""
			If bAddEmptyOptions Then DisplayDateCombos = DisplayDateCombos & "if (this.options[0].selected) {this.form." & sDayFieldName & ".options[0].selected = true;} "
			DisplayDateCombos = DisplayDateCombos & "ChangeDaysListGivenTheMonthAndYear(this.options[this.selectedIndex].value, this.form." & sYearFieldName & ".options[this.form." & sYearFieldName & ".selectedIndex].value, this.form." & sDayFieldName & ")"""
		End If
	DisplayDateCombos = DisplayDateCombos & ">"
		If bAddEmptyOptions Then
			DisplayDateCombos = DisplayDateCombos & "<OPTION VALUE=""0""></OPTION>"
		End If
		For iIndex = 1 To 12
			DisplayDateCombos = DisplayDateCombos & "<OPTION VALUE="""
				DisplayDateCombos = DisplayDateCombos & Right(("0" & iIndex), Len("00"))
			DisplayDateCombos = DisplayDateCombos & """"
			If iMonth > -1 Then
				If iMonth = iIndex Then DisplayDateCombos = DisplayDateCombos & " SELECTED=""1"""
			End If
			DisplayDateCombos = DisplayDateCombos & ">" & asMonthNames_es(iIndex) & "</OPTION>"
		Next
	DisplayDateCombos = DisplayDateCombos & "</SELECT>&nbsp;-&nbsp;"
	DisplayDateCombos = DisplayDateCombos & "<SELECT NAME=""" & sYearFieldName & """ ID=""" & sYearFieldName & "Cmb"" CLASS=""Lists"""
		If bUseJavaScriptValidation Then
			DisplayDateCombos = DisplayDateCombos & " onChange=""ChangeDaysListGivenTheMonthAndYear(this.form." & sMonthFieldName & ".options[this.form." & sMonthFieldName & ".selectedIndex].value, this.options[this.selectedIndex].value, this.form." & sDayFieldName & ")"""
		End If
	DisplayDateCombos = DisplayDateCombos & ">"
		If bAddEmptyOptions Then
			DisplayDateCombos = DisplayDateCombos & "<OPTION VALUE=""0""></OPTION>"
		End If
		For iIndex = iFirstYear To iLastYear
			DisplayDateCombos = DisplayDateCombos & "<OPTION VALUE=""" & iIndex & """"
				If iYear > -1 Then
					If iYear = iIndex Then DisplayDateCombos = DisplayDateCombos & " SELECTED=""1"""
				End If
			DisplayDateCombos = DisplayDateCombos & ">" & iIndex & "</OPTION>"
		Next
	DisplayDateCombos = DisplayDateCombos & "</SELECT></NOBR>"
	If bUseJavaScriptValidation And (iYear>0) And (iMonth>0) And (iDay>0) Then
		DisplayDateCombos = DisplayDateCombos & "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			DisplayDateCombos = DisplayDateCombos & "ChangeDaysListGivenTheMonthAndYear(document.all['" & sMonthFieldName & "'].options[document.all['" & sMonthFieldName & "'].selectedIndex].value, document.all['" & sYearFieldName & "'].options[document.all['" & sYearFieldName & "'].selectedIndex].value, document.all['" & sDayFieldName & "']);" & vbNewLine
			DisplayDateCombos = DisplayDateCombos & "SelectItemByText('" & iDay & "', false, document.all['" & sDayFieldName & "']);" & vbNewLine
		DisplayDateCombos = DisplayDateCombos & "//--></SCRIPT>" & vbNewLine
	End If

	Err.Clear
End Function

Function DisplayDateCombosUsingSerial(lDate, sPrefixFieldName, iFirstYear, iLastYear, bUseJavaScriptValidation, bAddEmptyOptions)
'************************************************************
'Purpose: To display 3 combo lists containing the days, months
'         and years. The given serial date will be selected
'Inputs:  lDate, sPrefixFieldName, iFirstYear, iLastYear, bUseJavaScriptValidation, , bAddEmptyOptions
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayDateCombosUsingSerial"
	Dim iYear
	Dim iMonth
	Dim iDay

	If (Not bAddEmptyOptions) And (Len(lDate) = 0) Then lDate = Left(GetSerialNumberForDate(""), Len("00000000"))
	iYear = CLng(Left(lDate, Len("0000")))
	If Err.number <> 0 Then
		iYear = 0
		Err.Clear()
	End If
	iMonth = CLng(Mid(lDate, Len("00000"), Len("00")))
	If Err.number <> 0 Then
		iMonth = 0
		Err.Clear()
	End If
	iDay = CLng(Right(lDate, Len("00")))
	If Err.number <> 0 Then
		iDay = 0
		Err.Clear()
	End If
	DisplayDateCombosUsingSerial = DisplayDateCombos(iYear, iMonth, iDay, sPrefixFieldName & "Year", sPrefixFieldName & "Month", sPrefixFieldName & "Day", iFirstYear, iLastYear, bUseJavaScriptValidation, bAddEmptyOptions)
End Function

Function DisplayDateFromSerialNumber(sSerialDate, iHour, iMinute, iSecond)
'************************************************************
'Purpose: To display a date using the name of the months
'Inputs:  sSerialDate, iHour, iMinute, iSecond
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayDateFromSerialNumber"
	Dim iYear
	Dim iMonth
	Dim iDay

	iYear = CLng(Left(sSerialDate, Len("1976")))
	If Err.number <> 0 Then
		iYear = 0
		Err.Clear()
	End If
	iMonth = CLng(Mid(sSerialDate, Len("19760"), Len("02")))
	If Err.number <> 0 Then
		iMonth = 0
		Err.Clear()
	End If
	iDay = CLng(Mid(sSerialDate, Len("1976021"), Len("11")))
	If Err.number <> 0 Then
		iDay = 0
		Err.Clear()
	End If

	DisplayDateFromSerialNumber = DisplayDate(iYear, iMonth, iDay, iHour, iMinute, iSecond)
	Err.Clear
End Function

Function DisplayNumericDateFromSerialNumber(lDate)
'************************************************************
'Purpose: To display a date using numbers
'Inputs:  lDate
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayNumericDateFromSerialNumber"

	DisplayNumericDateFromSerialNumber = Right(lDate, Len("11")) & "/" & Mid(lDate, Len("19760"), Len("02")) & "/" & Left(lDate, Len("1976"))
	Err.Clear
End Function

Function DisplaySerialDateAsHidden(lDate, sPrefixFieldName)
'************************************************************
'Purpose: To display 3 hidden fields containing the day, month
'         and year using the serial date
'Inputs:  lDate, sPrefixFieldName
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayDateAsHidden"
	Dim sYear
	Dim sMonth
	Dim sDay

	sYear = Left(lDate, Len("0000"))
	If Err.number <> 0 Then
		sYear = Year(Date())
		Err.Clear()
	End If
	sMonth = Mid(lDate, Len("00000"), Len("00"))
	If Err.number <> 0 Then
		sMonth = Right("0" & Month(Date()), Len("00"))
		Err.Clear()
	End If
	sDay = Mid(lDate, Len("0000000"), Len("00"))
	If Err.number <> 0 Then
		sDay = Right("0" & Day(Date()), Len("00"))
		Err.Clear()
	End If
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""" & sPrefixFieldName & "Year"" ID=""" & sPrefixFieldName & "YearHdn"" VALUE=""" & sYear & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""" & sPrefixFieldName & "Month"" ID=""" & sPrefixFieldName & "MonthHdn"" VALUE=""" & sMonth & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""" & sPrefixFieldName & "Day"" ID=""" & sPrefixFieldName & "DayHdn"" VALUE=""" & sDay & """ />"
End Function

Function DisplayShortDate(iYear, iMonth, iDay, iHour, iMinute, iSecond)
'************************************************************
'Purpose: To display a date using the name of the months
'Inputs:  iYear, iMonth, iDay, iHour, iMinute, iSecond
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayShortDate"

	DisplayShortDate = ""
	If (iYear > 0) Or (iMonth > 0) Or (iDay > 0) Then
		If iDay > 0 Then DisplayShortDate = DisplayShortDate & iDay & "/"
		If iMonth > 0 Then DisplayShortDate = DisplayShortDate & Left(asMonthNames_es(iMonth), Len("Feb")) & "/"
		If iYear > 0 Then DisplayShortDate = DisplayShortDate & iYear
		If Not IsNull(iHour) Then
			If (iHour > -1) And (iMinute > -1) Then
				DisplayShortDate = DisplayShortDate & ", " & Right(("0" & iHour), Len("00")) & ":" & Right(("0" & iMinute), Len("00"))
				If Not IsNull(iSecond) Then
					If iSecond > -1 Then DisplayShortDate = DisplayShortDate & ":" & Right(("0" & iSecond), Len("00"))
				End If
			End If
		End If
	End If

	Err.Clear
End Function

Function DisplayShortDateFromSerialNumber(sSerialDate, iHour, iMinute, iSecond)
'************************************************************
'Purpose: To display a date using the name of the months
'Inputs:  sSerialDate, iHour, iMinute, iSecond
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayShortDateFromSerialNumber"
	Dim iYear
	Dim iMonth
	Dim iDay

	iYear = CLng(Left(sSerialDate, Len("1976")))
	If Err.number <> 0 Then
		iYear = 0
		Err.Clear()
	End If
	iMonth = CLng(Mid(sSerialDate, Len("19760"), Len("02")))
	If Err.number <> 0 Then
		iMonth = 0
		Err.Clear()
	End If
	iDay = CLng(Mid(sSerialDate, Len("1976021"), Len("11")))
	If Err.number <> 0 Then
		iDay = 0
		Err.Clear()
	End If

	DisplayShortDateFromSerialNumber = DisplayShortDate(iYear, iMonth, iDay, iHour, iMinute, iSecond)
	Err.Clear
End Function

Function DisplayTimeCombos(iHour, iMinute, sHourFieldName, sMinuteFieldName, iFirstHour, iLastHour, iStep, bAddEmptyOptions)
'************************************************************
'Purpose: To display 2 combo lists containing the hours and
'         minutes. The given hour will be selected.
'Inputs:  iHour, iMinute, sHourFieldName, sMinuteFieldName, iFirstHour, iLastHour, iStep, bAddEmptyOptions
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayTimeCombos"
	Dim iIndex
	Dim sHTML

	sHTML = "<SELECT NAME=""" & sHourFieldName & """ ID=""" & sHourFieldName & "Cmb"" CLASS=""Lists"" onChange=""if (this.value=='" & iLastHour & "') this.form." & sMinuteFieldName& ".value='00';"">"
		If bAddEmptyOptions Then
			sHTML = sHTML & "<OPTION VALUE=""0""></OPTION>"
		End If
		For iIndex = iFirstHour To iLastHour
			sHTML = sHTML & "<OPTION VALUE="""
				sHTML = sHTML & Right(("0" & iIndex), Len("00"))
			sHTML = sHTML & """"
			If iHour > -1 Then
				If iHour = iIndex Then sHTML = sHTML & " SELECTED=""1"""
			End If
			sHTML = sHTML & ">" & iIndex
			If iStep >= 60 Then sHTML = sHTML & ":00"
			sHTML = sHTML & "</OPTION>"
		Next
	sHTML = sHTML & "</SELECT>"
	If iStep >= 60 Then
		sHTML = sHTML & "<INPUT TYPE=""HIDDEN"" NAME=""" & sMinuteFieldName & """ ID=""" & sMinuteFieldName & "Hdn"" VALUE=""0"" />"
	Else
		sHTML = sHTML & "&nbsp;:&nbsp;<SELECT NAME=""" & sMinuteFieldName & """ ID=""" & sMinuteFieldName & "Cmb"" CLASS=""Lists"" onChange=""if (this.form." & sHourFieldName & ".value=='" & iLastHour & "') this.value='00';"">"
			If bAddEmptyOptions Then
				sHTML = sHTML & "<OPTION VALUE=""0""></OPTION>"
			End If
			For iIndex = 0 To 59 Step iStep
				sHTML = sHTML & "<OPTION VALUE="""
					sHTML = sHTML & Right(("0" & iIndex), Len("00"))
				sHTML = sHTML & """"
				If iMinute > -1 Then
					If iMinute = iIndex Then sHTML = sHTML & " SELECTED=""1"""
				End If
				sHTML = sHTML & ">" & Right(("0" & iIndex), Len("00")) & "</OPTION>"
			Next
		sHTML = sHTML & "</SELECT>"
	End If

	DisplayTimeCombos = sHTML
	Err.Clear
End Function

Function DisplayTimeCombosUsingSerial(lHour, sPrefixFieldName, iFirstHour, iLastHour, iStep, bAddEmptyOptions)
'************************************************************
'Purpose: To display 2 combo lists containing the hours and
'         minutes. The given serial hour will be selected.
'Inputs:  lHour, sPrefixFieldName, iFirstHour, iLastHour, iStep, bAddEmptyOptions
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayTimeCombosUsingSerial"
	Dim sHour
	Dim iHour
	Dim iMinute

	sHour = Right(("0000" & lHour), Len("0000"))
	iHour = CLng(Left(sHour, Len("00")))
	If Err.number <> 0 Then
		iHour = 0
		Err.Clear()
	End If
	iMinute = CLng(Right(sHour, Len("00")))
	If Err.number <> 0 Then
		iMinute = 0
		Err.Clear()
	End If

	DisplayTimeCombosUsingSerial = DisplayTimeCombos(iHour, iMinute, sPrefixFieldName & "Hour", sPrefixFieldName & "Minute", iFirstHour, iLastHour, iStep, bAddEmptyOptions)
	Err.Clear
End Function

Function DisplayTimeFromSerialNumber(sSerialTime)
'************************************************************
'Purpose: To display the time using the format HH:MM:SS
'Inputs:  sSerialTime
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayTimeFromSerialNumber"
	Dim iHour
	Dim iMinute
	Dim iSecond

	iHour = CLng(Left(sSerialTime, Len("00")))
	If Err.number <> 0 Then
		iHour = Hour(Time())
		Err.Clear()
	End If
	iMinute = CLng(Mid(sSerialTime, Len("000"), Len("00")))
	If Err.number <> 0 Then
		iMinute = Minute(Time())
		Err.Clear()
	End If
	If Len(sSerialTime) >= 6 Then
		iSecond = CLng(Mid(sSerialTime, Len("00000"), Len("00")))
		If Err.number <> 0 Then
			iSecond = Second(Time())
			Err.Clear()
		End If
	End If

	If Len(sSerialTime) >= 6 Then
		DisplayTimeFromSerialNumber = iHour & ":" & Right(("0" & iMinute), Len("00")) & ":" & Right(("0" & iSecond), Len("00"))
	Else
		DisplayTimeFromSerialNumber = iHour & ":" & Right(("0" & iMinute), Len("00"))
	End If
	Err.Clear
End Function

Function DisplayTimeStamp(sMessage)
'************************************************************
'Purpose: To display a HTML comment with the date and time
'Inputs:  sMessage
'Outputs: A string representing a date
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayTimeStamp"
	Dim oDate
	oDate = Now()

	Response.Write vbNewLine & vbNewLine & "<!-- TIME STAMP: " & Year(oDate) & Right(("0" & Month(oDate)), Len("00")) & Right(("0" & Day(oDate)), Len("00")) & Right(("0" & Hour(oDate)), Len("00")) & Right(("0" & Minute(oDate)), Len("00")) & Right(("0" & Second(oDate)), Len("00"))
	If Len(sMessage) > 0 Then Response.Write vbNewLine & vbTab & sMessage & vbNewLine
	Response.Write " -->" & vbNewLine & vbNewLine

	Err.Clear
End Function

Function GetAntiquityFromDays(lDays, lStartDate, iYears, iMonths, iDays)
'************************************************************
'Purpose: To transform a number of days into years, months
'         and days
'Inputs:  lDays, lStartDate
'Outputs: iYears, iMonths, iDays
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetAntiquityFromDays"
	Dim oStartDate
	Dim oEndDate

	If lStartDate = 0 Then
		oStartDate = DateSerial(2000, 3, 1)
	Else
		oStartDate = GetDateFromSerialNumber(lStartDate)
	End If
	oEndDate = DateAdd("d", lDays - 1, oStartDate)
	Call GetAntiquityFromSerialDates(Left(GetSerialNumberForDate(oStartDate), Len("00000000")), Left(GetSerialNumberForDate(oEndDate), Len("00000000")), iYears, iMonths, iDays)

	GetAntiquityFromDays = Err.number
	Err.Clear
End Function

Function GetAntiquityFromSerialDates(lStartDate, lEndDate, iYears, iMonths, iDays)
'************************************************************
'Purpose: To get a number of years, months and days between two dates
'Inputs:  lStartDate, lEndDate
'Outputs: iYears, iMonths, iDays
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetAntiquityFromSerialDates"
	Dim oStartDate
	Dim oEndDate
	Dim asStartDate
	Dim asEndDate
	Dim lTemp
	Dim iIndex

	iYears = 0
	iMonths = 0
	iDays = 0
	If lStartDate = lEndDate Then
		iDays = 1
	Else
		If lEndDate < lStartDate Then
			lTemp = lStartDate
			lStartDate = lEndDate
			lEndDate = lTemp
		End If

		oEndDate = GetDateFromSerialNumber(lEndDate)

		oStartDate = GetDateFromSerialNumber(lStartDate)
		iYears = DateDiff("yyyy", oStartDate, oEndDate)

		oStartDate = DateAdd("yyyy", iYears, oStartDate)
		iMonths = DateDiff("m", oStartDate, oEndDate)
		If iMonths < 0 Then
			iMonths = 12 + iMonths
			iYears = iYears - 1
			oStartDate = DateAdd("yyyy", -1, oStartDate)
		End If

		oStartDate = DateAdd("m", iMonths, oStartDate)
		iDays = DateDiff("d", oStartDate, oEndDate) + 1
		If (iDays = 30) Or (iDays = 31) Then
			iDays = 0
			iMonths = iMonths + 1
		ElseIf iDays < 0 Then
			Select Case CInt(Mid(lStartDate, Len("00000"), Len("00")))
				Case 1, 3, 5, 7, 8, 10, 12
					iDays = 31 + iDays
				Case 2
					If CInt(Left(lStartDate, Len("0000"))) Mod 4 Then
						iDays = 29 + iDays
					Else
						iDays = 28 + iDays
					End If
				Case 4, 6, 9, 11
					iDays = 30 + iDays
			End Select
			iMonths = iMonths - 1
		End If
		If (iMonths = 12) Then
			iMonths = 0
			iYears = iYears + 1
		ElseIf iMonths < 0 Then
			iMonths = 12 + iMonths
			iYears = iYears - 1
		End If
	End If

	GetAntiquityFromSerialDates = Err.number
	Err.Clear
End Function

Function GetCreditsEndDate(lEndYear, lPayrollNumber, lCreditEndDate)
'************************************************************
'Purpose: To get the last day for a payroll for credits
'Inputs:  lStartYear, lPayrollNumber
'Outputs: lCreditEndDate as serial number
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetCreditsEndDate"
	Dim lErrorNumber
	Dim iMonth

	If (lPayrollNumber > 0) And (lPayrollNumber < 25) Then
		iMonth = CInt((CInt(lPayrollNumber) + 0.5) / 2)
		lCreditEndDate = lEndYear & Right("0" & iMonth, Len("MM"))
		If ((CInt(lPayrollNumber) + 0.5) Mod 2) = 0 Then
			lCreditEndDate = CLng(GetPayrollEndDate & "15")
		Else
			Select Case iMonth
				Case 1, 3, 5, 7, 8, 10, 12
					lCreditEndDate = CLng(GetPayrollEndDate & "31")
				Case 4, 6, 9, 11
					lCreditEndDate = CLng(GetPayrollEndDate & "30")
				Case 2
					If (lEndYear Mod 4)  = 0 Then
						lCreditEndDate = CLng(GetPayrollEndDate & "29")
					Else
						lCreditEndDate = CLng(GetPayrollEndDate & "28")
					End If
			End Select
		End If
	Else
		lErrorNumber = -1
		lCreditEndDate = 30000000
	End If

	GetCreditsEndDate = lErrorNumber
	Err.Clear
End Function

Function GetCreditsStartDate(lStartYear, lPayrollNumber, lCreditStartDate)
'************************************************************
'Purpose: To get the last day for a payroll for credits
'Inputs:  lStartYear, lPayrollNumber
'Outputs: A date as serial number
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetCreditsStartDate"
	Dim lErrorNumber
	Dim iMonth

	If (CInt(lPayrollNumber) > 0) And (CInt(lPayrollNumber) < 25) Then
		iMonth = CInt((CInt(lPayrollNumber) + 0.5) / 2)
		lCreditStartDate = lStartYear & Right("0" & iMonth, Len("MM"))
		If ((CInt(lPayrollNumber) + 0.5) Mod 2) = 0 Then
			lCreditStartDate = CLng(lCreditStartDate & "01")
		Else
			lCreditStartDate = CLng(lCreditStartDate & "16")
		End If
	Else
		lErrorNumber = -1
		lCreditStartDate = 0
	End If

	GetCreditsStartDate = lErrorNumber
	Err.Clear
End Function

Function GetDateRank(oRequest, sStartPrefix, sEndPrefix, bLimitEndDate, sRank)
'************************************************************
'Purpose: To get a date rank getting the data from the URL
'Inputs:  oRequest, sStartPrefix, sEndPrefix, bLimitEndDate
'Outputs: sRank
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetDateRank"
	Dim sStartDate
	Dim sEndDate
	Dim sTemp
	Dim bFromRequest
	Dim sID

	sID = oRequest(sStartPrefix & "Year").Item
	bFromRequest = (Err.number = 0)
	Err.clear

	sStartDate = "YYYYMMDD"
	If bFromRequest Then
		If Len(oRequest(sStartPrefix & "Year").Item) > 0 Then
			If CInt(oRequest(sStartPrefix & "Year").Item) > 0 Then
				sStartDate = Replace(sStartDate, "YYYY", Right(("0000" & oRequest(sStartPrefix & "Year").Item), Len("1976")), 1, -1, vbBinaryCompare)
			End If
		End If
		If Len(oRequest(sStartPrefix & "Month").Item) > 0 Then
			If CInt(oRequest(sStartPrefix & "Month").Item) > 0 Then
				sStartDate = Replace(sStartDate, "MM", Right(("0" & oRequest(sStartPrefix & "Month").Item), Len("02")), 1, -1, vbBinaryCompare)
			End If
		End If
		If Len(oRequest(sStartPrefix & "Day").Item) > 0 Then
			If CInt(oRequest(sStartPrefix & "Day").Item) > 0 Then
				sStartDate = Replace(sStartDate, "DD", Right(("0" & oRequest(sStartPrefix & "Day").Item), Len("11")), 1, -1, vbBinaryCompare)
			End If
		End If
	Else
		If Len(GetParameterFromURLString(oRequest, sStartPrefix & "Year")) > 0 Then
			If CInt(GetParameterFromURLString(oRequest, sStartPrefix & "Year")) > 0 Then
				sStartDate = Replace(sStartDate, "YYYY", Right(("0000" & GetParameterFromURLString(oRequest, sStartPrefix & "Year")), Len("1976")), 1, -1, vbBinaryCompare)
			End If
		End If
		If Len(GetParameterFromURLString(oRequest, sStartPrefix & "Month")) > 0 Then
			If CInt(GetParameterFromURLString(oRequest, sStartPrefix & "Month")) > 0 Then
				sStartDate = Replace(sStartDate, "MM", Right(("0" & GetParameterFromURLString(oRequest, sStartPrefix & "Month")), Len("02")), 1, -1, vbBinaryCompare)
			End If
		End If
		If Len(GetParameterFromURLString(oRequest, sStartPrefix & "Day")) > 0 Then
			If CInt(GetParameterFromURLString(oRequest, sStartPrefix & "Day")) > 0 Then
				sStartDate = Replace(sStartDate, "DD", Right(("0" & GetParameterFromURLString(oRequest, sStartPrefix & "Day")), Len("11")), 1, -1, vbBinaryCompare)
			End If
		End If
	End If
	sStartDate = Replace(Replace(Replace(sStartDate, "YYYY", N_FORM_START_YEAR, 1, -1, vbBinaryCompare), "MM", "01", 1, -1, vbBinaryCompare), "DD", "01", 1, -1, vbBinaryCompare)

	sEndDate = "YYYYMMDD"
	If bFromRequest Then
		If Len(oRequest(sEndPrefix & "Year").Item) > 0 Then
			If CInt(oRequest(sEndPrefix & "Year").Item) > 0 Then
				sEndDate = Replace(sEndDate, "YYYY", Right(("0000" & oRequest(sEndPrefix & "Year").Item), Len("1976")), 1, -1, vbBinaryCompare)
			End If
		End If
		If Len(oRequest(sEndPrefix & "Month").Item) > 0 Then
			If CInt(oRequest(sEndPrefix & "Month").Item) > 0 Then
				sEndDate = Replace(sEndDate, "MM", Right(("0" & oRequest(sEndPrefix & "Month").Item), Len("02")), 1, -1, vbBinaryCompare)
			End If
		End If
		If Len(oRequest(sEndPrefix & "Day").Item) > 0 Then
			If CInt(oRequest(sEndPrefix & "Day").Item) > 0 Then
				sEndDate = Replace(sEndDate, "DD", Right(("0" & oRequest(sEndPrefix & "Day").Item), Len("11")), 1, -1, vbBinaryCompare)
			End If
		End If
	Else
		If Len(GetParameterFromURLString(oRequest, sEndPrefix & "Year")) > 0 Then
			If CInt(GetParameterFromURLString(oRequest, sEndPrefix & "Year")) > 0 Then
				sEndDate = Replace(sEndDate, "YYYY", Right(("0000" & GetParameterFromURLString(oRequest, sEndPrefix & "Year")), Len("1976")), 1, -1, vbBinaryCompare)
			End If
		End If
		If Len(GetParameterFromURLString(oRequest, sEndPrefix & "Month")) > 0 Then
			If CInt(GetParameterFromURLString(oRequest, sEndPrefix & "Month")) > 0 Then
				sEndDate = Replace(sEndDate, "MM", Right(("0" & GetParameterFromURLString(oRequest, sEndPrefix & "Month")), Len("02")), 1, -1, vbBinaryCompare)
			End If
		End If
		If Len(GetParameterFromURLString(oRequest, sEndPrefix & "Day")) > 0 Then
			If CInt(GetParameterFromURLString(oRequest, sEndPrefix & "Day")) > 0 Then
				sEndDate = Replace(sEndDate, "DD", Right(("0" & GetParameterFromURLString(oRequest, sEndPrefix & "Day")), Len("11")), 1, -1, vbBinaryCompare)
			End If
		End If
	End If
	If bLimitEndDate Then
		sEndDate = Replace(Replace(Replace(sEndDate, "YYYY", Year(Date()), 1, -1, vbBinaryCompare), "MM", Right("0" & Month(Date()), Len("00")), 1, -1, vbBinaryCompare), "DD", Right("0" & Day(Date()), Len("00")), 1, -1, vbBinaryCompare)
	Else
		sEndDate = Replace(Replace(Replace(sEndDate, "YYYY", "3000", 1, -1, vbBinaryCompare), "MM", "12", 1, -1, vbBinaryCompare), "DD", "31", 1, -1, vbBinaryCompare)
	End If

	If CLng(sStartDate) > CLng(sEndDate) Then
		sTemp = sStartDate
		sStartDate = sEndDate
		sEndDate = sTemp
	End If
	sRank = "Del " & DisplayDate(Left(sStartDate, 4), Mid(sStartDate, 5, 2), Right(sStartDate, 2), -1, -1, -1) & " al " & DisplayDate(Left(sEndDate, 4), Mid(sEndDate, 5, 2), Right(sEndDate, 2), -1, -1, -1) & "."

	GetDateRank = Err.number
	Err.Clear
End Function

Function GetDateFromSerialNumber(sSerialDate)
'************************************************************
'Purpose: To create a string that represents a date using the
'         the correct format.
'Inputs:  sDate
'Outputs: A string representing a date
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetDateFromSerialNumber"

	If Len(sSerialDate) > 0 Then
		GetDateFromSerialNumber = DateSerial(Left(sSerialDate, Len("0000")), Mid(sSerialDate, Len("00000"), Len("00")), Right(sSerialDate, Len("00")))
	End If

	Err.Clear
End Function

Function GetLastDayFromMonth(lDate, iDay)
'************************************************************
'Purpose: To get last day of specific month of year
'Inputs:  lDate
'Outputs: iDay
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetLastDayFromMonth"
	Dim oStartDate
	Dim oEndDate
	Dim asStartDate
	Dim asEndDate
	Dim lTemp
	Dim iIndex

	Select Case CInt(Mid(lDate, Len("00000"), Len("00")))
		Case 1, 3, 5, 7, 8, 10, 12
			iDay = 31
		Case 2
			If (CInt(Left(lDate, Len("0000"))) Mod 4) = 0 Then
				iDay = 29
			Else
				iDay = 28
			End If
		Case 4, 6, 9, 11
			iDay = 30
	End Select

	GetLastDayFromMonth = Err.number
	Err.Clear
End Function

Function GetNextEndDateForVacations(sSerialDate, iJourneyType)
'************************************************************
'Purpose: To convert the serial date and add it the given days
'Inputs:  sSerialDate, iDays
'Outputs: A new serial date
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetNextEndDateForVacations"
	Dim oDate
	Dim iDay
	Dim lDateRevision

	If Len(sSerialDate) = 0 Then
		sSerialDate = Left(GetSerialNumberForDate(""), Len("00000000"))
	End If

	For iDay = 0 To 6
		lDateRevision = AddDaysToSerialDate(sSerialDate, (-1) * iDay)
		oDate = DateSerial(Left(lDateRevision, Len("0000")), Mid(lDateRevision, Len("00000"), Len("00")), Right(lDateRevision, Len("00")))
		Select Case iJourneyType
			Case 1
				If ((Weekday(oDate) <> vbSunday) And (Weekday(oDate) <> vbSaturday) And ( Not IsHoliday(oADODBConnection, lDateRevision, sErrorDescription))) Then
					Exit For
				End If
			Case 21
				If ((Weekday(oDate) <> vbTuesday) And (Weekday(oDate) <> vbThursday) And (Weekday(oDate) <> vbSaturday) And (Weekday(oDate) <> vbSunday)) Then
					Exit For
				End If
			Case 22
				If ((Weekday(oDate) <> vbMonday) And (Weekday(oDate) <> vbWednesday) And (Weekday(oDate) <> vbFriday)) Then
					Exit For
				End If
			Case 23
				If ((Weekday(oDate) = vbSaturday) And (Weekday(oDate) <> vbSunday)) Then
					Exit For
				End If
			Case 31, 32
				If (((Weekday(oDate) <> vbMonday) And (Weekday(oDate) <> vbTuesday) And (Weekday(oDate) <> vbWednesday) And (Weekday(oDate) <> vbThursday) And (Weekday(oDate) <> vbFriday)) Or (IsHoliday(oADODBConnection, lDateRevision, sErrorDescription))) Then
					Exit For
				End If
			Case 41
				If (((Weekday(oDate) <> vbMonday) And (Weekday(oDate) <> vbTuesday) And (Weekday(oDate) <> vbWednesday) And (Weekday(oDate) <> vbThursday) And (Weekday(oDate) <> vbFriday) And (Weekday(oDate) <> vbSunday)) Or (IsHoliday(oADODBConnection, lDateRevision, sErrorDescription))) Then
					Exit For
				End If
			Case 42
				If (((Weekday(oDate) <> vbMonday) And (Weekday(oDate) <> vbTuesday) And (Weekday(oDate) <> vbWednesday) And (Weekday(oDate) <> vbThursday) And (Weekday(oDate) <> vbFriday) And (Weekday(oDate) <> vbSaturday)) Or (IsHoliday(oADODBConnection, lDateRevision, sErrorDescription))) Then
					Exit For
				End If
		End Select
	Next

	GetNextEndDateForVacations = CLng(Year(oDate) & Right(("0" & Month(oDate)), Len("00")) & Right(("0" & Day(oDate)), Len("00")))
	Err.Clear
End Function

Function GetNextStartDateForVacations(sSerialDate, iJourneyType, iDaysAddForVacations)
'************************************************************
'Purpose: To convert the serial date and add it the given days
'Inputs:  sSerialDate, iDays
'Outputs: A new serial date
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetNextStartDateForVacations"
	Dim oDate
	Dim iDay
	Dim lDateRevision
	Dim iToNextDay

	If Len(sSerialDate) = 0 Then
		sSerialDate = Left(GetSerialNumberForDate(""), Len("00000000"))
	End If

	iToNextDay = 6
	For iDay = 0 To iToNextDay
		iDaysAddForVacations = iDaysAddForVacations + 1
		lDateRevision = AddDaysToSerialDate(sSerialDate, iDay)
		oDate = DateSerial(Left(lDateRevision, Len("0000")), Mid(lDateRevision, Len("00000"), Len("00")), Right(lDateRevision, Len("00")))
		Select Case iJourneyType
			Case 1
				If DayIsVacation(aAbsenceComponent, lDateRevision, sErrorDescription) Then
					iToNextDay = iToNextDay + 1
				Else
					If ((Weekday(oDate) <> vbSunday) And (Weekday(oDate) <> vbSaturday) And (Not IsHoliday(oADODBConnection, lDateRevision, sErrorDescription))) Then
						Exit For
					End If
				End If
			Case 21
				If ((Weekday(oDate) <> vbTuesday) And (Weekday(oDate) <> vbThursday) And (Weekday(oDate) <> vbSaturday) And (Weekday(oDate) <> vbSunday)) Then
					Exit For
				End If
			Case 22
				If ((Weekday(oDate) <> vbMonday) And (Weekday(oDate) <> vbWednesday) And (Weekday(oDate) <> vbFriday)) Then
					Exit For
				End If
			Case 23
				If ((Weekday(oDate) = vbSaturday) And (Weekday(oDate) <> vbSunday)) Then
					Exit For
				End If
			Case 31, 32
				If (((Weekday(oDate) <> vbMonday) And (Weekday(oDate) <> vbTuesday) And (Weekday(oDate) <> vbWednesday) And (Weekday(oDate) <> vbThursday) And (Weekday(oDate) <> vbFriday)) Or (IsHoliday(oADODBConnection, lDateRevision, sErrorDescription))) Then
					Exit For
				End If
			Case 41
				If (((Weekday(oDate) <> vbMonday) And (Weekday(oDate) <> vbTuesday) And (Weekday(oDate) <> vbWednesday) And (Weekday(oDate) <> vbThursday) And (Weekday(oDate) <> vbFriday) And (Weekday(oDate) <> vbSunday)) Or (IsHoliday(oADODBConnection, lDateRevision, sErrorDescription))) Then
					Exit For
				End If
			Case 42
				If (((Weekday(oDate) <> vbMonday) And (Weekday(oDate) <> vbTuesday) And (Weekday(oDate) <> vbWednesday) And (Weekday(oDate) <> vbThursday) And (Weekday(oDate) <> vbFriday) And (Weekday(oDate) <> vbSaturday)) Or (IsHoliday(oADODBConnection, lDateRevision, sErrorDescription))) Then
					Exit For
				End If
		End Select
	Next

	iDaysAddForVacations = iDaysAddForVacations - 1
	GetNextStartDateForVacations = CLng(Year(oDate) & Right(("0" & Month(oDate)), Len("00")) & Right(("0" & Day(oDate)), Len("00")))
	Err.Clear
End Function

Function GetPayrollStartDate(lPayrollID)
'************************************************************
'Purpose: To get the first day for a payroll
'Inputs:  lPayrollID
'Outputs: A date as serial number
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPayrollStartDate"
	Dim iDay
	Dim lTemp

	lTemp = lPayrollID
	If Len(lTemp) >= Len("00000000") Then
		lTemp = Left(lTemp, Len("00000000"))
		iDay = Int(Right(lTemp, Len("00")))
		If iDay < 16 Then
			GetPayrollStartDate = CLng(Left(lTemp, Len("000000")) & "01")
		Else
			GetPayrollStartDate = CLng(Left(lTemp, Len("000000")) & "16")
		End If
	Else
		GetPayrollStartDate = 0
	End If

	Err.Clear
End Function

Function GetPayrollEndDate(lPayrollID)
'************************************************************
'Purpose: To get the last day for a payroll
'Inputs:  lPayrollID
'Outputs: A date as serial number
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPayrollEndDate"
	Dim iYear
	Dim iMonth
	Dim iDay
	Dim lTemp

	lTemp = lPayrollID
	If Len(lTemp) >= Len("00000000") Then
		lTemp = Left(lTemp, Len("00000000"))
		iYear = Int(Left(lTemp, Len("YYYY")))
		iMonth = Int(Mid(lTemp, Len("YYYYM"), Len("MM")))
		iDay = Int(Right(lTemp, Len("DD")))
		GetPayrollEndDate = Left(lTemp, Len("YYYYMM"))
		If iDay < 16 Then
			GetPayrollEndDate = CLng(GetPayrollEndDate & "15")
		Else
			Select Case iMonth
				Case 1, 3, 5, 7, 8, 10, 12
					GetPayrollEndDate = CLng(GetPayrollEndDate & "31")
				Case 4, 6, 9, 11
					GetPayrollEndDate = CLng(GetPayrollEndDate & "30")
				Case 2
					If (iYear Mod 4)  = 0 Then
						GetPayrollEndDate = CLng(GetPayrollEndDate & "29")
					Else
						GetPayrollEndDate = CLng(GetPayrollEndDate & "28")
					End If
			End Select
		End If
	Else
		GetPayrollEndDate = 0
	End If

	Err.Clear
End Function

Function GetSerialNumberForDate(sDate)
'************************************************************
'Purpose: To create a string that represents a date using a
'         long integer.
'Inputs:  sDate
'Outputs: A Long integer representing a date
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetSerialNumberForDate"

	If Len(sDate) = 0 Then
		sDate = Now()
	End If

	GetSerialNumberForDate = Year(sDate) & Right(("0" & Month(sDate)), Len("00")) & Right(("0" & Day(sDate)), Len("00")) & Right(("0" & Hour(sDate)), Len("00")) & Right(("0" & Minute(sDate)), Len("00")) & Right(("0" & Second(sDate)), Len("00"))
	Err.Clear
End Function

Function GetSerialNumberFromURL(sPrefixName)
'************************************************************
'Purpose: To create a string that represents a date using the
'         URL
'Inputs:  sPrefixName
'Outputs: A Long integer representing a date
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetSerialNumberFromURL"
	Dim sYear
	Dim sMonth
	Dim sDay

	If Len(oRequest(sPrefixName & "Year").Item) = 0 Then
		sYear = Year(Date())
	Else
		sYear = oRequest(sPrefixName & "Year").Item
	End If
	If Len(oRequest(sPrefixName & "Month").Item) = 0 Then
		sMonth = Right("0" & Month(Date()), Len("00"))
	Else
		sMonth = Right("0" & oRequest(sPrefixName & "Month").Item, Len("00"))
	End If
	If Len(oRequest(sPrefixName & "Day").Item) = 0 Then
		sDay = Right("0" & Day(Date()), Len("00"))
	Else
		sDay = Right("0" & oRequest(sPrefixName & "Day").Item, Len("00"))
	End If

	GetSerialNumberFromURL = sYear & sMonth & sDay
	Err.Clear
End Function

Function GetStartAndEndDatesFromURL(sStartPrefix, sEndPrefix, sQueryColumn, bLimitEndDate, sCondition)
'************************************************************
'Purpose: To build a condition string getting the start and
'         end date from the URL
'Inputs:  sStartPrefix, sEndPrefix, sQueryColumn, bLimitEndDate
'Outputs: sCondition
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetStartAndEndDatesFromURL"
	Dim sStartDate
	Dim sEndDate
	Dim sTemp

	If (Len(oRequest(sStartPrefix & "Year").Item) > 0) Or (Len(oRequest(sStartPrefix & "Month").Item) > 0) Or (Len(oRequest(sStartPrefix & "Day").Item) > 0) Then
		sStartDate = "YYYYMMDD"
		If Len(oRequest(sStartPrefix & "Year").Item) > 0 Then
			If CInt(oRequest(sStartPrefix & "Year").Item) > 0 Then
				sStartDate = Replace(sStartDate, "YYYY", Right(("0000" & oRequest(sStartPrefix & "Year").Item), Len("1976")), 1, -1, vbBinaryCompare)
			End If
		End If
		If Len(oRequest(sStartPrefix & "Month").Item) > 0 Then
			If CInt(oRequest(sStartPrefix & "Month").Item) > 0 Then
				sStartDate = Replace(sStartDate, "MM", Right(("0" & oRequest(sStartPrefix & "Month").Item), Len("02")), 1, -1, vbBinaryCompare)
			End If
		End If
		If Len(oRequest(sStartPrefix & "Day").Item) > 0 Then
			If CInt(oRequest(sStartPrefix & "Day").Item) > 0 Then
				sStartDate = Replace(sStartDate, "DD", Right(("0" & oRequest(sStartPrefix & "Day").Item), Len("11")), 1, -1, vbBinaryCompare)
			End If
		End If
		sStartDate = Replace(Replace(Replace(sStartDate, "YYYY", "0000", 1, -1, vbBinaryCompare), "MM", "00", 1, -1, vbBinaryCompare), "DD", "00", 1, -1, vbBinaryCompare)
	End If

	If (Len(oRequest(sEndPrefix & "Year").Item) > 0) Or (Len(oRequest(sEndPrefix & "Month").Item) > 0) Or (Len(oRequest(sEndPrefix & "Day").Item) > 0) Then
		sEndDate = "YYYYMMDD"
		If Len(oRequest(sEndPrefix & "Year").Item) > 0 Then
			If CInt(oRequest(sEndPrefix & "Year").Item) > 0 Then
				sEndDate = Replace(sEndDate, "YYYY", Right(("0000" & oRequest(sEndPrefix & "Year").Item), Len("1976")), 1, -1, vbBinaryCompare)
			End If
		End If
		If Len(oRequest(sEndPrefix & "Month").Item) > 0 Then
			If CInt(oRequest(sEndPrefix & "Month").Item) > 0 Then
				sEndDate = Replace(sEndDate, "MM", Right(("0" & oRequest(sEndPrefix & "Month").Item), Len("02")), 1, -1, vbBinaryCompare)
			End If
		End If
		If Len(oRequest(sEndPrefix & "Day").Item) > 0 Then
			If CInt(oRequest(sEndPrefix & "Day").Item) > 0 Then
				sEndDate = Replace(sEndDate, "DD", Right(("0" & oRequest(sEndPrefix & "Day").Item), Len("11")), 1, -1, vbBinaryCompare)
			End If
		End If
		If bLimitEndDate Then
			If InStr(1, sEndDate, Year(Date()), vbBinaryCompare) = 1 Then
				sEndDate = Replace(Replace(Replace(sEndDate, "YYYY", Year(Date()), 1, -1, vbBinaryCompare), "MM", Right("0" & Month(Date()), Len("00")), 1, -1, vbBinaryCompare), "DD", Right("0" & Day(Date()), Len("00")), 1, -1, vbBinaryCompare)
			Else
				sEndDate = Replace(Replace(Replace(sEndDate, "YYYY", Year(Date()), 1, -1, vbBinaryCompare), "MM", "12", 1, -1, vbBinaryCompare), "DD", "31", 1, -1, vbBinaryCompare)
			End If
		Else
			sEndDate = Replace(Replace(Replace(sEndDate, "YYYY", "3000", 1, -1, vbBinaryCompare), "MM", "12", 1, -1, vbBinaryCompare), "DD", "31", 1, -1, vbBinaryCompare)
		End If
	End If
	If IsEmpty(sStartDate) Then sStartDate = 0
	If IsEmpty(sEndDate) Then sEndDate = 30000000
	If CLng(sStartDate) > CLng(sEndDate) Then
		sTemp = sStartDate
		sStartDate = sEndDate
		sEndDate = sTemp
	End If
	If (Len(oRequest("ReportID").Item) > 0) Then
		If (CLng(oRequest("ReportID").Item) = 711) Then
			If CLng(sStartDate) > 0 Then sCondition = sCondition & " And (JobsHistoryList.JobDate >= " & sStartDate & ")"
			If CLng(sEndDate) < 30000000 Then sCondition = sCondition & " And (JobsHistoryList.EndDate <= " & sEndDate & ")"
		Else
			If CLng(sStartDate) > 0 Then sCondition = sCondition & " And (" & sQueryColumn & " >= " & sStartDate & ")"
			If CLng(sEndDate) < 30000000 Then sCondition = sCondition & " And (" & sQueryColumn & " <= " & sEndDate & ")"
		End If
	Else
		If CLng(sStartDate) > 0 Then sCondition = sCondition & " And (" & sQueryColumn & " >= " & sStartDate & ")"
		If CLng(sEndDate) < 30000000 Then sCondition = sCondition & " And (" & sQueryColumn & " <= " & sEndDate & ")"
	End If

	GetStartAndEndDatesFromURL = Err.number
	Err.Clear
End Function

Function GetStartAndEndDatesFromURLAsNumbers(sStartPrefix, sEndPrefix, bLimitEndDate, lStartDate, lEndDate)
'************************************************************
'Purpose: To get the the start and end date from the URL as numbers
'Inputs:  sStartPrefix, sEndPrefix, bLimitEndDate
'Outputs: lStartDate, lEndDate
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetStartAndEndDatesFromURLAsNumbers"
	Dim lTemp

	If (Len(oRequest(sStartPrefix & "Year").Item) > 0) Or (Len(oRequest(sStartPrefix & "Month").Item) > 0) Or (Len(oRequest(sStartPrefix & "Day").Item) > 0) Then
		lStartDate = "YYYYMMDD"
		If Len(oRequest(sStartPrefix & "Year").Item) > 0 Then
			If CInt(oRequest(sStartPrefix & "Year").Item) > 0 Then
				lStartDate = Replace(lStartDate, "YYYY", Right(("0000" & oRequest(sStartPrefix & "Year").Item), Len("1976")), 1, -1, vbBinaryCompare)
			End If
		End If
		If Len(oRequest(sStartPrefix & "Month").Item) > 0 Then
			If CInt(oRequest(sStartPrefix & "Month").Item) > 0 Then
				lStartDate = Replace(lStartDate, "MM", Right(("0" & oRequest(sStartPrefix & "Month").Item), Len("02")), 1, -1, vbBinaryCompare)
			End If
		End If
		If Len(oRequest(sStartPrefix & "Day").Item) > 0 Then
			If CInt(oRequest(sStartPrefix & "Day").Item) > 0 Then
				lStartDate = Replace(lStartDate, "DD", Right(("0" & oRequest(sStartPrefix & "Day").Item), Len("11")), 1, -1, vbBinaryCompare)
			End If
		End If
		lStartDate = Replace(Replace(Replace(lStartDate, "YYYY", "0000", 1, -1, vbBinaryCompare), "MM", "00", 1, -1, vbBinaryCompare), "DD", "00", 1, -1, vbBinaryCompare)
	End If

	If (Len(oRequest(sEndPrefix & "Year").Item) > 0) Or (Len(oRequest(sEndPrefix & "Month").Item) > 0) Or (Len(oRequest(sEndPrefix & "Day").Item) > 0) Then
		lEndDate = "YYYYMMDD"
		If Len(oRequest(sEndPrefix & "Year").Item) > 0 Then
			If CInt(oRequest(sEndPrefix & "Year").Item) > 0 Then
				lEndDate = Replace(lEndDate, "YYYY", Right(("0000" & oRequest(sEndPrefix & "Year").Item), Len("1976")), 1, -1, vbBinaryCompare)
			End If
		End If
		If Len(oRequest(sEndPrefix & "Month").Item) > 0 Then
			If CInt(oRequest(sEndPrefix & "Month").Item) > 0 Then
				lEndDate = Replace(lEndDate, "MM", Right(("0" & oRequest(sEndPrefix & "Month").Item), Len("02")), 1, -1, vbBinaryCompare)
			End If
		End If
		If Len(oRequest(sEndPrefix & "Day").Item) > 0 Then
			If CInt(oRequest(sEndPrefix & "Day").Item) > 0 Then
				lEndDate = Replace(lEndDate, "DD", Right(("0" & oRequest(sEndPrefix & "Day").Item), Len("11")), 1, -1, vbBinaryCompare)
			End If
		End If
		If bLimitEndDate Then
			lEndDate = Replace(Replace(Replace(lEndDate, "YYYY", Year(Date()), 1, -1, vbBinaryCompare), "MM", Right("0" & Month(Date()), Len("00")), 1, -1, vbBinaryCompare), "DD", Right("0" & Day(Date()), Len("00")), 1, -1, vbBinaryCompare)
		Else
			lEndDate = Replace(Replace(Replace(lEndDate, "YYYY", "3000", 1, -1, vbBinaryCompare), "MM", "12", 1, -1, vbBinaryCompare), "DD", "31", 1, -1, vbBinaryCompare)
		End If
	End If
	If IsEmpty(lStartDate) Then
		lStartDate = 0
	Else
		lStartDate = CLng(lStartDate)
	End If
	If IsEmpty(lEndDate) Then
		lEndDate = 30000000
	Else
		lEndDate = CLng(lEndDate)
	End If
	If lStartDate > lEndDate Then
		lTemp = lStartDate
		lStartDate = lEndDate
		lEndDate = lTemp
	End If

	GetStartAndEndDatesFromURLAsNumbers = Err.number
	Err.Clear
End Function

Function GetSundaysForPayroll(lDate)
'************************************************************
'Purpose: To count how many Sundays are in a payroll given
'         a date
'Inputs:  lDate
'Outputs: The number of Sundays in a payroll
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetSundaysForPayroll"
	Dim oStartDate
	Dim oEndDate
	Dim iIndex

	oStartDate = GetDateFromSerialNumber(GetPayrollStartDate(lDate))
	oEndDate = GetDateFromSerialNumber(GetPayrollEndDate(lDate))
	GetSundaysForPayroll = 0
	Do While DateDiff("d", oStartDate, oEndDate) >= 0
		If Weekday(oStartDate) = 1 Then GetSundaysForPayroll = GetSundaysForPayroll + 1
		oStartDate = DateAdd("d", 1, oStartDate)
	Loop

	Err.Clear
End Function

Function GetTotalPayrolls(StartDate, RecordDate)
'************************************************************
'Purpose: To calculate the total number of payments 
'         between two dates
'Inputs:  StartDate, RecordDate
'Outputs: Total of payrolls between two dates
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetTotalPayrolls"
	Dim lTotalPayrolls
	Dim lTotalDays
	Dim lTotalMonths
	Dim lTotalYears

	lTotalPayrolls = CStr(RecordDate - StartDate)

	If Len(lTotalPayrolls) <= 2 Then
		lTotalDays = Right(lTotalPayrolls, 2)
		GetTotalPayrolls = CInt(lTotalDays / 15)
	ElseIf Len(TotalPayrolls) <= 4 Then
		lTotalDays = Right(lTotalPayrolls, 2)
		lTotalMonths = lTotalDays - lTotalDays
		GetTotalPayrolls = lTotalMonths * 2 + CInt(lTotalDays / 15)
	ElseIf Len(TotalPayrolls) <= 6 Then
		lTotalDays = Right(lTotalPayrolls, 2)
		lNumberMonth = Mid(lTotalPayrolls, 3, 2)
		lTotalYears = (CDbl(lTotalPayrolls) - (CDbl(lTotalMonths) * 100 + CDbl(lTotalDays))) / 10000
		GetTotalPayrolls = lTotalYears * 24 + lTotalMonths * 2 + CInt(lTotalDays / 15)
	Else
		GetTotalPayrolls = 0
	End If

	Err.Clear
End Function

Function GetYearsMonthsDays(lTotalDays)
'************************************************************
'Purpose: To transform the given number of days in years,
'         months and days
'Inputs:  lTotalDays
'Outputs: A string with the years, months and days
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetYearsMonthsDays"
	Dim lYears
	Dim lMonths
	Dim lDays

	lYears = Int(lTotalDays / 365)
	lDays = lTotalDays Mod 365
	lMonths = Int(lDays / 30)
	lDays = lDays Mod 30
	GetYearsMonthsDays = "Años: " & lYears & ", meses: " & lMonths & ", días: " & lDays
	Err.Clear
End Function

Function GetWeekStartDate(lDate)
'************************************************************
'Purpose: To convert the serial dates and get the working days
'         of the period
'Inputs:  sSerialDate, iDays
'Outputs: A new serial date
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetWeekStartDate"
	Dim oDate
	Dim iDays

	oDate = GetDateFromSerialNumber(lDate)
	iDays = Weekday(oDate, vbMonday) - 1
	GetWeekStartDate = CLng(Left(GetSerialNumberForDate(DateAdd("d", -iDays, oDate)), Len("00000000")))

	Err.Clear
End Function

Function GetWorkingDaysOfAbsencesPeriod(lStartDate, lEndDate, iJourneyType)
'************************************************************
'Purpose: To convert the serial dates and get the working days
'         of the period
'Inputs:  sSerialDate, iDays
'Outputs: A new serial date
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetWorkingDaysOfAbsencesPeriod"
	Dim oDate
	Dim iDay
	Dim iDaysCount
	Dim lDateRevision
	Dim iHolidayDaysCount
	Dim dStartDate
	Dim dEndDate

	iHolidayDaysCount = 0
	dStartDate = GetDateFromSerialNumber(lStartDate)
	dEndDate = GetDateFromSerialNumber(lEndDate)

	iDaysCount = DateDiff("d", dStartDate, dEndDate)
	For iDay = 0 To iDaysCount
		lDateRevision = AddDaysToSerialDate(lStartDate, iDay)
		oDate = DateSerial(Left(lDateRevision, Len("0000")), Mid(lDateRevision, Len("00000"), Len("00")), Right(lDateRevision, Len("00")))
		Select Case iJourneyType
			Case 1
				If ((Weekday(oDate) = vbSunday) Or (Weekday(oDate) = vbSaturday) Or (IsHoliday(oADODBConnection, lDateRevision, sErrorDescription))) Then
					iHolidayDaysCount = iHolidayDaysCount + 1
				End If
			Case 21
				If ((Weekday(oDate) = vbTuesday) Or (Weekday(oDate) = vbThursday) Or (Weekday(oDate) = vbSaturday) Or (Weekday(oDate) = vbSunday)) Then
					iHolidayDaysCount = iHolidayDaysCount + 1
				End If
			Case 22
				If ((Weekday(oDate) = vbMonday) Or (Weekday(oDate) = vbWednesday) Or (Weekday(oDate) = vbFriday)) Then
					iHolidayDaysCount = iHolidayDaysCount + 1
				End If
			Case 23
				If ((Weekday(oDate) = vbSaturday) Or (Weekday(oDate) = vbSunday)) Then
					iHolidayDaysCount = iHolidayDaysCount + 1
				End If
			Case 31, 32
				If (((Weekday(oDate) = vbMonday) Or (Weekday(oDate) = vbTuesday) Or (Weekday(oDate) = vbWednesday) Or (Weekday(oDate) = vbThursday) Or (Weekday(oDate) = vbFriday)) And (Not IsHoliday(oADODBConnection, lDateRevision, sErrorDescription))) Then
					iHolidayDaysCount = iHolidayDaysCount + 1
				End If
			Case 41
				If (((Weekday(oDate) = vbMonday) Or (Weekday(oDate) = vbTuesday) Or (Weekday(oDate) = vbWednesday) Or (Weekday(oDate) = vbThursday) Or (Weekday(oDate) = vbFriday) Or (Weekday(oDate) = vbSunday)) And (Not IsHoliday(oADODBConnection, lDateRevision, sErrorDescription))) Then
					iHolidayDaysCount = iHolidayDaysCount + 1
				End If
			Case 42
				If (((Weekday(oDate) = vbMonday) Or (Weekday(oDate) = vbTuesday) Or (Weekday(oDate) = vbWednesday) Or (Weekday(oDate) = vbThursday) Or (Weekday(oDate) = vbFriday) Or (Weekday(oDate) = vbSaturday)) And (Not IsHoliday(oADODBConnection, lDateRevision, sErrorDescription))) Then
					iHolidayDaysCount = iHolidayDaysCount + 1
				End If
		End Select
	Next

	GetWorkingDaysOfAbsencesPeriod = DateDiff("d", dStartDate, dEndDate) - iHolidayDaysCount + 1
	Err.Clear
End Function

Function IsHoliday(oADODBConnection, lDate, sErrorDescription)
'************************************************************
'Purpose: To verify if the given date is a holiday
'Inputs:  oADODBConnection, lDate
'Outputs: sErrorDescription. True if it's a holiday, otherwise False
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "IsHoliday"
	Dim lErrorNumber
	Dim oRecordset

	If Len(lDate) = 0 Then
		IsHoliday = False
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Holidays Where (Holiday=" & lDate & ")", "DateLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If (lErrorNumber = 0) Then
			IsHoliday = (Not oRecordset.EOF)
			oRecordset.Close
		Else
			sErrorDescription = "Error al verificar si la fecha es día festivo."
			IsHoliday = False
		End If
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function IsSunday(lDate)
'************************************************************
'Purpose: To verify if the given date is Sunday
'Inputs:  lDate
'Outputs: True if it's Sunday, otherwise False
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "IsSunday"
	Dim oDate

	If Len(lDate) = 0 Then
		IsSunday = False
	Else
		oDate = GetDateFromSerialNumber(lDate)
		If Weekday(oDate) = vbSunday Then
			IsSunday = True
		Else
			IsSunday = False
		End If
	End If

	Err.Clear
End Function

Function VerifyIfDateIsCorrect(lDate)
'************************************************************
'Purpose: To verify if the given date is Sunday
'Inputs:  lDate
'Outputs: True if it's Sunday, otherwise False
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyIfDateIsCorrect"
	Dim oDate

	If Len(lDate) = 0 Then
		VerifyIfDateIsCorrect = False
	Else
		oDate = GetDateFromSerialNumber(lDate)
		If (Err.number = 0) Then
			VerifyIfDateIsCorrect = True
		Else
			VerifyIfDateIsCorrect = False
		End If
	End If

	Err.Clear
End Function
%>