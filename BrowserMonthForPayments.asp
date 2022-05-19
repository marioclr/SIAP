<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
%>
<!-- #include file="Libraries/GlobalVariables.asp" -->
<!-- #include file="Libraries/CalendarComponent.asp" -->
<!-- #include file="Libraries/AbsenceComponent.asp" -->
<%
Dim sURLForCalendar
Dim iYear
Dim iMonth
Dim iDay
Dim sFilesNames
Dim sFormName
Dim sDateCombo
Dim sSpecialSelection
Dim sAbsenceIDs
Dim sCaseOptions
Dim iIndex
Dim lReasonID
Dim lJourneyID

sFormName = "AbsencesFrm"
sSpecialSelection = ""
If Len(oRequest("FormName").Item) > 0 Then
	sFormName = oRequest("FormName").Item
End If
If Len(oRequest("OnlySundays").Item) > 0 Then
	sSpecialSelection = "&OnlySundays=1"
ElseIf Len(oRequest("OnlyHolidays").Item) > 0 Then
	sSpecialSelection = "&OnlyHolidays=1"
End If
If Len(oRequest("ReasonID").Item) > 0 Then lReasonID = CLng(oRequest("ReasonID").Item)
If Len(oRequest("JourneyID").Item) > 0 Then 
	lJourneyID = CLng(oRequest("JourneyID").Item)
	sSpecialSelection = "&JourneyID=" & lJourneyID
Else
	lJourneyID = -1
End If
sDateCombo = oRequest("DateCombo").Item

If Len(oRequest("EmployeeDate").Item) > 0 Then
	aCalendarComponent(N_YEAR_CALENDAR) = CInt(Left(oRequest("EmployeeDate").Item, Len("YYYY")))
	aCalendarComponent(N_MONTH_CALENDAR) = CInt(Mid(oRequest("EmployeeDate").Item, Len("YYYYM"), Len("MM")))
	aCalendarComponent(N_DAY_CALENDAR) = CInt(Mid(oRequest("EmployeeDate").Item, Len("YYYYMMD"), Len("DD")))
ElseIf Len(oRequest("OcurredDate").Item) > 0 Then
	aCalendarComponent(N_YEAR_CALENDAR) = CInt(Left(oRequest("OcurredDate").Item, Len("YYYY")))
	aCalendarComponent(N_MONTH_CALENDAR) = CInt(Mid(oRequest("OcurredDate").Item, Len("YYYYM"), Len("MM")))
	aCalendarComponent(N_DAY_CALENDAR) = CInt(Mid(oRequest("OcurredDate").Item, Len("YYYYMMD"), Len("DD")))
End If
Call InitializeCalendarComponent(oRequest, aCalendarComponent)
Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
If InStr(1,sFormName, "AbsencesFrm", vbTextCompare) > 0 Then
	Call GetAbsencesDates(oRequest, oADODBConnection, aAbsenceComponent, lReasonID, aCalendarComponent(S_MARKED_DAYS_CALENDAR), sErrorDescription)
ElseIf InStr(1,sFormName, "EmployeeFrm", vbTextCompare) > 0 Then
	aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = aEmployeeComponent(N_ID_EMPLOYEE)
	Call GetAbsencesDates(oRequest, oADODBConnection, aAbsenceComponent, lReasonID, aCalendarComponent(S_MARKED_DAYS_CALENDAR), sErrorDescription)
Else
	Call GetHolidayDates(oRequest, oADODBConnection, aCalendarComponent, aCalendarComponent(S_MARKED_DAYS_CALENDAR), sErrorDescription)
End If
aCalendarComponent(S_TARGET_PAGE_CALENDAR) = "BrowserMonthForPayments.asp?FormName=" & sFormName & "&EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & sSpecialSelection
%>
<HTML>
	<HEAD>
		<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1" />
		<TITLE>Sistema Integral de Administración de Personal del ISSSTE. Calendario por Meses</TITLE>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/PopupItem.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/HTMLLists.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript"><!--		
			if (window.parent) {
				if (window.parent.document.<%Response.Write sFormName%>.OcurredDate) {
					window.parent.document.<%Response.Write sFormName%>.OcurredDate.value='<%Response.Write aCalendarComponent(N_YEAR_CALENDAR) & Right(("0" & aCalendarComponent(N_MONTH_CALENDAR)), Len("00")) & Right(("0" & aCalendarComponent(N_DAY_CALENDAR)), Len("00"))%>';
					<%Response.Write "var lJourneyID=" & lJourneyID & ";" & vbNewLine
					If B_ISSSTE And (Len(oRequest("FromArrow").Item) = 0) Then
						Response.Write "if (window.parent.document." & sFormName & ".AbsenceID) {" & vbNewLine
							Response.Write "if (IsAbsencesForPeriod(window.parent.document." & sFormName & ".AbsenceID.value) && (lJourneyID!=21 && lJourneyID!=22 && lJourneyID!=23)) {" & vbNewLine
								Response.Write "if (IsAttendanceControlUndefined(window.parent.document." & sFormName & ".AbsenceID.value) && (window.parent.document." & sFormName & ".OcurredDates.length > 0)) {" & vbNewLine
									Response.Write "alert('Para este tipo de incidencia solamente se requiere la fecha de inicio ya que el registro quedara por tiempo indefinido.');" & vbNewLine
								Response.Write "} else { " & vbNewLine
									Response.Write "if (window.parent.document." & sFormName & ".OcurredDates.length < 2) {" & vbNewLine
										Response.Write "AddItemToList('" & DisplayDate(aCalendarComponent(N_YEAR_CALENDAR), aCalendarComponent(N_MONTH_CALENDAR), aCalendarComponent(N_DAY_CALENDAR), -1, -1, -1) & "', '" & aCalendarComponent(N_YEAR_CALENDAR) & Right(("0" & aCalendarComponent(N_MONTH_CALENDAR)), Len("00")) & Right(("0" & aCalendarComponent(N_DAY_CALENDAR)), Len("00")) & "', null, window.parent.document." & sFormName & ".OcurredDates);"  & vbNewLine
									Response.Write "} else { " & vbNewLine
										Response.Write "alert('Este tipo de incidencia no permite capturas multiples. Solo por rango de fecha de inicio y final.');" & vbNewLine
									Response.Write "}"  & vbNewLine
								Response.Write "}"  & vbNewLine
							Response.Write "} else { " & vbNewLine
								'Response.Write "alert('IsAbsencesForPeriod regresa false');" & vbNewLine
								Response.Write "AddItemToList('" & DisplayDate(aCalendarComponent(N_YEAR_CALENDAR), aCalendarComponent(N_MONTH_CALENDAR), aCalendarComponent(N_DAY_CALENDAR), -1, -1, -1) & "', '" & aCalendarComponent(N_YEAR_CALENDAR) & Right(("0" & aCalendarComponent(N_MONTH_CALENDAR)), Len("00")) & Right(("0" & aCalendarComponent(N_DAY_CALENDAR)), Len("00")) & "', null, window.parent.document." & sFormName & ".OcurredDates);"  & vbNewLine
							Response.Write "}"  & vbNewLine
						Response.Write "} else { " & vbNewLine
							Response.Write "AddItemToList('" & DisplayDate(aCalendarComponent(N_YEAR_CALENDAR), aCalendarComponent(N_MONTH_CALENDAR), aCalendarComponent(N_DAY_CALENDAR), -1, -1, -1) & "', '" & aCalendarComponent(N_YEAR_CALENDAR) & Right(("0" & aCalendarComponent(N_MONTH_CALENDAR)), Len("00")) & Right(("0" & aCalendarComponent(N_DAY_CALENDAR)), Len("00")) & "', null, window.parent.document." & sFormName & ".OcurredDates);"  & vbNewLine
						Response.Write "}"  & vbNewLine
					Else
						Response.Write "window.parent.document." & sFormName & ".AbsenceDate.value='" & DisplayDate(aCalendarComponent(N_YEAR_CALENDAR), aCalendarComponent(N_MONTH_CALENDAR), aCalendarComponent(N_DAY_CALENDAR), -1, -1, -1) & "';"
					End If%>
				}
			}
			<%Response.Write "function IsAbsencesForPeriod(sValue) {" & vbNewLine
				If B_ISSSTE Then
					lErrorNumber = GetAbsenceIDsForPeriod(sAbsenceIDs, sErrorDescription)
					If (lErrorNumber = L_ERR_NO_RECORDS) Then
						Response.Write "return false;" & vbNewLine
						sErrorDescription = ""
						lErrorNumber = 0
					Else
						sCaseOptions = Split(sAbsenceIDs, "," , -1, vbBinaryCompare)
						Response.Write "switch (sValue) {" & vbNewLine
							For iIndex = 0 To UBound(sCaseOptions)
								Response.Write "case '" & CInt(sCaseOptions(iIndex)) & "':" & vbNewLine
									Response.Write "return true;" & vbNewLine
									Response.Write "break;" & vbNewLine
							Next
							Response.Write "default:" & vbNewLine
								Response.Write "return false;" & vbNewLine
								Response.Write "break;" & vbNewLine
						Response.Write "}" & vbNewLine
					End If
				End If
			Response.Write "} // End of IsAbsencesForPeriod" & vbNewLine
			Response.Write "function IsAttendanceControlUndefined(sValue) {" & vbNewLine
				Response.Write "switch (sValue) {" & vbNewLine
					Response.Write "case '50':" & vbNewLine
					Response.Write "case '51':" & vbNewLine
					Response.Write "case '54':" & vbNewLine
					Response.Write "case '55':" & vbNewLine
					Response.Write "case '56':" & vbNewLine
						Response.Write "return true;" & vbNewLine
						Response.Write "break;" & vbNewLine
					Response.Write "default:" & vbNewLine
						Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of IsAttendanceControlUndefined" & vbNewLine
			%>
		//--></SCRIPT>
		<LINK REL="STYLESHEET" TYPE="text/css" HREF="Styles/SIAP.css" />
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<TABLE WIDTH="1" BORDER="0" CELLPADDING="0" CELLSPACING="0"><TR>
			<%iYear = aCalendarComponent(N_YEAR_CALENDAR)
			iMonth = aCalendarComponent(N_MONTH_CALENDAR) - 1
			iDay = aCalendarComponent(N_DAY_CALENDAR)
			If iMonth <= 0 Then
				iMonth = iMonth + 12
				iYear = iYear - 1
			End If
			If (InStr(1, (",4,6,9,11,"), ("," & iMonth & ","), vbBinaryCompare) > 0) And (iDay > 30) Then
				iDay = 30
			ElseIf (iMonth = 2) And ((iYear Mod 4) = 0) And (iDay > 29) Then
				iDay = 29
			ElseIf (iMonth = 2) And (iDay > 28) Then
				iDay = 28
			End If
			sURLForCalendar = RemoveParameterFromURLString(ReplaceValueInURLString(ReplaceValueInURLString(ReplaceValueInURLString(ReplaceValueInURLString(oRequest, "Year", iYear), "Month", iMonth), "Day", iDay), "JourneyID", lJourneyID), "EmployeeDate")
			Response.Write "<TD VALIGN=""TOP""><A HREF=""" & GetASPFileName("") & "?" & sURLForCalendar & "&FromArrow=1""><IMG SRC=""Images/ArrLeftBlack.gif"" WIDTH=""7"" HEIGHT=""13"" BORDER=""0"" ALT=""Mes anterior"" /></A>&nbsp;</TD>"
			Response.Write "<TD VALIGN=""TOP"">"
				lErrorNumber = DisplayMonth(oRequest, aCalendarComponent, sErrorDescription)
				If Len(sErrorDescription) > 0 Then
					lErrorNumber = DisplayErrorMessage("Error", sErrorDescription)
				End If
			Response.Write "</TD>"
			iYear = aCalendarComponent(N_YEAR_CALENDAR)
			iMonth = aCalendarComponent(N_MONTH_CALENDAR) + 1
			iDay = aCalendarComponent(N_DAY_CALENDAR)
			If iMonth > 12 Then
				iMonth = iMonth - 12
				iYear = iYear + 1
			End If
			If (InStr(1, (",4,6,9,11,"), ("," & iMonth & ","), vbBinaryCompare) > 0) And (iDay > 30) Then
				iDay = 30
			ElseIf (iMonth = 2) And ((iYear Mod 4) = 0) And (iDay > 29) Then
				iDay = 29
			ElseIf (iMonth = 2) And (iDay > 28) Then
				iDay = 28
			End If
			sURLForCalendar = RemoveParameterFromURLString(ReplaceValueInURLString(ReplaceValueInURLString(ReplaceValueInURLString(ReplaceValueInURLString(oRequest, "Year", iYear), "Month", iMonth), "Day", iDay), "JourneyID", lJourneyID), "EmployeeDate")
			Response.Write "<TD VALIGN=""TOP"">&nbsp;<A HREF=""" & GetASPFileName("") & "?" & sURLForCalendar & "&FromArrow=1""><IMG SRC=""Images/ArrRightBlack.gif"" WIDTH=""7"" HEIGHT=""13"" BORDER=""0"" ALT=""Mes siguiente"" /></A></TD>"
			%>
			<TD>&nbsp;</TD>
			<TD BGCOLOR="#000000"><IMG SRC="Images/Transparent.gif" WIDTH="1" HEIGHT="1" /></TD>
			<TD>&nbsp;</TD>
			<TD VALIGN="TOP"><NOBR><FONT FACE="Arial" SIZE="2">
				<IMG SRC="Images/FrameRed.gif" WIDTH="20" HEIGHT="12" /> Día en curso<BR />
				<IMG SRC="Images/FrameGray.gif" WIDTH="20" HEIGHT="12" /> Día seleccionado<BR />
				<FONT COLOR="#FF0000"><B>&nbsp;<U>00</U></B></FONT><IMG SRC="Images/Transparent.gif" WIDTH="2" HEIGHT="1" /> Días festivos<BR />
				<B>&nbsp;<U>00</U></B><IMG SRC="Images/Transparent.gif" WIDTH="2" HEIGHT="1" /> Días con registros<BR />
			</FONT></NOBR></TD>
		</TR></TABLE>
	</BODY>
</HTML>
<%
Erase aCalendarComponent
%>