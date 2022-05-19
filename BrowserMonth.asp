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
<%
Dim sURLForCalendar
Dim iYear
Dim iMonth
Dim iDay
Dim sFilesNames
Dim bShowAll
Dim sFormName
Dim sDateCombo

bShowAll = (Len(oRequest("HideDesc").Item) = 0)
sFormName = oRequest("FormName").Item
sDateCombo = oRequest("DateCombo").Item

If Len(oRequest("ErrorLogDate").Item) > 0 Then
	aCalendarComponent(N_YEAR_CALENDAR) = CInt(Left(oRequest("ErrorLogDate").Item, Len("YYYY")))
	aCalendarComponent(N_MONTH_CALENDAR) = CInt(Mid(oRequest("ErrorLogDate").Item, Len("YYYYM"), Len("MM")))
	aCalendarComponent(N_DAY_CALENDAR) = CInt(Mid(oRequest("ErrorLogDate").Item, Len("YYYYMMD"), Len("DD")))
End If
Call InitializeCalendarComponent(oRequest, aCalendarComponent)
If Len(sDateCombo) = 0 Then
	lErrorNumber = GetLogFilesNames(aErrorLogComponent, "Log" & aCalendarComponent(N_YEAR_CALENDAR) & Right(("0" & aCalendarComponent(N_MONTH_CALENDAR)), Len("00")), sFilesNames, sErrorDescription)
	aCalendarComponent(S_MARKED_DAYS_CALENDAR) = Replace(Replace(sFilesNames, "Log", "", 1, -1, vbTextCompare), ".txt", "", 1, -1, vbTextCompare)
	If StrComp(Right(aCalendarComponent(S_MARKED_DAYS_CALENDAR), Len(".")), ".", vbBinaryCompare) = 0 Then aCalendarComponent(S_MARKED_DAYS_CALENDAR) = Left(aCalendarComponent(S_MARKED_DAYS_CALENDAR), (Len(aCalendarComponent(S_MARKED_DAYS_CALENDAR)) - Len(".")))
	aCalendarComponent(S_TARGET_PAGE_CALENDAR) = "ErrorLog.asp"
Else
	aCalendarComponent(S_TARGET_PAGE_CALENDAR) = GetASPFileName("") & "?FormName=" & sFormName & "&" & "DateCombo=" & sDateCombo & "&HideDesc="
	If Not bShowAll Then aCalendarComponent(S_TARGET_PAGE_CALENDAR) = aCalendarComponent(S_TARGET_PAGE_CALENDAR) & "1"
End If
aCalendarComponent(S_TARGET_FRAME_CALENDAR) = "_top"
%>
<HTML>
	<HEAD>
		<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1" />
		<TITLE>Sistema Integral de Administración de Personal del ISSSTE. Calendario por Meses</TITLE>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/PopupItem.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript"><!--
			<%If (Not bShowAll) And (Len(oRequest("Year").Item) > 0) And (Len(oRequest("FromArrow").Item) = 0) Then
				Response.Write "window.opener.document.all['" & sDateCombo & "Year'].value = '" & oRequest("Year").Item & "';" & vbNewLine
				Response.Write "window.opener.document.all['" & sDateCombo & "Month'].value = '" & Right(("0" & oRequest("Month").Item), Len("00")) & "';" & vbNewLine
				Response.Write "window.opener.ChangeDaysListGivenTheMonthAndYear(" & Right(("0" & oRequest("Month").Item), Len("00")) & ", " & oRequest("Year").Item & ", window.opener." & sFormName & "." & sDateCombo & "Day);" & vbNewLine
				Response.Write "window.opener.document.all['" & sDateCombo & "Day'].value = '" & Right(("0" & oRequest("Day").Item), Len("00")) & "';" & vbNewLine
				Response.Write "window.opener.focus();" & vbNewLine
				Response.Write "window.close();" & vbNewLine
			End If%>
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
			sURLForCalendar = RemoveParameterFromURLString(ReplaceValueInURLString(ReplaceValueInURLString(ReplaceValueInURLString(oRequest, "Year", iYear), "Month", iMonth), "Day", iDay), "ErrorLogDate")
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
			sURLForCalendar = RemoveParameterFromURLString(ReplaceValueInURLString(ReplaceValueInURLString(ReplaceValueInURLString(oRequest, "Year", iYear), "Month", iMonth), "Day", iDay), "ErrorLogDate")
			Response.Write "<TD VALIGN=""TOP"">&nbsp;<A HREF=""" & GetASPFileName("") & "?" & sURLForCalendar & "&FromArrow=1""><IMG SRC=""Images/ArrRightBlack.gif"" WIDTH=""7"" HEIGHT=""13"" BORDER=""0"" ALT=""Mes siguiente"" /></A></TD>"
			%>
			<TD>&nbsp;</TD>
			<TD BGCOLOR="#000000"><IMG SRC="Images/Transparent.gif" WIDTH="1" HEIGHT="1" /></TD>
			<TD>&nbsp;</TD>
			<TD VALIGN="TOP"><NOBR><FONT FACE="Arial" SIZE="2">
				<IMG SRC="Images/FrameRed.gif" WIDTH="20" HEIGHT="12" /> Día en curso<BR />
				<IMG SRC="Images/FrameGray.gif" WIDTH="20" HEIGHT="12" /> Día seleccionado<BR />
				<B>&nbsp;<U>00</U></B><IMG SRC="Images/Transparent.gif" WIDTH="2" HEIGHT="1" /> Días con bitácora<BR />
			</FONT></NOBR></TD>
		</TR></TABLE>
	</BODY>
</HTML>
<%
Erase aCalendarComponent
%>