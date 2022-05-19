<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Server.ScriptTimeout = 72000
%>
<!-- #include file="Libraries/GlobalVariables.asp" -->
<!-- #include file="Libraries/LoginComponent.asp" -->
<!-- #include file="Libraries/ConceptComponent.asp" -->
<!-- #include file="Libraries/_EmployeeChangesLKP.asp" -->
<!-- #include file="Libraries/ReportsLib.asp" -->
<%
Dim sAction
Dim bShowForm
Dim bShowTable
Dim lPayrollID
Dim iPayrollStatus
Dim sNames
Dim iSelectedTab
Dim lRecordID
Dim lRecordID2
Dim sAuthorize
Dim lSuccess
Dim lPositionID

If B_ISSSTE Then
Else
	If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_PAYROLL_PERMISSIONS) = N_PAYROLL_PERMISSIONS Then
	Else
		Response.Redirect "AccessDenied.asp?Permission=" & N_PAYROLL_PERMISSIONS
	End If
End If

sAction = oRequest("Action").Item
lRecordID = oRequest("RecordID").Item
lRecordID2 = oRequest("RecordID2").Item
sAuthorize = oRequest("Authorize").Item
bShowTable = False

Call InitializePayrollComponent(oRequest, aPayrollComponent)

aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
aHeaderComponent(S_TITLE_NAME_HEADER) = "Nómina"
lErrorNumber = GetLastPayrollStatus(oADODBConnection, lPayrollID, iPayrollStatus, sErrorDescription)
If Len(oRequest("CalculatePayroll").Item) > 0 Then
	aHeaderComponent(S_TITLE_NAME_HEADER) = "Prenómina"
	If Len(Application.Contents("SIAP_CalculatePayroll")) = 0 Then
		Application.Contents("SIAP_CalculatePayroll") = aLoginComponent(N_USER_ID_LOGIN) & LIST_SEPARATOR & aPayrollComponent(N_ID_PAYROLL) & LIST_SEPARATOR & GetSerialNumberForDate("")
		lErrorNumber = CalculatePayroll(oRequest, oADODBConnection, 2, aPayrollComponent, sErrorDescription)
	End If
	If lErrorNumber =  0 Then
		lErrorNumber = GetPayroll(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
		If lErrorNumber =  0 Then
			aPayrollComponent(N_CLOSED_PAYROLL) = 2
			lErrorNumber = ModifyPayroll(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
		End If
	End If
End If

bWaitMessage = True
Response.Cookies("SoS_SectionID") = 194
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/Export.js"></SCRIPT>
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<!-- #include file="_Header.asp" -->
		<%Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> > <A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <B>Prenómina</B><BR /><BR />"

		If lErrorNumber <> 0 Then
			Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
			lErrorNumber = 0
			sErrorDescription = ""
			Response.Write "<BR />"
		End If
		If bShowTable Then
			Response.Write "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0""><TR>" & vbNewLine
				Select Case sAction
					Case Else
						Response.Write "<TD WIDTH=""600"" VALIGN=""TOP"">" & vbNewLine
							If StrComp(oRequest("Action").Item, "UpdatePayroll", vbBinaryCompare) = 0 Then
								aPayrollComponent(S_QUERY_CONDITION_PAYROLL) = " And (IsClosed<>1)"
								lErrorNumber = DisplayPayrollsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, True, aPayrollComponent, sErrorDescription)
							Else
								lErrorNumber = DisplayPayrollsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, False, aPayrollComponent, sErrorDescription)
							End If
							If lErrorNumber <> 0 Then
								Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
								lErrorNumber = 0
								sErrorDescription = ""
								Response.Write "<BR />"
							Else
								Response.Write "<BR />"
								Response.Write "<INPUT TYPE=""BUTTON"" VALUE=""Regresar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=4'"" />"
							End If
						Response.Write "</TD>" & vbNewLine
				End Select
				Response.Write "<TD>&nbsp;</TD>"
				Response.Write "<TD BGCOLOR=""" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""1"" ><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
				Response.Write "<TD>&nbsp;</TD>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">" & vbNewLine
		End If
			If Len(oRequest("CalculatePayroll").Item) > 0 Then
				sNames = L_NO_INSTRUCTIONS_FLAGS & "," & L_OPEN_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_JOURNEY_FLAGS & "," & L_SHIFT_FLAGS & "," & L_LEVEL_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS
				lErrorNumber = DisplayFilterInformation(oRequest, sNames, False, "", sErrorDescription)
			End If
			Call DisplayModifyPayrollMessage(0, aPayrollComponent(N_ID_PAYROLL))
		If bShowTable Then
				Response.Write "</FONT></TD>" & vbNewLine
			Response.Write "</TR></TABLE>" & vbNewLine
		End If

		If lErrorNumber <> 0 Then
			Response.Write "<BR /><BR />"
			Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
			lErrorNumber = 0
			sErrorDescription = ""
			Response.Write "<BR />"
		End If%>
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>