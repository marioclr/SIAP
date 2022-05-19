<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Server.ScriptTimeout = 72000
response.buffer=true
%>
<!-- #include file="Libraries/GlobalVariables.asp" -->
<!-- #include file="Libraries/LoginComponent.asp" -->
<!-- #include file="Libraries/PayrollResumeForSarComponent.asp" -->
<!-- #include file="Libraries/ReportsQueries1000bLib.asp" -->
<!-- #include file="Libraries/ZIPLibrary.asp" -->

<%
Dim bAction
Dim sAction
Dim iStep
Dim sMessage
Dim lSuccess
Dim sError
Dim lStartMonth
Dim lEndMonth
Dim oPayrollRecordset
Dim iIndex
Dim lPeriod

sError = ""
lReasonID = 1
lPeriod = 0

sAction = oRequest("Action").Item
iStep = 1
If Len(oRequest("Step").Item) > 0 Then iStep = CInt(oRequest("Step").Item)

Call InitializePayrollResumeForSarComponent(oRequest, aPayrollResumeForSarComponent)

If B_ISSSTE Then
	Select Case CInt(Request.Cookies("SIAP_SectionID"))
		Case 1
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
		Case 4
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
		Case 7
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
		Case Else
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
	End Select
Else
	aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
End If

'aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
aHeaderComponent(S_TITLE_NAME_HEADER) = "Ejercicio bimestral del SAR"
bWaitMessage = True

%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<!-- #include file="_Header.asp" -->
		<%
		Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> >"
		Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > "
		Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=491"">Ejercicio bimestral del SAR</A> > "
		If strComp(oRequest("Action").Item,"ClosePeriod",vbBinaryCompare) = 0 Then
			Response.Write "<B>Cerrar bimestre</B>"
		ElseIf StrComp(oRequest("Action").Item, "EstrQna", vbBinaryCompare) = 0 Then
			Response.Write "<B>Generar ejercicio bimestral</B>"
		ElseIf StrComp(oRequest("Action").Item, "StartPeriod", vbBinaryCompare) = 0 Then
			Response.Write "<B>Iniciar nuevo periodo</B>"
		ElseIf StrComp(oRequest("Action").Item, "payrollCompare", vbBinaryCompare) = 0 Then
			Response.Write "<B>Generar comparativo de nóminas</B>"
		ElseIf StrComp(oRequest("Action").Item, "employeesMovements", vbBinaryCompare) = 0 Then
			Response.Write "<B>Generar proceso de altas, bajas y cambios</B>"
		ElseIf StrComp(oRequest("Action").Item, "deleteResume", vbBinaryCompare) = 0 Then 
			Response.Write "<B>Borrar resumen de nóminas</B>"
		End If
		Response.Write "<BR /><BR />"
		If strComp(oRequest("Action").Item,"ClosePeriod",vbBinaryCompare) = 0 Then
			lErrorNumber = ClosePeriod(oRequest, oADODBConnection, sErrorDescription)
			If lErrorNumber = 0 Then
				Call DisplayErrorMessage("Confirmación", "El periodo actual fue cerrado con éxito.")
				lErrorNumber = 0
			Else
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
				lErrorNumber = 0
			End If
			Response.Write "<BR />"
		ElseIf StrComp(oRequest("Action").Item, "EstrQna", vbBinaryCompare) = 0 Then
			lErrorNumber = distributePayments(oRequest, oADODBConnection, sErrorDescription)
			If lErrorNumber = 0 Then
				Call DisplayErrorMessage("Confirmación", "El ejercicio bimestral fue generado con éxito.")
				lErrorNumber = 0
			Else
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
				lErrorNumber = 0
			End If
			Response.Write "<BR />"
		ElseIf StrComp(oRequest("Action").Item, "employeesMovements", vbBinaryCompare) = 0 Then
			lErrorNumber = ComparePayrolls(oRequest, oADODBConnection, sErrorDescription)
			If lErrorNumber = 0 Then
				Call DisplayErrorMessage("Confirmación", "El proceso de altas y bajas fue generado con éxito.")
				lErrorNumber = 0
			Else
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
				lErrorNumber = 0
			End If
			Response.Write "<BR />"
            'Response.Redirect "Reports.asp?ReportID=1035"
		ElseIf StrComp(oRequest("Action").Item, "deleteResume", vbBinaryCompare) = 0 Then
			lErrorNumber = deletePayrollResume(oRequest, oADODBConnection, sErrorDescription)
			If lErrorNumber = 0 Then
				Call DisplayErrorMessage("Confirmación", "El resumen fue borrado con éxito.")
				lErrorNumber = 0
			Else
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
				lErrorNumber = 0
			End If
			Response.Write "<BR />"
		ElseIf StrComp(oRequest("Action").Item, "payrollCompare", vbBinaryCompare) = 0 Then
			lErrorNumber = paymentsCompare(oRequest, oADODBConnection, sErrorDescription)
			If lErrorNumber = 0 Then
				Call DisplayErrorMessage("Confirmación", "La comparación de nóminas concluyó con éxito.")
				lErrorNumber = 0
			Else
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
				lErrorNumber = 0
			End If
			Response.Write "<BR />"
		ElseIf StrComp(oRequest("Action").Item, "StartPeriod", vbBinaryCompare) = 0 Then
			If iStep = 1 Then
				Response.Write "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
					Response.Write "<TD WIDTH=""600"" VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"
					If lErrorNumber = 0 Then
						lErrorNumber = DisplayPeriodsList(oRequest, oADODBConnection, sErrorDescription)
					End If
					If lErrorNumber <> 0 Then
						Response.Write "<BR />"
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						sErrorDescription = ""
						bShowForm = True
					End If
					Response.Write "</FONT></TD>"
					Response.Write "<TD>&nbsp;</TD>"
					Response.Write "<TD BGCOLOR=""" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""1"" ><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
					Response.Write "<TD>&nbsp;</TD>"
					Response.Write "<TD WIDTH=""*"" VALIGN=""TOP"">"
					Response.Write "<DIV NAME=""CatalogDiv"" ID=""CatalogDiv"">"
						lErrorNumber = DisplayStartSarForm(oRequest, oADODBConnection, sErrorDescription)
					Response.Write "</DIV>"
					If lErrorNumber <> 0 Then
						Response.Write "<BR />"
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						sErrorDescription = ""
					End If
					Response.Write "</TD>"
				Response.Write "</TR></TABLE>"
			ElseIf iStep = 2 Then
				lErrorNumber = StartPeriod(oRequest, oADODBConnection, lPeriod, sErrorDescription)
				Response.Write "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
					Response.Write "<TD WIDTH=""600"" VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"
					If lErrorNumber = 0 Then
						lErrorNumber = DisplayPeriodsList(oRequest, oADODBConnection, sErrorDescription)
					Else
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						lErrorNumber = 0
						sErrorDescription = ""
					End If
					Response.Write "</FONT></TD>"
					Response.Write "<TD>&nbsp;</TD>"
					Response.Write "<TD BGCOLOR=""" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""1"" ><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
					Response.Write "<TD>&nbsp;</TD>"
					Response.Write "<TD WIDTH=""*"" VALIGN=""TOP"">"
					Response.Write "<DIV NAME=""CatalogDiv"" ID=""CatalogDiv"">"
						lErrorNumber = DisplayStartSarForm(oRequest, oADODBConnection, sErrorDescription)
					Response.Write "</DIV>"
					If lErrorNumber <> 0 Then
						Response.Write "<BR />"
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						sErrorDescription = ""
					End If
					Response.Write "</TD>"
				Response.Write "</TR></TABLE>"
			End If
		End If%>
			<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>