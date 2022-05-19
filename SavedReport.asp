<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
%>
<!-- #include file="Libraries/GlobalVariables.asp" -->
<!-- #include file="Libraries/LoginComponent.asp" -->
<!-- #include file="Libraries/ReportComponent.asp" -->
<!-- #include file="Libraries/ReportsLib.asp" -->
<%
Dim bDoAction
Dim bShowInfo

If B_ISSSTE Then
Else
	If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REPORTS_PERMISSIONS Then
	Else
		Response.Redirect "AccessDenied.asp?Permission=" & N_REPORTS_PERMISSIONS
	End If
End If

Select Case Request.Cookies("SIAP_SectionID")
	Case "1"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
	Case "2"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
	Case "3"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case "4"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
	Case "5"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = BUDGET_TOOLBAR
	Case "6"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = REPORTS_TOOLBAR
	Case "7"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
End Select
aHeaderComponent(S_TITLE_NAME_HEADER) = "Reportes Guardados"
bWaitMessage = False

bDoAction = (Len(oRequest("Add").Item) > 0) Or (Len(oRequest("Modify").Item) > 0) Or (Len(oRequest("Remove").Item) > 0)
bShowInfo = (Len(oRequest("New").Item) > 0) Or (Len(oRequest("Change").Item) > 0) Or (Len(oRequest("Delete").Item) > 0)

Call InitializeReportComponent(oRequest, aReportComponent)
If Len(oRequest("Add").Item) > 0 Then
	lErrorNumber = AddReport(oRequest, oADODBConnection, aReportComponent, sErrorDescription)
	If lErrorNumber = 0 Then Response.Redirect "SavedReport.asp?AllReports=" & oRequest("AllReports").Item & "&ReportToShow=" & oRequest("ReportToShow").Item & "&Time=" & Now()
ElseIf Len(oRequest("Modify").Item) > 0 Then
	lErrorNumber = ModifyReport(oRequest, oADODBConnection, aReportComponent, sErrorDescription)
	If lErrorNumber = 0 Then Response.Redirect "SavedReport.asp?AllReports=" & oRequest("AllReports").Item & "&ReportToShow=" & oRequest("ReportToShow").Item & "&Time=" & Now()
ElseIf Len(oRequest("Remove").Item) > 0 Then
	lErrorNumber = RemoveReport(oRequest, oADODBConnection, aReportComponent, sErrorDescription)
	If lErrorNumber = 0 Then Response.Redirect "SavedReport.asp?AllReports=" & oRequest("AllReports").Item & "&ReportToShow=" & oRequest("ReportToShow").Item & "&Time=" & Now()
ElseIf Len(oRequest("New").Item) > 0 Then
	aHeaderComponent(S_TITLE_NAME_HEADER) = "Guardar Reporte"
	aReportComponent(N_ID_REPORT) = -1
ElseIf Len(oRequest("Change").Item) > 0 Then
	aHeaderComponent(S_TITLE_NAME_HEADER) = "Modificar Reporte"
ElseIf Len(oRequest("Delete").Item) > 0 Then
	aHeaderComponent(S_TITLE_NAME_HEADER) = "Eliminar Reporte"
End If
Response.Cookies("SoS_SectionID") = 204
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<!-- #include file="_Header.asp" -->
		<%Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> > "
		Select Case Request.Cookies("SIAP_SectionID")
			Case "1"
				Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=19"">Reportes</A> > "
			Case "2"
				Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=24"">Reportes</A> > "
			Case "3"
				Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=34"">Reportes</A> > "
			Case "4"
				Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=49"">Reportes</A> > "
			Case "5"
				If (Len(oRequest("ReportToShow").Item) > 0) Or (aReportComponent(N_CONSTANT_ID_REPORT) = 1503) Then
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=5"">Presupuesto</A> > <A HREF=""Main_ISSSTE.asp?SectionID=53"">Costeo de plazas</A> > "
				Else
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=5"">Presupuesto</A> > <A HREF=""Main_ISSSTE.asp?SectionID=58"">Reportes</A> > "
				End If
			Case "6"
				Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=6"">Departamento técnico</A> > <A HREF=""Main_ISSSTE.asp?SectionID=64"">Reportes</A> > "
			Case "7"
				Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > "
			Case Else
				Response.Write "<A HREF=""Reports.asp"">Reportes</A> > "
		End Select
		Response.Write "<B>Reportes guardados</B><BR /><BR /><BR />"
		If (Not bShowInfo) And (lErrorNumber = 0) Then
			lErrorNumber = DisplayReportsInThreeSmallColumns(oRequest, oADODBConnection, aReportComponent, sErrorDescription)
			If lErrorNumber <> 0 Then
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
			End If
		Else
			If lErrorNumber <> 0 Then
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
				Response.Write "<BR />"
				lErrorNumber = 0
				sErrorDescription = ""
			End If
			lErrorNumber = DisplayReportForm(oRequest, oADODBConnection, GetASPFileName(""), aReportComponent, sErrorDescription)
		End If%>
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>