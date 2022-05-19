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
<!-- #include file="Libraries/EmployeeComponent.asp" -->
<!-- #include file="Libraries/JobsLib.asp" -->
<!-- #include file="Libraries/JobComponent.asp" -->
<!-- #include file="Libraries/ReportsLib.asp" -->
<%
Dim iSelectedTab
Dim bAction
Dim bError
Dim sError

sError = ""

If B_ISSSTE Then
Else
	If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_JOBS_PERMISSIONS) = N_JOBS_PERMISSIONS Then
	Else
		Response.Redirect "AccessDenied.asp?Permission=" & N_JOBS_PERMISSIONS
	End If
End If

Call InitializeJobComponent(oRequest, aJobComponent)
Call GetJobsURLValues(oRequest, iSelectedTab, bAction, aJobComponent(S_QUERY_CONDITION_JOB))

If Len(oRequest("AuthorizationFile").Item) > 0 Then
	aJobComponent(N_ACTIVE_JOB) = 1
	lErrorNumber = AddJobFile(oRequest, oADODBConnection, oRequest("sQuery").Item, 59, aJobComponent, sErrorDescription)
	Response.Redirect "UploadInfo.asp?Action=Jobs&ReasonID=59&MovementsSuccess=1"
End If

bError = False
If bAction And Len(oRequest("JobIDH").Item) = 0 Then
	lErrorNumber = DoJobsAction(oRequest, oADODBConnection, oRequest("Action").Item, sErrorDescription)
	bError = (lErrorNumber <> 0)
	If (lErrorNumber = 0) And (Len(oRequest("Remove").Item) > 0) Then
		aJobComponent(N_ID_JOB) = -1
		bAction = False
	End If
End If
If CLng(oRequest("JobIDH").Item) > 0 Then
	aJobComponent(N_ID_JOB) = oRequest("JobIDH").Item
	lErrorNumber = DoJobsAction(oRequest, oADODBConnection, oRequest("Action").Item, sErrorDescription)
	B_ISSSTE = True
		Response.Redirect "Jobs.asp?Action=Jobs&JobID=" & oRequest("JobIDH").Item & "&Change=1&Tab=2"
End If
'If aJobComponent(N_ID_JOB) > -1 Then
'	lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
'End If

If B_ISSSTE Then
	Select Case CInt(Request.Cookies("SIAP_SectionID"))
		Case 1
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
        Case 3
            aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
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
aHeaderComponent(S_TITLE_NAME_HEADER) = "Administración de plazas"
bWaitMessage = True
Response.Cookies("SoS_SectionID") = 192
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/Export.js"></SCRIPT>
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%If (Len(oRequest("DoSearch").Item) > 0) Or (aJobComponent(N_ID_JOB) > -1) Then
			aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Agregar una nueva plaza",_
					  "",_
					  "", "UploadInfo.asp?Action=Jobs&ReasonID=59", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_ModificacionDePlazas & ",", vbBinaryCompare) > 0),_
				Array("<LINE />",_
					  "",_
					  "", "", (Len(oRequest("DoSearch").Item) > 0)),_
				Array("Agregar un registro al historial",_
					  "",_
					  "", "Jobs.asp?Action=Jobs&JobID=" & aJobComponent(N_ID_JOB) & "&Tab=2&JobDate=0", (StrComp(oRequest("Action").Item, "Jobs", vbBinaryCompare) = 0) And (StrComp(oRequest("Tab").Item, "2", vbBinaryCompare) = 0)),_
				Array("Exportar a Excel",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=Jobs&Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (Len(oRequest("DoSearch").Item) > 0)),_
				Array("Imprimir",_
					  "",_
					  "", "javascript: SendReportToPrint('ReportDiv', '" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "')", False)_
			)
			aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 793
			aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
			aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 200
		End If%>
		<!-- #include file="_Header.asp" -->
		<%
			Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> > "
			If B_ISSSTE Then
				'Response.Write "Main_ISSSTE.asp?SectionID=1"
			Else
				Response.Write "HumanResources.asp"
			End If
			If Len(oRequest("New").Item) > 0 Then
				Response.Write """>Personal</A> > "
				Response.Write "<A HREF=""Jobs.asp"">Administración de plazas</A> > <B>Plazas</B>"
			Else
				If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <B>Administración de plazas</B>"
				ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
                    Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <B>Administración de plazas</B>"
                ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=71"">Personal</A> > <B>Consulta de plazas</B>"
				End If
				If aJobComponent(N_ID_JOB) > 0 Then
					Response.Write " > <B>Plaza No. " & CleanStringForHTML(aJobComponent(S_NUMBER_JOB)) & "</B>"
				End If
			End If
		Response.Write "<BR /><BR />"
		If Len(oRequest("Search").Item) > 0 Then
			Call DisplayJobsSearchForm(oRequest, oADODBConnection, True, sErrorDescription)
		ElseIf (Len(oRequest("New").Item) > 0) Or bError Then
			If bAction And (lErrorNumber <> 0) Then
				Response.Write "<BR /><BR />"
				Call DisplayErrorMessage("Error en la información de la plaza", sErrorDescription)
				lErrorNumber = 0
				Response.Write "<BR />"
			End If
			If aJobComponent(N_ID_JOB) > -1 Then
				lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
			End If
			Call DisplayJobForm(oRequest, oADODBConnection, GetASPFileName(""), aJobComponent, sErrorDescription)
		ElseIf ((Len(oRequest("Tab").Item) > 0) Or bAction) And (Not bError) And (Len(oRequest("Remove").Item) = 0) Then
			If bAction Then
				Response.Write "<BR />"
				If lErrorNumber = 0 Then
					Call DisplayErrorMessage("Confirmación", "La información de la plaza fue guardada con éxito.")
					lErrorNumber = 0
				Else
					Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
					lErrorNumber = 0
				End If
				Response.Write "<BR />"
			End If
			Call DisplayJobsTabs(oRequest, bError, sErrorDescription)
			Response.Write "<BR />"
			lErrorNumber = DisplayJobForms(oRequest, iSelectedTab, sErrorDescription)
			If lErrorNumber <> 0 Then
				Response.Write "<BR /><BR />"
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
				lErrorNumber = 0
				Response.Write "<BR />"
			End If
		ElseIf Len(oRequest("DoSearch").Item) > 0 Then
			lErrorNumber = DisplayJobsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, True, False, aJobComponent, sErrorDescription)
			If lErrorNumber = L_ERR_NO_RECORDS Then
				Call DisplayErrorMessage("Búsqueda vacía", sErrorDescription)
				lErrorNumber = 0
				sErrorDescription = ""
				Response.Write "<BR />"
				Call DisplayJobsSearchForm(oRequest, oADODBConnection, True, sErrorDescription)
			End If
		Else
			If Len(oRequest("Remove").Item) > 0 Then
				Call DisplayErrorMessage("Confirmación", "La información de la plaza fue eliminada con éxito.")
			End If
			Response.Write "<BR /><TABLE WIDTH=""720"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD WIDTH=""50%"" COLSPAN=""2"" VALIGN=""TOP""><FONT FACE=""ARIAL"" SIZE=""2""><B>Búsqueda de plazas</B></FONT></TD>"
					Response.Write "<TD WIDTH=""50%"" COLSPAN=""2"" VALIGN=""TOP"">"
					If StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_ModificacionDePlazas & ",", vbBinaryCompare) > 0 Then
						Response.Write "<FONT FACE=""ARIAL"" SIZE=""2""><B><A HREF=""UploadInfo.asp?Action=Jobs&ReasonID=59"">Agregar una nueva plaza</A></B></FONT>"
					End If
					Response.Write "</TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD ROWSPAN=""2"">&nbsp;&nbsp;&nbsp;</TD>"
					Response.Write "<TD WIDTH=""50%"" VALIGN=""TOP"">"
						Call DisplayJobsSearchForm(oRequest, oADODBConnection, False, sErrorDescription)
					Response.Write "</TD>"
					Response.Write "<TD ROWSPAN=""2"">&nbsp;&nbsp;&nbsp;</TD>"
					Response.Write "<TD WIDTH=""50%"" VALIGN=""TOP""><FONT FACE=""ARIAL"" SIZE=""2"">"
						If StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_ModificacionDePlazas & ",", vbBinaryCompare) > 0 Then
							Response.Write "Crea nuevas plazas indicando el puesto, centro de trabajo, centro de pago y vigencia de la plaza.<BR />"
						End If
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE><BR />"
			Response.Write "<IMG SRC=""Images/DotBlue.gif"" WIDTH=""980"" HEIGHT=""1"" /><BR /><BR />"
			aMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("212 Cambio de servicio",_
					  "Utilice un archivo para registrar las plazas y los nuevos servicios.",_
					  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=54", False),_
				Array("Cambio de datos a las plazas",_
					  "Utilice un archivo para registrar cambios a las plazas.",_
					  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Jobs&ReasonID=60", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_ModificacionDePlazas & ",", vbBinaryCompare) > 0),_
				Array("Cambio de puesto a las plazas",_
					  "Utilice un archivo para registrar cambios de puesto a las plazas.",_
					  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Jobs&ReasonID=61", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_ModificacionDePlazas & ",", vbBinaryCompare) > 0),_
				Array("<LINE />",_
					  "",_
					  "", "", True),_
				Array("Plazas ocupadas",_
					  "Obtenga un listado de las plazas que están asignadas a los empleados.",_
					  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=" & SPECIAL_JOBS_LIST_REPORTS & "&Template=" & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS & "," & L_OCCUPATION_TYPE_FLAGS & "," & L_JOB_START_DATE_FLAGS & "," & L_JOB_END_DATE_FLAGS & "&ReportStep=2&JobStatusID=1&Flags=" & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS & "," & L_DATE_FLAGS, True),_
				Array("Plazas vacantes",_
					  "Revise las plazas que están disponibles para ser asignadas.",_
					  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=" & SPECIAL_JOBS_LIST_REPORTS & "&Template=" & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS & "," & L_OCCUPATION_TYPE_FLAGS & "," & L_JOB_START_DATE_FLAGS & "," & L_JOB_END_DATE_FLAGS & "&ReportStep=2&JobStatusID=2&Flags=" & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_DATE_FLAGS, True),_
				Array("Plazas congeladas",_
					  "Lista de las plazas que están en estatus de congelada.",_
					  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=" & SPECIAL_JOBS_LIST_REPORTS & "&Template=" & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS & "," & L_OCCUPATION_TYPE_FLAGS & "," & L_JOB_START_DATE_FLAGS & "," & L_JOB_END_DATE_FLAGS & "&ReportStep=2&JobStatusID=3&Flags=" & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS & "," & L_DATE_FLAGS, True),_
				Array("Plazas con licencia",_
					  "Obtenga un listado de las plazas ocupadas por empleados que se encuentran en licencia.",_
					  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=" & SPECIAL_JOBS_LIST_REPORTS & "&Template=" & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS & "," & L_OCCUPATION_TYPE_FLAGS & "," & L_JOB_START_DATE_FLAGS & "," & L_JOB_END_DATE_FLAGS & "&ReportStep=2&JobStatusID=4&Flags=" & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS & "," & L_DATE_FLAGS, True),_
				Array("Plazas en interinatos",_
					  "Revise las plazas con licencia ocupadas por empleados interinos.",_
					  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=" & SPECIAL_JOBS_LIST_REPORTS & "&Template=" & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS & "," & L_OCCUPATION_TYPE_FLAGS & "," & L_JOB_START_DATE_FLAGS & "," & L_JOB_END_DATE_FLAGS & "&ReportStep=2&MovedEmployees=1&Flags=" & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS & "," & L_DATE_FLAGS, True),_
				Array("Plazas reservadas",_
					  "Lista de las plazas que están en estatus de reservada.",_
					  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=" & SPECIAL_JOBS_LIST_REPORTS & "&Template=" & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS & "," & L_OCCUPATION_TYPE_FLAGS & "," & L_JOB_START_DATE_FLAGS & "," & L_JOB_END_DATE_FLAGS & "&ReportStep=2&JobsOwners=1&Flags=" & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS & "," & L_DATE_FLAGS, True),_
				Array("Plazas creadas o modificadas",_
					  "Lista de las plazas creadas o modificadas en un periodo dado.",_
					  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=" & JOBS_LIST_BY_MODIFY_DATE & "&Template=" & L_JOB_NUMBER_FLAGS & "," & L_JOB_START_DATE_FLAGS & "," & L_JOB_END_DATE_FLAGS & "&ReportStep=2&JobsOwners=1&Flags=" & L_JOB_NUMBER_FLAGS & "," & L_DATE_FLAGS, True),_
				Array("xxx",_
					  "xxx",_
					  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=" & 0, False)_
			)
			aMenuComponent(B_USE_DIV_MENU) = True
			Response.Write "<TABLE WIDTH=""900"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Call DisplayMenuInThreeSmallColumns(aMenuComponent)
			Response.Write "</TABLE>"
		End If
		If lErrorNumber <> 0 Then
			Response.Write "<BR /><BR />"
			Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
			lErrorNumber = 0
			Response.Write "<BR />"
		End If%>
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>