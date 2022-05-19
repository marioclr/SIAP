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
<!-- #include file="Libraries/TaCoLibrary.asp" -->
<!-- #include file="Libraries/TaCoHistoryListComponent.asp" -->
<!-- #include file="Libraries/TaCoProjectsLib.asp" -->
<!-- #include file="Libraries/TaCoTaskComponent.asp" -->
<%
Dim iSelectedTab
Dim bAction
Dim sView
Dim lRecordID
Dim sImageSuffix
Dim sNames

Call GetProjectsURLValues(oRequest, iSelectedTab, bAction)
Call InitializeTaskComponent(oRequest, aTaskComponent)
Call InitializeHistoryComponent(oRequest, aHistoryComponent)
sView = oRequest("View").Item
lRecordID = -1

sImageSuffix = "Small"
If bAction Then
	lErrorNumber = DoProjectsAction(oRequest, oSIAPTACOADODBConnection, iSelectedTab, sErrorDescription)
End If

If (lErrorNumber = 0) And (aTaskComponent(N_ID_TASK) <> -1) And (aTaskComponent(N_PROJECT_ID_TASK) <> -1) Then
	lErrorNumber = GetTask(oRequest, oSIAPTACOADODBConnection, aTaskComponent, sErrorDescription)
End If

Select Case iGlobalSectionID
	Case 1
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
		Response.Cookies("SoS_SectionID") = 1373
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Seguimiento de procesos"
	Case 2
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Response.Cookies("SoS_SectionID") = 1373
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Seguimiento de procesos"
	Case 3
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
		Response.Cookies("SoS_SectionID") = 1373
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Seguimiento de procesos"
	Case 4
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
		Response.Cookies("SoS_SectionID") = 1373
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Seguimiento de procesos"
	Case 5
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = BUDGET_TOOLBAR
		Response.Cookies("SoS_SectionID") = 1373
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Seguimiento de procesos"
	Case 6
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = REPORTS_TOOLBAR
		Response.Cookies("SoS_SectionID") = 1632
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Tablero de control de procesos"
	Case 7
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
		Response.Cookies("SoS_SectionID") = 1750
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Tablero de control de procesos"
End Select
bWaitMessage = True
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/marketplace.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript"><!--
			var sCurrentWidgetDiv = '';

			function ShowTaskStatusFrm(bShowTabs) {
				HideWidget('');
				ToggleDisplay(document.all['TaskPanelDiv']);
				ToggleDisplay(document.all['PicTaskPanelDiv']);
				ToggleDisplay(document.all['AdvancePanelDiv']);
				ToggleDisplay(document.all['PicAdvancePanelDiv']);
				ToogleImage(document.images['ExpandArrowImg'], 'Images\/ArrExpandLf<%Response.Write sImageSuffix%>.gif', 'Images\/ArrExpandRg<%Response.Write sImageSuffix%>.gif');
				if (bShowTabs) {
					if (document.TaskStatusFrm) {
						ShowTaskTab(1);
						if (! IsDisplayed(document.all['TaskParametersDiv']))
							ToogleDiv('TaskParameters');
						document.TaskStatusFrm.TaskStatusPercentage.focus();
					}
				}
			} // End of ShowTaskStatusFrm

			function ShowWidget(sDivID) {
				HideWidget(sCurrentWidgetDiv);

				var oWidgetDiv = document.all['WidgetInfo' + sDivID + 'Div'];
				if (oWidgetDiv) {
					ShowPopupItem('WidgetInfo' + sDivID + 'Div', oWidgetDiv);
				}

				sCurrentWidgetDiv = sDivID;
			} // End of ShowWidget

			function HideWidget(sDivID) {
				var oWidgetDiv = document.all['WidgetInfo' + sDivID + 'Div'];
				if (oWidgetDiv)
					HidePopupItem('WidgetInfo' + sDivID + 'Div', oWidgetDiv);
			} // End of HideWidget
		//--></SCRIPT>
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<!-- #include file="_Header.asp" -->
		<%Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> > "
		Select Case iGlobalSectionID
			Case 1
				Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""TaCo.asp"">Tablero de Control de procesos</A> > "
				sNames = "Seguimiento de procesos"
			Case 2
				Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""TaCo.asp"">Tablero de Control de procesos</A> > "
				sNames = "Seguimiento de procesos"
			Case 3
				Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=37"">Planeación de recursos humanos</A> > "
				sNames = "Seguimiento de procesos"
			Case 4
				Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""TaCo.asp"">Tablero de Control de procesos</A> > "
				sNames = "Seguimiento de procesos"
			Case 5
				Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=5"">Presupuesto</A> > <A HREF=""TaCo.asp"">Tablero de Control de procesos</A> > "
				sNames = "Seguimiento de procesos"
			Case 6
				Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=6"">Departamento técnico</A> > <A HREF=""TaCo.asp"">Tablero de control de procesos</A> > "
				sNames = "Seguimiento de un proceso"
			Case 7
				Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > "
				If aTaskComponent(N_PROJECT_ID_TASK) = -1 Then
					sNames = "Tablero de control de procesos"
				End If
		End Select
		If aTaskComponent(N_PROJECT_ID_TASK) = -1 Then
			Response.Write "<B>" & sNames & "</B><BR /><BR />"
			If lErrorNumber <> 0 Then
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
				lErrorNumber = 0
				sErrorDescription = ""
				Response.Write "<BR />"
			End If
			lErrorNumber = DisplayProjects(oSIAPTACOADODBConnection, sErrorDescription)
			If lErrorNumber <> 0 Then
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
				lErrorNumber = 0
				sErrorDescription = ""
				Response.Write "<BR />"
			End If
		Else
			If iGlobalSectionID = 7 Then Response.Write "<A HREF=""Projects.asp"">Tablero de control de procesos</A> "
			Response.Write "<A HREF=""Projects.asp"">" & sNames & "</A> > "
			Call DisplayTaskPath(oSIAPTACOADODBConnection, aTaskComponent(N_PROJECT_ID_TASK), aTaskComponent(S_PATH_TASK), False, sErrorDescription)
			If lErrorNumber <> 0 Then
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
				lErrorNumber = 0
				sErrorDescription = ""
				Response.Write "<BR />"
			ElseIf Len(oRequest("UpdateTaskStatus").Item) > 0 Then
				Call DisplayInstructionsMessage("Confirmación", "El avance en la actividad fue registrado correctamente.")
				Response.Write "<BR />"
			End If
			lErrorNumber = DisplayTaskForms(oSIAPTACOADODBConnection, aTaskComponent, sErrorDescription)
		End If%>
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>