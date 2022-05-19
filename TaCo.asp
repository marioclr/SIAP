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
<!-- #include file="Libraries/CatalogsLib.asp" -->
<!-- #include file="Libraries/CatalogComponent.asp" -->
<!-- #include file="Libraries/TaCoLibrary.asp" -->
<%
Dim sAction
Dim sNames
Dim bDoAction
Dim bShowForm

If Len(oRequest) > 0 Then
	sAction = ""
	If Not IsEmpty(oRequest("Action")) Then
		sAction = oRequest("Action").Item
	End If
	bDoAction = ((Len(oRequest("Add")) > 0) Or (Len(oRequest("Modify")) > 0) Or (Len(oRequest("SetActive")) > 0) Or (Len(oRequest("Remove")) > 0) Or (Len(oRequest("Active")) > 0) Or (Len(oRequest("Deactive")) > 0) Or (Len(oRequest("Unlock")) > 0))

	Call InitializeCatalogs(oRequest)
	If bDoAction Then
		Call InitializeValuesForCatalogComponent(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
		lErrorNumber = DoAction(sAction, bShowForm, sErrorDescription)
	End If
End If

Select Case CInt(Request.Cookies("SIAP_SectionID"))
	Case 3
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
		Response.Cookies("SoS_SectionID") = 1372
	Case 4
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
		Response.Cookies("SoS_SectionID") = 1372
	Case 6
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = REPORTS_TOOLBAR
		Response.Cookies("SoS_SectionID") = 1631
End Select
aHeaderComponent(S_TITLE_NAME_HEADER) = "Tablero de control"
bWaitMessage = True
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%Select Case sAction
			Case "Tasks"
				sNames = Split(oRequest("TaskPath").Item, ",")
				aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
					Array("Agregar actividad nueva",_
						  "",_
						  "", GetASPFileName("") & "?Action=" & sAction & "&New=1&ProjectID=" & oRequest("ProjectID").Item & "&TaskID=-1&ParentID=" & oRequest("ParentID").Item, True),_
					Array("Agregar actividad existente",_
						  "",_
						  "", GetASPFileName("") & "?Action=" & sAction & "&Import=1&ProjectID=" & oRequest("ProjectID").Item & "&TaskID=-1&ParentID=" & oRequest("ParentID").Item, True)_
				)
			Case Else
				aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
					Array("Agregar registro",_
						  "",_
						  "", GetASPFileName("") & "?Action=" & sAction & "&New=1&" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "=-1&" & aCatalogComponent(S_URL_CATALOG), True)_
				)
		End Select
		aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 810
		aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
		aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 183%>
		<!-- #include file="_Header.asp" -->
		<%Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> > "
		Select Case CInt(Request.Cookies("SIAP_SectionID"))
			Case 3
				Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=37"">Planeación de recursos humanos</A> > "
				sNames = "Alta y modificación de procesos"
			Case 4
				Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Informática</A> > "
				If Len(sAction) > 0 Then Response.Write "<A HREF=""TaCo.asp"">Tablero de control de procesos</A> > "
				sNames = "Tablero de control de procesos"
			Case 6
				Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=6"">Departamento técnico</A> > "
				If Len(sAction) > 0 Then Response.Write "<A HREF=""TaCo.asp"">Tablero de control de procesos</A> > "
				sNames = "Tablero de control de procesos"
		End Select
		If Len(sAction) = 0 Then
			Response.Write "<B>" & sNames & "</B><BR />"
			Response.Write "<BR /><BR /><TABLE WIDTH=""720"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				aMenuComponent(A_ELEMENTS_MENU) = Array(_
					Array("Alta y modificación de procesos",_
						  "Registre y modifique los procesos y los objetivos, metas y actividades que lo conforman.",_
						  "Images/MnBudget.gif", "TaCo.asp?Action=Projects", True),_
					Array("Seguimiento de un proceso",_
						  "Revise y actualice el avance de los procesos.",_
						  "Images/MnSection63.gif", "Projects.asp", True)_
				)
				aMenuComponent(B_USE_DIV_MENU) = True
				Call DisplayMenuInTwoColumns(aMenuComponent)
			Response.Write "</TABLE>"
		Else
			Select Case sAction
				Case "Tasks"
					Response.Write "<A HREF=""TaCo.asp?Action=Projects"">Alta y modificación de procesos</A> > "
					Call DisplayTaskPath(oADODBConnection, aTaskComponent(N_PROJECT_ID_TASK), aTaskComponent(S_PATH_TASK), True, sErrorDescription)
				Case Else
					Response.Write "<B>Alta y modificación de procesos</B><BR /><BR />"
			End Select
			If lErrorNumber <> 0 Then
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
				lErrorNumber = 0
				sErrorDescription = ""
				bShowForm = (Len(oRequest("Add")) > 0)
			End If
			Response.Write "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
				Response.Write "<TD WIDTH=""600"" VALIGN=""TOP"">"
					Response.Write "<DIV NAME=""EntriesDiv"" ID=""EntriesDiv"" CLASS=""TableScrollDiv"">"
						If lErrorNumber = 0 Then
							aCatalogComponent(S_QUERY_CONDITION_CATALOG) = "((ProjectFile Like '" & S_WILD_CHAR & iGlobalSectionID & S_WILD_CHAR & "') Or (ProjectFile Like '" & S_WILD_CHAR & "-1" & S_WILD_CHAR & "'))"
							lErrorNumber = DisplayTables(sAction, sErrorDescription)
						End If
						If lErrorNumber <> 0 Then
							Response.Write "<BR />"
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							lErrorNumber = 0
							sErrorDescription = ""
							bShowForm = True
						End If
					Response.Write "</DIV>"
				Response.Write "</TD>"
				Response.Write "<TD>&nbsp;</TD>"
				Response.Write "<TD BGCOLOR=""" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""1"" ><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
				Response.Write "<TD>&nbsp;</TD>"
				Response.Write "<TD WIDTH=""*"" VALIGN=""TOP"">"
					Response.Write "<DIV NAME=""CatalogDiv"" ID=""CatalogDiv"">"
						lErrorNumber = DisplayForms(sAction, sErrorDescription)
					Response.Write "</DIV>"
					If lErrorNumber <> 0 Then
						Response.Write "<BR />"
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						sErrorDescription = ""
					End If
				Response.Write "</TD>"
			Response.Write "</TR></TABLE>"
			'<!-- END: CATALOGS -->
		End If
		%>
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>