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
<!-- #include file="Libraries/BudgetLib.asp" -->
<!-- #include file="Libraries/BudgetComponent.asp" -->
<%
Dim sSection
Dim bAction
Dim bError
Dim sNames
Dim aPath
Dim iIndex
Dim aTempBudgetComponent()
Dim sErrorMessage
Dim sBudgetPath

If B_ISSSTE Then
Else
	If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_BUDGET_PERMISSIONS) = N_BUDGET_PERMISSIONS Then
	Else
		Response.Redirect "AccessDenied.asp?Permission=" & N_BUDGET_PERMISSIONS
	End If
End If

Call InitializeBudgetComponent(oRequest, aBudgetComponent)
Call GetBudgetURLValues(oRequest, sSection, bAction, aBudgetComponent(S_QUERY_CONDITION_BUDGET))

bError = False
If bAction Then
	lErrorNumber = DoBudgetAction(oRequest, oADODBConnection, sSection, oRequest("Action").Item, sErrorDescription)
	bError = (lErrorNumber <> 0)
	sErrorMessage = sErrorDescription
End If

aHeaderComponent(L_SELECTED_OPTION_HEADER) = BUDGET_TOOLBAR
bWaitMessage = True
Select Case sSection
	Case "Budget"
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Clasificador por objeto del gasto"
		If (lErrorNumber = 0) And (aBudgetComponent(N_ID_BUDGET) > -1) And (Len(oRequest("View").Item) = 0) Then
			lErrorNumber = GetBudget(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
		End If
		Response.Cookies("SoS_SectionID") = 189
	Case "Money"
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Administración del presupuesto"
		Response.Cookies("SoS_SectionID") = 189
	Case "Program"
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Estructuras programáticas"
		If (lErrorNumber = 0) And (aBudgetComponent(N_ID_BUDGET) > -1) And (Len(oRequest("View").Item) = 0) Then
			lErrorNumber = GetProgram(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
		End If
		Response.Cookies("SoS_SectionID") = 189
End Select
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/Export.js"></SCRIPT>
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%sNames = ""
		Call GetProgramParent(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
		Call GetProgramPath(oRequest, oADODBConnection, aBudgetComponent, "")
		aPath = Split(aBudgetComponent(S_PATH_BUDGET), ",")
		If UBound(aPath) > 2 Then sNames = aPath(2)
		If StrComp(sSection, "Money", vbBinaryCompare) = 0 Then
			If FileExists(Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_56_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & Left(GetSerialNumberForDate(""), Len("00000000")) & ".txt"), sErrorDescription) Then
				aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
					Array("Imprimir adecuaciones",_
						  "",_
						  "", "javascript: OpenNewWindow('Export.asp?Action=ModifiedMoneys&Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True),_
					Array("Borrar bitácora de adecuaciones",_
						  "",_
						  "", "javascript: OpenNewWindow('Remove.asp?Action=Rep_56&UserID=" & aLoginComponent(N_USER_ID_LOGIN) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "', '', 'RemoveWnd', 320, 240, 'no', 'yes')", True)_
				)
			End If
		Else
			If Len(oRequest("BudgetID").Item) > 0 Then
				sBudgetPath = "javascript: OpenNewWindow('Export.asp?Action=Budgets&Excel=1&ParentID=" & oRequest("BudgetID").Item & "&AccessKey="
			Else
				sBudgetPath = "javascript: OpenNewWindow('Export.asp?Action=Budgets&Excel=1&AccessKey="
			End If
			aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Agregar un nuevo registro",_
					  "",_
					  "", "Budget.asp?Section=Budget&ParentID=" & aBudgetComponent(N_ID_BUDGET) & "&New=1", (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS) = N_ADD_PERMISSIONS) And (StrComp(sSection, "Budget", vbBinaryCompare) = 0))),_
				Array("Exportar a Excel",_
					  "",_
					  "", sBudgetPath & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (StrComp(sSection, "Budget", vbBinaryCompare) = 0)),_
				Array("Agregar un nuevo registro",_
					  "",_
					  "", "Budget.asp?Section=Money&ParentID=" & aBudgetComponent(N_ID_BUDGET) & "&New=1", (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS) = N_ADD_PERMISSIONS) And (StrComp(sSection, "Money", vbBinaryCompare) = 0))),_
				Array("Agregar un nuevo registro",_
					  "",_
					  "", "Budget.asp?Section=Program&ParentID=" & aBudgetComponent(N_ID_BUDGET) & "&New=1", (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS) = N_ADD_PERMISSIONS) And (StrComp(sSection, "Program", vbBinaryCompare) = 0))),_
				Array("Exportar a Excel",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=Programs&Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&ProgramYear=" & sNames & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", ((StrComp(sSection, "Program", vbBinaryCompare) = 0) And (Len(sNames) > 0)))_
			)
		End If
		aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 793
		aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
		aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 200
		%>
		<!-- #include file="_Header.asp" -->
		<%Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> > <A HREF=""Main_ISSSTE.asp?SectionID=5"">Presupuesto</A> > "
		Select Case sSection
			Case "Budget"
				If (aBudgetComponent(N_ID_BUDGET) = -1) And (aBudgetComponent(N_PARENT_ID_BUDGET) = -1) Then
					Response.Write "<B>Clasificador por objeto del gasto</B>"
				ElseIf (aBudgetComponent(N_ID_BUDGET) = -1) Then
					Redim aTempBudgetComponent(N_BUDGET_COMPONENT_SIZE)
					aTempBudgetComponent(N_ID_BUDGET) = aBudgetComponent(N_PARENT_ID_BUDGET)
					Response.Write "<A HREF=""Budget.asp?Section=Budget"">Clasificador por objeto del gasto</A> > "
					Call DisplayBudgetPath(oRequest, oADODBConnection, aTempBudgetComponent, "")
				Else
					Response.Write "<A HREF=""Budget.asp?Section=Budget"">Clasificador por objeto del gasto</A> > "
					Call DisplayBudgetPath(oRequest, oADODBConnection, aBudgetComponent, "")
				End If
			Case "Money"
				Response.Write "<B>Administración del presupuesto</B>"
			Case "Program"
				If (aBudgetComponent(N_ID_BUDGET) = -1) And (aBudgetComponent(N_PARENT_ID_BUDGET) = -1) Then
					Response.Write "<B>Estructuras programáticas</B>"
				ElseIf (aBudgetComponent(N_ID_BUDGET) = -1) Then
					Redim aTempBudgetComponent(N_BUDGET_COMPONENT_SIZE)
					aTempBudgetComponent(N_ID_BUDGET) = aBudgetComponent(N_PARENT_ID_BUDGET)
					Response.Write "<A HREF=""Budget.asp?Section=Program"">Estructuras programáticas</A> > "
					Call DisplayProgramPath(oRequest, oADODBConnection, aTempBudgetComponent, "")
				Else
					Response.Write "<A HREF=""Budget.asp?Section=Program"">Estructuras programáticas</A> > "
					Call DisplayProgramPath(oRequest, oADODBConnection, aBudgetComponent, "")
				End If
		End Select
		Response.Write "<BR /><BR />"

		If (lErrorNumber <> 0) Or bError Then
			Response.Write "<BR />"
			If Len(sErrorMessage) > 0 Then
				Call DisplayErrorMessage("Mensaje del sistema", sErrorMessage)
			Else
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
			End If
			lErrorNumber = 0
			Response.Write "<BR />"
		End If

		If StrComp(sSection, "Money", vbBinaryCompare) = 0 Then 
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
				Response.Write "<TD WIDTH=""480"" VALIGN=""TOP"">"
					Response.Write "<IFRAME SRC=""BudgetSearch.asp?Section=Money&"
						If Len(oRequest("ModifyMoneys").Item) > 0 Then
							Redim aTempBudgetComponent(N_BUDGET_COMPONENT_SIZE)
							For iIndex = N_AREA_ID_BUDGET To N_MONTH_BUDGET
								aTempBudgetComponent(iIndex) = aBudgetComponent(iIndex)(0)
							Next
							Response.Write GetMoneyAsURL(oRequest, oADODBConnection, aTempBudgetComponent, sErrorDescription) & "&Change=1&View=1"
						ElseIf StrComp(SERVER_NAME_FOR_LICENSE, "CASTOR", vbBinaryCompare) = 0 Then
							Response.Write "AreaID=120&ProgramDutyID=2&FundID=0&DutyID=2&ActiveDutyID=3&SpecificDutyID=-1&ProgramID=56&RegionID=0&BudgetUR=120&BudgetCT=123&BudgetAUX=0&LocationID=1&BudgetID1=1507&BudgetID2=10086&BudgetID3=20091&ConfineTypeID=1&ActivityID1=2&ActivityID2=40&ProcessID=1&BudgetYear=2010&Change=1&View=1"
						End If
					Response.Write """ NAME=""Search01IFrame"" FRAMEBORDER=""0"" WIDTH=""480"" HEIGHT=""800""></IFRAME>"
				Response.Write "</TD>"
				Response.Write "<TD>&nbsp;</TD>"
				Response.Write "<TD BGCOLOR=""#" & S_MAIN_COLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
				Response.Write "<TD>&nbsp;&nbsp;&nbsp;</TD>"
				Response.Write "<TD WIDTH=""600"" VALIGN=""TOP"">"
					Response.Write "<IFRAME SRC=""BudgetSearch.asp?Section=Money&"
						If Len(oRequest("ModifyMoneys").Item) > 0 Then
							Redim aTempBudgetComponent(N_BUDGET_COMPONENT_SIZE)
							For iIndex = N_AREA_ID_BUDGET To N_MONTH_BUDGET
								aTempBudgetComponent(iIndex) = aBudgetComponent(iIndex)(1)
							Next
							Response.Write GetMoneyAsURL(oRequest, oADODBConnection, aTempBudgetComponent, sErrorDescription) & "&Change=1&View=1"
						ElseIf StrComp(SERVER_NAME_FOR_LICENSE, "CASTOR", vbBinaryCompare) = 0 Then
							Response.Write "AreaID=107&ProgramDutyID=2&FundID=0&DutyID=2&ActiveDutyID=3&SpecificDutyID=-1&ProgramID=56&RegionID=0&BudgetUR=107&BudgetCT=0&BudgetAUX=0&LocationID=1&BudgetID1=1103&BudgetID2=10002&BudgetID3=20002&ConfineTypeID=1&ActivityID1=2&ActivityID2=42&ProcessID=1&BudgetYear=2010&Change=1&View=1"
						End If
					Response.Write """ NAME=""Search02IFrame"" FRAMEBORDER=""0"" WIDTH=""480"" HEIGHT=""800""></IFRAME>"
				Response.Write "</TD>"
			Response.Write "</TR></TABLE>" & vbNewLine
			lErrorNumber = DisplayMoneysForm(oRequest, oADODBConnection, GetASPFileName(""), aBudgetComponent, sErrorDescription)
		Else
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
				Response.Write "<TD WIDTH=""600"" VALIGN=""TOP"">"
					Select Case sSection
						Case "Budget"
							If Len(oRequest("View").Item) > 0 Then
								aBudgetComponent(S_QUERY_CONDITION_BUDGET) = " And (Budgets.ParentID=" & aBudgetComponent(N_ID_BUDGET) & ")"
							Else
								aBudgetComponent(S_QUERY_CONDITION_BUDGET) = " And (Budgets.ParentID=" & aBudgetComponent(N_PARENT_ID_BUDGET) & ")"
							End If
							lErrorNumber = DisplayBudgetTable(oRequest, oADODBConnection, DISPLAY_NOTHING, (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)), aBudgetComponent, sErrorDescription)
						Case "Money"
							lErrorNumber = DisplayMoneySearchForm(oRequest, oADODBConnection, GetASPFileName(""), aBudgetComponent, sErrorDescription)
							If lErrorNumber = 0 Then
								If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
									Response.Write "<DIV NAME=""EntriesDiv"" ID=""EntriesDiv"" STYLE=""width: 600px; height: 300px; overflow: auto;"">"
										lErrorNumber = DisplayMoneyTable(oRequest, oADODBConnection, (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)), aBudgetComponent, sErrorDescription)
									Response.Write "</DIV>"
								End If
							End If
						Case "Program"
							If Len(oRequest("View").Item) > 0 Then
								aBudgetComponent(S_QUERY_CONDITION_BUDGET) = " And (BudgetsAndPrograms.ParentID=" & aBudgetComponent(N_ID_BUDGET) & ")"
							Else
								aBudgetComponent(S_QUERY_CONDITION_BUDGET) = " And (BudgetsAndPrograms.ParentID=" & aBudgetComponent(N_PARENT_ID_BUDGET) & ")"
							End If
							lErrorNumber = DisplayProgramTable(oRequest, oADODBConnection, DISPLAY_NOTHING, (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)), aBudgetComponent, sErrorDescription)
					End Select
					Response.Write "<BR />"
					If (lErrorNumber <> 0) And (lErrorNumber <> L_ERR_NO_RECORDS) Then
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						sErrorDescription = ""
					End If
				Response.Write "</TD>"
				Response.Write "<TD>&nbsp;</TD>"
				Response.Write "<TD BGCOLOR=""#" & S_MAIN_COLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
				Response.Write "<TD>&nbsp;&nbsp;&nbsp;</TD>"
				Response.Write "<TD VALIGN=""TOP"">"
					Select Case sSection
						Case "Budget"
							If Len(oRequest("View").Item) > 0 Then
								aBudgetComponent(N_PARENT_ID_BUDGET) = aBudgetComponent(N_ID_BUDGET)
								aBudgetComponent(N_ID_BUDGET) = -1
							End If
							lErrorNumber = DisplayBudgetForm(oRequest, oADODBConnection, GetASPFileName(""), aBudgetComponent, sErrorDescription)
						Case "Money"
							lErrorNumber = DisplayMoneyForm(oRequest, oADODBConnection, GetASPFileName(""), aBudgetComponent, sErrorDescription)
						Case "Program"
							If Len(oRequest("View").Item) > 0 Then
								aBudgetComponent(N_PARENT_ID_BUDGET) = aBudgetComponent(N_ID_BUDGET)
								aBudgetComponent(N_ID_BUDGET) = -1
							End If
							lErrorNumber = DisplayProgramForm(oRequest, oADODBConnection, GetASPFileName(""), aBudgetComponent, sErrorDescription)
					End Select
				Response.Write "</TD>"
			Response.Write "</TR></TABLE>"
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