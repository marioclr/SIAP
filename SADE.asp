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
<!-- #include file="Libraries/ReportsLib.asp" -->
<!-- #include file="Libraries/SADELibrary.asp" -->
<!-- #include file="Libraries/SADECourseComponent.asp" -->
<!-- #include file="Libraries/SADEProfileComponent.asp" -->
<%
Dim iSectionID
Dim iStep
Dim lCourseID
Dim lEmployeeID
Dim bAction
Dim sNames

If B_ISSSTE Then
Else
	If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_SADE_PERMISSIONS) = N_SADE_PERMISSIONS Then
	Else
		Response.Redirect "AccessDenied.asp?Permission=" & N_SADE_PERMISSIONS
	End If
End If

iSectionID = -1
iStep = 1
lCourseID = -1
lEmployeeID = -1
If Len(oRequest("SectionID").Item) > 0 Then iSectionID = CInt(oRequest("SectionID").Item)
If Len(oRequest("Step").Item) > 0 Then iStep = CInt(oRequest("Step").Item)
If Len(oRequest("CourseID").Item) > 0 Then lCourseID = CLng(oRequest("CourseID").Item)
If Len(oRequest("EmployeeID").Item) > 0 Then lEmployeeID = CLng(oRequest("EmployeeID").Item)
bAction = (Len(oRequest("Add").Item) > 0) Or (Len(oRequest("Modify").Item) > 0) Or (Len(oRequest("Remove").Item) > 0) Or (Len(oRequest("SetActive").Item) > 0)

Call InitializeCourseComponent(oRequest, aCourseComponent)
Call InitializeProfileComponent(oRequest, aProfileComponent)

aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
If bAction Then
	lErrorNumber = DoSADEActions(oRequest, iSectionID, sErrorDescription)
End If
Select Case iSectionID
	Case 361
		If Len(oRequest("Diploma").Item) = 0 Then
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Registro de lista de cursos"
		Else
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Registro de diplomados"
		End If
	Case 362
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Cursos del programa anual de capacitación"
	Case 363
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Registro de personal para capacitación"
	Case 364
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Reportes"
	Case 365
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Seguimiento al programa autorizado de capacitación"
	Case 366
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Generación de diplomas de reconocimiento"
	Case 367
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Reporte de curriculum por empleado"
End Select

bWaitMessage = False
Response.Cookies("SoS_SectionID") = 36
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/Export.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript"><!--
			function CheckRadioList(oField){
				if (oField.checked) {
					return true;
				}
				for (var i=0; i<oField.length; i++) {
					if (oField[i].checked) {
						return true;
					}
				}

				alert('Favor de seleccionar un registro.');
				return false;
			} // End of CheckRadioList
		//--></SCRIPT>
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%Select Case iSectionID
			Case 361
				If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS) = N_ADD_PERMISSIONS Then
					aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Registrar un nuevo curso",_
							  "",_
							  "", "SADE.asp?SectionID=" & oRequest("SectionID").Item & "&New=1", (Len(oRequest("Diploma").Item) = 0)),_
						Array("Registrar un nuevo diplomado",_
							  "",_
							  "", "SADE.asp?SectionID=" & oRequest("SectionID").Item & "&New=1&Diploma=1", (Len(oRequest("Diploma").Item) > 0))_
					)
					aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 793
					aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
					aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 200
				End If
			Case 362
				If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS) = N_ADD_PERMISSIONS Then
					aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Registrar un nuevo curso",_
							  "",_
							  "", "SADE.asp?SectionID=" & oRequest("SectionID").Item & "&New=1", True)_
					)
					aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 793
					aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
					aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 200
				End If
			Case 365
				aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
					Array("Exportar a Excel",_
						  "",_
						  "", "javascript: OpenNewWindow('Export.asp?Action=Reports&Excel=1&ReportID=" & ISSSTE_1365_REPORTS & "&CourseID=" & lCourseID & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
				)
				aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 843
				aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
				aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 150
			Case 367
				aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
					Array("Exportar a Word",_
						  "",_
						  "", "javascript: OpenNewWindow('Export.asp?Action=Reports&Word=1&ReportID=" & ISSSTE_1367_REPORTS & "&EmployeeID=" & lEmployeeID & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToWord', 640, 480, 'yes', 'yes')", True)_
				)
				aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 843
				aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
				aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 150
		End Select%>
		<!-- #include file="_Header.asp" -->
		<%Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> > <A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=36"">Desarrollo humano</A> > "
		Select Case iSectionID
			Case Else
				Response.Write "<B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B>"
		End Select
		Response.Write "<BR /><BR />"

		If lErrorNumber <> 0 Then
			Response.Write "<BR />"
			Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
			lErrorNumber = 0
			sErrorDescription = ""
			Response.Write "<BR />"
		ElseIf bAction Then
			Call DisplayErrorMessage("Confirmación", "La información se guardó con éxito.")
			Response.Write "<BR />"
		End If

		Select Case iSectionID
			Case 361
				If Len(oRequest("Diploma").Item) > 0 Then aProfileComponent(S_QUERY_CONDITION_PROFILE) = " And (ID_Padre=0)"
				Response.Write "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
					Response.Write "<TD WIDTH=""600"" VALIGN=""TOP"">"
						lErrorNumber = DisplayProfilesTable(oRequest, oSIAPSADEADODBConnection, DISPLAY_NOTHING, True, sErrorDescription)
					Response.Write "</TD>"
					Response.Write "<TD>&nbsp;</TD>"
					Response.Write "<TD BGCOLOR=""" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""1"" ><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
					Response.Write "<TD>&nbsp;</TD>"
					Response.Write "<TD WIDTH=""100%"" VALIGN=""TOP"">"
						lErrorNumber = DisplayProfilesForm(oRequest, oSIAPSADEADODBConnection, aProfileComponent, sErrorDescription)
					Response.Write "</TD>"
				Response.Write "</TR></TABLE>"
			Case 362
				Response.Write "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
					Response.Write "<TD WIDTH=""600"" VALIGN=""TOP"">"
						lErrorNumber = DisplayCoursesTable(oRequest, oSIAPSADEADODBConnection, True, DISPLAY_NOTHING, True, sErrorDescription)
					Response.Write "</TD>"
					Response.Write "<TD>&nbsp;</TD>"
					Response.Write "<TD BGCOLOR=""" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""1"" ><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
					Response.Write "<TD>&nbsp;</TD>"
					Response.Write "<TD WIDTH=""100%"" VALIGN=""TOP"">"
						lErrorNumber = DisplayCoursesForm(oRequest, oSIAPSADEADODBConnection, aCourseComponent, sErrorDescription)
					Response.Write "</TD>"
				Response.Write "</TR></TABLE>"
			Case 363
				Select Case iStep
					Case 2
						Call GetNameFromTable(oADODBConnection, "SADE_Curso", lCourseID, "", "", sNames, sErrorDescription)
						Call DisplayInstructionsMessage("Instrucciones", "<B>Paso 2.</B> Utilizando el número de empleado, registre a los empleados que tomarán el curso <B>""" & sNames & """</B>.")
						lErrorNumber = DisplayCourseRegistrationForm(oRequest, oSIAPSADEADODBConnection, lCourseID, sErrorDescription)
					Case Else
						Call DisplayInstructionsMessage("Instrucciones", "<B>Paso 1.</B> Seleccione el curso para registrar a los empleados que lo tomarán.")
						Response.Write "<BR /><FORM NAME=""SADEFrm"" ID=""SADEFrm"" ACTION=""SADE.asp"" METHOD=""GET"" onSubmit=""return CheckRadioList(this.CourseID)"">"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""" & iSectionID & """ />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""2"" />"
							lErrorNumber = DisplayCoursesTable(oRequest, oSIAPSADEADODBConnection, False, DISPLAY_RADIO_BUTTONS, False, sErrorDescription)
							If lErrorNumber = 0 Then
								Response.Write "<BR />&nbsp;&nbsp;&nbsp;<INPUT TYPE=""SUBMIT"" NAME=""Continue"" ID=""ContinueBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
							ElseIf lErrorNumber = L_ERR_NO_RECORDS Then
								Call DisplayErrorMessage("No existen cursos registrados", "Para inscribir a los empleados en los cursos, es necesario que se den de alta los cursos con fecha de inicio a futuro.")
								lErrorNumber = 0
								sErrorDescription = ""
								Response.Write "<BR />"
							End If
						Response.Write "</FORM>"
				End Select
			Case 365
				Select Case iStep
					Case 2
						Call GetNameFromTable(oADODBConnection, SADE_PREFIX & "Curso", lCourseID, "", "", sNames, sErrorDescription)
						Call DisplayInstructionsMessage("Instrucciones", "<B>Paso 2.</B> Registre las calificaciones obtenidas por los empleados para el curso <B>""" & sNames & """</B>.")
						Response.Write "<BR />"
						Response.Write "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
							Response.Write "<TD WIDTH=""600"" VALIGN=""TOP"">"
								lErrorNumber = DisplayCourseEmployeesTable(oRequest, oSIAPSADEADODBConnection, lCourseID, False, False, sErrorDescription)
							Response.Write "</TD>"
							Response.Write "<TD>&nbsp;</TD>"
							Response.Write "<TD BGCOLOR=""" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""1"" ><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
							Response.Write "<TD>&nbsp;</TD>"
							Response.Write "<TD WIDTH=""100%"" VALIGN=""TOP"">"
								If lEmployeeID > -1 Then
									lErrorNumber = DisplayGradesForm(oRequest, oSIAPSADEADODBConnection, lCourseID, lEmployeeID, sErrorDescription)
								End If
							Response.Write "</TD>"
						Response.Write "</TR></TABLE>"
					Case Else
						Call DisplayInstructionsMessage("Instrucciones", "<B>Paso 1.</B> Seleccione el curso al que desea darle seguimiento.")
						Response.Write "<BR /><FORM NAME=""SADEFrm"" ID=""SADEFrm"" ACTION=""SADE.asp"" METHOD=""GET"" onSubmit=""return CheckRadioList(this.CourseID)"">"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""" & iSectionID & """ />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""2"" />"
							lErrorNumber = DisplayCoursesTable(oRequest, oSIAPSADEADODBConnection, False, DISPLAY_RADIO_BUTTONS, False, sErrorDescription)
							If lErrorNumber = 0 Then
								Response.Write "<BR />&nbsp;&nbsp;&nbsp;<INPUT TYPE=""SUBMIT"" NAME=""Continue"" ID=""ContinueBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
							End If
						Response.Write "</FORM>"
				End Select
			Case 366
				Select Case iStep
					Case 2
						Call GetNameFromTable(oADODBConnection, SADE_PREFIX & "Curso", lCourseID, "", "", sNames, sErrorDescription)
						Call DisplayInstructionsMessage("Instrucciones", "<B>Paso 2.</B> Seleccione al empleado que se le imprimirá su constancia por haber <B>aprobado</B> el curso <B>""" & sNames & """</B>.")
						Response.Write "<BR />"
						Response.Write "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
							Response.Write "<TD WIDTH=""600"" VALIGN=""TOP"">"
								lErrorNumber = DisplayCourseEmployeesTable(oRequest, oSIAPSADEADODBConnection, lCourseID, True, False, sErrorDescription)
							Response.Write "</TD>"
							Response.Write "<TD>&nbsp;</TD>"
							Response.Write "<TD BGCOLOR=""" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""1"" ><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
							Response.Write "<TD>&nbsp;</TD>"
							Response.Write "<TD WIDTH=""100%"" VALIGN=""TOP"">"
								If lEmployeeID > -1 Then
									lErrorNumber = DisplayCertificateForm(oRequest, oSIAPSADEADODBConnection, lCourseID, lEmployeeID, sErrorDescription)
								End If
							Response.Write "</TD>"
						Response.Write "</TR></TABLE>"
						If bAction Then
							lErrorNumber = PrintEmployeeCertificate(oRequest, oSIAPSADEADODBConnection, lCourseID, lEmployeeID, sErrorDescription)
						End If
					Case Else
						Call DisplayInstructionsMessage("Instrucciones", "<B>Paso 1.</B> Seleccione el curso para imprimir las constancias de los empleados.")
						Response.Write "<BR /><FORM NAME=""SADEFrm"" ID=""SADEFrm"" ACTION=""SADE.asp"" METHOD=""GET"" onSubmit=""return CheckRadioList(this.CourseID)"">"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""" & iSectionID & """ />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""2"" />"
							lErrorNumber = DisplayCoursesTable(oRequest, oSIAPSADEADODBConnection, False, DISPLAY_RADIO_BUTTONS, False, sErrorDescription)
							If lErrorNumber = 0 Then
								Response.Write "<BR />&nbsp;&nbsp;&nbsp;<INPUT TYPE=""SUBMIT"" NAME=""Continue"" ID=""ContinueBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
							End If
						Response.Write "</FORM>"
				End Select
			Case 367
				Select Case iStep
					Case 2
						lErrorNumber = DisplayEmployeeCurriculum(oRequest, oSIAPSADEADODBConnection, lEmployeeID, False, sErrorDescription)
					Case Else
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							Response.Write "function CheckEmployeeForm(oForm){" & vbNewLine
								Response.Write "if (oForm) {" & vbNewLine
									Response.Write "if (oForm.EmployeeID.value == '') {" & vbNewLine
										Response.Write "alert('Introduzca el número del empleado');" & vbNewLine
										Response.Write "return false;" & vbNewLine
										Response.Write "oForm.EmployeeID.focus();" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "}" & vbNewLine

								Response.Write "return true;" & vbNewLine
							Response.Write "} // End of CheckEmployeeForm" & vbNewLine
						Response.Write "//--></SCRIPT>" & vbNewLine
						Call DisplayInstructionsMessage("Instrucciones", "<B>Paso 1.</B> Introduzca el número de empleado para mostrar su currículum." &_
						"<BR /><FORM NAME=""SADEFrm"" ID=""SADEFrm"" ACTION=""SADE.asp"" METHOD=""GET"" onSubmit=""return CheckEmployeeForm(this)"">" &_
							"<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""" & iSectionID & """ />" &_
							"<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""2"" />" &_
							"<IMG SRC=""Images/Transparent.gif"" WIDTH=""32"" HEIGHT=""20"" ALIGN=""LEFT"" HSPACE=""5"" />" &_
							"No. de empleado: <INPUT TYPE=""TEXT"" NAME=""EmployeeID"" ID=""EmployeeIDTxt"" SIZE=""6"" MAXLENGTH=""6"" CLASS=""TextFields"" />" &_
							"&nbsp;&nbsp;&nbsp;<INPUT TYPE=""SUBMIT"" NAME=""Continue"" ID=""ContinueBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />" &_
						"</FORM>")
				End Select
		End Select

		If lErrorNumber <> 0 Then
			Response.Write "<BR />"
			Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
			lErrorNumber = 0
			sErrorDescription = ""
			Response.Write "<BR />"
		End If
		%>
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>