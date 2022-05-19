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
<!-- #include file="Libraries/AbsenceComponent.asp" -->
<!-- #include file="Libraries/ConceptComponent.asp" -->
<!-- #include file="Libraries/EmployeesLib.asp" -->
<!-- #include file="Libraries/EmployeeComponent.asp" -->
<!-- #include file="Libraries/JobComponent.asp" -->
<!-- #include file="Libraries/JobsLib.asp" -->
<!-- #include file="Libraries/PaymentsLib.asp" -->
<!-- #include file="Libraries/PaymentComponent.asp" -->
<!-- #include file="Libraries/ReportsLib.asp" -->
<!-- #include file="Libraries/ReportComponent.asp" -->
<!-- #include file="Libraries/UploadInfoLibrary.asp" -->
<!-- #include file="Libraries/ZIPLibrary.asp" -->
<%
Dim iSelectedTab
Dim iActionPading
Dim bAction
Dim bError
Dim sError
Dim sNames
Dim lReasonID
Dim lStatusID
Dim lEmployeeID
Dim lSuccess

If Len(oRequest("SuccessID").Item) > 0 Then	lSuccess = CLng(oRequest("SuccessID").Item)
If Len(oRequest("ReasonID").Item) > 0 Then	lReasonID = CLng(oRequest("ReasonID").Item)
If Len(oRequest("EmployeeID").Item) > 0 Then lEmployeeID = CLng(oRequest("EmployeeID").Item)
If Len(oRequest("StatusID").Item) > 0 Then aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = CLng(oRequest("StatusID").Item)
If (Len(oRequest("SaveChanges").Item) > 0) Then bAction = True

If B_ISSSTE Then
Else
	If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_EMPLOYEES_PERMISSIONS) = N_EMPLOYEES_PERMISSIONS Then
	Else
		Response.Redirect "AccessDenied.asp?Permission=" & N_EMPLOYEES_PERMISSIONS
	End If
End If

Select Case CInt(Request.Cookies("SIAP_SectionID"))
	Case 1
		Response.Cookies("SIAP_SubSectionID") = 11
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
	Case 4
		Response.Cookies("SIAP_SubSectionID") = 14
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
	Case 7
		Response.Cookies("SIAP_SubSectionID") = 17
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
	Case Else
		Response.Cookies("SIAP_SubSectionID") = 12
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
End Select
If (Len(oRequest("New").Item) > 0) And (CInt(oRequest("StatusID").Item) = -2) Then
	aHeaderComponent(S_TITLE_NAME_HEADER) = "Asignar número de empleado"
ElseIf B_ISSSTE Then
	aHeaderComponent(S_TITLE_NAME_HEADER) = "Consulta de personal"
Else
	aHeaderComponent(S_TITLE_NAME_HEADER) = "Empleados"
End If
bWaitMessage = True

Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
Call InitializeReportsComponent(oRequest, aReportsComponent)
Call GetEmployeesURLValues(oRequest, iSelectedTab, bAction, aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE))

bError = False
If (Len(oRequest("Register").Item) > 0) Or (Len(oRequest("Validate").Item) > 0) Or (Len(oRequest("Authorization").Item) > 0)  Or (Len(oRequest("RemoveValidate").Item) > 0)  Or (Len(oRequest("RemoveAuthorization").Item) > 0) Or (Len(oRequest("RemoveApply").Item) > 0) Or (Len(oRequest("LicenseValidate").Item) > 0)  Or (Len(oRequest("LicenseAuthorization").Item) > 0) Or (Len(oRequest("LicenseApply").Item) > 0)Or (Len(oRequest("ResumptionOfWorkValidate").Item) > 0)  Or (Len(oRequest("ResumptionOfWorkAuthorization").Item) > 0) Or (Len(oRequest("ResumptionOfWorkApply").Item) > 0) Or (Len(oRequest("RemoveMotion").Item) > 0) Then bAction = True

If bAction Then
	lErrorNumber = DoEmployeesAction(oRequest, oADODBConnection, oRequest("Action").Item, sErrorDescription)
	sError = sErrorDescription
	bError = (lErrorNumber <> 0)
	If (lErrorNumber = 0) And (Len(oRequest("Remove").Item) > 0) Then
		bAction = False
	End If
	If  (Len(oRequest("RemoveValidate").Item) > 0)  Or (Len(oRequest("RemoveAuthorization").Item) > 0) Or (Len(oRequest("RemoveApply").Item) > 0) Then
		bAction = False
		Response.Redirect "Main_ISSSTE.asp?SectionID=18"
	End If
	If  (Len(oRequest("LicenseValidate").Item) > 0)  Or (Len(oRequest("LicenseAuthorization").Item) > 0) Or (Len(oRequest("LicenseApply").Item) > 0) Then
		bAction = False
	End If
End If
If aEmployeeComponent(N_ID_EMPLOYEE) > -1 Then
	lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
End If
Response.Cookies("SoS_SectionID") = 193
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/Export.js"></SCRIPT>
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%If ((Len(oRequest("DoSearch").Item) > 0) Or (aEmployeeComponent(N_ID_EMPLOYEE) > -1)) And (CLng(oRequest("ReportID").Item) <> ISSSTE_1003_REPORTS) And (StrComp(oRequest("Tab").Item, "8", vbBinaryCompare) <> 0) Then
			aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Agregar un nuevo empleado",_
					  "",_
					  "", "Employees.asp?New=1", ((aEmployeeComponent(N_ID_EMPLOYEE) = -1) And ((Len(oRequest("AreaID").Item) > 0) Or (Len(oRequest("PositionID").Item) > 0)))),_
				Array("Agregar un registro al historial",_
					  "",_
					  "", "Employees.asp?EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&EmployeeDate=0&Tab=6&Change=1&ReportID=707", (((aLoginComponent(N_PROFILE_ID_LOGIN) <= 0) Or (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ModificacionDeAntiguedades & ",", vbBinaryCompare) > 0)) And (StrComp(oRequest("ReportID").Item, "707", vbBinaryCompare) = 0) And (StrComp(oRequest("Tab").Item, "6", vbBinaryCompare) = 0))),_
				Array("<LINE />",_
					  "",_
					  "", "", ((Len(oRequest("New").Item) = 0) And ((aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (Len(oRequest("PositionID").Item) > 0)))),_
				Array("Exportar a Excel",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=Employees&Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(RemoveParameterFromURLString(oRequest, "Action")) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (Len(oRequest("DoSearch").Item) > 0)),_
				Array("Exportar a Excel",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=Employees&Excel=1&Tab=" & iSelectedTab & "&EmployeeID=" & oRequest("EmployeeID").Item & "&FilterStartYear=" & oRequest("FilterStartYear").Item & "&FilterStartMonth=" & oRequest("FilterStartMonth").Item & "&FilterStartDay=" & oRequest("FilterStartDay").Item & "&FilterEndYear=" & oRequest("FilterEndYear").Item & "&FilterEndMonth=" & oRequest("FilterEndMonth").Item & "&FilterEndDay=" & oRequest("FilterEndDay").Item & "&PayrollID=" & oRequest("PayrollID").Item & "&ReportID=" & oRequest("ReportID").Item & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", ((aEmployeeComponent(N_ID_EMPLOYEE) > -1) And (InStr(1, ",3,4,6,", iSelectedTab, vbBinaryCompare) > 0))),_
				Array("Exportar a Excel",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=Reports&Excel=1&Tab=" & iSelectedTab & "&EmployeeID=" & oRequest("EmployeeID").Item & "&ReportID=" & oRequest("ReportID").Item & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", ((aEmployeeComponent(N_ID_EMPLOYEE) > -1) And (InStr(1, ",7,", iSelectedTab, vbBinaryCompare) > 0) And (CLng(oRequest("ReportID").Item) = ISSSTE_1116_REPORTS))),_
				Array("Exportar a Texto",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=Reports&Word=1&Tab=" & iSelectedTab & "&EmployeeID=" & oRequest("EmployeeID").Item & "&StartYear=" & oRequest("StartYear").Item & "&StartMonth=" & oRequest("StartMonth").Item & "&StartDay=" & oRequest("StartDay").Item & "&PayrollID=" & oRequest("PayrollID").Item & "&ReportID=" & oRequest("ReportID").Item & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", ((aEmployeeComponent(N_ID_EMPLOYEE) > -1) And (InStr(1, ",7,", iSelectedTab, vbBinaryCompare) > 0) And (CLng(oRequest("ReportID").Item) = ISSSTE_1002_REPORTS))),_
				Array("Exportar a Word",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=Reports&Word=1&Tab=" & iSelectedTab & "&EmployeeID=" & oRequest("EmployeeID").Item & "&ConceptID=" & oRequest("ConceptID").Item & "&StartYear=" & oRequest("StartYear").Item & "&StartMonth=" & oRequest("StartMonth").Item & "&StartDay=" & oRequest("StartDay").Item & "&EndYear=" & oRequest("EndYear").Item & "&EndMonth=" & oRequest("EndMonth").Item & "&EndDay=" & oRequest("EndDay").Item & "&ReportID=" & oRequest("ReportID").Item & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", ((aEmployeeComponent(N_ID_EMPLOYEE) > -1) And (InStr(1, ",7,", iSelectedTab, vbBinaryCompare) > 0) And (CLng(oRequest("ReportID").Item)=ISSSTE_1208_REPORTS))),_
				Array("Exportar a Word",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=Employees&Word=1&Tab=" & iSelectedTab & "&EmployeeID=" & oRequest("EmployeeID").Item & "&StartYear=" & oRequest("StartYear").Item & "&StartMonth=" & oRequest("StartMonth").Item & "&StartDay=" & oRequest("StartDay").Item & "&PayrollID=" & oRequest("PayrollID").Item & "&ReportID=" & oRequest("ReportID").Item & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", ((Len(oRequest("DoSearch").Item) = 0) And (aEmployeeComponent(N_ID_EMPLOYEE) > -1) And (InStr(1, ",1,2,5,", iSelectedTab, vbBinaryCompare) > 0))),_
                Array("Exportar reporte de recursos humanos",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=Employees&Web=1&Tab=" & iSelectedTab & "&EmployeeID=" & oRequest("EmployeeID").Item & "&StartYear=" & oRequest("StartYear").Item & "&StartMonth=" & oRequest("StartMonth").Item & "&StartDay=" & oRequest("StartDay").Item & "&PayrollID=" & oRequest("PayrollID").Item & "&ReportID=" & oRequest("ReportID").Item & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&ReportRH=++Reporte++', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", ((Len(oRequest("DoSearch").Item) = 0) And (aEmployeeComponent(N_ID_EMPLOYEE) > -1) And (InStr(1, ",1,2,5,", iSelectedTab, vbBinaryCompare) > 0))),_
				Array("Imprimir",_
					  "",_
					  "", "javascript: SendReportToPrint('ReportDiv', '" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "')", False)_
			)
			aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 793
			aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
			aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 200
		End If%>
		<!-- #include file="_Header.asp" -->
		<%Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> > "
			If B_ISSSTE Then
				If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > "
					sNames = "Consulta de personal"
				ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > "
					sNames = "Consulta de personal"
				ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=71"">Personal</A> > "
					sNames = "Consulta de personal"
				Else
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > "
					sNames = "Consulta de personal"
				End If
			Else
				Response.Write "<A HREF=""HumanResources.asp"">Personal</A> > "
				sNames = "Empleados"
			End If
			If ((Len(oRequest("New").Item) > 0) Or (Len(oRequest("Add").Item) > 0)) And (CInt(oRequest("StatusID").Item) = -2) Then
				Response.Write "<B>Asignación de número de empleado</B>"
			ElseIf (aEmployeeComponent(N_ID_EMPLOYEE) <> -1) And (Len(oRequest("DoSearch").Item) > 0) Then
				Response.Write "<A HREF=""Employees.asp"">Consulta de personal</A> > Resultado de la búsqueda</B><BR /><BR />"
			ElseIf (aEmployeeComponent(N_ID_EMPLOYEE) <> -1) And (Len(oRequest("DoSearch").Item) = 0) Then
				Response.Write "<A HREF=""Employees.asp"
				If Len(Request.Cookies("SIAP_SearchPath").Item) > 0 Then Response.Write "?" & Request.Cookies("SIAP_SearchPath").Item
				Response.Write """>" & sNames & "</A> > <B>" & CleanStringForHTML(aEmployeeComponent(S_NUMBER_EMPLOYEE) & ". " & aEmployeeComponent(S_NAME_EMPLOYEE) & " " & aEmployeeComponent(S_LAST_NAME_EMPLOYEE) & " " & aEmployeeComponent(S_LAST_NAME2_EMPLOYEE)) & "</B><BR /><BR />"
			Else
				Response.Write "<B>" & sNames & "</B><BR /><BR />"
			End If
		If Len(oRequest("Search").Item) > 0 Then
			Call DisplayEmployeesSearchForm(oRequest, oADODBConnection, GetASPFileName(""), True, sErrorDescription)
		ElseIf (Len(oRequest("New").Item) > 0) Or bError Then
			If bAction And ((lErrorNumber <> 0) Or (Len(sError) > 0))Then
				Response.Write "<BR /><BR />"
				If Len(sError) > 0 Then sErrorDescription = sError
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
				lErrorNumber = 0
				Response.Write "<BR />"
			End If
			Call DisplayEmployeesTabs(oRequest, bError, sErrorDescription)
			Response.Write "<BR />"
			lErrorNumber = DisplayEmployeeForms(oRequest, iSelectedTab, (CInt(oRequest("StatusID").Item) <> -2), sErrorDescription)
		ElseIf (Len(oRequest("Tab").Item) > 0) Or bAction Then
			If bAction Then
				Response.Write "<BR />"
				If lErrorNumber = 0 Then
					If (Len(oRequest("Add").Item) > 0) And (CInt(oRequest("StatusID").Item) = -2) Then
						Call DisplayErrorMessage("Confirmación", "Se asignó el número al nuevo empleado. <BR /><BR /><BR /><FORM ACTION=""Employees.asp"" METHOD=""GET""><INPUT TYPE=""HIDDEN"" NAME=""New"" VALUE=""1"" /><INPUT TYPE=""HIDDEN"" NAME=""StatusID"" VALUE=""-2"" />&nbsp;&nbsp;<INPUT TYPE=""SUBMIT"" VALUE=""Asignar otro número"" CLASS=""Buttons"" /><IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" /><INPUT TYPE=""BUTTON"" VALUE=""Regresar"" CLASS=""Buttons"" onClick=""javascript: window.location.href='Main_ISSSTE.asp?SectionID=1'"" /></FORM>")
					Else
						Call DisplayErrorMessage("Confirmación", "La información del empleado fue guardada con éxito.")
					End If
					lErrorNumber = 0
				Else
					Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
					lErrorNumber = 0
				End If
				Response.Write "<BR />"
			End If
			Call DisplayEmployeesTabs(oRequest, bError, sErrorDescription)
			Response.Write "<BR />"
			lErrorNumber = DisplayEmployeeForms(oRequest, iSelectedTab, (CInt(oRequest("StatusID").Item) <> -2), sErrorDescription)
			If lErrorNumber <> 0 Then
				Response.Write "<BR /><BR />"
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
				lErrorNumber = 0
				Response.Write "<BR />"
			End If
		ElseIf Len(oRequest("DoSearch").Item) > 0 Then
			lErrorNumber = DisplayEmployeesTable(oRequest, oADODBConnection, DISPLAY_NOTHING, True, False, aEmployeeComponent, sErrorDescription)
			If lErrorNumber <> 0 Then
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
				lErrorNumber = 0
				Response.Write "<BR />"
				lErrorNumber = DisplayEmployeesSearchForm(oRequest, oADODBConnection, GetASPFileName(""), True, sErrorDescription)
			End If
		Else
			If Len(oRequest("Remove").Item) > 0 Then
				Call DisplayErrorMessage("Confirmación", "La información del empleado fue eliminada con éxito.")
			End If
			Response.Write "<TABLE WIDTH=""720"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					If False Then
					Response.Write "<TD WIDTH=""50%"" COLSPAN=""2"" VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><B>Consulta de empleados</B></FONT></TD>"
					End If
					If B_ISSSTE Then
						Response.Write "<TD WIDTH=""50%"" COLSPAN=""2"" VALIGN=""TOP""></TD>"
					Else
						Response.Write "<TD WIDTH=""50%"" COLSPAN=""2"" VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><B><A HREF=""Employees.asp?New=1"">ALTA DE EMPLEADOS</A></B></FONT></TD>"
					End If
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD ROWSPAN=""2"">&nbsp;&nbsp;&nbsp;</TD>"
					Response.Write "<TD WIDTH=""50%"" VALIGN=""TOP"">"
						Call DisplayEmployeesSearchForm(oRequest, oADODBConnection, GetASPFileName(""), False, sErrorDescription)
					Response.Write "</TD>"
					Response.Write "<TD ROWSPAN=""2"">&nbsp;&nbsp;&nbsp;</TD>"
					Response.Write "<TD WIDTH=""50%"" VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"
						If B_ISSSTE And False Then
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								Response.Write "function CheckNewEmployeeFields(oForm) {" & vbNewLine
									Response.Write "if (oForm) {" & vbNewLine
										Response.Write "if (oForm.EmployeeNumber.value == '') {" & vbNewLine
											Response.Write "alert('Favor de introducir el número de empleado.');" & vbNewLine
											Response.Write "oForm.EmployeeNumber.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "return true;" & vbNewLine
								Response.Write "} // End of CheckNewEmployeeFields" & vbNewLine
							Response.Write "//--></SCRIPT>"

							Response.Write "<FORM NAME=""NewEmployeeFrm"" ID=""NewEmployeeFrm"" ACTION=""Employees.asp"" METHOD=""GET"" onSubmit=""return CheckNewEmployeeFields(this)"">"
								Response.Write "Número del empleado: <INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & oRequest("EmployeeNumber").Item & """ CLASS=""TextFields"" /><BR />"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StatusID"" ID=""StatusIDHdn"" VALUE=""-2"" />"
								Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""DoSearch"" ID=""DoSearchBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
							Response.Write "</FORM>"
						Else
							'Response.Write "<A HREF=""Employees.asp?New=1"">Registre la información del empleado de nuevo ingreso</A> e indique la plaza que va a ocupar.<BR />"
						End If
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE>"

			Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""980"" HEIGHT=""1"" /><BR /><BR />"
			aMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Asignación de número de empleado",_
					  "Asigne un número al nuevo empleado antes de darlo de alta.",_
					  "Images/MnLeftArrows.gif", "Employees.asp?New=1&StatusID=-2", False),_
				Array("<LINE />", "", "", "", False),_ 
				Array("Alta de empleados",_
					  "Alta de empleados de nuevo ingreso.",_
					  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=NewEmployees", False),_
				Array("Cambios a la información de los empleados",_
					  "Actualice la información de los empleados: centros de trabajo, centros de pago, horarios, turnos, servicios, jornadas, información personal, etc.",_
					  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=UpdateEmployeesData", False),_
				Array("Baja de empleados",_
					  "Realice bajas de empleados de manera masiva.",_
					  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesDrop", False),_	
				Array("Licencias a empleados",_
					  "Registre las licencias con goce de sueldo y sin goce de sueldo de los empleados.",_
					  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesLicenses", False),_
				Array("Reanudación de labores",_
					  "Modifique el estatus de los empleados reincorporándolos a sus labores anteriores.",_
					  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesResumptions", False),_
				Array("<LINE />", "", "", "", False),_
				Array("Incidencias",_
					  "Alta de retardos y ausencias.",_
					  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesAbsences", False),_
				Array("<LINE />", "", "", "", False),_
				Array("77. FONAC",_
					  "Registre a los empleados que de forma voluntaria se inscriben al Programa de Fondo de Ahorro Capitalizable.",_
					 "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesFONAC", False),_	 
				Array("<LINE />", "", "", "", False),_
				Array("79. SAR",_
					  "Registre el empleado y la cuota adicional establecida por el como aportación adicional a su cuenta individual del SAR.",_
					 "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesSAR", False),_
				Array("<LINE />", "", "", "", False),_
				Array("Cuentas bancarias",_
					  "Asocie las cuentas bancarias donde los empleados recibirán sus pagos.",_
					  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesAccounts", False),_
				Array("Bloqueo de pagos",_
					  "Especifique el listado de pagos emitidos que se cancelarán.",_
					  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesInactive", False),_
				Array("Revisiones salariales",_
					  "Indiqué a qué empleados se les realizará las revisiones a sus percepciones y deducciones.",_
					  "Images/MnLeftArrows.gif", "Reports.asp?Action=PayrollReviews", False),_
				Array("xxx",_
					  "xxx",_
					  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=" & 0, False)_
			)
			aMenuComponent(B_USE_DIV_MENU) = True
			Response.Write "<TABLE WIDTH=""900"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Call DisplayMenuInThreeSmallColumns(aMenuComponent)
			Response.Write "</TABLE>"

'			Response.Write "<TABLE WIDTH=""720"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
'				aMenuComponent(A_ELEMENTS_MENU) = Array(_
'					Array("Alta masiva",_
'						  "A través de un archivo de texto, suba la información de varios empleados.",_
'						  "Images/MnEmployees.gif", "Employees.asp?Upload=1", ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS) = N_ADD_PERMISSIONS))_
'				)
'				aMenuComponent(B_USE_DIV_MENU) = True
'				Call DisplayMenuInTwoColumns(aMenuComponent)
			Response.Write "</TABLE><BR />"
		End If
		If lErrorNumber <> 0 Then
			Call DisplayErrorMessage("Error al registrar la información", sErrorDescription)
			Response.Write "<BR />"
			lErrorNumber = 0
			sErrorDescription = ""
		End If%>
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>