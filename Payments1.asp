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
<!-- #include file="Libraries/CatalogComponent.asp" -->
<!-- #include file="Libraries/CatalogsLib.asp" -->
<!-- #include file="Libraries/PaymentsLib.asp" -->
<!-- #include file="Libraries/PaymentComponent.asp" -->
<!-- #include file="Libraries/ReportsLib.asp" -->
<!-- #include file="Libraries/ZIPLibrary.asp" -->
<%
Dim sAction
Dim bAction
Dim bError
Dim sError
Dim sCondition
Dim lPayrollID
Dim iStep
Dim bAllPrinted

If B_ISSSTE Then
Else
	If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_PAYMENTS_PERMISSIONS) = N_PAYMENTS_PERMISSIONS Then
	Else
		Response.Redirect "AccessDenied.asp?Permission=" & N_PAYMENTS_PERMISSIONS
	End If
	aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
	aHeaderComponent(S_TITLE_NAME_HEADER) = "Cheques"
End If

Call InitializePaymentComponent(oRequest, aPaymentComponent)
Call GetPaymentsURLValues(oRequest, sAction, bAction, aPaymentComponent(S_QUERY_CONDITION_PAYMENT))
lPayrollID = -1
If Len(oRequest("PayrollID").Item) > 0 Then lPayrollID = CLng(oRequest("PayrollID").Item)
iStep = 1
If Len(oRequest("Step").Item) > 0 Then iStep = CInt(oRequest("Step").Item)

bError = False
If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
	aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
Else
	aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
End If
Select Case sAction
	Case "BlockPayments"
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Bloqueo de depósitos"
		If Len(oRequest("ChangeStatus").Item) > 0 Then
			lErrorNumber = DoPaymentCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, sAction, sCondition, sErrorDescription)
		ElseIf (Len(oRequest("BlockEmployees").Item) > 0) Or (Len(oRequest("RemoveBlockPayments").Item) > 0) Then
			lErrorNumber = DoPaymentCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, sAction, sCondition, sErrorDescription)
		ElseIf Len(oRequest("UnblockEmployees").Item) > 0 Then
			lErrorNumber = DoPaymentCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, sAction, sCondition, sErrorDescription)
		End If
	Case "CancelPayments"
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Cancelación de pagos"
		If Len(oRequest("ChangeStatus").Item) > 0 Then
			lErrorNumber = DoPaymentCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, sAction, sCondition, sErrorDescription)
		End If
	Case "PaymentsMessages", "PrintPayments"
		If Len(oRequest("PosX1").Item) > 0 Then
			If IsNumeric(oRequest("PosX1").Item) Then Call SetOption(aOptionsComponent, CHECKS_LEFT_MARGIN1_OPTION, CLng(oRequest("PosX1").Item), sErrorDescription)
		End If
		If Len(oRequest("PosY1").Item) > 0 Then
			If IsNumeric(oRequest("PosY1").Item) Then Call SetOption(aOptionsComponent, CHECKS_TOP_MARGIN1_OPTION, CLng(oRequest("PosY1").Item), sErrorDescription)
		End If
		If Len(oRequest("PosX2").Item) > 0 Then
			If IsNumeric(oRequest("PosX2").Item) Then Call SetOption(aOptionsComponent, CHECKS_LEFT_MARGIN2_OPTION, CLng(oRequest("PosX2").Item), sErrorDescription)
		End If
		If Len(oRequest("PosY2").Item) > 0 Then
			If IsNumeric(oRequest("PosY2").Item) Then Call SetOption(aOptionsComponent, CHECKS_TOP_MARGIN2_OPTION, CLng(oRequest("PosY2").Item), sErrorDescription)
		End If
		lErrorNumber = ModifyOptions(oRequest, oADODBConnection, aOptionsComponent, sErrorDescription)
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Impresión de pagos"

		If iStep = 2 Then
			aCatalogComponent(S_TABLE_NAME_CATALOG) = "PaymentsMessages"
		Else
			aCatalogComponent(S_TABLE_NAME_CATALOG) = "PaymentsRecords"
		End If
		lErrorNumber = DoPaymentCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, sAction, sCondition, sErrorDescription)
		If Len(oRequest("UpdateStatus").Item) > 0 Then iStep = 3
	Case "PaymentsRecords"
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Asignación de folios"

		aCatalogComponent(S_TABLE_NAME_CATALOG) = "PaymentsRecords"
		lErrorNumber = DoPaymentCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, sAction, sCondition, sErrorDescription)
	Case "Reexpedition"
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Reexpedición"

		aCatalogComponent(S_TABLE_NAME_CATALOG) = "PaymentsRecords2"
		lErrorNumber = DoPaymentCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, sAction, sCondition, sErrorDescription)
	Case "RemovePaymentsRecords"
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Eliminar folios no impresos"

		aCatalogComponent(S_TABLE_NAME_CATALOG) = "PaymentsRecords"
		lErrorNumber = DoPaymentCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, sAction, sCondition, sErrorDescription)
	Case "Replacement"
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Reposiciones"

		aCatalogComponent(S_TABLE_NAME_CATALOG) = "PaymentsRecords"
		lErrorNumber = DoPaymentCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, sAction, sCondition, sErrorDescription)
End Select

If bAction Then
	If B_ISSSTE Then
	Else
		lErrorNumber = DoPaymentAction(oRequest, oADODBConnection, oRequest("Action").Item, sErrorDescription)
	End If
End If
sError = sErrorDescription
bError = (lErrorNumber <> 0)
If aPaymentComponent(N_ID_PAYMENT) > -1 Then
	lErrorNumber = GetPayment(oRequest, oADODBConnection, aPaymentComponent, sErrorDescription)
End If
Response.Cookies("SoS_SectionID") = 206
bWaitMessage = True
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
	</HEAD>
	<SCRIPT LANGUAGE="JavaScript"><!--
		var bReadyToPrint = false;
		var iPrintCounter = 0;
		var iStatusCounter = 0;
	//--></SCRIPT>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%Select Case sAction
			Case "PaymentsRecords", "Reexpedition", "Replacement"
				If (bAction Or (Len(oRequest("DisplayResults").Item) > 0)) And  (Not bError) Then
					aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Exportar a Excel",_
							  "",_
							  "", "javascript: OpenNewWindow('Export.asp?Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
					)
					aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 793
					aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
					aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 200
				End If
			Case "BlockPayments", "CancelPayments"
				If (Len(oRequest("DisplayResults").Item) > 0) Or (Len(oRequest("SearchBlocks").Item) > 0) Then
					aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Exportar a Excel",_
							  "",_
							  "", "javascript: OpenNewWindow('Export.asp?Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
					)
					aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 793
					aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
					aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 200
				End If
			Case "PaymentsMessages", "PrintPayments", "RemovePaymentsRecords"
			Case Else
				If (Len(oRequest.Item("DoSearch")) > 0) Or (aPaymentComponent(N_ID_PAYMENT) > -1) Then
					aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Agregar un nuevo cheque",_
							  "",_
							  "", "Payments.asp?New=1", (aPaymentComponent(N_ID_PAYMENT) = -1)),_
						Array("<LINE />",_
							  "",_
							  "", "", ((Len(oRequest.Item("New")) = 0) And (aPaymentComponent(N_ID_PAYMENT) = -1))),_
						Array("Exportar a Excel",_
							  "",_
							  "", "javascript: OpenNewWindow('Export.asp?Action=Payments&Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (Len(oRequest.Item("DoSearch")) > 0)),_
						Array("Exportar a Word",_
							  "",_
							  "", "javascript: OpenNewWindow('Export.asp?Action=Payments&Word=1&PaymentID=" & oRequest("PaymentID").Item & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (aPaymentComponent(N_ID_PAYMENT) > -1))_
					)
					aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 793
					aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
					aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 200
				End If
		End Select%>
		<!-- #include file="_Header.asp" -->
		<%Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> > "
		If B_ISSSTE Then
			If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
				Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=47"">Cheques y depósitos</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B>"
			Else
				Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=47"">Cheques y depósitos</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B>"
			End If
		Else
			If bAction Or aPaymentComponent(N_ID_PAYMENT) > -1 Then
				Response.Write "<A HREF=""Payments.asp"">Cheques y depósitos</A> > <B>Alta de cheques</B>"
			Else
				Response.Write "<B>Cheques y depósitos</B>"
			End If
		End If
		Response.Write "<BR /><BR />"

		If Len(sErrorDescription) > 0 Then
			Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
			lErrorNumber = 0
			Response.Write "<BR />"
		ElseIf bAction And (Len(sErrorDescription) = 0) Then
			Call DisplayErrorMessage("Confirmación", "La información se actualizó correctamente.")
			Response.Write "<BR />"
		End If
		Select Case sAction
			Case "BlockPayments"
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "function CheckBlockPaymentsFields(oForm) {" & vbNewLine
						Response.Write "if (oForm) {" & vbNewLine
							Response.Write "if (oForm.EmployeeNumber.value == '') {" & vbNewLine
								Response.Write "alert('Favor de especificar el número del empleado');" & vbNewLine
								Response.Write "oForm.EmployeeNumber.focus();" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
							If Len(oRequest("BlockEmployees").Item) > 0 Then
								Response.Write "if ((oForm.PayrollID.value == '') || (oForm.PayrollID.value == '-1')) {" & vbNewLine
									Response.Write "alert('Favor de seleccionar la quincena del bloqueo');" & vbNewLine
									Response.Write "oForm.PayrollID.focus();" & vbNewLine
									Response.Write "return false;" & vbNewLine
								Response.Write "}" & vbNewLine
							Else
								Response.Write "if ((oForm.CancelationPayrollID.value == '') || (oForm.CancelationPayrollID.value == '-1')) {" & vbNewLine
									Response.Write "alert('Favor de seleccionar la quincena de cancelación');" & vbNewLine
									Response.Write "oForm.CancelationPayrollID.focus();" & vbNewLine
									Response.Write "return false;" & vbNewLine
								Response.Write "}" & vbNewLine
							End If
						Response.Write "}" & vbNewLine
						Response.Write "return true;" & vbNewLine
					Response.Write "} // End of CheckBlockPaymentsFields" & vbNewLine

					Response.Write "function CheckMultipleBlockPaymentsFields(oForm) {" & vbNewLine
						Response.Write "if (oForm) {" & vbNewLine
							Response.Write "if (oForm.EmployeeNumbers.value == '') {" & vbNewLine
								Response.Write "alert('Favor de especificar el número de el(los) empleado(s)');" & vbNewLine
								Response.Write "oForm.EmployeeNumbers.focus();" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "return true;" & vbNewLine
					Response.Write "} // End of CheckMultipleBlockPaymentsFields" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine

				Response.Write "<IMG SRC=""Images/Crcl1.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: ShowDisplay(document.all['SearchFormDiv']); if(document.all['PeriodFormDiv'] != null) { HideDisplay(document.all['PeriodFormDiv']); } if(document.all['SearchBlockFormDiv'] != null) { HideDisplay(document.all['SearchBlockFormDiv']); }"">Búsqueda de bloqueos por número de empleado</A><BR /><BR />"
				Response.Write "<DIV NAME=""SearchFormDiv"" ID=""SearchFormDiv"""
					If (Len(oRequest("ShowForm").Item) > 0) And (StrComp(oRequest("ShowForm").Item, "1", vbBinaryCompare) <> 0) Then Response.Write " STYLE=""display: none"""
				Response.Write ">"
					Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""Payments.asp"" METHOD=""GET"" onSubmit=""return CheckBlockPaymentsFields(this)"">" & vbNewLine
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""BlockPayments"" />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ShowForm"" ID=""ShowFormHdn"" VALUE=""1"" />"
						Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número de empleado:&nbsp;</FONT></TD>"
								Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" VALUE=""" & oRequest("EmployeeNumber").Item & """ SIZE=""6"" MAXLENGTH=""6"" CLASS=""TextFields"" /></TD>"
							Response.Write "</TR>"
							Response.Write "<TR NAME=""PayrollDateDiv"" ID=""PayrollDateDiv"">"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Quincena del pago:&nbsp;</FONT></TD>"
								Response.Write "<TD><SELECT NAME=""PayrollID"" ID=""PayrollIDCmb"" SIZE=""1"" CLASS=""Lists"">"
									Response.Write "<OPTION VALUE="""">Todas</OPTION>"
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID<>0) And (IsClosed=1)", "PayrollID Desc", lPayrollID, "Ninguna;;;-1", sErrorDescription)
								Response.Write "</SELECT></TD>"
							Response.Write "</TR>"
						Response.Write "</TABLE><BR />"
						Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""DisplayResults"" ID=""DisplayResultsBtn"" VALUE=""Realizar Búsqueda"" CLASS=""Buttons"" />"
						Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
						Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=47'"" />"
					Response.Write "</FORM>"

					If Len(oRequest("DisplayResults").Item) > 0 Then
						Response.Write "<IMG SRC=""Images/DotBlue.gif"" WIDTH=""960"" HEIGHT=""1"" /><BR />"
						Response.Write "<FORM NAME=""ModifyFrm"" ID=""ModifyFrm"" ACTION=""Payments.asp"" METHOD=""POST"">" & vbNewLine
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeNumber"" ID=""EmployeeNumberHdn"" VALUE=""" & oRequest("EmployeeNumber").Item & """ />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PayrollID"" ID=""PayrollIDHdn"" VALUE=""" & lPayrollID & """ />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""BlockPayments"" />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ShowForm"" ID=""ShowFormHdn"" VALUE=""1"" />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DisplayResults"" ID=""DisplayResultsHdn"" VALUE=""1"" />"
							lErrorNumber = DisplayEmployeePaymentsTable(oRequest, oADODBConnection, 1, False, sErrorDescription)
							If lErrorNumber = 0 Then
								Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Quincena de cancelación:&nbsp;</B></FONT>"
								Response.Write "<SELECT NAME=""CancelationPayrollID"" ID=""CancelationPayrollIDCmb"" SIZE=""1"" CLASS=""Lists"">"
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID=0) And (IsClosed<>1)", "PayrollID Desc", lPayrollID, "No existen quincenas de cancelación;;;-1", sErrorDescription)
								Response.Write "</SELECT><BR /><BR />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Comentarios:<BR /></B></FONT>"
								Response.Write "<TEXTAREA NAME=""Description"" ID=""DescriptionTxtArea"" ROWS=""4"" COLS=""60"" CLASS=""TextFields""></TEXTAREA><BR /><BR />"

								Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ChangeStatus"" ID=""ChangeStatusBtn"" VALUE=""Modificar Depósitos"" CLASS=""Buttons"" />"
								Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
								Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=47'"" />"
								Response.Write "<BR /><BR /><BR />"
							End If
						Response.Write "</FORM>" & vbNewLine
					End If
				Response.Write "</DIV><BR />" & vbNewLine

				Response.Write "<IMG SRC=""Images/Crcl2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript:  if(document.all['SearchFormDiv'] != null) { HideDisplay(document.all['SearchFormDiv']); } ShowDisplay(document.all['PeriodFormDiv']);  if(document.all['SearchBlockFormDiv'] != null) { HideDisplay(document.all['SearchBlockFormDiv']); }"">Bloqueo anticipado de los depósitos por fecha de pago</A><BR /><BR />"
				Response.Write "<DIV NAME=""PeriodFormDiv"" ID=""PeriodFormDiv"""
					If StrComp(oRequest("ShowForm").Item, "2", vbBinaryCompare) <> 0 Then Response.Write " STYLE=""display: none"""
				Response.Write "><FORM NAME=""PeriodFrm"" ID=""PeriodFrm"" ACTION=""Payments.asp"" METHOD=""GET"" onSubmit=""return CheckMultipleBlockPaymentsFields(this)"">" & vbNewLine
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""BlockPayments"" />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ShowForm"" ID=""ShowFormHdn"" VALUE=""2"" />"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Empleados a los que se les bloquerán sus depósitos:</B><BR /></FONT>"
					Response.Write "<TEXTAREA NAME=""EmployeeNumbers"" ID=""EmployeeNumbers"" ROWS=""6"" COLS=""60"" CLASS=""TextFields""></TEXTAREA><BR /><BR />"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Bloquear la quincena </FONT>"
					Response.Write "<SELECT NAME=""PayrollID"" ID=""PayrollID"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID<>0) And (IsActive_4=1)", "PayrollID Desc", "", "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
					Response.Write "</SELECT>&nbsp;"
					Response.Write "<BR /><BR />"
					Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""BlockEmployees"" ID=""BlockEmployeesBtn"" VALUE=""Bloquear Depósitos"" CLASS=""Buttons"" />"
					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
					Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=47'"" />"
					Response.Write "<BR /><BR /><BR />"
				Response.Write "</FORM></DIV><BR />" & vbNewLine

				Response.Write "<IMG SRC=""Images/Crcl3.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['SearchFormDiv'] != null) { HideDisplay(document.all['SearchFormDiv']);} if(document.all['PeriodFormDiv'] != null) { HideDisplay(document.all['PeriodFormDiv']);} ShowDisplay(document.all['SearchBlockFormDiv']);"">Consulta de empleados con bloqueo anticipado de depósitos</A><BR /><BR />"
				Response.Write "<DIV NAME=""SearchBlockFormDiv"" ID=""SearchBlockFormDiv"" STYLE=""display: none"">"
					Response.Write "<FORM NAME=""SearchBlockFrm"" ID=""SearchBlockFrm"" METHOD=""POST"">"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & sAction & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartPage"" ID=""StartPageHdn"" VALUE=""1"" />"
						Response.Write "<B>Seleccione los datos para filtrar los registros:&nbsp;&nbsp;&nbsp;</B><BR />"
						Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" />Mostrar bloqueos de la nómina:&nbsp;"
						Response.Write "<SELECT NAME=""PayrollID"" ID=""PayrollID"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID<>0) And (IsActive_4=1)", "PayrollID Desc", CLng(oRequest("PayrollID").Item), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
						Response.Write "</SELECT>&nbsp;&nbsp;"
						Response.Write "<INPUT TYPE=""SUBMIT"" VALUE=""Consultar registros"" CLASS=""Buttons""><BR /><BR />"
						lErrorNumber = DisplayEmployeesBlockPaymentsTable(oRequest, oADODBConnection, False, aPaymentComponent, sErrorDescription)
						If lErrorNumber <> 0 Then
							Response.Write "<BR />"
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							lErrorNumber = 0
							sErrorDescription = ""
						End If
					Response.Write "</FORM>"
				Response.Write "</DIV>"

				If False Then
					Response.Write "<IMG SRC=""Images/Crcl2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: HideDisplay(document.all['SearchFormDiv']); ShowDisplay(document.all['PeriodFormDiv']); HideDisplay(document.all['SearchBlockFormDiv']);"">Bloqueo de los depósitos por periodo</A><BR /><BR />"
					Response.Write "<DIV NAME=""PeriodFormDiv"" ID=""PeriodFormDiv"""
						If StrComp(oRequest("ShowForm").Item, "2", vbBinaryCompare) <> 0 Then Response.Write " STYLE=""display: none"""
					Response.Write "><FORM NAME=""PeriodFrm"" ID=""PeriodFrm"" ACTION=""Payments.asp"" METHOD=""GET"" onSubmit=""return CheckMultipleBlockPaymentsFields(this)"">" & vbNewLine
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""BlockPayments"" />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ShowForm"" ID=""ShowFormHdn"" VALUE=""2"" />"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Empleados a los que se les bloquerán sus depósitos:</B><BR /></FONT>"
						Response.Write "<TEXTAREA NAME=""EmployeeNumbers"" ID=""EmployeeNumbers"" ROWS=""6"" COLS=""60"" CLASS=""TextFields""></TEXTAREA><BR /><BR />"

						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Bloquear desde el </FONT>"
						Response.Write DisplayDateCombosUsingSerial("", "StartPayment", N_FORM_START_YEAR, Year(Date()), True, False)
						Response.Write "<FONT FACE=""Arial"" SIZE=""2""> hasta el </FONT>"
						Response.Write DisplayDateCombosUsingSerial("", "EndPayment", N_FORM_START_YEAR, Year(Date()), True, False)
						Response.Write "<FONT FACE=""Arial"" SIZE=""2""><BR /><BR /></FONT>"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Comentarios:<BR /></B></FONT>"
						Response.Write "<TEXTAREA NAME=""Description"" ID=""DescriptionTxtArea"" ROWS=""4"" COLS=""60"" CLASS=""TextFields""></TEXTAREA><BR /><BR />"

						Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""BlockEmployees"" ID=""BlockEmployeesBtn"" VALUE=""Bloquear Depósitos"" CLASS=""Buttons"" />"
						Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
						Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=47'"" />"
						Response.Write "<BR /><BR /><BR />"
					Response.Write "</FORM></DIV><BR />" & vbNewLine

					Response.Write "<IMG SRC=""Images/Crcl3.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: HideDisplay(document.all['SearchFormDiv']); HideDisplay(document.all['PeriodFormDiv']); ShowDisplay(document.all['SearchBlockFormDiv']);"">Desbloqueo de depósitos por periodo</A><BR /><BR />"
					Response.Write "<DIV NAME=""SearchBlockFormDiv"" ID=""SearchBlockFormDiv"""
						If StrComp(oRequest("ShowForm").Item, "3", vbBinaryCompare) <> 0 Then Response.Write " STYLE=""display: none"""
					Response.Write "><FORM NAME=""SearchBlockFrm"" ID=""SearchBlockFrm"" ACTION=""Payments.asp"" METHOD=""GET"">" & vbNewLine
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""BlockPayments"" />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ShowForm"" ID=""ShowFormHdn"" VALUE=""3"" />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PaymentStatusID"" ID=""PaymentStatusIDHdn"" VALUE=""4"" />"

						Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Empleados a los que se les desbloquerán sus depósitos:<-B><BR /></FONT>"
						Response.Write "<TEXTAREA NAME=""EmployeeNumbers"" ID=""EmployeeNumbers"" ROWS=""6"" COLS=""60"" CLASS=""TextFields""></TEXTAREA><BR /><BR />"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Comentarios:<BR /></B></FONT>"
						Response.Write "<TEXTAREA NAME=""Description"" ID=""DescriptionTxtArea"" ROWS=""4"" COLS=""60"" CLASS=""TextFields""></TEXTAREA><BR /><BR />"

						Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""SearchBlocks"" ID=""SearchBlocksBtn"" VALUE=""Buscar Bloqueos"" CLASS=""Buttons"" />"
						Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
						Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=47'"" />"
					Response.Write "</FORM></DIV>" & vbNewLine

					If Len(oRequest("SearchBlocks").Item) > 0 Then
						Response.Write "<IMG SRC=""Images/DotBlue.gif"" WIDTH=""960"" HEIGHT=""1"" /><BR />"
						Response.Write "<FORM NAME=""ModifyFrm"" ID=""ModifyFrm"" ACTION=""Payments.asp"" METHOD=""POST"">" & vbNewLine
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeNumbers"" ID=""EmployeeNumbersHdn"" VALUE=""" & oRequest("EmployeeNumbers").Item & """ />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""BlockPayments"" />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ShowForm"" ID=""ShowFormHdn"" VALUE=""3"" />"
							lErrorNumber = DisplayEmployeePaymentsTable(oRequest, oADODBConnection, 1, False, sErrorDescription)
							If lErrorNumber = 0 Then
								Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Quincena de cancelación:&nbsp;</B></FONT>"
								Response.Write "<SELECT NAME=""CancelationPayrollID"" ID=""CancelationPayrollIDCmb"" SIZE=""1"" CLASS=""Lists"">"
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID=0) And (IsClosed<>1)", "PayrollID Desc", lPayrollID, "Ninguna;;;-1", sErrorDescription)
								Response.Write "</SELECT><BR /><BR />"

								Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ChangeStatus"" ID=""ChangeStatusBtn"" VALUE=""Bloquear Depósitos"" CLASS=""Buttons"" />"
								Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
								Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Regresar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=47'"" />"
							End If
						Response.Write "</FORM>" & vbNewLine
					End If
				End If
			Case "CancelPayments"
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "function CheckCancelPaymentsFields(oForm) {" & vbNewLine
						Response.Write "if (oForm) {" & vbNewLine
							Response.Write "if (oForm.EmployeeNumber.value == '') {" & vbNewLine
								Response.Write "alert('Favor de especificar el número del empleado');" & vbNewLine
								Response.Write "oForm.EmployeeNumber.focus();" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
							Response.Write "if ((oForm.CancelationPayrollID.value == '') || (oForm.CancelationPayrollID.value == '-1')) {" & vbNewLine
								Response.Write "alert('Favor de seleccionar la quincena de cancelación');" & vbNewLine
								Response.Write "oForm.CancelationPayrollID.focus();" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "return true;" & vbNewLine
					Response.Write "} // End of CheckCancelPaymentsFields" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine

				Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""Payments.asp"" METHOD=""GET"" onSubmit=""return CheckCancelPaymentsFields(this)"">" & vbNewLine
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""CancelPayments"" />"
					Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número de empleado:&nbsp;</FONT></TD>"
							Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" VALUE=""" & oRequest("EmployeeNumber").Item & """ SIZE=""6"" MAXLENGTH=""6"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR NAME=""PayrollDateDiv"" ID=""PayrollDateDiv"">"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Quincena del pago:&nbsp;</FONT></TD>"
							Response.Write "<TD><SELECT NAME=""PayrollID"" ID=""PayrollIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write "<OPTION VALUE="""">Todas</OPTION>"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID<>0) And (IsClosed=1)", "PayrollID Desc", lPayrollID, "Ninguna;;;-1", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de pago:&nbsp;</FONT></TD>"
							Response.Write "<TD><SELECT NAME=""PaymentType"" ID=""PaymentTypeCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write "<OPTION VALUE="""">Todos</OPTION>"
								Response.Write "<OPTION VALUE=""0"">Cheques</OPTION>"
								Response.Write "<OPTION VALUE=""1"">Depósitos</OPTION>"
							Response.Write "</SELECT></TD>"
						Response.Write "</TR>"
					Response.Write "</TABLE><BR />"
					Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""DisplayResults"" ID=""DisplayResultsBtn"" VALUE=""Realizar Búsqueda"" CLASS=""Buttons"" />"
					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
					Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=47'"" />"
				Response.Write "</FORM>" & vbNewLine
				If Len(oRequest("DisplayResults").Item) > 0 Then
					Response.Write "<IMG SRC=""Images/DotBlue.gif"" WIDTH=""960"" HEIGHT=""1"" /><BR />"
					Response.Write "<FORM NAME=""ModifyFrm"" ID=""ModifyFrm"" ACTION=""Payments.asp"" METHOD=""POST"" onSubmit=""return CheckCancelPaymentsFields(this)"">" & vbNewLine
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeNumber"" ID=""EmployeeNumberHdn"" VALUE=""" & oRequest("EmployeeNumber").Item & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PayrollID"" ID=""PayrollIDHdn"" VALUE=""" & lPayrollID & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""CancelPayments"" />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DisplayResults"" ID=""DisplayResultsHdn"" VALUE=""1"" />"
						lErrorNumber = DisplayEmployeePaymentsTable(oRequest, oADODBConnection, 0, False, sErrorDescription)
						If lErrorNumber = 0 Then
							Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Quincena de cancelación:&nbsp;</B></FONT>"
							Response.Write "<SELECT NAME=""CancelationPayrollID"" ID=""CancelationPayrollIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID=0) And (IsClosed<>1)", "PayrollID Desc", lPayrollID, "No existen quincenas de cancelación;;;-1", sErrorDescription)
							Response.Write "</SELECT><BR /><BR />"

							Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Comentarios:<BR /></B></FONT>"
							Response.Write "<TEXTAREA NAME=""Description"" ID=""DescriptionTxtArea"" ROWS=""4"" COLS=""60"" CLASS=""TextFields""></TEXTAREA><BR /><BR />"

							Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ChangeStatus"" ID=""ChangeStatusBtn"" VALUE=""Actualizar Estatus"" CLASS=""Buttons"" />"
							Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
						End If
						Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Regresar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=47'"" />"
					Response.Write "</FORM>" & vbNewLine
				End If
			Case "PaymentsMessages", "PrintPayments"
				Select Case iStep
					Case 1
						Response.Write "<FORM NAME=""PrintFrm"" ID=""PrintFrm"" ACTION=""Payments.asp"" METHOD=""GET"">"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""PrintPayments"" />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""" & iStep + 1 & """ />"
							Call DisplayInstructionsMessage("Instrucciones", "Seleccione la quincena para imprimir los cheques")
							Response.Write "<BR /><BR />"
							Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Quincena del pago:&nbsp;</FONT>"
							Response.Write "<SELECT NAME=""PayrollID"" ID=""PayrollIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID<>0) And (IsClosed=1)", "PayrollID Desc", lPayrollID, "Ninguna;;;-1", sErrorDescription)
							Response.Write "</SELECT><BR /><BR />"

							Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Continue"" ID=""ContinueBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
							Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
							Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=47';"" />"
						Response.Write "</FORM>" & vbNewLine
					Case 2
						Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
							Response.Write "<TD VALIGN=""TOP"">"
								Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchPositionsIFrame"" FRAMEBORDER=""1"" WIDTH=""575"" HEIGHT=""0""></IFRAME>"
								lErrorNumber = DisplayCatalogForm(oRequest, oADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)
							Response.Write "</TD>"
							Response.Write "<TD>&nbsp;</TD>"
							Response.Write "<TD BGCOLOR=""" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""1"" ><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
							Response.Write "<TD>&nbsp;</TD>"
							Response.Write "<TD VALIGN=""TOP"">"
								lErrorNumber = DisplayPaymentsMessages(oRequest, oADODBConnection, lPayrollID, False, sErrorDescription)
							Response.Write "</TD>"
						Response.Write "</TABLE>"
					Case 3
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							Response.Write "function CheckPrintPaymentsFields(oForm) {" & vbNewLine
								Response.Write "if (oForm) {" & vbNewLine
									Response.Write "if (!bReadyToPrint) {" & vbNewLine
										Response.Write "alert('Favor de seleccionar un registro');" & vbNewLine
										Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									Response.Write "if (oForm.PosX1.value > 10 || oForm.PosX1.value < -50) {" & vbNewLine
										Response.Write "alert('El desplazamiento izquierdo de la sección superior del formato de impresion del cheque debe estar entre -50 y 10');" & vbNewLine
										Response.Write "oForm.PosX1.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.PosY1.value > 20 || oForm.PosY1.value < -50) {" & vbNewLine
										Response.Write "alert('El desplazamiento superior de la sección superior del formato de impresion del cheque debe estar entre -50 y 20');" & vbNewLine
										Response.Write "oForm.PosY1.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.PosX2.value > 10 || oForm.PosX2.value < -50) {" & vbNewLine
										Response.Write "alert('El desplazamiento izquierdo de la sección desprendible del formato de impresion del cheque debe estar entre -50 y 10');" & vbNewLine
										Response.Write "oForm.PosX2.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.PosY2.value > 20 || oForm.PosY2.value < -50) {" & vbNewLine
										Response.Write "alert('El desplazamiento superior de la sección desprendible del formato de impresion del cheque debe estar entre -50 y 20');" & vbNewLine
										Response.Write "oForm.PosY2.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "}" & vbNewLine
								Response.Write "return true;" & vbNewLine
							Response.Write "} // End of CheckCancelPaymentsFields" & vbNewLine
						Response.Write "//--></SCRIPT>" & vbNewLine
						Response.Write "<FORM NAME=""PrintFrm"" ID=""PrintFrm"" ACTION=""Payments.asp"" METHOD=""GET"" onSubmit=""return CheckPrintPaymentsFields(this)"">"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""PrintPayments"" />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PayrollID"" ID=""PayrollIDHdn"" VALUE=""" & lPayrollID & """ />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""" & iStep + 1 & """ />"
							lErrorNumber = DisplayPaymentRecordsTable(oRequest, oADODBConnection, sAction, lPayrollID, False, bAllPrinted, sErrorDescription)
							If lErrorNumber = 0 Then
								If Not bAllPrinted Then
									lErrorNumber = DisplayPaymentsMarginsSettings(oRequest, oADODBConnection, sErrorDescription)
									Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Continue"" ID=""ContinueBtn"" VALUE=""Imprimir"" CLASS=""Buttons"" onClick=""bReadyToPrint = (iPrintCounter > 0);"" />"
									Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
								End If
							Else
								Response.Write "<BR />"
								Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
								lErrorNumber = 0
								Response.Write "<BR /><BR />"
							End If
							Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""UpdateStatus"" ID=""UpdateStatusBtn"" VALUE=""Cambiar Estatus"" CLASS=""Buttons"" onClick=""bReadyToPrint = (iStatusCounter > 0);"" />"
							Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
							Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Return"" ID=""ReturnBtn"" VALUE=""Regresar"" CLASS=""Buttons"" onClick=""window.location.href='Payments.asp?Action=PrintPayments&PayrollID=" & lPayrollID & "&Step=2';"" />"
						Response.Write "</FORM>" & vbNewLine
					Case 4
						Response.Write "<FORM NAME=""PrintFrm"" ID=""PrintFrm"" ACTION=""Payments.asp"" METHOD=""GET"">"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""PrintPayments"" />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PayrollID"" ID=""PayrollIDHdn"" VALUE=""" & lPayrollID & """ />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"

							lErrorNumber = PrintPayments(oRequest, oADODBConnection, oRequest("RecordID").Item, sErrorDescription)
							Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Return"" ID=""ReturnBtn"" VALUE=""Regresar"" CLASS=""Buttons"" onClick=""window.location.href='Payments.asp?Action=PrintPayments&PayrollID=" & lPayrollID & "&Step=3';"" />"
						Response.Write "</FORM>" & vbNewLine
				End Select
			Case "PaymentsRecords"
				If (bAction And (Len(oRequest("Remove").Item) > 0)) Or bError Then
					Response.Write "<FORM><INPUT TYPE=""BUTTON"" NAME=""Continue"" ID=""ContinueBtn"" VALUE=""Continuar"" CLASS=""Buttons"" onClick=""window.location.href='Payments.asp?Action=PaymentsRecords';"" /></FORM>" & vbNewLine
				ElseIf bAction Or (Len(oRequest("DisplayResults").Item) > 0) Then
					lErrorNumber = DisplayNewPaymentsTable(oRequest, oADODBConnection, False, aCatalogComponent, sErrorDescription)
					Response.Write "<FORM NAME=""ConfirmationFrm"" ID=""ConfirmationFrm"" ACTION=""Payments.asp"" METHOD=""GET"">"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""PaymentsRecords"" />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RecordID"" ID=""RecordIDHdn"" VALUE=""" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) & """ />"
						Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Continue"" ID=""ContinueBtn"" VALUE=""Continuar"" CLASS=""Buttons"" onClick=""window.location.href='Payments.asp?Action=PaymentsRecords';"" />"
						Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
						Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveRecordWngDiv']); ConfirmationFrm.Remove.focus()"" />"

						Response.Write "<BR /><BR />"
						Call DisplayWarningDiv("RemoveRecordWngDiv", "¿Está seguro que desea borrar las asignaciones de folios que se muestran arriba?")
					Response.Write "</FORM>" & vbNewLine
				Else
					Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchAccountsCatalogsIFrame"" FRAMEBORDER=""0"" WIDTH=""0"" HEIGHT=""0""></IFRAME><BR />"
					Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
						Response.Write "function ShowPaymentsRecordsFields(sBankID) {" & vbNewLine
							Response.Write "oForm = document.CatalogFrm;" & vbNewLine

							Response.Write "if (oForm) {" & vbNewLine
'								Response.Write "SearchRecord(sBankID, 'BankAccounts', 'SearchAccountsCatalogsIFrame', 'CatalogFrm.AccountID');" & vbNewLine
							Response.Write "}" & vbNewLine
						Response.Write "} // End of ShowPaymentsRecordsFields" & vbNewLine
					Response.Write "//--></SCRIPT>" & vbNewLine

					lErrorNumber = DisplayCatalogForm(oRequest, oADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)
					Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
						Response.Write "ShowPaymentsRecordsFields(document.CatalogFrm.BankID.value);" & vbNewLine
					Response.Write "//--></SCRIPT>" & vbNewLine
				End If
			Case "Reexpedition"
				If (bAction And (Len(oRequest("Remove").Item) > 0)) Or bError Then
					Response.Write "<FORM><INPUT TYPE=""BUTTON"" NAME=""Continue"" ID=""ContinueBtn"" VALUE=""Continuar"" CLASS=""Buttons"" onClick=""window.location.href='Payments.asp?Action=Reexpedition';"" /></FORM>" & vbNewLine
				ElseIf bAction Or (Len(oRequest("DisplayResults").Item) > 0) Then
					lErrorNumber = DisplayNewPaymentsTable(oRequest, oADODBConnection, False, aCatalogComponent, sErrorDescription)
					Response.Write "<FORM NAME=""ConfirmationFrm"" ID=""ConfirmationFrm"" ACTION=""Payments.asp"" METHOD=""GET"">"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""Reexpedition"" />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RecordID"" ID=""RecordIDHdn"" VALUE=""" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) & """ />"
						Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Continue"" ID=""ContinueBtn"" VALUE=""Continuar"" CLASS=""Buttons"" onClick=""window.location.href='Payments.asp?Action=Reexpedition';"" />"
						Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
						Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveRecordWngDiv']); ConfirmationFrm.Remove.focus()"" />"

						Response.Write "<BR /><BR />"
						Call DisplayWarningDiv("RemoveRecordWngDiv", "¿Está seguro que desea borrar las asignaciones de folios que se muestran arriba?")
					Response.Write "</FORM>" & vbNewLine
				Else
					Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
						Response.Write "function CheckEmployeeValidation() {" & vbNewLine
							Response.Write "var oForm = document.CatalogFrm;" & vbNewLine

							Response.Write "if (oForm) {" & vbNewLine
								Response.Write "if (oForm.CheckNumber.value == '') {" & vbNewLine
									Response.Write "alert('Favor de validar la existencia del empleado y del pago');" & vbNewLine
									Response.Write "oForm.EmployeeID.focus();" & vbNewLine
									Response.Write "return false;" & vbNewLine
								Response.Write "}" & vbNewLine
							Response.Write "}" & vbNewLine
							Response.Write "return true;" & vbNewLine
						Response.Write "} // End of CheckEmployeeValidation" & vbNewLine
						
						Response.Write "function SearchForPayment() {" & vbNewLine
							Response.Write "var bCorrect = true;" & vbNewLine
							Response.Write "if (bCorrect && (document.CatalogFrm.EmployeeID.value == '')) {" & vbNewLine
								Response.Write "alert('Favor de especificar el número del empleado.');" & vbNewLine
								Response.Write "bCorrect = false;" & vbNewLine
								Response.Write "document.CatalogFrm.EmployeeID.focus();" & vbNewLine
							Response.Write "}" & vbNewLine

							Response.Write "if (bCorrect)" & vbNewLine
								Response.Write "SearchRecord(document.CatalogFrm.EmployeeID.value, 'EmployeePayment&PaymentDate=' + document.CatalogFrm.PayrollID.value, 'SearchAccountsCatalogsIFrame', 'CatalogFrm.CheckNumber');" & vbNewLine
						Response.Write "} // End of SearchForPayment" & vbNewLine

						Response.Write "function ShowPaymentsRecordsFields(sBankID) {" & vbNewLine
							Response.Write "oForm = document.CatalogFrm;" & vbNewLine

							Response.Write "if (oForm) {" & vbNewLine
								Response.Write "SearchRecord(sBankID, 'BankAccounts', 'SearchAccountsCatalogsIFrame', 'CatalogFrm.AccountID');" & vbNewLine
							Response.Write "}" & vbNewLine
						Response.Write "} // End of ShowPaymentsRecordsFields" & vbNewLine
					Response.Write "//--></SCRIPT>" & vbNewLine

					'Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchAccountsCatalogsIFrame"" FRAMEBORDER=""0"" WIDTH=""0"" HEIGHT=""0""></IFRAME><BR />"

					aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(9) = Left(GetSerialNumberForDate(""), Len("00000000"))
					lErrorNumber = DisplayCatalogForm(oRequest, oADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)
					Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
						Response.Write "document.CatalogFrm.Action.value = 'Reexpedition';" & vbNewLine
						Response.Write "ShowPaymentsRecordsFields(document.CatalogFrm.BankID.value);" & vbNewLine
					Response.Write "//--></SCRIPT>" & vbNewLine
				End If
			Case "RemovePaymentsRecords"
				Select Case iStep
					Case 1
						Response.Write "<FORM NAME=""PrintFrm"" ID=""PrintFrm"" ACTION=""Payments.asp"" METHOD=""GET"">"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""RemovePaymentsRecords"" />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""" & iStep + 1 & """ />"
							Call DisplayInstructionsMessage("Instrucciones", "Seleccione la quincena para imprimir los cheques")
							Response.Write "<BR /><BR />"
							Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Quincena del pago:&nbsp;</FONT>"
							Response.Write "<SELECT NAME=""PayrollID"" ID=""PayrollIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID<>0) And (IsClosed=1)", "PayrollID Desc", lPayrollID, "Ninguna;;;-1", sErrorDescription)
							Response.Write "</SELECT><BR /><BR />"

							Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Continue"" ID=""ContinueBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
							Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
							Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=47';"" />"
						Response.Write "</FORM>" & vbNewLine
					Case 2
						If (bAction And (Len(oRequest("Remove").Item) > 0)) Or bError Then
							Response.Write "<FORM><INPUT TYPE=""BUTTON"" NAME=""Continue"" ID=""ContinueBtn"" VALUE=""Continuar"" CLASS=""Buttons"" onClick=""window.location.href='Payments.asp?Action=RemovePaymentsRecords';"" /></FORM>" & vbNewLine
						ElseIf Len(oRequest("DisplayResults").Item) > 0 Then
							lErrorNumber = DisplayNewPaymentsTable(oRequest, oADODBConnection, False, aCatalogComponent, sErrorDescription)
							Response.Write "<FORM NAME=""ConfirmationFrm"" ID=""ConfirmationFrm"" ACTION=""Payments.asp"" METHOD=""GET"">"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""RemovePaymentsRecords"" />"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""" & iStep & """ />"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RecordID"
									If Len(oRequest("RecordID2").Item) > 0 Then Response.Write "2"
								Response.Write """ ID=""RecordID"
									If Len(oRequest("RecordID2").Item) > 0 Then Response.Write "2"
								Response.Write "Hdn"" VALUE=""" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) & """ />"
								Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveRecordWngDiv']); ConfirmationFrm.Remove.focus()"" />"
								Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
								Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Return"" ID=""ReturnBtn"" VALUE=""Regresar"" CLASS=""Buttons"" onClick=""window.location.href='Payments.asp?Action=RemovePaymentsRecords';"" />"

								Response.Write "<BR /><BR />"
								Call DisplayWarningDiv("RemoveRecordWngDiv", "¿Está seguro que desea borrar las asignaciones de folios que se muestran arriba?")
							Response.Write "</FORM>" & vbNewLine
						Else
							Response.Write "<FORM NAME=""PaymentsFrm"" ID=""PaymentsFrm"" ACTION=""Payments.asp"" METHOD=""GET"" onSubmit=""if (bReadyToPrint) {return true;} else {alert('Favor de seleccionar un registro'); return false;}"">"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""RemovePaymentsRecords"" />"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""" & iStep & """ />"
								lErrorNumber = DisplayPaymentRecordsTable(oRequest, oADODBConnection, sAction, lPayrollID, False, bAllPrinted, sErrorDescription)
								If lErrorNumber = 0 Then
									Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""DisplayResults"" ID=""DisplayResultsBtn"" VALUE=""Continuar"" CLASS=""Buttons"" onClick=""bReadyToPrint = (iPrintCounter > 0);"" />"
									Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
								Else
									Response.Write "<BR />"
									Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
									lErrorNumber = 0
									Response.Write "<BR /><BR />"
								End If
								Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='Payments.asp?Action=RemovePaymentsRecords&PayrollID=" & lPayrollID & "&Step=1';"" />"

								Response.Write "<BR /><BR />"
								Call DisplayWarningDiv("RemoveRecordWngDiv", "¿Está seguro que desea borrar las asignaciones de folios que se muestran arriba?")
							Response.Write "</FORM>" & vbNewLine
						End If
					End Select
			Case "Replacement"
				If (bAction And (Len(oRequest("Remove").Item) > 0)) Or bError Then
					Response.Write "<FORM><INPUT TYPE=""BUTTON"" NAME=""Continue"" ID=""ContinueBtn"" VALUE=""Continuar"" CLASS=""Buttons"" onClick=""window.location.href='Payments.asp?Action=Replacement';"" /></FORM>" & vbNewLine
				ElseIf bAction Or (Len(oRequest("DisplayResults").Item) > 0) Then
					lErrorNumber = DisplayNewPaymentsTable(oRequest, oADODBConnection, False, aCatalogComponent, sErrorDescription)
					Response.Write "<FORM NAME=""ConfirmationFrm"" ID=""ConfirmationFrm"" ACTION=""Payments.asp"" METHOD=""GET"">"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""Replacement"" />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RecordID"" ID=""RecordIDHdn"" VALUE=""" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) & """ />"
						Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Continue"" ID=""ContinueBtn"" VALUE=""Continuar"" CLASS=""Buttons"" onClick=""window.location.href='Payments.asp?Action=Replacement';"" />"
						Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
						Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveRecordWngDiv']); ConfirmationFrm.Remove.focus()"" />"

						Response.Write "<BR /><BR />"
						Call DisplayWarningDiv("RemoveRecordWngDiv", "¿Está seguro que desea borrar las asignaciones de folios que se muestran arriba?")
					Response.Write "</FORM>" & vbNewLine
				Else
					aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,6,8,8,8,8,6,6,4,11,4,4,11,11,4,11,11,11,11"
					aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
					aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,-1,,,,,-1,-1,1,1,1,1,-1,-1,0,0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",-1"
					aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
					aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,-1,,,,,-1,-1,1,1,1,1,-1,-1,0,0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",-1"
					aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
					Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchAccountsCatalogsIFrame"" FRAMEBORDER=""0"" WIDTH=""0"" HEIGHT=""0""></IFRAME><BR />"
					Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
						Response.Write "function ShowPaymentsRecordsFields(sBankID) {" & vbNewLine
							Response.Write "oForm = document.CatalogFrm;" & vbNewLine

							Response.Write "if (oForm) {" & vbNewLine
								Response.Write "SearchRecord(sBankID, 'BankAccounts', 'SearchAccountsCatalogsIFrame', 'CatalogFrm.AccountID');" & vbNewLine
							Response.Write "}" & vbNewLine
						Response.Write "} // End of ShowPaymentsRecordsFields" & vbNewLine
					Response.Write "//--></SCRIPT>" & vbNewLine
					lErrorNumber = DisplayCatalogForm(oRequest, oADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)
					Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
						Response.Write "ShowPaymentsRecordsFields(document.CatalogFrm.BankID.value);" & vbNewLine
						Response.Write "document.CatalogFrm.Action.value = 'Replacement';" & vbNewLine
					Response.Write "//--></SCRIPT>" & vbNewLine
				End If
		End Select
		Select Case sAction
			Case "BlockPayments"
				If (Len(oRequest("StartPage").Item) > 0) Then
					Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
						Response.Write "if(document.all['SearchFormDiv'] != null) { HideDisplay(document.all['SearchFormDiv']);} if(document.all['PeriodFormDiv'] != null) { HideDisplay(document.all['PeriodFormDiv']);} ShowDisplay(document.all['SearchBlockFormDiv']);" & vbNewLine
					Response.Write "//--></SCRIPT>" & vbNewLine
				End If
		End Select
		If False Then
			lErrorNumber = ShowSignatures(oRequest, oADODBConnection, sErrorDescription)
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