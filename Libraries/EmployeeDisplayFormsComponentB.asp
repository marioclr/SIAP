<%
Function DisplayAnotherEmployeeForm(oRequest, oADODBConnection, sASPFileName, sAction, iLeftWidth, lReasonID, sAltDescription, sDescription, sErrorDescription)
'************************************************************
'Purpose: To display the information about a registration of child for the
'         employee from the database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sASPFileName, sAction, sURL, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayAnotherEmployeeForm"
	Dim sNames
	Dim lErrorNumber

	Response.Write "<FORM NAME=""AnotherConceptFrm"" ID=""AnotherConceptFrm"" ACTION=""" & sASPFileName & """ METHOD=""GET"">"
		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""" & iLeftWidth & """>&nbsp;</TD>"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""32""><IMG SRC=""Images/MnLeftArrows.gif"" WIDTH=""32"" HEIGHT=""32"" ALT=""" & sAltDescription & """ BORDER=""0"" /><BR /></TD>"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""350""><FONT FACE=""Arial"" SIZE=""2""><B>Otro empleado</B><BR /></FONT>"
				Response.Write "<DIV CLASS=""MenuOverflow""><FONT FACE=""Arial"" SIZE=""2"">" & sDescription & "</FONT></DIV></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""" & iLeftWidth & """>&nbsp;</TD>"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""32""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;&nbsp;&nbsp;</FONT></TD>"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""350""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;&nbsp;&nbsp;Número del empleado:&nbsp;</FONT><INPUT TYPE=""TEXT"" NAME=""EmployeeID"" ID=""EmployeeIDTxt"" SIZE=""6"" MAXLENGTH=""6"" CLASS=""TextFields"" /></TD>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & sAction & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE=""" & lReasonID & """ />"
				If InStr(1, sAction, "ServiceSheet", vbBinaryCompare) > 0 Then
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""" & lReasonID & """ />"
				End If
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""" & iLeftWidth & """>&nbsp;</TD>"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""32""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;&nbsp;&nbsp;</FONT></TD>"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""350""><INPUT TYPE=""SUBMIT"" NAME=""EmployeeConcept"" ID=""EmployeeConceptBtn"" VALUE=""Buscar empleado"" CLASS=""Buttons"" ALT=""" & sAltDescription & """/></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	Response.Write "</FORM>"

	DisplayAnotherEmployeeForm = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeeBeneficiaryForm(oRequest, oADODBConnection, sAction, sURL, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a registration of beneficiary for the
'         employee from the database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, sURL, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeBeneficiaryForm"
	Dim sNames
	Dim lErrorNumber

	If (aEmployeeComponent(N_ID_EMPLOYEE) <> -1) And (aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE) <> -1) Then
		lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckEmployeeBeneficiaryFields(oForm) {" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					If Len(oRequest("Delete").Item) > 0 Then Response.Write "return true;" & vbNewLine
					If StrComp(GetASPFileName(""), "Employees.asp", vbBinaryCompare) <> 0 Then
						Response.Write "if ((oForm.EmployeeID.value.length == 0) || (oForm.EmployeeID.value == '-1')) {" & vbNewLine
							Response.Write "alert('Favor de especificar el número de empleado.');" & vbNewLine
							Response.Write "oForm.EmployeeNumber.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (oForm.EmployeeNumber.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de especificar el número de empleado.');" & vbNewLine
							Response.Write "oForm.EmployeeNumber.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					End If
					Response.Write "if (oForm.BeneficiaryNumber.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de especificar el número del beneficiario.');" & vbNewLine
						Response.Write "oForm.BeneficiaryNumber.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (oForm.BeneficiaryName.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de especificar el nombre del beneficiario.');" & vbNewLine
						Response.Write "oForm.BeneficiaryName.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (oForm.BeneficiaryLastName.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de especificar el apellido paterno.');" & vbNewLine
						Response.Write "oForm.BeneficiaryLastName.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "oForm.AlimonyAmount.value = oForm.AlimonyAmount.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
					Response.Write "if (! CheckFloatValue(oForm.AlimonyAmount, 'el monto de la pensión alimenticia', N_MINIMUM_ONLY_FLAG, N_MINIMUM_OPEN_FLAG, 0, 0))" & vbNewLine
						Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckEmployeeBeneficiaryFields" & vbNewLine

		Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "<FORM NAME=""EmployeeBeneficiaryFrm"" ID=""EmployeeBeneficiaryFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckEmployeeBeneficiaryFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""EmployeesBeneficiaries"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SaveEmployeesMovements"" ID=""SaveEmployeeBeneficiariesHdn"" VALUE=""1"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Tab"" ID=""TabHdn"" VALUE=""1"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""" & oRequest("Step").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_EMPLOYEE) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BeneficiaryID"" ID=""BeneficiaryIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ActionHdn"" VALUE=""" & lReasonID & """ />"

			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				If StrComp(GetASPFileName(""), "Employees.asp", vbBinaryCompare) <> 0 Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. del empleado:&nbsp;</FONT></TD>"
						Response.Write "<TD>"
							Response.Write "<INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE="""" CLASS=""TextFields"" onChange=""document.EmployeeBeneficiaryFrm.EmployeeID.value='';"" />"
							Response.Write "<A HREF=""javascript: document.EmployeeBeneficiaryFrm.EmployeeID.value=''; SearchRecord(document.EmployeeBeneficiaryFrm.EmployeeNumber.value, 'EmployeeNumber', 'SearchEmployeeNumberIFrame', 'EmployeeBeneficiaryFrm.EmployeeID')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A>&nbsp;"
							Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""300"" HEIGHT=""22""></IFRAME>"
						Response.Write "</TD>"
					Response.Write "</TR>"
				End If
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Número del beneficiario:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""BeneficiaryNumber"" ID=""BeneficiaryNumberTxt"" VALUE="""
						If aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE) > 0 Then Response.Write aEmployeeComponent(N_NUMBER_BENEFICIARY_EMPLOYEE)
					Response.Write """ SIZE=""10"" MAXLENGTH=""10"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Nombre del beneficiario:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""BeneficiaryName"" ID=""BeneficiaryNameTxt"" VALUE=""" & aEmployeeComponent(S_NAME_BENEFICIARY_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Apellido paterno:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""BeneficiaryLastName"" ID=""BeneficiaryLastNameTxt"" VALUE=""" & aEmployeeComponent(S_LAST_NAME_BENEFICIARY_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Apellido materno:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""BeneficiaryLastName2"" ID=""BeneficiaryLastName2Txt"" VALUE=""" & aEmployeeComponent(S_LAST_NAME2_BENEFICIARY_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Fecha de nacimiento:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_BIRTH_DATE_BENEFICIARY_EMPLOYEE), "BeneficiaryBirth", N_FORM_START_YEAR, Year(Date()), True, True) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Fecha de inicio:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_START_DATE_BENEFICIARY_EMPLOYEE), "BeneficiaryStart", N_FORM_START_YEAR, Year(Date()), True, True) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Fecha de término:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_END_DATE_BENEFICIARY_EMPLOYEE), "BeneficiaryEnd", N_FORM_START_YEAR, Year(Date()) + 10, True, True) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Monto de la pensión alimenticia:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""AlimonyAmount"" ID=""AlimonyAmountTxt"" VALUE=""" & FormatNumber(aEmployeeComponent(D_ALIMONY_AMOUNT_BENEFICIARY_EMPLOYEE), 2, True, False, True) & """ SIZE=""20"" MAXLENGTH=""20"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Tipo de pensión:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><SELECT NAME=""AlimonyTypeID"" ID=""AlimonyTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "AlimonyTypes", "AlimonyTypeID", "AlimonyTypeID As Temp1, AlimonyTypeName", "", "AlimonyTypeID", aEmployeeComponent(N_ALIMONY_TYPE_ID_BENEFICIARY_EMPLOYEE), "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Centro de pago:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><SELECT NAME=""BeneficiaryPaymentCenterID"" ID=""BeneficiaryPaymentCenterIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "PaymentCenters", "PaymentCenterID", "PaymentCenterShortName, PaymentCenterName", "(Active=1)", "PaymentCenterShortName, PaymentCenterName", aEmployeeComponent(N_PAYMENT_CENTER_ID_BENEFICIARY_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR><TD COLSPAN=""2""><NOBR>"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Comentarios:<BR />"
					Response.Write "<TEXTAREA NAME=""BeneficiaryComments"" ID=""BeneficiaryCommentsTxtArea"" ROWS=""5"" COLS=""60"" MAXLENGTH=""2000"" CLASS=""TextFields"">" & aEmployeeComponent(S_COMMENTS_BENEFICIARY_EMPLOYEE) & "</TEXTAREA>"
				Response.Write "</NOBR></TD></TR>"
			Response.Write "</TABLE><BR />"
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				If Len(sURL) > 0 Then
					Response.Write "SendURLValuesToForm('" & sURL & "', document.EmployeeBeneficiaryFrm);" & vbNewLine
				End If
			Response.Write "//--></SCRIPT>" & vbNewLine

			If Len(oRequest("Delete").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveBeneficiaryWngDiv']); EmployeeBeneficiaryFrm.Remove.focus()"" />"
			ElseIf aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE) = -1 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			Else
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&Tab=1'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveBeneficiaryWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
		Response.Write "</FORM>"
	End If

	DisplayEmployeeBeneficiaryForm = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeeChildForm(oRequest, oADODBConnection, sASPFileName, sAction, sURL, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a registration of child for the
'         employee from the database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sASPFileName, sAction, sURL, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeChildForm"
	Dim sNames
	Dim lErrorNumber

	If (aEmployeeComponent(N_ID_EMPLOYEE) <> -1) And (aEmployeeComponent(N_ID_CHILD_EMPLOYEE) <> -1) Then
		lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckEmployeeChildFields(oForm) {" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					If Len(oRequest("Delete").Item) > 0 Then Response.Write "return true;" & vbNewLine
					If StrComp(GetASPFileName(""), "Employees.asp", vbBinaryCompare) <> 0 Then
						Response.Write "if ((oForm.EmployeeID.value.length == 0) || (oForm.EmployeeID.value == '-1')) {" & vbNewLine
							Response.Write "alert('Favor de especificar el número de empleado.');" & vbNewLine
							Response.Write "oForm.EmployeeNumber.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (oForm.EmployeeNumber.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de especificar el número de empleado.');" & vbNewLine
							Response.Write "oForm.EmployeeNumber.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					End If
					Response.Write "if (oForm.ChildName.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de especificar el nombre del hijo(a).');" & vbNewLine
						Response.Write "oForm.ChildName.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (oForm.ChildLastName.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de especificar el apellido paterno del hijo(a).');" & vbNewLine
						Response.Write "oForm.ChildLastName.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckEmployeeChildFields" & vbNewLine

		Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "<FORM NAME=""EmployeeChildFrm"" ID=""EmployeeChildFrm"" ACTION=""" & sASPFileName & """ METHOD=""POST"" onSubmit=""return CheckEmployeeChildFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & sAction & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SaveEmployeeChildren"" ID=""SaveEmployeesChildrenHdn"" VALUE=""1"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Tab"" ID=""TabHdn"" VALUE=""1"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""" & oRequest("Step").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_EMPLOYEE) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ChildID"" ID=""ChildIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_CHILD_EMPLOYEE) & """ />"

			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				If (StrComp(GetASPFileName(""), "UploadInfo.asp", vbBinaryCompare) = 0) And (StrComp(sAction, "ChildrenSchoolarships", vbBinaryCompare) = 0) Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. del empleado:&nbsp;</FONT></TD>"
						Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" VALUE=""" & aEmployeeComponent(N_ID_EMPLOYEE) & """ SIZE=""10"" MAXLENGTH=""10"" DISABLED=""DISABLED"" CLASS=""TextFields"" /></TD>"
					Response.Write "</TR>"
				ElseIf StrComp(GetASPFileName(""), "Employees.asp", vbBinaryCompare) <> 0 Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. del empleado:&nbsp;</FONT></TD>"
						Response.Write "<TD>"
							Response.Write "<INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE="""" CLASS=""TextFields"" onChange=""document.EmployeeChildFrm.EmployeeID.value='';"" />"
							Response.Write "<A HREF=""javascript: document.EmployeeChildFrm.EmployeeID.value=''; SearchRecord(document.EmployeeChildFrm.EmployeeNumber.value, 'EmployeeNumber', 'SearchEmployeeNumberIFrame', 'EmployeeChildFrm.EmployeeID')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A>&nbsp;"
							Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""300"" HEIGHT=""22""></IFRAME>"
						Response.Write "</TD>"
					Response.Write "</TR>"
				End If
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre del hijo(a):&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""ChildName"" ID=""ChildNameTxt"" VALUE=""" & aEmployeeComponent(S_NAME_CHILD_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido paterno:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""ChildLastName"" ID=""ChildLastNameTxt"" VALUE=""" & aEmployeeComponent(S_LAST_NAME_CHILD_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido materno:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""ChildLastName2"" ID=""ChildLastName2Txt"" VALUE=""" & aEmployeeComponent(S_LAST_NAME2_CHILD_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de nacimiento:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_BIRTH_DATE_CHILD_EMPLOYEE), "ChildBirth", N_FORM_START_YEAR, Year(Date()), True, True) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de término:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_END_DATE_CHILD_EMPLOYEE), "ChildEnd", N_FORM_START_YEAR, Year(Date()), True, True) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Grado escolar de la beca:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""LevelID"" ID=""LevelIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Schoolarships", "SchoolarshipID", "SchoolarshipName", "(Active=1)", "SchoolarshipID", aEmployeeComponent(N_CHILD_LEVEL_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE><BR />"
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				If Len(sURL) > 0 Then
					Response.Write "SendURLValuesToForm('" & sURL & "', document.EmployeeChildFrm);" & vbNewLine
				End If
			Response.Write "//--></SCRIPT>" & vbNewLine

			If Len(oRequest("Delete").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveChildWngDiv']); EmployeeChildFrm.Remove.focus()"" />"
			ElseIf aEmployeeComponent(N_ID_CHILD_EMPLOYEE) = -1 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			Else
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&Tab=1'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveChildWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
		Response.Write "</FORM>"
	End If

	DisplayEmployeeChildForm = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeeConceptForm(oRequest, oADODBConnection, sAction, sURL, sConceptIDs, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a concept for the
'         employee from the database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, sConceptIDs, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeConceptForm"
	Dim sNames
	Dim lErrorNumber

	If aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) > 0 Then
		lErrorNumber = GetEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
	End If

	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckConceptFields(oForm) {" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					If Len(oRequest("Delete").Item) > 0 Then Response.Write "return true;" & vbNewLine
					If StrComp(GetASPFileName(""), "Employees.asp", vbBinaryCompare) <> 0 Then
						Response.Write "if ((oForm.EmployeeID.value.length == 0) || (oForm.EmployeeID.value == '-1')) {" & vbNewLine
							Response.Write "alert('Favor de especificar el número de empleado.');" & vbNewLine
							Response.Write "oForm.EmployeeNumber.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					End If
					Response.Write "oForm.ConceptAmount.value = oForm.ConceptAmount.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
					Response.Write "if (! CheckFloatValue(oForm.ConceptAmount, 'el monto del concepto', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
						Response.Write "return false;" & vbNewLine
					If (InStr(1, sURL, ",EmployeesSafeSeparation,", vbBinaryCompare) = 0) Then
						Response.Write "oForm.ConceptMin.value = oForm.ConceptMin.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.ConceptMin, 'el monto mínimo del concepto', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "oForm.ConceptMax.value = oForm.ConceptMax.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.ConceptMax, 'el monto máximo del concepto', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "if (((oForm.ConceptQttyID.value == '2') || (oForm.ConceptQttyID.value == '8')) && (GetSelectedValues(oForm.AppliesToID) == '')) {" & vbNewLine
							Response.Write "alert('Seleccione el(los) concepto(s) que se utiliza(n) para calcular el concepto');" & vbNewLine
							Response.Write "oForm.AppliesToID.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					End If
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckConceptFields" & vbNewLine
			
			Response.Write "function ShowAmountFields(sValue) {" & vbNewLine
				Response.Write "var oForm = document.ConceptFrm;" & vbNewLine

				Response.Write "if (oForm) {" & vbNewLine
					Response.Write "HideDisplay(document.all['ConceptCurrencySpn']);" & vbNewLine
					Response.Write "HideDisplay(document.all['ConceptAppliesToSpn']);" & vbNewLine
					Response.Write "switch (sValue) {" & vbNewLine
						Response.Write "case '1':" & vbNewLine
							Response.Write "ShowDisplay(document.all['ConceptCurrencySpn']);" & vbNewLine
							Response.Write "break;" & vbNewLine
						Response.Write "case '2':" & vbNewLine
							Response.Write "ShowDisplay(document.all['ConceptAppliesToSpn']);" & vbNewLine
							Response.Write "break;" & vbNewLine
						Response.Write "case '8':" & vbNewLine
							Response.Write "ShowDisplay(document.all['ConceptAppliesToSpn']);" & vbNewLine
							Response.Write "break;" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of ShowAmountFields" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "<FORM NAME=""ConceptFrm"" ID=""ConceptFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckConceptFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""EmployeesConcepts"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SaveEmployeeConcept"" ID=""SaveEmployeeConceptHdn"" VALUE=""1"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Tab"" ID=""TabHdn"" VALUE=""3"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""" & oRequest("Step").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_EMPLOYEE) & """ />"

			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				If StrComp(GetASPFileName(""), "Employees.asp", vbBinaryCompare) <> 0 Then
					If StrComp(sConceptIDs, "87,120", vbBinaryCompare) <> 0 Then
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. del empleado:&nbsp;</FONT></TD>"
							Response.Write "<TD>"
								Response.Write "<INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE="""" CLASS=""TextFields"" onChange=""document.ConceptFrm.EmployeeID.value='';"" />"
								Response.Write "<A HREF=""javascript: document.ConceptFrm.EmployeeID.value=''; SearchRecord(document.ConceptFrm.EmployeeNumber.value, 'EmployeeNumber', 'SearchEmployeeNumberIFrame', 'ConceptFrm.EmployeeID')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A>&nbsp;"
								Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""22""></IFRAME>"
							Response.Write "</TD>"
						Response.Write "</TR>"
					Else
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. del empleado:&nbsp;</FONT></TD>"
							Response.Write "<TD>"
								Response.Write "<INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE="""" CLASS=""TextFields"" onChange=""document.ConceptFrm.EmployeeID.value='';"" />"
								Response.Write "<A HREF=""javascript: document.ConceptFrm.EmployeeID.value=''; SearchRecord(document.ConceptFrm.EmployeeNumber.value, 'EmployeeHeadNumber', 'SearchEmployeeNumberIFrame', 'ConceptFrm.EmployeeID')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A>&nbsp;"
								Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""600"" HEIGHT=""22""></IFRAME>"
							Response.Write "</TD>"
						Response.Write "</TR>"
					End If
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), "ConceptStart", N_FORM_START_YEAR, Year(Date()), True, True) & "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de fin de la vigencia:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE), "ConceptEnd", Year(Date()), Year(Date())+2, True, True) & "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR NAME=""PayrollDateDiv"" ID=""PayrollDateDiv"">"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Quincena de aplicación:&nbsp;</FONT></TD>"
						Response.Write "<TD><SELECT NAME=""EmployeePayrollDate"" ID=""EmployeePayrollDateCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""CheckStatusPayrolls()"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And (IsActive_7=1) And (PayrollTypeID=1)", "PayrollID Desc", aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
						Response.Write "</SELECT>&nbsp;"
						Response.Write "</TD>"
					Response.Write "</TR>"
				End If
				Response.Write "<TR NAME=""ConceptIDDiv"" ID=""ConceptIDDiv"">"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Concepto:&nbsp;</FONT></TD>"
					Response.Write "<TD>"
						If aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1 Then
							Response.Write "<SELECT NAME=""ConceptID"" ID=""ConceptIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								If Len(sConceptIDs) = 0 Then
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID>0) And (ConceptID Not In (" & aEmployeeComponent(S_EXCLUDED_CONCEPTS_ID_EMPLOYEE) & "))", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Else
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID>0) And (ConceptID In (" & sConceptIDs & "))", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								End If
							Response.Write "</SELECT>"
							If StrComp(GetASPFileName(""), "Employees.asp", vbBinaryCompare) <> 0 Then
								Response.Write "<A HREF=""javascript: if ((document.ConceptFrm.EmployeeID.value == '') || (document.ConceptFrm.EmployeeID.value == '-1')) {alert('Favor de especificar y validar el número de empleado');} else {SearchRecord(document.ConceptFrm.EmployeeNumber.value + '&ConceptID=' + document.ConceptFrm.ConceptID.value, 'EmployeeConcept', 'SearchEmployeeConceptIFrame', '')}""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Revisar si el empleado tiene registrado el concepto"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A>&nbsp;"
								Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeConceptIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""22""></IFRAME>"
							End If
						Else
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptID"" ID=""ConceptIDHdn"" VALUE=""" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & """ />"
							Call GetNameFromTable(oADODBConnection, "Concepts", aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
							Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT>"
						End If
					Response.Write "</TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					If (InStr(1, sURL, "EmployeesForRisk", vbBinaryCompare) > 0) Then
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Porcentaje:&nbsp;</FONT></TD>"
					Else
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Monto:&nbsp;</FONT></TD>"
					End If
					Response.Write "<TD>"
						If (InStr(1, sURL, "EmployeesForRisk", vbBinaryCompare) > 0) Then
							Response.Write "<SELECT NAME=""ConceptAmount"" ID=""ConceptAmountTxt"" SIZE=""1"" CLASS=""Lists""/>&nbsp;"
								Response.Write "<OPTION VALUE=10>10</OPTION>"
								Response.Write "<OPTION VALUE=20>20</OPTION>"
							Response.Write "</SELECT>&nbsp;"
						Else
							Response.Write "<INPUT TYPE=""TEXT"" NAME=""ConceptAmount"" ID=""ConceptAmountTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & FormatNumber(aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE), 2, True, False, True) & """ CLASS=""TextFields"" />&nbsp;"
						End If
						If (InStr(1, sURL, "EmployeesForRisk", vbBinaryCompare) > 0) Then
							Response.Write "<SELECT NAME=""ConceptQttyID"" ID=""ConceptQttyIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowAmountFields(this.value, 'Concept');"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "QttyID=2", "QttyID", "2", "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT>&nbsp;"
						Else						
							Response.Write "<SELECT NAME=""ConceptQttyID"" ID=""ConceptQttyIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowAmountFields(this.value, 'Concept');"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "", "QttyID", aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT>&nbsp;"						
						End If
						Response.Write "<SPAN ID=""ConceptCurrencySpn""><SELECT NAME=""CurrencyID"" ID=""CurrencyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Currencies", "CurrencyID", "CurrencyName", "(CurrencyID>-1) And (Active=1)", "CurrencyName", aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE), "Ninguna;;;-1", sErrorDescription)
						Response.Write "</SELECT></SPAN>"
						Response.Write "<SPAN ID=""ConceptAppliesToSpn"" STYLE=""display: none""><SELECT NAME=""AppliesToID"" ID=""AppliesToIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID In (Select ConceptID From EmployeesConceptsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")))", "ConceptShortName, ConceptName", aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT></SPAN>"
					Response.Write "</TD>"
				Response.Write "</TR>"
				If False Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo:&nbsp;</FONT></TD>"
						Response.Write "<TD><SELECT NAME=""ConceptTypeID"" ID=""ConceptTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "ConceptTypes", "ConceptTypeID", "ConceptTypeName", "", "ConceptTypeID", aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT></TD>"
					Response.Write "</TR>"
				Else
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptTypeID"" ID=""ConceptTypeIDHdn"" VALUE=""" & aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) & """ />"
				End If
				If False Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Mínimo:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
							Response.Write "<INPUT TYPE=""TEXT"" NAME=""ConceptMin"" ID=""ConceptMinTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & FormatNumber(aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE), 2, True, False, True) & """ CLASS=""TextFields"" />&nbsp;"
							Response.Write "<SELECT NAME=""ConceptMinQttyID"" ID=""ConceptMinQttyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "(QttyID In (1,3,5))", "QttyID", aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT>"
						Response.Write "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Máximo:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
							Response.Write "<INPUT TYPE=""TEXT"" NAME=""ConceptMax"" ID=""ConceptMaxTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & FormatNumber(aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE), 2, True, False, True) & """ CLASS=""TextFields"" />&nbsp;"
							Response.Write "<SELECT NAME=""ConceptMaxQttyID"" ID=""ConceptMaxQttyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "(QttyID In (1,3,5))", "QttyID", aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT>"
						Response.Write "</FONT></TD>"
					Response.Write "</TR>"
				Else
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptMin"" ID=""ConceptMinHdn"" VALUE=""0"" />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptMax"" ID=""ConceptMaxHdn"" VALUE=""0"" />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptMinQttyID"" ID=""ConceptMinQttyIDHdn"" VALUE=""1"" />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptMaxQttyID"" ID=""ConceptMaxQttyIDHdn"" VALUE=""1"" />"
				End If
				If False Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Ausencias:&nbsp;</FONT></TD>"
						Response.Write "<TD><SELECT NAME=""AbsenceTypeID"" ID=""AbsenceTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "AbsenceTypes", "AbsenceTypeID", "AbsenceTypeName", "", "AbsenceTypeID", aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Orden:&nbsp;</FONT></TD>"
						Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""ConceptOrder"" ID=""ConceptOrderTxt"" SIZE=""3"" MAXLENGTH=""3"" VALUE=""" & aEmployeeComponent(N_CONCEPT_ORDER_EMPLOYEE) & """ CLASS=""TextFields"" /></TD>"
					Response.Write "</TR>"
				Else
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AbsenceTypeID"" ID=""AbsenceTypeIDHdn"" VALUE=""" & aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptOrder"" ID=""ConceptOrderHdn"" VALUE=""" & aEmployeeComponent(N_CONCEPT_ORDER_EMPLOYEE) & """ />"
				End If
				If False Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Activo:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
							Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptActive"" ID=""ConceptActiveRd"" VALUE=""1"" "
								If aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 1 Then Response.Write " CHECKED=""1"""
							Response.Write " />Sí"
							Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""1"" />"
							Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptActive"" ID=""ConceptActiveRd"" VALUE=""0"""
								If aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0 Then Response.Write " CHECKED=""1"""
							Response.Write " />No"
						Response.Write "</FONT></TD>"
					Response.Write "</TR>"
				Else
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptActive"" ID=""ConceptActiveHdn"" VALUE=""" & aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) & """ />"
				End If

			Response.Write "</TABLE><BR />"
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "ShowAmountFields(document.ConceptFrm.ConceptQttyID.value, 'Concept');" & vbNewLine
				If Len(sURL) > 0 Then
					Response.Write "SendURLValuesToForm('" & sURL & "', document.ConceptFrm);" & vbNewLine
				End If
			Response.Write "//--></SCRIPT>" & vbNewLine

			If Len(oRequest("Delete").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveConceptWngDiv']); ConceptFrm.Remove.focus()"" />"
			ElseIf aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			Else
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&Tab=3'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveConceptWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
		Response.Write "</FORM>"
	End If

	DisplayEmployeeConceptForm = lErrorNumber
	Err.Clear
End Function

Function ShowEmployeeHistoryListForm(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display form to report the result
'         of the validation of information
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ShowEmployeeHistoryListForm"
	Dim oRecordset
	Dim sURL
	Dim lErrorNumber
	Dim sAlimonyTypeIDs
	Dim sCaseOptions
	Dim sPayrollsIDs
	Dim sEmployeeAntiquity
	Dim lAntiquityYears
	Dim lAntiquityMonths
	Dim lAntiquityDays
	Dim lCurrentDate
	Dim lStartDate
	Dim lEndDate
	Dim sCondition
	Dim iYears
	Dim iMonths
	Dim iDays
	Dim lLimitDate
	Dim sNames

	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckHistoryListFields(oForm) {" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					If Len(oRequest("Delete").Item) > 0 Then Response.Write "return true;" & vbNewLine
					Response.Write "if (oForm.JobID.value == '') {" & vbNewLine
						Response.Write "alert('Favor de introducir un número de plaza.');" & vbNewLine
						Response.Write "oForm.JobID.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (oForm.JobID.value != oForm.CheckJobID.value) {" & vbNewLine
						Response.Write "alert('Debe validar la nueva plaza que asignará al empleado.');" & vbNewLine
						Response.Write "oForm.JobID.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (oForm.NewPositionID.value == '-1') {" & vbNewLine
						Response.Write "alert('La nueva plaza no puede asignarse al empleado, verifique la información.');" & vbNewLine
						Response.Write "oForm.JobID.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if ((oForm.EndYear.value != '0') && (oForm.EndMonth.value != '0') && (oForm.EndDay.value != '0')) {" & vbNewLine
						Response.Write "if (parseInt('' + oForm.EndYear.value + oForm.EndMonth.value + oForm.EndDay.value) < parseInt('' + oForm.EmployeeYear.value + oForm.EmployeeMonth.value + oForm.EmployeeDay.value)) {" & vbNewLine
							Response.Write "alert('La fecha de término debe ser posterior a la fecha de inicio del registro.');" & vbNewLine
							Response.Write "oForm.EndYear.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckHistoryListFields" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "<FORM NAME=""HistoryListFrm"" ID=""HistoryListFrm"" ACTION=""" & sAction & """ METHOD=""GET"" onSubmit=""return CheckHistoryListFields(this)"">"
			Call DisplayURLParametersAsHiddenValues(aEmployeeComponent(S_URL_EMPLOYEE))
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""EmployeeHistoryList"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FilterStartYear"" ID=""FilterStartYearHdn"" VALUE=""" & oRequest("FilterStartYear").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FilterStartMonth"" ID=""FilterStartMonthHdn"" VALUE=""" & oRequest("FilterStartMonth").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FilterStartDay"" ID=""FilterStartDayHdn"" VALUE=""" & oRequest("FilterStartDay").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FilterEndYear"" ID=""FilterEndYearHdn"" VALUE=""" & oRequest("FilterEndYear").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FilterEndMonth"" ID=""FilterEndMonthHdn"" VALUE=""" & oRequest("FilterEndMonth").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FilterEndDay"" ID=""FilterEndDayHdn"" VALUE=""" & oRequest("FilterEndDay").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Tab"" ID=""TabHdn"" VALUE=""6"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""" & oRequest("Step").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReportID"" ID=""ReportIDHdn"" VALUE=""707"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_EMPLOYEE) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PayrollDate"" ID=""PayrollDateHdn"" VALUE=""" & oRequest("PayrollDate").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CheckJobID"" ID=""CheckJobIDHdn"" VALUE="""" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""NewPositionID"" ID=""NewPositionIDHdn"" VALUE="""" />"

			If aEmployeeComponent(N_ID_EMPLOYEE) <> -1 Then
				lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			End If
			If lErrorNumber = 0 Then
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Información general</B></FONT>"
				Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" />"
				Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del empleado:&nbsp;</FONT></TD>"
						If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
							Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" VALUE=""" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""SpecialTextFields"" onFocus=""document.EmployeeFrm.EmployeeName.focus()"" /></TD>"
						Else
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "</B></FONT></TD>"
						End If
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre(s):&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_NAME_EMPLOYEE) & "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido paterno:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_LAST_NAME_EMPLOYEE) & "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido materno:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_LAST_NAME2_EMPLOYEE) & "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartDate"" ID=""StartDateHdn"" VALUE=""" & aEmployeeComponent(N_START_DATE_EMPLOYEE) & """ />"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de tabulador:&nbsp;</FONT></TD>"
						Call GetNameFromTable(oADODBConnection, "EmployeeTypes", aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & sNames & "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Antigüedad ISSSTE:&nbsp;</FONT></TD>"
							If (CInt(aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE)) <> 6) And (CInt(aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE)) <> 7) Then
								lErrorNumber = CalculateEmployeeAntiquity(oADODBConnection, aEmployeeComponent, 0, sEmployeeAntiquity, lAntiquityYears, lAntiquityMonths, lAntiquityDays, sErrorDescription)
							Else
								lAntiquityYears = 0
								lAntiquityMonths = 0
								lAntiquityDays = 0
								sEmployeeAntiquity = "0 Años 0 Meses 0 Días"
							End If
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & sEmployeeAntiquity & "</TD>"
							Response.Write "<TD><INPUT TYPE=""HIDDEN"" NAME=""AntiquityYears"" ID=""AntiquityYearsTxt"" VALUE=""" & lAntiquityYears & """ CLASS=""TextFields"" />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AntiquityMonths"" ID=""AntiquityMonthsTxt"" VALUE=""" & lAntiquityMonths & """ CLASS=""TextFields"" />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AntiquityDays"" ID=""AntiquityDaysTxt"" VALUE=""" & lAntiquityDays & """ CLASS=""TextFields"" />"
							Response.Write "</TD>"
					Response.Write "</TR>"
				Response.Write "</TABLE>"
				If aEmployeeComponent(N_JOB_ID_EMPLOYEE) <> -1 Then
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Información de la plaza:</B></FONT>"
					Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" />"
					Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Plaza:&nbsp;</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>" & Right("000000" & aEmployeeComponent(N_JOB_ID_EMPLOYEE), Len("000000")) & "</B></FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Puesto:&nbsp;</FONT></TD>"
								Call GetNameFromTable(oADODBConnection, "Positions", aJobComponent(N_POSITION_ID_JOB), "", "", sNames, sErrorDescription)
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & sNames & "</FONT></TD>"
						Response.Write "</TR>"
						
						If aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 1 Then
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">GGN:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE) & "</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Clasificáción:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE) & "</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Integracion:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE) & "</FONT></TD>"
							Response.Write "</TR>"
						Else
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nivel:</FONT></TD><TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) & "</FONT></TD>"
							Response.Write "</TR>"
						End If
						Response.Write "</TR>"
'-----------------------------Se regresa funcionalidad anterior. 20130719
'							Response.Write "<SELECT NAME=""PositionID"" ID=""PositionIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowPopupItem('WaitDiv', parent.window.document.all['WaitDiv'], false); SearchRecord(this.value, 'PositionsCatalogsLKP', 'SearchPositionsCatalogsIFrame', 'JobFrm');"">"
'								sErrorDescription = "No se pudo obtener la información del registro."
'								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Positions.PositionID, Positions.PositionShortName, Positions.PositionName, Positions.EmployeeTypeID, EmployeeTypeName, CompanyName, GroupGradeLevelName, LevelName, ClassificationID, IntegrationID, WorkingHours From Positions, EmployeeTypes, Companies, GroupGradeLevels, Levels Where (Positions.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Positions.CompanyID=Companies.CompanyID) And (Positions.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (Positions.LevelID=Levels.LevelID) And (Positions.EndDate=30000000) And (EmployeeTypes.EndDate=30000000) And (Companies.EndDate=30000000) And (GroupGradeLevels.EndDate=30000000) And (Levels.EndDate=30000000) And (PositionID>-1) Order By Positions.PositionShortName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
'								If lErrorNumber = 0 Then
'									Do While Not oRecordset.EOF
'										Response.Write "<OPTION VALUE=""" & CStr(oRecordset.Fields("PositionID").Value) & """"
'											If aJobComponent(N_POSITION_ID_JOB) = CLng(oRecordset.Fields("PositionID").Value) Then Response.Write " SELECTED=""1"""
'										Response.Write ">" & CStr(oRecordset.Fields("PositionShortName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value) & " (Tabulador: " & CStr(oRecordset.Fields("EmployeeTypeName").Value) & ", Compañía: " & CStr(oRecordset.Fields("CompanyName").Value) & ", "
'											If CLng(oRecordset.Fields("EmployeeTypeID").Value) = 1 Then
'												Response.Write "GGN: " & CStr(oRecordset.Fields("GroupGradeLevelName").Value) & ", Clasificación:" & CStr(oRecordset.Fields("ClassificationID").Value) & ", Integración: " & CStr(oRecordset.Fields("IntegrationID").Value)
'											Else
'												Response.Write "Nivel: " & CStr(oRecordset.Fields("LevelName").Value)
'											End If
'										Response.Write ", Horas laboradas: " & CStr(oRecordset.Fields("WorkingHours").Value) & ")" & "</OPTION>"
'										oRecordset.MoveNext
'										If Err.number <> 0 Then Exit Do
'									Loop
'								End If
'							Response.Write "</SELECT>"
'-----------------------------
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de puesto:&nbsp;</FONT></TD>"
								Call GetNameFromTable(oADODBConnection, "PositionTypes", aJobComponent(N_POSITION_TYPE_ID_JOB), "", "", sNames, sErrorDescription)
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & sNames & "</FONT></TD>"
						Response.Write "</TR>"
					Response.Write "</TABLE>"
				End If
			Else
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Información general</B></FONT>"
				Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" />"
				Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
					Response.Write "<TR>"
						Response.Write "<TD>"
							Call DisplayErrorMessage("Error al obtener la información", "<BLOCKQUOTE>No se pudo obtener la información general que actualmente tiene el empleado.</BLOCKQUOTE>")
						Response.Write "</TD>"
					Response.Write "</TR>"
				Response.Write "</TABLE>"
			End If
			Response.Write "<IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
			Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Actualización de antigüedad</B></FONT>"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(0, "Employee", N_FORM_START_YEAR, Year(Date()) + 1, True, False) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de término:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(0, "End", N_FORM_START_YEAR, Year(Date()) + 1, True, True) & "</FONT></TD>"
				Response.Write "</TR>"
'				Response.Write "<TR>"
'					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del empleado:&nbsp;</FONT></TD>"
'					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" SIZE=""7"" MAXLENGTH=""7"" VALUE="""" CLASS=""TextFields"" /></TD>"
'				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número de la plaza:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						Response.Write "<INPUT TYPE=""TEXT"" NAME=""JobID"" ID=""JobIDTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE="""" CLASS=""TextFields"" onChange=""document.HistoryListFrm.NewPositionID.value='';"" />"
						Response.Write "<A HREF=""javascript: SearchRecord(document.HistoryListFrm.JobID.value,'PositionName&EmployeeYear='+HistoryListFrm.EmployeeYear.value+'&EmployeeMonth='+HistoryListFrm.EmployeeMonth.value+'&EmployeeDay='+HistoryListFrm.EmployeeDay.value+'&EndYear='+HistoryListFrm.EndYear.value+'&EndMonth='+HistoryListFrm.EndMonth.value+'&EndDay='+HistoryListFrm.EndDay.value,'SearchJobNumberIFrame', 'HistoryListFrm.NewPositionID')"">"
						Response.Write "<IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de la plaza"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A>"
					Response.Write "</TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Puesto:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & _
										"<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchJobNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""390"" HEIGHT=""20""></IFRAME>" & _
									"</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Estatus del empleado:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><SELECT NAME=""StatusID"" ID=""StatusIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "StatusEmployees", "StatusID", "StatusName", "", "StatusName", "", "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Tipo de movimiento:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><SELECT NAME=""ReasonID"" ID=""ReasonIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Reasons", "ReasonID", "ReasonShortName, ReasonName", "", "ReasonShortName", "", "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Activo:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						Response.Write "<INPUT TYPE=""RADIO"" NAME=""Active"" ID=""ActiveRd"" VALUE=""1"" />Sí"
						Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""1"" />"
						Response.Write "<INPUT TYPE=""RADIO"" NAME=""Active"" ID=""ActiveRd"" VALUE=""0"" />No"
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE><BR />"

			If Len(oRequest("Delete").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveConceptWngDiv']); ConceptFrm.Remove.focus()"" />"
			ElseIf Len(oRequest("EmployeeDate").Item) = 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			ElseIf StrComp(oRequest("EmployeeDate").Item, "0", vbBinaryCompare) = 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			Else
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			If CInt(Request.Cookies("SIAP_SubSectionID")) = 262 Then			
				Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&SectionID=262'"" />"
			Else
				Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&Tab=6&Change=1&ReportID=707'"" />"
			End If
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveConceptWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
		Response.Write "</FORM>"

		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "SendURLValuesToForm('EmployeeNumber=" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "&Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & "', document.HistoryListFrm);" & vbNewLine
			If Len(oRequest("EmployeeDate").Item) > 0 Then
				sErrorDescription = "No se pudieron obtener los registros de la base de datos."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate=" & oRequest("EmployeeDate").Item & ")", "EmployeeDisplayFormsComponentB.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						sURL = ""
						sURL = sURL & "EmployeeYear=" & Left(CStr(oRecordset.Fields("EmployeeDate").Value), Len("0000"))
						sURL = sURL & "&EmployeeMonth=" & Mid(CStr(oRecordset.Fields("EmployeeDate").Value), Len("00000"), Len("00"))
						sURL = sURL & "&EmployeeDay=" & Right(CStr(oRecordset.Fields("EmployeeDate").Value), Len("00"))
						If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
							sURL = sURL & "&EndYear=0&EndMonth=0&EndDay=0"
						Else
							sURL = sURL & "&EndYear=" & Left(CStr(oRecordset.Fields("EndDate").Value), Len("0000"))
							sURL = sURL & "&EndMonth=" & Mid(CStr(oRecordset.Fields("EndDate").Value), Len("00000"), Len("00"))
							sURL = sURL & "&EndDay=" & Right(CStr(oRecordset.Fields("EndDate").Value), Len("00"))
						End If
						sURL = sURL & "&EmployeeNumber=" & CStr(oRecordset.Fields("EmployeeNumber").Value)
						sURL = sURL & "&JobID=" & Right(("000000" & CStr(oRecordset.Fields("JobID").Value)), Len("000000"))
						sURL = sURL & "&CheckJobID=" & Right(("000000" & CStr(oRecordset.Fields("JobID").Value)), Len("000000"))
						sURL = sURL & "&EmployeeTypeID=" & CStr(oRecordset.Fields("EmployeeTypeID").Value)
						sURL = sURL & "&PositionTypeID=" & CStr(oRecordset.Fields("PositionTypeID").Value)
						sURL = sURL & "&StatusID=" & CStr(oRecordset.Fields("StatusID").Value)
						sURL = sURL & "&ReasonID=" & CStr(oRecordset.Fields("ReasonID").Value)
						sURL = sURL & "&Active=" & CStr(oRecordset.Fields("Active").Value)
						Response.Write "SendURLValuesToForm('" & sURL & "', document.HistoryListFrm);" & vbNewLine
					End If
					oRecordset.Close
				End If
			End If
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	Set oRecordset = Nothing
	ShowEmployeeHistoryListForm = lErrorNumber
	Err.Clear
End Function

Function DisplayFormForError(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display form to report the result
'         of the validation of information
'Inputs:  oRequest, oADODBConnection, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayFormForError"
	Dim sNames
	Dim lErrorNumber

	If (aEmployeeComponent(N_ID_EMPLOYEE) <> -1) Then
		lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckEmployeeErrorFields(oForm) {" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					If Len(oRequest("Delete").Item) > 0 Then Response.Write "return true;" & vbNewLine
					Response.Write "if (oForm.Comments.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de especificar los comentarios.');" & vbNewLine
						Response.Write "oForm.Comments.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckEmployeeErrorFields" & vbNewLine

		Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "<FORM NAME=""EmployeeErrorFrm"" ID=""EmployeeErrorFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""POST"" onSubmit=""return CheckEmployeeErrorFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""SaveEmployeeError"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeDate"" ID=""EmployeeDateHdn"" VALUE=""" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ActionHdn"" VALUE=""" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_EMPLOYEE) & """ />"
			If Len(oRequest("ValidationError").Item) > 0 Then
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ValidateFM1"" ID=""ValidateFM1Hdn"" VALUE=""1"" />"
			ElseIf Len(oRequest("AuthorizationError").Item) > 0 Then
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AuthorizationFM1"" ID=""AuthorizationFM1Hdn"" VALUE=""1"" />"
			End If
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del empleado:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_NAME_EMPLOYEE) & " " & aEmployeeComponent(S_LAST_NAME_EMPLOYEE) & " " & aEmployeeComponent(S_LAST_NAME2_EMPLOYEE) & "</FONT></TD>"
					Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de vigencia del movimiento:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Call GetNameFromTable(oADODBConnection, "Reasons", aEmployeeComponent(N_REASON_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de movimiento:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR><TD COLSPAN=""2"">"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Razones por las que no procede el movimiento:<BR /></FONT>"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2""><TEXTAREA NAME=""Comments"" ID=""CommentsTxtArea"" ROWS=""5"" COLS=""60"" MAXLENGTH=""2000"" CLASS=""TextFields"">" & aEmployeeComponent(S_REASON_FOR_REJECTION_COMMENTS_EMPLOYEE) & "</TEXTAREA></FONT>"
					Response.Write "</TD></TR>"
			Response.Write "</TABLE><BR />"
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				If Len(sURL) > 0 Then
					Response.Write "SendURLValuesToForm('" & sURL & "', document.EmployeeErrorFrm);" & vbNewLine
				End If
			Response.Write "//--></SCRIPT>" & vbNewLine

			If Len(oRequest("Delete").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveErrorWngDiv']); EmployeeErrorFrm.Remove.focus()"" />"
			ElseIf aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			Else
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Rechazar Movimiento"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&Tab=1'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveErrorWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
		Response.Write "</FORM>"
	End If

	DisplayFormForError = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeeAsHiddenFields(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about an employee using
'         hidden form fields
'Inputs:  oRequest, oADODBConnection, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeAsHiddenFields"

	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeNumber"" ID=""EmployeeNumberHdn"" VALUE=""" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeAccessKey"" ID=""EmployeeAccessKeyHdn"" VALUE=""" & aEmployeeComponent(S_ACCESS_KEY_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeePassword"" ID=""EmployeePasswordHdn"" VALUE=""" & aEmployeeComponent(S_PASSWORD_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeName"" ID=""EmployeeNameHdn"" VALUE=""" & aEmployeeComponent(S_NAME_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeLastName"" ID=""EmployeeLastNameHdn"" VALUE=""" & aEmployeeComponent(S_LAST_NAME_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeLastName2"" ID=""EmployeeLastName2Hdn"" VALUE=""" & aEmployeeComponent(S_LAST_NAME2_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CompanyID"" ID=""CompanyIDHdn"" VALUE=""" & aEmployeeComponent(N_COMPANY_ID_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobID"" ID=""JobIDHdn"" VALUE=""" & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ServiceID"" ID=""ServiceIDHdn"" VALUE=""" & aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDHdn"" VALUE=""" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PositionTypeID"" ID=""PositionTypeIDHdn"" VALUE=""" & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ClassificationID"" ID=""ClassificationIDHdn"" VALUE=""" & aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""GroupGradeLevelID"" ID=""GroupGradeLevelIDHdn"" VALUE=""" & aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IntegrationID"" ID=""IntegrationIDHdn"" VALUE=""" & aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JourneyID"" ID=""JourneyIDHdn"" VALUE=""" & aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ShiftID"" ID=""ShiftIDHdn"" VALUE=""" & aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartHour1"" ID=""StartHour1Hdn"" VALUE=""" & aEmployeeComponent(N_START_HOUR_1_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EndHour1"" ID=""EndHour1Hdn"" VALUE=""" & aEmployeeComponent(N_END_HOUR_1_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartHour2"" ID=""StartHour2Hdn"" VALUE=""" & aEmployeeComponent(N_START_HOUR_2_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EndHour2"" ID=""EndHour2Hdn"" VALUE=""" & aEmployeeComponent(N_END_HOUR_2_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartHour3"" ID=""StartHour3Hdn"" VALUE=""" & aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EndHour3"" ID=""EndHour3Hdn"" VALUE=""" & aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""WorkingHours"" ID=""WorkingHoursHdn"" VALUE=""" & aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""LevelID"" ID=""LevelIDHdn"" VALUE=""" & aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StatusID"" ID=""StatusIDHdn"" VALUE=""" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PaymentCenterID"" ID=""PaymentCenterIDHdn"" VALUE=""" & aEmployeeComponent(N_PAYMENT_CENTER_ID_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RiskLevel"" ID=""RiskLevelHdn"" VALUE=""" & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeEmail"" ID=""EmployeeEmailHdn"" VALUE=""" & aEmployeeComponent(S_EMAIL_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SocialSecurityNumber"" ID=""SocialSecurityNumberHdn"" VALUE=""" & aEmployeeComponent(S_SSN_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BirthDate"" ID=""BirthDateHdn"" VALUE=""" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartDate"" ID=""StartDateHdn"" VALUE=""" & aEmployeeComponent(N_START_DATE_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartDate2"" ID=""StartDate2Hdn"" VALUE=""" & aEmployeeComponent(N_START_DATE2_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CountryID"" ID=""CountryIDHdn"" VALUE=""" & aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RFC"" ID=""RFCHdn"" VALUE=""" & aEmployeeComponent(S_RFC_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CURP"" ID=""CURPHdn"" VALUE=""" & aEmployeeComponent(S_CURP_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""GenderID"" ID=""GenderIDHdn"" VALUE=""" & aEmployeeComponent(N_GENDER_ID_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""MaritalStatusID"" ID=""MaritalStatusIDHdn"" VALUE=""" & aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Active"" ID=""ActiveHdn"" VALUE=""" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeDate"" ID=""EmployeeDateHdn"" VALUE=""" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeEndDate"" ID=""EmployeeEndDateHdn"" VALUE=""" & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeComments"" ID=""EmployeeCommentsHdn"" VALUE=""" & aEmployeeComponent(S_COMMENTS_EMPLOYEE) & """ />"
	
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeAddress"" ID=""EmployeeAddressHdn"" VALUE=""" & aEmployeeComponent(S_ADDRESS_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeCity"" ID=""EmployeeCityHdn"" VALUE=""" & aEmployeeComponent(S_CITY_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeZipCode"" ID=""EmployeeZipCodeHdn"" VALUE=""" & aEmployeeComponent(S_ZIP_CODE_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StateID"" ID=""StateIDHdn"" VALUE=""" & aEmployeeComponent(N_ADDRESS_STATE_ID_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AddressCountryID"" ID=""AddressCountryIDHdn"" VALUE=""" & aEmployeeComponent(N_ADDRESS_COUNTRY_ID_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeePhone"" ID=""EmployeePhoneHdn"" VALUE=""" & aEmployeeComponent(S_EMPLOYEE_PHONE_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OfficePhone"" ID=""OfficePhoneHdn"" VALUE=""" & aEmployeeComponent(S_OFFICE_PHONE_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OfficeExt"" ID=""OfficeExtHdn"" VALUE=""" & aEmployeeComponent(S_EXT_OFFICE_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DocumentNumber1"" ID=""DocumentNumber1Hdn"" VALUE=""" & aEmployeeComponent(S_DOCUMENT_NUMBER_1_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DocumentNumber2"" ID=""DocumentNumber2Hdn"" VALUE=""" & aEmployeeComponent(S_DOCUMENT_NUMBER_2_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DocumentNumber3"" ID=""DocumentNumber3Hdn"" VALUE=""" & aEmployeeComponent(S_DOCUMENT_NUMBER_3_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeActivityID"" ID=""EmployeeActivityIDHdn"" VALUE=""" & aEmployeeComponent(N_ACTIVITY_ID_EMPLOYEE) & """ />"

	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ChildID"" ID=""ChildIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_CHILD_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ChildName"" ID=""ChildNameHdn"" VALUE=""" & aEmployeeComponent(S_NAME_CHILD_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ChildLastName"" ID=""ChildLastNameHdn"" VALUE=""" & aEmployeeComponent(S_LAST_NAME_CHILD_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ChildLastName2"" ID=""ChildLastName2Hdn"" VALUE=""" & aEmployeeComponent(S_LAST_NAME2_CHILD_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ChildBirthDate"" ID=""ChildBirthDateHdn"" VALUE=""" & aEmployeeComponent(N_BIRTH_DATE_CHILD_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ChildEndDate"" ID=""ChildEndDateHdn"" VALUE=""" & aEmployeeComponent(N_END_DATE_CHILD_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SchoolarshipID"" ID=""SchoolarshipIDHdn"" VALUE=""" & aEmployeeComponent(N_CHILD_LEVEL_ID_EMPLOYEE) & """ />"

	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BeneficiaryID"" ID=""BeneficiaryIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_CHILD_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BeneficiaryStartDate"" ID=""BeneficiaryStartDateHdn"" VALUE=""" & aEmployeeComponent(N_ID_CHILD_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BeneficiaryEndDate"" ID=""BeneficiaryEndDateHdn"" VALUE=""" & aEmployeeComponent(N_ID_CHILD_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BeneficiaryNumber"" ID=""BeneficiaryNumberHdn"" VALUE=""" & aEmployeeComponent(N_ID_CHILD_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BeneficiaryName"" ID=""BeneficiaryNameHdn"" VALUE=""" & aEmployeeComponent(N_ID_CHILD_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BeneficiaryLastName"" ID=""BeneficiaryLastNameHdn"" VALUE=""" & aEmployeeComponent(N_ID_CHILD_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BeneficiaryLastName2"" ID=""BeneficiaryLastName2Hdn"" VALUE=""" & aEmployeeComponent(N_ID_CHILD_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BeneficiaryBirthDate"" ID=""BeneficiaryBirthDateHdn"" VALUE=""" & aEmployeeComponent(N_ID_CHILD_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AlimonyAmount"" ID=""AlimonyAmountHdn"" VALUE=""" & aEmployeeComponent(N_ID_CHILD_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AlimonyTypeID"" ID=""AlimonyTypeIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_CHILD_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BeneficiaryPaymentCenterID"" ID=""BeneficiaryPaymentCenterIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_CHILD_EMPLOYEE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BeneficiaryComments"" ID=""BeneficiaryCommentsHdn"" VALUE=""" & aEmployeeComponent(N_ID_CHILD_EMPLOYEE) & """ />"

	DisplayEmployeeAsHiddenFields = Err.Number
	Err.Clear
End Function

Function DisplayEmployeeSafeSeparationForm(oRequest, oADODBConnection, sAction, sURL, sConceptIDs, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a concept for the
'         employee from the database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, sConceptIDs, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeSafeSeparationForm"
	Dim sNames
	Dim lErrorNumber

	If aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) > 0 Then
		lErrorNumber = GetEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
	End If

	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckConceptFields(oForm) {" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					If Len(oRequest("Delete").Item) > 0 Then Response.Write "return true;" & vbNewLine
					If StrComp(GetASPFileName(""), "Employees.asp", vbBinaryCompare) <> 0 Then
						Response.Write "if ((oForm.EmployeeID.value.length == 0) || (oForm.EmployeeID.value == '-1')) {" & vbNewLine
							Response.Write "alert('Favor de especificar el número de empleado.');" & vbNewLine
							Response.Write "oForm.EmployeeNumber.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					End If
					Response.Write "oForm.ConceptAmount.value = oForm.ConceptAmount.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
					Response.Write "if (! CheckFloatValue(oForm.ConceptAmount, 'el monto del concepto', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
						Response.Write "return false;" & vbNewLine
					If (InStr(1, sURL, ",EmployeesSafeSeparation,", vbBinaryCompare) = 0) Then
						Response.Write "oForm.ConceptMin.value = oForm.ConceptMin.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.ConceptMin, 'el monto mínimo del concepto', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "oForm.ConceptMax.value = oForm.ConceptMax.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.ConceptMax, 'el monto máximo del concepto', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "if (((oForm.ConceptQttyID.value == '2') || (oForm.ConceptQttyID.value == '8')) && (GetSelectedValues(oForm.AppliesToID) == '')) {" & vbNewLine
							Response.Write "alert('Seleccione el(los) concepto(s) que se utiliza(n) para calcular el concepto');" & vbNewLine
							Response.Write "oForm.AppliesToID.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					End If
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckConceptFields" & vbNewLine
			
			Response.Write "function ShowAmountFields(sValue) {" & vbNewLine
				Response.Write "var oForm = document.ConceptFrm;" & vbNewLine

				Response.Write "if (oForm) {" & vbNewLine
					Response.Write "HideDisplay(document.all['ConceptCurrencySpn']);" & vbNewLine
					Response.Write "HideDisplay(document.all['ConceptAppliesToSpn']);" & vbNewLine
					Response.Write "switch (sValue) {" & vbNewLine
						Response.Write "case '1':" & vbNewLine
							Response.Write "ShowDisplay(document.all['ConceptCurrencySpn']);" & vbNewLine
							Response.Write "break;" & vbNewLine
						Response.Write "case '2':" & vbNewLine
							Response.Write "ShowDisplay(document.all['ConceptAppliesToSpn']);" & vbNewLine
							Response.Write "break;" & vbNewLine
						Response.Write "case '8':" & vbNewLine
							Response.Write "ShowDisplay(document.all['ConceptAppliesToSpn']);" & vbNewLine
							Response.Write "break;" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of ShowAmountFields" & vbNewLine
			
			Response.Write "function ShowPercentSIFields(sValue) {" & vbNewLine
				Response.Write "var oForm = document.ConceptFrm;" & vbNewLine

				Response.Write "if (oForm) {" & vbNewLine
					Response.Write "HideDisplay(document.all['ConceptCurrencySpn']);" & vbNewLine
					Response.Write "HideDisplay(document.all['ConceptAppliesToSpn']);" & vbNewLine
					Response.Write "switch (sValue) {" & vbNewLine
						Response.Write "case '120':" & vbNewLine
							Response.Write "ShowDisplay(document.all['ConceptCurrencySpn']);" & vbNewLine
							Response.Write "break;" & vbNewLine
						Response.Write "case '2':" & vbNewLine
							Response.Write "ShowDisplay(document.all['ConceptAppliesToSpn']);" & vbNewLine
							Response.Write "break;" & vbNewLine
						Response.Write "case '8':" & vbNewLine
							Response.Write "ShowDisplay(document.all['ConceptAppliesToSpn']);" & vbNewLine
							Response.Write "break;" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of ShowAmountFields" & vbNewLine			
			
		Response.Write "//--></SCRIPT>" & vbNewLine

		Response.Write "<FORM NAME=""ConceptFrm"" ID=""ConceptFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckConceptFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""EmployeeSafeSeparation"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SaveEmployeeSafeSeparation"" ID=""SaveEmployeeSafeSeparationHdn"" VALUE=""1"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Tab"" ID=""TabHdn"" VALUE=""3"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""" & oRequest("Step").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_EMPLOYEE) & """ />"
			If (InStr(1, sURL, "EmployeesSafeSeparation", vbBinaryCompare) > 0) Then
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Type"" ID=""TypeHdn"" VALUE=""SI"" />"
			Else
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Type"" ID=""TypeHdn"" VALUE=""AE"" />"
			End If
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				If StrComp(GetASPFileName(""), "Employees.asp", vbBinaryCompare) <> 0 Then
					If StrComp(sConceptIDs, "87,120", vbBinaryCompare) <> 0 Then
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. del empleado:&nbsp;</FONT></TD>"
							Response.Write "<TD>"
								Response.Write "<INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE="""" CLASS=""TextFields"" onChange=""document.ConceptFrm.EmployeeID.value='';"" />"
								Response.Write "<A HREF=""javascript: document.ConceptFrm.EmployeeID.value=''; SearchRecord(document.ConceptFrm.EmployeeNumber.value, 'EmployeeNumber', 'SearchEmployeeNumberIFrame', 'ConceptFrm.EmployeeID')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A>&nbsp;"
								Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""22""></IFRAME>"
							Response.Write "</TD>"
						Response.Write "</TR>"
					Else
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. del empleado:&nbsp;</FONT></TD>"
							Response.Write "<TD>"
								Response.Write "<INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE="""" CLASS=""TextFields"" onChange=""document.ConceptFrm.EmployeeID.value='';"" />"
								Response.Write "<A HREF=""javascript: document.ConceptFrm.EmployeeID.value=''; SearchRecord(document.ConceptFrm.EmployeeNumber.value, 'EmployeeHeadNumber', 'SearchEmployeeNumberIFrame', 'ConceptFrm.EmployeeID')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A>&nbsp;"
								Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""600"" HEIGHT=""22""></IFRAME>"
							Response.Write "</TD>"
						Response.Write "</TR>"
					End If
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), "ConceptStart", N_FORM_START_YEAR, Year(Date()), True, True) & "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR NAME=""PayrollDateDiv"" ID=""PayrollDateDiv"">"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Quincena de aplicación:&nbsp;</FONT></TD>"
						Response.Write "<TD><SELECT NAME=""EmployeePayrollDate"" ID=""EmployeePayrollDateCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""CheckStatusPayrolls()"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And (IsActive_1=1) And (PayrollTypeID=1)", "PayrollID Desc", "2", "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT>&nbsp;"
						Response.Write "</TD>"
					Response.Write "</TR>"
				End If
				Response.Write "<TR NAME=""ConceptIDDiv"" ID=""ConceptIDDiv"">"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Concepto:&nbsp;</FONT></TD>"
					Response.Write "<TD>"
						If aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1 Then
							Response.Write "<SELECT NAME=""ConceptID"" ID=""ConceptIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								If (InStr(1, sURL, "EmployeesSafeSeparation", vbBinaryCompare) > 0) Then
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 120)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Else
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 87)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								End If
							Response.Write "</SELECT>"
							If StrComp(GetASPFileName(""), "Employees.asp", vbBinaryCompare) <> 0 Then
								Response.Write "<A HREF=""javascript: if ((document.ConceptFrm.EmployeeID.value == '') || (document.ConceptFrm.EmployeeID.value == '-1')) {alert('Favor de especificar y validar el número de empleado');} else {SearchRecord(document.ConceptFrm.EmployeeNumber.value + '&ConceptID=' + document.ConceptFrm.ConceptID.value, 'EmployeeConcept', 'SearchEmployeeConceptIFrame', '')}""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Revisar si el empleado tiene registrado el concepto"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A>&nbsp;"
								Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeConceptIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""22""></IFRAME>"
							End If
						Else
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptID"" ID=""ConceptIDHdn"" VALUE=""" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & """ />"
							Call GetNameFromTable(oADODBConnection, "Concepts", aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
							Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT>"
						End If
					Response.Write "</TD>"
					Response.Write "</TR>"
				Response.Write "<TR>"
					If (InStr(1, sURL, "EmployeesSafeSeparation", vbBinaryCompare) > 0) Then
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Porcentaje:&nbsp;</FONT></TD>"
					Else
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cantidad ($ o %):&nbsp;</FONT></TD>"
					End If
					Response.Write "<TD>"
						If (InStr(1, sURL, "EmployeesSafeSeparation", vbBinaryCompare) > 0) Then
							Response.Write "<SELECT NAME=""ConceptAmount"" ID=""ConceptAmountTxt"" SIZE=""1"" CLASS=""Lists"" />&nbsp;"
								Response.Write "<OPTION VALUE=0>0</OPTION>"
								Response.Write "<OPTION VALUE=2>2</OPTION>"
								Response.Write "<OPTION VALUE=4>4</OPTION>"
								Response.Write "<OPTION VALUE=5>5</OPTION>"
								Response.Write "<OPTION VALUE=10>10</OPTION>"
							Response.Write "</SELECT>&nbsp;"
						Else
							Response.Write "<INPUT TYPE=""TEXT"" NAME=""ConceptAmount"" ID=""ConceptAmountTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE="""" CLASS=""TextFields"" />&nbsp;"
						End If
						If (InStr(1, sURL, "EmployeesSafeSeparation", vbBinaryCompare) > 0) Then
							Response.Write "<SELECT NAME=""ConceptQttyID"" ID=""ConceptQttyIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowAmountFields(this.value);"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "QttyID=2", "QttyID", "2", "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT>&nbsp;"
							Response.Write "<SPAN ID=""ConceptCurrencySpn"">"
							Response.Write "</SPAN>"
							Response.Write "<SPAN ID=""ConceptAppliesToSpn"" STYLE=""display: none"">"
							Response.Write "</SPAN>"
						Else
							Response.Write "<SELECT NAME=""ConceptQttyID"" ID=""ConceptQttyIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowAmountFields(this.value);"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "QttyID=1 or QttyID=2", "QttyID", aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT>&nbsp;"
							Response.Write "<SPAN ID=""ConceptCurrencySpn""><SELECT NAME=""CurrencyID"" ID=""CurrencyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Currencies", "CurrencyID", "CurrencyName", "(CurrencyID=0) And (Active=1)", "CurrencyName", aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE), "Ninguna;;;-1", sErrorDescription)
							Response.Write "</SELECT></SPAN>"
							Response.Write "<SPAN ID=""ConceptAppliesToSpn"" STYLE=""display: none""><SELECT NAME=""AppliesToID"" ID=""AppliesToIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID In (Select ConceptID From EmployeesConceptsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")))", "ConceptShortName, ConceptName", aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT></SPAN>"
						End If
					Response.Write "</TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE><BR />"
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "ShowAmountFields(document.ConceptFrm.ConceptQttyID.value);" & vbNewLine
				If Len(sURL) > 0 Then
					Response.Write "SendURLValuesToForm('" & sURL & "', document.ConceptFrm);" & vbNewLine
				End If
			Response.Write "//--></SCRIPT>" & vbNewLine

			If Len(oRequest("Delete").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveConceptWngDiv']); ConceptFrm.Remove.focus()"" />"
			ElseIf aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			Else
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&Tab=3'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveConceptWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
		Response.Write "</FORM>"
	End If

	DisplayEmployeeSafeSeparationForm = lErrorNumber
	Err.Clear
End Function

Function DisplayFormForEmployee(oADODBConnection, bForExport, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the details for the given employee form
'Inputs:  oRequest, oADODBConnection, bForExport, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayFormForEmployee"
	Dim sFormContents
	Dim sField
	Dim sQueryForSource
	Dim asFields
	Dim asValues
	Dim sURL
	Dim iIndex
	Dim sMinType
	Dim sMaxType
	Dim bTemplate
	Dim asIDs
	Dim oFieldADODBConnection
	Dim oRecordset
	Dim oCatalogRecordset
	Dim sAnswer
	Dim asFiles
	Dim aEmailComponent
	Dim lErrorNumber

	sFormContents = ""
	If FileExists(Server.MapPath(TEMPLATES_PHYSICAL_PATH & "EmployeeForm.htm"), sErrorDescription) Then
		sFormContents = GetFileContents(Server.MapPath(TEMPLATES_PHYSICAL_PATH & "EmployeeForm.htm"), sErrorDescription)
	End If
	bTemplate = (Len(sFormContents) > 0)

	If Not bForExport Then
		sErrorDescription = "No se pudo obtener el formulario para el módulo."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeeFields Order By FormFieldID", "EmployeeDisplayFormsComponentB.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Response.Write "<FORM NAME=""EmployeeFieldsFrm"" ID=""EmployeeFieldsFrm"" ACTION=""" & GetASPFileName("") & """ METHID=""POST"">"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_EMPLOYEE) & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Tab"" ID=""TabHdn"" VALUE=""" & oRequest("Tab").Item & """ />"
				Do While Not oRecordset.EOF
					sQueryForSource = ""
					sQueryForSource = oRecordset.Fields("QueryForSource").Value
					Err.Clear
					If Len(sQueryForSource) > 0 Then
						Call TransformXMLTagsForEmployeeForm(aEmployeeComponent, False, sQueryForSource, sErrorDescription)
						lErrorNumber = CreateADODBConnection(CStr(oRecordset.Fields("DSNForSource").Value), "", "", CInt(oRecordset.Fields("ConnectionTypeForSource").Value), oFieldADODBConnection, sErrorDescription)
					End If

					sField = ""
					If lErrorNumber = 0 Then
						Select Case CLng(oRecordset.Fields("FieldTypeID").Value)
							Case 0 'Booleano
								sField = "<INPUT TYPE=""CHECKBOX"" NAME=""EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & """ ID=""EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Chk"" VALUE=""1"" "
									If Len(sQueryForSource) > 0 Then
										sErrorDescription = "No se pudo obtener el valor por default para el campo de texto del formulario."
										lErrorNumber = ExecuteSQLQuery(oFieldADODBConnection, sQueryForSource, "EmployeeDisplayFormsComponentB.asp", S_FUNCTION_NAME, 0, sErrorDescription, oCatalogRecordset)
										If lErrorNumber = 0 Then
											If Not oCatalogRecordset.EOF Then
												If CInt(oCatalogRecordset.Fields("Active").Value) = 1 Then sField = sField & " CHECKED=""1"""
											End If
										End If
									End If
									If Not IsNull(oRecordset.Fields("DefaultValue")) Then
										If Len(oRecordset.Fields("DefaultValue")) > 0 Then sField = sField & " CHECKED=""1"""
									End If
									If Not IsNull(oRecordset.Fields("JavaScriptCode")) Then sField = sField & CStr(oRecordset.Fields("JavaScriptCode").Value)
								sField = sField & " />"
								If Not bTemplate Then sField = sField & CStr(oRecordset.Fields("FormFieldText").Value) & "<BR />"
							Case 1 'Fecha
								If Not bTemplate Then sField = CStr(oRecordset.Fields("FormFieldText").Value) & ": "
								If Not IsNull(oRecordset.Fields("DefaultValue")) Then
									sAnswer = CStr(oRecordset.Fields("DefaultValue").Value)
								Else
									sAnswer = "0"
								End If
								If CLng(oRecordset.Fields("MinimumValue").Value) < 1 Then
									sField = sField & DisplayDateCombosUsingSerial(sAnswer, "EF__" & CStr(oRecordset.Fields("FormFieldID").Value), Year(Date()) - Abs(CLng(oRecordset.Fields("MinimumValue").Value)), Year(Date()) + Abs(CLng(oRecordset.Fields("MaximumValue").Value)), True, (CInt(oRecordset.Fields("IsOptional").Value) = 1))
								Else
									sField = sField & DisplayDateCombosUsingSerial(sAnswer, "EF__" & CStr(oRecordset.Fields("FormFieldID").Value), CLng(oRecordset.Fields("MinimumValue").Value), Year(Date()) + Abs(CLng(oRecordset.Fields("MaximumValue").Value)), True, (CInt(oRecordset.Fields("IsOptional").Value) = 1))
								End If
								If Not bTemplate Then sField = sField & "<BR />"
							Case 2, 4 'Flotante, Numérico
								If Not bTemplate Then sField = CStr(oRecordset.Fields("FormFieldText").Value) & ": "
								sField = sField & "<INPUT TYPE=""TEXT"" NAME=""EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & """ ID=""EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Txt"" SIZE=""20"" MAXLENGTH=""20"" VALUE="""
									If Not IsNull(oRecordset.Fields("DefaultValue")) Then sField = sField & CStr(oRecordset.Fields("DefaultValue").Value)
								sField = sField & """ CLASS=""TextFields"" "
									If Not IsNull(oRecordset.Fields("JavaScriptCode")) Then sField = sField & CStr(oRecordset.Fields("JavaScriptCode").Value)
								sField = sField & " />"
								If Not bTemplate Then sField = sField & "<BR />"
							Case 3 'Hora
								If Not bTemplate Then sField = CStr(oRecordset.Fields("FormFieldText").Value) & ": "
								If Not IsNull(oRecordset.Fields("DefaultValue")) Then
									sAnswer = CStr(oRecordset.Fields("DefaultValue").Value)
								Else
									sAnswer = "0"
								End If
								sField = sField & DisplayTimeCombosUsingSerial(sAnswer, "EF__" & CStr(oRecordset.Fields("FormFieldID").Value), 0, 24, 1, (CInt(oRecordset.Fields("IsOptional").Value) = 1))
								If Not bTemplate Then sField = sField & "<BR />"
							Case 5 'Texto
								If Not bTemplate Then sField = CStr(oRecordset.Fields("FormFieldText").Value) & ": "
								If CLng(oRecordset.Fields("FormFieldSize").Value) > 255 Then
									If Not bTemplate Then sField = sField & "<BR />"
									sField = sField & "<TEXTAREA NAME=""EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & """ ID=""EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & "TxtArea"" "
									If CLng(oRecordset.Fields("FormFieldSize").Value) < 1000 Then
										sField = sField & "ROWS=""6"""
									Else
										sField = sField & "ROWS=""20"""
									End If
									sField = sField & " COLS=""50"" MAXLENGTH=""" & CStr(oRecordset.Fields("FormFieldSize").Value) & """ CLASS=""TextFields"" "
										If Not IsNull(oRecordset.Fields("JavaScriptCode")) Then sField = sField & CStr(oRecordset.Fields("JavaScriptCode").Value)
									sField = sField & ">"
								Else
									sField = sField & "<INPUT TYPE=""TEXT"" NAME=""EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & """ ID=""EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Txt"" SIZE="""
									If CLng(oRecordset.Fields("FormFieldSize").Value) < 30 Then
										sField = sField & CStr(oRecordset.Fields("FormFieldSize").Value)
									Else
										sField = sField & "30"
									End If
									sField = sField & """ MAXLENGTH=""" & CStr(oRecordset.Fields("FormFieldSize").Value) & """ VALUE="""
								End If
									If Len(sQueryForSource) > 0 Then
										sErrorDescription = "No se pudo obtener el valor por default para el campo de texto del formulario."
										lErrorNumber = ExecuteSQLQuery(oFieldADODBConnection, sQueryForSource, "EmployeeDisplayFormsComponentB.asp", S_FUNCTION_NAME, 0, sErrorDescription, oCatalogRecordset)
										If lErrorNumber = 0 Then
											If Not oCatalogRecordset.EOF Then
												For iIndex = 0 To oCatalogRecordset.Fields.Count - 1
													sField = sField & CStr(oCatalogRecordset.Fields(iIndex).Value)
													Err.Clear
													If iIndex < oCatalogRecordset.Fields.Count - 1 Then sField = sField & " "
												Next
											End If
										End If
									ElseIf Not IsNull(oRecordset.Fields("DefaultValue")) Then
										sField = sField & CStr(oRecordset.Fields("DefaultValue").Value)
									End If
								If CLng(oRecordset.Fields("FormFieldSize").Value) > 255 Then
									sField = sField & "</TEXTAREA>"
									If Not bTemplate Then sField = sField & "<BR />"
								Else
									sField = sField & """ CLASS=""TextFields"" "
										If Not IsNull(oRecordset.Fields("JavaScriptCode")) Then sField = sField & CStr(oRecordset.Fields("JavaScriptCode").Value)
									sField = sField & " />"
								End If
								If Not bTemplate Then sField = sField & "<BR />"
							Case 6, 8 'Catálogo, Lista
								If lErrorNumber = 0 Then
									If Not bTemplate Then
										sField = CStr(oRecordset.Fields("FormFieldText").Value) & ": "
										If CLng(oRecordset.Fields("FieldTypeID").Value) = 8 Then sField = sField & "<BR />"
									End If
									sField = sField & "<SELECT NAME=""EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & """ ID=""EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Cmb"""
									If CLng(oRecordset.Fields("FieldTypeID").Value) = 6 Then
										sField = sField & " SIZE=""1"""
									Else
										sField = sField & " SIZE=""5"" MULTIPLE=""1"""
									End If
									sField = sField & " VALUE="""" CLASS=""Lists"" "
										If Not IsNull(oRecordset.Fields("JavaScriptCode")) Then sField = sField & CStr(oRecordset.Fields("JavaScriptCode").Value)
									sField = sField & " >"
										If Len(sQueryForSource) > 0 Then
											asFields = Split(sQueryForSource, LIST_SEPARATOR, -1, vbBinaryCompare)
											sAnswer = ""
											sAnswer = CStr(oRecordset.Fields("DefaultValue").Value)
											Err.Clear
											If Len(sAnswer) = 0 Then sAnswer = asFields(6)
											sField = sField & GenerateListOptionsFromQuery(oADODBConnection, asFields(0), asFields(1), asFields(3), asFields(4), asFields(5), sAnswer, "", sErrorDescription)
										End If
									sField = sField & "</SELECT>"
									If Not bTemplate Then sField = sField & "<BR />"
								End If
							Case 7, 9 'Catálogo jerárquico, Lista jerárquica
								If lErrorNumber = 0 Then
									If Not bTemplate Then
										sField = CStr(oRecordset.Fields("FormFieldText").Value) & ": "
										If CLng(oRecordset.Fields("FieldTypeID").Value) = 9 Then sField = sField & "<BR />"
									End If
									sField = sField & "<SELECT NAME=""EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & """ ID=""EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Cmb"""
									If CLng(oRecordset.Fields("FieldTypeID").Value) = 7 Then
										sField = sField & " SIZE=""1"""
									Else
										sField = sField & " SIZE=""5"" MULTIPLE=""1"""
									End If
									sField = sField & " VALUE="""" CLASS=""Lists"" "
										If Not IsNull(oRecordset.Fields("JavaScriptCode")) Then sField = sField & CStr(oRecordset.Fields("JavaScriptCode").Value)
									sField = sField & " >"
										asFields = Split(sQueryForSource, LIST_SEPARATOR, -1, vbBinaryCompare)
										sAnswer = ""
										sAnswer = CStr(oRecordset.Fields("DefaultValue").Value)
										Err.Clear
										If Len(sAnswer) = 0 Then sAnswer = asFields(6)
										sField = sField & GenerateHierarchyListOptionsFromQuery(oFieldADODBConnection, asFields(0), asFields(1), asFields(2), asFields(3), asFields(4), -1, asFields(5), sAnswer, "", "", sField, sErrorDescription)
									sField = sField & "</SELECT>"
									If Not bTemplate Then sField = sField & "<BR />"
								End If
							Case 10 'Archivo
								sField = "<IFRAME SRC=""BrowserFile.asp?FormID=" & CStr(oRecordset.Fields("FormID").Value) & "&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&FormFieldID=" & CStr(oRecordset.Fields("FormFieldID").Value) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & """ NAME=""Form_" & aEmployeeComponent(N_ID_EMPLOYEE) & "_" & CStr(oRecordset.Fields("FormFieldID").Value) & "_FilesIFrame"" FRAMEBORDER=""1"" WIDTH=""400"" HEIGHT=""150""></IFRAME>"
							Case 11 'Oculto
								sField = "<INPUT TYPE=""HIDDEN"" NAME=""EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & """ ID=""EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Hdn"" VALUE="""
									If Len(sQueryForSource) > 0 Then
										sErrorDescription = "No se pudo obtener el valor por default para el campo de texto del formulario."
										lErrorNumber = ExecuteSQLQuery(oFieldADODBConnection, sQueryForSource, "EmployeeDisplayFormsComponentB.asp", S_FUNCTION_NAME, 0, sErrorDescription, oCatalogRecordset)
										If lErrorNumber = 0 Then
											If Not oCatalogRecordset.EOF Then
												For iIndex = 0 To oCatalogRecordset.Fields.Count - 1
													sField = sField & CStr(oCatalogRecordset.Fields(iIndex).Value)
													Err.Clear
													If iIndex < oCatalogRecordset.Fields.Count - 1 Then sField = sField & " "
												Next
											End If
										End If
									ElseIf Not IsNull(oRecordset.Fields("DefaultValue")) Then
										sField = sField & CStr(oRecordset.Fields("DefaultValue").Value)
									End If
								sField = sField & """ />"
						End Select
						If bTemplate Then
							sFormContents = Replace(sFormContents, "<" & CStr(oRecordset.Fields("FormFieldName").Value) & " />", sField)
						Else
							Response.Write sField
						End If
					End If
					oRecordset.MoveNext
					If Err.Number <> 0 Then Exit Do
				Loop
				oRecordset.Close
			Response.Write "</FORM>"
		End If
	Else
		sErrorDescription = "No se pudieron obtener los valores para el formulario del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesInformation.FormFieldID, FormFieldName, Answer, FieldTypeID, QueryForSource From EmployeesInformation, EmployeeFields Where (EmployeesInformation.FormFieldID=EmployeeFields.FormFieldID) And (EmployeesInformation.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") Order By EmployeesInformation.FormFieldID", "EmployeeDisplayFormsComponentB.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				sAnswer = CStr(oRecordset.Fields("Answer").Value)
				Select Case CInt(oRecordset.Fields("FieldTypeID").Value)
					Case 0 'Booleano
						sFormContents = Replace(sFormContents, "<" & CStr(oRecordset.Fields("FormFieldName").Value) & " />", DisplayYesNo(CInt(sAnswer), False))
					Case 1 'Fecha
						sFormContents = Replace(sFormContents, "<" & CStr(oRecordset.Fields("FormFieldName").Value) & " />", DisplayDateAndTimeFromSerialNumber(sAnswer, ""))
					Case 3 'Hora
						sFormContents = Replace(sFormContents, "<" & CStr(oRecordset.Fields("FormFieldName").Value) & " />", DisplayTimeFromSerialNumber(sAnswer & "00"))
					Case 6, 8 'Catálogo, Lista
						asFields = Split(CStr(oRecordset.Fields("QueryForSource").Value), LIST_SEPARATOR, -1, vbBinaryCompare)
						sErrorDescription = "No se pudieron obtener los valores para el formulario del empleado."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select " & asFields(3) & " From " & asFields(0) & " Where (" & asFields(1) & "=" & sAnswer & ")", "EmployeeDisplayFormsComponentB.asp", S_FUNCTION_NAME, 0, sErrorDescription, oCatalogRecordset)
						If lErrorNumber = 0 Then
							If Not oCatalogRecordset.EOF Then
								sFormContents = Replace(sFormContents, "<" & CStr(oRecordset.Fields("FormFieldName").Value) & " />", CStr(oCatalogRecordset.Fields(asFields(3)).Value))
							End If
						End If
					Case 7, 9 'Catálogo jerárquico, Lista jerárquica
						sFormContents = Replace(sFormContents, "<" & CStr(oRecordset.Fields("FormFieldName").Value) & " />", CleanStringForHTML(sAnswer))
					Case Else 'Flotante, Numérico, Texto
						sFormContents = Replace(sFormContents, "<" & CStr(oRecordset.Fields("FormFieldName").Value) & " />", CleanStringForHTML(sAnswer))
				End Select
				oRecordset.MoveNext
				If Err.Number <> 0 Then Exit Do
			Loop
			oRecordset.Close
		End If

		sErrorDescription = "No se pudieron obtener los valores para el formulario del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select FormFieldID, FormFieldName From EmployeeFields Where (FieldTypeID=10) Order By FormFieldID", "EmployeeDisplayFormsComponentB.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				sAnswer = ""
				Call GetFolderContents(Server.MapPath(INSPECTIONS_PHYSICAL_PATH & "i" & aEmployeeComponent(N_ID_EMPLOYEE) & "_" & CStr(oRecordset.Fields("FormFieldID").Value)), False, sAnswer, sErrorDescription)
				asFiles = Split(sAnswer, LIST_SEPARATOR)
				sAnswer = ""
				For iIndex = 0 To UBound(asFiles)
					If (InStr(1, LCase(asFiles(iIndex)), ".bmp") <> 0) Or (InStr(1, LCase(asFiles(iIndex)), ".gif") <> 0) Or (InStr(1, LCase(asFiles(iIndex)), ".jpg") <> 0) Or (InStr(1, LCase(asFiles(iIndex)), ".png") <> 0) Then
						sAnswer = sAnswer & "<IMG SRC=""" & S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME & INSPECTIONS_PATH & "i" & aEmployeeComponent(N_ID_EMPLOYEE) & "_" & CStr(oRecordset.Fields("FormFieldID").Value) & "/" & asFiles(iIndex) & """ /><BR />"
					Else
						sAnswer = sAnswer & "<A HREF=""" & S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME & INSPECTIONS_PATH & "i" & aEmployeeComponent(N_ID_EMPLOYEE) & "_" & CStr(oRecordset.Fields("FormFieldID").Value) & "/" & asFiles(iIndex) & """ TARGET=""_blank"">" & asFiles(iIndex) & "</A><BR />"
					End If
				Next
				sFormContents = Replace(sFormContents, "<" & CStr(oRecordset.Fields("FormFieldName").Value) & " />", sAnswer)
				oRecordset.MoveNext
				If Err.Number <> 0 Then Exit Do
			Loop
		End If
	End If

	Call TransformXMLTagsForEmployeeForm(aEmployeeComponent, True, sFormContents, sErrorDescription)
	Response.Write sFormContents

	If Not bForExport Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			sErrorDescription = "No se pudieron obtener la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesInformation.FormFieldID, Answer, FieldTypeID From EmployeesInformation, EmployeeFields Where (EmployeesInformation.FormFieldID=EmployeeFields.FormFieldID) And (EmployeesInformation.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") Order By EmployeesInformation.FormFieldID", "EmployeeDisplayFormsComponentB.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				sURL = ""
				Do While Not oRecordset.EOF
					sAnswer = CStr(oRecordset.Fields("Answer").Value)
					Select Case CInt(oRecordset.Fields("FieldTypeID").Value)
						Case 1
							sURL = sURL & "EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Year=" & Left(sAnswer, 4) & "&"
							sURL = sURL & "EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Month=" & Mid(sAnswer, 5, 2) & "&"
							sURL = sURL & "EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Day=" & Right(sAnswer, 2) & "&"
						Case 3
							sURL = sURL & "EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Hour=" & Left(sAnswer, 2) & "&"
							sURL = sURL & "EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Minute=" & Right(sAnswer, 2) & "&"
						Case 8, 9
							asValues = Split(sAnswer, ", ", -1, vbBinaryCompare)
							For iIndex = 0 To UBound(asValues)
								Response.Write "SelectItemByValue('" & asValues(iIndex) & "', true, document." & sFormName & ".EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ");" & vbNewLine
								If Err.Number <> 0 Then Exit For
							Next
						Case Else
							sURL = sURL & "EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & "=" & CleanStringForJavaScript(sAnswer) & "&"
					End Select
					oRecordset.MoveNext
					If Err.Number <> 0 Then Exit Do
				Loop
				oRecordset.Close
				If Len(sURL) > 0 Then
					sURL = Left(sURL, (Len(sURL) - Len("&")))
					Response.Write "SendURLValuesToForm('" & sURL & "', document.EmployeeFieldsFrm);" & vbNewLine
				End If
			End If

			Response.Write "function CheckFormForModule(oForm) {" & vbNewLine
				sErrorDescription = "No se pudieron obtener los campos obligatorios para el formulario."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select FormFieldID, FormFieldText, FieldTypeID, FormFieldSize, LimitTypeID, MinimumValue, MaximumValue, IsOptional From EmployeeFields Order By FormFieldID", "EmployeeDisplayFormsComponentB.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						Select Case CInt(oRecordset.Fields("LimitTypeID").Value)
							Case 0  'Ninguno
								sMinType = "N_NO_RANK_FLAG"
								sMaxType = "N_CLOSED_FLAG"
							Case 1 'Sólo mínimo abierto
								sMinType = "N_MINIMUM_ONLY_FLAG"
								sMaxType = "N_OPEN_FLAG"
							Case 2  'Sólo máximo abierto
								sMinType = "N_MAXIMUM_ONLY_FLAG"
								sMaxType = "N_OPEN_FLAG"
							Case 3  'Mínimo abierto y máximo abierto
								sMinType = "N_BOTH_FLAG"
								sMaxType = "N_OPEN_FLAG"
							Case 5  'Sólo mínimo cerrado
								sMinType = "N_MINIMUM_ONLY_FLAG"
								sMaxType = "N_CLOSED_FLAG"
							Case 7  'Mínimo cerrado y máximo abierto
								sMinType = "N_BOTH_FLAG"
								sMaxType = "N_MAXIMUM_OPEN_FLAG"
							Case 10 'Sólo máximo cerrado
								sMinType = "N_MAXIMUM_ONLY_FLAG"
								sMaxType = "N_CLOSED_FLAG"
							Case 11 'Mínimo abierto y máximo cerrado
								sMinType = "N_BOTH_FLAG"
								sMaxType = "N_MINIMUM_OPEN_FLAG"
							Case 15 'Mínimo cerrado y máximo cerrado
								sMinType = "N_BOTH_FLAG"
								sMaxType = "N_CLOSED_FLAG"
						End Select
						Response.Write vbTab & "if (oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ") {" & vbNewLine
							Select Case CInt(oRecordset.Fields("FieldTypeID").Value)
								Case 2 'Flotante
									If CInt(oRecordset.Fields("IsOptional").Value) = 1 Then Response.Write vbTab & "if (oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".value != '') {" & vbNewLine
										Response.Write vbTab & vbTab & "oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".value = oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
										Response.Write vbTab & "if (! CheckFloatValue(oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ", '" & Replace(Replace(Replace(CStr(oRecordset.Fields("FormFieldText").Value), "\", "\\"), "/", "\/"), "'", "\'") & "', " & sMinType & ", " & sMaxType & ", " & CStr(oRecordset.Fields("MinimumValue").Value) & ", " & CStr(oRecordset.Fields("MaximumValue").Value) & ")) {" & vbNewLine
											Response.Write vbTab & vbTab & "oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".focus();" & vbNewLine
											Response.Write vbTab & vbTab & "return false;" & vbNewLine
										Response.Write vbTab & "}" & vbNewLine
									If CInt(oRecordset.Fields("IsOptional").Value) = 1 Then Response.Write vbTab & "}" & vbNewLine
								Case 4 'Numérico
									If CInt(oRecordset.Fields("IsOptional").Value) = 1 Then Response.Write vbTab & "if (oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".value != '') {" & vbNewLine
										Response.Write vbTab & vbTab & "oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".value = oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
										Response.Write vbTab & "if (! CheckIntegerValue(oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ", '" & Replace(Replace(Replace(CStr(oRecordset.Fields("FormFieldText").Value), "\", "\\"), "/", "\/"), "'", "\'") & "', " & sMinType & ", " & sMaxType & ", " & CStr(oRecordset.Fields("MinimumValue").Value) & ", " & CStr(oRecordset.Fields("MaximumValue").Value) & ")) {" & vbNewLine
											Response.Write vbTab & vbTab & "oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".focus();" & vbNewLine
											Response.Write vbTab & vbTab & "return false;" & vbNewLine
										Response.Write vbTab & "}" & vbNewLine
									If CInt(oRecordset.Fields("IsOptional").Value) = 1 Then Response.Write vbTab & "}" & vbNewLine
								Case 5 'Texto
									If CInt(oRecordset.Fields("IsOptional").Value) = 0 Then
										Response.Write vbTab & "if (oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".value == '') {" & vbNewLine
											Response.Write vbTab & vbTab & "alert('Favor de introducir la información para el campo " & Replace(Replace(Replace(CStr(oRecordset.Fields("FormFieldText").Value), "\", "\\"), "/", "\/"), "'", "\'") & ".');" & vbNewLine
											Response.Write vbTab & vbTab & "oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".focus();" & vbNewLine
											Response.Write vbTab & vbTab & "return false;" & vbNewLine
										Response.Write vbTab & "}" & vbNewLine
										If CInt(oRecordset.Fields("MinimumValue").Value) > 0 Then
											Response.Write vbTab & vbTab & "if (oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".value.length < " & CStr(oRecordset.Fields("MinimumValue").Value) & ") {" & vbNewLine
												Response.Write vbTab & vbTab & vbTab & "ShowTaskTab(2);" & vbNewLine
												Response.Write vbTab & vbTab & vbTab & "alert('El campo " & Replace(Replace(Replace(CStr(oRecordset.Fields("FormFieldText").Value), "\", "\\"), "/", "\/"), "'", "\'") & " requiere al menos " & CStr(oRecordset.Fields("MinimumValue").Value) & " caracteres.');" & vbNewLine
												Response.Write vbTab & vbTab & vbTab & "window.setTimeout('oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".focus()', 1000);" & vbNewLine
												Response.Write vbTab & vbTab & vbTab & "return false;" & vbNewLine
											Response.Write vbTab & vbTab & "}" & vbNewLine
										End If
									End If
							End Select
						Response.Write vbTab & "}" & vbNewLine
						oRecordset.MoveNext
						If Err.Number <> 0 Then Exit Do
					Loop
				End If
				If InStr(1, sFormContents, "function CheckTemplate(", vbBinaryCompare) > 0 Then
					Response.Write vbTab & "return CheckTemplate(oForm);" & vbNewLine
				End If
				Response.Write vbTab & "return true;" & vbNewLine
			Response.Write "} // End of CheckFormForModule" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	Set oRecordset = Nothing
	DisplayFormForEmployee = lErrorNumber
	Err.Clear
End Function

Function DisplayFormForHonoraryEmployee(oADODBConnection, bForExport, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the details for the given employee form
'Inputs:  oRequest, oADODBConnection, bForExport, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayFormForHonoraryEmployee"
	Dim sFormContents
	Dim sField
	Dim sQuery
	Dim sQueryForSource
	Dim asFields
	Dim asValues
	Dim sURL
	Dim iIndex
	Dim sMinType
	Dim sMaxType
	Dim bTemplate
	Dim asIDs
	Dim oFieldADODBConnection
	Dim oRecordset
	Dim oCatalogRecordset
	Dim sAnswer
	Dim asFiles
	Dim aEmailComponent
	Dim lErrorNumber
	Dim sGender
	Dim lEmployeeAge
	Dim iContador
	Dim oHistoryRecordset
	Dim lHistoryStartDate
	Dim lHistoryEndDate

'	sFormContents = ""
'	If FileExists(Server.MapPath(TEMPLATES_PHYSICAL_PATH & "HonoraryEmployeeForm.htm"), sErrorDescription) Then
'		sFormContents = GetFileContents(Server.MapPath(TEMPLATES_PHYSICAL_PATH & "HonoraryEmployeeForm.htm"), sErrorDescription)
'	End If
'	bTemplate = (Len(sFormContents) > 0)
	
	
	sQuery = "Select EmployeeDate, EndDate, ReasonID From EmployeesHistoryList Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID In (14,66)) Order By EndDate Desc"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayFormsComponentB.asp", S_FUNCTION_NAME, 0, sErrorDescription, oHistoryRecordset)
	If lErrorNumber = 0 Then
		If Not oHistoryRecordset.EOF Then
			Do While Not oHistoryRecordset.EOF
			
	sFormContents = ""
	If FileExists(Server.MapPath(TEMPLATES_PHYSICAL_PATH & "HonoraryEmployeeForm.htm"), sErrorDescription) Then
		sFormContents = GetFileContents(Server.MapPath(TEMPLATES_PHYSICAL_PATH & "HonoraryEmployeeForm.htm"), sErrorDescription)
	End If
	bTemplate = (Len(sFormContents) > 0)
			
				If Not bForExport Then
					sErrorDescription = "No se pudo obtener el formulario para el módulo."
					If CInt(oHistoryRecordset.Fields("ReasonID").Value) = 14 Then
						sFormContents = Replace(sFormContents, "BAJA", "ALTA")
						sFormContents = Replace(sFormContents, "<REASON_NAME />", "ALTA")
						sFormContents = Replace(sFormContents, "<EMPLOYEE_START_YEAR />",mid(oHistoryRecordset.Fields("EmployeeDate").Value,1,4))
						sFormContents = Replace(sFormContents, "<EMPLOYEE_START_MONTH />",mid(oHistoryRecordset.Fields("EmployeeDate").Value,5,2))
						sFormContents = Replace(sFormContents, "<EMPLOYEE_START_DAY />",mid(oHistoryRecordset.Fields("EmployeeDate").Value,7,2))
						sFormContents = Replace(sFormContents, "<EMPLOYEE_END_YEAR />",mid(oHistoryRecordset.Fields("EndDate").Value,1,4))
						sFormContents = Replace(sFormContents, "<EMPLOYEE_END_MONTH />",mid(oHistoryRecordset.Fields("EndDate").Value,5,2))
						sFormContents = Replace(sFormContents, "<EMPLOYEE_END_DAY />",mid(oHistoryRecordset.Fields("EndDate").Value,7,2))
					ElseIf CInt(oHistoryRecordset.Fields("ReasonID").Value) = 66 Then
						sFormContents = Replace(sFormContents, "<REASON_NAME />", "BAJA")
						sFormContents = Replace(sFormContents, "ALTA", "BAJA")
						sFormContents = Replace(sFormContents, "<EMPLOYEE_DROP_YEAR />",mid(oHistoryRecordset.Fields("EmployeeDate").Value,1,4))
						sFormContents = Replace(sFormContents, "<EMPLOYEE_DROP_MONTH />",mid(oHistoryRecordset.Fields("EmployeeDate").Value,5,2))
						sFormContents = Replace(sFormContents, "<EMPLOYEE_DROP_DAY />",mid(oHistoryRecordset.Fields("EmployeeDate").Value,7,2))
						sFormContents = Replace(sFormContents, "<EMPLOYEE_START_YEAR />","")
						sFormContents = Replace(sFormContents, "<EMPLOYEE_START_MONTH />","")
						sFormContents = Replace(sFormContents, "<EMPLOYEE_START_DAY />","")
						sFormContents = Replace(sFormContents, "<EMPLOYEE_END_YEAR />","")
						sFormContents = Replace(sFormContents, "<EMPLOYEE_END_MONTH />","")
						sFormContents = Replace(sFormContents, "<EMPLOYEE_END_DAY />","")
					End If
					lHistoryStartDate = oHistoryRecordset.Fields("EmployeeDate").Value
					lHistoryEndDate = oHistoryRecordset.Fields("EndDate").Value

					sQuery = "Select EmployeeNumber, EmployeeLastName, EmployeeLastName2, EmployeeName,  RFC, CURP, GenderID, MaritalStatusName, BirthDate," & _
								" jobUA.area1, jobUA.area2, jobUA.Clave, jobUA.UA" & _
							" From Employees, MaritalStatus, Jobs, Areas," & _
								" (Select JobID, jobs.AreaID, areas.AreaName as area1, jobs.ZoneID, ZoneName as area2, AreaCode as Clave, UA.AreaName as UA" & _
								 " From jobs, areas, zones, (select areaname, areaid from areas) as UA" & _
								 " Where JobID = " & aEmployeeComponent(N_ID_EMPLOYEE) & _
									" And Jobs.AreaID = Areas.AreaID" & _
									" And Jobs.ZoneID = Areas.ZoneID" & _
									" And Jobs.ZoneID = Zones.ZoneID" & _
									" And Areas.ParentID = UA.AreaID) As jobUA" & _
							" Where EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & _
								" And Employees.MaritalStatusID = MaritalStatus.MaritalStatusID" & _
								" And Jobs.JobID = Employees.JobID" & _
								" And Jobs.AreaID = Areas.AreaID"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayFormsComponentB.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					If Not oRecordset.EOF Then
						sFormContents = Replace(sFormContents, "<PARENT_AREA_NAME />", oRecordset.Fields("UA").Value)
						sFormContents = Replace(sFormContents, "<AREA_NAME />", oRecordset.Fields("area1").Value & " " & oRecordset.Fields("area2").Value)
						sFormContents = Replace(sFormContents, "<AREA_CODE />", oRecordset.Fields("Clave").Value)
						sFormContents = Replace(sFormContents, "<EMPLOYEE_LAST_NAME />", oRecordset.Fields("EmployeeLastName").Value)
						If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
							sFormContents = Replace(sFormContents, "<EMPLOYEE_LAST_NAME2 />", oRecordset.Fields("EmployeeLastName2").Value)
						Else
							sFormContents = Replace(sFormContents, "<EMPLOYEE_LAST_NAME2 />", " ")
						End If
						sFormContents = Replace(sFormContents, "<EMPLOYEE_NAME />", oRecordset.Fields("EmployeeName").Value)
						sFormContents = Replace(sFormContents, "<EMPLOYEE_RFC />", oRecordset.Fields("RFC").Value)
						sFormContents = Replace(sFormContents, "<EMPLOYEE_CURP />", oRecordset.Fields("CURP").Value)
						If oRecordset.Fields("GenderID") = 0 Then
							sGender = "M"
						Else
							sGender = "H"
						End If
						sFormContents = Replace(sFormContents, "<GENDER_SHORT_NAME />", sGender)
						sFormContents = Replace(sFormContents, "<MARITAL_STATUS />", oRecordset.Fields("MaritalStatusName").Value)
						sFormContents = Replace(sFormContents, "<EMPLOYEE_AGE />", CalculateEmployeeAge(oADODBConnection, aEmployeeComponent, lEmployeeAge, sErrorDescription))
					End If
					
					sQuery = "Select EmployeeActivityID, EmployeeAddress, EmployeeCity, EmployeeZipCode, StateName , CountryName," & _
							" DocumentNumber1, DocumentNumber2, DocumentNumber3, Nationality" & _
							" From EmployeesExtraInfo, countries, states" & _ 
							" Where EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE)& _
								" And EmployeesExtraInfo.CountryID = Countries.CountryID" & _
								" And EmployeesExtraInfo.StateID = States.StateID"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayFormsComponentB.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					If Not oRecordset.EOF Then
						sFormContents = Replace(sFormContents, "<ADDRESS_NAME />", oRecordset.Fields("EmployeeAddress").Value)
						sFormContents = Replace(sFormContents, "<ADDRESS_CITY />", oRecordset.Fields("EmployeeCity").Value)
						sFormContents = Replace(sFormContents, "<STATE_NAME />", oRecordset.Fields("StateName").Value)
						sFormContents = Replace(sFormContents, "<ADDRESS_ZIP_CODE />", oRecordset.Fields("EmployeeZipCode").Value)
						sFormContents = Replace(sFormContents, "<ADDRESS_NAME />", oRecordset.Fields("EmployeeAddress").Value)
						sFormContents = Replace(sFormContents, "<NATIONALITY />", oRecordset.Fields("Nationality").Value)
						sFormContents = Replace(sFormContents, "<DOCUMENT_1 />", oRecordset.Fields("DocumentNumber1").Value)
						sFormContents = Replace(sFormContents, "<DOCUMENT_2 />", oRecordset.Fields("DocumentNumber2").Value)
						sFormContents = Replace(sFormContents, "<DOCUMENT_3 />", oRecordset.Fields("DocumentNumber3").Value)
						If CInt(oRecordset.Fields("EmployeeActivityID").Value) = 1 Then
							sFormContents = Replace(sFormContents, "<MEDICAL_ACTIVITIES />", "X")
						ElseIf CInt(oRecordset.Fields("EmployeeActivityID").Value) = 2 Then
							sFormContents = Replace(sFormContents, "<TECHNICAL_ACTIVITIES />", "X")
						ElseIf CInt(oRecordset.Fields("EmployeeActivityID").Value) = 3 Then
							sFormContents = Replace(sFormContents, "<ADMINISTRATIVE_ACTIVITIES />", "X")
						End If
					End If
					sQuery = "Select ConceptAmount*2 as Salario From EmployeesConceptsLKP where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (StartDate = " & lHistoryStartDate & ")"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayFormsComponentB.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					If Not oRecordset.EOF Then
						sFormContents = Replace(sFormContents, "<CONCEPT_AMOUNT />", FormatNumber(oRecordset.Fields("Salario").Value, 2, True, False, True))
					End If
				End If

				Call TransformXMLTagsForEmployeeForm(aEmployeeComponent, True, sFormContents, sErrorDescription)
				Response.Write sFormContents

				If Not bForExport Then
					Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
						sErrorDescription = "No se pudieron obtener la información del empleado."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesInformation.FormFieldID, Answer, FieldTypeID From EmployeesInformation, EmployeeFields Where (EmployeesInformation.FormFieldID=EmployeeFields.FormFieldID) And (EmployeesInformation.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") Order By EmployeesInformation.FormFieldID", "EmployeeDisplayFormsComponentB.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							sURL = ""
							Do While Not oRecordset.EOF
								sAnswer = CStr(oRecordset.Fields("Answer").Value)
								Select Case CInt(oRecordset.Fields("FieldTypeID").Value)
									Case 1
										sURL = sURL & "EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Year=" & Left(sAnswer, 4) & "&"
										sURL = sURL & "EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Month=" & Mid(sAnswer, 5, 2) & "&"
										sURL = sURL & "EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Day=" & Right(sAnswer, 2) & "&"
									Case 3
										sURL = sURL & "EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Hour=" & Left(sAnswer, 2) & "&"
										sURL = sURL & "EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Minute=" & Right(sAnswer, 2) & "&"
									Case 8, 9
										asValues = Split(sAnswer, ", ", -1, vbBinaryCompare)
										For iIndex = 0 To UBound(asValues)
											Response.Write "SelectItemByValue('" & asValues(iIndex) & "', true, document." & sFormName & ".EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ");" & vbNewLine
											If Err.Number <> 0 Then Exit For
										Next
									Case Else
										sURL = sURL & "EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & "=" & CleanStringForJavaScript(sAnswer) & "&"
								End Select
								oRecordset.MoveNext
								If Err.Number <> 0 Then Exit Do
							Loop
							oRecordset.Close
							If Len(sURL) > 0 Then
								sURL = Left(sURL, (Len(sURL) - Len("&")))
								Response.Write "SendURLValuesToForm('" & sURL & "', document.EmployeeFieldsFrm);" & vbNewLine
							End If
						End If

						Response.Write "function CheckFormForModule(oForm) {" & vbNewLine
							sErrorDescription = "No se pudieron obtener los campos obligatorios para el formulario."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select FormFieldID, FormFieldText, FieldTypeID, FormFieldSize, LimitTypeID, MinimumValue, MaximumValue, IsOptional From EmployeeFields Order By FormFieldID", "EmployeeDisplayFormsComponentB.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								Do While Not oRecordset.EOF
									Select Case CInt(oRecordset.Fields("LimitTypeID").Value)
										Case 0  'Ninguno
											sMinType = "N_NO_RANK_FLAG"
											sMaxType = "N_CLOSED_FLAG"
										Case 1 'Sólo mínimo abierto
											sMinType = "N_MINIMUM_ONLY_FLAG"
											sMaxType = "N_OPEN_FLAG"
										Case 2  'Sólo máximo abierto
											sMinType = "N_MAXIMUM_ONLY_FLAG"
											sMaxType = "N_OPEN_FLAG"
										Case 3  'Mínimo abierto y máximo abierto
											sMinType = "N_BOTH_FLAG"
											sMaxType = "N_OPEN_FLAG"
										Case 5  'Sólo mínimo cerrado
											sMinType = "N_MINIMUM_ONLY_FLAG"
											sMaxType = "N_CLOSED_FLAG"
										Case 7  'Mínimo cerrado y máximo abierto
											sMinType = "N_BOTH_FLAG"
											sMaxType = "N_MAXIMUM_OPEN_FLAG"
										Case 10 'Sólo máximo cerrado
											sMinType = "N_MAXIMUM_ONLY_FLAG"
											sMaxType = "N_CLOSED_FLAG"
										Case 11 'Mínimo abierto y máximo cerrado
											sMinType = "N_BOTH_FLAG"
											sMaxType = "N_MINIMUM_OPEN_FLAG"
										Case 15 'Mínimo cerrado y máximo cerrado
											sMinType = "N_BOTH_FLAG"
											sMaxType = "N_CLOSED_FLAG"
									End Select
									Response.Write vbTab & "if (oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ") {" & vbNewLine
										Select Case CInt(oRecordset.Fields("FieldTypeID").Value)
											Case 2 'Flotante
												If CInt(oRecordset.Fields("IsOptional").Value) = 1 Then Response.Write vbTab & "if (oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".value != '') {" & vbNewLine
													Response.Write vbTab & vbTab & "oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".value = oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
													Response.Write vbTab & "if (! CheckFloatValue(oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ", '" & Replace(Replace(Replace(CStr(oRecordset.Fields("FormFieldText").Value), "\", "\\"), "/", "\/"), "'", "\'") & "', " & sMinType & ", " & sMaxType & ", " & CStr(oRecordset.Fields("MinimumValue").Value) & ", " & CStr(oRecordset.Fields("MaximumValue").Value) & ")) {" & vbNewLine
														Response.Write vbTab & vbTab & "oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".focus();" & vbNewLine
														Response.Write vbTab & vbTab & "return false;" & vbNewLine
													Response.Write vbTab & "}" & vbNewLine
												If CInt(oRecordset.Fields("IsOptional").Value) = 1 Then Response.Write vbTab & "}" & vbNewLine
											Case 4 'Numérico
												If CInt(oRecordset.Fields("IsOptional").Value) = 1 Then Response.Write vbTab & "if (oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".value != '') {" & vbNewLine
													Response.Write vbTab & vbTab & "oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".value = oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
													Response.Write vbTab & "if (! CheckIntegerValue(oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ", '" & Replace(Replace(Replace(CStr(oRecordset.Fields("FormFieldText").Value), "\", "\\"), "/", "\/"), "'", "\'") & "', " & sMinType & ", " & sMaxType & ", " & CStr(oRecordset.Fields("MinimumValue").Value) & ", " & CStr(oRecordset.Fields("MaximumValue").Value) & ")) {" & vbNewLine
														Response.Write vbTab & vbTab & "oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".focus();" & vbNewLine
														Response.Write vbTab & vbTab & "return false;" & vbNewLine
													Response.Write vbTab & "}" & vbNewLine
												If CInt(oRecordset.Fields("IsOptional").Value) = 1 Then Response.Write vbTab & "}" & vbNewLine
											Case 5 'Texto
												If CInt(oRecordset.Fields("IsOptional").Value) = 0 Then
													Response.Write vbTab & "if (oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".value == '') {" & vbNewLine
														Response.Write vbTab & vbTab & "alert('Favor de introducir la información para el campo " & Replace(Replace(Replace(CStr(oRecordset.Fields("FormFieldText").Value), "\", "\\"), "/", "\/"), "'", "\'") & ".');" & vbNewLine
														Response.Write vbTab & vbTab & "oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".focus();" & vbNewLine
														Response.Write vbTab & vbTab & "return false;" & vbNewLine
													Response.Write vbTab & "}" & vbNewLine
													If CInt(oRecordset.Fields("MinimumValue").Value) > 0 Then
														Response.Write vbTab & vbTab & "if (oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".value.length < " & CStr(oRecordset.Fields("MinimumValue").Value) & ") {" & vbNewLine
															Response.Write vbTab & vbTab & vbTab & "ShowTaskTab(2);" & vbNewLine
															Response.Write vbTab & vbTab & vbTab & "alert('El campo " & Replace(Replace(Replace(CStr(oRecordset.Fields("FormFieldText").Value), "\", "\\"), "/", "\/"), "'", "\'") & " requiere al menos " & CStr(oRecordset.Fields("MinimumValue").Value) & " caracteres.');" & vbNewLine
															Response.Write vbTab & vbTab & vbTab & "window.setTimeout('oForm.EF__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".focus()', 1000);" & vbNewLine
															Response.Write vbTab & vbTab & vbTab & "return false;" & vbNewLine
														Response.Write vbTab & vbTab & "}" & vbNewLine
													End If
												End If
										End Select
									Response.Write vbTab & "}" & vbNewLine
									oRecordset.MoveNext
									If Err.Number <> 0 Then Exit Do
								Loop
							End If
							If InStr(1, sFormContents, "function CheckTemplate(", vbBinaryCompare) > 0 Then
								Response.Write vbTab & "return CheckTemplate(oForm);" & vbNewLine
							End If
							Response.Write vbTab & "return true;" & vbNewLine
						Response.Write "} // End of CheckFormForModule" & vbNewLine
					Response.Write "//--></SCRIPT>" & vbNewLine
				End If
				oHistoryRecordset.MoveNext
			Loop
		Else
			sErrorDescription = "No se encontró el hisotrial del empleado."
			lErrorNumber = -1
		End If
	Else
		sErrorDescription = "No se pudo obtener la información del empleado."		
	End If
	
	Set oHistoryRecordset = Nothing
	Set oRecordset = Nothing
	DisplayFormForHonoraryEmployee = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesCreditsSearchForm(oRequest, oADODBConnection, sAction, bFull, sErrorDescription)
'************************************************************
'Purpose: To display the search HTML form
'Inputs:  oRequest, oADODBConnection, bFull
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesCreditsSearchForm"

	B_UPPERCASE = false
	Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""" & sAction & """ METHOD=""GET"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""ThirdUploadMovements"" />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ActionHdn"" VALUE=""" & oRequest("ReasonID").Item & """ />"
		If bFull Then Response.Write "<B>CONSULTA DE CARGAS DE CREDITOS DE TERCEROS</B><BR /><BR />"
		Response.Write "<TABLE"
			If Not bFull Then Response.Write " WIDTH=""400"""
		Response.Write " BORDER=""0"" CELLPADING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
			Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Archivo de carga de terceros:<BR /><BR /><BR /></FONT>"
			Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""ConceptFileName"" ID=""ConceptFileNameLst"" SIZE=""1"" CLASS=""Lists"">"
				Response.Write "<OPTION VALUE="""">Todos</OPTION>"
				Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Credits", "Distinct UploadedFileName", "UploadedFileName As UploadedFileName2", "UploadedFileName IS NOT NULL AND UploadedFileName <> ' '", "UploadedFileName", aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE), "", sErrorDescription)
			Response.Write "</SELECT><BR /><BR />"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""2"""
				If Not bFull Then Response.Write " ALIGN=""LEFT"""
				Response.Write "><INPUT TYPE=""SUBMIT"" NAME=""DoSearch"" ID=""DoSearchBtn"" VALUE=""Buscar Registros"" CLASS=""Buttons"" /></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	Response.Write "</FORM>"

	DisplayEmployeesCreditsSearchForm = Err.number
End Function
%>