<%
Function AddAlimonyTypes(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new Alimony Type into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddAlimonyTypes"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	lErrorNumber = GetNewIDFromTable(oADODBConnection, "AlimonyTypes", "AlimonyTypeID", "", 1, aEmployeeComponent(N_ALIMONY_TYPE_ID_BENEFICIARY_EMPLOYEE), sErrorDescription)
	If lErrorNumber = 0 Then
		sQuery = "INSERT INTO AlimonyTypes" & _
				 " (AlimonyTypeID, AlimonyTypeShortName, AlimonyTypeName, ConceptQttyID," & _
				 " AppliesToID, Active) Values" & _
				 " (" & aEmployeeComponent(N_ALIMONY_TYPE_ID_BENEFICIARY_EMPLOYEE) & _
				 ",	'" & aEmployeeComponent(S_ALIMONY_TYPE_SHORT_NAME) & "'" & _
				 ",	'" & aEmployeeComponent(S_ALIMONY_TYPE_NAME) & "'" & _
				 "," & aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) & _
				 ",'" & aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) & "'" & _
				 ", " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ")" 

		sErrorDescription = "No se pudo agregar la información de la cuenta bancaria del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AlimonyTypeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	Else
		sErrorDescription = "No se pudo obtener un identificador para el nuevo tipo de pensión alimenticia."
		lErrorNumber = -1
	End If

	Set oRecordset = Nothing
	AddAlimonyTypes = lErrorNumber
	Err.Clear
End Function

Function AddCreditorTypes(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new Alimony Type into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddCreditorTypes"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	lErrorNumber = GetNewIDFromTable(oADODBConnection, "CreditorsTypes", "CreditorTypeID", "", 1, aEmployeeComponent(N_CREDITOR_TYPE_ID_EMPLOYEE), sErrorDescription)
	If lErrorNumber = 0 Then
		sQuery = "INSERT INTO CreditorsTypes" & _
				 " (CreditorTypeID, CreditorTypeShortName, CreditorTypeName, ConceptQttyID," & _
				 " AppliesToID, Active) Values" & _
				 " (" & aEmployeeComponent(N_CREDITOR_TYPE_ID_EMPLOYEE) & _
				 ",	'" & aEmployeeComponent(S_CREDITOR_TYPE_SHORT_NAME) & "'" & _
				 ",	'" & aEmployeeComponent(S_CREDITOR_TYPE_NAME) & "'" & _
				 "," & aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) & _
				 ",'" & aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) & "'" & _
				 ", " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ")" 

		sErrorDescription = "No se pudo agregar la información de la cuenta bancaria del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AlimonyTypeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	Else
		sErrorDescription = "No se pudo obtener un identificador para el nuevo tipo de pensión alimenticia."
		lErrorNumber = -1
	End If

	Set oRecordset = Nothing
	AddCreditorTypes = lErrorNumber
	Err.Clear
End Function

Function DisplayAlimonyTypesForm(oRequest, oADODBConnection, sAction, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the form from register Alimony Types
'Inputs:  oRequest, oADODBConnection, sAction, bFull, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayAlimonyTypesForm"

	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function ShowAmountFields(sValue, sFieldsName) {" & vbNewLine
				Response.Write "var oForm = document.AlimonyTypesFrm;" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					Response.Write "if(document.all[sFieldsName + 'CurrencySpn'] != null) HideDisplay(document.all[sFieldsName + 'CurrencySpn']);" & vbNewLine
					Response.Write "if(document.all[sFieldsName + 'AppliesToSpn'] != null) HideDisplay(document.all[sFieldsName + 'AppliesToSpn']);" & vbNewLine
					Response.Write "switch (sValue) {" & vbNewLine
						Response.Write "case 1:" & vbNewLine
							Response.Write "ShowDisplay(document.all[sFieldsName + 'CurrencySpn']);" & vbNewLine
							Response.Write "break;" & vbNewLine
						Response.Write "case '2':" & vbNewLine
							Response.Write "ShowDisplay(document.all[sFieldsName + 'AppliesToSpn']);" & vbNewLine
							Response.Write "break;" & vbNewLine
						Response.Write "case '8':" & vbNewLine
							Response.Write "ShowDisplay(document.all[sFieldsName + 'AppliesToSpn']);" & vbNewLine
							Response.Write "break;" & vbNewLine
						Response.Write "case '11':" & vbNewLine
							Response.Write "ShowDisplay(document.all[sFieldsName + 'AppliesToSpn']);" & vbNewLine
							Response.Write "break;" & vbNewLine
						Response.Write "default:" & vbNewLine
							Response.Write "break;" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of ShowAmountFields" & vbNewLine
			Response.Write "function CheckAlimonyTypesFields(oForm) {" & vbNewLine
				Response.Write "var oForm = document.AlimonyTypesFrm;" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					If B_ISSSTE Then
						Select Case lReasonID
							Case ALIMONY_TYPES
								Response.Write "if (oForm.AlimonyTypeShortName.value.length == 0) {" & vbNewLine
									Response.Write "alert('Favor de introducir la clave del tipo de pensión alimenticia.');" & vbNewLine
									Response.Write "oForm.AlimonyTypeName.focus();" & vbNewLine
									Response.Write "return false;" & vbNewLine
								Response.Write "}" & vbNewLine
								Response.Write "if (oForm.AlimonyTypeName.value.length == 0) {" & vbNewLine
									Response.Write "alert('Favor de introducir el nombre del tipo de pensión alimenticia.');" & vbNewLine
									Response.Write "oForm.AlimonyTypeName.focus();" & vbNewLine
									Response.Write "return false;" & vbNewLine
								Response.Write "}" & vbNewLine
							Case CREDITORS_TYPES
								Response.Write "if (oForm.CreditorTypeShortName.value.length == 0) {" & vbNewLine
									Response.Write "alert('Favor de introducir la clave del tipo de pensión alimenticia.');" & vbNewLine
									Response.Write "oForm.CreditorTypeName.focus();" & vbNewLine
									Response.Write "return false;" & vbNewLine
								Response.Write "}" & vbNewLine
								Response.Write "if (oForm.CreditorTypeName.value.length == 0) {" & vbNewLine
									Response.Write "alert('Favor de introducir el nombre del tipo de pensión alimenticia.');" & vbNewLine
									Response.Write "oForm.CreditorTypeName.focus();" & vbNewLine
									Response.Write "return false;" & vbNewLine
								Response.Write "}" & vbNewLine
						End Select
						Response.Write "if (oForm.AppliesToID.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de selecciónar el(los) concepto(s) a los que se aplica la pensión.');" & vbNewLine
							Response.Write "oForm.AlimonyTypeName.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					End If
					Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckAbsenceFields" & vbNewLine

			Response.Write "function ShowHideAbsencesFields(sValue) {" & vbNewLine
				Response.Write "var oForm = document.AlimonyTypesFrm" & vbNewLine
				If Not B_ISSSTE Then
					Response.Write "if (oForm) {" & vbNewLine
						Response.Write "if (sValue == 0) {" & vbNewLine
							Response.Write "HideDisplay(document.all['AbsenceHoursDiv']);" & vbNewLine
						Response.Write "} else {" & vbNewLine
							Response.Write "ShowDisplay(document.all['AbsenceHoursDiv']);" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "}" & vbNewLine
				End If
			Response.Write "} // End of ShowHideAbsencesFields" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine

		Response.Write "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
			Response.Write "<TD WIDTH=""30%"" VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"
				Response.Write "<DIV NAME=""EntriesDiv"" ID=""EntriesDiv"" CLASS=""TableScrollDiv"">"
					Select Case lReasonID
						Case ALIMONY_TYPES
							lErrorNumber = DisplayAlimonyTypesTable(oRequest, oADODBConnection, False, sErrorDescription)
						Case CREDITORS_TYPES
							lErrorNumber = DisplayCreditorsTypesTable(oRequest, oADODBConnection, False, sErrorDescription)
					End Select
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
				Select Case lReasonID
					Case ALIMONY_TYPES
						If CLng(aEmployeeComponent(N_ALIMONY_TYPE_ID_BENEFICIARY_EMPLOYEE)) <> -1 Then
							lErrorNumber = GetAlimonyType(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						End If
					Case CREDITORS_TYPES
						If CLng(aEmployeeComponent(N_CREDITOR_TYPE_ID_EMPLOYEE)) <> -1 Then
							lErrorNumber = GetCreditorType(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						End If
				End Select
					Response.Write "<DIV NAME=""CatalogDiv"" ID=""CatalogDiv"">"
						Response.Write "<FORM NAME=""AlimonyTypesFrm"" ID=""AlimonyTypesFrm"" ACTION=""" & sAction & """ METHOD=""GET"" onSubmit=""return CheckAlimonyTypesFields(this)"">"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""AlimonyTypes"" />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SaveAlimonyTypesMovements"" ID=""ActionHdn"" VALUE=""1"" />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE=""" & lReasonID & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AlimonyTypeID"" ID=""AlimonyTypeIDHdn"" VALUE=""" & aEmployeeComponent(N_ALIMONY_TYPE_ID_BENEFICIARY_EMPLOYEE) & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CreditorTypeID"" ID=""CreditorTypeIDHdn"" VALUE=""" & aEmployeeComponent(N_CREDITOR_TYPE_ID_EMPLOYEE) & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Active"" ID=""ActiveHdn"" VALUE=""" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & """ />"
							Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
								'Response.Write "<TD WIDTH=""600"" VALIGN=""TOP"">"
								Select Case lReasonID
									Case ALIMONY_TYPES
										Response.Write "<TR>"
											Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Clave del tipo de pensión alimenticia:&nbsp;</NOBR></FONT></TD>"
											Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""AlimonyTypeShortName"" ID=""AlimonyTypeShortNameTxt"" VALUE=""" & aEmployeeComponent(S_ALIMONY_TYPE_SHORT_NAME) & """ SIZE=""10"" MAXLENGTH=""5"" CLASS=""TextFields"" /></TD>"
										Response.Write "</TR>"
										Response.Write "<TR>"
											Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Nombre del tipo de pensión alimenticia:&nbsp;</NOBR></FONT></TD>"
											Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""AlimonyTypeName"" ID=""AlimonyTypeNameTxt"" VALUE=""" & aEmployeeComponent(S_ALIMONY_TYPE_NAME) & """ SIZE=""50"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
										Response.Write "</TR>"
										Response.Write "<TR>"
											Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Unidad para aplicar la pensión:&nbsp;</FONT></TD>"
											Response.Write "<TD>"
													Response.Write "<SELECT NAME=""ConceptQttyID"" ID=""ConceptQttyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
														Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "(QttyID In (1,2))", "QttyID", aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
													Response.Write "</SELECT>"
											Response.Write "</FONT></TD>"
										Response.Write "</TR>"
									Case CREDITORS_TYPES
										Response.Write "<TR>"
											Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Clave del tipo de descuento para acreedores:&nbsp;</NOBR></FONT></TD>"
											Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""CreditorTypeShortName"" ID=""CreditorTypeShortNameTxt"" VALUE=""" & aEmployeeComponent(S_CREDITOR_TYPE_SHORT_NAME) & """ SIZE=""10"" MAXLENGTH=""5"" CLASS=""TextFields"" /></TD>"
										Response.Write "</TR>"
										Response.Write "<TR>"
											Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Nombre del tipo de descuento para acreedores:&nbsp;</NOBR></FONT></TD>"
											Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""CreditorTypeName"" ID=""CreditorTypeNameTxt"" VALUE=""" & aEmployeeComponent(S_CREDITOR_TYPE_NAME) & """ SIZE=""50"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
										Response.Write "</TR>"
										Response.Write "<TR>"
											Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Unidad para aplicar el tipo de descuento para acreedores:&nbsp;</FONT></TD>"
											Response.Write "<TD>"
													Response.Write "<SELECT NAME=""ConceptQttyID"" ID=""ConceptQttyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
														Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "(QttyID In (1,2))", "QttyID", aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
													Response.Write "</SELECT>"
											Response.Write "</FONT></TD>"
										Response.Write "</TR>"
								End Select
								Response.Write "<TR>"
									Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><NOBR>Conceptos sobre los que se aplica:&nbsp;</NOBR></FONT></TD>"
											Response.Write "<TD>"
											Response.Write "<SELECT NAME=""AppliesToID"" ID=""AppliesToIDCmb"" SIZE=""10"" MULTIPLE=""1"" CLASS=""Lists"">"
												Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "", "ConceptShortName, ConceptName", aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
											Response.Write "</SELECT>"
											Response.Write "<BR />"
									Response.Write "</FONT></TD>"
								Response.Write "</TR>"
								If False Then
									Response.Write "<TR>"
										Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><NOBR>Monto mínimo:&nbsp;</NOBR></FONT></TD>"
										Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
											Response.Write "<INPUT TYPE=""TEXT"" NAME=""ConceptMin"" ID=""ConceptMinTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & FormatNumber(aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE), 2, True, False, True) & """ CLASS=""TextFields"" />&nbsp;"
											Response.Write "<SELECT NAME=""ConceptMinQttyID"" ID=""ConceptMinQttyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
												Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "(QttyID In (1,3,13,23,24,25,33,34,35))", "QttyID", aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
											Response.Write "</SELECT>"
										Response.Write "</FONT></TD>"
									Response.Write "</TR>"
									Response.Write "<TR>"
										Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><NOBR>Monto máximo:&nbsp;</NOBR></FONT></TD>"
										Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
											Response.Write "<INPUT TYPE=""TEXT"" NAME=""ConceptMax"" ID=""ConceptMaxTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & FormatNumber(aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE), 2, True, False, True) & """ CLASS=""TextFields"" />&nbsp;"
											Response.Write "<SELECT NAME=""ConceptMaxQttyID"" ID=""ConceptMaxQttyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
												Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "(QttyID In (1,3,13,23,24,25,33,34,35))", "QttyID", aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
											Response.Write "</SELECT>"
										Response.Write "</FONT></TD>"
									Response.Write "</TR>"
								End If
							Response.Write "</TABLE>"
							Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""340"" HEIGHT=""1"" /><BR /><BR />"
							Select Case lReasonID
								Case ALIMONY_TYPES
									If aEmployeeComponent(N_ALIMONY_TYPE_ID_BENEFICIARY_EMPLOYEE) = -1 Then
										If InStr(1, sAction, "Employees") = 0 Then
											If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
										End If
									ElseIf Len(oRequest("Delete").Item) > 0 Then
										If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveAlimonyTypesWngDiv']); AlimonyTypesFrm.Remove.focus()"" />"
									Else
										If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
									End If
									Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
									If InStr(1, sAction, "Employees") = 0 Then
										Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=AlimonyTypes&ReasonID=" & lReasonID & "'"" />"
									End If
								Case CREDITORS_TYPES
									If aEmployeeComponent(N_CREDITOR_TYPE_ID_EMPLOYEE) = -1 Then
										If InStr(1, sAction, "Employees") = 0 Then
											If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
										End If
									ElseIf Len(oRequest("Delete").Item) > 0 Then
										If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveAlimonyTypesWngDiv']); AlimonyTypesFrm.Remove.focus()"" />"
									Else
										If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
									End If
									Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
									If InStr(1, sAction, "Employees") = 0 Then
										Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=AlimonyTypes&ReasonID=" & lReasonID & "'"" />"
									End If
							End Select
							Response.Write "<BR /><BR />"
							Call DisplayWarningDiv("RemoveAlimonyTypesWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
						Response.Write "</FORM>"
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
	DisplayAlimonyTypesForm = lErrorNumber
	Err.Clear
End Function

Function DisplayAlimonyTypesTable(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the AlimonyTypes list from the database
'Inputs:  oRequest, oADODBConnection, bForExport, aAbsenceComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayAlimonyTypesTable"
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sFontBegin
	Dim sFontEnd
	Dim sBoldBegin
	Dim sBoldEnd
	Dim sNames
	Dim lErrorNumber
	Dim sConceptNames

	lErrorNumber = GetAlimonyTypes(oRequest, oADODBConnection, aEmployeeComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><TABLE BORDER="""
			If bForExport Then
				Response.Write "1"
			Else
				Response.Write "0"
			End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			If Not bForExport And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
				asColumnsTitles = Split("Acciones,Clave,Nombre del tipo de pensión,Unidad,Conceptos sobre los que se aplica", ",", -1, vbBinaryCompare)
				asCellWidths = Split(",,,,,,1500",",", -1, vbBinaryCompare)
				asCellAlignments = Split("CENTER,,,,,", ",", -1, vbBinaryCompare)
			Else
				asColumnsTitles = Split("Clave, Nombre del tipo de pensión,Tipo de unidad,Conceptos sobre los cuales se aplica", ",", -1, vbBinaryCompare)
				asCellWidths = Split(",,,,,,",",", -1, vbBinaryCompare)
				asCellAlignments = Split("CENTER,,,,,", ",", -1, vbBinaryCompare)
			End If

			If bForExport Then
				lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
			Else
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If
			End If
			sBoldBegin = "<B>"
			sBoldEnd = "</B>"
			sFontBegin = ""
			sFontEnd = ""
			Do While Not oRecordset.EOF
				sConceptNames = ""
				If Not bForExport Then
					If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
						sRowContents = "<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&Delete=1&AlimonyTypeID=" & CStr(oRecordset.Fields("AlimonyTypeID").Value) & "&ReasonID=" & lReasonID & """>"
							sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar registro"" BORDER=""0"" />"
						sRowContents = sRowContents & "</A>&nbsp;"
					Else
							sRowContents = "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""8"" BORDER=""0"" />"
						sRowContents = sRowContents & "&nbsp;"
					End If
					If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
						sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&Change=1&AlimonyTypeID=" & CStr(oRecordset.Fields("AlimonyTypeID").Value) & "&ReasonID=" & lReasonID & """>"
							sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar registro"" BORDER=""0"" />"
						sRowContents = sRowContents & "</A>&nbsp;"
					End If
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AlimonyTypeShortName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AlimonyTypeName").Value))
				'sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ConceptAmount").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("QttyName").Value))
				Call GetConceptNamesFromAppliesToID(oADODBConnection, CStr(oRecordset.Fields("AppliesToID").Value), sConceptNames, sErrorDescription)
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(sConceptNames)
				sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
				sFontEnd = "</FONT>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
			Response.Write "</TABLE><BR /><BR />"
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en el catálogo de tipos de pensión alimenticia."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayAlimonyTypesTable = lErrorNumber
	Err.Clear
End Function

Function DisplayCreditorsTypesTable(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the AlimonyTypes list from the database
'Inputs:  oRequest, oADODBConnection, bForExport, aAbsenceComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayCreditorsTypesTable"
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sFontBegin
	Dim sFontEnd
	Dim sBoldBegin
	Dim sBoldEnd
	Dim sNames
	Dim lErrorNumber
	Dim sConceptNames

	lErrorNumber = GetCreditorsTypes(oRequest, oADODBConnection, aEmployeeComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><TABLE BORDER="""
			If bForExport Then
				Response.Write "1"
			Else
				Response.Write "0"
			End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			If Not bForExport And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
				asColumnsTitles = Split("Acciones,Clave,Nombre del tipo de descuento,Unidad,Conceptos sobre los que se aplica", ",", -1, vbBinaryCompare)
				asCellWidths = Split(",,,,,,1500",",", -1, vbBinaryCompare)
				asCellAlignments = Split("CENTER,,,,,", ",", -1, vbBinaryCompare)
			Else
				asColumnsTitles = Split("Clave, Nombre del tipo de descuento,Tipo de unidad,Conceptos sobre los cuales se aplica", ",", -1, vbBinaryCompare)
				asCellWidths = Split(",,,,,,",",", -1, vbBinaryCompare)
				asCellAlignments = Split("CENTER,,,,,", ",", -1, vbBinaryCompare)
			End If

			If bForExport Then
				lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
			Else
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If
			End If
			sBoldBegin = "<B>"
			sBoldEnd = "</B>"
			sFontBegin = ""
			sFontEnd = ""
			Do While Not oRecordset.EOF
				sConceptNames = ""
				If Not bForExport Then
					If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
						sRowContents = "<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&Delete=1&CreditorTypeID=" & CStr(oRecordset.Fields("CreditorTypeID").Value) & "&ReasonID=" & lReasonID & """>"
							sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar registro"" BORDER=""0"" />"
						sRowContents = sRowContents & "</A>&nbsp;"
					Else
							sRowContents = "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""8"" BORDER=""0"" />"
						sRowContents = sRowContents & "&nbsp;"
					End If
					If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
						sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&Change=1&CreditorTypeID=" & CStr(oRecordset.Fields("CreditorTypeID").Value) & "&ReasonID=" & lReasonID & """>"
							sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar registro"" BORDER=""0"" />"
						sRowContents = sRowContents & "</A>&nbsp;"
					End If
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CreditorTypeShortName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CreditorTypeName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("QttyName").Value))
				Call GetConceptNamesFromAppliesToID(oADODBConnection, CStr(oRecordset.Fields("AppliesToID").Value), sConceptNames, sErrorDescription)
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(sConceptNames)
				sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
				sFontEnd = "</FONT>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
			Response.Write "</TABLE><BR /><BR />"
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en el catálogo de tipos de pensión alimenticia."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayCreditorsTypesTable = lErrorNumber
	Err.Clear
End Function

Function GetAlimonyType(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a Alimony Type
'         from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetAlimonyType"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	If aEmployeeComponent(N_ALIMONY_TYPE_ID_BENEFICIARY_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del tipo de pensión alimenticia."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "AlimonyTypeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del tipo de pensión alimenticia."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From AlimonyTypes Where (AlimonyTypeID=" & aEmployeeComponent(N_ALIMONY_TYPE_ID_BENEFICIARY_EMPLOYEE)& ")", "AlimonyTypeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El tipo de pensión alimenticia especificada no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "AlimonyTypeComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aEmployeeComponent(S_ALIMONY_TYPE_SHORT_NAME) = CStr(oRecordset.Fields("AlimonyTypeShortName").Value)
				aEmployeeComponent(S_ALIMONY_TYPE_NAME) = CStr(oRecordset.Fields("AlimonyTypeName").Value)
				'aEmployeeComponent(D_ALIMONY_AMOUNT_BENEFICIARY_EMPLOYEE) = CLng(oRecordset.Fields("ConceptAmount").Value)
				aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = Replace(CStr(oRecordset.Fields("AppliesToID").Value), " ", "")
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = CLng(oRecordset.Fields("ConceptQttyID").Value)
				aEmployeeComponent(N_ACTIVE_EMPLOYEE) = CInt(oRecordset.Fields("Active").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetAlimonyType = lErrorNumber
	Err.Clear
End Function

Function GetAlimonyTypes(oRequest, oADODBConnection, aEmployeeComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the Alimony Types
'         from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetAlimonyTypes"
	Dim sSort
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) = Trim(aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE))
	If Len(aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE)) > 0 Then
		If InStr(1, aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE), "And ", vbBinaryCompare) <> 1 Then aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) = "And " & aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE)
	End If

	'sSort = aEmployeeComponent(S_SORT_COLUMN_EMPLOYEE)
	'If aEmployeeComponent(B_SORT_DESCENDING_EMPLOYEE) Then sSort = Replace(sSort, ", ", " Desc, ") & " Desc"
	sErrorDescription = "No se pudo obtener la información de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from AlimonyTypes, QttyValues Where AlimonyTypes.ConceptQttyID = QttyValues.QttyID " & aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE), "AlimonyTypeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)

	GetAlimonyTypes = lErrorNumber
	Err.Clear
End Function

Function GetAlimonyTypesForPercent(sAlimonyTypeIDs, sErrorDescription)
'************************************************************
'Purpose: To get the absences requerids for insert te
'         absence for employee
'Inputs:  oRequest, oADODBConnection, aAbsenceComponent
'Outputs: sAbsenceIDs, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetAlimonyTypesForPercent"
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información del tipo de pensión alimenticia."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From AlimonyTypes Where (ConceptQttyID=2)", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen pensiones alimenticias que apliquen por porcentaje en el catálogo."
		Else
			Do While Not oRecordset.EOF
				sAlimonyTypeIDs = sAlimonyTypeIDs & CStr(oRecordset.Fields("AlimonyTypeID").Value) & ","
				oRecordset.MoveNext
			Loop
			If (InStr(Right(sAlimonyTypeIDs,1),",") > 0) Then
				sAlimonyTypeIDs = Left(sAlimonyTypeIDs, Len(sAbsenceIDs) -1)
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	GetAlimonyTypesForPercent = lErrorNumber
	Err.Clear
End Function

Function GetAlimonyTypesTotalAmountForEmployee(lTotalAmount, sErrorDescription)
'************************************************************
'Purpose: To get the absences requerids for insert te
'         absence for employee
'Inputs:  oRequest, oADODBConnection, aAbsenceComponent
'Outputs: sAbsenceIDs, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetAlimonyTypesTotalAmountForEmployee"
	Dim oRecordset
	Dim lErrorNumber
	Dim sSpecialCondition

	If aEmployeeComponent(N_ALIMONY_TYPE_ID_BENEFICIARY_EMPLOYEE) <> -1 Then
		sSpecialCondition = " And (AlimonyTypes.AlimonyTypeID=" & aEmployeeComponent(N_ALIMONY_TYPE_ID_BENEFICIARY_EMPLOYEE) & ")"
	End If
	sErrorDescription = "No se pudo obtener la información de las pensiones del empleado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesBeneficiariesLKP.AlimonyTypeID, SUM(ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, AlimonyTypes Where (EmployeesBeneficiariesLKP.AlimonyTypeID=AlimonyTypes.AlimonyTypeID) And (ConceptQttyID = 2) And (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (((EmployeesBeneficiariesLKP.StartDate>=" & aEmployeeComponent(N_START_DATE_BENEFICIARY_EMPLOYEE) & ") And (EmployeesBeneficiariesLKP.EndDate<=" & aEmployeeComponent(N_END_DATE_BENEFICIARY_EMPLOYEE) & ")) Or ((EmployeesBeneficiariesLKP.EndDate>=" & aEmployeeComponent(N_START_DATE_BENEFICIARY_EMPLOYEE) & ") And (EmployeesBeneficiariesLKP.EndDate<=" & aEmployeeComponent(N_END_DATE_BENEFICIARY_EMPLOYEE) & ")) Or ((EmployeesBeneficiariesLKP.EndDate>=" & aEmployeeComponent(N_START_DATE_BENEFICIARY_EMPLOYEE) & ") And (EmployeesBeneficiariesLKP.StartDate<=" & aEmployeeComponent(N_END_DATE_BENEFICIARY_EMPLOYEE) & "))) " & sSpecialCondition & " Group By EmployeesBeneficiariesLKP.AlimonyTypeID", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			lTotalAmount = 0
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen pensión alimenticia que apliquen por porcentaje para este empleado."
		Else
			Do While Not oRecordset.EOF
				If Not IsNull(oRecordset.Fields("TotalAmount").Value) Then
					lTotalAmount = lTotalAmount & CLng(oRecordset.Fields("AlimonyTypeID").Value) & LIST_SEPARATOR & CLng(oRecordset.Fields("TotalAmount").Value) & SECOND_LIST_SEPARATOR
				Else
					lTotalAmount = lTotalAmount & CLng(oRecordset.Fields("AlimonyTypeID").Value) & LIST_SEPARATOR & "0" & SECOND_LIST_SEPARATOR
				End If
				oRecordset.MoveNext
			Loop
		End If
		oRecordset.Close
	Else
		lTotalAmount = 0
		lErrorNumber = L_ERR_NO_RECORDS
		sErrorDescription = "Error al obtener la información de las pensiones del empleado."
	End If

	Set oRecordset = Nothing
	GetAlimonyTypesTotalAmountForEmployee = lErrorNumber
	Err.Clear
End Function

Function GetCreditorType(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a Alimony Type
'         from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetCreditorType"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	If aEmployeeComponent(N_CREDITOR_TYPE_ID_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del tipo de pensión alimenticia."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "AlimonyTypeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del tipo de pensión alimenticia."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From CreditorsTypes Where (CreditorTypeID=" & aEmployeeComponent(N_CREDITOR_TYPE_ID_EMPLOYEE)& ")", "AlimonyTypeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El tipo de pensión alimenticia especificada no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "AlimonyTypeComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aEmployeeComponent(S_CREDITOR_TYPE_SHORT_NAME) = CStr(oRecordset.Fields("CreditorTypeShortName").Value)
				aEmployeeComponent(S_CREDITOR_TYPE_NAME) = CStr(oRecordset.Fields("CreditorTypeName").Value)
				aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = Replace(CStr(oRecordset.Fields("AppliesToID").Value), " ", "")
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = CLng(oRecordset.Fields("ConceptQttyID").Value)
				aEmployeeComponent(N_ACTIVE_EMPLOYEE) = CInt(oRecordset.Fields("Active").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetCreditorType = lErrorNumber
	Err.Clear
End Function

Function GetCreditorsTypes(oRequest, oADODBConnection, aEmployeeComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the Creditors Types
'         from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetCreditorsTypes"
	Dim sSort
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) = Trim(aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE))
	If Len(aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE)) > 0 Then
		If InStr(1, aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE), "And ", vbBinaryCompare) <> 1 Then aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) = "And " & aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE)
	End If

	sErrorDescription = "No se pudo obtener la información de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from CreditorsTypes, QttyValues Where CreditorsTypes.ConceptQttyID = QttyValues.QttyID " & aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE), "AlimonyTypeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)

	GetCreditorsTypes = lErrorNumber
	Err.Clear
End Function

Function GetCreditorTypesForPercent(sCreditorTypeIDs, sErrorDescription)
'************************************************************
'Purpose: To get the absences requerids for insert te
'         absence for employee
'Inputs:  oRequest, oADODBConnection, aAbsenceComponent
'Outputs: sAbsenceIDs, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetCreditorTypesForPercent"
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información del tipo de pensión alimenticia."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From CreditorsTypes Where (ConceptQttyID=2)", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen pensiones alimenticias que apliquen por porcentaje en el catálogo."
		Else
			Do While Not oRecordset.EOF
				sCreditorTypeIDs = sCreditorTypeIDs & CStr(oRecordset.Fields("CreditorTypeID").Value) & ","
				oRecordset.MoveNext
			Loop
			If (InStr(Right(sCreditorTypeIDs,1),",") > 0) Then
				sCreditorTypeIDs = Left(sCreditorTypeIDs, Len(sAbsenceIDs) -1)
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	GetCreditorTypesForPercent = lErrorNumber
	Err.Clear
End Function

Function GetCreditorTypesTotalAmountForEmployee(lTotalAmount, sErrorDescription)
'************************************************************
'Purpose: To get the absences requerids for insert te
'         absence for employee
'Inputs:  oRequest, oADODBConnection, aAbsenceComponent
'Outputs: sAbsenceIDs, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetCreditorTypesTotalAmountForEmployee"
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información de las pensiones del empleado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesCreditorsLKP.CreditorTypeID, SUM(ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, CreditorsTypes Where (EmployeesCreditorsLKP.CreditorTypeID=CreditorsTypes.CreditorTypeID) And (ConceptQttyID = 2) And (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeesCreditorsLKP.StartDate<=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ") And (EmployeesCreditorsLKP.EndDate>=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ") Group By EmployeesCreditorsLKP.CreditorTypeID", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			lTotalAmount = 0
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen pensión alimenticia que apliquen por porcentaje para este empleado."
		Else
			Do While Not oRecordset.EOF
				If Not IsNull(oRecordset.Fields("TotalAmount").Value) Then
					lTotalAmount = lTotalAmount & CLng(oRecordset.Fields("CreditorTypeID").Value) & LIST_SEPARATOR & CLng(oRecordset.Fields("TotalAmount").Value) & SECOND_LIST_SEPARATOR
				Else
					lTotalAmount = lTotalAmount & CLng(oRecordset.Fields("CreditorTypeID").Value) & LIST_SEPARATOR & "0" & SECOND_LIST_SEPARATOR
				End If
				oRecordset.MoveNext
			Loop
		End If
		oRecordset.Close
	Else
		lTotalAmount = 0
		lErrorNumber = L_ERR_NO_RECORDS
		sErrorDescription = "Error al obtener la información de las pensiones del empleado."
	End If

	Set oRecordset = Nothing
	GetCreditorTypesTotalAmountForEmployee = lErrorNumber
	Err.Clear
End Function

Function GetConceptNamesFromAppliesToID(oADODBConnection, sAppliesToIDs, sConceptNames, sErrorDescription)
'************************************************************
'Purpose: To get the names of a concept list
'Inputs:  oADODBConnection, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetConceptNamesFromAppliesToID"
	Dim lErrorNumber
	Dim oRecordset

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from Concepts Where (ConceptID IN(" & sAppliesToIDs & "))", "AlimonyTypeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Do While Not oRecordset.EOF
				sConceptNames = sConceptNames & " " & CStr(oRecordset.Fields("ConceptName").Value) & ","
				oRecordset.MoveNext
			Loop
			If (InStr(Right(sAbsencesConceptNamesIDs,1),",") > 0) Then
				sConceptNames = Left(sConceptNames, Len(sConceptNames) -1)
			End If
			oRecordset.Close
		Else
			sErrorDescription = "No se encontraron los conceptos indicados"
		End If
	Else
		sErrorDescription = "Error al obtener los nombres de los conceptos."
		lErrorNumber = -1
	End If

	Set oRecordset = Nothing
	GetConceptNamesFromAppliesToID = lErrorNumber
	Err.Clear
End Function

Function ModifyAlimonyType(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To modify the Alimony Type information in
'         the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************

	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyAlimonyType"
	Dim oRecordset
	Dim lErrorNumber
	Dim sField
	Dim bComponentInitialized
	Dim sQuery

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ALIMONY_TYPE_ID_BENEFICIARY_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del tipo de pensión alimenticia."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "AlimonyTypeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sQuery = "Update AlimonyTypes Set AlimonyTypeShortName = '" & aEmployeeComponent(S_ALIMONY_TYPE_SHORT_NAME) & "'" & _
				 ", AlimonyTypeName = '" & aEmployeeComponent(S_ALIMONY_TYPE_NAME) & "'" & _
				 ", ConceptQttyID = " & aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) & _
				 ", AppliesToID = '" & aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) & "'" & _
				 ", Active = " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & _
				 " Where AlimonyTypeID= " & aEmployeeComponent(N_ALIMONY_TYPE_ID_BENEFICIARY_EMPLOYEE)

		sErrorDescription = "No se pudo modificar la información del tipo de pensión alimenticia."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AlimonyTypeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	ModifyAlimonyType = lErrorNumber
	Err.Clear
End Function

Function RemoveAlimonyType(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a Alimony Type record from database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveAlimonyType"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If aEmployeeComponent(N_ALIMONY_TYPE_ID_BENEFICIARY_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del tipo de pensión alimenticia."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "AlimonyTypeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If VerifyExistenceOfAlimonyTypeForBeneficiaries(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
			lErrorNumber = -1
			sErrorDescription = "No se puede eliminar la información del tipo de pensión alimenticia, debido a que existen beneficiarios con este tipo de pensión."
		Else
			sErrorDescription = "No se pudo eliminar la información del tipo de pensión alimenticia."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From AlimonyTypes Where (AlimonyTypeID=" & aEmployeeComponent(N_ALIMONY_TYPE_ID_BENEFICIARY_EMPLOYEE) & ")", "AlimonyTypeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
		End If
	End If

	RemoveAlimonyType = lErrorNumber
	Err.Clear
End Function

Function VerifyExistenceOfAlimonyTypeForBeneficiaries(oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To verify if an absence already exist in database
'Inputs:  oADODBConnection, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyExistenceOfAlimonyTypeForBeneficiaries"
	Dim lErrorNumber
	Dim oRecordset

	If aEmployeeComponent(N_ALIMONY_TYPE_ID_BENEFICIARY_EMPLOYEE) = -1 Then
		VerifyExistenceOfAlimonyTypeForBeneficiaries = True
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesBeneficiariesLKP Where (AlimonyTypeID = " & aEmployeeComponent(N_ALIMONY_TYPE_ID_BENEFICIARY_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				oRecordset.Close
				VerifyExistenceOfAlimonyTypeForBeneficiaries = True
			Else
				VerifyExistenceOfAlimonyTypeForBeneficiaries = False
			End If
		Else
			sErrorDescription = "Error al verificar si existe el tipo de pensión registrada para algún beneficiario de pensión alimenticia."
			VerifyExistenceOfAlimonyTypeForBeneficiaries = True
		End If
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function
%>