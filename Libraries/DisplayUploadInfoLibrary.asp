<%
Function DisplayChildrenSchoolarshipsColumns(sFileName, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayChildrenSchoolarshipsColumns"
	Dim iColumns
	Dim iIndex
	Dim lErrorNumber

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "Indique a qué campo pertenece cada columna del archivo.")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""UploadAbsencesFrm"" ID=""UploadAbsencesFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
			For iIndex = 1 To iColumns
				Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
				Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeID"">No. del empleado</OPTION>"
					Response.Write "<OPTION VALUE=""ChildName"">Nombre del hijo(a)</OPTION>"
					Response.Write "<OPTION VALUE=""ChildLastName"">Apellido paterno del hijo(a)</OPTION>"
					Response.Write "<OPTION VALUE=""ChildLastName2"">Apellido materno del hijo(a)</OPTION>"
					Response.Write "<OPTION VALUE=""LevelID"">Grado escolar</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateYYYYMMDD"">Fecha de nacimiento (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateDDMMYYYY"">Fecha de nacimiento (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateMMDDYYYY"">Fecha de nacimiento (MM-DD-AAAA)</OPTION>"
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			Response.Write "<BR />"
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
				Response.Write "var bDuplicated = false;" & vbNewLine
				Response.Write "var sFields = '';" & vbNewLine

				For iIndex = 1 To iColumns
					Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
						Response.Write "if (sFields.search(eval('/' + oForm.Column" & iIndex & ".value + '/gi')) == -1)" & vbNewLine
							Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
						Response.Write "else" & vbNewLine
							Response.Write "bDuplicated = true;" & vbNewLine
					Response.Write "}" & vbNewLine
				Next
				Response.Write "if (bDuplicated) {" & vbNewLine
					Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene los números de los empleados.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/ChildName/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el nombre del hijo(a).');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/ChildLastName/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el apellido paterno del hijo(a).');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/LevelID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el grado escolar de la beca.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/OcurredDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha de nacimiento de los hijos.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/OcurredDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de nacimiento de los hijos con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	DisplayChildrenSchoolarshipsColumns = lErrorNumber
	Err.Clear
End Function

Function DisplayConceptsValuesColumns(sFileName, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayConceptsValuesColumns"
	Dim iColumns
	Dim iIndex
	Dim lErrorNumber

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "Indique a qué campo pertenece cada columna del archivo.")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""UploadAbsencesFrm"" ID=""UploadAbsencesFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
			For iIndex = 1 To iColumns
				Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
				Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeTypeID"">Tipo de tabulador</OPTION>"
					Response.Write "<OPTION VALUE=""PositionTypeID"">Tipo de puesto</OPTION>"
					Response.Write "<OPTION VALUE=""PositionShortNames"">Puesto</OPTION>"
					Response.Write "<OPTION VALUE=""LevelID"">Nivel</OPTION>"
					Response.Write "<OPTION VALUE=""GroupGradeLevelID"">Grupo grado nivel</OPTION>"
					Response.Write "<OPTION VALUE=""ClassificationID"">Clasificación</OPTION>"
					Response.Write "<OPTION VALUE=""IntegrationID"">Integración</OPTION>"
					Response.Write "<OPTION VALUE=""WorkingHours"">Jornada</OPTION>"
					Response.Write "<OPTION VALUE=""EconomicZoneID"">Zona económica</OPTION>"
					Response.Write "<OPTION VALUE=""ConceptID"">Clave concepto de pago</OPTION>"
					Response.Write "<OPTION VALUE=""ConceptAmount"">Monto quincenal</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredStartDateYYYYMMDD"">Fecha de inicio vigencia (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredStartDateDDMMYYYY"">Fecha de inicio vigencia (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredStartDateMMDDYYYY"">Fecha de inicio vigencia (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredEndDateYYYYMMDD"">Fecha de fin vigencia (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredEndDateDDMMYYYY"">Fecha de fin vigencia (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredEndDateMMDDYYYY"">Fecha de fin vigencia (MM-DD-AAAA)</OPTION>"
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			Response.Write "<BR />"
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
				Response.Write "var bDuplicated = false;" & vbNewLine
				Response.Write "var sFields = '';" & vbNewLine

				For iIndex = 1 To iColumns
					Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
						Response.Write "if (sFields.search(eval('/' + oForm.Column" & iIndex & ".value + '/gi')) == -1)" & vbNewLine
							Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
						Response.Write "else" & vbNewLine
							Response.Write "bDuplicated = true;" & vbNewLine
					Response.Write "}" & vbNewLine
				Next

				Response.Write "if (bDuplicated) {" & vbNewLine
					Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/EmployeeTypeID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el tipo de tabulador.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/PositionShortNames/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el código del puesto.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/ConceptID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la clave del concepto de pago.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/ConceptAmount/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el monto de pago.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/OcurredStartDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredStartDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredStartDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha de inicio vigengia.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/StartDateYYYYMMDD/gi) != -1) && ((sFields.search(/StartDateDDMMYYYY/gi) != -1) || (sFields.search(/StartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/StartDateDDMMYYYY/gi) != -1) && ((sFields.search(/StartDateYYYYMMDD/gi) != -1) || (sFields.search(/StartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/StartDateMMDDYYYY/gi) != -1) && ((sFields.search(/StartDateDDMMYYYY/gi) != -1) || (sFields.search(/StartDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de inicio vigengia con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	DisplayConceptsValuesColumns = lErrorNumber
	Err.Clear
End Function

Function DisplayDocumentsForLicensesColumns(sFileName, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayDocumentsForLicensesColumns"
	Dim iColumns
	Dim iIndex
	Dim lErrorNumber

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "Indique a qué campo pertenece cada columna del archivo.")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""UploadAbsencesFrm"" ID=""UploadAbsencesFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
			For iIndex = 1 To iColumns
				Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
				Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeID"">No. del empleado</OPTION>"
					Response.Write "<OPTION VALUE=""DocumentForLicenseNumber"">No. del oficio</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDocumentDateYYYYMMDD"">Fecha del documento (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDocumentDateDDMMYYYY"">Fecha del documento (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDocumentDateMMDDYYYY"">Fecha del documento (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""RequestNumber"">No. de la solicitud</OPTION>"
					Response.Write "<OPTION VALUE=""LicenseSyndicateTypeID"">Tipo de la licencia</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredStartDateYYYYMMDD"">Fecha inicio de la licencia sindical (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredStartDateDDMMYYYY"">Fecha inicio de la licencia sindical (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredStartDateMMDDYYYY"">Fecha inicio de la licencia sindical (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredEndDateYYYYMMDD"">Fecha fin de la licencia sindical (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredEndDateDDMMYYYY"">Fecha fin de la licencia sindical (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredEndDateMMDDYYYY"">Fecha fin de la licencia sindical (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""DocumentTemplate"">Nombre de la plantilla</OPTION>"
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			Response.Write "<BR />"
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
				Response.Write "var bDuplicated = false;" & vbNewLine
				Response.Write "var sFields = '';" & vbNewLine

				For iIndex = 1 To iColumns
					Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
						Response.Write "if (sFields.search(eval('/' + oForm.Column" & iIndex & ".value + '/gi')) == -1)" & vbNewLine
							Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
						Response.Write "else" & vbNewLine
							Response.Write "bDuplicated = true;" & vbNewLine
					Response.Write "}" & vbNewLine
				Next

				Response.Write "if (bDuplicated) {" & vbNewLine
					Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene los números de los empleados.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/OcurredDocumentDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredDocumentDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredDocumentDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha del documento.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/OcurredDocumentDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredDocumentDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDocumentDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDocumentDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredDocumentDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredDocumentDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDocumentDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredDocumentDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDocumentDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha del documento.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/OcurredStartDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredStartDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredStartDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha de inicio de la licencia sindical.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de inicio de la licencia sindical con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/OcurredEndDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredEndDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredEndDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha de término de la licencia sindical.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/OcurredEndDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredEndDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredEndDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredEndDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredEndDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredEndDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredEndDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredEndDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredEndDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de término de la licencia sindical con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/DocumentForLicenseNumber/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene los números de oficios.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/RequestNumber/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene los números de solicitudes.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/LicenseSyndicateTypeID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene los tipos de licencias.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/DocumentTemplate/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene los nombres de las plantilla.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	DisplayDocumentsForLicensesColumns = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesDropColumns(sFileName, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesDropColumns"
	Dim iColumns
	Dim iIndex
	Dim lErrorNumber

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "Indique a qué campo pertenece cada columna del archivo.")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""UploadAbsencesFrm"" ID=""UploadAbsencesFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
			For iIndex = 1 To iColumns
				Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
				Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeID"">No. del empleado</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateYYYYMMDD"">Fecha de baja (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateDDMMYYYY"">Fecha de baja (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateMMDDYYYY"">Fecha de baja (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""DocumentNumber"">No. de oficio</OPTION>"
					Response.Write "<OPTION VALUE=""Reasons"">Observaciones</OPTION>"
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			Response.Write "<BR />"
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
				Response.Write "var bDuplicated = false;" & vbNewLine
				Response.Write "var sFields = '';" & vbNewLine

				For iIndex = 1 To iColumns
					Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
						Response.Write "if (sFields.search(eval('/' + oForm.Column" & iIndex & ".value + '/gi')) == -1)" & vbNewLine
							Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
						Response.Write "else" & vbNewLine
							Response.Write "bDuplicated = true;" & vbNewLine
					Response.Write "}" & vbNewLine
				Next

				Response.Write "if (bDuplicated) {" & vbNewLine
					Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene los números de los empleados.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/OcurredDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha de ocurencia.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/OcurredDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de ocurencia con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	DisplayEmployeesDropColumns = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesAbsencesColumns(lReasonID, sFileName, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesAbsencesColumns"
	Dim iColumns
	Dim iIndex
	Dim lErrorNumber

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "Indique a qué campo pertenece cada columna del archivo.")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""UploadAbsencesFrm"" ID=""UploadAbsencesFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE="&lReasonID&" />"
			For iIndex = 1 To iColumns
				Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
				Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeID"">No. del empleado</OPTION>"
					Select Case lReasonID
						Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS
						Case Else
							Response.Write "<OPTION VALUE=""AbsenceID"">Clave de incidencia</OPTION>"
					End Select
					Response.Write "<OPTION VALUE=""OcurredDateYYYYMMDD"">Fecha de ocurrencia (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateDDMMYYYY"">Fecha de ocurrencia (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateMMDDYYYY"">Fecha de ocurrencia (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""PayrollDateYYYYMMDD"">Fecha de aplicación (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""PayrollDateDDMMYYYY"">Fecha de aplicación (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""PayrollDateMMDDYYYY"">Fecha de aplicación (MM-DD-AAAA)</OPTION>"
					Select Case lReasonID
						Case EMPLOYEES_EXTRAHOURS
							Response.Write "<OPTION VALUE=""AbsenceHours"">Horas extras</OPTION>"
						Case EMPLOYEES_SUNDAYS
						Case 1
						Case Else
							Response.Write "<OPTION VALUE=""DocumentNumber"">No. de oficio</OPTION>"
							Response.Write "<OPTION VALUE=""AbsenceHours"">Horas de retardo</OPTION>"
							Response.Write "<OPTION VALUE=""JustificationID"">Justificación</OPTION>"
							Response.Write "<OPTION VALUE=""AppliesForPunctuality"">¿Aplica para puntualidad?</OPTION>"
					End Select
					Response.Write "<OPTION VALUE=""Reasons"">Observaciones</OPTION>"
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			Response.Write "<BR />"
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
				Response.Write "var bDuplicated = false;" & vbNewLine
				Response.Write "var sFields = '';" & vbNewLine

				For iIndex = 1 To iColumns
					Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
						Response.Write "if (sFields.search(eval('/' + oForm.Column" & iIndex & ".value + '/gi')) == -1)" & vbNewLine
							Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
						Response.Write "else" & vbNewLine
							Response.Write "bDuplicated = true;" & vbNewLine
					Response.Write "}" & vbNewLine
				Next

				Response.Write "if (bDuplicated) {" & vbNewLine
					Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene los números de los empleados.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Select Case lReasonID
					Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS
					Case Else
						Response.Write "if (sFields.search(/AbsenceID/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se especificó qué columna contiene los tipos de ausencia.');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
				End Select
				Response.Write "if ((sFields.search(/OcurredDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha de ocurrencia.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/OcurredDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de ocurrencia con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/PayrollDateYYYYMMDD/gi) == -1) && (sFields.search(/PayrollDateDDMMYYYY/gi) == -1) && (sFields.search(/PayrollDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha de aplicación.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/PayrollDateYYYYMMDD/gi) != -1) && ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) || (sFields.search(/PayrollDateMMDDYYYY/gi) != -1))) || ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) && ((sFields.search(/PayrollDateYYYYMMDD/gi) != -1) || (sFields.search(/PayrollDateMMDDYYYY/gi) != -1))) || ((sFields.search(/PayrollDateMMDDYYYY/gi) != -1) && ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) || (sFields.search(/PayrollDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de aplicación con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Select Case lReasonID
					Case EMPLOYEES_EXTRAHOURS
						Response.Write "if (sFields.search(/AbsenceHours/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se especificó qué columna contiene el número de horas extras.');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
				End Select
				If False Then
					Response.Write "if (sFields.search(/Reasons/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se especificó qué columna contiene las observaciones del registro.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "return true;" & vbNewLine
				End If
			Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	DisplayEmployeesAbsencesColumns = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesAdjustmentsColumns(sFileName, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesAdjustmentsColumns"
	Dim iColumns
	Dim iIndex
	Dim lErrorNumber

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "Indique a qué campo pertenece cada columna del archivo.")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""UploadAdjustmentsFrm"" ID=""UploadAdjustmentsFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE=""-58"" />"
			For iIndex = 1 To iColumns
				Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
				Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeID"">Número de Empleado</OPTION>"
					Response.Write "<OPTION VALUE=""ConceptID"">Número de Concepto</OPTION>"
					Response.Write "<OPTION VALUE=""ConceptAmount"">Cantidad a ajustar</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredMissingDateYYYYMMDD"">Fecha de omisión de pago (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredMissingDateDDMMYYYY"">Fecha de omisión de pago (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredMissingDateMMDDYYYY"">Fecha de omisión de pago (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredPayrollDateYYYYMMDD"">Fecha de aplicación de nómina (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredPayrollDateDDMMYYYY"">Fecha de aplicación de nómina (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredPayrollDateMMDDYYYY"">Fecha de aplicación de nómina (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""BeneficiaryName"">Nombre del beneficiario</OPTION>"
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			Response.Write "<BR />"
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
				Response.Write "var bDuplicated = false;" & vbNewLine
				Response.Write "var sFields = '';" & vbNewLine
				For iIndex = 1 To iColumns
					Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
						Response.Write "if (sFields.search(eval('/' + oForm.Column" & iIndex & ".value + '/gi')) == -1)" & vbNewLine
							Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
						Response.Write "else" & vbNewLine
							Response.Write "bDuplicated = true;" & vbNewLine
					Response.Write "}" & vbNewLine
				Next
				Response.Write "if (bDuplicated) {" & vbNewLine
					Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene Número de Empleado.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/ConceptID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene Número de Concepto.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/ConceptAmount/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene Cantidad a ajustar.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/OcurredMissingDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredMissingDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredMissingDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la Fecha de omisión de pago.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/OcurredMissingDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredMissingDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredMissingDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredMissingDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredMissingDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredMissingDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredMissingDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredMissingDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredMissingDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de omisión de pago con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/OcurredPayrollDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredPayrollDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredPayrollDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha de aplicación de nómina.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/PayrollDateYYYYMMDD/gi) != -1) && ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) || (sFields.search(/PayrollDateMMDDYYYY/gi) != -1))) || ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) && ((sFields.search(/PayrollDateYYYYMMDD/gi) != -1) || (sFields.search(/PayrollDateMMDDYYYY/gi) != -1))) || ((sFields.search(/PayrollDateMMDDYYYY/gi) != -1) && ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) || (sFields.search(/PayrollDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de aplicacion de nómina.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	DisplayEmployeesAdjustmentsColumns = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesBankAccountColumns(lReasonID, sFileName, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesBankAccountColumns"
	Dim iColumns
	Dim iIndex
	Dim lErrorNumber

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "Indique a qué campo pertenece cada columna del archivo.")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""BankAccountsFrm"" ID=""BankAccountsFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE="&lReasonID&" />"
			For iIndex = 1 To iColumns
				Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
				Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeID"">No. del empleado</OPTION>"
					Response.Write "<OPTION VALUE=""BankID"">ID del banco</OPTION>"
					Response.Write "<OPTION VALUE=""StartDateYYYYMMDD"">Fecha (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""StartDateDDMMYYYY"">Fecha (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""StartDateMMDDYYYY"">Fecha (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""AccountNumber"">No. de cuenta</OPTION>"
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			Response.Write "<BR />"
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
				Response.Write "var bDuplicated = false;" & vbNewLine
				Response.Write "var sFields = '';" & vbNewLine

				For iIndex = 1 To iColumns
					Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
						Response.Write "if (sFields.search(eval('/' + oForm.Column" & iIndex & ".value + '/gi')) == -1)" & vbNewLine
							Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
						Response.Write "else" & vbNewLine
							Response.Write "bDuplicated = true;" & vbNewLine
					Response.Write "}" & vbNewLine
				Next
				Response.Write "if (bDuplicated) {" & vbNewLine
					Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene los números de los empleados.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/BankID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha de ocurencia.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/StartDateYYYYMMDD/gi) == -1) && (sFields.search(/StartDateDDMMYYYY/gi) == -1) && (sFields.search(/StartDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha de registro de la cuenta bancaria.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/StartDateYYYYMMDD/gi) != -1) && ((sFields.search(/StartDateDDMMYYYY/gi) != -1) || (sFields.search(/StartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/StartDateDDMMYYYY/gi) != -1) && ((sFields.search(/StartDateYYYYMMDD/gi) != -1) || (sFields.search(/StartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/StartDateMMDDYYYY/gi) != -1) && ((sFields.search(/StartDateDDMMYYYY/gi) != -1) || (sFields.search(/StartDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de registro de la cuenta bancaria con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/AccountNumber/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el número de cuenta.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	DisplayEmployeesBankAccountColumns = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesBeneficiariesDebitColumns(lReasonID, sFileName, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesBeneficiariesDebitColumns"
	Dim iColumns
	Dim iIndex
	Dim lErrorNumber

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "Indique a qué campo pertenece cada columna del archivo.")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""UploadEmployeesSafeSeparationFrm"" ID=""UploadEmployeesSafeSeparationFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE="&lReasonID&" />"
			For iIndex = 1 To iColumns
				Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
				Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeID"">Número de Empleado</OPTION>"
					Response.Write "<OPTION VALUE=""StartDateYYYYMMDD"">Fecha de inicio (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""StartDateDDMMYYYY"">Fecha de inicio (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""StartDateMMDDYYYY"">Fecha de inicio (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""EndDateYYYYMMDD"">Fecha de termino de la prestación (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""EndDateDDMMYYYY"">Fecha de termino de la prestación (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""EndDateMMDDYYYY"">Fecha de termino de la prestación (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""PayrollDateYYYYMMDD"">Fecha de aplicación de nómina (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""PayrollDateDDMMYYYY"">Fecha de aplicación de nómina (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""PayrollDateMMDDYYYY"">Fecha de aplicación de nómina (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""ConceptAmount"">Importe del adeudo</OPTION>"
					Response.Write "<OPTION VALUE=""BeneficiaryNumber"">Número de beneficiario</OPTION>"
					Response.Write "<OPTION VALUE=""ConceptComments"">Observaciones</OPTION>"
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			Response.Write "<BR />"
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
				Response.Write "var bDuplicated = false;" & vbNewLine
				Response.Write "var sFields = '';" & vbNewLine
				For iIndex = 1 To iColumns
					Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
						Response.Write "if (sFields.search(eval('/' + oForm.Column" & iIndex & ".value + '/gi')) == -1)" & vbNewLine
							Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
						Response.Write "else" & vbNewLine
							Response.Write "bDuplicated = true;" & vbNewLine
					Response.Write "}" & vbNewLine
				Next
				Response.Write "if (bDuplicated) {" & vbNewLine
					Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
                Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
                    Response.Write "alert('No se especificó qué columna contiene el Número de Empleado.');" & vbNewLine
                    Response.Write "return false;" & vbNewLine
                Response.Write "}" & vbNewLine
                Response.Write "if ((sFields.search(/StartDateYYYYMMDD/gi) == -1) && (sFields.search(/StartDateDDMMYYYY/gi) == -1) && (sFields.search(/StartDateMMDDYYYY/gi) == -1)) {" & vbNewLine
                        Response.Write "alert('No se especificó qué columna contiene la fecha de inicio.');" & vbNewLine
                        Response.Write "return false;" & vbNewLine
                Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/StartDateYYYYMMDD/gi) != -1) && ((sFields.search(/StartDateDDMMYYYY/gi) != -1) || (sFields.search(/StartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/StartDateDDMMYYYY/gi) != -1) && ((sFields.search(/StartDateYYYYMMDD/gi) != -1) || (sFields.search(/StartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/StartDateMMDDYYYY/gi) != -1) && ((sFields.search(/StartDateDDMMYYYY/gi) != -1) || (sFields.search(/StartDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de inicio con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/EndDateYYYYMMDD/gi) == -1) && (sFields.search(/EndDateDDMMYYYY/gi) == -1) && (sFields.search(/EndDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha de fin de la prestación.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/EndDateYYYYMMDD/gi) != -1) && ((sFields.search(/EndDateDDMMYYYY/gi) != -1) || (sFields.search(/EndDateMMDDYYYY/gi) != -1))) || ((sFields.search(/EndDateDDMMYYYY/gi) != -1) && ((sFields.search(/EndDateYYYYMMDD/gi) != -1) || (sFields.search(/EndDateMMDDYYYY/gi) != -1))) || ((sFields.search(/EndDateMMDDYYYY/gi) != -1) && ((sFields.search(/EndDateDDMMYYYY/gi) != -1) || (sFields.search(/EndDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de fin con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
                Response.Write "if ((sFields.search(/PayrollDateYYYYMMDD/gi) == -1) && (sFields.search(/PayrollStartDateDDMMYYYY/gi) == -1) && (sFields.search(/PayrollDateMMDDYYYY/gi) == -1)) {" & vbNewLine
                        Response.Write "alert('No se especificó qué columna contiene la fecha de aplicación en nómina.');" & vbNewLine
                        Response.Write "return false;" & vbNewLine
                Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/PayrollDateYYYYMMDD/gi) != -1) && ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) || (sFields.search(/PayrollDateMMDDYYYY/gi) != -1))) || ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) && ((sFields.search(/PayrollDateYYYYMMDD/gi) != -1) || (sFields.search(/PayrollDateMMDDYYYY/gi) != -1))) || ((sFields.search(/PayrollDateMMDDYYYY/gi) != -1) && ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) || (sFields.search(/PayrollDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de aplicación en nómina con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
                Response.Write "if (sFields.search(/ConceptAmount/gi) == -1) {" & vbNewLine
                        Response.Write "alert('No se especificó qué columna contiene el Importe del adeudo.');" & vbNewLine
                        Response.Write "return false;" & vbNewLine
                Response.Write "}" & vbNewLine
                Response.Write "if (sFields.search(/BeneficiaryNumber/gi) == -1) {" & vbNewLine
                        Response.Write "alert('No se especificó qué columna contiene el número del beneficiario.');" & vbNewLine
                        Response.Write "return false;" & vbNewLine
                Response.Write "}" & vbNewLine
                Response.Write "if (sFields.search(/BeneficiaryNumber/gi) == -1) {" & vbNewLine
                        Response.Write "alert('No se especificó qué columna contiene el número del beneficiario.');" & vbNewLine
                        Response.Write "return false;" & vbNewLine
                Response.Write "}" & vbNewLine
			Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If
	DisplayEmployeesBeneficiariesDebitColumns = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesSafeSeparationColumns(lReasonID, sAction, sFileName, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  sAction, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesSafeSeparationColumns"
	Dim iColumns
	Dim iIndex
	Dim lErrorNumber

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "Indique a qué campo pertenece cada columna del archivo.")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""UploadEmployeesSafeSeparationFrm"" ID=""UploadEmployeesSafeSeparationFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE="&lReasonID&" />"
			For iIndex = 1 To iColumns
				Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
				Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeID"">Número de Empleado</OPTION>"
					'Response.Write "<OPTION VALUE=""ConceptID"">Clave del Concepto</OPTION>"                    
					Response.Write "<OPTION VALUE=""OcurredStartDateYYYYMMDD"">Fecha de inicio (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredStartDateDDMMYYYY"">Fecha de inicio (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredStartDateMMDDYYYY"">Fecha de inicio (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""EndDateYYYYMMDD"">Fecha de termino de la prestación (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""EndDateDDMMYYYY"">Fecha de termino de la prestación (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""EndDateMMDDYYYY"">Fecha de termino de la prestación (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""PayrollDateYYYYMMDD"">Fecha de aplicación de nómina (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""PayrollDateDDMMYYYY"">Fecha de aplicación de nómina (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""PayrollDateMMDDYYYY"">Fecha de aplicación de nómina (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""ConceptComments"">Observaciones</OPTION>"
					If lReasonID = EMPLOYEES_SAFE_SEPARATION Then
						Response.Write "<OPTION VALUE=""ConceptAmount"">Porcentaje</OPTION>"
					Else
						Response.Write "<OPTION VALUE=""ConceptAmount"">Cantidad (en $ o %)</OPTION>"
						Response.Write "<OPTION VALUE=""ConceptQttyID"">Tipo de unidad de la cantidad (1=$ o 2=%)</OPTION>"
					End If
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			Response.Write "<BR />"
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
				Response.Write "var bDuplicated = false;" & vbNewLine
				Response.Write "var sFields = '';" & vbNewLine
				For iIndex = 1 To iColumns
					Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
						Response.Write "if (sFields.search(eval('/' + oForm.Column" & iIndex & ".value + '/gi')) == -1)" & vbNewLine
							Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
						Response.Write "else" & vbNewLine
							Response.Write "bDuplicated = true;" & vbNewLine
					Response.Write "}" & vbNewLine
				Next
				Response.Write "if (bDuplicated) {" & vbNewLine
					Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
                Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
                    Response.Write "alert('No se especificó qué columna contiene el Número de Empleado.');" & vbNewLine
                    Response.Write "return false;" & vbNewLine
                Response.Write "}" & vbNewLine
		If False Then
	                Response.Write "if (sFields.search(/ConceptID/gi) == -1) {" & vbNewLine
        	                Response.Write "alert('No se especificó qué columna contiene la Clave del Concepto.');" & vbNewLine
                	        Response.Write "return false;" & vbNewLine	
	                Response.Write "}" & vbNewLine
		End If
                Response.Write "if ((sFields.search(/OcurredStartDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredStartDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredStartDateMMDDYYYY/gi) == -1)) {" & vbNewLine
                        Response.Write "alert('No se especificó qué columna contiene la fecha de inicio.');" & vbNewLine
                        Response.Write "return false;" & vbNewLine
                Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de inicio con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/EndDateYYYYMMDD/gi) == -1) && (sFields.search(/EndDateDDMMYYYY/gi) == -1) && (sFields.search(/EndDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha de fin de la prestación.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/EndDateYYYYMMDD/gi) != -1) && ((sFields.search(/EndDateDDMMYYYY/gi) != -1) || (sFields.search(/EndDateMMDDYYYY/gi) != -1))) || ((sFields.search(/EndDateDDMMYYYY/gi) != -1) && ((sFields.search(/EndDateYYYYMMDD/gi) != -1) || (sFields.search(/EndDateMMDDYYYY/gi) != -1))) || ((sFields.search(/EndDateMMDDYYYY/gi) != -1) && ((sFields.search(/EndDateDDMMYYYY/gi) != -1) || (sFields.search(/EndDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de fin con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
                Response.Write "if ((sFields.search(/PayrollDateYYYYMMDD/gi) == -1) && (sFields.search(/PayrollStartDateDDMMYYYY/gi) == -1) && (sFields.search(/PayrollDateMMDDYYYY/gi) == -1)) {" & vbNewLine
                        Response.Write "alert('No se especificó qué columna contiene la fecha de aplicación en nómina.');" & vbNewLine
                        Response.Write "return false;" & vbNewLine
                Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/PayrollDateYYYYMMDD/gi) != -1) && ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) || (sFields.search(/PayrollDateMMDDYYYY/gi) != -1))) || ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) && ((sFields.search(/PayrollDateYYYYMMDD/gi) != -1) || (sFields.search(/PayrollDateMMDDYYYY/gi) != -1))) || ((sFields.search(/PayrollDateMMDDYYYY/gi) != -1) && ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) || (sFields.search(/PayrollDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de aplicación en nómina.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
                Response.Write "if (sFields.search(/ConceptAmount/gi) == -1) {" & vbNewLine
                        Response.Write "alert('No se especificó qué columna contiene el Monto.');" & vbNewLine
                        Response.Write "return false;" & vbNewLine
                Response.Write "}" & vbNewLine
		If lReasonID <> EMPLOYEES_SAFE_SEPARATION Then
			Response.Write "if (sFields.search(/ConceptQttyID/gi) == -1) {" & vbNewLine
                        	Response.Write "alert('No se especificó qué columna contiene el tipo de cantidad ($ o %).');" & vbNewLine
	                        Response.Write "return false;" & vbNewLine
			Response.Write "}" & vbNewLine
                End If
			Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If
	DisplayEmployeesSafeSeparationColumns = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesAntiquitiesColumns(sFileName, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesAntiquitiesColumns"
	Dim iColumns
	Dim iIndex
	Dim lErrorNumber

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "Indique a qué campo pertenece cada columna del archivo.")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""UploadAbsencesFrm"" ID=""UploadAbsencesFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
			For iIndex = 1 To iColumns
				Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
				Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeID"">No. del empleado</OPTION>"
					Response.Write "<OPTION VALUE=""ConceptAmount"">Monto a pagar quincenalmente</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateYYYYMMDD"">Fecha de inicio (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateDDMMYYYY"">Fecha de inicio (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateMMDDYYYY"">Fecha de inicio (MM-DD-AAAA)</OPTION>"
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			Response.Write "<BR />"
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
				Response.Write "var bDuplicated = false;" & vbNewLine
				Response.Write "var sFields = '';" & vbNewLine

				For iIndex = 1 To iColumns
					Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
						Response.Write "if (sFields.search(eval('/' + oForm.Column" & iIndex & ".value + '/gi')) == -1)" & vbNewLine
							Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
						Response.Write "else" & vbNewLine
							Response.Write "bDuplicated = true;" & vbNewLine
					Response.Write "}" & vbNewLine
				Next

				Response.Write "if (bDuplicated) {" & vbNewLine
					Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene los números de los empleados.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/ConceptAmount/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el monto a pagar.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/OcurredDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha de inicio del pago.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/OcurredDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de inicio del pago con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	DisplayEmployeesAntiquitiesColumns = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesAssignNumberColumns(sFileName, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesAssignNumberColumns"
	Dim iColumns
	Dim iIndex
	Dim lErrorNumber

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "Indique a qué campo pertenece cada columna del archivo.")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""UploadAbsencesFrm"" ID=""UploadAbsencesFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
			For iIndex = 1 To iColumns
				Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
				Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeName"">Nombre</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeLastName"">Apellido paterno</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeLastName2"">Apellido materno</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeTypeID"">Clave tipo de tabulador</OPTION>"
					Response.Write "<OPTION VALUE=""RFC"">RFC</OPTION>"
					Response.Write "<OPTION VALUE=""CURP"">CURP</OPTION>"
					'Response.Write "<OPTION VALUE=""EmployeeEmail"">Correo electrónico</OPTION>"
					'Response.Write "<OPTION VALUE=""SocialSecurityNumber"">No. Seg. Social</OPTION>"
					'Response.Write "<OPTION VALUE=""CountryID"">Clave del país</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDocumentDateYYYYMMDD"">Fecha de nacimiento (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDocumentDateDDMMYYYY"">Fecha de nacimiento (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDocumentDateMMDDYYYY"">Fecha de nacimiento (MM-DD-AAAA)</OPTION>"
					'Response.Write "<OPTION VALUE=""GenderID"">Clave del género</OPTION>"
					'Response.Write "<OPTION VALUE=""MaritalStatusID"">Clave del estado civil</OPTION>"
					'Response.Write "<OPTION VALUE=""EmployeeAddress"">Domicilio</OPTION>"
					'Response.Write "<OPTION VALUE=""EmployeeCity"">Ciudad</OPTION>"
					'Response.Write "<OPTION VALUE=""EmployeeZipCode"">Código postal</OPTION>"
					'Response.Write "<OPTION VALUE=""StateID"">Estado</OPTION>"
					'Response.Write "<OPTION VALUE=""EmployeePhone"">Teléfono casa</OPTION>"
					'Response.Write "<OPTION VALUE=""OfficePhone"">Teléfono oficina</OPTION>"
					'Response.Write "<OPTION VALUE=""OfficeExt"">Extensión</OPTION>"
					'Response.Write "<OPTION VALUE=""DocumentNumber1"">Clave de elector</OPTION>"
					'Response.Write "<OPTION VALUE=""DocumentNumber2"">Cédula profesional</OPTION>"
					'Response.Write "<OPTION VALUE=""DocumentNumber3"">Cartilla del servicio militar</OPTION>"
					'Response.Write "<OPTION VALUE=""EmployeeActivityID"">Actividad</OPTION>"
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			Response.Write "<BR />"
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
				Response.Write "var bDuplicated = false;" & vbNewLine
				Response.Write "var sFields = '';" & vbNewLine

				For iIndex = 1 To iColumns
					Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
						Response.Write "if (sFields.search(eval('/' + oForm.Column" & iIndex & ".value + '/gi')) == -1)" & vbNewLine
							Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
						Response.Write "else" & vbNewLine
							Response.Write "bDuplicated = true;" & vbNewLine
					Response.Write "}" & vbNewLine
				Next

				Response.Write "if (bDuplicated) {" & vbNewLine
					Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/OcurredDocumentDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredDocumentDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredDocumentDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha de nacimiento.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/OcurredDocumentDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredDocumentDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDocumentDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDocumentDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredDocumentDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredDocumentDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDocumentDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredDocumentDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDocumentDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de nacimiento de los hijos con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/EmployeeTypeID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la clave del tipo de tabulador.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/EmployeeName/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el nombre.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/EmployeeLastName/gi) == -1) && (sFields.search(/EmployeeLastName2/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el apellido paterno o materno.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/RFC/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el RFC.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/CURP/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el CURP.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	DisplayEmployeesAssignNumberColumns = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesChildrenColumns(sFileName, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesChildrenColumns"
	Dim iColumns
	Dim iIndex
	Dim lErrorNumber

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "Indique a qué campo pertenece cada columna del archivo.")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""UploadAbsencesFrm"" ID=""UploadAbsencesFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
			For iIndex = 1 To iColumns
				Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
				Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeID"">No. del empleado</OPTION>"
					Response.Write "<OPTION VALUE=""ChildName"">Nombre del hijo(a)</OPTION>"
					Response.Write "<OPTION VALUE=""ChildLastName"">Apellido paterno del hijo(a)</OPTION>"
					Response.Write "<OPTION VALUE=""ChildLastName2"">Apellido materno del hijo(a)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateYYYYMMDD"">Fecha de nacimiento (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateDDMMYYYY"">Fecha de nacimiento (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateMMDDYYYY"">Fecha de nacimiento (MM-DD-AAAA)</OPTION>"
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			Response.Write "<BR />"
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
				Response.Write "var bDuplicated = false;" & vbNewLine
				Response.Write "var sFields = '';" & vbNewLine

				For iIndex = 1 To iColumns
					Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
						Response.Write "if (sFields.search(eval('/' + oForm.Column" & iIndex & ".value + '/gi')) == -1)" & vbNewLine
							Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
						Response.Write "else" & vbNewLine
							Response.Write "bDuplicated = true;" & vbNewLine
					Response.Write "}" & vbNewLine
				Next

				Response.Write "if (bDuplicated) {" & vbNewLine
					Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene los números de los empleados.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/ChildName/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el nombre del hijo(a).');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/ChildLastName/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el apellido paterno del hijo(a).');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/OcurredDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha de nacimiento de los hijos.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/OcurredDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de nacimiento de los hijos con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	DisplayEmployeesChildrenColumns = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesConceptColumns(sFileName, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesConceptColumns"
	Dim iColumns
	Dim iIndex
	Dim lErrorNumber

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "Indique a qué campo pertenece cada columna del archivo.")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""UploadEmployeesConceptFrm"" ID=""UploadEmployeesConceptFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
			For iIndex = 1 To iColumns
				Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
				Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeID"">Número de Empleado</OPTION>"
					Response.Write "<OPTION VALUE=""ConceptID"">Número de Concepto</OPTION>"                    
					Response.Write "<OPTION VALUE=""OcurredStartDateYYYYMMDD"">Fecha de inicio (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredStartDateDDMMYYYY"">Fecha de inicio (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredStartDateMMDDYYYY"">Fecha de inicio (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""ConceptAmount"">Monto</OPTION>"
					Response.Write "<OPTION VALUE=""ConceptMin"">Mínimo</OPTION>"
					Response.Write "<OPTION VALUE=""ConceptMax"">Máximo</OPTION>"
					Response.Write "<OPTION VALUE=""AbsenceTypeID"">Ausencias</OPTION>"
					Response.Write "<OPTION VALUE=""Active"">Activo</OPTION>"
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			Response.Write "<BR />"
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
				Response.Write "var bDuplicated = false;" & vbNewLine
				Response.Write "var sFields = '';" & vbNewLine
				For iIndex = 1 To iColumns
					Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
						Response.Write "if (sFields.search(eval('/' + oForm.Column" & iIndex & ".value + '/gi')) == -1)" & vbNewLine
							Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
						Response.Write "else" & vbNewLine
							Response.Write "bDuplicated = true;" & vbNewLine
					Response.Write "}" & vbNewLine
				Next
				Response.Write "if (bDuplicated) {" & vbNewLine
					Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
                Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
                    Response.Write "alert('No se especificó qué columna contiene número de empleado.');" & vbNewLine
                    Response.Write "return false;" & vbNewLine
                Response.Write "}" & vbNewLine
                Response.Write "if (sFields.search(/ConceptID/gi) == -1) {" & vbNewLine
                        Response.Write "alert('No se especificó qué columna contiene número de concepto.');" & vbNewLine
                        Response.Write "return false;" & vbNewLine
                Response.Write "}" & vbNewLine
                Response.Write "if ((sFields.search(/OcurredStartDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredStartDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredStartDateMMDDYYYY/gi) == -1)) {" & vbNewLine
                        Response.Write "alert('No se especificó qué columna contiene la fecha de inicio.');" & vbNewLine
                        Response.Write "return false;" & vbNewLine
                Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de inicio con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
                Response.Write "if (sFields.search(/ConceptAmount/gi) == -1) {" & vbNewLine
                        Response.Write "alert('No se especificó qué columna contiene el importe.');" & vbNewLine
                        Response.Write "return false;" & vbNewLine
                Response.Write "}" & vbNewLine
                Response.Write "if (sFields.search(/ConceptMin/gi) == -1) {" & vbNewLine
                        Response.Write "alert('No se especificó qué columna contiene el mínimo.');" & vbNewLine
                        Response.Write "return false;" & vbNewLine
                Response.Write "}" & vbNewLine
                Response.Write "if (sFields.search(/ConceptMax/gi) == -1) {" & vbNewLine
                        Response.Write "alert('No se especificó qué columna contiene el máximo.');" & vbNewLine
                        Response.Write "return false;" & vbNewLine
                Response.Write "}" & vbNewLine
                Response.Write "if (sFields.search(/AbsenceTypeID/gi) == -1) {" & vbNewLine
                        Response.Write "alert('No se especificó qué columna contiene la clave de ausencia.');" & vbNewLine
                        Response.Write "return false;" & vbNewLine
                Response.Write "}" & vbNewLine
                Response.Write "if (sFields.search(/Active/gi) == -1) {" & vbNewLine
                        Response.Write "alert('No se especificó qué columna contiene si es Activo.');" & vbNewLine
                        Response.Write "return false;" & vbNewLine
                Response.Write "}" & vbNewLine
			Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If
	DisplayEmployeesConceptColumns = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesExtraHoursColumns(sFileName, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesExtraHoursColumns"
	Dim iColumns
	Dim iIndex
	Dim lErrorNumber

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "Indique a qué campo pertenece cada columna del archivo.")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""UploadAbsencesFrm"" ID=""UploadAbsencesFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
			For iIndex = 1 To iColumns
				Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
				Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeID"">No. del empleado</OPTION>"
					Response.Write "<OPTION VALUE=""AbsenceHours"">Horas extras</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateYYYYMMDD"">Fecha de ocurrencia (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateDDMMYYYY"">Fecha de ocurrencia (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateMMDDYYYY"">Fecha de ocurrencia (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""DocumentNumber"">No. de documento</OPTION>"
					Response.Write "<OPTION VALUE=""Reasons"">Observaciones</OPTION>"
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			Response.Write "<BR />"
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
				Response.Write "var bDuplicated = false;" & vbNewLine
				Response.Write "var sFields = '';" & vbNewLine

				For iIndex = 1 To iColumns
					Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
						Response.Write "if (sFields.search(eval('/' + oForm.Column" & iIndex & ".value + '/gi')) == -1)" & vbNewLine
							Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
						Response.Write "else" & vbNewLine
							Response.Write "bDuplicated = true;" & vbNewLine
					Response.Write "}" & vbNewLine
				Next

				Response.Write "if (bDuplicated) {" & vbNewLine
					Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene los números de los empleados.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/AbsenceHours/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el número de horas extras.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/OcurredDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha de la incidencia.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/OcurredDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de nacimiento de los hijos con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	DisplayEmployeesExtraHoursColumns = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesFeaturesColumns(lReasonID, sAction, sFileName, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesFeaturesColumns"
	Dim iColumns
	Dim iIndex
	Dim lErrorNumber
	Dim sTextAlert

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "Indique a qué campo pertenece cada columna del archivo.")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""UploadEmployeesFeaturesFrm"" ID=""UploadEmployeesFeaturesFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE="&lReasonID&" />"
			For iIndex = 1 To iColumns
				Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
				Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeID"">Número de empleado</OPTION>"
					If (lReasonID = -89) Or (lReasonID = EMPLOYEES_NON_EXCENT) Or (lReasonID = EMPLOYEES_EXCENT) Or (lReasonID = EMPLOYEES_NIGHTSHIFTS) Or (lReasonID = EMPLOYEES_CHILDREN_SCHOOLARSHIPS) Or (lReasonID = EMPLOYEES_EXTRAHOURS) Or (lReasonID = EMPLOYEES_GLASSES) Or (lReasonID = EMPLOYEES_FAMILY_DEATH) Or (lReasonID = EMPLOYEES_PROFESSIONAL_DEGREE) Or (lReasonID = EMPLOYEES_MOTHERAWARD) Or (lReasonID = EMPLOYEES_ANTIQUITY_25_AND_30_YEARS) Then
						Response.Write "<OPTION VALUE=""StartDateYYYYMMDD"">Fecha (AAAAMMDD)</OPTION>"
						Response.Write "<OPTION VALUE=""StartDateDDMMYYYY"">Fecha (DD-MM-AAAA)</OPTION>"
						Response.Write "<OPTION VALUE=""StartDateMMDDYYYY"">Fecha (MM-DD-AAAA)</OPTION>"
					Else
						Response.Write "<OPTION VALUE=""StartDateYYYYMMDD"">Fecha de inicio (AAAAMMDD)</OPTION>"
						Response.Write "<OPTION VALUE=""StartDateDDMMYYYY"">Fecha de inicio (DD-MM-AAAA)</OPTION>"
						Response.Write "<OPTION VALUE=""StartDateMMDDYYYY"">Fecha de inicio (MM-DD-AAAA)</OPTION>"
						Response.Write "<OPTION VALUE=""EndDateYYYYMMDD"">Fecha de término (AAAAMMDD)</OPTION>"
						Response.Write "<OPTION VALUE=""EndDateDDMMYYYY"">Fecha de término (DD-MM-AAAA)</OPTION>"
						Response.Write "<OPTION VALUE=""EndDateMMDDYYYY"">Fecha de término (MM-DD-AAAA)</OPTION>"
					End If
					Response.Write "<OPTION VALUE=""PayrollDateYYYYMMDD"">Quincena de aplicación (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""PayrollDateDDMMYYYY"">Quincena de aplicación (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""PayrollDateMMDDYYYY"">Quincena de aplicación (MM-DD-AAAA)</OPTION>"
					Select Case lReasonID
						Case -89
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe de deducción no gravables</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de deducción no gravables"
						Case 53 ' EmployeesForRisk
							Response.Write "<OPTION VALUE=""ConceptAmount"">Porcentaje de la prestación</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el porcentaje de la prestación"
						Case EMPLOYEES_ANTIQUITIES
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe para la compensación por antigüedad</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de la diferencia con el puesto superior"
						Case EMPLOYEES_ADDITIONALSHIFT
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe por el turno opcional</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe por el turno opcional"
						Case EMPLOYEES_CONCEPT_08
							Response.Write "<OPTION VALUE=""ConceptAmount"">Percepcion Adicional</OPTION>"
							sTextAlert = "No se especificó qué columna contiene la Percepcion Adicional"
						Case EMPLOYEES_CONCEPT_16
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe de devolución por deducciones indebidas</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de devolución por deducciones indebidas"
						Case EMPLOYEES_CHILDREN_SCHOOLARSHIPS
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe de la beca de hijos de trabajadores</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de la beca de hijos de trabajadores"
						Case EMPLOYEES_GLASSES
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe de la ayuda de anteojos</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de la ayuda de anteojos"
						Case EMPLOYEES_ANUAL_AWARD
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe del estimulo</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe del estimulo"
						Case EMPLOYEES_FAMILY_DEATH
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe de la ayuda de muerte del familiar</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de la ayuda de muerte del familiar"
						Case EMPLOYEES_ANTIQUITY_25_AND_30_YEARS
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe de concepto 41</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de concepto 41"
						Case EMPLOYEES_PROFESSIONAL_DEGREE
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe de la ayuda para tesis</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de la ayuda para tesis"					
						Case EMPLOYEES_MONTHAWARD
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe del premio al trabajador del mes</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de la ayuda para tesis"
						Case EMPLOYEES_SPORTS_HELP
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe del apoyo al deporte</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe del apoyo al deporte"
						Case EMPLOYEES_SPORTS
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe de la cuota deportiva</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de la cuota deportiva"
						Case EMPLOYEES_BENEFICIARIES
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe para pensió alimenticia</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe para pensió alimenticia"
						Case EMPLOYEES_NON_EXCENT
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe de deducción por cobro de sueldos indebidos</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de deducción por cobro de sueldos indebidos"
						Case EMPLOYEES_EXCENT
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe para otras deducciones</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe para otras deducciones"
						Case EMPLOYEES_MOTHERAWARD
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe del Premio del 10 de Mayo</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe del Premio del 10 de Mayo"
						Case EMPLOYEES_BENEFICIARIES_DEBIT
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe para adeudo pensión alimenticia</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe para adeudo pensión alimenticia"
						Case EMPLOYEES_ADD_SAFE_SEPARATION
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe para seguro adicional de separación individualizado para el personal de mando</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe para seguro adicional de separación individualizado"
						Case EMPLOYEES_NIGHTSHIFTS
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe para jornada nocturna adicional por día festivo</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe para jornada nocturna adicional"
						Case EMPLOYEES_CONCEPT_C3
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe de la ayuda para tesis</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de la ayuda para tesis"
						Case EMPLOYEES_LICENSES
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe de retenciones por exceso de licencias médicas</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de retenciones por exceso de licencias médicas"
						Case EMPLOYEES_SAFE_SEPARATION
							Response.Write "<OPTION VALUE=""ConceptAmount"">Seguro de separación</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Seguro de separación"
						Case EMPLOYEES_CARLOAN
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe del préstamo automóvil servidores públicos de mando superior</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe del préstamo automóvil servidores públicos de mando superior"
					End Select
					Response.Write "<OPTION VALUE=""ConceptComments"">Observaciones</OPTION>"
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			Response.Write "<BR />"
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
				Response.Write "var bDuplicated = false;" & vbNewLine
				Response.Write "var sFields = '';" & vbNewLine
				For iIndex = 1 To iColumns
					Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
						Response.Write "if (sFields.search(eval('/' + oForm.Column" & iIndex & ".value + '/gi')) == -1)" & vbNewLine
							Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
						Response.Write "else" & vbNewLine
							Response.Write "bDuplicated = true;" & vbNewLine
					Response.Write "}" & vbNewLine
				Next
				Response.Write "if (bDuplicated) {" & vbNewLine
					Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el número de empleado.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/StartDateYYYYMMDD/gi) == -1) && (sFields.search(/StartDateDDMMYYYY/gi) == -1) && (sFields.search(/StartDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha de inicio.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/StartDateYYYYMMDD/gi) != -1) && ((sFields.search(/StartDateDDMMYYYY/gi) != -1) || (sFields.search(/StartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/StartDateDDMMYYYY/gi) != -1) && ((sFields.search(/StartDateYYYYMMDD/gi) != -1) || (sFields.search(/StartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/StartDateMMDDYYYY/gi) != -1) && ((sFields.search(/StartDateDDMMYYYY/gi) != -1) || (sFields.search(/StartDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de inicio con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				If (lReasonID <> -89) And (lReasonID <> EMPLOYEES_NON_EXCENT) And (lReasonID <> EMPLOYEES_EXCENT) And (lReasonID = EMPLOYEES_NIGHTSHIFTS) And (lReasonID = EMPLOYEES_CHILDREN_SCHOOLARSHIPS) And (lReasonID = EMPLOYEES_EXTRAHOURS) And (lReasonID = EMPLOYEES_GLASSES) And (lReasonID = EMPLOYEES_FAMILY_DEATH) And (lReasonID = EMPLOYEES_PROFESSIONAL_DEGREE) And (lReasonID = EMPLOYEES_MOTHERAWARD) And (lReasonID = EMPLOYEES_ANTIQUITY_25_AND_30_YEARS) Then
					Response.Write "if ((sFields.search(/EndDateYYYYMMDD/gi) == -1) && (sFields.search(/EndDateDDMMYYYY/gi) == -1) && (sFields.search(/EndDateMMDDYYYY/gi) == -1)) {" & vbNewLine
						Response.Write "alert('No se especificó qué columna contiene la fecha de término.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (((sFields.search(/EndDateYYYYMMDD/gi) != -1) && ((sFields.search(/EndDateDDMMYYYY/gi) != -1) || (sFields.search(/EndDateMMDDYYYY/gi) != -1))) || ((sFields.search(/EndDateDDMMYYYY/gi) != -1) && ((sFields.search(/EndDateYYYYMMDD/gi) != -1) || (sFields.search(/EndDateMMDDYYYY/gi) != -1))) || ((sFields.search(/EndDateMMDDYYYY/gi) != -1) && ((sFields.search(/EndDateDDMMYYYY/gi) != -1) || (sFields.search(/EndDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
						Response.Write "alert('No puede seleccionar más de una vez la fecha de término con diferente formato.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
				End If
				Response.Write "if ((sFields.search(/PayrollDateYYYYMMDD/gi) == -1) && (sFields.search(/PayrollDateDDMMYYYY/gi) == -1) && (sFields.search(/PayrollDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la quincena de aplicación.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/PayrollDateYYYYMMDD/gi) != -1) && ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) || (sFields.search(/PayrollDateMMDDYYYY/gi) != -1))) || ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) && ((sFields.search(/PayrollDateYYYYMMDD/gi) != -1) || (sFields.search(/PayrollDateMMDDYYYY/gi) != -1))) || ((sFields.search(/PayrollDateMMDDYYYY/gi) != -1) && ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) || (sFields.search(/PayrollDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la quincena de aplicación con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
					Response.Write "if (sFields.search(/ConceptAmount/gi) == -1) {" & vbNewLine
						Response.Write "alert('" & sTextAlert & "');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If
	DisplayEmployeesFeaturesColumns = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesLicensesColumns(sFileName, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesLicensesColumns"
	Dim iColumns
	Dim iIndex
	Dim lErrorNumber

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "Indique a qué campo pertenece cada columna del archivo.")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""UploadAbsencesFrm"" ID=""UploadAbsencesFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
			For iIndex = 1 To iColumns
				Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
				Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeID"">No. del empleado</OPTION>"
					Response.Write "<OPTION VALUE=""LicenseTypeID"">Tipo de licencia (0-Sin goce de sueldo, 1-Con goce de sueldo</OPTION>"
					Response.Write "<OPTION VALUE=""AbsenceID"">Clave de incidencia</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateStartYYYYMMDD"">Fecha de inicio (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateStartDDMMYYYY"">Fecha de inicio (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateStartMMDDYYYY"">Fecha de inicio (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateEndYYYYMMDD"">Fecha de fin (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateEndDDMMYYYY"">Fecha de fin (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateEndMMDDYYYY"">Fecha de fin licencia (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""DocumentNumber"">No. de documento</OPTION>"
					Response.Write "<OPTION VALUE=""Reasons"">Observaciones</OPTION>"
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			Response.Write "<BR />"
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
				Response.Write "var bDuplicated = false;" & vbNewLine
				Response.Write "var sFields = '';" & vbNewLine

				For iIndex = 1 To iColumns
					Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
						Response.Write "if (sFields.search(eval('/' + oForm.Column" & iIndex & ".value + '/gi')) == -1)" & vbNewLine
							Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
						Response.Write "else" & vbNewLine
							Response.Write "bDuplicated = true;" & vbNewLine
					Response.Write "}" & vbNewLine
				Next

				Response.Write "if (bDuplicated) {" & vbNewLine
					Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene los números de los empleados.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/LicenseTypeID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene los números de los empleados.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/AbsenceID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la clave de incidencia.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/OcurredDateStartYYYYMMDD/gi) == -1) && (sFields.search(/OcurredDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha de inicio de la licencia.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/OcurredDateStartYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredDateStartDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateStartMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateStartDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredDateStartYYYYMMDD/gi) != -1) || (sFields.search(/OcurredDateStartMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateStartMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredDateStartDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateStartYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de inicio de la licencia con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/OcurredDateEndYYYYMMDD/gi) == -1) && (sFields.search(/OcurredDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha de fin de la licencia.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/OcurredDateEndYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredDateEndDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateEndMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateEndDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredDateEndYYYYMMDD/gi) != -1) || (sFields.search(/OcurredDateEndMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateEndMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredDateEndDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateEndYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de fin de la licencia con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	DisplayEmployeesLicensesColumns = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesNightShiftColumns(sFileName, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesNightShiftColumns"
	Dim iColumns
	Dim iIndex
	Dim lErrorNumber

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "Indique a qué campo pertenece cada columna del archivo.")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""UploadAbsencesFrm"" ID=""UploadAbsencesFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
			For iIndex = 1 To iColumns
				Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
				Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeID"">No. del empleado</OPTION>"
					Response.Write "<OPTION VALUE=""AbsenceID"">Tipo</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateYYYYMMDD"">Fecha de ocurrencia (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateDDMMYYYY"">Fecha de ocurrencia (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateMMDDYYYY"">Fecha de ocurrencia (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""DocumentNumber"">No. de oficio</OPTION>"
					Response.Write "<OPTION VALUE=""Reasons"">Observaciones</OPTION>"
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			Response.Write "<BR />"
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
				Response.Write "var bDuplicated = false;" & vbNewLine
				Response.Write "var sFields = '';" & vbNewLine

				For iIndex = 1 To iColumns
					Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
						Response.Write "if (sFields.search(eval('/' + oForm.Column" & iIndex & ".value + '/gi')) == -1)" & vbNewLine
							Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
						Response.Write "else" & vbNewLine
							Response.Write "bDuplicated = true;" & vbNewLine
					Response.Write "}" & vbNewLine
				Next

				Response.Write "if (bDuplicated) {" & vbNewLine
					Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene los números de los empleados.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/AbsenceID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene los tipos de ausencia.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/OcurredDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha de las ausencias.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/OcurredDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de las ausencias con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	DisplayEmployeesNightShiftColumns = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesSundaysColumns(sFileName, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesSundaysColumns"
	Dim iColumns
	Dim iIndex
	Dim lErrorNumber

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "Indique a qué campo pertenece cada columna del archivo.")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""UploadAbsencesFrm"" ID=""UploadAbsencesFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
			For iIndex = 1 To iColumns
				Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
				Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeID"">No. del empleado</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateYYYYMMDD"">Fecha de ocurrencia (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateDDMMYYYY"">Fecha de ocurrencia (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateMMDDYYYY"">Fecha de ocurrencia (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""DocumentNumber"">No. de documento</OPTION>"
					Response.Write "<OPTION VALUE=""Reasons"">Observaciones</OPTION>"
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			Response.Write "<BR />"
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
				Response.Write "var bDuplicated = false;" & vbNewLine
				Response.Write "var sFields = '';" & vbNewLine

				For iIndex = 1 To iColumns
					Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
						Response.Write "if (sFields.search(eval('/' + oForm.Column" & iIndex & ".value + '/gi')) == -1)" & vbNewLine
							Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
						Response.Write "else" & vbNewLine
							Response.Write "bDuplicated = true;" & vbNewLine
					Response.Write "}" & vbNewLine
				Next

				Response.Write "if (bDuplicated) {" & vbNewLine
					Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene los números de los empleados.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/OcurredDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha de la incidencia.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/OcurredDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de la incidencia con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	DisplayEmployeesSundaysColumns = lErrorNumber
	Err.Clear
End Function

Function DisplayJobsColumns(sFileName, sAction, lReasonID, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
        On Error Resume Next
        Const S_FUNCTION_NAME = "DisplayJobsColumns"
        Dim iColumns
        Dim iIndex
        Dim lErrorNumber

        Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "Indique a qué campo pertenece cada columna del archivo. <BR /> * Información requerida.")
        Response.Write "<BR />"
        lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
        If lErrorNumber = 0 Then
                Response.Write "<FORM NAME=""UploadEmployeesRequirementsFM1Frm"" ID=""UploadEmployeesRequirementsFM1Frm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
                        Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
                        Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
                        Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE="&lReasonID&" />"
                        For iIndex = 1 To iColumns
                                Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
                                Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
									Select Case lReasonID
										Case 54
											Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
											Response.Write "<OPTION VALUE=""JobID"">Número de plaza*</OPTION>"
											Response.Write "<OPTION VALUE=""ServiceID"">Clave del servicio*</OPTION>"
										Case 60
											Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
											Response.Write "<OPTION VALUE=""JobID"">Número de plaza</OPTION>"
											Response.Write "<OPTION VALUE=""AreaID"">Centro de trabajo</OPTION>"
											Response.Write "<OPTION VALUE=""PaymentCenterID"">Centro de pago</OPTION>"
											Response.Write "<OPTION VALUE=""ServiceID"">Clave de servicio</OPTION>"
											Response.Write "<OPTION VALUE=""ShiftID"">Clave de horario</OPTION>"
											Response.Write "<OPTION VALUE=""JourneyID"">Clave de turno</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateYYYYMMDD"">Fecha de inicio de vigencia* (AAAAMMDD)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateDDMMYYYY"">Fecha de inicio de vigencia* (DD-MM-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateMMDDYYYY"">Fecha de inicio de vigencia* (MM-DD-AAAA)</OPTION>"
										Case 61
											Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
											Response.Write "<OPTION VALUE=""JobID"">Número de plaza</OPTION>"
											Response.Write "<OPTION VALUE=""PositionID"">Clave del puesto</OPTION>"
											Response.Write "<OPTION VALUE=""LevelID"">Clave del nivel</OPTION>"
											Response.Write "<OPTION VALUE=""GroupGradeLevelID"">Clave de GGN</OPTION>"
											Response.Write "<OPTION VALUE=""IntegrationID"">Clave de integración</OPTION>"
											Response.Write "<OPTION VALUE=""ClassificationID"">Clave de clasificación</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateYYYYMMDD"">Fecha de inicio de vigencia* (AAAAMMDD)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateDDMMYYYY"">Fecha de inicio de vigencia* (DD-MM-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateMMDDYYYY"">Fecha de inicio de vigencia* (MM-DD-AAAA)</OPTION>"
										Case Else
											Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
											Response.Write "<OPTION VALUE=""AreaID"">Centro de trabajo</OPTION>"
											Response.Write "<OPTION VALUE=""PaymentCenterID"">Centro de pago</OPTION>"
											Response.Write "<OPTION VALUE=""ServiceID"">Clave de servicio</OPTION>"
											Response.Write "<OPTION VALUE=""PositionID"">Clave del puesto</OPTION>"
											Response.Write "<OPTION VALUE=""LevelID"">Clave del nivel</OPTION>"
											Response.Write "<OPTION VALUE=""GroupGradeLevelID"">Clave de GGN</OPTION>"
											Response.Write "<OPTION VALUE=""IntegrationID"">Clave de integración</OPTION>"
											Response.Write "<OPTION VALUE=""ClassificationID"">Clave de clasificación</OPTION>"
											Response.Write "<OPTION VALUE=""JobTypeID"">Clave del tipo de ocupación</OPTION>"
											Response.Write "<OPTION VALUE=""ShiftID"">Clave de horario</OPTION>"
											Response.Write "<OPTION VALUE=""JourneyID"">Clave de turno</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateYYYYMMDD"">Fecha de inicio de vigencia* (AAAAMMDD)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateDDMMYYYY"">Fecha de inicio de vigencia* (DD-MM-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateMMDDYYYY"">Fecha de inicio de vigencia* (MM-DD-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndDateYYYYMMDD"">Fecha de fin de vigencia (AAAAMMDD)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndDateDDMMYYYY"">Fecha de fin de vigencia (DD-MM-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndDateMMDDYYYY"">Fecha de fin de vigencia (MM-DD-AAAA)</OPTION>"
									End Select
                                Response.Write "</SELECT>"
                                Response.Write "<BR />"
                        Next
                        Response.Write "<BR />"
                        Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
                Response.Write "</FORM>"
                Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
                        Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
                                Response.Write "var bDuplicated = false;" & vbNewLine
                                Response.Write "var sFields = '';" & vbNewLine
                                For iIndex = 1 To iColumns
                                        Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
                                                Response.Write "if (sFields.search(eval('/' + oForm.Column" & iIndex & ".value + '/gi')) == -1)" & vbNewLine
                                                        Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
                                                Response.Write "else" & vbNewLine
                                                        Response.Write "bDuplicated = true;" & vbNewLine
                                        Response.Write "}" & vbNewLine
                                Next
                                Response.Write "if (bDuplicated) {" & vbNewLine
                                        Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
                                        Response.Write "return false;" & vbNewLine
                                Response.Write "}" & vbNewLine
								Select Case lReasonID
									Case 54
											Response.Write "if (sFields.search(/JobID/gi) == -1) {" & vbNewLine
											    Response.Write "alert('No se especificó qué columna contiene el número de plaza.');" & vbNewLine
											    Response.Write "return false;" & vbNewLine
											Response.Write "}" & vbNewLine
											Response.Write "if (sFields.search(/ServiceID/gi) == -1) {" & vbNewLine
											        Response.Write "alert('No se especificó qué columna contiene el servicio.');" & vbNewLine
											        Response.Write "return false;" & vbNewLine
											Response.Write "}" & vbNewLine
									Case 60
										Response.Write "if (sFields.search(/JobID/gi) == -1) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene el número de plaza.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if ((sFields.search(/OcurredStartDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredStartDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredStartDateMMDDYYYY/gi) == -1)) {" & vbNewLine
											Response.Write "alert('No se especificó qué columna contiene la fecha de inicio de vigencia de la plaza.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (((sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
											Response.Write "alert('No puede seleccionar más de una vez la fecha de inicio de vigencia de la plaza con diferente formato.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if ((sFields.search(/AreaID/gi) == -1) && (sFields.search(/PaymentCenterID/gi) == -1) && (sFields.search(/ServiceID/gi) == -1) && (sFields.search(/ShiftID/gi) == -1) && (sFields.search(/JourneyID/gi) == -1)) {" & vbNewLine
										Response.Write "alert('No se especificó qué columna contiene al menos un dato de la plaza a cambiar.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									Case 61
										Response.Write "if (sFields.search(/JobID/gi) == -1) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene el número de plaza.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if ((sFields.search(/OcurredStartDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredStartDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredStartDateMMDDYYYY/gi) == -1)) {" & vbNewLine
											Response.Write "alert('No se especificó qué columna contiene la fecha de inicio de vigencia de la plaza.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (((sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
											Response.Write "alert('No puede seleccionar más de una vez la fecha de inicio de vigencia de la plaza con diferente formato.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (sFields.search(/PositionID/gi) == -1) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene la clave del puesto.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if ((sFields.search(/LevelID/gi) == -1) && (sFields.search(/GroupGradeLevelID/gi) == -1)) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene el nivel o el GGN.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if ((sFields.search(/GroupGradeLevelID/gi) != -1) && (sFields.search(/IntegrationID/gi) == -1) && (sFields.search(/ClassificationID/gi) == -1)) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene la clave de integración y clasificación.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									Case Else
										Response.Write "if (sFields.search(/AreaID/gi) == -1) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene el centro de trabajo.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (sFields.search(/PaymentCenterID/gi) == -1) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene el centro de pago.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (sFields.search(/PositionID/gi) == -1) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene la clave del puesto.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if ((sFields.search(/LevelID/gi) == -1) && (sFields.search(/GroupGradeLevelID/gi) == -1)) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene el nivel o el GGN.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if ((sFields.search(/GroupGradeLevelID/gi) != -1) && (sFields.search(/IntegrationID/gi) == -1) && (sFields.search(/ClassificationID/gi) == -1)) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene la clave de integración y clasificación.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (sFields.search(/JobTypeID/gi) == -1) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene la clave del tipo de ocupación.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (sFields.search(/ShiftID/gi) == -1) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene la clave de horario.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (sFields.search(/JourneyID/gi) == -1) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene la clave de turno.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (sFields.search(/ServiceID/gi) == -1) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene la clave de servicio.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if ((sFields.search(/OcurredStartDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredStartDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredStartDateMMDDYYYY/gi) == -1)) {" & vbNewLine
											Response.Write "alert('No se especificó qué columna contiene la fecha de inicio de vigencia de la plaza.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (((sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
											Response.Write "alert('No puede seleccionar más de una vez la fecha de inicio de vigencia de la plaza con diferente formato.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
								End Select
                        Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
                Response.Write "//--></SCRIPT>" & vbNewLine
        End If
        DisplayJobsColumns = lErrorNumber
        Err.Clear
End Function

Function DisplayMedicalAreasColumns(sFileName, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayMedicalAreasColumns"
	Dim iColumns
	Dim iIndex
	Dim lErrorNumber

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "Indique a qué campo pertenece cada columna del archivo.")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""UploadAbsencesFrm"" ID=""UploadAbsencesFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
			For iIndex = 1 To iColumns
				Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
				Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""MedicalAreasID"">Número de la fila</OPTION>"
					Response.Write "<OPTION VALUE=""CompanyID"">Clave empresa</OPTION>"
					Response.Write "<OPTION VALUE=""MedicalAreasTypeID"">Tipo de reporte UNIMED</OPTION>"
					Response.Write "<OPTION VALUE=""PositionID"">Puesto</OPTION>"
					Response.Write "<OPTION VALUE=""ServiceID"">Servicio</OPTION>"
					Response.Write "<OPTION VALUE=""ColumnNumber"">No. de Anexo</OPTION>"
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			Response.Write "<BR />"
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
				Response.Write "var bDuplicated = false;" & vbNewLine
				Response.Write "var sFields = '';" & vbNewLine

				For iIndex = 1 To iColumns
					Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
						Response.Write "if (sFields.search(eval('/' + oForm.Column" & iIndex & ".value + '/gi')) == -1)" & vbNewLine
							Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
						Response.Write "else" & vbNewLine
							Response.Write "bDuplicated = true;" & vbNewLine
					Response.Write "}" & vbNewLine
				Next

				Response.Write "if (bDuplicated) {" & vbNewLine
					Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/MedicalAreasID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el número de fila del excel.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/CompanyID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la clave de la empresa.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/MedicalAreasTypeID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el tipo de reporte UNIMED.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/PositionID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la clave del puesto.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/ServiceID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la clave del servicio.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/ColumnNumber/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el número del anexo.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	DisplayMedicalAreasColumns = lErrorNumber
	Err.Clear
End Function

Function DisplayNewEmployeesColumns(sFileName, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayNewEmployeesColumns"
	Dim iColumns
	Dim iIndex
	Dim lErrorNumber

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "Indique a qué campo pertenece cada columna del archivo.")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""UploadAbsencesFrm"" ID=""UploadAbsencesFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
			For iIndex = 1 To iColumns
				Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
				Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeName"">Nombre del empleado</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeLastName"">Apellido paterno</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeLastName2"">Apellido materno</OPTION>"
					Response.Write "<OPTION VALUE=""RFC"">RFC</OPTION>"
					Response.Write "<OPTION VALUE=""CURP"">CURP</OPTION>"
					Response.Write "<OPTION VALUE=""GenderID"">Género(0-Femenino, 1-Masculino)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateYYYYMMDD"">Fecha de ingreso (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateDDMMYYYY"">Fecha de ingreso (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateMMDDYYYY"">Fecha de ingreso (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""JobID"">Número de plaza</OPTION>"
					Response.Write "<OPTION VALUE=""AreaID"">Clave del servicio</OPTION>"
					Response.Write "<OPTION VALUE=""PaymentCenterID"">Clave del centro de pago</OPTION>"
					Response.Write "<OPTION VALUE=""LevelID"">Nivel del empleado</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeTypeShortName"">Tipo de empleado</OPTION>"
					Response.Write "<OPTION VALUE=""ConceptID"">Riesgos profesionales (S/N)</OPTION>"
					Response.Write "<OPTION VALUE=""DocumentNumber"">No. de oficio</OPTION>"
					Response.Write "<OPTION VALUE=""Reasons"">Observaciones</OPTION>"
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			Response.Write "<BR />"
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
				Response.Write "var bDuplicated = false;" & vbNewLine
				Response.Write "var sFields = '';" & vbNewLine

				For iIndex = 1 To iColumns
					Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
						Response.Write "if (sFields.search(eval('/' + oForm.Column" & iIndex & ".value + '/gi')) == -1)" & vbNewLine
							Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
						Response.Write "else" & vbNewLine
							Response.Write "bDuplicated = true;" & vbNewLine
					Response.Write "}" & vbNewLine
				Next

				Response.Write "if (bDuplicated) {" & vbNewLine
					Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/EmployeeName/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna que contiene el nombre del empleado.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/EmployeeLastName/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna que contiene el apellido paterno del empleado.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/EmployeeLastName2/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna que contiene el apellido materno del empleado.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/RFC/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna que contiene el RFC del empleado.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/CURP/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna que contiene el CURP del empleado.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/GenderID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna que contiene el género del empleado.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/JobID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna que contiene el número de plaza del empleado.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/PaymentCenterID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna que contiene la clave del centro de pago del empleado.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/EmployeeTypeShortName/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna que contiene el tipó del empleado del empleado.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/ConceptID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna que contiene si el empleado tiene riesgos profesionales.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/OcurredDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha de ingreso.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/OcurredDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de ingreso con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	DisplayNewEmployeesColumns = lErrorNumber
	Err.Clear
End Function

Function DisplayRegisterEmployeesColumns(sFileName, sAction, lReasonID, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
        On Error Resume Next
        Const S_FUNCTION_NAME = "DisplayRegisterEmployeesColumns"
        Dim iColumns
        Dim iIndex
        Dim lErrorNumber

        Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "Indique a qué campo pertenece cada columna del archivo. <BR /> * Información requerida.")
        Response.Write "<BR />"
        lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
        If lErrorNumber = 0 Then
                Response.Write "<FORM NAME=""UploadEmployeesRequirementsFM1Frm"" ID=""UploadEmployeesRequirementsFM1Frm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
                        Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
                        Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
                        Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE="&lReasonID&" />"
                        For iIndex = 1 To iColumns
                                Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
                                Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
									Select Case lReasonID
										Case 1,5,6,10,2,4,8,3
											Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
											Response.Write "<OPTION VALUE=""EmployeeID"">No. de empleado*</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateYYYYMMDD"">Fecha de baja* (AAAAMMDD)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateDDMMYYYY"">Fecha de baja* (DD-MM-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateMMDDYYYY"">Fecha de baja* (MM-DD-AAAA)</OPTION>"
										Case 12
											Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
											Response.Write "<OPTION VALUE=""EmployeeID"">No. de empleado*</OPTION>"
											Response.Write "<OPTION VALUE=""JobID"">Número de plaza*</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateYYYYMMDD"">Fecha de inicio de vigencia* (AAAAMMDD)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateDDMMYYYY"">Fecha de inicio de vigencia* (DD-MM-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateMMDDYYYY"">Fecha de inicio de vigencia* (MM-DD-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndDateYYYYMMDD"">Fecha de fin de vigencia (AAAAMMDD)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndDateDDMMYYYY"">Fecha de fin de vigencia (DD-MM-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndDateMMDDYYYY"">Fecha de fin de vigencia (MM-DD-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartHour1HHMM"">Hora de entrada 1* (HHMM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartHour1HH_MM"">Hora de entrada 1* (HH:MM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndHour1HHMM"">Hora de salida 1* (HHMM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndHour1HH_MM"">Hora de salida 1* (HH:MM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartHour2HHMM"">Hora de entrada 2 (HHMM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartHour2HH_MM"">Hora de entrada 2 (HH:MM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndHour2HHMM"">Hora de salida 2 (HHMM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndHour2HH_MM"">Hora de salida 2 (HH:MM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartHour3HHMM"">Hora de entrada T/A (HHMM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartHour3HH_MM"">Hora de entrada T/A (HH:MM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndHour3HHMM"">Hora de salida T/A (HHMM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndHour3HH_MM"">Hora de salida T/A (HH:MM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""RiskLevel"">Nivel de riesgo (0,1,2)</OPTION>"
										Case 13
											Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
											Response.Write "<OPTION VALUE=""EmployeeID"">No. de empleado*</OPTION>"
											Response.Write "<OPTION VALUE=""JobID"">Número de plaza*</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateYYYYMMDD"">Fecha de inicio de vigencia* (AAAAMMDD)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateDDMMYYYY"">Fecha de inicio de vigencia* (DD-MM-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateMMDDYYYY"">Fecha de inicio de vigencia* (MM-DD-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndDateYYYYMMDD"">Fecha de fin de vigencia* (AAAAMMDD)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndDateDDMMYYYY"">Fecha de fin de vigencia* (DD-MM-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndDateMMDDYYYY"">Fecha de fin de vigencia* (MM-DD-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""ServiceID"">Clave del servicio</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartHour1HHMM"">Hora de entrada 1* (HHMM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartHour1HH_MM"">Hora de entrada 1* (HH:MM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndHour1HHMM"">Hora de salida 1* (HHMM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndHour1HH_MM"">Hora de salida 1* (HH:MM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartHour2HHMM"">Hora de entrada 2 (HHMM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartHour2HH_MM"">Hora de entrada 2 (HH:MM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndHour2HHMM"">Hora de salida 2 (HHMM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndHour2HH_MM"">Hora de salida 2 (HH:MM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""RiskLevel"">Nivel de riesgo (0,1,2)</OPTION>"
										Case 14
											Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
											Response.Write "<OPTION VALUE=""EmployeeID"">No. de empleado*</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateYYYYMMDD"">Fecha de inicio de vigencia* (AAAAMMDD)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateDDMMYYYY"">Fecha de inicio de vigencia* (DD-MM-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateMMDDYYYY"">Fecha de inicio de vigencia* (MM-DD-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndDateYYYYMMDD"">Fecha de fin de vigencia (AAAAMMDD)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndDateDDMMYYYY"">Fecha de fin de vigencia (DD-MM-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndDateMMDDYYYY"">Fecha de fin de vigencia (MM-DD-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""AreaID"">Clave del centro de trabajo</OPTION>"
											Response.Write "<OPTION VALUE=""PaymentCenterID"">Clave del centro de pago</OPTION>"
											Response.Write "<OPTION VALUE=""ServiceID"">Clave del servicio</OPTION>"
											Response.Write "<OPTION VALUE=""JourneyID"">Clave del turno</OPTION>"
											Response.Write "<OPTION VALUE=""ShiftID"">Clave del horario</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartHour1HHMM"">Hora de entrada 1* (HHMM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartHour1HH_MM"">Hora de entrada 1* (HH:MM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndHour1HHMM"">Hora de salida 1* (HHMM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndHour1HH_MM"">Hora de salida 1* (HH:MM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartHour2HHMM"">Hora de entrada 2 (HHMM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartHour2HH_MM"">Hora de entrada 2 (HH:MM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndHour2HHMM"">Hora de salida 2 (HHMM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndHour2HH_MM"">Hora de salida 2 (HH:MM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""ConceptAmount"">Monto quincenal</OPTION>"
										Case 21
											Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
											Response.Write "<OPTION VALUE=""EmployeeID"">No. de empleado*</OPTION>"
											Response.Write "<OPTION VALUE=""JobID"">Número de plaza*</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateYYYYMMDD"">Fecha de inicio de vigencia* (AAAAMMDD)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateDDMMYYYY"">Fecha de inicio de vigencia* (DD-MM-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateMMDDYYYY"">Fecha de inicio de vigencia* (MM-DD-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndDateYYYYMMDD"">Fecha de fin de vigencia* (AAAAMMDD)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndDateDDMMYYYY"">Fecha de fin de vigencia* (DD-MM-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndDateMMDDYYYY"">Fecha de fin de vigencia* (MM-DD-AAAA)</OPTION>"
										Case 54
											Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
											Response.Write "<OPTION VALUE=""JobID"">Número de plaza*</OPTION>"
											Response.Write "<OPTION VALUE=""ServiceID"">Clave del servicio*</OPTION>"
										Case Else
											Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
											Response.Write "<OPTION VALUE=""EmployeeID"">No. de empleado*</OPTION>"
											Response.Write "<OPTION VALUE=""JobID"">Número de plaza*</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateYYYYMMDD"">Fecha de inicio de vigencia* (AAAAMMDD)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateDDMMYYYY"">Fecha de inicio de vigencia* (DD-MM-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartDateMMDDYYYY"">Fecha de inicio de vigencia* (MM-DD-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndDateYYYYMMDD"">Fecha de fin de vigencia (AAAAMMDD)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndDateDDMMYYYY"">Fecha de fin de vigencia (DD-MM-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndDateMMDDYYYY"">Fecha de fin de vigencia (MM-DD-AAAA)</OPTION>"
											Response.Write "<OPTION VALUE=""ServiceID"">Clave del servicio</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartHour1HHMM"">Hora de entrada 1* (HHMM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartHour1HH_MM"">Hora de entrada 1* (HH:MM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndHour1HHMM"">Hora de salida 1* (HHMM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndHour1HH_MM"">Hora de salida 1* (HH:MM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartHour2HHMM"">Hora de entrada 2 (HHMM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartHour2HH_MM"">Hora de entrada 2 (HH:MM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndHour2HHMM"">Hora de salida 2 (HHMM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndHour2HH_MM"">Hora de salida 2 (HH:MM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartHour3HHMM"">Hora de entrada T/A (HHMM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredStartHour3HH_MM"">Hora de entrada T/A (HH:MM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndHour3HHMM"">Hora de salida T/A (HHMM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""OcurredEndHour3HH_MM"">Hora de salida T/A (HH:MM) 24 hrs.</OPTION>"
											Response.Write "<OPTION VALUE=""RiskLevel"">Nivel de riesgo (0,1,2)</OPTION>"
									End Select
                                Response.Write "</SELECT>"
                                Response.Write "<BR />"
                        Next
                        Response.Write "<BR />"
                        Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
                Response.Write "</FORM>"
                Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
                        Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
                                Response.Write "var bDuplicated = false;" & vbNewLine
                                Response.Write "var sFields = '';" & vbNewLine
                                For iIndex = 1 To iColumns
                                        Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
                                                Response.Write "if (sFields.search(eval('/' + oForm.Column" & iIndex & ".value + '/gi')) == -1)" & vbNewLine
                                                        Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
                                                Response.Write "else" & vbNewLine
                                                        Response.Write "bDuplicated = true;" & vbNewLine
                                        Response.Write "}" & vbNewLine
                                Next
                                Response.Write "if (bDuplicated) {" & vbNewLine
                                        Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
                                        Response.Write "return false;" & vbNewLine
                                Response.Write "}" & vbNewLine
								Select Case lReasonID
									Case 1,5,6,10,2,4,8,3
										Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
										    Response.Write "alert('No se especificó qué columna contiene el número de empleado.');" & vbNewLine
										    Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if ((sFields.search(/OcurredStartDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredStartDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredStartDateMMDDYYYY/gi) == -1)) {" & vbNewLine
											Response.Write "alert('No se especificó qué columna contiene la fecha de baja del empleado.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									Case 13
										Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene el número de empleado.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (sFields.search(/JobID/gi) == -1) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene el número de plaza.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if ((sFields.search(/OcurredStartDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredStartDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredStartDateMMDDYYYY/gi) == -1)) {" & vbNewLine
											Response.Write "alert('No se especificó qué columna contiene la fecha de inicio de vigencia del movimiento.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (((sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
											Response.Write "alert('No puede seleccionar más de una vez la fecha de inicio de vigencia del movimiento con diferente formato.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if ((sFields.search(/OcurredEndDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredEndDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredEndDateMMDDYYYY/gi) == -1)) {" & vbNewLine
											Response.Write "alert('No se especificó qué columna contiene la fecha de fin de vigencia del movimiento.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (((sFields.search(/OcurredEndDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredEndDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredEndDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredEndDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredEndDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredEndDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredEndDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredEndDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredEndDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
											Response.Write "alert('No puede seleccionar más de una vez la fecha de fin de vigencia del movimiento con diferente formato.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if ((sFields.search(/OcurredStartHour1HHMM/gi) == -1) && (sFields.search(/OcurredStartHour1HH_MM/gi) == -1)) {" & vbNewLine
											Response.Write "alert('No se especificó qué columna contiene la hora de entrada del empleado.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if ((sFields.search(/OcurredEndHour1HHMM/gi) == -1) && (sFields.search(/OcurredEndHour1HH_MM/gi) == -1)) {" & vbNewLine
											Response.Write "alert('No se especificó qué columna contiene la hora de salida del empleado.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (sFields.search(/ServiceID/gi) == -1) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene la clave del servicio.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									Case 14
										Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene el número de empleado.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (sFields.search(/AreaID/gi) == -1) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene la clave del centro de trabajo.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (sFields.search(/PaymentCenterID/gi) == -1) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene la clave del centro de pago.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (sFields.search(/ServiceID/gi) == -1) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene la clave del servicio.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (sFields.search(/ConceptAmount/gi) == -1) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene el monto quincenal.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if ((sFields.search(/OcurredStartDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredStartDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredStartDateMMDDYYYY/gi) == -1)) {" & vbNewLine
											Response.Write "alert('No se especificó qué columna contiene la fecha de inicio de vigencia del movimiento.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (((sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
											Response.Write "alert('No puede seleccionar más de una vez la fecha de inicio de vigencia del movimiento con diferente formato.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if ((sFields.search(/OcurredEndDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredEndDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredEndDateMMDDYYYY/gi) == -1)) {" & vbNewLine
											Response.Write "alert('No se especificó qué columna contiene la fecha de fin de vigencia del movimiento.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (((sFields.search(/OcurredEndDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredEndDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredEndDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredEndDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredEndDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredEndDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredEndDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredEndDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredEndDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
											Response.Write "alert('No puede seleccionar más de una vez la fecha de fin de vigencia del movimiento con diferente formato.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if ((sFields.search(/OcurredStartHour1HHMM/gi) == -1) && (sFields.search(/OcurredStartHour1HH_MM/gi) == -1)) {" & vbNewLine
											Response.Write "alert('No se especificó qué columna contiene la hora de entrada del empleado.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if ((sFields.search(/OcurredEndHour1HHMM/gi) == -1) && (sFields.search(/OcurredEndHour1HH_MM/gi) == -1)) {" & vbNewLine
											Response.Write "alert('No se especificó qué columna contiene la hora de salida del empleado.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									Case 21
										Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene el número de empleado.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (sFields.search(/JobID/gi) == -1) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene el número de plaza.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if ((sFields.search(/OcurredStartDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredStartDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredStartDateMMDDYYYY/gi) == -1)) {" & vbNewLine
											Response.Write "alert('No se especificó qué columna contiene la fecha de inicio de vigencia del movimiento.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (((sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
											Response.Write "alert('No puede seleccionar más de una vez la fecha de inicio de vigencia del movimiento con diferente formato.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if ((sFields.search(/OcurredEndDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredEndDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredEndDateMMDDYYYY/gi) == -1)) {" & vbNewLine
											Response.Write "alert('No se especificó qué columna contiene la fecha de fin de vigencia del movimiento.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (((sFields.search(/OcurredEndDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredEndDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredEndDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredEndDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredEndDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredEndDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredEndDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredEndDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredEndDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
											Response.Write "alert('No puede seleccionar más de una vez la fecha de fin de vigencia del movimiento con diferente formato.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									Case 54
											Response.Write "if (sFields.search(/JobID/gi) == -1) {" & vbNewLine
											    Response.Write "alert('No se especificó qué columna contiene el número de plaza.');" & vbNewLine
											    Response.Write "return false;" & vbNewLine
											Response.Write "}" & vbNewLine
											Response.Write "if (sFields.search(/ServiceID/gi) == -1) {" & vbNewLine
											        Response.Write "alert('No se especificó qué columna contiene el servicio.');" & vbNewLine
											        Response.Write "return false;" & vbNewLine
											Response.Write "}" & vbNewLine
									Case Else
										Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene el número de empleado.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (sFields.search(/JobID/gi) == -1) {" & vbNewLine
										        Response.Write "alert('No se especificó qué columna contiene el número de plaza.');" & vbNewLine
										        Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if ((sFields.search(/OcurredStartDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredStartDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredStartDateMMDDYYYY/gi) == -1)) {" & vbNewLine
											Response.Write "alert('No se especificó qué columna contiene la fecha de inicio de vigencia del movimiento.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (((sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
											Response.Write "alert('No puede seleccionar más de una vez la fecha de inicio de vigencia del movimiento con diferente formato.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if ((sFields.search(/OcurredStartHour1HHMM/gi) == -1) && (sFields.search(/OcurredStartHour1HH_MM/gi) == -1)) {" & vbNewLine
											Response.Write "alert('No se especificó qué columna contiene la hora de entrada del empleado.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if ((sFields.search(/OcurredEndHour1HHMM/gi) == -1) && (sFields.search(/OcurredEndHour1HH_MM/gi) == -1)) {" & vbNewLine
											Response.Write "alert('No se especificó qué columna contiene la hora de salida del empleado.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
								End Select
                        Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
                Response.Write "//--></SCRIPT>" & vbNewLine
        End If
        DisplayRegisterEmployeesColumns = lErrorNumber
        Err.Clear
End Function

Function DisplayResumptionOfWorkColumns(sFileName, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayResumptionOfWorkColumns"
	Dim iColumns
	Dim iIndex
	Dim lErrorNumber

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "Indique a qué campo pertenece cada columna del archivo (*Información requerida).")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""UploadAbsencesFrm"" ID=""UploadAbsencesFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
			For iIndex = 1 To iColumns
				Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
				Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeID"">No. del empleado</OPTION>"
					Response.Write "<OPTION VALUE=""AbsenceID"">Reanunación</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateYYYYMMDD"">Fecha de reanudación (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateDDMMYYYY"">Fecha de reanudación (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateMMDDYYYY"">Fecha de reanudación (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""DocumentNumber"">No. de oficio</OPTION>"
					Response.Write "<OPTION VALUE=""Reasons"">Observaciones</OPTION>"
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			Response.Write "<BR />"
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
				Response.Write "var bDuplicated = false;" & vbNewLine
				Response.Write "var sFields = '';" & vbNewLine

				For iIndex = 1 To iColumns
					Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
						Response.Write "if (sFields.search(eval('/' + oForm.Column" & iIndex & ".value + '/gi')) == -1)" & vbNewLine
							Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
						Response.Write "else" & vbNewLine
							Response.Write "bDuplicated = true;" & vbNewLine
					Response.Write "}" & vbNewLine
				Next

				Response.Write "if (bDuplicated) {" & vbNewLine
					Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene los números de los empleados.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/OcurredDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha de reanudación.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/OcurredDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de reanudación con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	DisplayResumptionOfWorkColumns = lErrorNumber
	Err.Clear
End Function

Function DisplayUpdateEmployeeDataColumns(sFileName, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayUpdateEmployeeDataColumns"
	Dim iColumns
	Dim iIndex
	Dim lErrorNumber

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "Indique a qué campo pertenece cada columna del archivo.")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""UploadAbsencesFrm"" ID=""UploadAbsencesFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
			For iIndex = 1 To iColumns
				Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
				Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeID"">No. del empleado</OPTION>"
					Response.Write "<OPTION VALUE=""AbsenceID"">Tipo movimiento</OPTION>"
					Response.Write "<OPTION VALUE=""DocumentNumber"">No. de oficio</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeName"">Nombre del empleado</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeLastName"">Apellido paterno</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeLastName2"">Apellido materno</OPTION>"
					Response.Write "<OPTION VALUE=""RFC"">RFC</OPTION>"
					Response.Write "<OPTION VALUE=""CURP"">CURP</OPTION>"
					Response.Write "<OPTION VALUE=""CURP"">Clave centro de pago</OPTION>"
					Response.Write "<OPTION VALUE=""CURP"">Clave horario</OPTION>"
					Response.Write "<OPTION VALUE=""CURP"">Clave turno</OPTION>"
					Response.Write "<OPTION VALUE=""CURP"">Clave servicio</OPTION>"
					Response.Write "<OPTION VALUE=""CURP"">Clave jornada</OPTION>"
					Response.Write "<OPTION VALUE=""GenderID"">Género(0-Femenino, 1-Masculino)</OPTION>"
					Response.Write "<OPTION VALUE=""DocumentNumber"">No. de oficio</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateYYYYMMDD"">Fecha de movimiento (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateDDMMYYYY"">Fecha de movimiento (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateMMDDYYYY"">Fecha de movimiento (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""PaymentCenterID"">Clave del centro de pago</OPTION>"
					Response.Write "<OPTION VALUE=""AreaID"">Clave del servicio</OPTION>"
					Response.Write "<OPTION VALUE=""ConceptID"">Riesgos profesionales (S/N)</OPTION>"
					Response.Write "<OPTION VALUE=""Reasons"">Observaciones</OPTION>"
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			Response.Write "<BR />"
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
				Response.Write "var bDuplicated = false;" & vbNewLine
				Response.Write "var sFields = '';" & vbNewLine

				For iIndex = 1 To iColumns
					Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
						Response.Write "if (sFields.search(eval('/' + oForm.Column" & iIndex & ".value + '/gi')) == -1)" & vbNewLine
							Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
						Response.Write "else" & vbNewLine
							Response.Write "bDuplicated = true;" & vbNewLine
					Response.Write "}" & vbNewLine
				Next
				Response.Write "if (bDuplicated) {" & vbNewLine
					Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene los números de los empleados.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/OcurredDateYYYYMMDD/gi) == -1) && (sFields.search(/OcurredDateDDMMYYYY/gi) == -1) && (sFields.search(/OcurredDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la fecha del movimiento.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/OcurredDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha del movimiento con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	DisplayUpdateEmployeeDataColumns = lErrorNumber
	Err.Clear
End Function

Function DisplayUploadForm(sAction, lEmployeeTypeID, lReasonID)
'************************************************************
'Purpose: To display the TEXTAREA field and the FILE field to
'         upload the text data to upload
'Inputs:  sAction
'************************************************************
	On Error Resume Next
	Dim sRequiredFields
	Dim sNumber
	Dim sMessage
	Dim bShowSection3
	Const S_FUNCTION_NAME = "DisplayUploadForm"

	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		Response.Write "var bReady = false;" & vbNewLine
	Response.Write "//--></SCRIPT>" & vbNewLine

	Select Case sAction
		Case "FONAC", "JobServices", "MedicalAreas", "Third"
		Case "ThirdUploadMovements"
			Response.Write "<IMG SRC=""Images/Crcl1.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: ShowDisplay(document.all['AbsencesFormDiv']); if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };"">Seleccione el archivo para buscar los registros</A><BR /><BR />"
		Case Else
			If (lReasonID <> 54) And (lReasonID <> 60) And (lReasonID <> 61) Then
				Response.Write "<IMG SRC=""Images/Crcl1.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: ShowDisplay(document.all['AbsencesFormDiv']); if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };"">Deseo registrar la información en línea</A><BR /><BR />"
			End If
	End Select
	Response.Write "<DIV NAME=""AbsencesFormDiv"" ID=""AbsencesFormDiv"">"
		Select Case sAction
			Case "Third"
				'Call DisplayInstructionsMessage("Carga de registros solo a través de archivo", "Para cargar la información es necesario que la información del archivo tenga el formato correcto para el concepto. Una vez este seguro de eso, de clic en la sección 2 y siga las instrucciones.")
				Response.Write "<BR /><BR />"
			Case "ChildrenSchoolarships"
				lErrorNumber = DisplayEmployeesSearchForm(oRequest, oADODBConnection, GetASPFileName(""), False, sErrorDescription)
				If (Len(oRequest("DoSearch").Item) > 0) Or (Len(oRequest("Change").Item) > 0) Then
					lErrorNumber = GetEmployeeChildren(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					If Len(aEmployeeComponent(S_NAME_EMPLOYEE)) > 0 Then
						If aEmployeeComponent(N_ID_CHILD_EMPLOYEE) > -1 Then
							lErrorNumber = DisplayEmployeeChildrenTable(oRequest, oADODBConnection, "ChildrenSchoolarships", DISPLAY_NOTHING, True, False, aEmployeeComponent, sErrorDescription)
						Else
							sErrorDescription = "Este empleado no tiene dados de alta hijos. Acuda al área correspondiente con el acta de nacimiento para que sea dado de alta."
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						End If	
					Else
						sErrorDescription = "No existen registros que cumplan con los criterios de la búsqueda."
						Call DisplayErrorMessage("Búsqueda vacía", sErrorDescription)
					End If
				End If
				If (Len(oRequest("Change").Item) > 0) Or (Len(oRequest("Delete").Item) > 0) Then
					lErrorNumber = DisplayEmployeeChildForm(oRequest, oADODBConnection, GetASPFileName(""), "ChildrenSchoolarships", "", aEmployeeComponent, sErrorDescription)
					If lErrorNumber <> 0 Then
						Response.Write "<BR />"
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						sErrorDescription = ""
					End If
				End If
				Response.Write "<BR />"
			Case "ConceptsValues"
				aConceptComponent(N_RECORD_ID_CONCEPT) = -1
				aConceptComponent(N_ID_CONCEPT) = -1
				aConceptComponent(N_JOB_STATUS_ID_CONCEPT) = -1
				aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) = -1
				aConceptComponent(N_INTEGRATION_ID_CONCEPT) = -1
				aConceptComponent(N_JOURNEY_ID_CONCEPT) = -1
				aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) = -1
				aConceptComponent(N_SERVICE_ID_CONCEPT) = -1
				aConceptComponent(N_ANTIQUITY_ID_CONCEPT) = -1
				aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) = -1
				aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) = -1
				aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) = -1
				aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) = -1
				aConceptComponent(N_GENDER_ID_CONCEPT) = -1
				aConceptComponent(N_HAS_CHILDREN_CONCEPT) = 0
				aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) = -1
				aConceptComponent(N_HAS_SYNDICATE_CONCEPT) = -1
				aConceptComponent(D_CONCEPT_MIN_CONCEPT) = 0
				aConceptComponent(D_CONCEPT_MAX_CONCEPT) = 0
				aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) = -1
				aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) = 20100701
				aConceptComponent(S_SHORT_NAME_CONCEPT) = ""
				aConceptComponent(S_NAME_CONCEPT) = ""
				aConceptComponent(N_CONCEPT_TYPE_ID_CONCEPT) = 3
				aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) = -1
				Call DisplayConceptValuesForm(oRequest, oADODBConnection, "UploadInfo.asp", False, lEmployeeTypeID, aConceptComponent, sErrorDescription)
			Case "CreditFOVISSSTE"
				Call DisplayEmployeeConceptForm(oRequest, oADODBConnection, GetASPFileName(""), "Step=1&Action=CreditFOVISSSTE", "64", aEmployeeComponent, sErrorDescription)
			Case "DocumentsForLicenses"
				Call DisplayDocumentsForLicensesForm(oRequest, oADODBConnection, GetASPFileName(""), "Step=1&Action=DocumentsForLicenses", aEmployeeComponent, sErrorDescription)
			Case "EmployeesAssignJob"
				Call DisplayEmployeeForm(oRequest, oADODBConnection, "Employees.asp", "EmployeesAssignJob", ",0,", lReasonID, aEmployeeComponent, sErrorDescription)
			Case "EmployeesAdditionalShift"
				sRequiredFields = "Número de empleado, la fecha de registro, la fecha de aplicación en nómina y la cantidad por el turno opcional"
				Call DisplayEmployeeConceptForm(oRequest, oADODBConnection, GetASPFileName(""), "Step=1&Action=EmployeesAdditionalShift", "7", aEmployeeComponent, sErrorDescription)
			Case "EmployeesAdditionalCompensation"
				Call DisplayEmployeeConceptForm(oRequest, oADODBConnection, GetASPFileName(""), "Step=1&Action=EmployeesAdditionalCompensation", "8", aEmployeeComponent, sErrorDescription)
			Case "EmployeesAdjustments"
				Call DisplayEmployeeForm(oRequest, oADODBConnection, GetASPFileName(""), "EmployeesAdjustments", ",7,", lReasonID, aEmployeeComponent, sErrorDescription)
			Case "EmployeesAntiquities"
				sRequiredFields = "Número de empleado, la fecha de registro, la fecha de aplicación en nómina y la cantidad de la compensación."
				Call DisplayEmployeeConceptForm(oRequest, oADODBConnection, GetASPFileName(""), "Step=1&Action=EmployeesAntiquities", "5", aEmployeeComponent, sErrorDescription)
			Case "EmployeesAssignNumber"
				sRequiredFields = "Nombre del empleado, apellido paterno, apellido materno, clave del tipo de tabulador, RFC, CURP, fecha de nacimiento del empleado"
				Call DisplayEmployeeForm(oRequest, oADODBConnection, GetASPFileName(""), "EmployeesAssignNumber", ",1,", 0, aEmployeeComponent, sErrorDescription)
			Case "EmployeesAnualAward"
				sRequiredFields = "Número de empleado, la fecha de registro, la fecha de aplicación en nómina y la cantidad."
				Call DisplayEmployeeConceptForm(oRequest, oADODBConnection, GetASPFileName(""), "Step=1&Action=EmployeesAnualAward", "94", aEmployeeComponent, sErrorDescription)
			Case "EmployeesChildren"
				Call DisplayEmployeeChildForm(oRequest, oADODBConnection, GetASPFileName(""), "EmployeesChildren", "Step=1&Action=EmployeesChildren", aEmployeeComponent, sErrorDescription)
			Case "EmployeesCarLoan"
				sRequiredFields = "Número de empleado, la fecha de registro, la fecha de aplicación en nómina y la cantidad"
				Call DisplayEmployeeConceptForm(oRequest, oADODBConnection, GetASPFileName(""), "Step=1&Action=EmployeesCarLoan", "74", aEmployeeComponent, sErrorDescription)
			Case "EmployeesChanges"
				Call DisplayEmployeeForm(oRequest, oADODBConnection, GetASPFileName(""), "", ",0,", lReasonID, aEmployeeComponent, sErrorDescription)
			Case "EmployeesDrop"
				Call DisplayEmployeeForm(oRequest, oADODBConnection, "Employees.asp", "EmployeesDrop", ",0,", lReasonID, aEmployeeComponent, sErrorDescription)
			Case "EmployeesExtraHours"
				Call DisplayAbsenceForm(oRequest, oADODBConnection, GetASPFileName(""), "201", "", aAbsenceComponent, sErrorDescription)
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "ShowDisplay(document.all['AbsenceHoursDiv']);" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
			Case "EmployeesFamilyDeath"
				sRequiredFields = "Número de empleado, la fecha de registro, la fecha de aplicación en nómina y la cantidad de la ayuda"
				Call DisplayEmployeeConceptForm(oRequest, oADODBConnection, GetASPFileName(""), "Step=1&Action=EmployeesFamilyDeath", "45", aEmployeeComponent, sErrorDescription)
			Case "EmployeesForRisk"
				sRequiredFields = "Número de empleado, la fecha de registro, la fecha de aplicación en nómina y el porcentaje para la compensación"
				Call DisplayEmployeeConceptForm(oRequest, oADODBConnection, GetASPFileName(""), "Step=1&Action=EmployeesForRisk", "4", aEmployeeComponent, sErrorDescription)
			Case "EmployeesGlasses"
				sRequiredFields = "Número de empleado, la fecha de registro, la fecha de aplicación en nómina y la cantidad de la ayuda"
				Call DisplayEmployeeConceptForm(oRequest, oADODBConnection, GetASPFileName(""), "Step=1&Action=EmployeesGlasses", "24", aEmployeeComponent, sErrorDescription)
			Case "EmployeesLicenses"
				sRequiredFields = "Número de empleado, la fecha de registro, la fecha de aplicación en nómina y el porcentaje"
				Call DisplayEmployeeForm(oRequest, oADODBConnection, "Employees.asp", "EmployeesLicenses", ",0,", lReasonID, aEmployeeComponent, sErrorDescription)
			Case "EmployeeMonthAward"
				sRequiredFields = "Número de empleado, la fecha de registro, la fecha de aplicación en nómina y la cantidad por el premio"
				Call DisplayEmployeeConceptForm(oRequest, oADODBConnection, GetASPFileName(""), "Step=1&Action=EmployeeMonthAward", "50", aEmployeeComponent, sErrorDescription)
			Case "EmployeesMovements"
				Select Case lReasonID
					Case EMPLOYEES_SAFE_SEPARATION, 53, EMPLOYEES_CONCEPT_08, EMPLOYEES_ADDITIONALSHIFT
						sRequiredFields = "Número de empleado, fecha de inicio, fecha de fin, quincena de aplicación, porcentaje y observaciones"
					Case EMPLOYEES_ANTIQUITIES, EMPLOYEES_MONTHAWARD, EMPLOYEES_CARLOAN, EMPLOYEES_CONCEPT_C3, EMPLOYEES_LICENSES, EMPLOYEES_CONCEPT_16, EMPLOYEES_HELP_COMISSION, EMPLOYEES_SAFEDOWN, EMPLOYEES_ANUAL_AWARD, EMPLOYEES_FONAC_CONCEPT
						sRequiredFields = "Número de empleado, fecha de inicio, fecha de fin, quincena de aplicación, importe y observaciones"
					Case -89, EMPLOYEES_NON_EXCENT, EMPLOYEES_EXCENT, EMPLOYEES_NIGHTSHIFTS, EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_GLASSES, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_MOTHERAWARD, EMPLOYEES_ANTIQUITY_25_AND_30_YEARS
						sRequiredFields = "Número de empleado, fecha, quincena de aplicación, importe y observaciones"
					Case EMPLOYEES_SPORTS_HELP, EMPLOYEES_SPORTS
						sRequiredFields = "Número de empleado, fecha de inicio, fecha de fin, quincena de aplicación, importe (con valor 0) y observaciones"
					Case EMPLOYEES_ADD_SAFE_SEPARATION
						sRequiredFields = "Número de empleado, fecha de inicio, fecha de fin, quincena de aplicación, cantidad, tipo de unidad de la cantidad ($ o %) y observaciones"
					Case EMPLOYEES_DOCUMENTS_FOR_LICENSES
						sRequiredFields = "No. del empleado, Fecha del documento, Número de la solicitud, Tipo de licencia, Fecha inicio de la licencia sindical, fecha fin de la licencia sindical, Nombre de la plantilla"
					Case -58
						sRequiredFields = "Número de empleado, clave del concepto, importe reclamado, fecha de omisión, quincena de aplicación y nombre del beneficiario (opcional)"
					Case 1,2,3,4,5,6,8,10
						sRequiredFields = "Número de empleado, fecha de baja"
					Case 12
						sRequiredFields = "Número de empleado, número de plaza, fecha de inicio vigencia, fecha de fin de vigencia (opcional), Horario de entrada 1, Horario de salida 1, horario de entrada 2 (opc), horario de salida 2 (opc), horario turno opcional 1, horario salida turno opcional, nivel de riesgo"
					Case 13
						sRequiredFields = "Número de empleado, número de plaza, fecha de inicio vigencia, fecha de fin de vigencia, Horario de entrada 1, Horario de salida 1, horario de entrada 2 (opc), horario de salida 2 (opc), clave del servicio"
					Case 14
						sRequiredFields = "Número de empleado, fecha de inicio vigencia, fecha de fin de vigencia, clave de centro de trabajo, clave de centro de pago, clave de servicio, clave de turno, clave de horario, monto quincenal"
					Case 17,18
						sRequiredFields = "Número de empleado, número de plaza, fecha de inicio vigencia, fecha de fin de vigencia (opcional), Horario de entrada 1, Horario de salida 1, horario de entrada 2 (opc), horario de salida 2 (opc), horario turno opcional 1, horario salida turno opcional, clave del servicio, nivel de riesgo"
					Case 21
						sRequiredFields = "Número de empleado, número de plaza, fecha de inicio vigencia, fecha de fin de vigencia (opcional)"
					Case 28
						sRequiredFields = "Número de empleado"
					Case 50
						sRequiredFields = "Número de empleado, número de plaza, fecha de inicio vigencia, fecha de fin de vigencia (opcional), clave del servicio, nivel de riesgo"
					Case 52
						sRequiredFields = "Número de empleado, fecha de inicio vigencia, fecha de fin de vigencia (opcional), clave de centro de trabajo, clave de centro de pago, clave del servicio"
					Case 54
						sRequiredFields = "Número de plaza, clave del servicio"
					Case 53
						sRequiredFields = "Número de empleado, la fecha de registro, la fecha de aplicación en nómina y el porcentaje para la compensación"
					Case EMPLOYEES_EXTRAHOURS
						sRequiredFields = "Número de empleado, la fecha de registro, la fecha de aplicación en nómina, el número de horas extraordinarias y observaciones"
					Case EMPLOYEES_SUNDAYS
						sRequiredFields = "Número de empleado, la fecha del domingo, la fecha de aplicación en nómina y observaciones"
					Case EMPLOYEES_BENEFICIARIES_DEBIT
						sRequiredFields = "Número de empleado, fecha de inicio, fecha de fin, quincena de aplicación, importe, número del beneficiario y observaciones"
					Case EMPLOYEES_BANK_ACCOUNTS
						sRequiredFields = "Número de empleado, ID del banco, y número de cuenta"
					Case Else
						sRequiredFields = "Número de empleado, número de plaza, fecha de inicio vigencia, fecha de fin de vigencia (opcional), Horario de entrada 1, Horario de salida 1, horario de entrada 2 (opc), horario de salida 2 (opc), horario turno opcional 1, horario salida turno opcional, clave del servicio, nivel de riesgo"
				End Select
				Select Case lReasonID
					Case 54
					'Case EMPLOYEES_ADD_BENEFICIARIES
					'	Call DisplayEmployeeBeneficiaryForm(oRequest, oADODBConnection, GetASPFileName(""), GetASPFileName(""), aEmployeeComponent, sError)
					Case Else
						Call DisplayEmployeeForm(oRequest, oADODBConnection, GetASPFileName(""), "EmployeesMovements", lEmployeeID, lReasonID, aEmployeeComponent, sError)
				End Select				
			Case "EmployeesNew"
				lErrorNumber = GetEmployeeByStatus(oRequest, oADODBConnection, "1", aEmployeeComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					Call DisplayPendingEmployeesTable(oRequest, oADODBConnection, False, "EmployeesNew", "1", 0, 12, aEmployeeComponent, sErrorDescription)
				End If
			Case "EmployeesProfessionalDegree"
				sRequiredFields = "Número de empleado, la fecha de registro, la fecha de aplicación en nómina y la cantidad de ayuda para impresión"
				Call DisplayEmployeeConceptForm(oRequest, oADODBConnection, GetASPFileName(""), "Step=1&Action=EmployeesGlasses", "46", aEmployeeComponent, sErrorDescription)
			Case "EmployeesSafeSeparation", "EmployeesAddSafeSeparation"
				If sAction = "EmployeesSafeSeparation" Then
					sRequiredFields = "Número de empleado, la fecha de registro, la fecha de aplicación en nómina y el porcentaje para el seguro"
					Call DisplayEmployeeSafeSeparationForm(oRequest, oADODBConnection, GetASPFileName(""), "Step=1&Action=" & sAction, "120", aEmployeeComponent, sErrorDescription)
				Else
					sRequiredFields = "Número de empleado, la fecha de registro, la fecha de aplicación en nómina y la cantidad ($ o %) para el seguro adicional"
					Call DisplayEmployeeSafeSeparationForm(oRequest, oADODBConnection, GetASPFileName(""), "Step=1&Action=" & sAction, "87", aEmployeeComponent, sErrorDescription)
				End If
			Case "EmployeeSports"
				sRequiredFields = "Número de empleado, la fecha de registro, la fecha de aplicación en nómina y la cantidad de cuota deportiva"
				Call DisplayEmployeeConceptForm(oRequest, oADODBConnection, GetASPFileName(""), "Step=1&Action=EmployeeSports", "69", aEmployeeComponent, sErrorDescription)
			Case "EmployeesSundays"
				Call DisplayAbsenceForm(oRequest, oADODBConnection, GetASPFileName(""), "202", "OnlySundays=1", aAbsenceComponent, sErrorDescription)
			Case "FONAC"
			Case "Jobs"
				Select Case lReasonID
					Case 59
						sRequiredFields = "Centro de trabajo, centro de pago, clave de servicio, clave del puesto, clave del nivel ó (clave de GGN, clave de integración y clave de clasificación), clave del tipo de ocupación, clave de horario, clave de turno, fecha de inicio de vigencia, fecha de fin de vigencia (opcional)"
						Call DisplayJobForm(oRequest, oADODBConnection, GetASPFileName(""), aJobComponent, sErrorDescription)
					Case 60
						sRequiredFields = "Número de plaza*, clave del centro de trabajo, clave del centro de pago, clave del servicio, clave del turno, clave del horario, fecha de inicio de vigencia* (*Campos requeridos)"
					Case 61
						sRequiredFields = "Número de plaza*, fecha de inicio*, clave del puesto*, (nivel ó si es plaza de funcionario: clave de GGN, clave de clasificación, clave de integración)* (*requeridos)"
				End Select
			Case "JobServices"
			Case "MedicalAreas"
			Case "MetLifeInsurance1"
				Call DisplayEmployeeConceptForm(oRequest, oADODBConnection, GetASPFileName(""), "Step=1&Action=MetLifeInsurance1", "65", aEmployeeComponent, sErrorDescription)
			Case "MetLifeInsurance2"
				Call DisplayEmployeeConceptForm(oRequest, oADODBConnection, GetASPFileName(""), "Step=1&Action=MetLifeInsurance2", "66", aEmployeeComponent, sErrorDescription)
			Case "MortgageCredit"
				Call DisplayEmployeeConceptForm(oRequest, oADODBConnection, GetASPFileName(""), "Step=1&Action=MortgageCredit", "58", aEmployeeComponent, sErrorDescription)
			Case "MortgageInsurance"
				Call DisplayEmployeeConceptForm(oRequest, oADODBConnection, GetASPFileName(""), "Step=1&Action=MortgageInsurance", "57", aEmployeeComponent, sErrorDescription)
			Case "PersonalLoan"
				Call DisplayEmployeeConceptForm(oRequest, oADODBConnection, GetASPFileName(""), "Step=1&Action=MediumTermLoan", "61", aEmployeeComponent, sErrorDescription)
			Case "ResumptionOfWork"
				Call DisplayEmployeeForm(oRequest, oADODBConnection, "Employees.asp", "ResumptionOfWork", ",0,", lReasonID, aEmployeeComponent, sErrorDescription)
			Case "ThirdUploadMovements"
				lErrorNumber = DisplayEmployeesCreditsSearchForm(oRequest, oADODBConnection, GetASPFileName(""), False, sErrorDescription)
			Case Else
				sRequiredFields = "Numero de empleado, clave de incidencia, fecha de ocurrencia, fecha de aplicación, Observaciones"
				Call DisplayAbsenceForm(oRequest, oADODBConnection, GetASPFileName(""), lReasonID, "", aAbsenceComponent, sErrorDescription)
		End Select
	Response.Write "</DIV>"
	Select Case sAction
		Case "FONAC"
		Case "MedicalAreas", "JobServices"
			Response.Write "<IMG SRC=""Images/Crcl1.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; ShowDisplay(document.all['UploadInfoFormDiv']); if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };""><FONT FACE=""Arial"" SIZE=""2"">Deseo subir la información a través de un archivo</FONT></A><BR /><BR />"
			Response.Write "<DIV NAME=""UploadInfoFormDiv"" ID=""UploadInfoFormDiv"">"
		Case "ThirdUploadMovements"
			Response.Write "<DIV NAME=""UploadInfoFormDiv"" ID=""UploadInfoFormDiv"" STYLE=""display: none"">"
		Case Else
			Select Case lReasonID
				Case 54, 60, 61
					Response.Write "<IMG SRC=""Images/Crcl1.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; ShowDisplay(document.all['UploadInfoFormDiv']); if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };""><FONT FACE=""Arial"" SIZE=""2"">Deseo subir la información a través de un archivo</FONT></A><BR /><BR />"
					Response.Write "<DIV NAME=""UploadInfoFormDiv"" ID=""UploadInfoFormDiv"">"
				Case 57, 58
				Case 17, 18, 21, 28, 17, 18, 28, 29, 30, 31, 32, 33, 34, 43, 44, 45, 46, 47, 48, 36, 37, 38, 39, 40, 41, 51, 50, 26, 57, EMPLOYEES_ADD_BENEFICIARIES, EMPLOYEES_THIRD_PROCESS, EMPLOYEES_THIRD_CONCEPT, CANCEL_EMPLOYEES_CONCEPTS, CANCEL_EMPLOYEES_SSI, CANCEL_EMPLOYEES_C04
					Response.Write "<DIV NAME=""UploadInfoFormDiv"" ID=""UploadInfoFormDiv"" STYLE=""display: none"">"
				Case Else
					Response.Write "<IMG SRC=""Images/Crcl2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; ShowDisplay(document.all['UploadInfoFormDiv']); if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };""><FONT FACE=""Arial"" SIZE=""2"">Deseo subir la información a través de un archivo</FONT></A><BR /><BR />"
					Response.Write "<DIV NAME=""UploadInfoFormDiv"" ID=""UploadInfoFormDiv"" STYLE=""display: none"">"
			End Select
	End Select
	If (lReasonID <> 57) And (lReasonID <> 58) Then
		Select Case lReasonID
			Case EMPLOYEES_THIRD_CONCEPT
				Call DisplayInstructionsMessage("Captura manual de registros", "En esta sección solo está habilitada la captura manual de registros. Si requiere cargar registros por medio de un archivo, ir a la sección correspondiente.")
				Response.Write "<BR />"
			Case Else
				If sAction = "ThirdUploadMovements" Then
					Call DisplayInstructionsMessage("Consulta y activación de registros", "En esta sección solo está habilitada la captura manual de registros. Si requiere cargar registros por medio de un archivo, ir a la sección correspondiente.")
					Response.Write "<BR />"
				Else
					Response.Write "<FORM NAME=""UploadInfoFrm"" ID=""UploadInfoFrm"" METHOD=""POST"" onSubmit=""return bReady"">"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""2"" />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ActionHdn"" VALUE=""" & lReasonID & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Success"" ID=""ActionHdn"" VALUE=""" & lSuccess & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & lEmployeeID & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ErrorDescription"" ID=""EmployeeIDHdn"" VALUE=""" & sError & """ />"
						Select Case sAction
							Case "Third"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ThirdConcept"" ID=""ThirdConceptHdn"" VALUE=""" & oRequest("ThirdConcept").Item & """ />"
								Call DisplayInstructionsMessage("Paso 1. Introduzca el archivo a utilizar", "<BLOCKQUOTE><OL>" & _
																	"<LI><B>Seleccione el archivo de texto que contiene la información a subir con el botón Examinar.</B></LI>" & _
																	"<LI><B>De click en el botón Continuar.</B></LI>" & _
																"</OL></BLOCKQUOTE>")
							Case Else
								Call DisplayInstructionsMessage("Paso 1. Introduzca el archivo a utilizar", "<BLOCKQUOTE><OL>" & _
																	"<LI>Abra el documento de Excel con la información que desea subir.</LI>" & _
																	"<LI>Copie únicamente las celdas que contienen la información deseada.</LI>" & _
																	"<LI>Pegue dicha información en la caja de texto.</LI>" & _
																	"<LI><B>O seleccione el archivo de texto que contiene la información a subir.</B></LI>" & _
																"</OL></BLOCKQUOTE>")
						End Select
						Response.Write "<BR />"
						If Len(sRequiredFields) > 0 Then
							If lReasonID = 0 Then
								Response.Write "Para la asignación de número de empleado se requiere: <B>" & sRequiredFields & "</B>."
							Else
								Response.Write "Para este concepto se requiere: <B>" & sRequiredFields & "</B>."
							End If
							Response.Write "<BR />"
							Response.Write "<BR />"
						End If
						If sAction = "Third" Then
						Else
							Response.Write "<TEXTAREA NAME=""RawData"" ID=""RawDataTxtArea"" ROWS=""10"" COLS=""119"" CLASS=""TextFields"" onChange=""bReady = (this.value != '')""></TEXTAREA><BR /><BR />"
							Response.Write "<INPUT TYPE=""SUBMIT"" VALUE=""Continuar"" CLASS=""Buttons"" />"
						End If
					Response.Write "</FORM>"
					Response.Write "<IFRAME SRC=""BrowserFileForInfo.asp?Action=" & oRequest("Action").Item & "&UserID=" & aLoginComponent(N_USER_ID_LOGIN) & """ NAME=""UploadInfoIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""100""></IFRAME>"
					Response.Write "<BR />"
				End If
		End Select
		Response.Write "</DIV>"
	End If
	Select Case sAction
		Case "Absences"
			Response.Write "<IMG SRC=""Images/Crcl3.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; ShowDisplay(document.all['UploadValidateInfoFormDiv']); if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };"">Registros en proceso de aplicación</A><BR /><BR />"
			Response.Write "<DIV NAME=""UploadValidateInfoFormDiv"" ID=""UploadValidateInfoFormDiv"" STYLE=""display: none"">"
				Response.Write "<FORM NAME=""UploadValidateInfoFrm"" ID=""UploadValidateInfoFrm"" METHOD=""POST"" onSubmit=""return bReady"">"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & sAction & """ />"
					If CInt(Request.Cookies("SIAP_SectionID")) <> 7 Then
						Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""AuthorizationFile"" ID=""ModifyBtn"" VALUE=""Aplicar Movimientos Seleccionados"" CLASS=""Buttons""/>"
					End If
					Response.Write "<BR /><BR />"
					lErrorNumber = DisplayPendingEmployeesAbscencesTable(oRequest, oADODBConnection, 0, False, lReasonID, sAction, aEmployeeComponent, sErrorDescription)
					If lErrorNumber <> 0 Then
						Response.Write "<BR />"
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						sErrorDescription = ""
					End If
				Response.Write "</FORM>"
			Response.Write "</DIV>"
		Case "ConceptsValues"
		Case "EmployeesAssignNumber"
			Response.Write "<IMG SRC=""Images/Crcl3.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; ShowDisplay(document.all['UploadValidateInfoFormDiv']); if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };"">Registros en proceso de asignación de plaza</A><BR /><BR />"
			Response.Write "<DIV NAME=""UploadValidateInfoFormDiv"" ID=""UploadValidateInfoFormDiv"" STYLE=""display: none"">"
				Response.Write "<FORM NAME=""UploadValidateInfoFrm"" ID=""UploadValidateInfoFrm"" METHOD=""POST"" onSubmit=""return bReady"">"
					lErrorNumber = DisplayPendingEmployeesTable(oRequest, oADODBConnection, False, "EmployeesMovements", lReasonID, 0, aEmployeeComponent, sErrorDescription)
					If lErrorNumber <> 0 Then
						Response.Write "<BR />"
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						sErrorDescription = ""
					End If
				Response.Write "</FORM>"
			Response.Write "</DIV>"
		Case "EmployeesMovements"
			bShowSection3 = True
			Select Case lReasonID
				Case 17, 18, 28, 29, 30, 31, 32, 33, 34, 43, 44, 45, 46, 47, 48, 36, 37, 38, 39, 40, 41
					sNumber = "2"
					sMessage = "Registros en proceso de aplicación"
				Case 54
					sNumber = "2"
					sMessage = "Consulta de plazas que cambiaron servicio el día de hoy"
				Case 57, 58, 59, EMPLOYEES_DOCUMENTS_FOR_LICENSES, 21,51,50,26, CANCEL_EMPLOYEES_CONCEPTS, CANCEL_EMPLOYEES_SSI, CANCEL_EMPLOYEES_C04
					bShowSection3 = False
				Case EMPLOYEES_THIRD_CONCEPT
					sNumber = "3"
					sMessage = "Consulta de registros de terceros capturados manualmente en proceso de aplicación"
				Case EMPLOYEES_ADD_BENEFICIARIES
					sNumber = "3"
					sMessage = "Consulta de beneficiarios del empleado"
				Case EMPLOYEES_BENEFICIARIES_DEBIT
					sNumber = "3"
					sMessage = "Consulta de adeudo de pensión alimenticia por aplicar"
				Case Else
					sNumber = "3"
					sMessage = "Registros en proceso de aplicación"
			End Select
			If bShowSection3 Then
				Response.Write "<IMG SRC=""Images/Crcl" & sNumber & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; ShowDisplay(document.all['UploadValidateInfoFormDiv']); if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };"">" & sMessage & "</A><BR /><BR />"
			End If
			Response.Write "<DIV NAME=""UploadValidateInfoFormDiv"" ID=""UploadValidateInfoFormDiv"" STYLE=""display: none"">"
				Response.Write "<FORM NAME=""UploadValidateInfoFrm"" ID=""UploadValidateInfoFrm"" METHOD=""POST"">"
					Select Case lReasonID
						Case -89
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 100
						Case -74
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 86
						Case EMPLOYEES_DOCUMENTS_FOR_LICENSES
						Case EMPLOYEES_THIRD_CONCEPT
						Case EMPLOYEES_SAFE_SEPARATION
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 120
						Case EMPLOYEES_ADD_SAFE_SEPARATION
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 87
						Case 12
						Case EMPLOYEES_ANTIQUITIES
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 5
						Case EMPLOYEES_ADDITIONALSHIFT
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 7
						Case EMPLOYEES_GLASSES
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 24


						Case EMPLOYEES_FAMILY_DEATH
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 45
						Case EMPLOYEES_PROFESSIONAL_DEGREE
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 46
						Case EMPLOYEES_MONTHAWARD
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 50
						Case EMPLOYEES_SPORTS_HELP
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 165
						Case EMPLOYEES_SPORTS
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 69
						Case EMPLOYEES_CARLOAN
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 74
						Case 53
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 4
						Case EMPLOYEES_CONCEPT_08
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 8
						Case EMPLOYEES_CHILDREN_SCHOOLARSHIPS
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 22
						Case EMPLOYEES_LICENSES
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 104
						Case EMPLOYEES_CONCEPT_16
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 19
						Case EMPLOYEES_NON_EXCENT
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 72
						Case EMPLOYEES_EXCENT
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 73
						Case EMPLOYEES_MOTHERAWARD
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 26
						Case EMPLOYEES_HELP_COMISSION
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 63
						Case EMPLOYEES_SAFEDOWN
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 67
						Case EMPLOYEES_ANUAL_AWARD
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 32
						Case EMPLOYEES_EXTRAHOURS
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 201
						Case EMPLOYEES_SUNDAYS
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 202
						Case EMPLOYEES_BENEFICIARIES
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 70
						Case EMPLOYEES_BENEFICIARIES_DEBIT
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 86
						Case EMPLOYEES_FONAC_CONCEPT
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 77
						Case EMPLOYEES_NIGHTSHIFTS
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 93
						Case 59
						Case 54
					End Select
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & sAction & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SaveEmployeesMovements"" ID=""SaveEmployeesMovementsHdn"" VALUE=""1"" />"
					If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""AuthorizationFile"" ID=""ModifyBtn"" VALUE=""Aplicar Movimientos Seleccionados"" CLASS=""Buttons""/>"
					Response.Write "<BR /><BR />"
					Select Case lReasonID
						Case EMPLOYEES_THIRD_CONCEPT
							lErrorNumber = DisplayPendingEmployeesCreditsTable(oRequest, oADODBConnection, 0, False, sAction, lReasonID, aEmployeeComponent, sErrorDescription)
						Case EMPLOYEES_ADD_BENEFICIARIES
							lErrorNumber = DisplayPendingEmployeesbeneficiariesTable(oRequest, oADODBConnection, False, sAction, lReasonID, aEmployeeComponent, sErrorDescription)
						Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS
							'lErrorNumber = DisplayPendingEmployeesConceptsTable(oRequest, oADODBConnection, False, sAction, lReasonID, aEmployeeComponent, sErrorDescription)
							lErrorNumber = DisplayPendingEmployeesTable(oRequest, oADODBConnection, False, sAction, lReasonID, 0, aEmployeeComponent, sErrorDescription)
						Case Else
							lErrorNumber = DisplayPendingEmployeesTable(oRequest, oADODBConnection, False, sAction, lReasonID, 0, aEmployeeComponent, sErrorDescription)
					End Select
					If lErrorNumber <> 0 Then
						Response.Write "<BR />"
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						sErrorDescription = ""
					End If
				Response.Write "</FORM>"
			Response.Write "</DIV>"
		Case "FONAC"
		Case "Jobs"
			If lReasonID = 60 Then
				Response.Write "<IMG SRC=""Images/Crcl2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; ShowDisplay(document.all['UploadValidateInfoFormDiv']); if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };"">Consultar las plazas que cambiaron datos el día de hoy</A><BR /><BR />"
				Response.Write "<DIV NAME=""UploadValidateInfoFormDiv"" ID=""UploadValidateInfoFormDiv"" STYLE=""display: none"">"
					Response.Write "<FORM NAME=""UploadValidateInfoFrm"" ID=""UploadValidateInfoFrm"" METHOD=""POST"" onSubmit=""return bReady"">"
						lErrorNumber = DisplayPendingJobsTable(oRequest, oADODBConnection, False, sAction, lReasonID, 0, aJobComponent, sErrorDescription)
						If lErrorNumber <> 0 Then
							Response.Write "<BR />"
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							lErrorNumber = 0
							sErrorDescription = ""
						End If
					Response.Write "</FORM>"
				Response.Write "</DIV>"
			ElseIf lReasonID = 61 Then
			Else
				Response.Write "<IMG SRC=""Images/Crcl3.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; ShowDisplay(document.all['UploadValidateJobInfoFormDiv']); if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };"">Consulta de registros en proceso de aplicación</A><BR /><BR />"
				Response.Write "<DIV NAME=""UploadValidateJobInfoFormDiv"" ID=""UploadValidateJobInfoFormDiv"" STYLE=""display: none"">"
					Response.Write "<FORM NAME=""UploadValidateJobInfoFrm"" ACTION=""Jobs.asp"" ID=""UploadValidateJobInfoFrm"" METHOD=""POST"">"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & sAction & """ />"
						Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""AuthorizationFile"" ID=""ModifyBtn"" VALUE=""Aplicar Movimientos Seleccionados"" CLASS=""Buttons""/>"
						Response.Write "<BR /><BR />"
						lErrorNumber = DisplayPendingJobsTable(oRequest, oADODBConnection, False, sAction, lReasonID, 0, aJobComponent, sErrorDescription)
						If lErrorNumber <> 0 Then
							Response.Write "<BR />"
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							lErrorNumber = 0
							sErrorDescription = ""
						End If
					Response.Write "</FORM>"
				Response.Write "</DIV>"
			End If
		Case "MedicalAreas"
		Case "Third"
		Case "ThirdUploadMovements"
			Response.Write "<IMG SRC=""Images/Crcl3.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; ShowDisplay(document.all['UploadValidateInfoFormDiv']); if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };"">Active los registros cargados desde el archivo del tercero</A><BR /><BR />"
			Response.Write "<DIV NAME=""UploadValidateInfoFormDiv"" ID=""UploadValidateInfoFormDiv"" STYLE=""display: none"">"
				Response.Write "<FORM NAME=""UploadValidateInfoFrm"" ID=""UploadValidateInfoFrm"" METHOD=""POST"">"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & sAction & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SaveEmployeesMovements"" ID=""SaveEmployeesMovementsHdn"" VALUE=""1"" />"
					If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""AuthorizationFile"" ID=""ModifyBtn"" VALUE=""Aplicar Movimientos Seleccionados"" CLASS=""Buttons""/>"
					Response.Write "<BR /><BR />"
					lErrorNumber = DisplayPendingEmployeesCreditsTable(oRequest, oADODBConnection, 0, False, sAction, EMPLOYEES_THIRD_PROCESS, aEmployeeComponent, sErrorDescription)
					If lErrorNumber <> 0 Then
						Response.Write "<BR />"
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						sErrorDescription = ""
					End If
				Response.Write "</FORM>"
			Response.Write "</DIV>"
		Case Else
			Response.Write "<IMG SRC=""Images/Crcl3.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; ShowDisplay(document.all['UploadValidateInfoFormDiv']); if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };"">Consulta de registros en proceso de aplicación</A><BR /><BR />"
			Response.Write "<DIV NAME=""UploadValidateInfoFormDiv"" ID=""UploadValidateInfoFormDiv"" STYLE=""display: none"">"
				Response.Write "<FORM NAME=""UploadValidateInfoFrm"" ID=""UploadValidateInfoFrm"" METHOD=""POST"" onSubmit=""return bReady"">"
					lErrorNumber = DisplayPendingEmployeesTable(oRequest, oADODBConnection, False, sAction, lReasonID, 0, aEmployeeComponent, sErrorDescription)
					If lErrorNumber <> 0 Then
						Response.Write "<BR />"
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						sErrorDescription = ""
					End If
				Response.Write "</FORM>"
			Response.Write "</DIV>"
	End Select
	Select Case sAction
		Case "EmployeesAssignNumber"
		Case "Absences"
			Response.Write "<IMG SRC=""Images/Crcl4.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; ShowDisplay(document.all['ConceptInfoFormDiv']);""><FONT FACE=""Arial"" SIZE=""2"">Consulta de registros existentes para el empleado</FONT></A><BR /><BR />"
			Response.Write "<DIV NAME=""ConceptInfoFormDiv"" ID=""ConceptInfoFormDiv"" STYLE=""display: none"">"
			Response.Write "<FORM NAME=""ConceptInfoFrm"" ID=""ConceptInfoFrm"" METHOD=""POST"" onSubmit=""return bReady"">"
				lErrorNumber = DisplayPendingEmployeesAbscencesTable(oRequest, oADODBConnection, 1, False, lReasonID, sAction, aEmployeeComponent, sErrorDescription)
				If lErrorNumber <> 0 Then
					Response.Write "<BR />"
					Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
					lErrorNumber = 0
					sErrorDescription = ""
				End If
			Response.Write "</FORM>"
		Case Else
			If (lReasonID <= -58) And (lReasonID <> EMPLOYEES_THIRD_PROCESS) And (lReasonID <> EMPLOYEES_ADD_BENEFICIARIES) Then
					Response.Write "<IMG SRC=""Images/Crcl4.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; ShowDisplay(document.all['ConceptInfoFormDiv']);""><FONT FACE=""Arial"" SIZE=""2"">Consulta de registros existentes para el empleado</FONT></A><BR /><BR />"
					Response.Write "<DIV NAME=""ConceptInfoFormDiv"" ID=""ConceptInfoFormDiv"" STYLE=""display: none"">"
					Response.Write "<FORM NAME=""ConceptInfoFrm"" ID=""ConceptInfoFrm"" METHOD=""POST"" onSubmit=""return bReady"">"
						Select Case lReasonID
							Case CANCEL_EMPLOYEES_CONCEPTS, CANCEL_EMPLOYEES_SSI, CANCEL_EMPLOYEES_C04
								lErrorNumber = DisplayPendingEmployeesConceptsTable(oRequest, oADODBConnection, False, sAction, lReasonID, aEmployeeComponent, sErrorDescription)
							Case EMPLOYEES_THIRD_CONCEPT
								lErrorNumber = DisplayPendingEmployeesCreditsTable(oRequest, oADODBConnection, 1, False, sAction, lReasonID, aEmployeeComponent, sErrorDescription)
							Case Else
								lErrorNumber = DisplayPendingEmployeesTable(oRequest, oADODBConnection, False, sAction, lReasonID, 1, aEmployeeComponent, sErrorDescription)
						End Select
						If lErrorNumber <> 0 Then
							Response.Write "<BR />"
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							lErrorNumber = 0
							sErrorDescription = ""
						End If
					Response.Write "</FORM>"
			End If
	End Select
	Response.Write "</DIV>"
	DisplayUploadForm = Err.Number
	Err.Clear
End Function
%>