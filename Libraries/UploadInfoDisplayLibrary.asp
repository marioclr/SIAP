<!-- #include file="ReportsQueries1200Lib.asp" -->
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

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo.</BLOCKQUOTE>")
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

Function DisplayConceptsValuesColumns(sFileName, lEmployeeTypeID, bFull, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns, lEmployeeTypeID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayConceptsValuesColumns"
	Dim iColumns
	Dim iIndex
	Dim lErrorNumber

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo.</BLOCKQUOTE>")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""UploadAbsencesFrm"" ID=""UploadAbsencesFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDHdn"" VALUE=""" & lEmployeeTypeID & """ />"
			For iIndex = 1 To iColumns
				Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
				Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					If bFull Then
							Response.Write "<OPTION VALUE=""EmployeeTypeID"">Tipo de tabulador</OPTION>"
					End If
					Response.Write "<OPTION VALUE=""CompanyID"">Compañía</OPTION>"
					Response.Write "<OPTION VALUE=""ConceptID"">Clave concepto de pago</OPTION>"
					If bFull Then
							Response.Write "<OPTION VALUE=""PositionTypeID"">Tipo de puesto</OPTION>"
					End If
					Response.Write "<OPTION VALUE=""PositionShortNames"">Clave del puesto</OPTION>"
					If Not bFull Then
						Select Case lEmployeeTypeID
							Case 0
								Response.Write "<OPTION VALUE=""PositionTypeID"">Tipo de puesto</OPTION>"
								Response.Write "<OPTION VALUE=""LevelID"">Nivel</OPTION>"
								Response.Write "<OPTION VALUE=""WorkingHours"">Horas laboradas(Jornada)</OPTION>"
								Response.Write "<OPTION VALUE=""EconomicZoneID"">Zona económica</OPTION>"
							Case 1
								Response.Write "<OPTION VALUE=""GroupGradeLevelID"">Grupo grado nivel</OPTION>"
								Response.Write "<OPTION VALUE=""ClassificationID"">Clasificación</OPTION>"
								Response.Write "<OPTION VALUE=""IntegrationID"">Integración</OPTION>"
							Case 2
								Response.Write "<OPTION VALUE=""PositionTypeID"">Tipo de puesto</OPTION>"
								Response.Write "<OPTION VALUE=""LevelID"">Nivel</OPTION>"
								Response.Write "<OPTION VALUE=""EconomicZoneID"">Zona económica</OPTION>"
							Case 3
								Response.Write "<OPTION VALUE=""LevelID"">Nivel</OPTION>"
							Case 4, 5, 6
								Response.Write "<OPTION VALUE=""LevelID"">Nivel</OPTION>"
								Response.Write "<OPTION VALUE=""EconomicZoneID"">Zona económica</OPTION>"
						End Select
					Else
						Response.Write "<OPTION VALUE=""GroupGradeLevelID"">Grupo grado nivel</OPTION>"
						Response.Write "<OPTION VALUE=""LevelID"">Nivel</OPTION>"
						Response.Write "<OPTION VALUE=""EmployeeStatusID"">Estatus del empleado</OPTION>"
						Response.Write "<OPTION VALUE=""JobStatusID"">Estatus de la plaza</OPTION>"
						Response.Write "<OPTION VALUE=""ClassificationID"">Clasificación</OPTION>"
						Response.Write "<OPTION VALUE=""IntegrationID"">Integración</OPTION>"
						Response.Write "<OPTION VALUE=""JourneyID"">Jornada</OPTION>"
						Response.Write "<OPTION VALUE=""WorkingHours"">Horas laboradas</OPTION>"
						Response.Write "<OPTION VALUE=""AdditionalShift"">Turno opcional</OPTION>"
						Response.Write "<OPTION VALUE=""EconomicZoneID"">Zona económica</OPTION>"
						Response.Write "<OPTION VALUE=""ServiceID"">Servicio</OPTION>"
						Response.Write "<OPTION VALUE=""AntiquityID"">Antigüedad en el ISSSTE</OPTION>"
						Response.Write "<OPTION VALUE=""Antiquity2ID"">Antigüedad consecutiva</OPTION>"
						Response.Write "<OPTION VALUE=""Antiquity3ID"">Antigüedad en el ISSSTE con plaza de base</OPTION>"
						Response.Write "<OPTION VALUE=""Antiquity4ID"">Antigüedad federal</OPTION>"
						Response.Write "<OPTION VALUE=""ForRisk"">Riesgos profesionales</OPTION>"
						Response.Write "<OPTION VALUE=""GenderID"">Género</OPTION>"
						Response.Write "<OPTION VALUE=""HasChildren"">Hijos</OPTION>"
						Response.Write "<OPTION VALUE=""SchoolarshipID"">Escolaridad (hijos)</OPTION>"
						Response.Write "<OPTION VALUE=""HasSyndicate"">Sindicalizado</OPTION>"
					End If
					Response.Write "<OPTION VALUE=""OcurredStartDateYYYYMMDD"">Fecha de inicio vigencia (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredStartDateDDMMYYYY"">Fecha de inicio vigencia (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredStartDateMMDDYYYY"">Fecha de inicio vigencia (MM-DD-AAAA)</OPTION>"
					If CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
						Response.Write "<OPTION VALUE=""ConceptAmount"">Monto quincenal</OPTION>"
					Else
						Response.Write "<OPTION VALUE=""ConceptAmount"">Monto mensual</OPTION>"
					End If
					If bFull Then
							Response.Write "<OPTION VALUE=""ConceptCurrencyID"">Unidad del Monto</OPTION>"
							Response.Write "<OPTION VALUE=""AppliesToID"">Conceptos sobre los que aplica</OPTION>"
							Response.Write "<OPTION VALUE=""ConceptMin"">Monto mínimo</OPTION>"
							Response.Write "<OPTION VALUE=""ConceptMinQttyID"">Unidad del monto mínimo</OPTION>"
							Response.Write "<OPTION VALUE=""ConceptMax"">Monto máximo</OPTION>"
							Response.Write "<OPTION VALUE=""ConceptMaxQttyID"">Unidad del monto máximo</OPTION>"
					End If
					Response.Write "<OPTION VALUE=""OcurredEndDateYYYYMMDD"">Fecha de fin vigencia (AAAAMMDD) (opcional)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredEndDateDDMMYYYY"">Fecha de fin vigencia (DD-MM-AAAA) (opcional)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredEndDateMMDDYYYY"">Fecha de fin vigencia (MM-DD-AAAA) (opcional)</OPTION>"
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			If bFull Then
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "if(document.all['Column1'] != null) SendURLValuesToForm('Column1=EmployeeTypeID', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column2'] != null) SendURLValuesToForm('Column2=CompanyID', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column3'] != null) SendURLValuesToForm('Column3=ConceptID', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column4'] != null) SendURLValuesToForm('Column4=PositionTypeID', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column5'] != null) SendURLValuesToForm('Column5=PositionShortNames', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column6'] != null) SendURLValuesToForm('Column6=LevelID', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column7'] != null) SendURLValuesToForm('Column7=GroupGradeLevelID', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column8'] != null) SendURLValuesToForm('Column8=EmployeeStatusID', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column9'] != null) SendURLValuesToForm('Column9=JobStatusID', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column10'] != null) SendURLValuesToForm('Column10=ClassificationID', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column11'] != null) SendURLValuesToForm('Column11=IntegrationID', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column12'] != null) SendURLValuesToForm('Column12=JourneyID', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column13'] != null) SendURLValuesToForm('Column13=WorkingHours', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column14'] != null) SendURLValuesToForm('Column14=AdditionalShift', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column15'] != null) SendURLValuesToForm('Column15=EconomicZoneID', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column16'] != null) SendURLValuesToForm('Column16=ServiceID', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column17'] != null) SendURLValuesToForm('Column17=AntiquityID', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column18'] != null) SendURLValuesToForm('Column18=Antiquity2ID', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column19'] != null) SendURLValuesToForm('Column19=Antiquity3ID', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column20'] != null) SendURLValuesToForm('Column20=Antiquity4ID', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column21'] != null) SendURLValuesToForm('Column21=ForRisk', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column22'] != null) SendURLValuesToForm('Column22=GenderID', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column23'] != null) SendURLValuesToForm('Column23=HasChildren', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column24'] != null) SendURLValuesToForm('Column24=SchoolarshipID', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column25'] != null) SendURLValuesToForm('Column25=HasSyndicate', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column26'] != null) SendURLValuesToForm('Column26=ConceptAmount', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column27'] != null) SendURLValuesToForm('Column27=OcurredStartDateDDMMYYYY', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column28'] != null) SendURLValuesToForm('Column28=OcurredEndDateDDMMYYYY', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column29'] != null) SendURLValuesToForm('Column29=ConceptCurrencyID', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column30'] != null) SendURLValuesToForm('Column30=AppliesToID', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column31'] != null) SendURLValuesToForm('Column31=ConceptMin', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column32'] != null) SendURLValuesToForm('Column32=ConceptMinQttyID', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column33'] != null) SendURLValuesToForm('Column33=ConceptMax', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "if(document.all['Column34'] != null) SendURLValuesToForm('Column34=ConceptMaxQttyID', document.UploadAbsencesFrm);" & vbNewLine
				Response.Write "//--></SCRIPT>"
			End If
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
				Response.Write "if (sFields.search(/CompanyID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la compañía.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/ConceptID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la clave del concepto de pago.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/PositionShortNames/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el código del puesto.');" & vbNewLine
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
				Response.Write "if (((sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1) && ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) && ((sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1) || (sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/OcurredStartDateMMDDYYYY/gi) != -1) && ((sFields.search(/OcurredStartDateDDMMYYYY/gi) != -1) || (sFields.search(/OcurredStartDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de inicio vigengia con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				If Not bFull Then
					Select Case lEmployeeTypeID
						Case 0
							Response.Write "if (sFields.search(/PositionTypeID/gi) == -1) {" & vbNewLine
								Response.Write "alert('No se especificó qué columna contiene el tipo de puesto.');" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
							Response.Write "if (sFields.search(/LevelID/gi) == -1) {" & vbNewLine
								Response.Write "alert('No se especificó qué columna contiene el nivel del puesto.');" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
							Response.Write "if (sFields.search(/WorkingHours/gi) == -1) {" & vbNewLine
								Response.Write "alert('No se especificó qué columna contiene las Horas laboradas(Jornada).');" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
							Response.Write "if (sFields.search(/EconomicZoneID/gi) == -1) {" & vbNewLine
								Response.Write "alert('No se especificó qué columna contiene la Zona económica.');" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
						Case 1
							Response.Write "if (sFields.search(/GroupGradeLevelID/gi) == -1) {" & vbNewLine
								Response.Write "alert('No se especificó qué columna contiene el Grupo grado nivel del puesto.');" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
							Response.Write "if (sFields.search(/ClassificationID/gi) == -1) {" & vbNewLine
								Response.Write "alert('No se especificó qué columna contiene la Clasificación del puesto.');" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
							Response.Write "if (sFields.search(/IntegrationID/gi) == -1) {" & vbNewLine
								Response.Write "alert('No se especificó qué columna contiene la Integración del puesto.');" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
						Case 2
							Response.Write "if (sFields.search(/PositionTypeID/gi) == -1) {" & vbNewLine
								Response.Write "alert('No se especificó qué columna contiene el tipo de puesto.');" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
							Response.Write "if (sFields.search(/LevelID/gi) == -1) {" & vbNewLine
								Response.Write "alert('No se especificó qué columna contiene el nivel del puesto.');" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
							Response.Write "if (sFields.search(/EconomicZoneID/gi) == -1) {" & vbNewLine
								Response.Write "alert('No se especificó qué columna contiene la Zona económica.');" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
						Case 3
							Response.Write "if (sFields.search(/LevelID/gi) == -1) {" & vbNewLine
								Response.Write "alert('No se especificó qué columna contiene el nivel del puesto.');" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
						Case 4, 5, 6
							Response.Write "if (sFields.search(/LevelID/gi) == -1) {" & vbNewLine
								Response.Write "alert('No se especificó qué columna contiene el nivel del puesto.');" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
							Response.Write "if (sFields.search(/EconomicZoneID/gi) == -1) {" & vbNewLine
								Response.Write "alert('No se especificó qué columna contiene la Zona económica.');" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
						Case Else
					End Select
				Else
					Response.Write "if (sFields.search(/LevelID/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se especificó qué columna contiene el nivel del puesto.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (sFields.search(/GroupGradeLevelID/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se especificó qué columna contiene la Zona económica.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (sFields.search(/EmployeeStatusID/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se especificó qué columna contiene el nivel del puesto.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (sFields.search(/JobStatusID/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se especificó qué columna contiene la Zona económica.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (sFields.search(/ClassificationID/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se especificó qué columna contiene el nivel del puesto.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (sFields.search(/IntegrationID/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se especificó qué columna contiene la Zona económica.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (sFields.search(/JourneyID/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se especificó qué columna contiene el nivel del puesto.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (sFields.search(/WorkingHours/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se especificó qué columna contiene la Zona económica.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (sFields.search(/AdditionalShift/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se especificó qué columna contiene el nivel del puesto.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (sFields.search(/EconomicZoneID/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se especificó qué columna contiene la Zona económica.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (sFields.search(/ServiceID/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se especificó qué columna contiene el nivel del puesto.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (sFields.search(/AntiquityID/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se especificó qué columna contiene la Zona económica.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (sFields.search(/Antiquity2ID/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se especificó qué columna contiene el nivel del puesto.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (sFields.search(/Antiquity3ID/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se especificó qué columna contiene la Zona económica.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (sFields.search(/Antiquity4ID/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se especificó qué columna contiene el nivel del puesto.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (sFields.search(/ForRisk/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se especificó qué columna contiene la Zona económica.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (sFields.search(/GenderID/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se especificó qué columna contiene el nivel del puesto.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (sFields.search(/HasChildren/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se especificó qué columna contiene la Zona económica.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (sFields.search(/SchoolarshipID/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se especificó qué columna contiene el nivel del puesto.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (sFields.search(/HasSyndicate/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se especificó qué columna contiene la Zona económica.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
				End If
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

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo.</BLOCKQUOTE>")
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

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo.</BLOCKQUOTE>")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""UploadAbsencesFrm"" ID=""UploadAbsencesFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UploadFile"" ID=""UploadFileHdn"" VALUE=""" & oRequest("UploadFile").Item & """ />"
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
					Response.Write "<OPTION VALUE=""OcurredDateYYYYMMDD"">Fecha de ocurrencia   (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateDDMMYYYY"">Fecha de ocurrencia (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""OcurredDateMMDDYYYY"">Fecha de ocurrencia (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""EndDateYYYYMMDD"">Fecha de termino   (AAAAMMDD) (opcional)</OPTION>"
					Response.Write "<OPTION VALUE=""EndDateDDMMYYYY"">Fecha de termino (DD-MM-AAAA) (opcional)</OPTION>"
					Response.Write "<OPTION VALUE=""EndDateMMDDYYYY"">Fecha de termino (MM-DD-AAAA) (opcional)</OPTION>"
					Response.Write "<OPTION VALUE=""PayrollDateYYYYMMDD"">Quincena de aplicación   (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""PayrollDateDDMMYYYY"">Quincena de aplicación (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""PayrollDateMMDDYYYY"">Quincena de aplicación (MM-DD-AAAA)</OPTION>"
					Select Case lReasonID
						Case EMPLOYEES_EXTRAHOURS
							Response.Write "<OPTION VALUE=""AbsenceHours"">Horas extras</OPTION>"
						Case EMPLOYEES_SUNDAYS
						Case 0, 1
							Response.Write "<OPTION VALUE=""VacationPeriod"">Periódo de vacaciones/estimulo(opcional)</OPTION>"
							Response.Write "<OPTION VALUE=""PeriodYear"">Año del periódo de vacaciones/estimulo(opcional)</OPTION>"
							Response.Write "<OPTION VALUE=""ForJustificationID"">Incidencia a justificar(opcional)</OPTION>"
						Case Else
							Response.Write "<OPTION VALUE=""DocumentNumber"">No. de oficio</OPTION>"
							Response.Write "<OPTION VALUE=""AbsenceHours"">Horas de retardo</OPTION>"
							Response.Write "<OPTION VALUE=""JustificationID"">Justificación</OPTION>"
							Response.Write "<OPTION VALUE=""AppliesForPunctuality"">¿Aplica para puntualidad?</OPTION>"
					End Select
					Response.Write "<OPTION VALUE=""Reasons"">Observaciones(opcional)</OPTION>"
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			Select Case lReasonID
				Case 0, 1
					Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
						Response.Write "if(document.all['Column1'] != null) SendURLValuesToForm('Column1=EmployeeID', document.UploadAbsencesFrm);" & vbNewLine
						Response.Write "if(document.all['Column2'] != null) SendURLValuesToForm('Column2=AbsenceID', document.UploadAbsencesFrm);" & vbNewLine
						Response.Write "if(document.all['Column3'] != null) SendURLValuesToForm('Column3=OcurredDateDDMMYYYY', document.UploadAbsencesFrm);" & vbNewLine
						Response.Write "if(document.all['Column4'] != null) SendURLValuesToForm('Column4=EndDateDDMMYYYY', document.UploadAbsencesFrm);" & vbNewLine
						Response.Write "if(document.all['Column5'] != null) SendURLValuesToForm('Column5=PayrollDateYYYYMMDD', document.UploadAbsencesFrm);" & vbNewLine
						Response.Write "if(document.all['Column6'] != null) SendURLValuesToForm('Column6=VacationPeriod', document.UploadAbsencesFrm);" & vbNewLine
						Response.Write "if(document.all['Column7'] != null) SendURLValuesToForm('Column7=PeriodYear', document.UploadAbsencesFrm);" & vbNewLine
						Response.Write "if(document.all['Column8'] != null) SendURLValuesToForm('Column8=ForJustificationID', document.UploadAbsencesFrm);" & vbNewLine
					Response.Write "//--></SCRIPT>"
			End Select
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
					Response.Write "alert('No se especificó qué columna contiene la quincena de aplicación.');" & vbNewLine
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

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo.</BLOCKQUOTE>")
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
					Response.Write "<OPTION VALUE=""EmployeeID"">Número de empleado</OPTION>"
					Response.Write "<OPTION VALUE=""ConceptID"">Clave del concepto</OPTION>"
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

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo.</BLOCKQUOTE>")
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

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo.</BLOCKQUOTE>")
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

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo.<BLOCKQUOTE>")
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
					Response.Write "<OPTION VALUE=""BankID"">Clave del banco</OPTION>"
					Response.Write "<OPTION VALUE=""AccountNumber"">No. de cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""PayrollDateYYYYMMDD"">Quincena de aplicación (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""PayrollDateDDMMYYYY"">Quincena de aplicación (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""PayrollDateMMDDYYYY"">Quincena de aplicación (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""EndDateYYYYMMDD"">Fecha de fin (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""EndDateDDMMYYYY"">Fecha de fin (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""EndDateMMDDYYYY"">Fecha de fin (MM-DD-AAAA)</OPTION>"
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
				Response.Write "if (sFields.search(/BankID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la clave del banco.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/AccountNumber/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el número de cuenta.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if ((sFields.search(/PayrollDateYYYYMMDD/gi) == -1) && (sFields.search(/PayrollDateDDMMYYYY/gi) == -1) && (sFields.search(/PayrollDateMMDDYYYY/gi) == -1)) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene la quincena de aplicación.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((sFields.search(/PayrollDateYYYYMMDD/gi) != -1) && ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) || (sFields.search(/PayrollDateMMDDYYYY/gi) != -1))) || ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) && ((sFields.search(/PayrollDateYYYYMMDD/gi) != -1) || (sFields.search(/PayrollDateMMDDYYYY/gi) != -1))) || ((sFields.search(/PayrollDateMMDDYYYY/gi) != -1) && ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) || (sFields.search(/PayrollDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la quincena de aplicación con diferente formato.');" & vbNewLine
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

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo.</BLOCKQUOTE>")
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

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo.</BLOCKQUOTE>")
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

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo.</BLOCKQUOTE>")
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

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo.</BLOCKQUOTE>")
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

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo.</BLOCKQUOTE>")
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

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo.</BLOCKQUOTE>")
	Response.Write "<BR />"
	lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""UploadEmployeesFeaturesFrm"" ID=""UploadEmployeesFeaturesFrm"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE="&lReasonID&" />"
			Select Case lReasonID
				Case EMPLOYEES_MOTHERAWARD, EMPLOYEES_HELP_COMISSION, EMPLOYEES_SAFEDOWN, EMPLOYEES_FONAC_CONCEPT
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptAmount"" ID=""ConceptAmountHdn"" VALUE=""1"" />"
				Case EMPLOYEES_EFFICIENCY_AWARD
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PayrollDateYYYYMMDD"" ID=""PayrollDateYYYYMMDDHdn"" VALUE=""" & CLng(oRequest("PayrollID").Item) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FileName"" ID=""FileNameHdn"" VALUE=""" & CStr(oRequest("FileName").Item) & """ />"
			End Select
			For iIndex = 1 To iColumns
				Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
				Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""EmployeeID"">Número de empleado</OPTION>"
					Select Case lReasonID
						Case -89, EMPLOYEES_ANUAL_AWARD, EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_CONCEPT_C3, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_GLASSES, EMPLOYEES_MONTHAWARD, EMPLOYEES_MOTHERAWARD, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_FONAC_ADJUSTMENT
							Response.Write "<OPTION VALUE=""PayrollDateYYYYMMDD"">Quincena de aplicación (AAAAMMDD)</OPTION>"
							Response.Write "<OPTION VALUE=""PayrollDateDDMMYYYY"">Quincena de aplicación (DD-MM-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""PayrollDateMMDDYYYY"">Quincena de aplicación (MM-DD-AAAA)</OPTION>"
						Case EMPLOYEES_FONAC_CONCEPT
							Response.Write "<OPTION VALUE=""PayrollDateYYYYMMDD"">Quincena de aplicación (AAAAMMDD)</OPTION>"
							Response.Write "<OPTION VALUE=""PayrollDateDDMMYYYY"">Quincena de aplicación (DD-MM-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""PayrollDateMMDDYYYY"">Quincena de aplicación (MM-DD-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""EndDateYYYYMMDD"">Fecha de término (AAAAMMDD) (opcional)</OPTION>"
							Response.Write "<OPTION VALUE=""EndDateDDMMYYYY"">Fecha de término (DD-MM-AAAA) (opcional)</OPTION>"
							Response.Write "<OPTION VALUE=""EndDateMMDDYYYY"">Fecha de término (MM-DD-AAAA) (opcional)</OPTION>"
						Case EMPLOYEES_NIGHTSHIFTS
							Response.Write "<OPTION VALUE=""StartDateYYYYMMDD"">Fecha del día festivo (AAAAMMDD)   (si es más de uno, separelos con ',')</OPTION>"
							Response.Write "<OPTION VALUE=""StartDateDDMMYYYY"">Fecha del día festivo (DD-MM-AAAA) (si es más de uno, separelos con ',')</OPTION>"
							Response.Write "<OPTION VALUE=""StartDateMMDDYYYY"">Fecha del día festivo (MM-DD-AAAA) (si es más de uno, separelos con ',')</OPTION>"
							Response.Write "<OPTION VALUE=""PayrollDateYYYYMMDD"">Quincena de aplicación (AAAAMMDD)</OPTION>"
							Response.Write "<OPTION VALUE=""PayrollDateDDMMYYYY"">Quincena de aplicación (DD-MM-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""PayrollDateMMDDYYYY"">Quincena de aplicación (MM-DD-AAAA)</OPTION>"
						Case EMPLOYEES_EFFICIENCY_AWARD
						Case EMPLOYEES_GRADE
							Response.Write "<OPTION VALUE=""YearID"">Año</OPTION>"
							Response.Write "<OPTION VALUE=""PayrollDateYYYYMMDD"">Quincena a considerar (AAAAMMDD)</OPTION>"
							Response.Write "<OPTION VALUE=""PayrollDateDDMMYYYY"">Quincena a considerar (DD-MM-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""PayrollDateMMDDYYYY"">Quincena a considerar (MM-DD-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""EmployeeGrade"">Calificación</OPTION>"
						Case Else
							Response.Write "<OPTION VALUE=""StartDateYYYYMMDD"">Fecha de inicio (AAAAMMDD)</OPTION>"
							Response.Write "<OPTION VALUE=""StartDateDDMMYYYY"">Fecha de inicio (DD-MM-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""StartDateMMDDYYYY"">Fecha de inicio (MM-DD-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""EndDateYYYYMMDD"">Fecha de término (AAAAMMDD)</OPTION>"
							Response.Write "<OPTION VALUE=""EndDateDDMMYYYY"">Fecha de término (DD-MM-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""EndDateMMDDYYYY"">Fecha de término (MM-DD-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""PayrollDateYYYYMMDD"">Quincena de aplicación (AAAAMMDD)</OPTION>"
							Response.Write "<OPTION VALUE=""PayrollDateDDMMYYYY"">Quincena de aplicación (DD-MM-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""PayrollDateMMDDYYYY"">Quincena de aplicación (MM-DD-AAAA)</OPTION>"
					End Select
					Select Case lReasonID
						Case -89
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe de deducción no gravables</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de deducción no gravables"
						Case 53 ' EmployeesForRisk
							Response.Write "<OPTION VALUE=""ConceptAmount"">Porcentaje de la prestación</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el porcentaje de la prestación"
						Case EMPLOYEES_ADD_SAFE_SEPARATION
							Response.Write "<OPTION VALUE=""ConceptAmount"">Cantidad (en $ o %)</OPTION>"
							Response.Write "<OPTION VALUE=""ConceptQttyID"">Tipo de unidad de la cantidad (1=$ o 2=%)</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el importe o porcentaje del AE."
						'Case EMPLOYEES_ADDITIONALSHIFT
						'	Response.Write "<OPTION VALUE=""ConceptAmount"">Importe por el turno opcional</OPTION>"
						'	sTextAlert = "No se especificó qué columna contiene el Importe por el turno opcional"
						Case EMPLOYEES_ANTIQUITIES
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe para la compensación por antigüedad</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de la diferencia con el puesto superior"
						Case EMPLOYEES_ANUAL_AWARD
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe del estimulo</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe del estimulo"
						Case EMPLOYEES_BENEFICIARIES
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe para pensió alimenticia</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe para pensió alimenticia"
						Case EMPLOYEES_BENEFICIARIES_DEBIT
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe para adeudo pensión alimenticia</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe para adeudo pensión alimenticia"
						Case EMPLOYEES_CARLOAN
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe del préstamo automóvil servidores públicos de mando superior</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe del préstamo automóvil servidores públicos de mando superior"
						'Case EMPLOYEES_CONCEPT_08
						'	Response.Write "<OPTION VALUE=""ConceptAmount"">Percepcion Adicional</OPTION>"
						'	sTextAlert = "No se especificó qué columna contiene la Percepcion Adicional"
						Case EMPLOYEES_CONCEPT_16
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe de devolución por deducciones indebidas</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de devolución por deducciones indebidas"
						Case EMPLOYEES_CONCEPT_C3
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe de la ayuda para tesis</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de la ayuda para tesis"
						Case EMPLOYEES_CHILDREN_SCHOOLARSHIPS
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe de la beca de hijos de trabajadores</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de la beca de hijos de trabajadores"
						Case EMPLOYEES_EXCENT
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe para otras deducciones</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe para otras deducciones"
						Case EMPLOYEES_FAMILY_DEATH
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe de la ayuda de muerte del familiar</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de la ayuda de muerte del familiar"
						Case EMPLOYEES_FONAC_ADJUSTMENT
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe del ajuste FONAC</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe del ajuste FONAC"
						Case EMPLOYEES_GLASSES
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe de la ayuda de anteojos</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de la ayuda de anteojos"
						Case EMPLOYEES_LICENSES
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe de retenciones por exceso de licencias médicas</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de retenciones por exceso de licencias médicas"
						Case EMPLOYEES_MONTHAWARD
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe del premio al trabajador del mes</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de la ayuda para tesis"
						Case EMPLOYEES_MOTHERAWARD
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe del Premio del 10 de Mayo</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe del Premio del 10 de Mayo"
						Case EMPLOYEES_NON_EXCENT
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe de deducción por cobro de sueldos indebidos</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de deducción por cobro de sueldos indebidos"
						Case EMPLOYEES_NIGHTSHIFTS
							'Response.Write "<OPTION VALUE=""ConceptAmount"">Importe para jornada nocturna adicional por día festivo</OPTION>"
							'sTextAlert = "No se especificó qué columna contiene el Importe para jornada nocturna adicional"
						Case EMPLOYEES_PROFESSIONAL_DEGREE
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe de la ayuda para tesis</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de la ayuda para tesis"
						Case EMPLOYEES_SAFE_SEPARATION
							Response.Write "<OPTION VALUE=""ConceptAmount"">Porcentaje del seguro de separación</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el porcentaje del seguro de separación."
						Case EMPLOYEES_SPORTS_HELP
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe del apoyo al deporte</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe del apoyo al deporte"
						Case EMPLOYEES_SPORTS
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe de la cuota deportiva</OPTION>"
							sTextAlert = "No se especificó qué columna contiene el Importe de la cuota deportiva"
						Case EMPLOYEES_EFFICIENCY_AWARD
							Response.Write "<OPTION VALUE=""ConceptAmount"">Importe del estímulo a la productividad, calidad y eficacia</OPTION>"
							sTextAlert = "No se especificó qué columna contiene del estímulo a la productividad, calidad y eficacia"
					End Select
					Select Case lReasonID
						Case EMPLOYEES_NIGHTSHIFTS, EMPLOYEES_EFFICIENCY_AWARD, EMPLOYEES_GRADE
						Case Else
							Response.Write "<OPTION VALUE=""ConceptComments"">Observaciones (opcional)</OPTION>"
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
				Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el número de empleado.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Select Case lReasonID
					Case -89, EMPLOYEES_ANUAL_AWARD, EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_CONCEPT_C3, EMPLOYEES_EXCENT, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_GLASSES, EMPLOYEES_MONTHAWARD, EMPLOYEES_MOTHERAWARD, EMPLOYEES_NON_EXCENT, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_FONAC_ADJUSTMENT
						Response.Write "if ((sFields.search(/PayrollDateYYYYMMDD/gi) == -1) && (sFields.search(/PayrollDateDDMMYYYY/gi) == -1) && (sFields.search(/PayrollDateMMDDYYYY/gi) == -1)) {" & vbNewLine
							Response.Write "alert('No se especificó qué columna contiene la quincena de aplicación.');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (((sFields.search(/PayrollDateYYYYMMDD/gi) != -1) && ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) || (sFields.search(/PayrollDateMMDDYYYY/gi) != -1))) || ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) && ((sFields.search(/PayrollDateYYYYMMDD/gi) != -1) || (sFields.search(/PayrollDateMMDDYYYY/gi) != -1))) || ((sFields.search(/PayrollDateMMDDYYYY/gi) != -1) && ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) || (sFields.search(/PayrollDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
							Response.Write "alert('No puede seleccionar más de una vez la quincena de aplicación con diferente formato.');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					Case EMPLOYEES_FONAC_CONCEPT
						Response.Write "if ((sFields.search(/PayrollDateYYYYMMDD/gi) == -1) && (sFields.search(/PayrollDateDDMMYYYY/gi) == -1) && (sFields.search(/PayrollDateMMDDYYYY/gi) == -1)) {" & vbNewLine
							Response.Write "alert('No se especificó qué columna contiene la quincena de aplicación.');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (((sFields.search(/PayrollDateYYYYMMDD/gi) != -1) && ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) || (sFields.search(/PayrollDateMMDDYYYY/gi) != -1))) || ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) && ((sFields.search(/PayrollDateYYYYMMDD/gi) != -1) || (sFields.search(/PayrollDateMMDDYYYY/gi) != -1))) || ((sFields.search(/PayrollDateMMDDYYYY/gi) != -1) && ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) || (sFields.search(/PayrollDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
							Response.Write "alert('No puede seleccionar más de una vez la quincena de aplicación con diferente formato.');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					Case EMPLOYEES_NIGHTSHIFTS
						Response.Write "if ((sFields.search(/StartDateYYYYMMDD/gi) == -1) && (sFields.search(/StartDateDDMMYYYY/gi) == -1) && (sFields.search(/StartDateMMDDYYYY/gi) == -1)) {" & vbNewLine
							Response.Write "alert('No se especificó qué columna contiene la fecha del día festivo.');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (((sFields.search(/StartDateYYYYMMDD/gi) != -1) && ((sFields.search(/StartDateDDMMYYYY/gi) != -1) || (sFields.search(/StartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/StartDateDDMMYYYY/gi) != -1) && ((sFields.search(/StartDateYYYYMMDD/gi) != -1) || (sFields.search(/StartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/StartDateMMDDYYYY/gi) != -1) && ((sFields.search(/StartDateDDMMYYYY/gi) != -1) || (sFields.search(/StartDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
							Response.Write "alert('No puede seleccionar más de una vez la fecha del día festivo con diferente formato.');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if ((sFields.search(/PayrollDateYYYYMMDD/gi) == -1) && (sFields.search(/PayrollDateDDMMYYYY/gi) == -1) && (sFields.search(/PayrollDateMMDDYYYY/gi) == -1)) {" & vbNewLine
							Response.Write "alert('No se especificó qué columna contiene la quincena de aplicación.');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (((sFields.search(/PayrollDateYYYYMMDD/gi) != -1) && ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) || (sFields.search(/PayrollDateMMDDYYYY/gi) != -1))) || ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) && ((sFields.search(/PayrollDateYYYYMMDD/gi) != -1) || (sFields.search(/PayrollDateMMDDYYYY/gi) != -1))) || ((sFields.search(/PayrollDateMMDDYYYY/gi) != -1) && ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) || (sFields.search(/PayrollDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
							Response.Write "alert('No puede seleccionar más de una vez la quincena de aplicación con diferente formato.');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					Case EMPLOYEES_EFFICIENCY_AWARD
					Case EMPLOYEES_GRADE
						Response.Write "if (sFields.search(/YearID/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se especificó qué columna contiene el año de la calificación.');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if ((sFields.search(/PayrollDateYYYYMMDD/gi) == -1) && (sFields.search(/PayrollDateDDMMYYYY/gi) == -1) && (sFields.search(/PayrollDateMMDDYYYY/gi) == -1)) {" & vbNewLine
							Response.Write "alert('No se especificó qué columna contiene la quincena de aplicación.');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (((sFields.search(/PayrollDateYYYYMMDD/gi) != -1) && ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) || (sFields.search(/PayrollDateMMDDYYYY/gi) != -1))) || ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) && ((sFields.search(/PayrollDateYYYYMMDD/gi) != -1) || (sFields.search(/PayrollDateMMDDYYYY/gi) != -1))) || ((sFields.search(/PayrollDateMMDDYYYY/gi) != -1) && ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) || (sFields.search(/PayrollDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
							Response.Write "alert('No puede seleccionar más de una vez la quincena de aplicación con diferente formato.');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					Case Else
						Response.Write "if ((sFields.search(/StartDateYYYYMMDD/gi) == -1) && (sFields.search(/StartDateDDMMYYYY/gi) == -1) && (sFields.search(/StartDateMMDDYYYY/gi) == -1)) {" & vbNewLine
							Response.Write "alert('No se especificó qué columna contiene la fecha de inicio.');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (((sFields.search(/StartDateYYYYMMDD/gi) != -1) && ((sFields.search(/StartDateDDMMYYYY/gi) != -1) || (sFields.search(/StartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/StartDateDDMMYYYY/gi) != -1) && ((sFields.search(/StartDateYYYYMMDD/gi) != -1) || (sFields.search(/StartDateMMDDYYYY/gi) != -1))) || ((sFields.search(/StartDateMMDDYYYY/gi) != -1) && ((sFields.search(/StartDateDDMMYYYY/gi) != -1) || (sFields.search(/StartDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
							Response.Write "alert('No puede seleccionar más de una vez la fecha de inicio con diferente formato.');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if ((sFields.search(/EndDateYYYYMMDD/gi) == -1) && (sFields.search(/EndDateDDMMYYYY/gi) == -1) && (sFields.search(/EndDateMMDDYYYY/gi) == -1)) {" & vbNewLine
							Response.Write "alert('No se especificó qué columna contiene la fecha de término.');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (((sFields.search(/EndDateYYYYMMDD/gi) != -1) && ((sFields.search(/EndDateDDMMYYYY/gi) != -1) || (sFields.search(/EndDateMMDDYYYY/gi) != -1))) || ((sFields.search(/EndDateDDMMYYYY/gi) != -1) && ((sFields.search(/EndDateYYYYMMDD/gi) != -1) || (sFields.search(/EndDateMMDDYYYY/gi) != -1))) || ((sFields.search(/EndDateMMDDYYYY/gi) != -1) && ((sFields.search(/EndDateDDMMYYYY/gi) != -1) || (sFields.search(/EndDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
							Response.Write "alert('No puede seleccionar más de una vez la fecha de término con diferente formato.');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if ((sFields.search(/PayrollDateYYYYMMDD/gi) == -1) && (sFields.search(/PayrollDateDDMMYYYY/gi) == -1) && (sFields.search(/PayrollDateMMDDYYYY/gi) == -1)) {" & vbNewLine
							Response.Write "alert('No se especificó qué columna contiene la quincena de aplicación.');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (((sFields.search(/PayrollDateYYYYMMDD/gi) != -1) && ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) || (sFields.search(/PayrollDateMMDDYYYY/gi) != -1))) || ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) && ((sFields.search(/PayrollDateYYYYMMDD/gi) != -1) || (sFields.search(/PayrollDateMMDDYYYY/gi) != -1))) || ((sFields.search(/PayrollDateMMDDYYYY/gi) != -1) && ((sFields.search(/PayrollDateDDMMYYYY/gi) != -1) || (sFields.search(/PayrollDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
							Response.Write "alert('No puede seleccionar más de una vez la quincena de aplicación con diferente formato.');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
				End Select
				Select Case lReasonID
					Case EMPLOYEES_MOTHERAWARD, EMPLOYEES_HELP_COMISSION, EMPLOYEES_SAFEDOWN, EMPLOYEES_FONAC_CONCEPT, EMPLOYEES_GRADE
					Case EMPLOYEES_CONCEPT_08, EMPLOYEES_ADDITIONALSHIFT, EMPLOYEES_NIGHTSHIFTS
					Case EMPLOYEES_ADD_SAFE_SEPARATION
						Response.Write "if (sFields.search(/ConceptQttyID/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se especificó qué columna contiene el tipo de cantidad ($ o %).');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					Case Else
						Response.Write "if (sFields.search(/ConceptAmount/gi) == -1) {" & vbNewLine
							Response.Write "alert('" & sTextAlert & "');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
				End Select
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

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo.</BLOCKQUOTE>")
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

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo.</BLOCKQUOTE>")
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

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo.</BLOCKQUOTE>")
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
					Response.Write "alert('No se especificó qué columna contiene el importe.');" & vbNewLine
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

Function DisplayEmployeesSpecialJourneysColumns(lReasonID, sFileName, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesSpecialJourneysColumns"
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
				Response.Write "if (sFields.search(/AbsenceID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene los tipos de ausencia.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
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

	DisplayEmployeesSpecialJourneysColumns = lErrorNumber
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

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo.</BLOCKQUOTE>")
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

		Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo. <BR /> * Información requerida.</BLOCKQUOTE>")
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
								Response.Write "<OPTION VALUE=""StatusID"">Clave del estatus de la plaza</OPTION>"
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
								Response.Write "<OPTION VALUE=""JourneyID"">Jornada</OPTION>"
								Response.Write "<OPTION VALUE=""ServiceID"">Servicio</OPTION>"
								Response.Write "<OPTION VALUE=""OcurredStartDateYYYYMMDD"">Fecha de inicio de vigencia* (AAAAMMDD)</OPTION>"
								Response.Write "<OPTION VALUE=""OcurredStartDateDDMMYYYY"">Fecha de inicio de vigencia* (DD-MM-AAAA)</OPTION>"
								Response.Write "<OPTION VALUE=""OcurredStartDateMMDDYYYY"">Fecha de inicio de vigencia* (MM-DD-AAAA)</OPTION>"
							Case 59
								Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
								Response.Write "<OPTION VALUE=""JobID"">Plaza (Opcional)</OPTION>"
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
								Response.Write "<OPTION VALUE=""WorkingHours"">Jornada</OPTION>"
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
				If (strComp(oRequest("Action").Item, "ProcessForSar",VbBinaryCompare) = 0) Then
					Response.Write "if (sFields.search(/SocietyID/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con el identificador de sociedad');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
					Response.Write "if (sFields.search(/CompanyID/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con la empresa');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
					Response.Write "if (sFields.search(/PeriodID/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con el bimestre');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
					Response.Write "if (sFields.search(/CLC/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con el número de CLC');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
					Response.Write "if (sFields.search(/BankID/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con la clave del banco');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
					Response.Write "if (sFields.search(/PaymentDate/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con la fecha de pago');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
					Response.Write "if (sFields.search(/EmployeeType/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con el tipo de empleado');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
					Response.Write "if (sFields.search(/Income/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con los ingresos');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
					Response.Write "if (sFields.search(/Deductions/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con las deducciones');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
					Response.Write "if (sFields.search(/NetIncome/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con el ingreso líquido');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
					Response.Write "if (sFields.search(/Cpt_01/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con el importe por concepto 01');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
					Response.Write "if (sFields.search(/Cpt_04/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con el importe por concepto 04');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
					Response.Write "if (sFields.search(/Cpt_05/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con el importe por concepto 05');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
					Response.Write "if (sFields.search(/Cpt_06/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con el importe por concepto 06');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
					Response.Write "if (sFields.search(/Cpt_07/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con el importe por concepto 07');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
					Response.Write "if (sFields.search(/Cpt_08/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con el importe por concepto 08');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
					Response.Write "if (sFields.search(/Cpt_11/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con el importe por concepto 11');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
					Response.Write "if (sFields.search(/Cpt_44/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con el importe por concepto 44');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
					Response.Write "if (sFields.search(/Cpt_b2/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con el importe por concepto B2');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
					Response.Write "if (sFields.search(/Cpt_7s/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con el importe por concepto 7S');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
				Else
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
							Response.Write "if ((sFields.search(/AreaID/gi) == -1) && (sFields.search(/PaymentCenterID/gi) == -1) && (sFields.search(/ServiceID/gi) == -1) && (sFields.search(/ShiftID/gi) == -1) && (sFields.search(/StatusID/gi) == -1) && (sFields.search(/JourneyID/gi) == -1)) {" & vbNewLine
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
							Response.Write "if (sFields.search(/WorkingHours/gi) == -1) {" & vbNewLine
								Response.Write "alert('No se especificó qué columna contiene la jornada.');" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
					End Select
				End If
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

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo.</BLOCKQUOTE>")
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

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo.</BLOCKQUOTE>")
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

Function DisplayPositionsSpecialJourneysLKPColumns(sFileName, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPositionsSpecialJourneysLKPColumns"
	Dim iColumns
	Dim iIndex
	Dim lErrorNumber

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo.</BLOCKQUOTE>")
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
					Response.Write "<OPTION VALUE=""PositionShortName"">Puesto</OPTION>"
					Response.Write "<OPTION VALUE=""LevelID"">Nivel</OPTION>"
					Response.Write "<OPTION VALUE=""WorkingHours"">Horas de trabajo (Jornada)</OPTION>"
					Response.Write "<OPTION VALUE=""ServiceID"">Servicio</OPTION>"
					Response.Write "<OPTION VALUE=""CenterTypeID"">Tipo de centro de trabajo</OPTION>"
					Response.Write "<OPTION VALUE=""StartDateYYYYMMDD"">Fecha de inicio (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""StartDateDDMMYYYY"">Fecha de inicio (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""StartDateMMDDYYYY"">Fecha de inicio (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""EndDateYYYYMMDD"">Fecha de termino (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""EndDateDDMMYYYY"">Fecha de termino (DD-MM-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""EndDateMMDDYYYY"">Fecha de termino (MM-DD-AAAA)</OPTION>"
					Response.Write "<OPTION VALUE=""IsActive1"">Aplica para guardias   (1-Si; 0-No)</OPTION>"
					Response.Write "<OPTION VALUE=""IsActive2"">Aplica para suplencias (1-Si; 0-No)</OPTION>"
					Response.Write "<OPTION VALUE=""IsActive3"">Aplica para rezago q.  (1-Si; 0-No)</OPTION>"
					Response.Write "<OPTION VALUE=""IsActive4"">Aplica para PROVAC     (1-Si; 0-No)</OPTION>"
				Response.Write "</SELECT>"
				Response.Write "<BR />"
			Next
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "SendURLValuesToForm('Column1=PositionShortName&Column2=LevelID&Column3=WorkingHours&Column4=ServiceID&Column5=CenterTypeID&Column6=StartDateYYYYMMDD&Column7=EndDateYYYYMMDD&Column8=IsActive1&Column9=IsActive2&Column10=IsActive3&Column11=IsActive4', document.UploadAbsencesFrm);" & vbNewLine
			Response.Write "//--></SCRIPT>"
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
				Response.Write "if (sFields.search(/PositionShortName/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el código del puesto.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/LevelID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el nivel del puesto.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/WorkingHours/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene las horas de trabajo (Jornada) del puesto.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/ServiceID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el servicio.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/CenterTypeID/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna contiene el tipo de centro de trabajo.');" & vbNewLine
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
				Response.Write "if (((sFields.search(/EndDateYYYYMMDD/gi) != -1) && ((sFields.search(/EndDateDDMMYYYY/gi) != -1) || (sFields.search(/EndDateMMDDYYYY/gi) != -1))) || ((sFields.search(/EndDateDDMMYYYY/gi) != -1) && ((sFields.search(/EndDateYYYYMMDD/gi) != -1) || (sFields.search(/EndDateMMDDYYYY/gi) != -1))) || ((sFields.search(/EndDateMMDDYYYY/gi) != -1) && ((sFields.search(/EndDateDDMMYYYY/gi) != -1) || (sFields.search(/EndDateYYYYMMDD/gi) != -1)))) {" & vbNewLine
					Response.Write "alert('No puede seleccionar más de una vez la fecha de termino con diferente formato.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/IsActive1/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna indica si aplica para guardias.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/IsActive2/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna indica si aplica para suplencias.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/IsActive3/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna indica si aplica para rezago q.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (sFields.search(/IsActive4/gi) == -1) {" & vbNewLine
					Response.Write "alert('No se especificó qué columna indica si aplica para PROVAC.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of CheckColumnsToUpload" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	DisplayPositionsSpecialJourneysLKPColumns = lErrorNumber
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

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo. <BR /> * Información requerida.</BLOCKQUOTE>")
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
					Case 1, 2, 3, 4, 5, 6, 7, 8, 10, 62, 63, 66, 78, 79, 80, 81
						Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
						Response.Write "<OPTION VALUE=""EmployeeID"">No. de empleado*</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredStartDateYYYYMMDD"">Fecha de baja* (AAAAMMDD)</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredStartDateDDMMYYYY"">Fecha de baja* (DD-MM-AAAA)</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredStartDateMMDDYYYY"">Fecha de baja* (MM-DD-AAAA)</OPTION>"
					Case 12, 13
						Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
						Response.Write "<OPTION VALUE=""EmployeeID"">No. de empleado*</OPTION>"
						Response.Write "<OPTION VALUE=""JobID"">Número de plaza*</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredStartDateYYYYMMDD"">Fecha de inicio de vigencia* (AAAAMMDD)</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredStartDateDDMMYYYY"">Fecha de inicio de vigencia* (DD-MM-AAAA)</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredStartDateMMDDYYYY"">Fecha de inicio de vigencia* (MM-DD-AAAA)</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredEndDateYYYYMMDD"">Fecha de fin de vigencia (AAAAMMDD)</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredEndDateDDMMYYYY"">Fecha de fin de vigencia (DD-MM-AAAA)</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredEndDateMMDDYYYY"">Fecha de fin de vigencia (MM-DD-AAAA)</OPTION>"
						Response.Write "<OPTION VALUE=""EmployeeName"">Nombre</OPTION>"
						Response.Write "<OPTION VALUE=""EmployeeLastName"">Apellido paterno</OPTION>"
						Response.Write "<OPTION VALUE=""EmployeeLastName2"">Apellido materno</OPTION>"
						Response.Write "<OPTION VALUE=""RFC"">RFC</OPTION>"
						Response.Write "<OPTION VALUE=""CURP"">CURP</OPTION>"
						Response.Write "<OPTION VALUE=""SocialSecurityNumber"">No. Seg. Social</OPTION>"
						Response.Write "<OPTION VALUE=""CountryID"">Clave de la nacionalidad*</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredBirthDateYYYYMMDD"">Fecha de nacimiento (AAAAMMDD)</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredBirthDateDDMMYYYY"">Fecha de nacimiento (DD-MM-AAAA)</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredBirthDateMMDDYYYY"">Fecha de nacimiento (MM-DD-AAAA)</OPTION>"
						Response.Write "<OPTION VALUE=""MaritalStatusID"">Clave del estado civil*</OPTION>"
						Response.Write "<OPTION VALUE=""EmployeeAddress"">Domicilio (calle, número y colonia)*</OPTION>"
						Response.Write "<OPTION VALUE=""EmployeeCity"">Delegación o municipio*</OPTION>"
						Response.Write "<OPTION VALUE=""EmployeeZipCode"">Código postal*</OPTION>"
						Response.Write "<OPTION VALUE=""StateID"">Clave del estado*</OPTION>"
						Response.Write "<OPTION VALUE=""EmployeeEmail"">Correo electrónico</OPTION>"
						Response.Write "<OPTION VALUE=""EmployeePhone"">Teléfono casa</OPTION>"
						Response.Write "<OPTION VALUE=""OfficePhone"">Teléfono oficina</OPTION>"
						Response.Write "<OPTION VALUE=""OfficeExt"">Ext. oficina</OPTION>"
						Response.Write "<OPTION VALUE=""DocumentNumber1"">Clave de elector</OPTION>"
						Response.Write "<OPTION VALUE=""DocumentNumber2"">Cédula profesional</OPTION>"
						Response.Write "<OPTION VALUE=""DocumentNumber3"">Número de cartilla militar</OPTION>"
						Response.Write "<OPTION VALUE=""EmployeeActivityID"">Clave de actividad del empleado</OPTION>"
						Response.Write "<OPTION VALUE=""ShiftID"">Clave del horario del empleado</OPTION>"
						Response.Write "<OPTION VALUE=""Comments"">Observaciones del movimiento</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredStartHour3HHMM"">Hora de entrada T/A (HHMM) 24 hrs.</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredStartHour3HH_MM"">Hora de entrada T/A (HH:MM) 24 hrs.</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredEndHour3HHMM"">Hora de salida T/A (HHMM) 24 hrs.</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredEndHour3HH_MM"">Hora de salida T/A (HH:MM) 24 hrs.</OPTION>"
						Response.Write "<OPTION VALUE=""RiskLevel"">Nivel de riesgos profesionales (0,1,2)</OPTION>"
					Case 14
						Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
						Response.Write "<OPTION VALUE=""EmployeeID"">No. de empleado*</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredStartDateYYYYMMDD"">Fecha de inicio de vigencia* (AAAAMMDD)</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredStartDateDDMMYYYY"">Fecha de inicio de vigencia* (DD-MM-AAAA)</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredStartDateMMDDYYYY"">Fecha de inicio de vigencia* (MM-DD-AAAA)</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredEndDateYYYYMMDD"">Fecha de fin de vigencia (AAAAMMDD)</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredEndDateDDMMYYYY"">Fecha de fin de vigencia (DD-MM-AAAA)</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredEndDateMMDDYYYY"">Fecha de fin de vigencia (MM-DD-AAAA)</OPTION>"
						Response.Write "<OPTION VALUE=""EmployeeName"">Nombre</OPTION>"
						Response.Write "<OPTION VALUE=""EmployeeLastName"">Apellido paterno</OPTION>"
						Response.Write "<OPTION VALUE=""EmployeeLastName2"">Apellido materno</OPTION>"
						Response.Write "<OPTION VALUE=""RFC"">RFC</OPTION>"
						Response.Write "<OPTION VALUE=""CURP"">CURP</OPTION>"
						Response.Write "<OPTION VALUE=""SocialSecurityNumber"">No. Seg. Social</OPTION>"
						Response.Write "<OPTION VALUE=""CountryID"">Clave de la nacionalidad*</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredBirthDateYYYYMMDD"">Fecha de nacimiento (AAAAMMDD)</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredBirthDateDDMMYYYY"">Fecha de nacimiento (DD-MM-AAAA)</OPTION>"
						Response.Write "<OPTION VALUE=""OcurredBirthDateMMDDYYYY"">Fecha de nacimiento (MM-DD-AAAA)</OPTION>"
						Response.Write "<OPTION VALUE=""MaritalStatusID"">Clave del estado civil*</OPTION>"
						Response.Write "<OPTION VALUE=""EmployeeAddress"">Domicilio (calle, número y colonia)*</OPTION>"
						Response.Write "<OPTION VALUE=""EmployeeCity"">Delegación o municipio*</OPTION>"
						Response.Write "<OPTION VALUE=""EmployeeZipCode"">Código postal*</OPTION>"
						Response.Write "<OPTION VALUE=""StateID"">Clave del estado*</OPTION>"
						Response.Write "<OPTION VALUE=""EmployeeEmail"">Correo electrónico</OPTION>"
						Response.Write "<OPTION VALUE=""EmployeePhone"">Teléfono casa</OPTION>"
						Response.Write "<OPTION VALUE=""OfficePhone"">Teléfono oficina</OPTION>"
						Response.Write "<OPTION VALUE=""OfficeExt"">Ext. oficina</OPTION>"
						Response.Write "<OPTION VALUE=""DocumentNumber1"">Clave de elector</OPTION>"
						Response.Write "<OPTION VALUE=""DocumentNumber2"">Cédula profesional</OPTION>"
						Response.Write "<OPTION VALUE=""DocumentNumber3"">Número de cartilla militar</OPTION>"
						Response.Write "<OPTION VALUE=""EmployeeActivityID"">Clave de actividad del empleado</OPTION>"
						Response.Write "<OPTION VALUE=""Comments"">Observaciones del movimiento</OPTION>"
						Response.Write "<OPTION VALUE=""CompanyID"">Clave de la empresa</OPTION>"
						Response.Write "<OPTION VALUE=""AreaID"">Clave del centro de trabajo</OPTION>"
						Response.Write "<OPTION VALUE=""PaymentCenterID"">Clave del centro de pago</OPTION>"
						Response.Write "<OPTION VALUE=""ServiceID"">Clave del servicio</OPTION>"
						Response.Write "<OPTION VALUE=""JourneyID"">Clave del turno</OPTION>"
						Response.Write "<OPTION VALUE=""ShiftID"">Clave del horario</OPTION>"
						Response.Write "<OPTION VALUE=""ConceptAmount"">Monto mensual</OPTION>"
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
						Case 12, 13
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
						Case 14
							Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
								Response.Write "alert('No se especificó qué columna contiene el número de empleado.');" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
							Response.Write "if (sFields.search(/CompanyID/gi) == -1) {" & vbNewLine
								Response.Write "alert('No se especificó qué columna contiene la clave de la empresa.');" & vbNewLine
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

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo (*Información requerida).</BLOCKQUOTE>")
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

	Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo.</BLOCKQUOTE>")
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
	Const S_FUNCTION_NAME = "DisplayUploadForm"
	Dim sRequiredFields
	Dim sNumber
	Dim sMessage
	Dim bShowSection3
	Dim oRecordset
	Dim lErrorNumber
	Dim oRequestForBanks
	Dim iStarPageForBanks

	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		Response.Write "var bReady = false;" & vbNewLine
		Response.Write "var bFileReady = false;" & vbNewLine
		Response.Write "var lFlag = 0;" & vbNewLine
		Response.Write "function CheckPayrollFields(oForm) {" & vbNewLine
			'Response.Write "alert('Entro en funcion: CheckPayrollFields.');" & vbNewLine
			Response.Write "if (oForm) {" & vbNewLine
				Response.Write "switch (lFlag) {" & vbNewLine
					Response.Write "case 1:" & vbNewLine
						Response.Write "if (parseInt(oForm.AppliedDate.value)==-1) {" & vbNewLine
							Response.Write "alert('No existen nóminas abiertas para el registro de movimientos.');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "break;" & vbNewLine
					Response.Write "}" & vbNewLine
			Response.Write "}" & vbNewLine
			Response.Write "return true;" & vbNewLine
		Response.Write "}" & vbNewLine
	Response.Write "//--></SCRIPT>" & vbNewLine

	Select Case sAction
		Case "FONAC", "JobServices", "MedicalAreas", "ProfessionalRisk", "Third"
		Case "ApplyAbsences"
			Response.Write "<IMG SRC=""Images/Crcl1.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: ShowDisplay(document.all['AbsencesFormDiv']); if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };"">Lista de incidencias pendientes por aplicar</A><BR /><BR />"
		Case "ThirdUploadMovements"
			Response.Write "<IMG SRC=""Images/Crcl1.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: ShowDisplay(document.all['AbsencesFormDiv']); if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };"">Seleccione el archivo para buscar los registros</A><BR /><BR />"
		Case "ProcessForSar"
		Case "ServiceSheet"
			Response.Write "<IMG SRC=""Images/Crcl1.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: ShowDisplay(document.all['AbsencesFormDiv']); if(document.all['UploadValidateServiceSheetDiv'] != null) { HideDisplay(document.all['UploadValidateServiceSheetDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };"">Deseo registrar solicitudes de hojas únicas de servicio</A><BR /><BR />"
		Case Else
			If (InStr(sAction, "ConceptsValues") > 0) And CInt(Request.Cookies("SIAP_SubSectionID")) = 32 Then
			ElseIf (lReasonID = CANCEL_EMPLOYEES_CONCEPTS) Or (lReasonID = CANCEL_EMPLOYEES_C04) Then
				Response.Write "<IMG SRC=""Images/Crcl1.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: ShowDisplay(document.all['AbsencesFormDiv']); if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };"">Mostrar la información general del empleado</A><BR /><BR />"
			ElseIf (lReasonID <> 54) And (lReasonID <> 60) And (lReasonID <> 61) Then
				Response.Write "<IMG SRC=""Images/Crcl1.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: ShowDisplay(document.all['AbsencesFormDiv']); if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };"">Deseo registrar la información en línea</A><BR /><BR />"
			End If
	End Select
	Response.Write "<DIV NAME=""AbsencesFormDiv"" ID=""AbsencesFormDiv"">"
		Select Case sAction
			Case "ApplyAbsences"
				Response.Write "<FORM NAME=""UploadValidateInfoFrm"" ID=""UploadValidateInfoFrm"" METHOD=""POST"" onSubmit=""return CheckPayrollFields(this)"">"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & sAction & """ />"
					If CInt(Request.Cookies("SIAP_SectionID")) <> 7 Then
						Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""10"">"
							Response.Write "<TR>"
								Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""AuthorizationFile"" ID=""ModifyBtn"" VALUE=""Aplicar incidencias en proceso"" CLASS=""Buttons""/>"
							Response.Write "</TR>"
'							Response.Write "<TR NAME=""PayrollDateDiv"" ID=""PayrollDateDiv"">"
'								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Quincena de aplicación:&nbsp;</FONT></TD>"
'								Response.Write "<TD><SELECT NAME=""AppliedDate"" ID=""AppliedDate"" SIZE=""1"" CLASS=""Lists"">"
'									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID=1) And (IsClosed<>1)", "PayrollID Desc", "", "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
'								Response.Write "</SELECT></TD>"
'							Response.Write "</TR>"
'							Response.Write "<TR NAME=""SuspensionsDiv"" ID=""SuspensionsDiv"">"
'								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Aplicar solo suspensiones:&nbsp;</FONT><INPUT TYPE=""CHECKBOX"" NAME=""OnlySuspension"" ID=""OnlySuspensionChk"" VALUE=""1""></TD>"
'								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Aplicar solo tipos de registro:&nbsp;</FONT><INPUT TYPE=""CHECKBOX"" NAME=""OnlyAttendanceControl"" ID=""OnlyAttendanceControlChk"" VALUE=""1""></TD>"
'							Response.Write "</TR>"
'						Response.Write "</TABLE>"
					End If
					lErrorNumber = DisplayAbsencesForApplyTable(oRequest, oADODBConnection, False, sErrorDescription)
					If lErrorNumber <> 0 Then
						Response.Write "<BR />"
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						sErrorDescription = ""
					End If
				Response.Write "</FORM>"
				'Response.Write "</DIV>"
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
				If CInt(Request.Cookies("SIAP_SubSectionID")) = 32 Then
				Else
					'aConceptComponent(N_RECORD_ID_CONCEPT) = -1
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
					Select Case lEmployeeTypeID
						Case 0
							sRequiredFields = "Compañía, Clave concepto de pago, Tipo de puesto, Puesto, Nivel, Horas laboradas(Jornada), Zona económica, Monto quincenal en pesos, Fecha de inicio de la vigencia"
						Case 1
							sRequiredFields = "Compañía, Clave concepto de pago, Puesto, Grupo grado y nivel, Clasificación, Integración, Monto quincenal en pesos, Fecha de inicio de la vigencia"
						Case 2
							sRequiredFields = "Compañía, Clave concepto de pago, Tipo de puesto, Puesto, Nivel, Zona económica, Monto quincenal en pesos, Fecha de inicio de la vigencia"
						Case 3
							sRequiredFields = "Compañía, Clave concepto de pago, Puesto, Nivel, Monto quincenal en pesos, Fecha de inicio de la vigencia"
						Case 4, 5, 6
							sRequiredFields = "Compañía, Clave concepto de pago, Puesto, Nivel, Zona económica, Monto quincenal en pesos, Fecha de inicio de la vigencia"
					End Select
					Call DisplayConceptValuesForm(oRequest, oADODBConnection, "UploadInfo.asp", False, lEmployeeTypeID, aConceptComponent, sErrorDescription)
				End If
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
			Case "EmployeesAssignTemporalNumber"
				sRequiredFields = "Nombre del empleado, apellido paterno, apellido materno, clave del tipo de tabulador, RFC, CURP, fecha de nacimiento del empleado"
				Call DisplayEmployeeForm(oRequest, oADODBConnection, GetASPFileName(""), "EmployeesAssignTemporalNumber", ",1,", 67, aEmployeeComponent, sErrorDescription)
			Case "EmployeesAnualAward"
				sRequiredFields = "Número de empleado, la fecha de registro, la fecha de aplicación en nómina y la cantidad."
				Call DisplayEmployeeConceptForm(oRequest, oADODBConnection, GetASPFileName(""), "Step=1&Action=EmployeesAnualAward", "94", aEmployeeComponent, sErrorDescription)
			Case "EmployeesConcepts"
				lErrorNumber = DisplayEmployeeConceptForm(oRequest, oADODBConnection, GetASPFileName(""), "EmployeesConcepts", EMPLOYEES_CONCEPTS, aEmployeeComponent, sErrorDescription)
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
					Case EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_GLASSES, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_MONTHAWARD, EMPLOYEES_NON_EXCENT, EMPLOYEES_EXCENT, EMPLOYEES_CONCEPT_C3, EMPLOYEES_MOTHERAWARD, -89, EMPLOYEES_ANUAL_AWARD, EMPLOYEES_FONAC_ADJUSTMENT, EMPLOYEES_ANTIQUITY_25_AND_30_YEARS
						sRequiredFields = "Número de empleado, quincena de aplicación, importe y observaciones(opcional)"
					Case EMPLOYEES_FONAC_CONCEPT
						sRequiredFields = "Número de empleado, quincena de aplicación, fecha de fin(opcional) y observaciones(opcional)"
					Case EMPLOYEES_SAFE_SEPARATION, EMPLOYEES_CONCEPT_7S, EMPLOYEES_ANTIQUITIES, EMPLOYEES_FOR_RISK, EMPLOYEES_FONAC_CONCEPT
						sRequiredFields = "Número de empleado, fecha de inicio, fecha de fin, quincena de aplicación, importe y observaciones(opcional)"
					Case EMPLOYEES_CONCEPT_08, EMPLOYEES_ADDITIONALSHIFT
						sRequiredFields = "Número de empleado, fecha de inicio, fecha de fin, quincena de aplicación y observaciones"
					Case EMPLOYEES_MOTHERAWARD, EMPLOYEES_HELP_COMISSION, EMPLOYEES_SAFEDOWN, EMPLOYEES_FONAC_CONCEPT
						sRequiredFields = "Número de empleado, fecha de inicio, fecha de fin, quincena de aplicación, y observaciones"
					Case EMPLOYEES_MONTHAWARD, EMPLOYEES_CARLOAN, EMPLOYEES_CONCEPT_C3, EMPLOYEES_LICENSES, EMPLOYEES_CONCEPT_16, EMPLOYEES_HELP_COMISSION, EMPLOYEES_SAFEDOWN, EMPLOYEES_ANUAL_AWARD, EMPLOYEES_SPORTS_HELP, EMPLOYEES_ANTIQUITY_25_AND_30_YEARS
						sRequiredFields = "Número de empleado, fecha de inicio, fecha de fin, quincena de aplicación, importe y observaciones"
					Case -89, EMPLOYEES_NON_EXCENT, EMPLOYEES_EXCENT
						sRequiredFields = "Número de empleado, fecha de inicio, fecha de fin, quincena de aplicación, importe y observaciones"
					Case EMPLOYEES_ANTIQUITIES, EMPLOYEES_FOR_RISK, EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_GLASSES, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_MOTHERAWARD
						sRequiredFields = "Número de empleado, fecha, quincena de aplicación, importe y observaciones"
					Case EMPLOYEES_NIGHTSHIFTS
						sRequiredFields = "Número de empleado, fecha del día festivo y quincena de aplicación<BR /><BR /><B>Nota: </B> Si se va a registrar más de un día festivo indiquelos en la celda correspondiente separando las fechas con comas."
					Case EMPLOYEES_SPORTS
						sRequiredFields = "Número de empleado, fecha de inicio, fecha de fin, quincena de aplicación, importe (con valor 0) y observaciones"
					Case EMPLOYEES_ADD_SAFE_SEPARATION
						sRequiredFields = "Número de empleado, fecha de inicio, fecha de fin, quincena de aplicación, cantidad, tipo de unidad de la cantidad ($ o %) y observaciones"
					Case EMPLOYEES_DOCUMENTS_FOR_LICENSES
						sRequiredFields = "No. del empleado, Fecha del documento, Número de la solicitud, Tipo de licencia, Fecha inicio de la licencia sindical, fecha fin de la licencia sindical, Nombre de la plantilla"
					Case -58
						sRequiredFields = "Número de empleado, clave del concepto, importe reclamado, fecha de omisión, quincena de aplicación y nombre del beneficiario (opcional)"
					Case 1, 2, 3, 4, 5, 7, 6, 8, 10, 62, 63, 66, 78, 79, 80, 81, 101, 102, 103, 104, 105, 106
						sRequiredFields = "Número de empleado, fecha de baja"
					Case 12, 13
						sRequiredFields = "Número de empleado*, número de plaza*, fecha de inicio vigencia*, fecha de fin de vigencia (opc), nombre (opc), apellido paterno (opc), apellido materno (opc), RFC (opc), CURP (opc), número de seguro social (opc), clave de la nacionalidad*, fecha de nacimiento (opc), clave de estado civil*, domicilio (calle, número y colonia)*, delegación o municipio*, código postal*, clave del estado*, correo electrónico (opc), teléfono casa (opc), teléfono oficina (opc), ext. oficina (opc), clave de elector (opc), cédula profesional (opc), número cartilla militar (opc), clave de actividades del empleado (opc), clave del horario del empleado*, observaciones del movimiento (opc), horario entrada turno opcional (opc), horario salida turno opcional (opc), nivel de riesgos profesionales (opc)"
					Case 14
						sRequiredFields = "Número de empleado*, fecha de inicio vigencia*, fecha de fin de vigencia*, nombre (opc), apellido paterno (opc), apellido materno (opc), RFC (opc), CURP (opc), número de seguro social (opc), clave de la nacionalidad*, fecha de nacimiento (opc), clave de estado civil*, domicilio (calle, número y colonia)*, delegación o municipio*, código postal*, clave del estado*, correo electrónico (opc), teléfono casa (opc), teléfono oficina (opc), ext. oficina (opc), clave de elector (opc), cédula profesional (opc), número cartilla militar (opc), clave de actividades del empleado (opc), observaciones del movimiento (opc), clave de la empresa*, clave de centro de trabajo*, clave de centro de pago*, clave de servicio*, clave de turno*, clave de horario*, monto mensual*"
					Case 17, 18
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
						iStarPageForBanks = CInt(oRequest("StartPage").Item)
						sRequiredFields = "Número de empleado, clave del banco, número de cuenta, quincena de aplicación y fecha de termino (omitir la fecha de termino si la cuenta es por tiempo ilimitado)<BR /><BR /><B>Nota: </B> Si el empleado recibe su pago por cheque, el número de cuenta deberá indicarse con un un punto."
						sErrorDescription = "No se pudo obtener la información de los registros."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Banks Where (BankID>-1) And (Active=1) Order By BankID", "UploadInfoDisplayLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							sRequiredFields = sRequiredFields & "<BR /><B>Bancos: </B>  "
							If Not oRecordset.EOF Then
								Do While Not oRecordset.EOF
									sRequiredFields = sRequiredFields & CleanStringForHTML(CStr(oRecordset.Fields("BankShortName").Value) & ". " & CStr(oRecordset.Fields("BankName").Value)) & ", "
									oRecordset.MoveNext
								Loop
								sRequiredFields = Left(sRequiredFields, (Len(sRequiredFields) - Len(", ")))
							End If
							oRecordset.Close
						End If
					Case EMPLOYEES_GRADE
						sRequiredFields = "Número de empleado, año, quincena a considerar, calificación"
					Case EMPLOYEES_ADD_BENEFICIARIES, EMPLOYEES_CREDITORS
						sRequiredFields = "No se cuenta con carga masiva para estos módulos"
					Case Else
						sRequiredFields = "Número de empleado, número de plaza, fecha de inicio vigencia, fecha de fin de vigencia (opcional), Horario de entrada 1, Horario de salida 1, horario de entrada 2 (opc), horario de salida 2 (opc), horario turno opcional 1, horario salida turno opcional, clave del servicio, nivel de riesgo"
				End Select
				Select Case lReasonID
					Case 54
					Case ALIMONY_TYPES, CREDITORS_TYPES
						Call DisplayAlimonyTypesForm(oRequest, oADODBConnection, GetASPFileName(""), lReasonID, aEmployeeComponent, sErrorDescription)
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
						sRequiredFields = "Centro de trabajo, centro de pago, clave de servicio, clave del puesto, clave del nivel ó (clave de GGN, clave de integración y clave de clasificación), clave del tipo de ocupación, clave de horario, clave de turno, jornada, fecha de inicio de vigencia, fecha de fin de vigencia (opcional)"
						Call DisplayJobForm(oRequest, oADODBConnection, GetASPFileName(""), aJobComponent, sErrorDescription)
					Case 60
						sRequiredFields = "Número de plaza*, clave del centro de trabajo, clave del centro de pago, clave del servicio, clave del turno, clave del horario, clave de estatus de la plaza (O,V,C,L,E,CA,LI), fecha de inicio de vigencia* (*Campos requeridos)"
					Case 61
						sRequiredFields = "Número de plaza*, fecha de inicio*, clave del puesto*, (nivel ó si es plaza de funcionario: clave de GGN, clave de clasificación, clave de integración)*, Jornada, Servicio (*requeridos)"
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
			Case "PayrollRevision"
				lErrorNumber = DisplayPayrollRevisionForm(oRequest, oADODBConnection, aPayrollRevisionComponent, sErrorDescription)
			Case "PersonalLoan"
				Call DisplayEmployeeConceptForm(oRequest, oADODBConnection, GetASPFileName(""), "Step=1&Action=MediumTermLoan", "61", aEmployeeComponent, sErrorDescription)
			Case "ProfessionalRisk"
				sRequiredFields = "Rama*, Tipo de centro de trabajo*, Clave del puesto*, Clave del servicio*, Monto de riesgo*, fecha de inicio de vigencia* (*Campos requeridos)"
			Case "ResumptionOfWork"
				Call DisplayEmployeeForm(oRequest, oADODBConnection, "Employees.asp", "ResumptionOfWork", ",0,", lReasonID, aEmployeeComponent, sErrorDescription)
			Case "ThirdUploadMovements"
				lErrorNumber = DisplayEmployeesCreditsSearchForm(oRequest, oADODBConnection, GetASPFileName(""), False, sErrorDescription)
			Case "ProcessForSar"
'				If StrComp(oRequest("Load").Item, "PayrollSummary", vbBinaryCompare) = 0 Then
'					sRequiredFields = "Sociedad*, Empresa*, Bimestre*, CLC*, Banco*, Fecha de pago*, Tipo de tabulador*, Percepciones*, Deducciones*, Líquido*, Cpto_01*, Cpto_04*, Cpto_05*, Cpto_06*, Cpto_07*, Cpto_08*, Cpto_11*, Cpto_44*, Cpto_b2*, Cpto_7s* (*Campos requeridos)"
'				ElseIf (StrComp(oRequest("Load").Item, "BanamexCensus", vbBinaryCompare) = 0) Or _
'						StrComp(oRequest("Load").Item, "ConsarFile", vbBinaryCompare) = 0) Then
'					Response.Write "<IFRAME SRC=""BrowserFile.asp?Action=" & oRequest("Load").Item & "&UserID=" & aLoginComponent(N_USER_ID_LOGIN) & """ NAME=""UploadInfoIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""96""></IFRAME>"
'				End If
			Case "ServiceSheet"
				Call DisplayEmployeeForm(oRequest, oADODBConnection, "Main_ISSSTE.asp", "ServiceSheet", -1, EMPLOYEES_SERVICE_SHEET, aEmployeeComponent, sErrorDescription)
			Case Else
				'sRequiredFields = "Numero de empleado, clave de incidencia, fecha de ocurrencia, fecha de aplicación, Observaciones"
				sRequiredFields = "Número de empleado, clave de incidencia, fecha de ocurrencia, fecha de termino(opcional), quincena de aplicación, incidencia a justificar(opcional, cuando se agrega justificación), observaciones(opcional), periodo de vacaciones/estimulo(opcional) y año del periodo de vacaciones/estimulo(opcional)"
				Call DisplayAbsenceForm(oRequest, oADODBConnection, GetASPFileName(""), lReasonID, "", aAbsenceComponent, sErrorDescription)
		End Select
	Response.Write "</DIV>"
	Select Case sAction
		Case "FONAC", "ApplyAbsences", "PayrollRevision", "EmployeesConcepts", "ServiceSheet"
		Case "MedicalAreas", "JobServices", "Third"
			Response.Write "<IMG SRC=""Images/Crcl1.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; ShowDisplay(document.all['UploadInfoFormDiv']); if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };""><FONT FACE=""Arial"" SIZE=""2"">Deseo subir la información a través de un archivo</FONT></A><BR /><BR />"
			Response.Write "<DIV NAME=""UploadInfoFormDiv"" ID=""UploadInfoFormDiv"">"
		Case "ThirdUploadMovements"
			Response.Write "<DIV NAME=""UploadInfoFormDiv"" ID=""UploadInfoFormDiv"" STYLE=""display: none"">"
		Case "ProcessForSar"
'			Response.Write "<IMG SRC=""Images/Crcl1.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; ShowDisplay(document.all['UploadInfoFormDiv']); if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };""><FONT FACE=""Arial"" SIZE=""2"">Deseo subir la información a través de un archivo</FONT></A><BR /><BR />"
'			Response.Write "<DIV NAME=""UploadInfoFormDiv"" ID=""UploadInfoFormDiv"">"
'			Call DisplayInstructionsMessage("Paso 1. Introduzca el archivo a utilizar", "<BLOCKQUOTE><OL>" & _
'				"<LI>Abra el documento de Excel con la información que desea subir.</LI>" & _
'				"<LI>Copie únicamente las celdas que contienen la información deseada.</LI>" & _
'				"<LI>Pegue dicha información en la caja de texto.</LI>" & _
'				"<LI><B>O seleccione el archivo de texto que contiene la información a subir.</B></LI>" & _
'				"</OL></BLOCKQUOTE>")
		Case "ProfessionalRisk"
			Response.Write "<IMG SRC=""Images/Crcl1.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; ShowDisplay(document.all['UploadInfoFormDiv']); if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };""><FONT FACE=""Arial"" SIZE=""2"">Deseo subir la información a través de un archivo</FONT></A><BR /><BR />"
			Response.Write "<DIV NAME=""UploadInfoFormDiv"" ID=""UploadInfoFormDiv"">"
		Case Else
			Select Case lReasonID
				Case 54, 60, 61
					Response.Write "<IMG SRC=""Images/Crcl1.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; ShowDisplay(document.all['UploadInfoFormDiv']); if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };""><FONT FACE=""Arial"" SIZE=""2"">Deseo subir la información a través de un archivo</FONT></A><BR /><BR />"
					Response.Write "<DIV NAME=""UploadInfoFormDiv"" ID=""UploadInfoFormDiv"">"
				Case 57, 58
				Case 18, 21, 28, 29, 30, 31, 32, 33, 34, 43, 44, 45, 46, 47, 48, 36, 37, 38, 39, 40, 41, 51, 50, 26, 57, 67, EMPLOYEES_ADD_BENEFICIARIES, EMPLOYEES_CREDITORS, EMPLOYEES_THIRD_PROCESS, EMPLOYEES_THIRD_CONCEPT, CANCEL_EMPLOYEES_CONCEPTS, CANCEL_EMPLOYEES_SSI, CANCEL_EMPLOYEES_C04
					Response.Write "<DIV NAME=""UploadInfoFormDiv"" ID=""UploadInfoFormDiv"" STYLE=""display: none"">"
				Case Else
					If (InStr(sAction, "ConceptsValues") > 0) And CInt(Request.Cookies("SIAP_SubSectionID")) = 32 Then
					Else
						Response.Write "<IMG SRC=""Images/Crcl2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; ShowDisplay(document.all['UploadInfoFormDiv']); if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };""><FONT FACE=""Arial"" SIZE=""2"">Deseo subir la información a través de un archivo</FONT></A><BR /><BR />"
						Response.Write "<DIV NAME=""UploadInfoFormDiv"" ID=""UploadInfoFormDiv"" STYLE=""display: none"">"
					End If
			End Select
	End Select
	If (InStr(sAction, "ConceptsValues") > 0) And CInt(Request.Cookies("SIAP_SubSectionID")) = 32 Then
	ElseIf(lReasonID <> 57) And (lReasonID <> 58) And (lReasonID <> 67) And (lReasonID <> EMPLOYEES_SERVICE_SHEET) Then
		Select Case lReasonID
			Case CANCEL_EMPLOYEES_CONCEPTS, CANCEL_EMPLOYEES_C04
			Case EMPLOYEES_THIRD_CONCEPT
				Call DisplayInstructionsMessage("Captura manual de registros", "En esta sección solo está habilitada la captura manual de registros. Si requiere cargar registros por medio de un archivo, ir a la sección correspondiente.")
				Response.Write "<BR />"
			Case Else
				If sAction = "ThirdUploadMovements" Then
					Call DisplayInstructionsMessage("Consulta y activación de registros", "En esta sección solo está habilitada la captura manual de registros. Si requiere cargar registros por medio de un archivo, ir a la sección correspondiente.")
					Response.Write "<BR />"
				ElseIf (sAction = "ApplyAbsences") Or (sAction = "PayrollRevision") Or (sAction = "EmployeesConcepts") Or (sAction = "ServiceSheet") Then
					Response.Write "<BR />"
				Else
					Response.Write "<FORM NAME=""UploadInfoFrm"" ID=""UploadInfoFrm"" METHOD=""POST"" onSubmit=""return bReady"">"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""2"" />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ActionHdn"" VALUE=""" & lReasonID & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Success"" ID=""ActionHdn"" VALUE=""" & lSuccess & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & lEmployeeID & """ />"
						If InStr(1, sAction, "ProcessForSar") > 0 Then 
							If Len(oRequest("Load").Item) > 0 Then
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Load"" ID=""LoadHdn"" VALUE=""" & oRequest("Load").Item & """ />"
							End If
						End If
						If InStr(1, sAction, "ConceptsValues") Then
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDHdn"" VALUE=""" & lEmployeeTypeID & """ />"
						End If
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ErrorDescription"" ID=""EmployeeIDHdn"" VALUE=""" & sError & """ />"
						Select Case sAction
							Case "ApplyAbsences", "ServiceSheet"
							Case "Third"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ThirdConcept"" ID=""ThirdConceptHdn"" VALUE=""" & oRequest("ThirdConcept").Item & """ />"
								Call DisplayInstructionsMessage("Paso 1. Introduzca el archivo a utilizar", "<BLOCKQUOTE><OL>" & _
																	"<LI><B>Seleccione el archivo de texto que contiene la información a subir con el botón Examinar.</B></LI>" & _
																	"<LI><B>De click en el botón Continuar.</B></LI>" & _
																"</OL></BLOCKQUOTE>")
							Case "ProfessionalRisk"
								Call DisplayInstructionsMessage("Paso 1. Introduzca el archivo a utilizar", "<BLOCKQUOTE><OL>" & _
									"<LI>Abra el documento de Excel con la información que desea subir.</LI>" & _
									"<LI>Copie únicamente las celdas que contienen la información deseada.</LI>" & _
									"<LI>Pegue dicha información en la caja de texto.</LI>" & _
									"<LI>O seleccione el archivo de texto que contiene la información a subir. <BR />(Con el botón Examinar para agregar un archivo local o seleccione un archivo previamente transferido de la lista.)</LI>" & _
									"</OL></BLOCKQUOTE>")
							Case Else
								Select Case lReasonID
									Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS
										Call DisplayInstructionsMessage("Paso 1. Introduzca el archivo a utilizar", "<BLOCKQUOTE><OL>" & _
																			"<LI>Abra el documento de Excel con la información que desea subir.</LI>" & _
																			"<LI>Copie únicamente las celdas que contienen la información deseada.</LI>" & _
																			"<LI>Pegue dicha información en la caja de texto.</LI>" & _
																			"<LI>O seleccione el archivo de texto que contiene la información a subir. <BR />(Con el botón Examinar para agregar un archivo local o seleccione un archivo previamente transferido de la lista.)</LI>" & _
																		"</OL></BLOCKQUOTE>")
									Case Else
										Call DisplayInstructionsMessage("Paso 1. Introduzca el archivo a utilizar", "<BLOCKQUOTE><OL>" & _
																			"<LI>Abra el documento de Excel con la información que desea subir.</LI>" & _
																			"<LI>Copie únicamente las celdas que contienen la información deseada.</LI>" & _
																			"<LI>Pegue dicha información en la caja de texto.</LI>" & _
																			"<LI><B>O seleccione el archivo de texto que contiene la información a subir.</B></LI>" & _
																		"</OL></BLOCKQUOTE>")
								End Select
						End Select
						Response.Write "<BR />"
						If Len(sRequiredFields) > 0 Then
							If lReasonID = 0 Then
								Response.Write "<B>Para la asignación de número de empleado se requiere: </B>" & sRequiredFields & "."
							Else
								Response.Write "<B>Para este concepto se requiere: </B>" & sRequiredFields & "."
							End If
							Response.Write "<BR />"
							Response.Write "<BR />"
						End If
						Select Case sAction
							Case "Third"
							Case Else
								Response.Write "<TEXTAREA NAME=""RawData"" ID=""RawDataTxtArea"" ROWS=""10"" COLS=""119"" CLASS=""TextFields"" onChange=""bReady = (this.value != '')""></TEXTAREA><BR /><BR />"
								Response.Write "<INPUT TYPE=""SUBMIT"" VALUE=""Continuar"" CLASS=""Buttons"" />"
						End Select
					Response.Write "</FORM>"
					Response.Write "<IFRAME SRC=""BrowserFileForInfo.asp?Action=" & oRequest("Action").Item & "&UserID=" & aLoginComponent(N_USER_ID_LOGIN) & """ NAME=""UploadInfoIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""60""></IFRAME>"
					Response.Write "<BR />"
				End If
		End Select
		Select Case lReasonID
			Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS, 300
				Response.Write "<FORM NAME=""UploadFileInfoFrm"" ID=""UploadFileInfoFrm"" METHOD=""POST"" onSubmit=""return (document.UploadFileInfoFrm.UploadFile.value.length>0)"">"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""2"" />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ActionHdn"" VALUE=""" & lReasonID & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Success"" ID=""ActionHdn"" VALUE=""" & lSuccess & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & lEmployeeID & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ThirdConcept"" ID=""ThirdConceptHdn"" VALUE=""" & CStr(oRequest("ThirdConcept").Item) & """ />"

					Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
						Response.Write "<TD COLSPAN=""2"" VALIGN=""TOP"">"
							Response.Write "<B><FONT FACE=""Arial"" SIZE=""2"">Selecciones un archivo previamente transferido:</B><BR /><BR /></FONT>"
						Response.Write "</TD></TR>"
						Response.Write "<TR><TD VALIGN=""TOP"">"
							Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Archivo seleccionado:<BR /></FONT>"
							Response.Write "<TEXTAREA NAME=""UploadFile"" ID=""UploadFileTxtArea"" ROWS=""7"" COLS=""50"" CLASS=""TextFields"" onChange=""bFileReady = (this.value != '')""></TEXTAREA><BR />"
						Response.Write "</TD>"
						Response.Write "<TD>&nbsp;&nbsp;&nbsp;</TD>"
						Response.Write "<TD VALIGN=""TOP"">"
							If lReasonID = 300 Then
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Archivos de prestaciones previamente transferidos:<BR /></FONT>"
								Response.Write "<IFRAME SRC=""BrowserFile.asp?Action=Discos&UserID=" & aLoginComponent(N_USER_ID_LOGIN) & """ NAME=""UploadInfoIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""96""></IFRAME>"
							Else
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Archivos de prestaciones previamente transferidos:<BR /></FONT>"
								Response.Write "<IFRAME SRC=""BrowserFile.asp?Action=Prestaciones&UserID=" & aLoginComponent(N_USER_ID_LOGIN) & """ NAME=""UploadInfoIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""96""></IFRAME>"
							End If
						Response.Write "</TD>"
					Response.Write "</TR></TABLE><BR />"
					Response.Write "<INPUT TYPE=""SUBMIT"" VALUE=""Continuar"" CLASS=""Buttons"">"
				Response.Write "</FORM>"
		End Select
		Response.Write "</DIV>"
	End If
	Select Case sAction
		Case "ApplyAbsences", "ProfessionalRisk"
		Case "Absences"
			Response.Write "<IMG SRC=""Images/Crcl3.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; ShowDisplay(document.all['UploadValidateInfoFormDiv']); if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };"">Incidencias en proceso de aplicación</A><BR /><BR />"
			Response.Write "<DIV NAME=""UploadValidateInfoFormDiv"" ID=""UploadValidateInfoFormDiv"" STYLE=""display: none"">"
				Response.Write "<FORM NAME=""UploadValidateInfoFrm"" ID=""UploadValidateInfoFrm"" METHOD=""POST"" onSubmit=""return CheckPayrollFields(this)"">"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & sAction & """ />"
					aAbsenceComponent(N_ACTIVE_ABSENCE)= 0
					lErrorNumber = DisplayAbsencesTable(oRequest, oADODBConnection, DISPLAY_NOTHING, False, aAbsenceComponent, sErrorDescription)
					If lErrorNumber <> 0 Then
						Response.Write "<BR />"
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						sErrorDescription = ""
					End If
				Response.Write "</FORM>"
			Response.Write "</DIV>"
		Case "ConceptsValues"
			If CInt(Request.Cookies("SIAP_SubSectionID")) = 32 Then
			Else
				Response.Write "<IMG SRC=""Images/Crcl3.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; ShowDisplay(document.all['UploadValidateInfoFormDiv']); if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };"">Consulta los registros en proceso para el tipo de tabulador</A><BR /><BR />"
				Response.Write "<DIV NAME=""UploadValidateInfoFormDiv"" ID=""UploadValidateInfoFormDiv"" STYLE=""display: none"">"
					Response.Write "<FORM NAME=""UploadValidateInfoFrm"" ID=""UploadValidateInfoFrm"" METHOD=""POST"" onSubmit=""return CheckPayrollFields(this)"">"
						Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""10"">"
							Response.Write "<TR>"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & sAction & """ />"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptValuesAction"" ID=""ConceptValuesActionHdn"" VALUE=""1"" />"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDHdn"" VALUE=""" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & """ />"
								If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<TD><INPUT TYPE=""SUBMIT"" NAME=""AuthorizationFile"" ID=""ModifyBtn"" VALUE=""Aplicar los movimientos en proceso"" CLASS=""Buttons""/></TD>"
								Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
								If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<TD><INPUT TYPE=""SUBMIT"" NAME=""RemoveFile"" ID=""RemoveBtn"" VALUE=""Eliminar los movimientos en proceso"" CLASS=""Buttons""/></TD>"
							Response.Write "</TR>"
						Response.Write "</TABLE>"
						aConceptComponent(N_STATUS_ID_CONCEPT) = 0
						lErrorNumber = DisplayConceptValuesTableSP(oRequest, oADODBConnection, lEmployeeTypeID, False, sErrorDescription)
						If lErrorNumber <> 0 Then
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							lErrorNumber = 0
							sErrorDescription = ""
						End If
					Response.Write "</FORM>"
				Response.Write "</DIV>"
			End If
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
		Case "EmployeesConcepts"
			If False Then
				Response.Write "<IMG SRC=""Images/Crcl2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; ShowDisplay(document.all['UploadValidateInfoFormDiv']); if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };"">Prestaciones vigentes del empleado</A><BR /><BR />"
				Response.Write "<DIV NAME=""UploadValidateInfoFormDiv"" ID=""UploadValidateInfoFormDiv"" STYLE=""display: none"">"
					Response.Write "<FORM NAME=""UploadValidateInfoFrm"" ID=""UploadValidateInfoFrm"" METHOD=""POST"" onSubmit=""return bReady"">"
						'lErrorNumber = DisplayPendingEmployeesTable(oRequest, oADODBConnection, False, "EmployeesMovements", lReasonID, 0, aEmployeeComponent, sErrorDescription)
						lErrorNumber = DisplayEmployeesFeaturesTable(oRequest, oADODBConnection, False, sAction, 1, aEmployeeComponent, sErrorDescription)
						If lErrorNumber <> 0 Then
							Response.Write "<BR />"
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							lErrorNumber = 0
							sErrorDescription = ""
						End If
					Response.Write "</FORM>"
				Response.Write "</DIV>"
			End If
		Case "EmployeesMovements", "EmployeesAssignTemporalNumber"
			bShowSection3 = True
			Select Case lReasonID
				Case -58
					sNumber = "3"
					sMessage = "Reclamos de pago en proceso de aplicación"
				Case 17, 18, 21, 26, 28, 29, 30, 31, 32, 33, 34, 36, 37, 38, 39, 40, 41, 43, 44, 45, 46, 47, 48, 50, 51, 67
					sNumber = "2"
					sMessage = "Registros en proceso de aplicación"
				Case 54
					sNumber = "2"
					sMessage = "Consulta de plazas que cambiaron servicio el día de hoy"
				Case 57, 58, 59, EMPLOYEES_DOCUMENTS_FOR_LICENSES, CANCEL_EMPLOYEES_SSI
					bShowSection3 = False
				Case CANCEL_EMPLOYEES_CONCEPTS, CANCEL_EMPLOYEES_C04
					sNumber = "2"
					sMessage = "Consulta de prestaciones para el empleado"
				Case EMPLOYEES_ADD_BENEFICIARIES
					sNumber = "2"
					sMessage = "Consulta de registros de beneficiarios en proceso de los empleados"
				Case EMPLOYEES_BANK_ACCOUNTS
					sNumber = "3"
					sMessage = "Consulta de cuentas bancarias en proceso de aplicación"
				Case EMPLOYEES_BENEFICIARIES_DEBIT
					sNumber = "3"
					sMessage = "Consulta de adeudo de pensión alimenticia por aplicar"
				Case EMPLOYEES_CREDITORS
					sNumber = "2"
					sMessage = "Consulta de registros de acreedores en proceso de los empleados"
				Case EMPLOYEES_GRADE
					sNumber = "3"
					sMessage = "Consulta de registros de calificación en proceso de los empleados"
				Case EMPLOYEES_THIRD_CONCEPT
					sNumber = "2"
					sMessage = "Consulta de registros de terceros capturados manualmente en proceso de aplicación"
				Case Else
					sNumber = "3"
					sMessage = "Registros en proceso de aplicación"
			End Select
			If bShowSection3 Then
				Response.Write "<IMG SRC=""Images/Crcl" & sNumber & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; ShowDisplay(document.all['UploadValidateInfoFormDiv']); if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };""><FONT FACE=""Arial"" SIZE=""2"">" & sMessage & "</FONT></A><BR /><BR />"
			End If
			Response.Write "<DIV NAME=""UploadValidateInfoFormDiv"" ID=""UploadValidateInfoFormDiv"" STYLE=""display: none"">"
				Response.Write "<FORM NAME=""UploadValidateInfoFrm"" ID=""UploadValidateInfoFrm"" METHOD=""POST"">"
					Select Case lReasonID
						Case -89
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 100
						Case -74
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 86
						Case 12
						Case 53
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 4
						Case 54
						Case 59
						Case EMPLOYEES_DOCUMENTS_FOR_LICENSES
						Case EMPLOYEES_SAFE_SEPARATION
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 120
						Case EMPLOYEES_THIRD_CONCEPT
						Case EMPLOYEES_ADD_SAFE_SEPARATION
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 87
						Case EMPLOYEES_ADDITIONALSHIFT
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 7
						Case EMPLOYEES_ANTIQUITIES
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 5
						Case EMPLOYEES_ANUAL_AWARD
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 32
						Case EMPLOYEES_BENEFICIARIES
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 70
						Case EMPLOYEES_BENEFICIARIES_DEBIT
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 86
						Case EMPLOYEES_CARLOAN
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 74
						Case EMPLOYEES_CHILDREN_SCHOOLARSHIPS
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 22
						Case EMPLOYEES_CONCEPT_08
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 8
						Case EMPLOYEES_CONCEPT_16
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 19
						Case EMPLOYEES_EXCENT
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 73
						Case EMPLOYEES_EXTRAHOURS
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 201
							aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE)
						Case EMPLOYEES_NON_EXCENT
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 72
						Case EMPLOYEES_FAMILY_DEATH
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 45
						Case EMPLOYEES_FONAC_CONCEPT
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 77
						Case EMPLOYEES_GLASSES
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 24
						Case EMPLOYEES_HELP_COMISSION
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 63
						Case EMPLOYEES_LICENSES
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 104
						Case EMPLOYEES_MONTHAWARD
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 50
						Case EMPLOYEES_MOTHERAWARD
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 26
						Case EMPLOYEES_NIGHTSHIFTS
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 93
						Case EMPLOYEES_PROFESSIONAL_DEGREE
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 46
						Case EMPLOYEES_SAFEDOWN
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 67
						Case EMPLOYEES_SPORTS_HELP
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 165
						Case EMPLOYEES_SPORTS
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 69
						Case EMPLOYEES_SUNDAYS
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 202
							aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE)
						Case EMPLOYEES_ANTIQUITY_25_AND_30_YEARS
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 44							
					End Select
					Select Case lReasonID
						Case CANCEL_EMPLOYEES_CONCEPTS, CANCEL_EMPLOYEES_C04
						Case Else
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & sAction & """ />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SaveEmployeesMovements"" ID=""SaveEmployeesMovementsHdn"" VALUE=""1"" />"
							Select Case lReasonID
								Case -58
									Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""10"">"
										Response.Write "<TR>"
											If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
												If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_RegistroDeReclamos & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RegistroDeReclamosCyA & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RegistroDeReclamos & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_RegistroDeReclamos & ",", vbBinaryCompare) > 0) Then Response.Write "<TD><INPUT TYPE=""SUBMIT"" NAME=""AuthorizationFile"" ID=""ModifyBtn"" VALUE=""Aplicar reclamos en proceso"" CLASS=""Buttons""/></TD>"
											Else
												If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_RegistroDeReclamos & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RegistroDeReclamosCyA & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RegistroDeReclamos & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_RegistroDeReclamos & ",", vbBinaryCompare) > 0) Then Response.Write "<TD><INPUT TYPE=""SUBMIT"" NAME=""AuthorizationFile"" ID=""ModifyBtn"" VALUE=""Aplicar los relcamos en proceso del empleado "" CLASS=""Buttons""/></TD>"
											End If
											Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
											If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
												If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_RegistroDeReclamos & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RegistroDeReclamosCyA & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RegistroDeReclamos & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_RegistroDeReclamos & ",", vbBinaryCompare) > 0) Then Response.Write "<TD><INPUT TYPE=""SUBMIT"" NAME=""RemoveFile"" ID=""RemoveBtn"" VALUE=""Eliminar reclamos en proceso"" CLASS=""Buttons""/></TD>"
											Else
												If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_RegistroDeReclamos & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RegistroDeReclamosCyA & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RegistroDeReclamos & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_RegistroDeReclamos & ",", vbBinaryCompare) > 0) Then Response.Write "<TD><INPUT TYPE=""SUBMIT"" NAME=""RemoveFile"" ID=""RemoveBtn"" VALUE=""Eliminar reclamos en proceso del empleado"" CLASS=""Buttons""/></TD>"
											End If
										Response.Write "</TR>"
									Response.Write "</TABLE>"
								Case EMPLOYEES_EXTRAHOURS
									If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
										If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_09_RemuneracionPorHorasExtraordinarias & ",", vbBinaryCompare) > 0) And (Request.Cookies("SIAP_SectionID") <> 7) Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""AuthorizationFile"" ID=""ModifyBtn"" VALUE=""Aplicar Horas extras en proceso"" CLASS=""Buttons""/>"
									Else
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_EMPLOYEE) & """ />"
										If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_09_RemuneracionPorHorasExtraordinarias & ",", vbBinaryCompare) > 0) And (Request.Cookies("SIAP_SectionID") <> 7) Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""AuthorizationFile"" ID=""ModifyBtn"" VALUE=""Aplicar Horas extras para este empleado"" CLASS=""Buttons""/>"
									End If
								Case EMPLOYEES_SUNDAYS
									If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
										If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_14_PrimasDominicales & ",", vbBinaryCompare) > 0) And (Request.Cookies("SIAP_SectionID") <> 7) Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""AuthorizationFile"" ID=""ModifyBtn"" VALUE=""Aplicar Primas dominicales en proceso"" CLASS=""Buttons""/>"
									Else
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_EMPLOYEE) & """ />"
										If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_14_PrimasDominicales & ",", vbBinaryCompare) > 0) And (Request.Cookies("SIAP_SectionID") <> 7) Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""AuthorizationFile"" ID=""ModifyBtn"" VALUE=""Aplicar Primas dominicales para este empleado"" CLASS=""Buttons""/>"
									End If
								Case Else
									If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
										If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) And (Request.Cookies("SIAP_SectionID") <> 7) Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""AuthorizationFile"" ID=""ModifyBtn"" VALUE=""Aplicar Movimientos Seleccionados"" CLASS=""Buttons""/>"
									Else
										If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) And (Request.Cookies("SIAP_SectionID") <> 7) Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""AuthorizationFile"" ID=""ModifyBtn"" VALUE=""Aplicar Movimientos del empleado"" CLASS=""Buttons""/>"
									End If
							End Select
					End Select
					If (CInt(Request.Cookies("SIAP_SectionID")) <> 7) Then
						Response.Write "<BR /><BR />"
					End If
					Select Case lReasonID
						Case EMPLOYEES_BANK_ACCOUNTS
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE="&lReasonID&" />"
							'lErrorNumber = DisplayPendingEmployeesTable(oRequest, oADODBConnection, False, sAction, lReasonID, 0, aEmployeeComponent, sErrorDescription)
							aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0
							If (CInt(oRequest("RowsType").Item) = 1) Then
								iStarPageForBanks = 0
							End If
							lErrorNumber = DisplayEmployeesBanksAccountsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, False, iStarPageForBanks, aEmployeeComponent, sErrorDescription)
						Case EMPLOYEES_GRADE
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE="&lReasonID&" />"
							aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0
							'If (CInt(oRequest("RowsType").Item) = 1) Then
							'	iStarPageForBanks = 0
							'End If
							lErrorNumber = DisplayEmployeesGradesTable(oRequest, oADODBConnection, False, aEmployeeComponent, sErrorDescription)
						Case CANCEL_EMPLOYEES_CONCEPTS, CANCEL_EMPLOYEES_C04
							aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) = "1"
							lErrorNumber = DisplayPendingEmployeesConceptsTable(oRequest, oADODBConnection, False, sAction, lReasonID, aEmployeeComponent, sErrorDescription)
						Case EMPLOYEES_THIRD_CONCEPT
							lErrorNumber = DisplayPendingEmployeesCreditsTable(oRequest, oADODBConnection, 0, False, sAction, lReasonID, aEmployeeComponent, sErrorDescription)
						Case EMPLOYEES_ADD_BENEFICIARIES
							aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0
							lErrorNumber = DisplayPendingEmployeesbeneficiariesTable(oRequest, oADODBConnection, False, sAction, lReasonID, aEmployeeComponent, sErrorDescription)
						Case EMPLOYEES_CREDITORS
							aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0
							lErrorNumber = DisplayEmployeesCreditorsTable(oRequest, oADODBConnection, False, sAction, lReasonID, aEmployeeComponent, sErrorDescription)
						Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE="&lReasonID&" />"
							aAbsenceComponent(N_ACTIVE_ABSENCE)= 0
							lErrorNumber = DisplayAbsencesTable(oRequest, oADODBConnection, DISPLAY_NOTHING, False, aAbsenceComponent, sErrorDescription)
							If lErrorNumber <> 0 Then
								Response.Write "<BR />"
								Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
								lErrorNumber = 0
								sErrorDescription = ""
							End If
						Case Else
							lErrorNumber = DisplayPendingEmployeesTable(oRequest, oADODBConnection, False, sAction, lReasonID, 0, aEmployeeComponent, sErrorDescription)
					End Select
					If lErrorNumber <> 0 Then
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
			Response.Write "<IMG SRC=""Images/Crcl2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; ShowDisplay(document.all['UploadValidateInfoFormDiv']); if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };"">Active los registros cargados desde el archivo del tercero</A><BR /><BR />"
			Response.Write "<DIV NAME=""UploadValidateInfoFormDiv"" ID=""UploadValidateInfoFormDiv"" STYLE=""display: none"">"
				Response.Write "<FORM NAME=""UploadValidateInfoFrm"" ID=""UploadValidateInfoFrm"" METHOD=""POST"" onSubmit=""return CheckPayrollFields(this)"">"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & sAction & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SaveEmployeesMovements"" ID=""SaveEmployeesMovementsHdn"" VALUE=""1"" />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptFileName"" ID=""ConceptFileNameHdn"" VALUE=""" & oRequest("ConceptFileName").Item & """ />"
					Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""10"">"
						Response.Write "<TR NAME=""PayrollDateDiv"" ID=""PayrollDateDiv"">"
							Response.Write "<TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2"">Quincena de aplicación:&nbsp;</FONT>"'</TD>"
							'Response.Write "<TD><SELECT NAME=""AppliedDate"" ID=""AppliedDate"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write "<SELECT NAME=""AppliedDate"" ID=""AppliedDate"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And (IsActive_7=1) And (PayrollTypeID=1)", "PayrollID Desc", "", "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Response.Write "</TR>"
					Response.Write "</TABLE>"
					Response.Write "<BR />"
					Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""10"">"
						Response.Write "<TR>"
							If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then Response.Write "<TD><INPUT TYPE=""SUBMIT"" NAME=""AuthorizationFile"" ID=""ModifyBtn"" VALUE=""Aplicar terceros del archivo seleccionado"" CLASS=""Buttons"" onClick=""lFlag = 1;""/></TD>"
							Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
							If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then Response.Write "<TD><INPUT TYPE=""SUBMIT"" NAME=""RemoveFile"" ID=""RemoveBtn"" VALUE=""Eliminar terceros del archivo seleccionado"" CLASS=""Buttons"" onClick=""lFlag = 2;""/></TD>"
						Response.Write "</TR>"
					Response.Write "</TABLE>"
					Response.Write "<BR />"
					lErrorNumber = DisplayPendingEmployeesCreditsTable(oRequest, oADODBConnection, 0, False, sAction, EMPLOYEES_THIRD_PROCESS, aEmployeeComponent, sErrorDescription)
					If lErrorNumber <> 0 Then
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						sErrorDescription = ""
					End If
				Response.Write "</FORM>"
			Response.Write "</DIV>"
		Case "PayrollRevision"
			Response.Write "<IMG SRC=""Images/Crcl2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; ShowDisplay(document.all['UploadValidateInfoFormDiv']); if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };"">Consulta los registros de revisión del empleado</A><BR /><BR />"
			Response.Write "<DIV NAME=""UploadValidateInfoFormDiv"" ID=""UploadValidateInfoFormDiv"" STYLE=""display: none"">"
				Response.Write "<FORM NAME=""UploadValidateInfoFrm"" ID=""UploadValidateInfoFrm"" METHOD=""POST"" onSubmit=""return CheckPayrollFields(this)"">"
					lErrorNumber = DisplayPayrollRevisionTable(oRequest, oADODBConnection, False, aPayrollRevisionComponent, sErrorDescription)
					If lErrorNumber <> 0 Then
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						sErrorDescription = ""
					End If
				Response.Write "</FORM>"
			Response.Write "</DIV>"
		Case "ProcessForSar"
				Response.Write "</FORM>"
			Response.Write "</DIV>"
		Case "ServiceSheet"
			Response.Write "<IMG SRC=""Images/Crcl2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadValidateServiceSheetDiv'] != null) { HideDisplay(document.all['UploadValidateServiceSheetDiv']) }; ShowDisplay(document.all['UploadValidateInfoFormDiv']); if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };"">Consulta las hojas de servicio en proceso de autorización</A><BR /><BR />"
			Response.Write "<DIV NAME=""UploadValidateInfoFormDiv"" ID=""UploadValidateInfoFormDiv"" STYLE=""display: none"">"
				Response.Write "<FORM NAME=""UploadValidateInfoFrm"" ID=""UploadValidateInfoFrm"" METHOD=""POST"" onSubmit=""return bReady"">"
					aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0
					lErrorNumber = DisplayEmployeesDocumentsTable(oRequest, oADODBConnection, False, aEmployeeComponent, sErrorDescription)
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
		Case "ServiceSheet"
			Response.Write "<IMG SRC=""Images/Crcl3.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; ShowDisplay(document.all['UploadValidateServiceSheetDiv']); if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };"">Consulta el detalle de la hoja de servicio en proceso de autorización</A><BR /><BR />"
			Response.Write "<DIV NAME=""UploadValidateServiceSheetDiv"" ID=""UploadValidateServiceSheetDiv"" STYLE=""display: none"">"
				Response.Write "<FORM NAME=""UploadValidateServiceSheetFrm"" ID=""UploadValidateServiceSheetFrm"" METHOD=""POST"" onSubmit=""return bReady"">"
					'aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0
					lErrorNumber = BuildReport1203a(oRequest, oADODBConnection, sErrorDescription)
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
		Case "ConceptsValues"
			If CInt(Request.Cookies("SIAP_SubSectionID")) = 32 Then
				Response.Write "<IMG SRC=""Images/Crcl1.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: ShowDisplay(document.all['ConceptInfoFormDiv']);""><FONT FACE=""Arial"" SIZE=""2"">Consulta los registros activos para el tipo de tabulador</FONT></A><BR /><BR />"
			Else
				Response.Write "<IMG SRC=""Images/Crcl4.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; ShowDisplay(document.all['ConceptInfoFormDiv']);""><FONT FACE=""Arial"" SIZE=""2"">Consulta los registros activos para el tipo de tabulador</FONT></A><BR /><BR />"
			End If
			Response.Write "<DIV NAME=""ConceptInfoFormDiv"" ID=""ConceptInfoFormDiv"" STYLE=""display: none"">"
				Response.Write "<FORM NAME=""ConceptInfoFrm"" ID=""ConceptInfoFrm"" ACTION=""UploadInfo.asp"" METHOD=""POST"">"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDHdn"" VALUE=""" & oRequest("EmployeeTypeID").Item & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartPage"" ID=""StartPageHdn"" VALUE=""1"" />"
                    Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PositionID"" ID=""PositionIDHdn"" VALUE=""" & aConceptComponent(N_POSITION_ID_CONCEPT) & """ />"
					Response.Write "<B>Seleccione los datos para filtrar los registros:&nbsp;&nbsp;&nbsp;</B><BR />"
					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" />Mostrar los registros del &nbsp;"
					Response.Write DisplayDateCombos(oRequest("StartForValue1Year").Item, oRequest("StartForValue1Month").Item, oRequest("StartForValue1Day").Item, "StartForValue1Year", "StartForValue1Month", "StartForValue1Day", N_FORM_START_YEAR, Year(Date()), True, True)
					Response.Write "&nbsp;al&nbsp;"
					Response.Write DisplayDateCombos(oRequest("EndForValueYear").Item, oRequest("EndForValueMonth").Item, oRequest("EndForValueDay").Item, "EndForValueYear", "EndForValueMonth", "EndForValueDay", N_FORM_START_YEAR, Year(Date()), True, True)
					Response.Write "<BR />"
					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" />Mostrar los registros del Puesto:&nbsp;"
					If aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) = -1 Then
						Response.Write "<SELECT NAME=""PositionIDTemp"" ID=""PositionIDTempCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""var asTemp = this.value.split(','); document.ConceptInfoFrm.PositionID.value = asTemp[0];"">"
							Response.Write "<OPTION VALUE=""-1"">Todos</OPTION>"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Positions", "PositionID", "PositionShortName, PositionName, 'Cia:' As Temp, CompanyID", "(PositionID>-1)", "PositionShortName", aConceptComponent(N_POSITION_ID_CONCEPT), "", sErrorDescription)
						Response.Write "</SELECT>"
					ElseIf aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) = 1 Then
						Response.Write "<SELECT NAME=""PositionIDTemp"" ID=""PositionIDTempCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""var asTemp = this.value.split(','); document.ConceptInfoFrm.PositionID.value = asTemp[0];"">"
							Response.Write "<OPTION VALUE=""-1"">Todos</OPTION>"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Positions, GroupGradeLevels", "PositionID, Positions.GroupGradeLevelID, ClassificationID, IntegrationID, 'Cia:' As Temp, CompanyID", "PositionShortName, PositionName, 'GGN:' As Temp1, GroupGradeLevelShortName, 'Clas:' As Temp2, ClassificationID, 'Int:' As Temp3, IntegrationID", "(Positions.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (PositionID>-1) And (EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ")", "PositionShortName", aConceptComponent(N_POSITION_ID_CONCEPT), "", sErrorDescription)
						Response.Write "</SELECT>"
					Else
						Response.Write "<SELECT NAME=""PositionIDTemp"" ID=""PositionIDTempCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""var asTemp = this.value.split(','); document.ConceptInfoFrm.PositionID.value = asTemp[0];"">"
							Response.Write "<OPTION VALUE=""-1"">Todos</OPTION>"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Positions", "PositionID, LevelID", "PositionShortName, PositionName, 'Nivel:' As Temp, LevelID, 'Cia:' As Temp, CompanyID", "(PositionID>-1) And (EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ")", "PositionShortName", aConceptComponent(N_POSITION_ID_CONCEPT) & "," & aConceptComponent(N_LEVEL_ID_CONCEPT), "", sErrorDescription)
						Response.Write "</SELECT>"
					End If
					Response.Write "<BR />"
					Response.Write "<INPUT TYPE=""SUBMIT"" VALUE=""Ver Reporte"" CLASS=""Buttons""><BR />"
					Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""960"" HEIGHT=""1"" /><BR />"
					aConceptComponent(N_STATUS_ID_CONCEPT) = 1
					lErrorNumber = DisplayConceptValuesTableSP(oRequest, oADODBConnection, lEmployeeTypeID, False, sErrorDescription)
					If lErrorNumber <> 0 Then
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						sErrorDescription = ""
					End If
				Response.Write "</FORM>"
			Response.Write "</DIV>"
		Case "EmployeesAssignNumber", "ApplyAbsences"
		Case "Absences"
			Response.Write "<IMG SRC=""Images/Crcl4.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; ShowDisplay(document.all['ConceptInfoFormDiv']);""><FONT FACE=""Arial"" SIZE=""2"">Consulta de incidencias existentes para el empleado</FONT></A><BR /><BR />"
			Response.Write "<DIV NAME=""ConceptInfoFormDiv"" ID=""ConceptInfoFormDiv"" STYLE=""display: none"">"
			Response.Write "<FORM NAME=""ConceptInfoFrm"" ID=""ConceptInfoFrm"" ACTION=""UploadInfo.asp"" METHOD=""POST"">"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_EMPLOYEE) & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AbsenceReview"" ID=""AbsenceReviewHdn"" VALUE=""1"" />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Tab"" ID=""TabHdn"" VALUE=""" & iSelectedTab & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartPage"" ID=""StartPageHdn"" VALUE=""1"" />"
				Response.Write "<B>Seleccione los datos para filtrar el historial:&nbsp;&nbsp;&nbsp;</B><BR />"
				Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" />Mostrar el historial del &nbsp;"
				Response.Write DisplayDateCombos(oRequest("FilterStartYear").Item, oRequest("FilterStartMonth").Item, oRequest("FilterStartDay").Item, "FilterStartYear", "FilterStartMonth", "FilterStartDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "&nbsp;al&nbsp;"
				Response.Write DisplayDateCombos(oRequest("FilterEndYear").Item, oRequest("FilterEndMonth").Item, oRequest("FilterEndDay").Item, "FilterEndYear", "FilterEndMonth", "FilterEndDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "<BR />"
				Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" />Mostrar las incidencias de: <SELECT NAME=""AbsenceID"" ID=""AbsenceIDCmb"" CLASS=""Lists"">"
				Response.Write "<OPTION VALUE=""-1"">Todas las claves</OPTION>"
				Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Absences", "AbsenceID", "AbsenceShortName, AbsenceName", "(AbsenceID>0) And (AbsenceID<100) And (Active=1)", "AbsenceID", aAbsenceComponent(N_ABSENCE_ID_ABSENCE), "", sErrorDescription)
				Response.Write "</SELECT><BR />"
				Response.Write "<INPUT TYPE=""SUBMIT"" VALUE=""Consultar incidencias"" CLASS=""Buttons"" /><BR />"
				Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""960"" HEIGHT=""1"" /><BR />"
				aAbsenceComponent(N_ACTIVE_ABSENCE)= 1
				lErrorNumber = DisplayAbsencesTable(oRequest, oADODBConnection, DISPLAY_NOTHING, False, aAbsenceComponent, sErrorDescription)
				If lErrorNumber <> 0 Then
					Response.Write "<BR />"
					Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
					lErrorNumber = 0
					sErrorDescription = ""
				End If
			Response.Write "</FORM>"
		Case "ServiceSheet"
			Response.Write "<IMG SRC=""Images/Crcl4.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadValidateServiceSheetDiv'] != null) { HideDisplay(document.all['UploadValidateServiceSheetDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; ShowDisplay(document.all['ConceptInfoFormDiv']);""><FONT FACE=""Arial"" SIZE=""2"">Consulta de hojas de servicio aprobadas para el empleado</FONT></A><BR /><BR />"
			Response.Write "<DIV NAME=""ConceptInfoFormDiv"" ID=""ConceptInfoFormDiv"" STYLE=""display: none"">"
				Response.Write "<FORM NAME=""ConceptInfoForm"" ID=""ConceptInfoFrm"" METHOD=""POST"" onSubmit=""return bReady"">"
					If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
						Call DisplayErrorMessage("Mensaje del sistema", "Introduzca un número de empleado para consultar las solicitudes de hojas únicas de servicio.")
					Else
						'lErrorNumber = DisplaySavedZIPReports(oRequest, oADODBConnection, 1203, sErrorDescription)
						aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) = aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) & " And (bPrinted=1)"
						lErrorNumber = DisplayEmployeesDocumentsTable(oRequest, oADODBConnection, False, aEmployeeComponent, sErrorDescription)
						If lErrorNumber <> 0 Then
							Response.Write "<BR />"
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							lErrorNumber = 0
							sErrorDescription = ""
						End If
					End If
				Response.Write "</FORM>"
			Response.Write "</DIV>"
		Case Else
			If ((lReasonID <= -58) And (lReasonID <> EMPLOYEES_THIRD_PROCESS)) Or (lReasonID = EMPLOYEES_FOR_RISK) Then
				Select Case lReasonID
					Case CANCEL_EMPLOYEES_CONCEPTS, CANCEL_EMPLOYEES_C04
						sNumber = "3"
						sMessage = "Consulta de registros cancelados para el empleado"
					Case CANCEL_EMPLOYEES_SSI
						sNumber = "2"
						sMessage = "Consulta de registros existentes para el empleado"
					Case EMPLOYEES_ADD_BENEFICIARIES
						sNumber = "3"
						sMessage = "Consulta de registros de beneficiarios existentes para los empleados"
					Case EMPLOYEES_BANK_ACCOUNTS
						sNumber = "4"
						sMessage = "Consulta de registros existentes para el empleado"
					Case EMPLOYEES_CREDITORS
						sNumber = "3"
						sMessage = "Consulta de registros de acreedores existentes para los empleados"
					Case EMPLOYEES_GRADE
						sNumber = "4"
						sMessage = "Consulta de registros de calificaciones existentes para los empleados"
					Case EMPLOYEES_THIRD_CONCEPT
						sNumber = "3"
						sMessage = "Consulta de registros de terceros existentes para el empleado"
					Case Else
						sNumber = "4"
						sMessage = "Consulta de registros existentes para el empleado"
				End Select
				Response.Write "<IMG SRC=""Images/Crcl" & sNumber & ".gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; ShowDisplay(document.all['ConceptInfoFormDiv']);""><FONT FACE=""Arial"" SIZE=""2"">" & sMessage & "</FONT></A><BR /><BR />"
				Response.Write "<DIV NAME=""ConceptInfoFormDiv"" ID=""ConceptInfoFormDiv"" STYLE=""display: none"">"
				Response.Write "<FORM NAME=""ConceptInfoFrm"" ID=""ConceptInfoFrm"" METHOD=""POST"" onSubmit=""return bReady"">"
					Select Case lReasonID
						Case CANCEL_EMPLOYEES_CONCEPTS, CANCEL_EMPLOYEES_SSI, CANCEL_EMPLOYEES_C04
							If (lReasonID = CANCEL_EMPLOYEES_CONCEPTS) Or (lReasonID = CANCEL_EMPLOYEES_C04) Then
								aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) = "2"
							End If
							lErrorNumber = DisplayPendingEmployeesConceptsTable(oRequest, oADODBConnection, False, sAction, lReasonID, aEmployeeComponent, sErrorDescription)
						Case EMPLOYEES_THIRD_CONCEPT
							lErrorNumber = DisplayPendingEmployeesCreditsTable(oRequest, oADODBConnection, 1, False, sAction, lReasonID, aEmployeeComponent, sErrorDescription)
						Case CANCEL_EMPLOYEES_CONCEPTS 
							lErrorNumber = DisplayPendingEmployeesTable(oRequest, oADODBConnection, False, sAction, lReasonID, 1, aEmployeeComponent, sErrorDescription)
						Case EMPLOYEES_ADD_BENEFICIARIES
							aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 1
							lErrorNumber = DisplayPendingEmployeesbeneficiariesTable(oRequest, oADODBConnection, False, sAction, lReasonID, aEmployeeComponent, sErrorDescription)
						Case EMPLOYEES_CREDITORS
							aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 1
							lErrorNumber = DisplayEmployeesCreditorsTable(oRequest, oADODBConnection, False, sAction, lReasonID, aEmployeeComponent, sErrorDescription)
						Case EMPLOYEES_FOR_RISK, EMPLOYEES_CONCEPT_08 , EMPLOYEES_ADDITIONALSHIFT
							lErrorNumber = DisplayConcepts040708HistoryList(oRequest, oADODBConnection, False, False, aEmployeeComponent, sErrorDescription)
						Case EMPLOYEES_BANK_ACCOUNTS
							aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 1
							If (Len(oRequest("StartPage").Item) > 0) Then
								If (CInt(oRequest("RowsType").Item) = 0) Then
									iStarPageForBanks = 0
								Else
 									iStarPageForBanks = CInt(oRequest("StartPage").Item)
								End If
							End If
							lErrorNumber = DisplayEmployeesBanksAccountsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, False, iStarPageForBanks, aEmployeeComponent, sErrorDescription)
						Case EMPLOYEES_GRADE
							aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 1
							lErrorNumber = DisplayEmployeesGradesTable(oRequest, oADODBConnection, False, aEmployeeComponent, sErrorDescription)
						Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS
							If lReasonID = EMPLOYEES_EXTRAHOURS Then
								aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 201
							Else
								aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 202
							End If
							aAbsenceComponent(N_ACTIVE_ABSENCE)= 1
							lErrorNumber = DisplayAbsencesTable(oRequest, oADODBConnection, DISPLAY_NOTHING, False, aAbsenceComponent, sErrorDescription)
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
	Select Case sAction
		Case "Absences"
			If (Len(oRequest("StartPage").Item) > 0) Then
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; ShowDisplay(document.all['ConceptInfoFormDiv']);" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
			End If
		Case "ConceptsValues"
			If CInt(Request.Cookies("SIAP_SubSectionID")) = 32 Then
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "if(document.all['AbsencesFormDiv'] != null) ShowDisplay(document.all['ConceptInfoFormDiv']);" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
			Else
				If (Len(oRequest("StartPage").Item) > 0) Then
					Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
						Response.Write "if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; ShowDisplay(document.all['ConceptInfoFormDiv']);" & vbNewLine
					Response.Write "//--></SCRIPT>" & vbNewLine
				End If
			End If
		Case "EmployeesMovements"
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			If lReasonID = EMPLOYEES_BANK_ACCOUNTS Then
				If (Len(oRequest("StartPage").Item) > 0) Then
					If CInt(oRequest("RowsType").Item) = 0 Then
						Response.Write "if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { ShowDisplay(document.all['UploadValidateInfoFormDiv']) }; HideDisplay(document.all['ConceptInfoFormDiv']);" & vbNewLine
					Else
						Response.Write "if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; ShowDisplay(document.all['ConceptInfoFormDiv']);" & vbNewLine
					End If
				End If
				Response.Write "ShowAmountFields(document.EmployeeFrm.BankID.value, 'SucursalSpn');" & vbNewLine
			End If
			Response.Write "//--></SCRIPT>" & vbNewLine
		Case "ThirdUploadMovements"
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			If lReasonID = EMPLOYEES_THIRD_PROCESS Then
				If Len(oRequest("ConceptFileName").Item) > 0 Then
					Response.Write "if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { ShowDisplay(document.all['UploadValidateInfoFormDiv']) };" & vbNewLine
				Else
					Response.Write "if(document.all['AbsencesFormDiv'] != null) { ShowDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) };" & vbNewLine
				End If
			End If
			Response.Write "//--></SCRIPT>" & vbNewLine
		Case "ServiceSheet"
			If (Len(oRequest("ShowServiceSheet").Item) > 0) Then
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "if(document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; ShowDisplay(document.all['UploadValidateServiceSheetDiv']); if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) };" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
			ElseIf (Len(oRequest("GenerateReport").Item) > 0) Then
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "if (document.all['AbsencesFormDiv'] != null) { HideDisplay(document.all['AbsencesFormDiv']) }; if (document.all['UploadValidateServiceSheetDiv'] != null) { HideDisplay(document.all['UploadValidateServiceSheetDiv']) }; if (document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; ShowDisplay(document.all['ConceptInfoFormDiv']);" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
			End If
	End Select
	DisplayUploadForm = Err.Number
	Err.Clear
End Function

Function DiplaySarProcessColumns(sFileName, sAction, lReasonID, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
		On Error Resume Next
		Const S_FUNCTION_NAME = "DiplaySarProcessColumns"
		Dim iColumns
		Dim iIndex
		Dim lErrorNumber

		Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo. <BR /> * Información requerida.</BLOCKQUOTE>")
		Response.Write "<BR />"
		lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
		If lErrorNumber = 0 Then
			Response.Write "<FORM NAME=""UploadProfessionalRiskMatrix"" ID=""UploadProfessionalRiskMatrix"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Load"" ID=""LoadHdn"" VALUE=""" & oRequest("Load").Item & """ />"
				If Len(oRequest("UploadFile").Item) > 0 Then Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UploadFile"" ID=""UploadFileHdn"" VALUE=""" & oRequest("UploadFile").Item & """ />"
				For iIndex = 1 To iColumns
					Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
					Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
						If (strComp(oRequest("Load").Item, "PayrollSummary",VbBinaryCompare) = 0) Then
							Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
							Response.Write "<OPTION VALUE=""SocietyID"""
								If iIndex = 1 Then Response.Write "SELECTED"
							Response.Write ">Sociedad</OPTION>"
							Response.Write "<OPTION VALUE=""CompanyID"""
								If iIndex = 2 Then Response.Write "SELECTED"
							Response.Write ">Empresa</OPTION>"
							Response.Write "<OPTION VALUE=""PeriodID"""
								If iIndex = 3 Then Response.Write "SELECTED"
							Response.Write ">Bimestre</OPTION>"
							Response.Write "<OPTION VALUE=""CLC"""
								If iIndex = 4 Then Response.Write "SELECTED"
							Response.Write ">CLC</OPTION>"
							Response.Write "<OPTION VALUE=""BankID"""
								If iIndex = 5 Then Response.Write "SELECTED"
							Response.Write ">Banco</OPTION>"
							Response.Write "<OPTION VALUE=""PaymentDateYYYYMMDD"""
								If iIndex = 6 Then Response.Write "SELECTED"
							Response.Write ">Fecha de pago (AAAAMMDD)</OPTION>"
							Response.Write "<OPTION VALUE=""PaymentDateDDMMYYYY"">Fecha de pago (DD-MM-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""PaymentDateMMDDYYYY"">Fecha de pago (MM-DD-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""EmployeeType"""
								If iIndex = 7 Then Response.Write "SELECTED"
							Response.Write ">Tabulador</OPTION>"
							Response.Write "<OPTION VALUE=""Income"""
								If iIndex = 8 Then Response.Write "SELECTED"
							Response.Write ">Percepciones</OPTION>"
							Response.Write "<OPTION VALUE=""Deductions"""
								If iIndex = 9 Then Response.Write "SELECTED"
							Response.Write ">Deducciones</OPTION>"
							Response.Write "<OPTION VALUE=""NetIncome"""
								If iIndex = 10 Then Response.Write "SELECTED"
							Response.Write ">Líquido</OPTION>"
							Response.Write "<OPTION VALUE=""Cpt_01"""
								If iIndex = 11 Then Response.Write "SELECTED"
							Response.Write ">Cpto_01</OPTION>"
							Response.Write "<OPTION VALUE=""Cpt_04"""
								If iIndex = 12 Then Response.Write "SELECTED"
							Response.Write ">Cpto_04</OPTION>"
							Response.Write "<OPTION VALUE=""Cpt_05"""
								If iIndex = 13 Then Response.Write "SELECTED"
							Response.Write ">Cpto_05</OPTION>"
							Response.Write "<OPTION VALUE=""Cpt_06"""
								If iIndex = 14 Then Response.Write "SELECTED"
							Response.Write ">Cpto_06</OPTION>"
							Response.Write "<OPTION VALUE=""Cpt_07"""
								If iIndex = 15 Then Response.Write "SELECTED"
							Response.Write ">Cpto_07</OPTION>"
							Response.Write "<OPTION VALUE=""Cpt_08"""
								If iIndex = 16 Then Response.Write "SELECTED"
							Response.Write ">Cpto_08</OPTION>"
							Response.Write "<OPTION VALUE=""Cpt_11"""
								If iIndex = 17 Then Response.Write "SELECTED"
							Response.Write ">Cpto_11</OPTION>"
							Response.Write "<OPTION VALUE=""Cpt_44"""
								If iIndex = 18 Then Response.Write "SELECTED"
							Response.Write ">Cpto_44</OPTION>"
							Response.Write "<OPTION VALUE=""Cpt_B2"""
								If iIndex = 19 Then Response.Write "SELECTED"
							Response.Write ">Cpto_B2</OPTION>"
							Response.Write "<OPTION VALUE=""Cpt_7S"""
								If iIndex = 20 Then Response.Write "SELECTED"
							Response.Write ">Cpto_7S</OPTION>"
						ElseIf (StrComp(oRequest("Load").Item, "BanamexCensus", vbBinaryCompare) = 0) Then
							Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
							Response.Write "<OPTION VALUE=""u_version"""
								If iIndex = 1 Then Response.Write "SELECTED"
							Response.Write ">Versión</OPTION>"
							Response.Write "<OPTION VALUE=""employeeID"""
								If iIndex = 2 Then Response.Write "SELECTED"
							Response.Write ">Núm. Empleado</OPTION>"
							Response.Write "<OPTION VALUE=""rfc"""
								If iIndex = 3 Then Response.Write "SELECTED"
							Response.Write ">Filiación</OPTION>"
							Response.Write "<OPTION VALUE=""curp"""
								If iIndex = 4 Then Response.Write "SELECTED"
							Response.Write ">CURP</OPTION>"
							Response.Write "<OPTION VALUE=""SocialSecurityNumber"""
								If iIndex = 5 Then Response.Write "SELECTED"
							Response.Write ">NSS</OPTION>"
							Response.Write "<OPTION VALUE=""EmployeeLastName"""
								If iIndex = 6 Then Response.Write "SELECTED"
							Response.Write ">Apellido paterno</OPTION>"
							Response.Write "<OPTION VALUE=""EmployeeLastName2"""
								If iIndex = 7 Then Response.Write "SELECTED"
							Response.Write ">Apellido materno</OPTION>"
							Response.Write "<OPTION VALUE=""EmployeeName"""
								If iIndex = 8 Then Response.Write "SELECTED"
							Response.Write ">Nombre</OPTION>"
							Response.Write "<OPTION VALUE=""CT"""
								If iIndex = 9 Then Response.Write "SELECTED"
							Response.Write ">CT</OPTION>"
							Response.Write "<OPTION VALUE=""birthDateYYYYMMDD"""
								If iIndex = 10 Then Response.Write "SELECTED"
							Response.Write ">Fecha de nacimiento (AAAAMMDD)</OPTION>"
							Response.Write "<OPTION VALUE=""birthDateDDMMYYYY"">Fecha de nacimiento (DD-MM-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""birthDateMMDDYYYY"">Fecha de nacimiento (MM-DD-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""birthState"""
								If iIndex = 11 Then Response.Write "SELECTED"
							Response.Write ">Estado de Nacimiento</OPTION>"
							Response.Write "<OPTION VALUE=""GenderShortName"""
								If iIndex = 12 Then Response.Write "SELECTED"
							Response.Write ">Sexo</OPTION>"
							Response.Write "<OPTION VALUE=""JoinDateYYYYMMDD"""
								If iIndex = 13 Then Response.Write "SELECTED"
							Response.Write ">Fecha de ingreso (AAAAMMDD)</OPTION>"
							Response.Write "<OPTION VALUE=""JoinDateDDMMYYYY"">Fecha de ingreso (DD-MM-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""JoinDateMMDDYYYY"">Fecha de ingreso (MM-DD-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""CotDateYYYYMMDD"""
								If iIndex = 14 Then Response.Write "SELECTED"
							Response.Write ">Fecha de cotización (AAAAMMDD)</OPTION>"
							Response.Write "<OPTION VALUE=""CotDateDDMMYYYY"">Fecha de cotización (DD-MM-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""CotDateMMDDYYYY"">Fecha de cotización (MM-DD-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""Salary"""
								If iIndex = 15 Then Response.Write "SELECTED"
							Response.Write ">Salario base</OPTION>"
							Response.Write "<OPTION VALUE=""Fovi"""
								If iIndex = 16 Then Response.Write "SELECTED"
							Response.Write ">FOVISSSTE</OPTION>"
							Response.Write "<OPTION VALUE=""Period"""
								If iIndex = 17 Then Response.Write "SELECTED"
							Response.Write ">Bimestre</OPTION>"
							Response.Write "<OPTION VALUE=""Status"""
								If iIndex = 18 Then Response.Write "SELECTED"
							Response.Write ">Estatus</OPTION>"
							Response.Write "<OPTION VALUE=""ChangeFlag"""
								If iIndex = 19 Then Response.Write "SELECTED"
							Response.Write ">Abre_Cierra</OPTION>"
							Response.Write "<OPTION VALUE=""MaritalStatusID"""
								If iIndex = 20 Then Response.Write "SELECTED"
							Response.Write ">Estado Civil</OPTION>"
							Response.Write "<OPTION VALUE=""Address"""
							If iIndex = 21 Then Response.Write "SELECTED"
							Response.Write ">Domicilio</OPTION>"
							Response.Write "<OPTION VALUE=""Colony"""
								If iIndex = 22 Then Response.Write "SELECTED"
							Response.Write ">Colonia</OPTION>"
							Response.Write "<OPTION VALUE=""city"""
								If iIndex = 23 Then Response.Write "SELECTED"
							Response.Write ">Municipio</OPTION>"
							Response.Write "<OPTION VALUE=""ZipZone"""
								If iIndex = 24 Then Response.Write "SELECTED"
							Response.Write ">C.P.</OPTION>"
							Response.Write "<OPTION VALUE=""State"""
								If iIndex = 25 Then Response.Write "SELECTED"
							Response.Write ">Entidad Federativa</OPTION>"
							Response.Write "<OPTION VALUE=""Nombram"""
								If iIndex = 26 Then Response.Write "SELECTED"
							Response.Write ">Nombramiento</OPTION>"
							Response.Write "<OPTION VALUE=""Afore"""
								If iIndex = 27 Then Response.Write "SELECTED"
							Response.Write ">Afore</OPTION>"
							Response.Write "<OPTION VALUE=""ICEFA"""
								If iIndex = 28 Then Response.Write "SELECTED"
							Response.Write ">Clave ICEFA</OPTION>"
							Response.Write "<OPTION VALUE=""ICNumber"""
								If iIndex = 29 Then Response.Write "SELECTED"
							Response.Write ">Número de control interno</OPTION>"
							Response.Write "<OPTION VALUE=""mot_baja"""
								If iIndex = 30 Then Response.Write "SELECTED"
							Response.Write ">Motivo de baja</OPTION>"
							Response.Write "<OPTION VALUE=""Salary_v"""
								If iIndex = 31 Then Response.Write "SELECTED"
							Response.Write ">Salario Base V</OPTION>"
							Response.Write "<OPTION VALUE=""FullPay"""
								If iIndex = 32 Then Response.Write "SELECTED"
							Response.Write ">Salario Integrado</OPTION>"
							Response.Write "<OPTION VALUE=""WorkingDays"""
								If iIndex = 33 Then Response.Write "SELECTED"
							Response.Write ">Días laborados</OPTION>"
							Response.Write "<OPTION VALUE=""InabilityDays"""
								If iIndex = 34 Then Response.Write "SELECTED"
							Response.Write ">Días incapacidad</OPTION>"
							Response.Write "<OPTION VALUE=""AbsenceDays"""
								If iIndex = 35 Then Response.Write "SELECTED"
							Response.Write ">Días ausentados</OPTION>"
							Response.Write "<OPTION VALUE=""EmployeeContributions"""
								If iIndex = 36 Then Response.Write "SELECTED"
							Response.Write ">Aportación empleado</OPTION>"
							Response.Write "<OPTION VALUE=""EmployeecontributionsAmount"""
								If iIndex = 37 Then Response.Write "SELECTED"
							Response.Write ">Importe aportación</OPTION>"
							Response.Write "<OPTION VALUE=""StartDateYYYYMMDD"""
								If iIndex = 38 Then Response.Write "SELECTED"
							Response.Write ">Fecha de inicio (AAAAMMDD)</OPTION>"
							Response.Write "<OPTION VALUE=""StartDateDDMMYYYY"">Fecha de inicio (DD-MM-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""StartDateMMDDYYYY"">Fecha de inicio (MM-DD-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""EndDateYYYYMMDD"""
								If iIndex = 39 Then Response.Write "SELECTED"
							Response.Write ">Fecha fin (AAAAMMDD)</OPTION>"
							Response.Write "<OPTION VALUE=""EndDateDDMMYYYY"">Fecha fin (DD-MM-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""EndDateMMDDYYYY"">Fecha fin (MM-DD-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""Comments"""
								If iIndex = 40 Then Response.Write "SELECTED"
							Response.Write ">Comentarios</OPTION>"
						ElseIf (StrComp(oRequest("Load").Item, "SarCensus", vbBinaryCompare) = 0) Then
							Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
							Response.Write "<OPTION VALUE=""u_version"""
								If iIndex = 1 Then Response.Write "SELECTED"
							Response.Write ">Empresa</OPTION>"
							Response.Write "<OPTION VALUE=""rfc"""
								If iIndex = 2 Then Response.Write "SELECTED"
							Response.Write ">Filiación</OPTION>"
							Response.Write "<OPTION VALUE=""curp"""
								If iIndex = 3 Then Response.Write "SELECTED"
							Response.Write ">CURP</OPTION>"
							Response.Write "<OPTION VALUE=""SocialSecurityNumber"""
								If iIndex = 4 Then Response.Write "SELECTED"
							Response.Write ">NSS</OPTION>"
							Response.Write "<OPTION VALUE=""EmployeeLastName"""
								If iIndex = 5 Then Response.Write "SELECTED"
							Response.Write ">Apellido paterno</OPTION>"
							Response.Write "<OPTION VALUE=""EmployeeLastName2"""
								If iIndex = 6 Then Response.Write "SELECTED"
							Response.Write ">Apellido materno</OPTION>"
							Response.Write "<OPTION VALUE=""EmployeeName"""
								If iIndex = 7 Then Response.Write "SELECTED"
							Response.Write ">Nombre</OPTION>"
							Response.Write "<OPTION VALUE=""pageID"""
								If iIndex = 8 Then Response.Write "SELECTED"
							Response.Write ">ID_Pag</OPTION>"
							Response.Write "<OPTION VALUE=""CT"""
								If iIndex = 9 Then Response.Write "SELECTED"
							Response.Write ">CT</OPTION>"
							Response.Write "<OPTION VALUE=""birthDateYYYYMMDD"""
								If iIndex = 10 Then Response.Write "SELECTED"
							Response.Write ">Fecha de nacimiento (AAAAMMDD)</OPTION>"
							Response.Write "<OPTION VALUE=""birthDateDDMMYYYY"">Fecha de nacimiento (DD-MM-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""birthDateMMDDYYYY"">Fecha de nacimiento (MM-DD-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""birthState"""
								If iIndex = 11 Then Response.Write "SELECTED"
							Response.Write ">Estado de Nacimiento</OPTION>"
							Response.Write "<OPTION VALUE=""GenderShortName"""
								If iIndex = 12 Then Response.Write "SELECTED"
							Response.Write ">Sexo</OPTION>"
							Response.Write "<OPTION VALUE=""MaritalStatusID"""
								If iIndex = 13 Then Response.Write "SELECTED"
							Response.Write ">Estado Civil</OPTION>"
							Response.Write "<OPTION VALUE=""Address"""
								If iIndex = 14 Then Response.Write "SELECTED"
							Response.Write ">Domicilio</OPTION>"
							Response.Write "<OPTION VALUE=""Colony"""
								If iIndex = 15 Then Response.Write "SELECTED"
							Response.Write ">Colonia</OPTION>"
							Response.Write "<OPTION VALUE=""city"""
								If iIndex = 16 Then Response.Write "SELECTED"
							Response.Write ">Municipio</OPTION>"
							Response.Write "<OPTION VALUE=""ZipZone"""
								If iIndex = 17 Then Response.Write "SELECTED"
							Response.Write ">C.P.</OPTION>"
							Response.Write "<OPTION VALUE=""State"""
								If iIndex = 18 Then Response.Write "SELECTED"
							Response.Write ">Entidad Federativa</OPTION>"
							Response.Write "<OPTION VALUE=""Nombram"""
								If iIndex = 19 Then Response.Write "SELECTED"
							Response.Write ">Nombramiento</OPTION>"
							Response.Write "<OPTION VALUE=""EmployeeID"""
								If iIndex = 20 Then Response.Write "SELECTED"
							Response.Write ">Número de empleado</OPTION>"
							Response.Write "<OPTION VALUE=""ICEFA"""
								If iIndex = 21 Then Response.Write "SELECTED"
							Response.Write ">Clave ICEFA</OPTION>"
							Response.Write "<OPTION VALUE=""Afore"""
								If iIndex = 22 Then Response.Write "SELECTED"
							Response.Write ">Afore</OPTION>"
							Response.Write "<OPTION VALUE=""JoinDateYYYYMMDD"""
								If iIndex = 23 Then Response.Write "SELECTED"
							Response.Write ">Fecha de ingreso (AAAAMMDD)</OPTION>"
							Response.Write "<OPTION VALUE=""JoinDateDDMMYYYY"">Fecha de ingreso (DD-MM-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""JoinDateMMDDYYYY"">Fecha de ingreso (MM-DD-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""StartDateYYYYMMDD"""
								If iIndex = 24 Then Response.Write "SELECTED"
							Response.Write ">Fecha de cotización (AAAAMMDD)</OPTION>"
							Response.Write "<OPTION VALUE=""StartDateDDMMYYYY"">Fecha de cotización (DD-MM-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""StartDateMMDDYYYY"">Fecha de cotización (MM-DD-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""Fovi"""
								If iIndex = 25 Then Response.Write "SELECTED"
							Response.Write ">FOVISSSTE</OPTION>"
							Response.Write "<OPTION VALUE=""WorkingDays"""
								If iIndex = 26 Then Response.Write "SELECTED"
							Response.Write ">Días laborados</OPTION>"
							Response.Write "<OPTION VALUE=""InabilityDays"""
								If iIndex = 27 Then Response.Write "SELECTED"
							Response.Write ">Días incapacidad</OPTION>"
							Response.Write "<OPTION VALUE=""AbsenceDays"""
								If iIndex = 28 Then Response.Write "SELECTED"
							Response.Write ">Días ausentados</OPTION>"
							Response.Write "<OPTION VALUE=""FullPay"""
								If iIndex = 29 Then Response.Write "SELECTED"
							Response.Write ">Salario Integrado</OPTION>"
							Response.Write "<OPTION VALUE=""Salary"""
								If iIndex = 30 Then Response.Write "SELECTED"
							Response.Write ">Salario Base</OPTION>"
							Response.Write "<OPTION VALUE=""Salary_v"""
								If iIndex = 31 Then Response.Write "SELECTED"
							Response.Write ">Salario Base V</OPTION>"
							Response.Write "<OPTION VALUE=""EmployeeContributions"""
								If iIndex = 32 Then Response.Write "SELECTED"
							Response.Write ">Aportación empleado</OPTION>"
							Response.Write "<OPTION VALUE=""EmployeecontributionsAmount"""
								If iIndex = 33 Then Response.Write "SELECTED"
							Response.Write ">Importe aportación</OPTION>"
						ElseIf (strComp(oRequest("Load").Item, "ConsarFile",VbBinaryCompare) = 0) Then
							Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
							Response.Write "<OPTION VALUE=""cve"""
								If iIndex = 1 Then Response.Write "SELECTED"
							Response.Write ">CVE</OPTION>"
							Response.Write "<OPTION VALUE=""rfc"""
								If iIndex = 2 Then Response.Write "SELECTED"
							Response.Write ">Filiación</OPTION>"
							Response.Write "<OPTION VALUE=""curp"""
								If iIndex = 3 Then Response.Write "SELECTED"
							Response.Write ">CURP</OPTION>"
							Response.Write "<OPTION VALUE=""SocialSecurityNumber"""
								If iIndex = 4 Then Response.Write "SELECTED"
							Response.Write ">NSS</OPTION>"
							Response.Write "<OPTION VALUE=""EmployeeLastName"""
								If iIndex = 5 Then Response.Write "SELECTED"
							Response.Write ">Apellido paterno</OPTION>"
							Response.Write "<OPTION VALUE=""EmployeeLastName2"""
								If iIndex = 6 Then Response.Write "SELECTED"
							Response.Write ">Apellido materno</OPTION>"
							Response.Write "<OPTION VALUE=""EmployeeName"""
								If iIndex = 7 Then Response.Write "SELECTED"
							Response.Write ">Nombre</OPTION>"
							Response.Write "<OPTION VALUE=""nombram"""
								If iIndex = 8 Then Response.Write "SELECTED"
							Response.Write ">Nombramiento</OPTION>"
							Response.Write "<OPTION VALUE=""icefa"""
								If iIndex = 9 Then Response.Write "SELECTED"
							Response.Write ">CVE ICEFA</OPTION>"
							Response.Write "<OPTION VALUE=""JoinDateYYYYMMDD"""
								If iIndex = 10 Then Response.Write "SELECTED"
							Response.Write ">Fecha de ingreso (AAAAMMDD)</OPTION>"
							Response.Write "<OPTION VALUE=""JoinDateDDMMYYYY"">Fecha de ingreso (DD-MM-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""JoinDateMMDDYYYY"">Fecha de ingreso (MM-DD-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""CotDateYYYYMMDD"""
								If iIndex = 11 Then Response.Write "SELECTED"
							Response.Write ">Fecha de cotización (AAAAMMDD)</OPTION>"
							Response.Write "<OPTION VALUE=""CotDateDDMMYYYY"">Fecha de cotización (DD-MM-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""CotDateMMDDYYYY"">Fecha de cotización (MM-DD-AAAA)</OPTION>"
							Response.Write "<OPTION VALUE=""fovi"""
								If iIndex = 12 Then Response.Write "SELECTED"
							Response.Write ">Crédito FOVISSSTE</OPTION>"
							Response.Write "<OPTION VALUE=""workingDays"""
								If iIndex = 13 Then Response.Write "SELECTED"
							Response.Write ">Días cotizados</OPTION>"
							Response.Write "<OPTION VALUE=""inabilityDays"""
								If iIndex = 14 Then Response.Write "SELECTED"
							Response.Write ">Días incapacidad</OPTION>"
							Response.Write "<OPTION VALUE=""absenceDays"""
								If iIndex = 15 Then Response.Write "SELECTED"
							Response.Write ">Días de ausencia</OPTION>"
							Response.Write "<OPTION VALUE=""salary"""
								If iIndex = 16 Then Response.Write "SELECTED"
							Response.Write ">Salario base</OPTION>"
							Response.Write "<OPTION VALUE=""salaryV"""
								If iIndex = 17 Then Response.Write "SELECTED"
							Response.Write ">Salario base V</OPTION>"
							Response.Write "<OPTION VALUE=""sar"""
								If iIndex = 18 Then Response.Write "SELECTED"
							Response.Write ">SAR</OPTION>"
							Response.Write "<OPTION VALUE=""entityCV"""
								If iIndex = 19 Then Response.Write "SELECTED"
							Response.Write ">CV Patrón</OPTION>"
							Response.Write "<OPTION VALUE=""employeeCV"""
								If iIndex = 20 Then Response.Write "SELECTED"
							Response.Write ">CV Empleado</OPTION>"
							Response.Write "<OPTION VALUE=""foviAmount"""
								If iIndex = 21 Then Response.Write "SELECTED"
							Response.Write ">Fovissste</OPTION>"
							Response.Write "<OPTION VALUE=""employeeSaving"""
								If iIndex = 22 Then Response.Write "SELECTED"
							Response.Write ">Ahorro empleado</OPTION>"
							Response.Write "<OPTION VALUE=""entitySaving"""
								If iIndex = 23 Then Response.Write "SELECTED"
							Response.Write ">Ahorro dependencia</OPTION>"
						End If
					Response.Write "</SELECT>"
					Response.Write "<BR />"
				Next
				Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ProcessFile"" ID=""ProcessFileBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
			Response.Write "</FORM>"
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "function CheckColumnsToUpload(oForm) {" & vbNewLine
					Response.Write "var bDuplicated = false;" & vbNewLine
					Response.Write "var sFields = '';" & vbNewLine
					For iIndex = 1 To iColumns
						Response.Write "if (oForm.Column" & iIndex & ".value != 'NA') {" & vbNewLine
							Response.Write "if (sFields.search(eval('/,' + oForm.Column" & iIndex & ".value + ',/gi')) == -1)" & vbNewLine
								Response.Write "sFields += oForm.Column" & iIndex & ".value + ',';" & vbNewLine
							Response.Write "else" & vbNewLine
								Response.Write "bDuplicated = true;" & vbNewLine
						Response.Write "}" & vbNewLine
					Next
					Response.Write "if (bDuplicated) {" & vbNewLine
						Response.Write "alert('Existen columnas duplicadas.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					If (strComp(oRequest("Load").Item, "PayrollSummary",VbBinaryCompare) = 0) Then
						Response.Write "if (sFields.search(/SocietyID/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con el identificador de la rama');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/CompanyID/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con el identificador del servicio');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/PeriodID/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con el identificador del puesto');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/CLC/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con el identificador del tipo de centro de trabajo');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/BankID/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con monto del riesgo');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/PaymentDate/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con monto del riesgo');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/EmployeeType/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con monto del riesgo');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/Income/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con monto del riesgo');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/Deductions/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con monto del riesgo');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/NetIncome/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con monto del riesgo');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/Cpt_01/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con monto del riesgo');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/Cpt_04/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con monto del riesgo');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/Cpt_05/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con monto del riesgo');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/Cpt_06/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con monto del riesgo');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/Cpt_06/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con monto del riesgo');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/Cpt_07/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con monto del riesgo');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/Cpt_08/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con monto del riesgo');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/Cpt_11/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con monto del riesgo');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/Cpt_44/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con monto del riesgo');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/Cpt_B2/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con monto del riesgo');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/Cpt_7S/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con monto del riesgo');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
					ElseIf (strComp(oRequest("Load").Item, "BanamexCensus",VbBinaryCompare) = 0) Then
						Response.Write "if (sFields.search(/u_version/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con la versión');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/EmployeeID/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con el identificador del empleado');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/rfc/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con el identificador del RFC');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/curp/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con el identificador de la CURP');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/SocialSecurityNumber/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con monto del número de seguro social');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/EmployeeLastName/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con el apellido paterno');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/EmployeeLastName2/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con el apellido materno');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/EmployeeName/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con el nombre');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/ct/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con el valor CT');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/birthDate/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con la fecha de nacimiento');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/birthState/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con el estado de nacimiento');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/GenderShortName/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con el genero del empleado');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/JoinDate/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con la fecha de ingreso');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/CotDate/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo fecha_cot');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/Salary/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo con monto del salario');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/Fovi/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo fovi');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/Period/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo periodo');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/Status/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo estatus');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/ChangeFlag/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo abre-cierra');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/MaritalStatusID/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo estado civil');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/Address/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo dirección');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/Colony/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo colonia');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/city/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo ciudad');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/ZipZone/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo código postal');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/State/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo estado');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/Nombram/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo nombramiento');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/Afore/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo Afore');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/ICEFA/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo ICEFA');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/ICNumber/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo número de control interno');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/mot_baja/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo mot_baja');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/Salary_v/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo salario base v');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/FullPay/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo salario integral');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/WorkingDays/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo días laborados');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/InabilityDays/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo días de incapacidad');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/AbsenceDays/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo días de ausencia');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/EmployeeContributions/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo contribuciones del empleado');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/EmployeecontributionsAmount/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo monto de las contribuciones');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/StartDate/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo fecha de inicio');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
						Response.Write "if (sFields.search(/EndDate/gi) == -1) {" & vbNewLine
							Response.Write "alert('No se ha establecido el campo fecha fin');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"
					End If
				Response.Write "}" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
		End If	

End Function

Function DisplayProfessionalRiskColumns(sFileName, sAction, lReasonID, sErrorDescription)
'************************************************************
'Purpose: To show the uploaded file columns
'Inputs:  iColumns
'Outputs: sErrorDescription
'************************************************************
		On Error Resume Next
		Const S_FUNCTION_NAME = "DisplayProfessionalRiskColumns"
		Dim iColumns
		Dim iIndex
		Dim lErrorNumber
		
		Call DisplayInstructionsMessage("Paso 2. Identifique las columnas del archivo", "<BLOCKQUOTE>Indique a qué campo pertenece cada columna del archivo. <BR /> * Información requerida.</BLOCKQUOTE>")
		Response.Write "<BR />"
		lErrorNumber = ShowUploadedFile(sFileName, iColumns, sErrorDescription)
		If lErrorNumber = 0 Then
			Response.Write "<FORM NAME=""UploadProfessionalRiskMatrix"" ID=""UploadProfessionalRiskMatrix"" METHOD=""POST"" onSubmit=""return CheckColumnsToUpload(this)"">"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""3"" />"
				For iIndex = 1 To iColumns
					Response.Write "&nbsp;&nbsp;Columna " & iIndex & ":&nbsp;"
					Response.Write "<SELECT NAME=""Column" & iIndex & """ ID=""Column" & iIndex & "Cmb"" CLASS=""Lists"" SIZE=""1"">"
					Response.Write "<OPTION VALUE=""NA"">Esta columna no se toma en cuenta</OPTION>"
					Response.Write "<OPTION VALUE=""BranchID"">Rama</OPTION>"
					Response.Write "<OPTION VALUE=""CenterTypeID"">Tipo de centro de trabajo</OPTION>"
					Response.Write "<OPTION VALUE=""PositionID"">Clave del puesto</OPTION>"
					Response.Write "<OPTION VALUE=""ServiceID"">Clave del servicio</OPTION>"
					Response.Write "<OPTION VALUE=""RiskLevel"">Monto de riesgo</OPTION>"
					Response.Write "<OPTION VALUE=""StartDate"">Fecha de inicio de vigencia* (AAAAMMDD)</OPTION>"
					Response.Write "<OPTION VALUE=""EndDate"">Fecha de fin de vigencia (AAAAMMDD)</OPTION>"
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
					Response.Write "if (sFields.search(/BranchID/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con el identificador de la rama');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
					Response.Write "if (sFields.search(/ServiceID/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con el identificador del servicio');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
					Response.Write "if (sFields.search(/PositionID/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con el identificador del puesto');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
					Response.Write "if (sFields.search(/CenterTypeID/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con el identificador del tipo de centro de trabajo');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
					Response.Write "if (sFields.search(/RiskLevel/gi) == -1) {" & vbNewLine
						Response.Write "alert('No se ha establecido el campo con monto del riesgo');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"
				Response.Write "}" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
		End If
		DisplayProfessionalRiskColumns = lErrorNumber
		Err.Clear
End Function
%>