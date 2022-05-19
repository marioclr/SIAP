<%
Function Display261SearchForm(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the search form for the EmployeesKardex table.
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "Display261SearchForm"

	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		Response.Write "function CheckEmployeeField(oForm) {" & vbNewLine
			Response.Write "if (oForm) {" & vbNewLine
				Response.Write "if ((oForm.EmployeeID.value.length == 0) || (oForm.EmployeeID.value == '-1')) {" & vbNewLine
					Response.Write "alert('Favor de especificar el número de empleado.');" & vbNewLine
					Response.Write "oForm.EmployeeID.focus();" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "else{" & vbNewLine
					Response.Write "return true;" & vbNewLine
			Response.Write "}}" & vbNewLine
		Response.Write "} // End of CheckEmployeeField" & vbNewLine
	Response.Write "//--></SCRIPT>" & vbNewLine
	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Indique los criterios que se utilizarán en la búsqueda de registros:</B></FONT><BR />"
	Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""Main_ISSSTE.asp"" METHOD=""GET"" onSubmit=""return CheckEmployeeField(this)"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""261"" />"
		Response.Write "<TABLE BORDER=""0"" CELLPADING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del empleado:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeID"" ID=""EmployeeIDTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE=""" & oRequest("EmployeeID").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<DONT_EXPORT><TD COLSPAN=""2"" ALIGN=""RIGHT""><BR /><INPUT TYPE=""BUTTON"" NAME=""Back"" ID=""BackBtn"" VALUE=""Regresar"" onClick=""window.history.go(-1);"" CLASS="" Buttons "" /></DONT_EXPORT>"
				Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""1"" />"
				Response.Write "<SPAN NAME=""ContinueSpn"" ID=""ContinueSpn""><INPUT TYPE=""SUBMIT"" NAME=""DoSearch"" ID=""DoSearchBtn"" VALUE=""Buscar Registros"" CLASS=""Buttons"" /></SPAN></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	Response.Write "</FORM>"

	Display261SearchForm = Err.number
	Err.Clear
End Function

Function Display262SearchForm(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the search form for the EmployeesKardex table.
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "Display262SearchForm"

	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		Response.Write "function CheckEmployeeField(oForm) {" & vbNewLine
			Response.Write "if (oForm) {" & vbNewLine
				Response.Write "if ((oForm.EmployeeID.value.length == 0) || (oForm.EmployeeID.value == '-1')) {" & vbNewLine
					Response.Write "alert('Favor de especificar el número de empleado.');" & vbNewLine
					Response.Write "oForm.EmployeeID.focus();" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "else{" & vbNewLine
					Response.Write "return true;" & vbNewLine
			Response.Write "}}" & vbNewLine
		Response.Write "} // End of CheckEmployeeField" & vbNewLine
	Response.Write "//--></SCRIPT>" & vbNewLine
	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Indique los criterios que se utilizarán en la búsqueda de registros:</B></FONT><BR />"
	Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""Main_ISSSTE.asp"" METHOD=""GET"" onSubmit=""return CheckEmployeeField(this)"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""262"" />"
		Response.Write "<TABLE BORDER=""0"" CELLPADING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del empleado:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeID"" ID=""EmployeeIDTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE=""" & oRequest("EmployeeID").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<DONT_EXPORT><TD COLSPAN=""2"" ALIGN=""RIGHT""><BR /><INPUT TYPE=""BUTTON"" NAME=""Back"" ID=""BackBtn"" VALUE="" Regresar "" onClick=""window.history.go(-1);"" CLASS=""Buttons"" /></DONT_EXPORT>"
				Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""1"" />"
				Response.Write "<SPAN NAME=""ContinueSpn"" ID=""ContinueSpn""><INPUT TYPE=""SUBMIT"" NAME=""DoSearch"" ID=""DoSearchBtn"" VALUE=""Buscar Registros"" CLASS=""Buttons"" /></SPAN></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	Response.Write "</FORM>"

	Display262SearchForm = Err.number
	Err.Clear
End Function

Function Display267SearchForm(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the search form for the EmployeesKardex table.
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "Display267SearchForm"

	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Indique los criterios que se utilizarán en la búsqueda de registros:</B></FONT><BR />"
	Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""Main_ISSSTE.asp"" METHOD=""GET"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""267"" />"
		Response.Write "<TABLE BORDER=""0"" CELLPADING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del empleado:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeID"" ID=""EmployeeIDTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE=""" & oRequest("EmployeeID").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fecha de entrega:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("StartYear").Item), CInt(oRequest("StartMonth").Item), CInt(oRequest("StartDay").Item), "StartYear", "StartMonth", "StartDay", 2009, Year(Date()), True, True)
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("EndYear").Item), CInt(oRequest("EndMonth").Item), CInt(oRequest("EndDay").Item), "EndYear", "EndMonth", "EndDay", 2009, Year(Date()), True, True)
				Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""2"" ALIGN=""RIGHT""><BR /><INPUT TYPE=""SUBMIT"" NAME=""DoSearch"" ID=""DoSearchBtn"" VALUE=""Buscar Registros"" CLASS=""Buttons"" /></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	Response.Write "</FORM>"

	Display267SearchForm = Err.number
	Err.Clear
End Function 

Function Display281SearchForm(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the search form for the EmployeesKardex5
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "Display281SearchForm"

	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Indique los criterios que se utilizarán en la búsqueda de registros:</B></FONT><BR />"
	Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""Main_ISSSTE.asp"" METHOD=""GET"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""281"" />"
		Response.Write "<TABLE BORDER=""0"" CELLPADING=""0"" CELLSPACING=""0"">"
			If CInt(Request.Cookies("SIAP_SectionID")) = 2 Then
				Response.Write "<TR>"
					Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Tipo de registro:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""Kardex5TypeID"" ID=""Kardex5TypeIDLst"" SIZE=""3"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
						Response.Write "<OPTION VALUE="""">Todos</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Kardex5Types", "Kardex5TypeID", "Kardex5TypeName", "", "Kardex5TypeName", oRequest("Kardex5TypeID").Item, "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
			End If
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Referido por:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""Kardex5OriginID"" ID=""Kardex5OriginIDLst"" SIZE=""3"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Kardex5Origins", "Kardex5OriginID", "Kardex5OriginName", "", "Kardex5OriginName", oRequest("Kardex5OriginID").Item, "Ninguno;;;-1", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre de empleado:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeName"" ID=""EmployeeNameTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("EmployeeName").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido paterno:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName"" ID=""EmployeeLastNameTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("EmployeeLastName").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido materno:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName2"" ID=""EmployeeLastName2Txt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("EmployeeLastName2").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Puesto:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""PositionID"" ID=""PositionIDCmb"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Positions", "PositionID", "PositionShortName, PositionName", "(PositiontypeID=1) And (EndDate=30000000) And (Active=1)", "PositionShortName, PositionName", oRequest("PositionID").Item, "Ninguno;;;-1", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fecha de registro:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("StartStartYear").Item), CInt(oRequest("StartStartMonth").Item), CInt(oRequest("StartStartDay").Item), "StartStartYear", "StartStartMonth", "StartStartDay", 2009, Year(Date()), True, True)
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("EndStartYear").Item), CInt(oRequest("EndStartMonth").Item), CInt(oRequest("EndStartDay").Item), "EndStartYear", "EndStartMonth", "EndStartDay", 2009, Year(Date()), True, True)
				Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""2"" ALIGN=""RIGHT""><BR /><INPUT TYPE=""SUBMIT"" NAME=""DoSearch"" ID=""DoSearchBtn"" VALUE=""Buscar Registros"" CLASS=""Buttons"" /></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	Response.Write "</FORM>"

	Display281SearchForm = Err.number
	Err.Clear
End Function 

Function Display282SearchForm(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the search form for the EmployeesKardex4
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "Display282SearchForm"

	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Indique los criterios que se utilizarán en la búsqueda de registros:</B></FONT><BR />"
	Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""Main_ISSSTE.asp"" METHOD=""GET"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""282"" />"
		Response.Write "<TABLE BORDER=""0"" CELLPADING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Registro de:&nbsp;</FONT></TD>"
				If CInt(Request.Cookies("SIAP_SectionID")) = 2 Then
					Response.Write "<TD><SELECT NAME=""KardexChangeTypeID"" ID=""KardexChangeTypeIDLst"" SIZE=""6"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
						Response.Write "<OPTION VALUE="""">Todos</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "KardexChangeTypes", "KardexChangeTypeID", "KardexChangeTypeName", "", "KardexChangeTypeName", oRequest("KardexChangeTypeID").Item, "Ninguno;;;-1", sErrorDescription)
				Else
					Response.Write "<TD><SELECT NAME=""KardexChangeTypeID"" ID=""KardexChangeTypeIDLst"" SIZE=""3"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
						Response.Write "<OPTION VALUE=""0,1"">Todos</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "KardexChangeTypes", "KardexChangeTypeID", "KardexChangeTypeName", "(KardexChangeTypeID In (0,1))", "KardexChangeTypeName", oRequest("KardexChangeTypeID").Item, "Ninguno;;;-1", sErrorDescription)
				End If
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número de empleado:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeID"" ID=""EmployeeIDTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE=""" & oRequest("EmployeeID").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número de plaza:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""JobID"" ID=""JobIDTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE=""" & oRequest("JobID").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Puesto:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""PositionID"" ID=""PositionIDCmb"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Positions", "PositionID", "PositionShortName, PositionName", "(EndDate=30000000) And (Active=1)", "PositionShortName, PositionName", oRequest("PositionID").Item, "Ninguno;;;-1", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fecha de registro:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("StartStartYear").Item), CInt(oRequest("StartStartMonth").Item), CInt(oRequest("StartStartDay").Item), "StartStartYear", "StartStartMonth", "StartStartDay", 2009, Year(Date()), True, True)
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("EndStartYear").Item), CInt(oRequest("EndStartMonth").Item), CInt(oRequest("EndStartDay").Item), "EndStartYear", "EndStartMonth", "EndStartDay", 2009, Year(Date()), True, True)
				Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fecha de resolución:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("StartEndYear").Item), CInt(oRequest("StartEndMonth").Item), CInt(oRequest("StartEndDay").Item), "StartEndYear", "StartEndMonth", "StartEndDay", 2009, Year(Date()), True, True)
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("EndEndYear").Item), CInt(oRequest("EndEndMonth").Item), CInt(oRequest("EndEndDay").Item), "EndEndYear", "EndEndMonth", "EndEndDay", 2009, Year(Date()), True, True)
				Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""2"" ALIGN=""RIGHT""><BR /><INPUT TYPE=""SUBMIT"" NAME=""DoSearch"" ID=""DoSearchBtn"" VALUE=""Buscar Registros"" CLASS=""Buttons"" /></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	Response.Write "</FORM>"

	Display282SearchForm = Err.number
	Err.Clear
End Function 

Function Display38SearchForm(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the search form for the Areas and the
'         PaymentCenters tables
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "Display38SearchForm"

	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Indique los criterios que se utilizarán en la búsqueda de registros:</B></FONT><BR />"
	Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""Main_ISSSTE.asp"" METHOD=""GET"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""38"" />"
		Response.Write "<TABLE BORDER=""0"" CELLPADING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Clave:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""ShortName"" ID=""ShortNameTxt"" SIZE=""5"" MAXLENGTH=""5"" VALUE=""" & oRequest("ShortName").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""Name"" ID=""NameTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("Name").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""2"" ALIGN=""RIGHT""><BR /><INPUT TYPE=""SUBMIT"" NAME=""DoSearch"" ID=""DoSearchBtn"" VALUE=""Buscar Registros"" CLASS=""Buttons"" /></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	Response.Write "</FORM>"

	Display38SearchForm = Err.number
	Err.Clear
End Function 

Function Display62SearchForm(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the search form for the employees
'         linceses in the DocumentsForLicenses table
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "Display62SearchForm"

	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Indique los criterios que se utilizarán en la búsqueda de registros:</B></FONT><BR />"
	Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""Main_ISSSTE.asp"" METHOD=""GET"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""62"" />"
		Response.Write "<TABLE BORDER=""0"" CELLPADING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. de oficio:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""DocumentForLicenseNumber"" ID=""DocumentForLicenseNumberTxt"" SIZE=""25"" MAXLENGTH=""25"" VALUE=""" & oRequest("S_DOCUMENT_FOR_LICENSE_NUMBER_EMPLOYEE").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. de oficio de cancelación:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""DocumentForCancelLicenseNumber"" ID=""DocumentForCancelLicenseNumberTxt"" SIZE=""25"" MAXLENGTH=""25"" VALUE=""" & oRequest("S_DOCUMENT_FOR_CANCEL_LICENSE_NUMBER_EMPLOYEE").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""2"" ALIGN=""RIGHT""><BR /><INPUT TYPE=""SUBMIT"" NAME=""DoSearch"" ID=""DoSearchBtn"" VALUE=""Buscar Registros"" CLASS=""Buttons"" /></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	Response.Write "</FORM>"

	Display62SearchForm = Err.number
	Err.Clear
End Function 

Function Display351SearchForm(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the search form for the EmployeesKardex table.
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "Display351SearchForm"

	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Indique los criterios que se utilizarán en la búsqueda de registros:</B></FONT><BR />"
	Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""Main_ISSSTE.asp"" METHOD=""GET"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""351"" />"
		Response.Write "<TABLE BORDER=""0"" CELLPADING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Inicio de trámite:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("StartStartYear").Item), CInt(oRequest("StartStartMonth").Item), CInt(oRequest("StartStartDay").Item), "StartStartYear", "StartStartMonth", "StartStartDay", N_FORM_START_YEAR, Year(Date()), True, True)
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("EndStartYear").Item), CInt(oRequest("EndStartMonth").Item), CInt(oRequest("EndStartDay").Item), "EndStartYear", "EndStartMonth", "EndStartDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Procedimiento para:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""KardexTypeID"" ID=""KardexTypeIDLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "KardexTypes", "KardexTypeID", "KardexTypeName", "(KardexTypeID>-1)", "KardexTypeID", "", "", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Propuesto por:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""KardexOriginID"" ID=""KardexOriginIDLst"" SIZE=""4"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "KardexOrigins", "KardexOriginID", "KardexOriginName", "", "KardexOriginName", "", "", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""PersonName"" ID=""PersonNameTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("PersonName").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido paterno:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""PersonLastName"" ID=""PersonLastNameTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("PersonLastName").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido materno:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""PersonLastName2"" ID=""PersonLastName2Txt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("PersonLastName2").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Puesto:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""PositionID"" ID=""PositionIDLst"" SIZE=""6"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Positions", "PositionID", "PositionShortName, PositionName", "(PositionTypeID In (1,2)) And (PositionID>-1) And (Active=1)", "PositionShortName", "", "", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Unidad administrativa:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""AreaID"" ID=""AreaIDLst"" SIZE=""6"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(AreaID>-1) And (ParentID=-1) And (Active=1)", "AreaCode", "", "", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR><TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2""><INPUT TYPE=""CHECKBOX"" NAME=""HasDate1"" ID=""HasDate1Chk"" VALUE=""1"" />&nbsp;Enviados a registro de bolsa de trabajo</FONT></TD></TR>"
			Response.Write "<TR><TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2""><INPUT TYPE=""CHECKBOX"" NAME=""HasDate2"" ID=""HasDate2Chk"" VALUE=""1"" />&nbsp;Enviados a registro de escalafón</FONT></TD></TR>"
			Response.Write "<TR><TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2""><INPUT TYPE=""CHECKBOX"" NAME=""HasDate3"" ID=""HasDate3Chk"" VALUE=""1"" />&nbsp;Enviados al área de recursos humanos</FONT></TD></TR>"
			Response.Write "<TR><TD COLSPAN=""2"" ALIGN=""RIGHT""><BR /><INPUT TYPE=""SUBMIT"" NAME=""DoSearch"" ID=""DoSearchBtn"" VALUE=""Buscar Registros"" CLASS=""Buttons"" /></TD></TR>"
		Response.Write "</TABLE>"
	Response.Write "</FORM>"

	Display351SearchForm = Err.number
	Err.Clear
End Function 

Function Display352SearchForm(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the search form for the EmployeesKardex table.
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "Display352SearchForm"

	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Indique los criterios que se utilizarán en la búsqueda de registros:</B></FONT><BR />"
	Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""Main_ISSSTE.asp"" METHOD=""GET"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""352"" />"
		Response.Write "<TABLE BORDER=""0"" CELLPADING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""PersonName"" ID=""PersonNameTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("PersonName").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido paterno:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""PersonLastName"" ID=""PersonLastNameTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("PersonLastName").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido materno:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""PersonLastName2"" ID=""PersonLastName2Txt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("PersonLastName2").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
'			Response.Write "<TR>"
'				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Tipo de puesto:&nbsp;</FONT></TD>"
'				Response.Write "<TD><SELECT NAME=""PositionTypeID"" ID=""PositionTypeIDLst"" SIZE=""1"" CLASS=""Lists"">"
'					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
'					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "PositionTypes2", "PositionTypeID", "PositionTypeName", "", "PositionTypeName", "", "", sErrorDescription)
'				Response.Write "</SELECT></TD>"
'			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Inicio de trámite:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("StartStartYear").Item), CInt(oRequest("StartStartMonth").Item), CInt(oRequest("StartStartDay").Item), "StartStartYear", "StartStartMonth", "StartStartDay", N_FORM_START_YEAR, Year(Date()), True, True)
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("EndStartYear").Item), CInt(oRequest("EndStartMonth").Item), CInt(oRequest("EndStartDay").Item), "EndStartYear", "EndStartMonth", "EndStartDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Procedimiento para:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""KardexTypeID"" ID=""KardexTypeIDCmb"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "KardexTypes", "KardexTypeID", "KardexTypeName", "(KardexTypeID>-1) And (Active=1)", "KardexTypeName", oRequest("KardexTypeID").Item, "Ninguno;;;-1", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Puesto:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""PositionID"" ID=""PositionIDCmb"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Positions", "PositionID", "PositionShortName, PositionName", "(PositionTypeID In (1,2)) And (PositionID>-1) And (Active=1)", "PositionShortName, PositionName", oRequest("PositionID").Item, "Ninguno;;;-1", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Unidad administrativa:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""AreaID"" ID=""AreaIDCmb"" SIZE=""6"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
				'onChange=""if (this.value == '-1') {document.HierarchyMenuIFrame.location.href = 'HierarchyMenu.asp';} else {document.HierarchyMenuIFrame.location.href = 'HierarchyMenu.asp?Action=SubAreas&TargetField=SearchFrm.SubAreaID&AreaID=' + this.value;}"">"
					Response.Write "<OPTION VALUE="""">Todas</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(ParentID=-1) And (AreaID>-1) And (EndDate=30000000)", "AreaCode, AreaName", oRequest("AreaID").Item, "", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
'			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2"">"
'					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SubAreaID"" ID=""SubAreaIDHdn"" VALUE=""" & oRequest("SubAreaID").Item & """ />"
'					Response.Write "<IFRAME SRC=""HierarchyMenu.asp"
'						If Len(oRequest("SubAreaID").Item) > 0 Then
'							Response.Write "?Action=SubAreas&TargetField=SearchFrm.SubAreaID&AreaID=" & oRequest("AreaID").Item & "&SubAreaID=" & oRequest("SubAreaID").Item
'						End If
'					Response.Write """ NAME=""HierarchyMenuIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""105""></IFRAME>"
'				Response.Write "</FONT></TD>"
'			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""2"" ALIGN=""RIGHT""><BR /><INPUT TYPE=""SUBMIT"" NAME=""DoSearch"" ID=""DoSearchBtn"" VALUE=""Buscar Registros"" CLASS=""Buttons"" /></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	Response.Write "</FORM>"

	Display352SearchForm = Err.number
	Err.Clear
End Function 

Function Display353SearchForm(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the search form for the EmployeesKardex table.
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "Display353SearchForm"

	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Indique los criterios que se utilizarán en la búsqueda de registros:</B></FONT><BR />"
	Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""Main_ISSSTE.asp"" METHOD=""GET"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""353"" />"
		Response.Write "<TABLE BORDER=""0"" CELLPADING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Procedimiento para:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""KardexTypeID"" ID=""KardexTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "KardexTypes", "KardexTypeID", "KardexTypeName", "(Active=1)", "KardexTypeName", oRequest("KardexTypeID").Item, "Ninguno;;;-1", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Inicio de trámite:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("StartStartYear").Item), CInt(oRequest("StartStartMonth").Item), CInt(oRequest("StartStartDay").Item), "StartStartYear", "StartStartMonth", "StartStartDay", N_FORM_START_YEAR, Year(Date()), True, True)
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("EndStartYear").Item), CInt(oRequest("EndStartMonth").Item), CInt(oRequest("EndStartDay").Item), "EndStartYear", "EndStartMonth", "EndStartDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Propuesto por:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""KardexOriginID"" ID=""KardexOriginIDCmb"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "KardexOrigins", "KardexOriginID", "KardexOriginName", "(Active=1)", "KardexOriginName", oRequest("KardexOriginID").Item, "Ninguno;;;-1", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""PersonName"" ID=""PersonNameTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("PersonName").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido paterno:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""PersonLastName"" ID=""PersonLastNameTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("PersonLastName").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido materno:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""PersonLastName2"" ID=""PersonLastName2Txt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("PersonLastName2").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Puesto:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""PositionID"" ID=""PositionIDCmb"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Positions", "PositionID", "PositionShortName, PositionName", "(EndDate=30000000) And (Active=1)", "PositionShortName, PositionName", oRequest("PositionID").Item, "Ninguno;;;-1", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Áreas:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""AreaID"" ID=""AreaIDLst"" SIZE=""1"" CLASS=""Lists"" onChange=""if (this.value == '-1') {document.HierarchyMenuIFrame.location.href = 'HierarchyMenu.asp';} else {document.HierarchyMenuIFrame.location.href = 'HierarchyMenu.asp?Action=SubAreas&TargetField=' + this.form.name + '.SubAreaID&AreaID=' + this.value;}"">"
					Response.Write "<OPTION VALUE="""">Todas</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(ParentID=-1) And (AreaID>-1) And (EndDate=30000000) And (Active=1)", "AreaCode, AreaName", oRequest("AreaID").Item, "", sErrorDescription)
				Response.Write "</SELECT>"
			Response.Write "</TR>"
			Response.Write "<TR><TD COLSPAN=""2"" VALIGN=""TOP"">"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SubAreaID"" ID=""SubAreaIDHdn"" VALUE=""" & oRequest("SubAreaID").Item & """ />"
				Response.Write "<IFRAME SRC=""HierarchyMenu.asp"
					If Len(oRequest("SubAreaID").Item) > 0 Then
						Response.Write "?Action=SubAreas&TargetField=ReportFrm.SubAreaID&AreaID=" & oRequest("AreaID").Item & "&SubAreaID=" & oRequest("SubAreaID").Item
					End If
				Response.Write """ NAME=""HierarchyMenuIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""105""></IFRAME>"
			Response.Write "</TD></TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Requisitos documentales por grupo:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""RequirementsTypeID"" ID=""RequirementsTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "RequirementsTypes", "RequirementsTypeID", "RequirementsTypeName", "(Active=1)", "RequirementsTypeName", oRequest("RequirementsTypeID").Item, "Ninguno;;;-1", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fecha de recepción de documentos:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("StartDocumentsYear").Item), CInt(oRequest("StartDocumentsMonth").Item), CInt(oRequest("StartDocumentsDay").Item), "StartDocumentsYear", "StartDocumentsMonth", "StartDocumentsDay", N_FORM_START_YEAR, Year(Date()), True, True)
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("EndDocumentsYear").Item), CInt(oRequest("EndDocumentsMonth").Item), CInt(oRequest("EndDocumentsDay").Item), "EndDocumentsYear", "EndDocumentsMonth", "EndDocumentsDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fecha de evaluación de conocimientos:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("StartKnowledgeYear").Item), CInt(oRequest("StartKnowledgeMonth").Item), CInt(oRequest("StartKnowledgeDay").Item), "StartKnowledgeYear", "StartKnowledgeMonth", "StartKnowledgeDay", N_FORM_START_YEAR, Year(Date()), True, True)
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("EndKnowledgeYear").Item), CInt(oRequest("EndKnowledgeMonth").Item), CInt(oRequest("EndKnowledgeDay").Item), "EndKnowledgeYear", "EndKnowledgeMonth", "EndKnowledgeDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Estatus de la evaluación de conocimientos:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""KnowledgeStatusID"" ID=""KnowledgeStatusIDCmb"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "StatusKnowledges", "StatusID", "StatusName", "(Active=1)", "StatusName", oRequest("KnowledgeStatusID").Item, "Ninguno;;;-1", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fecha de evaluación psicológica:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("StartPsychologicYear").Item), CInt(oRequest("StartPsychologicMonth").Item), CInt(oRequest("StartPsychologicDay").Item), "StartPsychologicYear", "StartPsychologicMonth", "StartPsychologicDay", N_FORM_START_YEAR, Year(Date()), True, True)
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("EndPsychologicYear").Item), CInt(oRequest("EndPsychologicMonth").Item), CInt(oRequest("EndPsychologicDay").Item), "EndPsychologicYear", "EndPsychologicMonth", "EndPsychologicDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Estatus de la evaluación psicológica:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""PsychologicStatusID"" ID=""PsychologicStatusIDCmb"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "StatusPsychologics", "StatusID", "StatusName", "(Active=1)", "StatusName", oRequest("KnowledgeStatusID").Item, "Ninguno;;;-1", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"

			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fecha de envío a registro en bolsa de trabajo:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("StartRegistration1Year").Item), CInt(oRequest("StartRegistration1Month").Item), CInt(oRequest("StartRegistration1Day").Item), "StartRegistration1Year", "StartRegistration1Month", "StartRegistration1Day", N_FORM_START_YEAR, Year(Date()), True, True)
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("EndRegistration1Year").Item), CInt(oRequest("EndRegistration1Month").Item), CInt(oRequest("EndRegistration1Day").Item), "EndRegistration1Year", "EndRegistration1Month", "EndRegistration1Day", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fecha de envío a registro en escalafón:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("StartRegistration2Year").Item), CInt(oRequest("StartRegistration2Month").Item), CInt(oRequest("StartRegistration2Day").Item), "StartRegistration2Year", "StartRegistration2Month", "StartRegistration2Day", N_FORM_START_YEAR, Year(Date()), True, True)
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("EndRegistration2Year").Item), CInt(oRequest("EndRegistration2Month").Item), CInt(oRequest("EndRegistration2Day").Item), "EndRegistration2Year", "EndRegistration2Month", "EndRegistration2Day", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fecha de envío al área de recursos humanos:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("StartRegistration3Year").Item), CInt(oRequest("StartRegistration3Month").Item), CInt(oRequest("StartRegistration3Day").Item), "StartRegistration3Year", "StartRegistration3Month", "StartRegistration3Day", N_FORM_START_YEAR, Year(Date()), True, True)
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("EndRegistration3Year").Item), CInt(oRequest("EndRegistration3Month").Item), CInt(oRequest("EndRegistration3Day").Item), "EndRegistration3Year", "EndRegistration3Month", "EndRegistration3Day", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""2"" ALIGN=""RIGHT""><BR /><INPUT TYPE=""SUBMIT"" NAME=""DoSearch"" ID=""DoSearchBtn"" VALUE=""Buscar Registros"" CLASS=""Buttons"" /></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	Response.Write "</FORM>"

	Display353SearchForm = Err.number
	Err.Clear
End Function 

Function Display356SearchForm(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the search form for the EmployeesKardex2 table.
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "Display356SearchForm"

	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Indique los criterios que se utilizarán en la búsqueda de registros:</B></FONT><BR />"
	Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""Main_ISSSTE.asp"" METHOD=""GET"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""" & oRequest("SectionID").Item & """ />"
		Response.Write "<TABLE BORDER=""0"" CELLPADING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del empleado:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE=""" & oRequest("EmployeeNumber").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeName"" ID=""EmployeeNameTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("EmployeeName").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido paterno:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName"" ID=""EmployeeLastNameTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("EmployeeLastName").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido materno:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName2"" ID=""EmployeeLastName2Txt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("EmployeeLastName2").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			If CInt(oRequest("SectionID").Item) <> 356 Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Empresas:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""CompanyID"" ID=""CompanyID"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write "<OPTION VALUE=""-1"">Ninguna</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Companies", "CompanyID", "CompanyName", "(ParentID>-1) And (EndDate=30000000) And (Active=1)", "CompanyName", oRequest("CompanyID").Item, "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
			End If
			If CInt(oRequest("SectionID").Item) = 356 Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Puesto:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""PositionID"" ID=""PositionIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todos</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Positions", "PositionID", "PositionShortName, PositionName", "(EndDate=30000000) And (Active=1)", "PositionShortName, PositionName", oRequest("PositionID").Item, "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Unidad administrativa:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""AreaID"" ID=""AreaIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todos</OPTION>"
						'Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(ParentID=-1) And (EndDate=30000000) And (Active=1)", "AreaCode, AreaName", oRequest("AreaID").Item, "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(0, "xxx", N_FORM_START_YEAR, Year(Date()), True, True) & "</FONT></TD>"
				Response.Write "</TR>"
			Else
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de tabulador:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todos</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EmployeeTypes", "EmployeeTypeID", "EmployeeTypeName", "(EndDate=30000000) And (Active=1)", "EmployeeTypeName", oRequest("EmployeeTypeID").Item, "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Áreas:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""AreaID"" ID=""AreaIDLst"" SIZE=""1"" CLASS=""Lists"" onChange=""if (this.value == '-1') {document.HierarchyMenuIFrame.location.href = 'HierarchyMenu.asp';} else {document.HierarchyMenuIFrame.location.href = 'HierarchyMenu.asp?Action=SubAreas&TargetField=' + this.form.name + '.SubAreaID&AreaID=' + this.value;}"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(ParentID=-1) And (AreaID>-1) And (EndDate=30000000) And (Active=1)", "AreaCode, AreaName", oRequest("AreaID").Item, "", sErrorDescription)
					Response.Write "</SELECT><BR />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SubAreaID"" ID=""SubAreaIDHdn"" VALUE=""" & oRequest("SubAreaID").Item & """ />"
					Response.Write "<IFRAME SRC=""HierarchyMenu.asp"
						If Len(oRequest("SubAreaID").Item) > 0 Then
							Response.Write "?Action=SubAreas&TargetField=ReportFrm.SubAreaID&AreaID=" & oRequest("AreaID").Item & "&SubAreaID=" & oRequest("SubAreaID").Item
						End If
					Response.Write """ NAME=""HierarchyMenuIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""105""></IFRAME></TD>"
				Response.Write "</TR>"
			End If
			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""2"" ALIGN=""RIGHT""><BR /><INPUT TYPE=""SUBMIT"" NAME=""DoSearch"" ID=""DoSearchBtn"" VALUE=""Buscar Registros"" CLASS=""Buttons"" /></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	Response.Write "</FORM>"

	Display356SearchForm = Err.number
	Err.Clear
End Function 

Function Display371SearchResults(oRequest, oADODBConnection, bDocsLibrary, sErrorDescription)
'************************************************************
'Purpose: To display the search results for the Documents table
'Inputs:  oRequest, oADODBConnection, bDocsLibrary
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "Display371SearchResults"
	Dim sTable
	Dim sIcon
	Dim oRecordset
	Dim aReportMenu()
	Dim sReportMenuData
	Dim sTemp
	Dim iIndex
	Dim lErrorNumber

	sTable = "Documents"
	sIcon = "MnLeftArrows"
	If bDocsLibrary Then
		sTable = "DocsLibrary"
		sIcon = "MnDocument"
	End If
	sErrorDescription = "No se pudo obtener la información de la normateca."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From " & sTable & " Where (DocumentID>-1) Order By DocumentName, StartDate Desc", "Main_ISSSTELib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			iIndex = 0
			Do While Not oRecordset.EOF
				iIndex = iIndex + 1
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			ReDim aReportMenu(iIndex)
			iIndex = 0
			oRecordset.MoveFirst
			Do While Not oRecordset.EOF
				sReportMenuData = ""
				sReportMenuData = CleanStringForHTML(CStr(oRecordset.Fields("DocumentName").Value)) & LIST_SEPARATOR
				sTemp = ""
				sTemp = CStr(oRecordset.Fields("Description").Value)
				Err.Clear
				sReportMenuData = sReportMenuData & CleanStringForHTML(sTemp)
				If (Not bDocsLibrary) Or ((aLoginComponent(N_USER_PERMISSIONS4_LOGIN) And N_31_PERMISSIONS4) = N_31_PERMISSIONS4) Then
					If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then sReportMenuData = sReportMenuData & "<BR /><A HREF=""" & GetASPFileName("") & "?SectionID=371&Change=1&DocumentID=" & CStr(oRecordset.Fields("DocumentID").Value) & """><IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" HSPACE=""3"" />Modificar</A><IMG SRC=""Images/Transparent.gif"" WIDTH=""40"" HEIGHT=""1"" />"
					If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS Then sReportMenuData = sReportMenuData & "<A HREF=""" & GetASPFileName("") & "?SectionID=371&Delete=1&DocumentID=" & CStr(oRecordset.Fields("DocumentID").Value) & """><IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar"" BORDER=""0"" HSPACE=""3"" />Eliminar</A>"
				End If
				sReportMenuData = sReportMenuData & LIST_SEPARATOR & "Images/" & sIcon & ".gif" & LIST_SEPARATOR & _
								  UPLOADED_PATH & sTable & "/" & CStr(oRecordset.Fields("FilePath").Value) & """ TARGET=""Docs"
				sReportMenuData = sReportMenuData & LIST_SEPARATOR & "-1"
				aReportMenu(iIndex) = Split(sReportMenuData, LIST_SEPARATOR, -1, vbBinaryCompare)
				iIndex = iIndex + 1
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
			aMenuComponent(A_ELEMENTS_MENU) = aReportMenu
			aMenuComponent(B_USE_DIV_MENU) = True
			Response.Write "<TABLE WIDTH=""900"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				If bDocsLibrary Then
					Call DisplayMenuInTwoColumns(aMenuComponent)
				Else
					Call DisplayMenuInThreeSmallColumns(aMenuComponent)
				End If
			Response.Write "</TABLE>"
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			If bDocsLibrary Then
				sErrorDescription = "No existen documentos registrados en la normateca."
			Else
				sErrorDescription = "No existen procediminetos registrados en el sistema."
			End If
		End If
	End If

	Set oRecordset = Nothing
	Display371SearchResults = lErrorNumber
	Err.Clear
End Function 

Function Display352SearchResults(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the search results for the EmployeesKardex3 tables
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "Display352SearchResults"
	Dim oRecordset
	Dim sFontBegin
	Dim sFontEnd
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesKardex3.*, KardexTypeName, PositionShortName, PositionName From EmployeesKardex3, KardexTypes, Positions, Areas Where (EmployeesKardex3.KardexTypeID=KardexTypes.KardexTypeID) And (EmployeesKardex3.PositionID=Positions.PositionID) And (EmployeesKardex3.AreaID=Areas.AreaID) And (Areas.EndDate=30000000) And (Positions.EndDate=30000000) " & aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " Order By EmployeesKardex3.StartDate, PersonLastName, PersonLastName2, PersonName", "Main_ISSSTELib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Response.Write "<TABLE BORDER="""
			If bForExport Then
				Response.Write "1"
			Else
				Response.Write "0"
			End If
		Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			asColumnsTitles = Split("Procedimiento para,Nombre del candidato,Puesto,Fecha de inicio de trámite,Estatus,Fecha del último estatus", ",", -1, vbBinaryCompare)
			asCellWidths = Split("200,200,200,200,200,200", ",", -1, vbBinaryCompare)

			If bForExport Then
				lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
			Else
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If
			End If

			asCellAlignments = Split(",,,,,", ",", -1, vbBinaryCompare)
			Do While Not oRecordset.EOF
				sFontBegin = "<FONT COLOR=""#00D200"">"
				sFontEnd = "</FONT>"
				If CLng(Left(GetSerialNumberForDate(DateAdd("m", -6, Now())), Len("00000000"))) >= CLng(oRecordset.Fields("StartDate").Value) Then
					sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
				ElseIf CLng(Left(GetSerialNumberForDate(DateAdd("m", -3, Now())), Len("00000000"))) >= CLng(oRecordset.Fields("StartDate").Value) Then
					sFontBegin = "<FONT COLOR=""#D2D200"">"
				End If
				sRowContents = sFontBegin & CleanStringForHTML(CStr(oRecordset.Fields("KardexTypeName").Value)) & sFontEnd
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & CleanStringForHTML(CStr(oRecordset.Fields("PersonLastName").Value) & " " & CStr(oRecordset.Fields("PersonLastName2").Value) & ", " & CStr(oRecordset.Fields("PersonName").Value)) & sFontEnd
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value)) & sFontEnd
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1) & sFontEnd
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin
					If CLng(oRecordset.Fields("Registration1Date").Value) > 0 Then
						sRowContents = sRowContents & "Envío a registro en bolsa de trabajo"
					ElseIf CLng(oRecordset.Fields("Registration2Date").Value) > 0 Then
						sRowContents = sRowContents & "Envío a registro en escalafón"
					ElseIf CLng(oRecordset.Fields("Registration3Date").Value) > 0 Then
						sRowContents = sRowContents & "Envío al área de recursos humanos"
					ElseIf CLng(oRecordset.Fields("PsychologicDate").Value) > 0 Then
						sRowContents = sRowContents & "Evaluación psicológica"
					ElseIf CLng(oRecordset.Fields("KnowledgeDate").Value) > 0 Then
						sRowContents = sRowContents & "Evaluación de conocimientos"
					ElseIf CLng(oRecordset.Fields("DocumentsDate").Value) > 0 Then
						sRowContents = sRowContents & "Recepción de documentos"
					Else
						sRowContents = sRowContents & "Inicio de trámite"
					End If
				sRowContents = sRowContents & sFontEnd
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("ModifyDate").Value), -1, -1, -1) & sFontEnd

				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If

				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
		Response.Write "</TABLE><BR />"
	End If

	Display352SearchResults = lErrorNumber
	Err.Clear
End Function 

Function Display353SearchResults(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the search results for the EmployeesKardex3 tables
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "Display353SearchResults"
	Dim sNames
	Dim oRecordset
	Dim sFontBegin
	Dim sFontEnd
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesKardex3.*, KardexTypeName, KardexOriginName, PositionShortName, PositionName, AreaCode, AreaName, RequirementsTypeName, StatusKnowledges.StatusID As KnowledgeStatusID, StatusKnowledges.StatusName As KnowledgeStatusName, StatusPsychologics.StatusID As PsychologicStatusID, StatusPsychologics.StatusName As PsychologicStatusName From EmployeesKardex3, KardexTypes, KardexOrigins, Positions, Areas, RequirementsTypes, StatusKnowledges, StatusPsychologics Where (EmployeesKardex3.KardexTypeID=KardexTypes.KardexTypeID) And (EmployeesKardex3.KardexOriginID=KardexOrigins.KardexOriginID) And (EmployeesKardex3.PositionID=Positions.PositionID) And (EmployeesKardex3.AreaID=Areas.AreaID) And (EmployeesKardex3.RequirementsTypeID=RequirementsTypes.RequirementsTypeID) And (EmployeesKardex3.RequirementsTypeID=RequirementsTypes.RequirementsTypeID) And (EmployeesKardex3.KnowledgeStatusID=StatusKnowledges.StatusID) And (EmployeesKardex3.PsychologicStatusID=StatusPsychologics.StatusID) And (Areas.EndDate=30000000) And (Positions.EndDate=30000000) " & aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " Order By EmployeesKardex3.StartDate, PersonLastName, PersonLastName2, PersonName", "Main_ISSSTELib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Response.Write "<TABLE BORDER="""
			If bForExport Then
				Response.Write "1"
			Else
				Response.Write "0"
			End If
		Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			asColumnsTitles = Split("Procedimiento para,Fecha de inicio de trámite,Propuesto por,Nombre del aspirante,Puesto,Área,Requisitos documentales por grupo,Requisitos documentales,Fecha entrega documentos,Fecha de la evaluación de conocimientos,Estatus de la evaluación de conocimientos,Fecha de la evaluación psicológica,Estatus de la evaluación psicológica,Fecha de envío a registro en bolsa de trabajo,Fecha de envío a registro de escalafón,Fecha de envío al área de RH", ",", -1, vbBinaryCompare)
			asCellWidths = Split("200,200,200,200,200,200,200,200,200,200,200,200,200,200,200", ",", -1, vbBinaryCompare)

			If bForExport Then
				lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
			Else
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If
			End If

			asCellAlignments = Split(",,,,,,,,,,,,,,", ",", -1, vbBinaryCompare)
			Do While Not oRecordset.EOF
				sFontBegin = "<FONT COLOR=""#00D200"">"
				sFontEnd = "</FONT>"
				If CLng(Left(GetSerialNumberForDate(DateAdd("m", -6, Now())), Len("00000000"))) >= CLng(oRecordset.Fields("StartDate").Value) Then
					sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
				ElseIf CLng(Left(GetSerialNumberForDate(DateAdd("m", -3, Now())), Len("00000000"))) >= CLng(oRecordset.Fields("StartDate").Value) Then
					sFontBegin = "<FONT COLOR=""#D2D200"">"
				End If
				sRowContents = sFontBegin & CleanStringForHTML(CStr(oRecordset.Fields("KardexTypeName").Value)) & sFontEnd
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1) & sFontEnd
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & CleanStringForHTML(CStr(oRecordset.Fields("KardexOriginName").Value)) & sFontEnd
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & CleanStringForHTML(CStr(oRecordset.Fields("PersonLastName").Value) & " " & CStr(oRecordset.Fields("PersonLastName2").Value) & ", " & CStr(oRecordset.Fields("PersonName").Value)) & sFontEnd
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value)) & sFontEnd
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value) & ". " & CStr(oRecordset.Fields("AreaName").Value)) & sFontEnd
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & CleanStringForHTML(CStr(oRecordset.Fields("RequirementsTypeName").Value)) & sFontEnd
				Call GetNameFromTable(oADODBConnection, "KardexRequirements", CStr(oRecordset.Fields("Requirements").Value), "", "<BR />", sNames, sErrorDescription)
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sNames & sFontEnd
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin
					If CLng(oRecordset.Fields("DocumentsDate").Value) > 0 Then
						sRowContents = sRowContents & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("DocumentsDate").Value), -1, -1, -1)
					Else
						sRowContents = sRowContents & "<CENTER>---</CENTER>"
					End If
				sRowContents = sRowContents & sFontEnd
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin
					If CLng(oRecordset.Fields("KnowledgeDate").Value) > 0 Then
						sRowContents = sRowContents & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("KnowledgeDate").Value), -1, -1, -1)
					Else
						sRowContents = sRowContents & "<CENTER>---</CENTER>"
					End If
				sRowContents = sRowContents & sFontEnd
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin
					If CLng(oRecordset.Fields("KnowledgeStatusID").Value) > -1 Then
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("KnowledgeStatusName").Value))
					Else
						sRowContents = sRowContents & "&nbsp;"
					End If
				sRowContents = sRowContents & sFontEnd
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin
					If CLng(oRecordset.Fields("PsychologicDate").Value) > 0 Then
						sRowContents = sRowContents & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("PsychologicDate").Value), -1, -1, -1)
					Else
						sRowContents = sRowContents & "<CENTER>---</CENTER>"
					End If
				sRowContents = sRowContents & sFontEnd
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin
					If CLng(oRecordset.Fields("PsychologicStatusID").Value) > -1 Then
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("PsychologicStatusName").Value))
					Else
						sRowContents = sRowContents & "&nbsp;"
					End If
				sRowContents = sRowContents & sFontEnd
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin
					If CLng(oRecordset.Fields("Registration1Date").Value) > 0 Then
						sRowContents = sRowContents & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("Registration1Date").Value), -1, -1, -1)
					ElseIf (CLng(oRecordset.Fields("Registration2Date").Value) > 0) Or (CLng(oRecordset.Fields("Registration3Date").Value) > 0) Then
						sRowContents = sRowContents & "<CENTER>No aplica</CENTER>"
					Else
						sRowContents = sRowContents & "&nbsp;"
					End If
				sRowContents = sRowContents & sFontEnd
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin
					If CLng(oRecordset.Fields("Registration2Date").Value) > 0 Then
						sRowContents = sRowContents & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("Registration2Date").Value), -1, -1, -1)
					ElseIf (CLng(oRecordset.Fields("Registration1Date").Value) > 0) Or (CLng(oRecordset.Fields("Registration3Date").Value) > 0) Then
						sRowContents = sRowContents & "<CENTER>No aplica</CENTER>"
					Else
						sRowContents = sRowContents & "&nbsp;"
					End If
				sRowContents = sRowContents & sFontEnd
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin
					If CLng(oRecordset.Fields("Registration3Date").Value) > 0 Then
						sRowContents = sRowContents & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("Registration3Date").Value), -1, -1, -1)
					ElseIf (CLng(oRecordset.Fields("Registration1Date").Value) > 0) Or (CLng(oRecordset.Fields("Registration2Date").Value) > 0) Then
						sRowContents = sRowContents & "<CENTER>No aplica</CENTER>"
					Else
						sRowContents = sRowContents & "&nbsp;"
					End If
				sRowContents = sRowContents & sFontEnd

				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If

				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
		Response.Write "</TABLE><BR />"
	End If

	Display353SearchResults = lErrorNumber
	Err.Clear
End Function 

Function Display356SearchResults(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the search results for the EmployeesKardex2 tables
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "Display356SearchResults"
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, Requirement1, Requirement2, JobNumber, PositionShortName, PositionName, AreaCode, AreaName From Employees, EmployeesKardex2, Jobs, Positions, Areas Where (Employees.EmployeeID=EmployeesKardex2.EmployeeID) And (Employees.JobID=Jobs.JobID) And (Jobs.PositionID=Positions.PositionID) And (Jobs.AreaID=Areas.AreaID) " & aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " Order By EmployeeLastName, EmployeeLastName2, EmployeeName", "Main_ISSSTELib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Response.Write "<TABLE BORDER="""
			If bForExport Then
				Response.Write "1"
			Else
				Response.Write "0"
			End If
		Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			asColumnsTitles = Split("No. de empleado,Nombre,Puesto,Fecha de registro,Unidad administrativa", ",", -1, vbBinaryCompare)
			asCellWidths = Split("200,200,200,200,200", ",", -1, vbBinaryCompare)

			If bForExport Then
				lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
			Else
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If
			End If

			asCellAlignments = Split(",,,,", ",", -1, vbBinaryCompare)
			Do While Not oRecordset.EOF
				sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
				If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & ", " & CStr(oRecordset.Fields("EmployeeName").Value))
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & ", " & CStr(oRecordset.Fields("EmployeeName").Value))
				End If
				'sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("JobNumber").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & "???"
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value) & ". " & CStr(oRecordset.Fields("AreaName").Value))

				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If

				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
		Response.Write "</TABLE><BR />"
	End If

	Display356SearchResults = lErrorNumber
	Err.Clear
End Function 

Function Display369SearchResults(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the search results for the EmployeesKardex2 tables
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "Display369SearchResults"
	Dim oRecordset
	Dim sIDs
	Dim sTemp
	Dim sNames
	Dim iIndex
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeName, EmployeeLastName, EmployeeLastName2, EmployeeNumber, JobNumber, PositionShortName, PositionName, AreaCode, AreaName, SchoolarshipName, CourseNames_1, CourseNames_2, CourseNames_3, CourseNames_4, CourseNames_5, CourseNames_6, CourseNames_7 From Employees, SADE_NewCourse, Jobs, Positions, Areas, Schoolarships Where (Employees.EmployeeID=SADE_NewCourse.EmployeeID) And (Employees.JobID=Jobs.JobID) And (Jobs.PositionID=Positions.PositionID) And (Jobs.AreaID=Areas.AreaID) And (SADE_NewCourse.SchoolarshipID=Schoolarships.SchoolarshipID) " & aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " Order By EmployeeLastName, EmployeeLastName2, EmployeeName", "Main_ISSSTELib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Response.Write "<TABLE WIDTH=""400"" BORDER="""
			If bForExport Then
				Response.Write "1"
			Else
				Response.Write "0"
			End If
		Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			asColumnsTitles = Split("Nombre,No. de empleado,Plaza,Puesto,Centro de trabajo,Escolaridad,Cursos", ",", -1, vbBinaryCompare)
			asCellWidths = Split("200,100,100,200,200,200,200", ",", -1, vbBinaryCompare)

			If bForExport Then
				lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
			Else
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If
			End If

			asCellAlignments = Split(",,,,,CENTER,CENTER", ",", -1, vbBinaryCompare)
			Do While Not oRecordset.EOF
				If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
					sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & ", " & CStr(oRecordset.Fields("EmployeeName").Value))
				Else
					sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & CStr(oRecordset.Fields("EmployeeName").Value))
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("JobNumber").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value) & ". " & CStr(oRecordset.Fields("AreaName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SchoolarshipName").Value))
				sIDs = -1
				For iIndex = 1 To 7
					sTemp = ""
					sTemp = CStr(oRecordset.Fields("CourseNames_" & iIndex).Value)
					If Len(sTemp) > 0 Then sIDs = sIDs & "," & sTemp
				Next
				Err.Clear
				Call GetNameFromTable(oADODBConnection, "SADE_Perfiles", sIDs, "", "<BR />", sNames, sErrorDescription)
				sRowContents = sRowContents & TABLE_SEPARATOR & sNames

				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If

				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
		Response.Write "</TABLE><BR />"
	End If

	Display369SearchResults = lErrorNumber
	Err.Clear
End Function 

Function Display38SearchResults(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the search results for the Areas and the
'         PaymentCenters tables
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "Display38SearchResults"
	Dim sCondition
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	sCondition = ""
	If Len(oRequest("ShortName").Item) > 0 Then sCondition = sCondition & " And (AreaCode='" & Right(("00000" & oRequest("ShortName").Item), Len("00000")) & "')"
	If Len(oRequest("Name").Item) > 0 Then sCondition = sCondition & " And (AreaName Like '" & S_WILD_CHAR & oRequest("Name").Item & S_WILD_CHAR & "')"
	sErrorDescription = "No se pudieron obtener los registros de la base de datos."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaCode, AreaName From Areas Where (ParentID>-1) " & sCondition & " Order By AreaCode, AreaName", "Main_ISSSTELib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>CENTROS DE TRABAJO</B></FONT><BR />"
		Response.Write "<TABLE WIDTH=""400"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			asColumnsTitles = Split("Clave,Nombre", ",", -1, vbBinaryCompare)
			asCellWidths = Split("150,250", ",", -1, vbBinaryCompare)

			If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
				lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			Else
				lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			End If

			asCellAlignments = Split(",", ",", -1, vbBinaryCompare)
			Do While Not oRecordset.EOF
				sRowContents = ""
				sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value))

				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)

				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
		Response.Write "</TABLE><BR />"
	End If

	sCondition = ""
	If Len(oRequest("ShortName").Item) > 0 Then sCondition = sCondition & " And (PaymentCenterShortName='" & Right(("00000" & oRequest("ShortName").Item), Len("00000")) & "')"
	If Len(oRequest("Name").Item) > 0 Then sCondition = sCondition & " And (PaymentCenterName Like '" & S_WILD_CHAR & oRequest("Name").Item & S_WILD_CHAR & "')"
	sErrorDescription = "No se pudieron obtener los registros de la base de datos."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaymentCenterShortName, PaymentCenterName From PaymentCenters Where (PaymentCenterID>-1) " & sCondition & " Order By PaymentCenterShortName, PaymentCenterName", "Main_ISSSTELib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>CENTROS DE PAGO</B></FONT><BR />"
		Response.Write "<TABLE WIDTH=""400"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			asColumnsTitles = Split("Clave,Nombre", ",", -1, vbBinaryCompare)
			asCellWidths = Split("150,250", ",", -1, vbBinaryCompare)

			If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
				lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			Else
				lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			End If

			asCellAlignments = Split(",", ",", -1, vbBinaryCompare)
			Do While Not oRecordset.EOF
				sRowContents = ""
				sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterName").Value))

				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)

				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
		Response.Write "</TABLE>"
	End If

	Display38SearchResults = lErrorNumber
	Err.Clear
End Function 

Function Display423SearchForm(oRequest, oADODBConnection, iSectionID, sErrorDescription)
'************************************************************
'Purpose: To display the search form for the EmployeesKardex5
'Inputs:  oRequest, oADODBConnection, iSectionID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "Display423SearchForm"

	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		Response.Write "function AddEmployeeIDToSearchList() {" & vbNewLine
			Response.Write "var oForm = document.SearchFrm;" & vbNewLine
			Response.Write "if (oForm.EmployeeID.value != '') {" & vbNewLine
				Response.Write "oForm.EmployeeID.value = '000000' + oForm.EmployeeID.value;" & vbNewLine
				Response.Write "AddItemToList(oForm.EmployeeID.value.substr(oForm.EmployeeID.value.length - 6), oForm.EmployeeID.value.substr(oForm.EmployeeID.value.length - 6), null, oForm.EmployeeIDs)" & vbNewLine
				Response.Write "SelectAllItemsFromList(oForm.EmployeeIDs);" & vbNewLine
				Response.Write "oForm.EmployeeID.value = '';" & vbNewLine
				Response.Write "oForm.EmployeeID.focus();" & vbNewLine
			Response.Write "}" & vbNewLine
		Response.Write "} // End of AddDisasterIDToSearchList" & vbNewLine

		Response.Write "function Show423Fields(sValue) {" & vbNewLine
			Response.Write "if (sValue == '1') {" & vbNewLine
				Response.Write "ShowDisplay(document.all['EmployeeNumberDiv']);" & vbNewLine
				Response.Write "HideDisplay(document.all['EmployeeNameDiv']);" & vbNewLine
				Response.Write "HideDisplay(document.all['EmployeeLastNameDiv']);" & vbNewLine
				Response.Write "HideDisplay(document.all['EmployeeLastName2Div']);" & vbNewLine
				Response.Write "HideDisplay(document.all['EmployeeRFCDiv']);" & vbNewLine
			Response.Write "} else {" & vbNewLine
				Response.Write "HideDisplay(document.all['EmployeeNumberDiv']);" & vbNewLine
				Response.Write "ShowDisplay(document.all['EmployeeNameDiv']);" & vbNewLine
				Response.Write "ShowDisplay(document.all['EmployeeLastNameDiv']);" & vbNewLine
				Response.Write "ShowDisplay(document.all['EmployeeLastName2Div']);" & vbNewLine
				Response.Write "ShowDisplay(document.all['EmployeeRFCDiv']);" & vbNewLine
			Response.Write "}" & vbNewLine
		Response.Write "} // End of Show423Fields" & vbNewLine
	Response.Write "//--></SCRIPT>" & vbNewLine
	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Indique los criterios que se utilizarán en la búsqueda de registros:</B></FONT><BR />"
	Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""Main_ISSSTE.asp"" METHOD=""GET"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""" & iSectionID & """ />"
		Response.Write "<TABLE BORDER=""0"" CELLPADING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR NAME=""EmployeeNumberDiv"" ID=""EmployeeNumberDiv"">"
				Response.Write "<TD VALIGN=""TOP"">"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Número de empleado:<BR /></FONT>"
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
					Response.Write "&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""EmployeeID"" ID=""EmployeeIDTxt"" SIZE=""10"" MAXLENGTH=""6"" VALUE=""" & oRequest("EmployeeID").Item & """ CLASS=""TextFields"" />"
					Response.Write "&nbsp;&nbsp;<A HREF=""javascript: AddEmployeeIDToSearchList();""><IMG SRC=""Images/BtnCrclAdd.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Agregar"" BORDER=""0"" /></A>&nbsp;&nbsp;<BR />"
				Response.Write "</TD>"
					Response.Write "<TD VALIGN=""TOP""><BR />"
						Response.Write "<SELECT NAME=""EmployeeIDs"" ID=""EmployeeIDsCmb"" SIZE=""6"" MULTIPLE=""1"" CLASS=""Lists"" STYLE=""width: 100px;"">"
							If Len(oRequest("EmployeeIDs").Item) > 0 Then
								For Each oItem In oRequest("EmployeeIDs")
									Response.Write "<OPTION VALUE=""" & oItem & """ SELECTED=""1"">" & oItem & "</OPTION>"
								Next
							End If
						Response.Write "</SELECT>"
						Response.Write "&nbsp;<A HREF=""javascript: RemoveSelectedItemsFromList(null, document.ReportFrm.EmployeeIDs); SelectAllItemsFromList(document.ReportFrm.EmployeeIDs);""><IMG SRC=""Images/BtnCrclDelete.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Quitar"" BORDER=""0""></A><BR />"
					Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR NAME=""EmployeeNameDiv"" ID=""EmployeeNameDiv"" STYLE=""display: none"">"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre del empleado:</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeName"" ID=""EmployeeNameTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("EmployeeName").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR NAME=""EmployeeLastNameDiv"" ID=""EmployeeLastNameDiv"" STYLE=""display: none"">"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido paterno:</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName"" ID=""EmployeeLastNameTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("EmployeeLastName").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR NAME=""EmployeeLastName2Div"" ID=""EmployeeLastName2Div"" STYLE=""display: none"">"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido materno:</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName2"" ID=""EmployeeLastName2Txt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("EmployeeLastName2").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR NAME=""EmployeeRFCDiv"" ID=""EmployeeRFCDiv"" STYLE=""display: none"">"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">RFC:</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeRFC"" ID=""EmployeeRFCTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("EmployeeRFC").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			If CInt(iSectionID) = 424 Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del empleado a suplir:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""OriginalEmployeeID"" ID=""OriginalEmployeeIDTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE=""" & oRequest("OriginalEmployeeID").Item & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
			End If
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Adscripción:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""AreaID"" ID=""AreaIDLst"" SIZE=""1"" CLASS=""Lists"" onChange=""if (this.value == '-1') {document.HierarchyMenuIFrame.location.href = 'HierarchyMenu.asp';} else {document.HierarchyMenuIFrame.location.href = 'HierarchyMenu.asp?Action=SubAreas&TargetField=' + this.form.name + '.SubAreaID&AreaID=' + this.value;}"">"
					Response.Write "<OPTION VALUE="""">Todas</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(ParentID>-1) And (AreaID>-1) And (EndDate=30000000) And (Active=1) And (CenterTypeID In (Select Distinct CenterTypeID From PositionsSpecialJourneysLKP))", "AreaCode, AreaName", oRequest("AreaID").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR />"
			Response.Write "</TR>"
			If CInt(iSectionID) <> 425 Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Turno:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""JourneyID"" ID=""JourneyIDLst"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "SpecialJourneys", "JourneyID", "JourneyShortName, JourneyName", "(RecordTypeID In (-1," & iSectionID & ")) And (Active=1)", "JourneyShortName", oRequest("JourneyID").Item, "", sErrorDescription)
					Response.Write "</SELECT><BR />"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Movimiento:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""MovementID"" ID=""MovementIDLst"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "SpecialJourneysMovements", "MovementID", "MovementShortName, MovementName", "(RecordTypeID In (-1," & iSectionID & ")) And (Active=1)", "MovementShortName", oRequest("MovementID").Item, "", sErrorDescription)
					Response.Write "</SELECT><BR />"
				Response.Write "</TR>"
			End If
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("StartStartYear").Item), CInt(oRequest("StartStartMonth").Item), CInt(oRequest("StartStartDay").Item), "StartStartYear", "StartStartMonth", "StartStartDay", 2009, Year(Date()), True, True)
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("EndStartYear").Item), CInt(oRequest("EndStartMonth").Item), CInt(oRequest("EndStartDay").Item), "EndStartYear", "EndStartMonth", "EndStartDay", 2009, Year(Date()), True, True)
				Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fecha de término:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("StartEndYear").Item), CInt(oRequest("StartEndMonth").Item), CInt(oRequest("StartEndDay").Item), "StartEndYear", "StartEndMonth", "StartEndDay", 2009, Year(Date()), True, True)
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("EndEndYear").Item), CInt(oRequest("EndEndMonth").Item), CInt(oRequest("EndEndDay").Item), "EndEndYear", "EndEndMonth", "EndEndDay", 2009, Year(Date()), True, True)
				Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Tipo de personal:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
					'Response.Write "<INPUT TYPE=""Radio"" NAME=""Internal"" ID=""InternalRd"" VALUE="""""
					'	If Len(oRequest("Internal").Item) = 0 Then
					'		Response.Write " CHECKED=""1"""
					'	End If
					'Response.Write " />Ambos&nbsp;&nbsp;&nbsp;"
					Response.Write "<INPUT TYPE=""Radio"" NAME=""Internal"" ID=""InternalRd"" VALUE=""1"""
						If StrComp(oRequest("Internal").Item, "0", vbBinaryCompare) <> 0 Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " onClick=""Show423Fields(this.value);"" />Interno&nbsp;&nbsp;&nbsp;<INPUT TYPE=""Radio"" NAME=""Internal"" ID=""InternalRd"" VALUE=""0"""
						If StrComp(oRequest("Internal").Item, "0", vbBinaryCompare) = 0 Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " onClick=""Show423Fields(this.value);"" />Externo<BR /><BR />"
				Response.Write "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""2"" ALIGN=""RIGHT""><BR /><INPUT TYPE=""SUBMIT"" NAME=""DoSearch"" ID=""DoSearchBtn"" VALUE=""Buscar Registros"" CLASS=""Buttons"" /></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	Response.Write "</FORM>"

	Display423SearchForm = Err.number
	Err.Clear
End Function 

Function DoCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, bSearchForm, bShowForm, sErrorDescription)
'************************************************************
'Purpose: To initialize and add, modify, or remove entires in
'         the EmployeesKardex or EmployeesKardex2 tables
'Inputs:  oRequest, oADODBConnection, aCatalogComponent
'Outputs: aCatalogComponent, bAction, bSearchForm, bShowForm, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DoCatalogsAction"
	Dim lErrorNumber

	bSearchForm = (Len(oRequest("Search").Item) > 0)
	bShowForm = ((Len(oRequest("New").Item) > 0) Or (Len(oRequest("Change").Item) > 0) Or (Len(oRequest("Delete").Item) > 0))
	Call InitializeCatalogs(oRequest)
	Call InitializeValuesForCatalogComponent(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
	aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) = CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)))

	Select Case oRequest("SectionID").Item
		Case "261"
		Case "267"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(9) = GenerateRandomCharactersSecuence(176) & "="
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(10) = GenerateRandomHexadecimalSecuence(8) & "-" & GenerateRandomHexadecimalSecuence(4) & "-" & GenerateRandomHexadecimalSecuence(4) & "-" & GenerateRandomHexadecimalSecuence(4) & "-" & GenerateRandomHexadecimalSecuence(12)
		Case Else
			aCatalogComponent(S_QUERY_CONDITION_CATALOG) = ""
	End Select
	If StrComp(GetASPFileName(""), "Export.asp", vbBinaryCompare) <> 0 Then
		If Len(oRequest("Add").Item) > 0 Then
			Select Case aCatalogComponent(S_TABLE_NAME_CATALOG)
				Case "EmployeesKardex2"
					sErrorDescription = "No se pudo agregar la información del registro de escalafón."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesKardex2 Where (EmployeeID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) & ")", "Main_ISSSTELib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				Case "PaymentsRecords"
					aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(3) = Replace(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(3), " ", "")
				Case "EmployeesDocs"
					aEmployeeComponent(N_ID_EMPLOYEE) = aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1)
					lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
			End Select
			If lErrorNumber = 0 Then
				'lErrorNumber = AddCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
				lErrorNumber = DoAction(CStr(oRequest("Action").Item), bShowForm, sErrorDescription)
			End If
			bAction = True
		ElseIf Len(oRequest("Modify").Item) > 0 Then
			lErrorNumber = ModifyCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
			bAction = True
		ElseIf Len(oRequest("Remove").Item) > 0 Then
			If StrComp(aCatalogComponent(S_TABLE_NAME_CATALOG), "Documents", vbBinaryCompare) = 0 Then
				lErrorNumber = GetCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					Call DeleteFile(SYSTEM_PHYSICAL_PATH & UPLOADED_PHYSICAL_PATH & "Documents/" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2), "")
				End If
			End If
			If StrComp(aCatalogComponent(S_TABLE_NAME_CATALOG), "DocsLibrary", vbBinaryCompare) = 0 Then
				lErrorNumber = GetCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					Call DeleteFile(SYSTEM_PHYSICAL_PATH & UPLOADED_PHYSICAL_PATH & "DocsLibrary/" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2), "")
				End If
			End If
			lErrorNumber = RemoveCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
			If lErrorNumber = 0 Then aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG))=-1
			bAction = True
		End If
	End If
	Select Case oRequest("SectionID").Item
		Case "261"
			If Len(oRequest("PrevAntiquityDate").Item) > 0 Then
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " And (EmployeeID=" & oRequest("EmployeeID").Item & ") And (AntiquityDate=" & oRequest("PrevAntiquityDate").Item & ")"
			ElseIf Len(oRequest("AntiquityDate").Item) > 0 Then
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " And (EmployeeID=" & oRequest("EmployeeID").Item & ") And (AntiquityDate=" & oRequest("AntiquityDate").Item & ")"
			ElseIf Len(oRequest("AntiquityDateYear").Item) > 0 Then
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " And (EmployeeID=" & oRequest("EmployeeID").Item & ") And (AntiquityDate=" & oRequest("AntiquityDateYear").Item & oRequest("AntiquityDateMonth").Item & oRequest("AntiquityDateDay").Item & ")"
			ElseIf Len(oRequest("EmployeeID").Item) > 0 Then
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " And (EmployeeID=" & oRequest("EmployeeID").Item & ")"
			End If
		Case "423", "424", "425", "426"
			If (lErrorNumber = 0) And (Len(oRequest("Add").Item) > 0) Or (Len(oRequest("Modify").Item) > 0) Then
				sErrorDescription = "No se pudo actualizar la información del empleado externo."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesSpecialJourneys Set EmployeeName='" & oRequest("EmployeeName").Item & "', EmployeeLastName='" & oRequest("EmployeeLastName").Item & "', EmployeeLastName2='" & oRequest("EmployeeLastName2").Item & "' Where (EmployeeID>=800000) And (RFC='" & oRequest("RFC").Item & "')", "Main_ISSSTELib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
	End Select
	If lErrorNumber = 0 Then
		If aCatalogComponent(N_ID_CATALOG) > -1 Then
			If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG))) > -1 Then
				lErrorNumber = GetCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
			End If
		Else
			lErrorNumber = GetCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
		End If
	End If
	If Len(oRequest("DoSearch").Item) > 0 Then
		aCatalogComponent(S_QUERY_CONDITION_CATALOG) = ""
		Select Case oRequest("SectionID").Item
			Case "261"
				If Len(oRequest("EmployeeID").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeeID=" & oRequest("EmployeeID").Item & ")"
				End If
				If Len(oRequest("AntiquityDate").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (AntiquityDate=" & oRequest("AntiquityDate").Item & ")"
				End If
			Case "267"
				If Len(oRequest("EmployeeID").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeeID=" & oRequest("EmployeeID").Item & ")"
				End If
				If ((CInt(oRequest("StartYear").Item) > 0) And (CInt(oRequest("StartMonth").Item) > 0) And (CInt(oRequest("StartDay").Item) > 0)) Or ((CInt(oRequest("EndYear").Item) > 0) And (CInt(oRequest("EndMonth").Item) > 0) And (CInt(oRequest("EndDay").Item) > 0)) Then Call GetStartAndEndDatesFromURL("Start", "End", "DocumentDate", True, aCatalogComponent(S_QUERY_CONDITION_CATALOG))
			Case "281"
				If Len(oRequest("Kardex5TypeID").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesKardex5.Kardex5TypeID In (" & oRequest("Kardex5TypeID").Item & "))"
				End If
				If Len(oRequest("Kardex5OriginID").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesKardex5.Kardex5OriginID In (" & oRequest("Kardex5OriginID").Item & "))"
				End If
				If Len(oRequest("EmployeeName").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeeName Like '" & S_WILD_CHAR & oRequest("EmployeeName").Item & S_WILD_CHAR & "')"
				End If
				If Len(oRequest("EmployeeLastName").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeeLastName Like '" & S_WILD_CHAR & oRequest("EmployeeLastName").Item & S_WILD_CHAR & "')"
				End If
				If Len(oRequest("EmployeeLastName2").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeeLastName2 Like '" & S_WILD_CHAR & oRequest("EmployeeLastName2").Item & S_WILD_CHAR & "')"
				End If
				If ((CInt(oRequest("StartStartYear").Item) > 0) And (CInt(oRequest("StartStartMonth").Item) > 0) And (CInt(oRequest("StartStartDay").Item) > 0)) Or ((CInt(oRequest("EndStartYear").Item) > 0) And (CInt(oRequest("EndStartMonth").Item) > 0) And (CInt(oRequest("EndStartDay").Item) > 0)) Then Call GetStartAndEndDatesFromURL("StartStart", "EndStart", "StartDate", True, aCatalogComponent(S_QUERY_CONDITION_CATALOG))
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (Positions.EndDate=30000000) And (Positions.Active=1) And (Branches.EndDate=30000000) And (Branches.Active=1)"
			Case "282"
				If Len(oRequest("KardexChangeTypeID").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesKardex4.KardexChangeTypeID In (" & oRequest("KardexChangeTypeID").Item & "))"
				ElseIf CInt(Request.Cookies("SIAP_SectionID")) <> 2 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesKardex4.KardexChangeTypeID In (0,1))"
				End If
				If Len(oRequest("EmployeeID").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeeID=" & oRequest("EmployeeID").Item & ")"
				End If
				If Len(oRequest("JobID").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (JobID=" & oRequest("JobID").Item & ")"
				End If
				If Len(oRequest("PositionID").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (Positions.PositionID In (" & oRequest("PositionID").Item & "))"
				End If
				If ((CInt(oRequest("StartStartYear").Item) > 0) And (CInt(oRequest("StartStartMonth").Item) > 0) And (CInt(oRequest("StartStartDay").Item) > 0)) Or ((CInt(oRequest("EndStartYear").Item) > 0) And (CInt(oRequest("EndStartMonth").Item) > 0) And (CInt(oRequest("EndStartDay").Item) > 0)) Then Call GetStartAndEndDatesFromURL("StartStart", "EndStart", "StartDate", True, aCatalogComponent(S_QUERY_CONDITION_CATALOG))
				If ((CInt(oRequest("StartEndYear").Item) > 0) And (CInt(oRequest("StartEndMonth").Item) > 0) And (CInt(oRequest("StartEndDay").Item) > 0)) Or ((CInt(oRequest("EndEndYear").Item) > 0) And (CInt(oRequest("EndEndMonth").Item) > 0) And (CInt(oRequest("EndEndDay").Item) > 0)) Then Call GetStartAndEndDatesFromURL("StartEnd", "EndEnd", "EndDate", True, aCatalogComponent(S_QUERY_CONDITION_CATALOG))
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (Positions.EndDate=30000000) And (Positions.Active=1) And (Journeys.EndDate=30000000) And (Journeys.Active=1) And (Areas.EndDate=30000000) And (Areas.ParentID<>-1) And (Services.EndDate=30000000) And (Services.Active=1) And (Branches.EndDate=30000000) And (Branches.Active=1)"
			Case "423", "424", "425", "426"
				If Len(oRequest("SectionID").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesSpecialJourneys.SpecialJourneyID In (" & oRequest("SectionID").Item & "))"
				End If
				If Len(oRequest("EmployeeIDs").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesSpecialJourneys.EmployeeID In (" & oRequest("EmployeeIDs").Item & "))"
				ElseIf Len(oRequest("EmployeeID").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesSpecialJourneys.EmployeeID In (" & oRequest("EmployeeID").Item & "))"
				End If
				If Len(oRequest("EmployeeName").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesSpecialJourneys.EmployeeName Like ('" & S_WILD_CHAR & oRequest("EmployeeName").Item & S_WILD_CHAR & "'))"
				End If
				If Len(oRequest("EmployeeLastName").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesSpecialJourneys.EmployeeLastName Like ('" & S_WILD_CHAR & oRequest("EmployeeLastName").Item & S_WILD_CHAR & "'))"
				End If
				If Len(oRequest("EmployeeLastName2").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesSpecialJourneys.EmployeeLastName2 Like ('" & S_WILD_CHAR & oRequest("EmployeeLastName2").Item & S_WILD_CHAR & "'))"
				End If
				If Len(oRequest("EmployeeRFC").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesSpecialJourneys.RFC Like ('" & S_WILD_CHAR & oRequest("EmployeeRFC").Item & S_WILD_CHAR & "'))"
				End If
				If Len(oRequest("OriginalEmployeeID").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesSpecialJourneys.OriginalEmployeeID In (" & oRequest("OriginalEmployeeID").Item & "))"
				End If
				If Len(oRequest("AreaID").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesSpecialJourneys.AreaID In (" & oRequest("AreaID").Item & "))"
				End If
				If Len(oRequest("JourneyID").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesSpecialJourneys.JourneyID In (" & oRequest("JourneyID").Item & "))"
				End If
				If Len(oRequest("MovementID").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesSpecialJourneys.MovementID In (" & oRequest("MovementID").Item & "))"
				End If
				If ((CInt(oRequest("StartStartYear").Item) > 0) And (CInt(oRequest("StartStartMonth").Item) > 0) And (CInt(oRequest("StartStartDay").Item) > 0)) Or ((CInt(oRequest("EndStartYear").Item) > 0) And (CInt(oRequest("EndStartMonth").Item) > 0) And (CInt(oRequest("EndStartDay").Item) > 0)) Then Call GetStartAndEndDatesFromURL("StartStart", "EndStart", "EmployeesSpecialJourneys.StartDate", True, aCatalogComponent(S_QUERY_CONDITION_CATALOG))
				If ((CInt(oRequest("StartEndYear").Item) > 0) And (CInt(oRequest("StartEndMonth").Item) > 0) And (CInt(oRequest("StartEndDay").Item) > 0)) Or ((CInt(oRequest("EndEndYear").Item) > 0) And (CInt(oRequest("EndEndMonth").Item) > 0) And (CInt(oRequest("EndEndDay").Item) > 0)) Then Call GetStartAndEndDatesFromURL("StartEnd", "EndEnd", "EmployeesSpecialJourneys.EndDate", True, aCatalogComponent(S_QUERY_CONDITION_CATALOG))
				If Len(oRequest("Internal").Item) > 0 Then
					If CInt(oRequest("Internal").Item) = 1 Then
						aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesSpecialJourneys.EmployeeID<800000)"
					ElseIf CInt(oRequest("Internal").Item) = 0 Then
						aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesSpecialJourneys.EmployeeID>=800000)"
					End If
				End If
			Case Else
				If Len(oRequest("EmployeeNumber").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeeNumber Like '" & S_WILD_CHAR & oRequest("EmployeeNumber").Item & S_WILD_CHAR & "')"
				End If
				If Len(oRequest("PersonNameName").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (PersonNameName Like '" & S_WILD_CHAR & oRequest("PersonNameName").Item & S_WILD_CHAR & "')"
				End If
				If Len(oRequest("PersonNameLastName").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (PersonNameLastName Like '" & S_WILD_CHAR & oRequest("PersonNameLastName").Item & S_WILD_CHAR & "')"
				End If
				If Len(oRequest("PersonNameLastName2").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (PersonNameLastName2 Like '" & S_WILD_CHAR & oRequest("PersonNameLastName2").Item & S_WILD_CHAR & "')"
				End If
				If Len(oRequest("EmployeeName").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeeName Like '" & S_WILD_CHAR & oRequest("EmployeeName").Item & S_WILD_CHAR & "')"
				End If
				If Len(oRequest("EmployeeLastName").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeeLastName Like '" & S_WILD_CHAR & oRequest("EmployeeLastName").Item & S_WILD_CHAR & "')"
				End If
				If Len(oRequest("EmployeeLastName2").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeeLastName2 Like '" & S_WILD_CHAR & oRequest("EmployeeLastName2").Item & S_WILD_CHAR & "')"
				End If
				If Len(oRequest("KardexTypeID").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesKardex3.KardexTypeID In (" & Replace(oRequest("KardexTypeID").Item, " ", "") & "))"
				End If
				If ((CInt(oRequest("StartStartYear").Item) > 0) And (CInt(oRequest("StartStartMonth").Item) > 0) And (CInt(oRequest("StartStartDay").Item) > 0)) Or ((CInt(oRequest("EndStartYear").Item) > 0) And (CInt(oRequest("EndStartMonth").Item) > 0) And (CInt(oRequest("EndStartDay").Item) > 0)) Then Call GetStartAndEndDatesFromURL("StartStart", "EndStart", "EmployeesKardex3.StartDate", True, aCatalogComponent(S_QUERY_CONDITION_CATALOG))
				If Len(oRequest("KardexOriginID").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesKardex3.KardexOriginID In (" & Replace(oRequest("KardexOriginID").Item, " ", "") & "))"
				End If
				If Len(oRequest("PersonName").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (PersonName Like '" & S_WILD_CHAR & oRequest("PersonName").Item & S_WILD_CHAR & "')"
				End If
				If Len(oRequest("PersonLastName").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (PersonLastName Like '" & S_WILD_CHAR & oRequest("PersonLastName").Item & S_WILD_CHAR & "')"
				End If
				If Len(oRequest("PersonLastName2").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (PersonLastName2 Like '" & S_WILD_CHAR & oRequest("PersonLastName2").Item & S_WILD_CHAR & "')"
				End If
				If Len(oRequest("RequirementsTypeID").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesKardex3.RequirementsTypeID In (" & Replace(oRequest("RequirementsTypeID").Item, " ", "") & "))"
				End If
				If Len(oRequest("Requirements").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (Requirements Like '" & S_WILD_CHAR & oRequest("Requirements").Item & S_WILD_CHAR & "')"
				End If
				If ((CInt(oRequest("StartDocumentsYear").Item) > 0) And (CInt(oRequest("StartDocumentsMonth").Item) > 0) And (CInt(oRequest("StartDocumentsDay").Item) > 0)) Or ((CInt(oRequest("EndDocumentsYear").Item) > 0) And (CInt(oRequest("EndDocumentsMonth").Item) > 0) And (CInt(oRequest("EndDocumentsDay").Item) > 0)) Then Call GetStartAndEndDatesFromURL("StartDocuments", "EndDocuments", "EmployeesKardex3.DocumentsDate", True, aCatalogComponent(S_QUERY_CONDITION_CATALOG))
				If ((CInt(oRequest("StartKnowledgeYear").Item) > 0) And (CInt(oRequest("StartKnowledgeMonth").Item) > 0) And (CInt(oRequest("StartKnowledgeDay").Item) > 0)) Or ((CInt(oRequest("EndKnowledgeYear").Item) > 0) And (CInt(oRequest("EndKnowledgeMonth").Item) > 0) And (CInt(oRequest("EndKnowledgeDay").Item) > 0)) Then Call GetStartAndEndDatesFromURL("StartKnowledge", "EndKnowledge", "EmployeesKardex3.KnowledgeDate", True, aCatalogComponent(S_QUERY_CONDITION_CATALOG))
				If Len(oRequest("KnowledgeStatusID").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesKardex3.KnowledgeStatusID In (" & Replace(oRequest("KnowledgeStatusID").Item, " ", "") & "))"
				End If
				If ((CInt(oRequest("StartPsychologicYear").Item) > 0) And (CInt(oRequest("StartPsychologicMonth").Item) > 0) And (CInt(oRequest("StartPsychologicDay").Item) > 0)) Or ((CInt(oRequest("EndPsychologicYear").Item) > 0) And (CInt(oRequest("EndPsychologicMonth").Item) > 0) And (CInt(oRequest("EndPsychologicDay").Item) > 0)) Then Call GetStartAndEndDatesFromURL("StartPsychologic", "EndPsychologic", "EmployeesKardex3.PsychologicDate", True, aCatalogComponent(S_QUERY_CONDITION_CATALOG))
				If Len(oRequest("PsychologicStatusID").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesKardex3.PsychologicStatusID In (" & Replace(oRequest("PsychologicStatusID").Item, " ", "") & "))"
				End If
				If ((CInt(oRequest("StartRegistration1Year").Item) > 0) And (CInt(oRequest("StartRegistration1Month").Item) > 0) And (CInt(oRequest("StartRegistration1Day").Item) > 0)) Or ((CInt(oRequest("EndRegistration1Year").Item) > 0) And (CInt(oRequest("EndRegistration1Month").Item) > 0) And (CInt(oRequest("EndRegistration1Day").Item) > 0)) Then Call GetStartAndEndDatesFromURL("StartRegistration1", "EndRegistration1", "EmployeesKardex3.Registration1Date", True, aCatalogComponent(S_QUERY_CONDITION_CATALOG))
				If ((CInt(oRequest("StartRegistration2Year").Item) > 0) And (CInt(oRequest("StartRegistration2Month").Item) > 0) And (CInt(oRequest("StartRegistration2Day").Item) > 0)) Or ((CInt(oRequest("EndRegistration2Year").Item) > 0) And (CInt(oRequest("EndRegistration2Month").Item) > 0) And (CInt(oRequest("EndRegistration2Day").Item) > 0)) Then Call GetStartAndEndDatesFromURL("StartRegistration2", "EndRegistration2", "EmployeesKardex3.Registration2Date", True, aCatalogComponent(S_QUERY_CONDITION_CATALOG))
				If ((CInt(oRequest("StartRegistration3Year").Item) > 0) And (CInt(oRequest("StartRegistration3Month").Item) > 0) And (CInt(oRequest("StartRegistration3Day").Item) > 0)) Or ((CInt(oRequest("EndRegistration3Year").Item) > 0) And (CInt(oRequest("EndRegistration3Month").Item) > 0) And (CInt(oRequest("EndRegistration3Day").Item) > 0)) Then Call GetStartAndEndDatesFromURL("StartRegistration3", "EndRegistration3", "EmployeesKardex3.Registration3Date", True, aCatalogComponent(S_QUERY_CONDITION_CATALOG))

				If Len(oRequest("PositionTypeID").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesKardex.PositionTypeID In (" & Replace(oRequest("PositionTypeID").Item, " ", "") & "))"
				End If
				If Len(oRequest("CompanyID").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (Employees.CompanyID In (" & Replace(oRequest("CompanyID").Item, " ", "") & "))"
				End If

				If InStr(1, ",351,352,", "," & oRequest("SectionID").Item & ",", vbBinaryCompare) > 0 Then
					If Len(oRequest("PositionID").Item) > 0 Then
						aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesKardex3.PositionID In (" & Replace(oRequest("PositionID").Item, " ", "") & "))"
					End If
					If Len(oRequest("AreaID").Item) > 0 Then
						aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeesKardex3.AreaID In (" & Replace(oRequest("AreaID").Item, " ", "") & "))"
					End If
					If (Len(oRequest("HasDate1").Item) > 0) Or (Len(oRequest("HasDate2").Item) > 0) Or (Len(oRequest("HasDate3").Item) > 0) Then
						aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And ("
							If Len(oRequest("HasDate1").Item) > 0 Then aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & "(Registration1Date>0) Or "
							If Len(oRequest("HasDate2").Item) > 0 Then aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & "(Registration2Date>0) Or "
							If Len(oRequest("HasDate3").Item) > 0 Then aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & "(Registration3Date>0) Or "
							aCatalogComponent(S_QUERY_CONDITION_CATALOG) = Left(aCatalogComponent(S_QUERY_CONDITION_CATALOG), (Len(aCatalogComponent(S_QUERY_CONDITION_CATALOG)) - Len(" Or ")))
						aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & ")"
					End If
				Else
					If Len(oRequest("PositionID").Item) > 0 Then
						aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (Positions.PositionID In (" & Replace(oRequest("PositionID").Item, " ", "") & "))"
					End If
					If Len(oRequest("SubAreaID").Item) > 0 Then
						aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (Areas.AreaID In (" & oRequest("SubAreaID").Item & "))"
					ElseIf Len(oRequest("AreaID").Item) > 0 Then
						aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (Areas.AreaPath Like '" & S_WILD_CHAR & "," & oRequest("AreaID").Item & "," & S_WILD_CHAR & "')"
					End If
				End If
				If Len(oRequest("EmployeeTypeID").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (Employees.EmployeeTypeID In (" & Replace(oRequest("EmployeeTypeID").Item, " ", "") & "))"
				End If
				If Len(oRequest("DocumentForLicenseNumber").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (DocumentForLicenseNumber Like '" & S_WILD_CHAR & oRequest("DocumentForLicenseNumber").Item & S_WILD_CHAR & "')"
				End If
				If Len(oRequest("DocumentForCancelLicenseNumber").Item) > 0 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (DocumentForCancelLicenseNumber Like '" & S_WILD_CHAR & oRequest("DocumentForCancelLicenseNumber").Item & S_WILD_CHAR & "')"
				End If
		End Select
	End If
	aCatalogComponent(N_ACTIVE_CATALOG) = -1
	aCatalogComponent(S_URL_CATALOG) = "SectionID=" & oRequest("SectionID").Item
	Select Case oRequest("SectionID").Item
		Case 261
			aCatalogComponent(S_URL_CATALOG) = aCatalogComponent(S_URL_CATALOG) & "&EmployeeID=<FIELD_0 />&AntiquityDate=<FIELD_1 />&PrevAntiquityDate=<FIELD_1 />"
	End Select

	DoCatalogsAction = lErrorNumber
	Err.Clear
End Function
%>