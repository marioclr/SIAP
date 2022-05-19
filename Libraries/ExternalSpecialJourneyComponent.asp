<%
Function AddExternalEmployee(oRequest, oADODBConnection, aSpecialJourneyComponent, sErrorDescription)
'************************************************************
'Purpose: Add a child of an employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddExternalEmployee"
	Dim lErrorNumber
	Dim oRecordset
	Dim iRedordID

	sErrorDescription = "No se pudo obtener un identificador para el nuevo registro."
	lErrorNumber = GetNewIDFromTable(oADODBConnection, "ExternalSpecialJourneys", "ExternalID", "", 1, iRedordID, sErrorDescription)
	aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EXTERNAL_ID) = iRedordID
	aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NUMBER) = CStr(iRedordID)
	If lErrorNumber = 0 Then
		sErrorDescription = "Error al insertar el nuevo registro de personal externo."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into ExternalSpecialJourneys (ExternalID, EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, SPEPID, Active) Values (" & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EXTERNAL_ID) &", '" & Replace(UCase(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NUMBER)), "'", "´") & "', '" & Replace(UCase(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NAME)), "'", "´") & "', '" & Replace(UCase(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME)), "'", "´") & "', '" &  Replace(UCase(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME2)), "'", "´") & "', '" &  Replace(UCase(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_RFC)), "'", "´") & "', '" &  Replace(UCase(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_CURP)), "'", "´") & "', '" &  Replace(UCase(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_SPEP)), "'", "´") & "', 0)", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	AddExternalEmployee = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfExternalID(aSpecialJourneyComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific employee exists in the database
'Inputs:  aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfExternalID"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	If aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) < 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del empleado en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From ExternalSpecialJourneys Where (ExternalID=" & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) & ")", "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				sErrorDescription = "El número de empleado no esta registrado en la base de datos."
				lErrorNumber = L_ERR_NO_RECORDS
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfExternalID = lErrorNumber
	Err.Clear
End Function

Function DisplayExternalJourneyForm(oRequest, oADODBConnection, sAction, aSpecialJourneyComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a concept from the
'		  database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, aSpecialJourneyComponent
'Outputs: aSpecialJourneyComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayExternalJourneyForm"
	Dim lErrorNumber

	If aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) <> -1 Then
		lErrorNumber = GetExternalEmployee(oRequest, oADODBConnection, aSpecialJourneyComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckConceptFields(oForm) {" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					Response.Write "return true;" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of CheckConceptFields" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "<FORM NAME=""ExternalFrm"" ID=""ExternalFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckConceptFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""ExternalSpecialJourney"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ExternalID"" ID=""ExternalIDHdn"" VALUE=""" & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EXTERNAL_ID) & """ />"
			'Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartDate"" ID=""StartDateHdn"" VALUE=""" & aSpecialJourneyComponent(N_START_DATE_CONCEPT) & """ />"

			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre(s):&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeName"" ID=""EmployeeNameTxt"" VALUE=""" & aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NAME) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido paterno:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName"" ID=""EmployeeLastNameTxt"" VALUE=""" & aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido materno:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName2"" ID=""EmployeeLastName2Txt"" VALUE=""" & aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME2) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">RFC:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""RFC"" ID=""RFCTxt"" VALUE=""" & aSpecialJourneyComponent(S_SPECIAL_JOURNEY_RFC) & """ SIZE=""13"" MAXLENGTH=""13"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">CURP:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""CURP"" ID=""CURPTxt"" VALUE=""" & aSpecialJourneyComponent(S_SPECIAL_JOURNEY_CURP) & """ SIZE=""18"" MAXLENGTH=""18"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Identificador SPEP:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""SPEP"" ID=""SPEPTxt"" VALUE=""" & aSpecialJourneyComponent(S_SPECIAL_JOURNEY_SPEP) & """ SIZE=""15"" MAXLENGTH=""15"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Folio de autorización:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""Folio"" ID=""FolioTxt"" VALUE=""" & aSpecialJourneyComponent(S_SPECIAL_JOURNEY_FOLIO) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE><BR />"
			If aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EXTERNAL_ID) = -1 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			ElseIf Len(oRequest("Delete").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveExternalWngDiv']); ExternalFrm.Remove.focus()"" />"
			ElseIf Len(oRequest("Change").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=" & oRequest("Action").Item & "&ConceptID=" & aSpecialJourneyComponent(N_ID_CONCEPT) & "&StartDate=" & aSpecialJourneyComponent(N_START_DATE_CONCEPT) & "'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveExternalWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
		Response.Write "</FORM>"
	End If

	DisplayExternalJourneyForm = lErrorNumber
	Err.Clear
End Function

Function VerifyRequerimentsForExternalSpecialJourneys(oADODBConnection, aSpecialJourneyComponent, sErrorDescription)
'************************************************************
'Purpose: To verify employee status requirements to register absences
'Inputs:  oADODBConnection, lReasonID, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyRequerimentsForExternalSpecialJourneys"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim bComponentInitialized

	bComponentInitialized = aSpecialJourneyComponent(B_COMPONENT_INITIALIZED_SPECIAL_JOURNEY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeSpecialJourneyComponent(oRequest, aSpecialJourneyComponent)
	End If

	If (aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) = -1) Then
		sErrorDescription = "No se especificó el identificador del empleado para agregar incidencias."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
		VerifyRequerimentsForExternalSpecialJourneys = False
	Else
		lErrorNumber = CheckExistencyOfExternalID(aSpecialJourneyComponent, sErrorDescription)
		If lErrorNumber = 0 Then
			VerifyRequerimentsForExternalSpecialJourneys = True
		Else
			VerifyRequerimentsForExternalSpecialJourneys = False
		End If
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function DisplayExternalEmployeeTable(oRequest, oADODBConnection, lIDColumn, bUseLinks, aSpecialJourneyComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the concepts
'		  from the database in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, aSpecialJourneyComponent
'Outputs: aSpecialJourneyComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayExternalEmployeeTable"
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
	Dim lErrorNumber
	Dim iStatusID
	Dim iRecordCounter
	Dim sCondition
	Dim iConceptsStatusID

	iStatusID = aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ACTIVE)
	sCondition = ""
	If iStatusID = 0 Then
		'sCondition = sCondition & " And (ExternalSpecialJourneys.Active<=" & iStatusID & ")"
	Else
		'sCondition = sCondition & " And (ExternalSpecialJourneys.Active=" & iStatusID & ")"
	End If
	aSpecialJourneyComponent(S_QUERY_CONDITION_SPECIAL_JOURNEY) = aSpecialJourneyComponent(S_QUERY_CONDITION_SPECIAL_JOURNEY) & sCondition
	lErrorNumber = GetExternalEmployees(oRequest, oADODBConnection, aSpecialJourneyComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If Not bForExport Then Call DisplayIncrementalFetch(oRequest, CInt(oRequest("StartPage").Item), ROWS_CATALOG, oRecordset)
			Response.Write "<TABLE WIDTH=""550"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bUseLinks Then
					asColumnsTitles = Split("&nbsp;,No. Empleado,Nombre,RFC, CURP,ID SPEP,Acciones", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,80,150,90,90,80,80,80,80", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("&nbsp;,No. Empleado,Nombre,RFC, CURP,ID SPEP", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,80,150,90,80,80,80", ",", -1, vbBinaryCompare)
				End If
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If

				iRecordCounter = 0
				asCellAlignments = Split(",,,CENTER,CENTER,CENTER", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sFontBegin = ""
					sFontEnd = ""
					sBoldBegin = ""
					sBoldEnd = ""
					If (StrComp(CStr(oRecordset.Fields("ExternalID").Value), oRequest("ExternalID").Item, vbBinaryCompare) = 0) Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					'If CInt(oRecordset.Fields("IsDeduction").Value) = 1 Then
					'	sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
					'	sFontEnd = "</FONT>"
					'End If
					sRowContents = ""
					Select Case lIDColumn
						Case DISPLAY_RADIO_BUTTONS
							sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""ExternalID"" ID=""ExternalIDRd"" VALUE=""" & CStr(oRecordset.Fields("ExternalID").Value) & """ />"
						Case DISPLAY_CHECKBOXES
							sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""ExternalID"" ID=""ExternalIDChk"" VALUE=""" & CStr(oRecordset.Fields("ExternalID").Value) & """ />"
						Case Else
							sRowContents = sRowContents & "&nbsp;"
					End Select
					sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
						sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ExternalID=" & CStr(oRecordset.Fields("ExternalID").Value) & """"
					sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & sBoldEnd & sFontEnd & "</A>"

					sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
						sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ExternalID=" & CStr(oRecordset.Fields("ExternalID").Value) & """"
					sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value)) & sBoldEnd & sFontEnd & "</A>"

					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("SPEPID").Value)) & sBoldEnd & sFontEnd
					If bUseLinks Then
						'sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
						If CInt(oRecordset.Fields("ConceptID").Value) > 0 Then
							iConceptsStatusID =  CInt(oRecordset.Fields("StatusID").Value)
							sRowContents = sRowContents & TABLE_SEPARATOR 
							If iConceptsStatusID <= 0 Then
								Select Case iConceptsStatusID
									Case 0
										sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;&nbsp;"
									Case -1
										sRowContents = sRowContents & "<IMG SRC=""Images/IcnExclamationSmall.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Al agregar este registro se detectaron registros posteriores que serán ajustados al aplicar este registro"" BORDER=""0"" />"
										sRowContents = sRowContents & "&nbsp;&nbsp;"
									Case -2
										sRowContents = sRowContents & "<IMG SRC=""Images/IcnInformation.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Al agregar este registro se detectaron registros dentro de los efectos de este que se ajustaran al aplicar este registro"" BORDER=""0"" />"
										sRowContents = sRowContents & "&nbsp;&nbsp;"
									Case -3
										sRowContents = sRowContents & "<IMG SRC=""Images/IcnExclamationSmall.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Al agregar este registro se detectaron registros que cubren todo el periodo de este, los cuales se se ajustaran al aplicar este registro"" BORDER=""0"" />"
										sRowContents = sRowContents & "&nbsp;&nbsp;"
								End Select
								sRowContents = sRowContents & "&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=Concepts&ConceptID=" & CInt(oRecordset.Fields("ConceptID").Value) & "&StartDate=" & CStr(oRecordset.Fields("StartDate").Value) & "&Delete=1"">"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;&nbsp;"
								sRowContents = sRowContents & "&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=Concepts&ConceptID=" & CInt(oRecordset.Fields("ConceptID").Value) & "&StartDate=" & CStr(oRecordset.Fields("StartDate").Value) & "&Apply=1"">"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;&nbsp;"
							Else
								If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
									sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Concepts&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&StartDate=" & CStr(oRecordset.Fields("StartDate").Value) & "&Tab=1&Change=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
								End If

								sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&StartDate=" & CStr(oRecordset.Fields("StartDate").Value) & """>"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnCurrency.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar Tabuladores"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
							End If

							If False And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
								sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Concepts&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&Tab=1&Delete=1"">"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;"
							End If
						End If
					Else
						If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
							sRowContents = sRowContents & TABLE_SEPARATOR
							sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=External&ExternalID=" & CStr(oRecordset.Fields("ExternalID").Value) & "&Tab=1&Change=1"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;"
						End If
						If ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) And (Not B_PORTAL) Then
							sRowContents = sRowContents & TABLE_SEPARATOR
							sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Users&ExternalID=" & CStr(oRecordset.Fields("ExternalID").Value) & "&Delete=1"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;"
						End If
					End If

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					oRecordset.MoveNext
					iRecordCounter = iRecordCounter + 1
					If (bUseLinks) And (iRecordCounter >= ROWS_CATALOG) Then Exit Do
					If Err.number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayExternalEmployeeTable = lErrorNumber
	Err.Clear
End Function

Function GetExternalEmployee(oRequest, oADODBConnection, aSpecialJourneyComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about an absence for the
'         employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aSpecialJourneyComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetExternalEmployee"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aSpecialJourneyComponent(B_COMPONENT_INITIALIZED_SPECIAL_JOURNEY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeSpecialJourneyComponent(oRequest, aSpecialJourneyComponent)
	End If

	If aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del colaborador externo para obtener la información del registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From ExternalSpecialJourneys Where (ExternalID=" & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EXTERNAL_ID) & ")", "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El registro especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NUMBER) = CStr(oRecordset.Fields("EmployeeNumber").Value)
				aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NAME) = CStr(oRecordset.Fields("EmployeeName").Value)
				aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME) = CStr(oRecordset.Fields("EmployeeLastName").Value)
				aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME2) = CStr(oRecordset.Fields("EmployeeLastName2").Value)
				aSpecialJourneyComponent(S_SPECIAL_JOURNEY_RFC) = CStr(oRecordset.Fields("RFC").Value)
				aSpecialJourneyComponent(S_SPECIAL_JOURNEY_CURP) = CStr(oRecordset.Fields("CURP").Value)
				aSpecialJourneyComponent(S_SPECIAL_JOURNEY_SPEP) = CStr(oRecordset.Fields("SPEPID").Value)
				aSpecialJourneyComponent(B_SPECIAL_JOURNEY_EXIST_EXTERNAL) = True
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetExternalEmployee = lErrorNumber
	Err.Clear
End Function

Function GetExternalEmployees(oRequest, oADODBConnection, aSpecialJourneyComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the concepts from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aSpecialJourneyComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetExternalEmployees"
	Dim sCondition
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aSpecialJourneyComponent(B_COMPONENT_INITIALIZED_SPECIAL_JOURNEY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeSpecialJourneyComponent(oRequest, aSpecialJourneyComponent)
	End If

	If (Len(aSpecialJourneyComponent(S_QUERY_CONDITION_SPECIAL_JOURNEY)) > 0) Then
		sCondition = Trim(aSpecialJourneyComponent(S_QUERY_CONDITION_SPECIAL_JOURNEY))
		If InStr(1, sCondition, "And ", vbTextCompare) <> 1 Then
			sCondition = "And " & sCondition
		End If
	End If
	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From ExternalSpecialJourneys Where (ExternalID>800000) " & sCondition & " Order By ExternalID", "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	GetExternalEmployees = lErrorNumber
	Err.Clear
End Function

Function ModifyExternalEmployee(oRequest, oADODBConnection, aSpecialJourneyComponent, sErrorDescription)
'************************************************************
'Purpose: To modify a child from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyExternalEmployee"
	Dim lErrorNumber
	Dim oRecordset

	sErrorDescription = "Error al modificar la información de la base de datos"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ExternalSpecialJourneys Set EmployeeName='" & Replace(UCase(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NAME)), "'", "") & "', EmployeeLastName='" & Replace(UCase(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME)), "'", "") & "', EmployeeLastName2='" & Replace(UCase(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME2)), "'", "") & "', RFC='" & Replace(UCase(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_RFC)), "'", "") & "', CURP='" & Replace(UCase(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_CURP)), "'", "") & "', SPEPID='" & Replace(UCase(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_SPEP)), "'", "") & "', Active=" & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ACTIVE) & " Where (ExternalID=" & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EXTERNAL_ID) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

	ModifyExternalEmployee = lErrorNumber
	Err.Clear
End Function

Function RemoveExternalEmployee(oRequest, oADODBConnection, aSpecialJourneyComponent, sErrorDescription)
'************************************************************
'Purpose: To remove an absence for the employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aSpecialJourneyComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveExternalEmployee"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aSpecialJourneyComponent(B_COMPONENT_INITIALIZED_SPECIAL_JOURNEY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeSpecialJourneyComponent(oRequest, aSpecialJourneyComponent)
	End If

	If (aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EXTERNAL_ID) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador para eliminar la información del registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ExternalSpecialJourneyComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From ExternalSpecialJourneys Where (ExternalID=" & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EXTERNAL_ID) & ")", "ExternalSpecialJourneyComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	RemoveExternalEmployee = lErrorNumber
	Err.Clear
End Function
%>