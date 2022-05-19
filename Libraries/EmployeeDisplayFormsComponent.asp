<%
Dim sRequirements
Dim asRequirements(300)
Dim sReadOnly
Dim bActivate
Dim bVisible
Dim bReadOnly
Dim sDisplayFormCaseOptions
Dim sEmployeeDisplayFormAntiquity
Dim lDisplayFormAntiquityYears
Dim lDisplayFormAntiquityMonths
Dim lDisplayFormAntiquityDays
Dim lDisplayFormCurrentDate

Function DisplayEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about an employee from the
'         database
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployee"
	Dim sNames
	Dim oRecordset
	Dim lErrorNumber

	If aEmployeeComponent(N_ID_EMPLOYEE) <> -1 Then
		lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><FONT FACE=""Arial"" SIZE=""2"">"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Número de empleado:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aEmployeeComponent(S_NUMBER_EMPLOYEE)) & "</FONT></TD>"
				Response.Write "</TR>"

				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Clave de acceso:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aEmployeeComponent(S_ACCESS_KEY_EMPLOYEE)) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Nombre del empleado:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aEmployeeComponent(S_NAME_EMPLOYEE)) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Apellido paterno:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aEmployeeComponent(S_LAST_NAME_EMPLOYEE)) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Apellido materno:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aEmployeeComponent(S_LAST_NAME2_EMPLOYEE)) & "</FONT></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE>"

			Response.Write "<BR /><HR /><BR />"

			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				If aEmployeeComponent(N_COMPANY_ID_EMPLOYEE) > -1 Then
					Response.Write "<TR>"
						Call GetNameFromTable(oADODBConnection, "Companies", aEmployeeComponent(N_COMPANY_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Empresa:&nbsp;</B></FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
					Response.Write "</TR>"
				End If
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Número de plaza:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						If aEmployeeComponent(N_JOB_ID_EMPLOYEE) > -1 Then
							Response.Write CleanStringForHTML(aEmployeeComponent(N_JOB_ID_EMPLOYEE))
						Else
							Response.Write "<B>NO TIENE PLAZA ASIGNADA</B>"
						End If
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Call GetNameFromTable(oADODBConnection, "Services", aEmployeeComponent(N_SERVICE_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Servicio:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Call GetNameFromTable(oADODBConnection, "EmployeeTypes", aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Tipo de tabulador:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Call GetNameFromTable(oADODBConnection, "PositionTypes", aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Tipo de puesto:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				If aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE) > -1 Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Clasificación:&nbsp;</B></FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE) & "</FONT></TD>"
					Response.Write "</TR>"
				End If
				Response.Write "<TR>"
					Call GetNameFromTable(oADODBConnection, "GroupGradeLevels", aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Grupo, grado, nivel:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Integración:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE)) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Call GetNameFromTable(oADODBConnection, "Levels", aEmployeeComponent(N_LEVEL_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Nivel:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Call GetNameFromTable(oADODBConnection, "PaymentCenters", aEmployeeComponent(N_PAYMENT_CENTER_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Centro de pago:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE>"

			Response.Write "<BR /><HR /><BR />"

			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Call GetNameFromTable(oADODBConnection, "Journeys", aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Jornada:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Call GetNameFromTable(oADODBConnection, "Shifts", aEmployeeComponent(N_SHIFT_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Turno:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				If (aEmployeeComponent(N_START_HOUR_1_EMPLOYEE) > 0) And (aEmployeeComponent(N_END_HOUR_1_EMPLOYEE) > 0) And (aEmployeeComponent(N_START_HOUR_2_EMPLOYEE) > 0) And (aEmployeeComponent(N_END_HOUR_2_EMPLOYEE) > 0) And (aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) > 0) And (aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) > 0) Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Horario:&nbsp;</B></FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
							If (aEmployeeComponent(N_START_HOUR_1_EMPLOYEE) > 0) And (aEmployeeComponent(N_END_HOUR_1_EMPLOYEE) > 0) Then DisplayTimeFromSerialNumber (aEmployeeComponent(N_START_HOUR_1_EMPLOYEE)) & " a " & DisplayTimeFromSerialNumber(aEmployeeComponent(N_END_HOUR_1_EMPLOYEE))
							If (aEmployeeComponent(N_START_HOUR_2_EMPLOYEE) > 0) And (aEmployeeComponent(N_END_HOUR_2_EMPLOYEE) > 0) Then Response.Write " y de " & DisplayTimeFromSerialNumber(aEmployeeComponent(N_START_HOUR_2_EMPLOYEE)) & " a " & DisplayTimeFromSerialNumber(aEmployeeComponent(N_END_HOUR_2_EMPLOYEE))
							If (aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) > 0) And (aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) > 0) Then Response.Write "<BR /><B>Turno opcional:&nbsp;</B>" & DisplayTimeFromSerialNumber(aEmployeeComponent(N_START_HOUR_2_EMPLOYEE)) & " a " & DisplayTimeFromSerialNumber(aEmployeeComponent(N_END_HOUR_2_EMPLOYEE))
						Response.Write "</FONT></TD>"
					Response.Write "</TR>"
				End If
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Horas laboradas:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						If aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) > -1 Then
							Response.Write aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE)
						Else
							Response.Write "N/A"
						End If
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE>"

			Response.Write "<BR /><HR /><BR />"

			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				If Len(aEmployeeComponent(S_EMAIL_EMPLOYEE)) > 0 Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Correo electrónico:&nbsp;</B></FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><A HREF=""mailto: " & aEmployeeComponent(S_EMAIL_EMPLOYEE) & """>" & CleanStringForHTML(aEmployeeComponent(S_EMAIL_EMPLOYEE)) & "</A></FONT></TD>"
					Response.Write "</TR>"
				End If
				If Len(aEmployeeComponent(S_SSN_EMPLOYEE)) > 0 Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Número de seguro social:&nbsp;</B></FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aEmployeeComponent(S_SSN_EMPLOYEE)) & "</FONT></TD>"
					Response.Write "</TR>"
				End If
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Fecha de nacimiento:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE), -1, -1, -1) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Fecha de inicio en el ISSSTE:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aEmployeeComponent(N_START_DATE_EMPLOYEE), -1, -1, -1) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Fecha de inicio en gobierno:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aEmployeeComponent(N_START_DATE2_EMPLOYEE), -1, -1, -1) & "</FONT></TD>"
				Response.Write "</TR>"
				If Not B_ISSSTE Then
					Response.Write "<TR>"
						Call GetNameFromTable(oADODBConnection, "Countries", aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>País de origen:&nbsp;</B></FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
					Response.Write "</TR>"
				End If
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>RFC:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aEmployeeComponent(S_RFC_EMPLOYEE)) & "</FONT></TD>"
				Response.Write "</TR>"
				If Len(aEmployeeComponent(S_CURP_EMPLOYEE)) > 0 Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>CURP:&nbsp;</B></FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aEmployeeComponent(S_CURP_EMPLOYEE)) & "</FONT></TD>"
					Response.Write "</TR>"
				End If
				Response.Write "<TR>"
					Call GetNameFromTable(oADODBConnection, "Genders", aEmployeeComponent(N_GENDER_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Género:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Call GetNameFromTable(oADODBConnection, "MaritalStatus", aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Estado civil:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Call GetNameFromTable(oADODBConnection, "StatusEmployees", aEmployeeComponent(N_STATUS_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Estatus:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Activo:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayYesNo(aEmployeeComponent(N_ACTIVE_EMPLOYEE), True) & "</FONT></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE>"

			Response.Write "<BR /><HR /><BR />"

			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR><TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2"">"
					Response.Write "<B>Dirección:<BR /></B>"
					Response.Write CleanStringForHTML(aEmployeeComponent(S_ADDRESS_EMPLOYEE))
				Response.Write "</FONT></TD></TR>"
				If Len(aEmployeeComponent(S_CITY_EMPLOYEE)) > 0 Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Ciudad:&nbsp;</B></FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aEmployeeComponent(S_CITY_EMPLOYEE)) & "</FONT></TD>"
					Response.Write "</TR>"
				End If
				If Len(aEmployeeComponent(S_ZIP_CODE_EMPLOYEE)) > 0 Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Código Postal:&nbsp;</B></FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aEmployeeComponent(S_ZIP_CODE_EMPLOYEE)) & "</FONT></TD>"
					Response.Write "</TR>"
				End If
				Response.Write "<TR>"
					Call GetNameFromTable(oADODBConnection, "States", aEmployeeComponent(N_ADDRESS_STATE_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Estado:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				If Not B_ISSSTE Then
					Response.Write "<TR>"
						Call GetNameFromTable(oADODBConnection, "Countries", aEmployeeComponent(N_ADDRESS_COUNTRY_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>País:&nbsp;</B></FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
					Response.Write "</TR>"
				End If
			Response.Write "</TABLE>"
		Response.Write "</FONT></DIV>"
	End If

	DisplayEmployee = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeeExport(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about an employee from the
'         database
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeExport"
	Dim sNames
	Dim oRecordset
	Dim lErrorNumber
	Dim sPictureSCR
	Dim strFoto
	Dim strRutaFoto
	Dim sSchoolarShip

	'Elementos de Credencialización
	Dim strFirmaSCR
	Dim strIdentificaSCR
	Dim strHuellaD1SCR
	Dim strHuellaD2SCR
	Dim strHuellaD3SCR
	Dim strHuellaD4SCR
	Dim strHuellaD5SCR
	Dim strHuellaI1SCR
	Dim strHuellaI2SCR
	Dim strHuellaI3SCR
	Dim strHuellaI4SCR
	Dim strHuellaI5SCR

	Dim strRutaFirma
	Dim strRutaIdentifica
	Dim strRutaHuellaD1
	Dim strRutaHuellaD2
	Dim strRutaHuellaD3
	Dim strRutaHuellaD4
	Dim strRutaHuellaD5
	Dim strRutaHuellaI1
	Dim strRutaHuellaI2
	Dim strRutaHuellaI3
	Dim strRutaHuellaI4
	Dim strRutaHuellaI5

	If aEmployeeComponent(N_ID_EMPLOYEE) <> -1 Then
		lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
	End If

	strRutaIdentifica = CStr(Server.MapPath("Uploaded Files\e" & aEmployeeComponent(N_ID_EMPLOYEE) & "\" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_IDENTIFICACION.jpg"))
	strRutaHuellaI1 = CStr(Server.MapPath("Uploaded Files\e" & aEmployeeComponent(N_ID_EMPLOYEE) & "\" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_HUELLA_IZQUIERDA1.jpg"))
	strRutaHuellaI2 = CStr(Server.MapPath("Uploaded Files\e" & aEmployeeComponent(N_ID_EMPLOYEE) & "\" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_HUELLA_IZQUIERDA2.jpg"))
	strRutaHuellaI3 = CStr(Server.MapPath("Uploaded Files\e" & aEmployeeComponent(N_ID_EMPLOYEE) & "\" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_HUELLA_IZQUIERDA3.jpg"))
	strRutaHuellaI4 = CStr(Server.MapPath("Uploaded Files\e" & aEmployeeComponent(N_ID_EMPLOYEE) & "\" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_HUELLA_IZQUIERDA4.jpg"))
	strRutaHuellaI5 = CStr(Server.MapPath("Uploaded Files\e" & aEmployeeComponent(N_ID_EMPLOYEE) & "\" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_HUELLA_IZQUIERDA5.jpg"))
	strRutaHuellaD1 = CStr(Server.MapPath("Uploaded Files\e" & aEmployeeComponent(N_ID_EMPLOYEE) & "\" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_HUELLA_DERECHA_1.jpg"))
	strRutaHuellaD2 = CStr(Server.MapPath("Uploaded Files\e" & aEmployeeComponent(N_ID_EMPLOYEE) & "\" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_HUELLA_DERECHA_2.jpg"))
	strRutaHuellaD3 = CStr(Server.MapPath("Uploaded Files\e" & aEmployeeComponent(N_ID_EMPLOYEE) & "\" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_HUELLA_DERECHA_3.jpg"))
	strRutaHuellaD4 = CStr(Server.MapPath("Uploaded Files\e" & aEmployeeComponent(N_ID_EMPLOYEE) & "\" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_HUELLA_DERECHA_4.jpg"))
	strRutaHuellaD5 = CStr(Server.MapPath("Uploaded Files\e" & aEmployeeComponent(N_ID_EMPLOYEE) & "\" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_HUELLA_DERECHA_5.jpg"))
	strRutaFirma = CStr(Server.MapPath("Uploaded Files\e" & aEmployeeComponent(N_ID_EMPLOYEE) & "\" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_FIRMA.jpg"))
	strRutaFoto = CStr(Server.MapPath("Uploaded Files\e" & aEmployeeComponent(N_ID_EMPLOYEE) & "\" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_FOTO.jpg"))

	sPictureSCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Uploaded Files/e" & aEmployeeComponent(N_ID_EMPLOYEE) & "/" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & ".jpg"
	strFirmaSCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Uploaded Files/e" & aEmployeeComponent(N_ID_EMPLOYEE) & "/" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_FIRMA.jpg"
	strIdentificaSCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Uploaded Files/e" & aEmployeeComponent(N_ID_EMPLOYEE) & "/" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_IDENTIFICACION.jpg"
	strHuellaD1SCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Uploaded Files/e" & aEmployeeComponent(N_ID_EMPLOYEE) & "/" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_HUELLA_DERECHA_1.jpg"
	strHuellaD2SCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Uploaded Files/e" & aEmployeeComponent(N_ID_EMPLOYEE) & "/" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_HUELLA_DERECHA_2.jpg"
	strHuellaD3SCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Uploaded Files/e" & aEmployeeComponent(N_ID_EMPLOYEE) & "/" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_HUELLA_DERECHA_3.jpg"
	strHuellaD4SCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Uploaded Files/e" & aEmployeeComponent(N_ID_EMPLOYEE) & "/" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_HUELLA_DERECHA_4.jpg"
	strHuellaD5SCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Uploaded Files/e" & aEmployeeComponent(N_ID_EMPLOYEE) & "/" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_HUELLA_DERECHA_5.jpg"
	strHuellaI1SCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Uploaded Files/e" & aEmployeeComponent(N_ID_EMPLOYEE) & "/" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_HUELLA_IZQUIERDA1.jpg"
	strHuellaI2SCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Uploaded Files/e" & aEmployeeComponent(N_ID_EMPLOYEE) & "/" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_HUELLA_IZQUIERDA2.jpg"
	strHuellaI3SCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Uploaded Files/e" & aEmployeeComponent(N_ID_EMPLOYEE) & "/" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_HUELLA_IZQUIERDA3.jpg"
	strHuellaI4SCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Uploaded Files/e" & aEmployeeComponent(N_ID_EMPLOYEE) & "/" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_HUELLA_IZQUIERDA4.jpg"
	strHuellaI5SCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Uploaded Files/e" & aEmployeeComponent(N_ID_EMPLOYEE) & "/" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "_HUELLA_IZQUIERDA5.jpg"

	If Not FileExists(strRutaFoto, "") Then
		sPictureSCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Images/nofoto.jpg"
	End If
	If Not FileExists(strRutaFirma, "") Then
		strFirmaSCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Images/nofoto.jpg"
	End If
	If Not FileExists(strRutaIdentifica, "") Then
		strIdentificaSCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Images/nofoto.jpg"
	End If
	If Not FileExists(strRutaHuellaD1, "") Then
		strHuellaD1SCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Images/nofoto.jpg"
	End If
	If Not FileExists(strRutaHuellaD2, "") Then
		strHuellaD2SCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Images/nofoto.jpg"
	End If
	If Not FileExists(strRutaHuellaD3, "") Then
		strHuellaD3SCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Images/nofoto.jpg"
	End If
	If Not FileExists(strRutaHuellaD4, "") Then
		strHuellaD4SCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Images/nofoto.jpg"
	End If
	If Not FileExists(strRutaHuellaD5, "") Then
		strHuellaD5SCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Images/nofoto.jpg"
	End If
	If Not FileExists(strRutaHuellaI1, "") Then
		strHuellaI1SCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Images/nofoto.jpg"
	End If
	If Not FileExists(strRutaHuellaI2, "") Then
		strHuellaI2SCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Images/nofoto.jpg"
	End If
	If Not FileExists(strRutaHuellaI3, "") Then
		strHuellaI3SCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Images/nofoto.jpg"
	End If
	If Not FileExists(strRutaHuellaI4, "") Then
		strHuellaI4SCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Images/nofoto.jpg"
	End If
	If Not FileExists(strRutaHuellaI5, "") Then
		strHuellaI5SCR = "http://"& Request.ServerVariables("SERVER_NAME") & "/SIAP/Images/nofoto.jpg"
	End If

	If lErrorNumber = 0 Then
		Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><FONT FACE=""Arial"" SIZE=""2"">"
			Response.Write "<TABLE CELLSPACING=""3"" CELLPADDING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD COLSPAN=2><IMG SRC=""http://"& Request.ServerVariables("SERVER_NAME")&"/SIAP/Images/LogoISSSTE.gif"" WIDTH=""200"" HEIGHT=""90"" ALT=""LogoISSSTE"" BORDER=""0"" /></TD>"
					Response.Write "<TD COLSPAN=3><h1><FONT COLOR=""#C0C0C0"">Inventario de Recursos Humanos</FONT></h1></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD COLSPAN=2></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD COLSPAN=5><FONT FACE=""Arial"" SIZE=""1""><p ALIGN=right><b>NÚMERO DE EMPLEADO</b>: "
						Response.Write CleanStringForHTML(aEmployeeComponent(S_NUMBER_EMPLOYEE)) & "</FONT>"
					Response.Write "</p></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE>"

			Response.Write "<TABLE>"
				Response.Write "<TR>"
					Response.Write "<TD>"
						Response.Write "<TABLE>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><b></b></FONT></TD>"
								Response.Write "<TD COLSPAN = 2><FONT FACE=""Arial"" SIZE=""2""><b>DATOS PERSONALES</b></FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD ROWSPAN=10><IMG SRC=""" & sPictureSCR & """ WIDTH=""120"" HEIGHT=""150"" ALT=""Picture"" BORDER=""0"" /></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD width = 30><FONT FACE=""Arial"" SIZE=""1""><b>NOMBRE COMPLETO</b></FONT></TD>"
								Response.Write "<TD COLSPAN=3><FONT FACE=""Arial"" SIZE=""1"">"& CleanStringForHTML(aEmployeeComponent(S_NAME_EMPLOYEE)) &" "& CleanStringForHTML(aEmployeeComponent(S_LAST_NAME_EMPLOYEE)) &" "& CleanStringForHTML(aEmployeeComponent(S_LAST_NAME2_EMPLOYEE)) &"</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><B>FECHA DE NACIMIENTO</B></FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">" & DisplayDateFromSerialNumber(aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE), -1, -1, -1) & "</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><B>LUGAR DE NACIMIENTO</B></FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">" & CleanStringForHTML(aEmployeeComponent(S_EMPLOYEE_BIRTHPLACE)) & "</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Call GetNameFromTable(oADODBConnection, "Genders", aEmployeeComponent(N_GENDER_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><b>GÉNERO</b></FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><b>RFC</b></FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">" & CleanStringForHTML(aEmployeeComponent(S_RFC_EMPLOYEE)) & "</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><b>CURP</b></FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">" & CleanStringForHTML(aEmployeeComponent(S_CURP_EMPLOYEE)) & "</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><b>IFE</b></FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">"& CleanStringForHTML(aEmployeeComponent(S_DOCUMENT_NUMBER_1_EMPLOYEE)) &"</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Call GetNameFromTable(oADODBConnection, "Nationalities", aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><b>NACIONALIDAD</b></FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
							Response.Write "</TR>" 
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><b>No DE SEGURIDAD SOCIAL</b></FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">"& CleanStringForHTML(aEmployeeComponent(S_SSN_EMPLOYEE)) &"</FONT></TD>"
								'Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">" & CleanStringForHTML((aEmployeeComponent(S_SSN_EMPLOYEE)) & "</FONT></TD>"
							Response.Write "</TR>"
						Response.Write "</TABLE>"
					Response.Write "</TD>"

					Response.Write "<TD>"
						Response.Write "<TABLE>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><b>FIRMA:</b></FONT></TD>"
									Response.Write "<TD>"
										Response.Write "<IMG SRC=""" & strFirmaSCR & """ WIDTH=""200"" HEIGHT=""150"" ALT=""FIRMA"" BORDER=""0"" />"
									Response.Write "</TD>"
							Response.Write "</TR>"
						Response.Write "</TABLE>"
					Response.Write "</TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE><HR />"

			Response.Write "<TABLE>"
				Response.Write "<TR>"
					Response.Write "<TD>"
						Response.Write "<TABLE>"
							Response.Write "<TR>"
								Call GetNameFromTable(oADODBConnection, "MaritalStatus", aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><b>ESTADO CIVIL</b></FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><b>GRUPO SANGUÍNEO</b></FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">"& CleanStringForHTML(aEmployeeComponent(S_EMPLOYEE_BLOODTYPE)) &"</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><b>CORREO ELECTRÓNICO</b></FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><a href=""mailto: " & aEmployeeComponent(S_EMAIL_EMPLOYEE) & """>" & CleanStringForHTML(aEmployeeComponent(S_EMAIL_EMPLOYEE)) & "</a></FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								'Call GetNameFromTable(oADODBConnection, "ScoolarShips",sSchoolarShip, "", "", sNames, sErrorDescription)
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><b>ESCOLARIDAD</b></FONT></TD>"
								Call GetNameFromTable(oADODBConnection, "Schoolarships", aEmployeeComponent(N_EMPLOYEE_SCHOOLARSHIP_ID), "", "", sNames, sErrorDescription)
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><b>ESPECIALIDAD</b></FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">"& CleanStringForHTML(aEmployeeComponent(S_EMPLOYEE_SPECIALISM)) &"</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><b>IDIOMAS</b></FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">"& CleanStringForHTML(aEmployeeComponent(S_EMPLOYEE_LANGUAGES)) &"</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><b>CARTILLA</b></FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">" & CleanStringForHTML(aEmployeeComponent(S_DOCUMENT_NUMBER_3_EMPLOYEE)) & "</FONT></TD>"
							Response.Write "</TR>"
						Response.Write "</TABLE>"
					Response.Write "</TD>"
						Response.Write "<TD>"
							Response.Write "<TABLE>"
								Response.Write "<TR>"
									Response.Write "<TD>"
										Response.Write "<IMG SRC=""" & strHuellaD1SCR & """ WIDTH=""100"" HEIGHT=""120"" ALT=""HUELLA_DERECHA_1"" BORDER=""0"" />"
										Response.Write "<IMG SRC=""" & strHuellaD2SCR & """ WIDTH=""100"" HEIGHT=""120"" ALT=""HUELLA_DERECHA_2"" BORDER=""0"" />"
										Response.Write "<IMG SRC=""" & strHuellaD3SCR & """ WIDTH=""100"" HEIGHT=""120"" ALT=""HUELLA_DERECHA_3"" BORDER=""0"" />"
										Response.Write "<IMG SRC=""" & strHuellaD4SCR & """ WIDTH=""100"" HEIGHT=""120"" ALT=""HUELLA_DERECHA_4"" BORDER=""0"" />"
										Response.Write "<IMG SRC=""" & strHuellaD5SCR & """ WIDTH=""100"" HEIGHT=""120"" ALT=""HUELLA_DERECHA_5"" BORDER=""0"" />"
									Response.Write "</TD>"
								Response.Write "</TR>"
								Response.Write "<TR>"
									Response.Write "<TD>"
										Response.Write "<IMG SRC=""" & strHuellaI1SCR & """ WIDTH=""100"" HEIGHT=""120"" ALT=""HUELLA_IZQUIERDA1"" BORDER=""0"" />"
										Response.Write "<IMG SRC=""" & strHuellaI2SCR & """ WIDTH=""100"" HEIGHT=""120"" ALT=""HUELLA_IZQUIERDA2"" BORDER=""0"" />"
										Response.Write "<IMG SRC=""" & strHuellaI3SCR & """ WIDTH=""100"" HEIGHT=""120"" ALT=""HUELLA_IZQUIERDA3"" BORDER=""0"" />"
										Response.Write "<IMG SRC=""" & strHuellaI4SCR & """ WIDTH=""100"" HEIGHT=""120"" ALT=""HUELLA_IZQUIERDA4"" BORDER=""0"" />"
										Response.Write "<IMG SRC=""" & strHuellaI5SCR & """ WIDTH=""100"" HEIGHT=""120"" ALT=""HUELLA_IZQUIERDA5"" BORDER=""0"" />"
									Response.Write "</TD>"
								Response.Write "</TR>"
							Response.Write "</TABLE>"
						Response.Write "</TD>"
					Response.Write "</TR>"
				Response.Write "</TABLE><HR />"
			'If Len(aEmployeeComponent(S_ADDRESS_EMPLOYEE))>0 Then
				Response.Write "<TABLE>"
					Response.Write "<TR>"
						Response.Write "<TD>"
							Response.Write "<TABLE>"
								Response.Write "<TR>"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><B>DIRECCIÓN</B></FONT></TD>"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">" & CleanStringForHTML(aEmployeeComponent(S_ADDRESS_EMPLOYEE)) & "</FONT></TD>"
								Response.Write "</TR>"
								Response.Write "<TR>"
									Call GetNameFromTable(oADODBConnection, "States", aEmployeeComponent(N_ADDRESS_STATE_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><B>ESTADO</B></FONT></TD>"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
								Response.Write "</TR>"
								'If Len(aEmployeeComponent(S_CITY_EMPLOYEE)) > 0 Then
									Response.Write "<TR>"
										Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><B>DELEGACIÓN O MUNICIPIO</B></FONT></TD>"
										Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">" & CleanStringForHTML(aEmployeeComponent(S_CITY_EMPLOYEE)) & "</FONT></TD>"
									Response.Write "</TR>"
								'End If
								'If Len(aEmployeeComponent(S_ZIP_CODE_EMPLOYEE)) > 0 Then
									Response.Write "<TR>"
										Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><B>CÓDIGO POSTAL</B></FONT></TD>"
										Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">" & CleanStringForHTML(aEmployeeComponent(S_ZIP_CODE_EMPLOYEE)) & "</FONT></TD>"
									Response.Write "</TR>"
								'End If
								Response.Write "<TR>"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><b>TELÉFONO CASA </b></FONT></TD>"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">" & CleanStringForHTML(aEmployeeComponent(S_EMPLOYEE_PHONE_EMPLOYEE)) & "</FONT></TD>"
								Response.Write "</TR>"
								Response.Write "<TR>"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><b>TELÉFONO CELULAR </b></FONT></TD>"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">" & CleanStringForHTML(aEmployeeComponent(S_EMPLOYEE_CELLPHONE)) & "</FONT></TD>"
								Response.Write "</TR>"
								Response.Write "<TR>"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><b>TELÉFONO OFICINA</b></FONT></TD>"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">" & CleanStringForHTML(aEmployeeComponent(S_OFFICE_PHONE_EMPLOYEE)) & "</FONT></TD>"
								Response.Write "</TR>"
								Response.Write "<TR>"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><b>EXTENSIÓN</b></FONT></TD>"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">" & CleanStringForHTML(aEmployeeComponent(S_EXT_OFFICE_EMPLOYEE)) & "</FONT></TD>"
								Response.Write "</TR>"
							Response.Write "</TABLE>"
						Response.Write "</TD>"
						Response.Write "<TD rowspan=""8"">"
							Response.Write "<TABLE>"
								Response.Write "<TR>"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><b>IFE:</b></FONT></TD>"
										Response.Write "<TD>"
											Response.Write "<IMG SRC=""" & strIdentificaSCR & """ WIDTH=""500"" HEIGHT=""350"" ALT=""IFE"" BORDER=""0"" />"
										Response.Write "</TD>"
								Response.Write "</TR>"
							Response.Write "</TABLE>"
						Response.Write "</TD>"
					Response.Write "</TR>"
				Response.Write "</TABLE><HR />"

				Response.Write "<HR />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""1""><B>BENEFICIARIOS DE PAGO DE DEFUNCIÓN</B></FONT><BR />"
				Response.Write "<BR />"
				Response.Write "<TABLE>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><b>NOMBRE DEL BENEFICIARIO 1:&nbsp;</b></FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">" & CleanStringForHTML(aEmployeeComponent(S_EMPLOYEE_DEATH_BENEFICIARY)) & "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1""><b>NOMBRE DEL BENEFICIARIO 2:&nbsp;</b></FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">" & CleanStringForHTML(aEmployeeComponent(S_EMPLOYEE_DEATH_BENEFICIARY2)) & "</FONT></TD>"
					Response.Write "</TR>" 
				Response.Write "</TABLE>"
				Response.Write "<HR />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""1""><B>HIJOS</B></FONT><BR />"
				Response.Write "<BR />"
				Response.Write "<TABLE>"
					Response.Write"<TR COLSPAN=""4"">"
						lErrorNumber = DisplayEmployeeChildrenTable(oRequest, oADODBConnection, "Employees", DISPLAY_NOTHING, True, True, aEmployeeComponent, sErrorDescription)	
					Response.Write "</TR>"
				Response.Write "</TABLE>"
			'End If
		Response.Write "</FONT></DIV>"
	End If

	DisplayEmployeeExport = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeeForm(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about an employee from the
'         database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, bFull, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeForm"
	Dim sNames
	Dim oRecordset
	Dim lErrorNumber
	Dim iIndex
	Dim sError
	Dim sStatusEmployeesIDs1
	Dim sDate
	Dim sPayrollsIDs
	Dim bAllow

	bVisible = True
	bReadOnly = False
	bAllow = False
	sDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	bActivate = True
	sReadOnly = " READONLY=""READONLY"""
	lErrorNumber = 0
	sError = sErrorDescription

	Select Case sURL
		Case "EmployeesAssignNumber", "EmployeesAssignTemporalNumber"
			aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE) = ",5,134,"
			sReadOnly = ""
		Case ",EmployeeManagement,"
			sReadOnly = " READONLY=""READONLY"""
			bActivate = False
			bReadOnly = True
			If (lReasonID = -1) And (CInt(oRequest("Tab").Item) = 2) Then bVisible = False
			lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudieron obtener las secciones que mostrar"
				lErrorNumber = GetSectionsToShow(oADODBConnection, aEmployeeComponent, lReasonID, sErrorDescription)
				aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE) = "," & aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE) & ","
			End If
		Case Else
			If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
				aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE) = ",101,"
			Else
				sErrorDescription = "No se pudieron obtener las secciones que mostrar"
				lErrorNumber = GetSectionsToShow(oADODBConnection, aEmployeeComponent, lReasonID, sErrorDescription)
				If lErrorNumber = 0 Then
					If lReasonID = EMPLOYEES_GRADE Then
						aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE) = "3,120,125,128"
					ElseIf lReasonID = EMPLOYEES_SERVICE_SHEET Then
						aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE) = "3,120,125,128"
					End If
					aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE) = "," & aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE) & ","
					If (lReasonID = 12) Or (lReasonID = 13) Or (lReasonID = 14) Or (lReasonID = 17) Or (lReasonID = 18) Or (lReasonID = 57) Or (lReasonID = 58) Then
						sReadOnly = ""
						bActivate = True
						bReadOnly = False 
					End If
					lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						If lErrorNumber = 0 Then
							If Len(oRequest("ModifyConcept").Item) > 0 Then
								lErrorNumber = GetEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							End If
						Else
							If Len(sErrorDescription) = 0 Then sErrorDescription = "No tiene permisos para realizar movimientos a empleados que pertenecen a otro centro de trabajo."
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							Response.Write "<BR /><BR />"
							Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""30"" HEIGHT=""1"" />" & "<A HREF=""UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & """><B>Consultar otro empleado</B></A>"
							Response.Write "<BR /><BR />"
						End If
						If lErrorNumber = 0 Then
							Select Case lReasonID
								Case EMPLOYEES_THIRD_CONCEPT, EMPLOYEES_BENEFICIARIES_DEBIT
									If CInt(oRequest("CreditChange").Item) = 1 Then
										lErrorNumber = GetEmployeeCredit(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
									End If
								Case EMPLOYEES_CHILDREN_SCHOOLARSHIPS, CANCEL_EMPLOYEES_CONCEPTS, CANCEL_EMPLOYEES_SSI, CANCEL_EMPLOYEES_C04
									bActivate = False
									bReadOnly = True
								Case EMPLOYEES_FOR_RISK, EMPLOYEES_ANTIQUITIES
									If aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) <> 1 Then
										bVisible = False
										bActivate = False
										bReadOnly = True
										Call DisplayErrorMessage("Mensaje del sistema", "Este movimiento solo se puede otorgar a personal con puesto de base")
									End If
									If lReasonID = EMPLOYEES_FOR_RISK Then
										bActivate = True
									End If
								'Case 26
								'	If aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) <> 1 Then
								'		bVisible = False
								'		bActivate = False
								'		bReadOnly = True
								'		sError = "Los cambios por permuta de plazas solo se pueden otorgar a personal con puesto de base"
								'	End If
								Case 28
									aJobComponent(N_ID_JOB) = aEmployeeComponent(N_JOB_ID_EMPLOYEE)
									If (aJobComponent(N_ID_JOB) > 0) Then
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JobID From EmployeesHistoryList Where EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & " And JobID <> -3 And ReasonID In (37,38,39,40,41,43,44,45,46,47,48,49) Order By EmployeeDate Desc", "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
										aJobComponent(N_ID_JOB) = oRecordset.Fields("JobID").Value
										aEmployeeComponent(N_JOB_ID_EMPLOYEE) = aJobComponent(N_ID_JOB)
										lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
										If (aJobComponent(N_STATUS_ID_JOB) <> 2) And (aJobComponent(N_STATUS_ID_JOB) <> 4) And (aJobComponent(N_STATUS_ID_JOB) <> 5) And (InStr(1, ",54,58,62,66,70,86,", aEmployeeComponent(N_STATUS_ID_EMPLOYEE), vbBinaryCompare) = 0) Then
											sError = "El empleado no puede reanudar si su plaza se encuentra en el estatus de ocupada."
											bActivate = False
										End If
									ElseIf (aJobComponent(N_ID_JOB) = -3) Then
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JobId From JobsHistoryList Where EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & " Order By EndDate Desc", "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
										aJobComponent(N_ID_JOB) = oRecordset.Fields("JobID").Value
										aEmployeeComponent(N_JOB_ID_EMPLOYEE) = aJobComponent(N_ID_JOB)
										lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
										If (aJobComponent(N_STATUS_ID_JOB) <> 2) And (aJobComponent(N_STATUS_ID_JOB) <> 4) And (aJobComponent(N_STATUS_ID_JOB) <> 5) And (InStr(1, ",54,58,62,66,70,86,", aEmployeeComponent(N_STATUS_ID_EMPLOYEE), vbBinaryCompare) = 0) Then
											sError = "El empleado no puede reanudar si su plaza se encuentra en el estatus de ocupada."
											bActivate = False
										End If
									Else
										sError = "El empleado no tiene asignada una plaza"
										bActivate = False
									End If
								Case 29, 33, 34, 36, 37, 38, 39, 40, 41, 43, 44, 45, 46, 47, 48 '30, 31, 32, 
									If aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) <> 1 Then
										bVisible = False
										bActivate = False
										bReadOnly = True
										sError = "Los movimientos de licencia solo se pueden otorgar a personal con puesto de base"
									End If
								Case 58
									lErrorNumber = CheckExistencyOfEmployeeJob(aEmployeeComponent, sErrorDescription)
									If lErrorNumber <> 0 Then
										Call DisplayErrorMessage("Mensaje del sistema", "No se puede reasignar número de empleado si el empleado ha tenido alguna vez alguna plaza en el Instituto.")
									End If
								Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS
									If CInt(oRequest("SundayChange").Item) = 1 Then
										lErrorNumber = GetEmployeeSundays(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
									End If
									If aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 1 Then
										Call DisplayErrorMessage("Mensaje del sistema", "El empleado pertenece al tabulador de funcionarios")
										bVisible = False
									End If
								Case EMPLOYEES_CONCEPT_08
									If aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) <> 2 Then
										bActivate = True
										bVisible = False
										Call DisplayErrorMessage("Mensaje del sistema", "El empleado no tiene puesto de confianza")
									ElseIf (aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 1) Then
										bVisible = False
										Call DisplayErrorMessage("Mensaje del sistema", "El empleado pertenece al tabulador de funcionarios")
									End If
								Case EMPLOYEES_NIGHTSHIFTS
									If (aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) <> 21) And (aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) <> 22) and (aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) <> 23) Then
										bVisible = False
										Call DisplayErrorMessage("Mensaje del sistema", "El empleado no tiene el turno 21, 22 ó 23")
									End If
								Case EMPLOYEES_ADD_BENEFICIARIES, EMPLOYEES_CREDITORS
									If CInt(oRequest("BeneficiaryChange").Item) = 1 Then
										Select Case lReasonID
											Case EMPLOYEES_ADD_BENEFICIARIES
												lErrorNumber = GetEmployeeBeneficiary(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
												If lErrorNumber = 0 Then
													lErrorNumber = GetAlimonyType(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
												End If
											Case EMPLOYEES_CREDITORS
												lErrorNumber = GetEmployeeCreditor(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
												If lErrorNumber = 0 Then
													lErrorNumber = GetCreditorType(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
												End If
										End Select
									End If
								Case EMPLOYEES_BANK_ACCOUNTS
									If CInt(oRequest("BankAccountChange").Item) = 1 Then
										lErrorNumber = GetEmployeeBankAccount(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
									End If
								Case EMPLOYEES_ADDITIONALSHIFT
									bActivate = True
								Case EMPLOYEES_SERVICE_SHEET
									If CInt(oRequest("ServiceSheetChange").Item) = 1 Then
										lErrorNumber = GetEmployeesDocument(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
									End If
							End Select
							If aEmployeeComponent(N_DOCUMENT_FOR_LICENSE_ID_EMPLOYEE) <> -1 Then
								lErrorNumber = GetDocumentsForLicenses(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							End If
							If (lReasonID = 12 Or lReasonID = 13 Or lReasonID = 17 Or lReasonID = 18 Or lReasonID = 28) And (aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 7 or aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 12 or aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 13) Then
								bActivate = False
								sError = "Este empleado pertenece al tabulador de honorarios"
							End If
							If lReasonID = 66 Then
								If (aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) <> 7 and aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) <> 12 and aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) <> 13) Then
									sError = "El empleado no corresponde al tipo de tabulador de honorarios"
									bActivate = False
								End If
							End If
							If (lReasonID <> 12) And (lReasonID <> 13) And (lReasonID <> 14) And (lReasonID <> 17) And (lReasonID <> 18) And (lReasonID <> 57) Then
								bReadOnly = True
							End If
						End If
						If lErrorNumber = 0 Then
							If lReasonID <> 57 Then
								lErrorNumber = GetEmployeeStatusToValidateTheMovement(oADODBConnection, lReasonID, aEmployeeComponent, sErrorDescription)
							End If
							If lErrorNumber <> 0 Then
								If (aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 7 or aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 12 or aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 13) Or _
									(InStr(1, ",13,17,28,43,44,45,46,47,48,68,78,82,90,94,98,102,106,110,114,118,140,", "," & lReasonID & ",", vbBinaryCompare) > 0) Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, EmployeeDate, EndDate, StatusID, ReasonID From EmployeesHistoryList Where EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "Order By EndDate Desc", "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
									lDisplayFormCurrentDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
									If CLng(oRecordset.Fields("EndDate").Value) > lDisplayFormCurrentDate Then
										Call GetNameFromTable(oADODBConnection, "StatusEmployees", aEmployeeComponent(N_STATUS_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
										sError = "El empleado se encuentra en el estatus: " & sNames
										sReadOnly = " READONLY=""READONLY"""
										bActivate = False
										lErrorNumber = 0
										If lReasonID = 57 Then
											Call DisplayErrorMessage("Mensaje del sistema", "El empleado se encuentra en el estatus: " & sNames)
										End If
									Else
										sError = ""
										bActivate = True
									End If
								Else
									Call GetNameFromTable(oADODBConnection, "StatusEmployees", aEmployeeComponent(N_STATUS_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
									sError = "El empleado se encuentra en el estatus: " & sNames
									sReadOnly = " READONLY=""READONLY"""
									bActivate = False
									lErrorNumber = 0
									If lReasonID = 57 Then
										Call DisplayErrorMessage("Mensaje del sistema", "El empleado se encuentra en el estatus: " & sNames)
									End If
								End If
							Else
								If lReasonID = 14 Then
									If (aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) <> 7 And aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) <> 12 and aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) <> 13) Then
										sError = "El número de empleado no corresponde al tipo de tabulador de honorarios"
										bActivate = False
									End If
								ElseIf lReasonID = 66 Then
									If (aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) <> 7 and aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) <> 12 and aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) <> 13) Then
										sError = "El empleado no corresponde al tipo de tabulador de honorarios"
										bActivate = False
									End If
								ElseIf (lReasonID = 1) Or (lReasonID = 2) Or (lReasonID = 3) Or (lReasonID = 4) Or (lReasonID = 5) Or (lReasonID = 6) Or (lReasonID = 8) Or (lReasonID = 10) Or (lReasonID = 62) Or (lReasonID = 63) Then
									If (aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 7 or aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 12 or aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 13)   Then
										sError = "El número de empleado corresponde al tipo de tabulador de honorarios"
										bActivate = False
									End If
								End If
							End If
						End If
					Else
						sErrorDescription = "El número de empleado no se encuentra registrado en el sistema"
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						Response.Write "<BR /><BR />"
						Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""30"" HEIGHT=""1"" />" & "<A HREF=""UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & """><B>Consultar otro empleado</B></A>"
						Response.Write "<BR /><BR />"
					End If
				End If
			End If
	End Select
	If lErrorNumber = 0 Then
		If lErrorNumber = 0 Then
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "var bCheckReason = true;" & vbNewLine
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",5,", vbBinaryCompare) > 0) Then
					If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
						Response.Write "var asConsecutiveIDs = new Array("
							If (InStr(1, sURL, "EmployeesAssignTemporalNumber", vbBinaryCompare) > 0) Then
								Call GenerateJavaScriptArrayFromQuery(oADODBConnection, "ConsecutiveIDs2", "IDType", "CurrentID", "", "IDType", sErrorDescription)
							Else
								Call GenerateJavaScriptArrayFromQuery(oADODBConnection, "ConsecutiveIDs", "IDType", "CurrentID", "", "IDType", sErrorDescription)
							End If
						Response.Write "['-1', '']);" & vbNewLine
					End If
				End If
				Response.Write "function TotalDays(iDay, iMonth, iYear){" & vbNewLine
					Response.Write "iMonth = (iMonth + 9) % 12;" & vbNewLine
					Response.Write "iYear = iYear - Math.floor(iMonth/10);" & vbNewLine
					Response.Write "return (365 * iYear + Math.floor(iYear/4) - Math.floor(iYear/100) + Math.floor(iYear/400) + Math.floor((iMonth * 306 + 5)/10) + iDay - 1)" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "function GetDifferenceBetweenDates(iDay1, iMonth1, iYear1, iDay2, iMonth2, iYear2){" & vbNewLine
					Response.Write "return TotalDays(iDay2, iMonth2, iYear2) - TotalDays(iDay1, iMonth1, iYear1)" & vbNewLine
				Response.Write "}" & vbNewLine
				If lReasonID = EMPLOYEES_ADD_BENEFICIARIES Then
					Response.Write "var aAlimonyTypes = new Array("
						Call GenerateJavaScriptArrayFromQuery(oADODBConnection, "AlimonyTypes, QttyValues", "AlimonyTypeID", "ConceptQttyID, QttyName", "(AlimonyTypes.ConceptQttyID=QttyValues.QttyID) And (AlimonyTypes.Active=1)", "AlimonyTypeID", sErrorDescription)
					Response.Write "['-1', '-1', '']);" & vbNewLine
				End If
				If lReasonID = EMPLOYEES_CREDITORS Then
					Response.Write "var aAlimonyTypes = new Array("
						Call GenerateJavaScriptArrayFromQuery(oADODBConnection, "CreditorsTypes, QttyValues", "CreditorTypeID", "ConceptQttyID, QttyName", "(CreditorsTypes.ConceptQttyID=QttyValues.QttyID) And (CreditorsTypes.Active=1)", "CreditorTypeID", sErrorDescription)
					Response.Write "['-1', '-1', '']);" & vbNewLine
				End If
				If lReasonID < 0 Then
					Response.Write "function ShowAmountFields(sValue, sFieldsName) {" & vbNewLine
						Select Case lReasonID
							Case EMPLOYEES_SERVICE_SHEET
							Case EMPLOYEES_ADD_BENEFICIARIES, EMPLOYEES_CREDITORS
								Response.Write "var oForm = document.EmployeeBeneficiaryFrm;" & vbNewLine
								Response.Write "if (oForm) {" & vbNewLine
									Response.Write "for (var i=0; i<aAlimonyTypes.length; i++)" & vbNewLine
										Response.Write "if (aAlimonyTypes[i][0] == sValue) {" & vbNewLine
											Response.Write "RemoveAllItemsFromList(null, oForm.ConceptQttyID);" & vbNewLine
											Response.Write "AddItemToList(aAlimonyTypes[i][2], aAlimonyTypes[i][1], null, oForm.ConceptQttyID);" & vbNewLine
											Response.Write "break;" & vbNewLine
										Response.Write "}" & vbNewLine
								Response.Write "}" & vbNewLine
							Case EMPLOYEES_BANK_ACCOUNTS
								Response.Write "if (document.EmployeeFrm.Cheque.checked) {" & vbNewLine
									Response.Write "HideDisplay(document.all['AccountNumberSpn']);" & vbNewLine
									Response.Write "HideDisplay(document.all['SucursalSpn']);" & vbNewLine
								Response.Write "} else {" & vbNewLine
									Response.Write "HideDisplay(document.all[sFieldsName]);" & vbNewLine
									Response.Write "if (sValue==3) {" & vbNewLine
										Response.Write "ShowDisplay(document.all[sFieldsName]);" & vbNewLine
									Response.Write "} else {" & vbNewLine
										Response.Write "HideDisplay(document.all[sFieldsName]);" & vbNewLine
										Response.Write "document.EmployeeFrm.Sucursal.value='';" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "}" & vbNewLine
							Case Else
								Response.Write "var oForm = document.EmployeeFrm;" & vbNewLine
								Response.Write "if (oForm) {" & vbNewLine
									Response.Write "HideDisplay(document.all[sFieldsName + 'CurrencySpn']);" & vbNewLine
									Response.Write "HideDisplay(document.all[sFieldsName + 'AppliesToSpn']);" & vbNewLine
									Response.Write "switch (sValue) {" & vbNewLine
										Response.Write "case '1':" & vbNewLine
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
									Response.Write "}" & vbNewLine
								Response.Write "}" & vbNewLine
						End Select
					Response.Write "} // End of ShowAmountFields" & vbNewLine
					If lReasonID = EMPLOYEES_BANK_ACCOUNTS Then
						Response.Write "function VerifyAccountNumberLength(sAcount, iBankId)" & vbNewLine
						Response.Write "{" & vbNewLine
							Response.Write "switch (iBankId) {" & vbNewLine
								Response.Write "case '1':" & vbNewLine
									Response.Write "if (sAcount.length == 20) {" & vbNewLine
										Response.Write "return true;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "break;" & vbNewLine
								Response.Write "case '3':" & vbNewLine
									Response.Write "if (sAcount.length == 20) {" & vbNewLine
										Response.Write "return true;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "break;" & vbNewLine
								Response.Write "case '14':" & vbNewLine
									Response.Write "if (sAcount.length == 18) {" & vbNewLine
										Response.Write "return true;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "break;" & vbNewLine
								Response.Write "case '17':" & vbNewLine
									Response.Write "if (sAcount.length == 11) {" & vbNewLine
										Response.Write "return true;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "break;" & vbNewLine
								Response.Write "case '24':" & vbNewLine
									Response.Write "if (sAcount.length == 10) {" & vbNewLine
										Response.Write "return true;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "break;" & vbNewLine
							Response.Write "}" & vbNewLine
						Response.Write "}" & vbNewLine
					End If
				End If
				If (lReasonID > 0) Or (lReasonID = -64) Or (lReasonID = -75) Then
					Response.Write "function ShowHideApplyMovementButton(sValue) {" & vbNewLine
						Response.Write "var oForm = document.EmployeeFrm" & vbNewLine
							Response.Write "if (oForm) {" & vbNewLine
								Response.Write "if (sValue == 0) {" & vbNewLine
									Response.Write "HideDisplay(document.all['ApplyMovementDiv']);" & vbNewLine
								Response.Write "} else {" & vbNewLine
									Response.Write "ShowDisplay(document.all['ApplyMovementDiv']);" & vbNewLine
								Response.Write "}" & vbNewLine
							Response.Write "}" & vbNewLine
					Response.Write "} // End of ShowHideApplyMovementButton" & vbNewLine
					Response.Write "function IsPayrollOpenToApplyMovements(sValue) {" & vbNewLine
						lErrorNumber = GetPayrollsEnableToApplyMovements(sPayrollsIDs, CInt(Request.Cookies("SIAP_SectionID")), sErrorDescription)
						If (lErrorNumber = L_ERR_NO_RECORDS) Then
							Response.Write "return false;" & vbNewLine
							sErrorDescription = ""
							lErrorNumber = 0
						Else
							sDisplayFormCaseOptions = Split(sPayrollsIDs, "," , -1, vbBinaryCompare)
							Response.Write "switch (sValue) {" & vbNewLine
								Response.Write "case '-2':" & vbNewLine
								For iIndex = 0 To UBound(sDisplayFormCaseOptions)
									Response.Write "case '" & CStr(sDisplayFormCaseOptions(iIndex)) & "':" & vbNewLine
								Next
								Response.Write "return true;" & vbNewLine
								Response.Write "default:" & vbNewLine
									Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
						End If
					Response.Write "} // End of IsPayrollOpenToApplyMovements" & vbNewLine
					Response.Write "function CheckStatusPayrolls() {" & vbNewLine
						Response.Write "if (IsPayrollOpenToApplyMovements(document.EmployeeFrm.EmployeePayrollDateCmb.value)) {" & vbNewLine
							Response.Write "ShowDisplay(document.all['ApplyMovementDiv'])" & vbNewLine
						Response.Write "}"  & vbNewLine
						Response.Write "else {"  & vbNewLine
						Response.Write "HideDisplay(document.all['ApplyMovementDiv'])" & vbNewLine
						Response.Write "}"  & vbNewLine
					Response.Write "} // End of CheckStatusPayrolls" & vbNewLine
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",", vbBinaryCompare) > 0) Then
					Response.Write "function CheckEmployeeFields(oForm) {" & vbNewLine
						If lReasonID <> -96 Then
							If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",1,", vbBinaryCompare) > 0) Or (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",101,", vbBinaryCompare) > 0) Then
								Response.Write "if (oForm) {" & vbNewLine
									Response.Write "if ((oForm.EmployeeNumber.value.length == 0) || (oForm.EmployeeNumber.value == '')) {" & vbNewLine
										Response.Write "alert('Favor de introducir el número de empleado.');" & vbNewLine
										Response.Write "oForm.EmployeeNumber.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "}" & vbNewLine
							End If
							If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",103,", vbBinaryCompare) > 0) Then
								Response.Write "if (oForm) {" & vbNewLine
									Response.Write "if (oForm.EmployeeNumber2.value == '') {" & vbNewLine
										Response.Write "alert('Favor de introducir el número de empleado.');" & vbNewLine
										Response.Write "oForm.EmployeeNumber2.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "}" & vbNewLine
							End If
							If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then
								If aEmployeeComponent(N_STATUS_REASON_ID_EMPLOYEE) = 2 Or aEmployeeComponent(N_STATUS_REASON_ID_EMPLOYEE) = 3 Then
									Response.Write "if (oForm) {" & vbNewLine
										Response.Write "if (bCheckReason) {" & vbNewLine
											Response.Write "if (oForm.Comments.value == '') {" & vbNewLine
												Response.Write "alert('Favor de introducir la razón del rechazo.');" & vbNewLine
												Response.Write "oForm.Comments.focus();" & vbNewLine
												Response.Write "return false;" & vbNewLine
											Response.Write "}" & vbNewLine
										Response.Write "}" & vbNewLine
									Response.Write "}" & vbNewLine
								End If
							End If
							If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",2,", vbBinaryCompare) > 0) Then
								Response.Write "if (oForm) {" & vbNewLine
									Response.Write "if ((oForm.JobID.value == null) || (oForm.JobID.value == '') || (oForm.JobID.value == -1) || (oForm.JobID.value == '-1')) {" & vbNewLine
										Response.Write "alert('Favor de especificar el número de plaza.');" & vbNewLine
										If lReasonID <> 14 Then
											Response.Write "oForm.JobNumber.focus();" & vbNewLine
										End If
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "}" & vbNewLine
							End If
							If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",8,", vbBinaryCompare) > 0) Then
								Response.Write "if (oForm) {" & vbNewLine
									Response.Write "if ((parseInt('1' + oForm.EmployeeDay.value) - 100) * (parseInt('1' + oForm.EmployeeMonth.value) - 100) * parseInt(oForm.EmployeeYear.value) == 0) {" & vbNewLine
										Response.Write "alert('La fecha de inicio de la vigencia es requerida');" & vbNewLine
										Response.Write "oForm.EmployeeDay.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "}" & vbNewLine
							End If
							If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",107,", vbBinaryCompare) > 0) Then
								If lReasonID = -64 Or lReasonID = -75 Then
								Response.Write "if (oForm) {" & vbNewLine
									Response.Write "if (parseInt(oForm.StartHour3.value) * parseInt(oForm.EndHour3.value) == 0) {" & vbNewLine
										Response.Write "alert('No ha indicado el horario correspondiente');" & vbNewLine
										Response.Write "oForm.StartHour3.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if ((parseInt('1' + oForm.StartHour3.value) - 100) == (parseInt('1' + oForm.EndHour3.value) - 100)) {" & vbNewLine
										Response.Write "alert('No ha indicado un horario correcto');" & vbNewLine
										Response.Write "oForm.StartHour3.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if ((parseInt('1' + oForm.StartHour3.value) - 100) > (parseInt('1' + oForm.EndHour3.value) - 100)) {" & vbNewLine
										Response.Write "alert('La hora inicial no puede ser mayor a la final');" & vbNewLine
										Response.Write "oForm.StartHour3.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "}" & vbNewLine
								End If
							End If
							If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",112,", vbBinaryCompare) > 0) Then
								Response.Write "if (oForm) {" & vbNewLine
								If (lReasonID <> 12) And (lReasonID <> 17) And (lReasonID <> 21) And (lReasonID <> 47) And (lReasonID <> 50) And (lReasonID <> 51) And (lReasonID <> 68) Then
									Response.Write "if (((parseInt('1' + oForm.EmployeeEndDay.value) - 100) + (parseInt('1' + oForm.EmployeeEndMonth.value) - 100) + parseInt(oForm.EmployeeEndYear.value)) > 0 ) {" & vbNewLine
										Response.Write "if ((parseInt('1' + oForm.EmployeeEndDay.value) - 100) * (parseInt('1' + oForm.EmployeeEndMonth.value) - 100) * parseInt(oForm.EmployeeEndYear.value) == 0 ) {" & vbNewLine
											Response.Write "alert('Favor de introducir la fecha final de la vigencia');" & vbNewLine
											Response.Write "oForm.EmployeeEndDay.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									If (InStr(1,",1,2,3,4,5,6,8,10,12,17,18,21,26,28,29,37,39,40,44,45,46,47,50,51,57,62,63,66,68,","," &lReasonID & ",", vbBinaryCompare) = 0) Then
										Response.Write "if ((parseInt('1' + oForm.EmployeeEndDay.value) - 100) * (parseInt('1' + oForm.EmployeeEndMonth.value) - 100) * parseInt(oForm.EmployeeEndYear.value) == 0 ) {" & vbNewLine
											Response.Write "alert('Favor de introducir la fecha final de la vigencia');" & vbNewLine
											Response.Write "oForm.EmployeeEndDay.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									End If
									Response.Write "}" & vbNewLine
								End If
									Select Case lReasonID
										Case 14
											Response.Write "if (GetDifferenceBetweenDates(parseInt(oForm.EmployeeDay.value), parseInt(oForm.EmployeeMonth.value), parseInt(oForm.EmployeeYear.value), parseInt(oForm.EmployeeEndDay.value), parseInt(oForm.EmployeeEndMonth.value), parseInt(oForm.EmployeeEndYear.value)) > 366) {" & vbNewLine
												Response.Write "alert('La vigencia del contrato de honorarios no puede ser mayor a 1 año.');" & vbNewLine
												Response.Write "oForm.EmployeeDay.focus();" & vbNewLine
												Response.Write "return false;" & vbNewLine
											Response.Write "}" & vbNewLine
										Case 37
											Response.Write "if (GetDifferenceBetweenDates(parseInt(oForm.EmployeeDay.value), parseInt(oForm.EmployeeMonth.value), parseInt(oForm.EmployeeYear.value), parseInt(oForm.EmployeeEndDay.value), parseInt(oForm.EmployeeEndMonth.value), parseInt(oForm.EmployeeEndYear.value)) > 180) {" & vbNewLine
												Response.Write "alert('El período autorizado para prórroga de licencia sin goce de sueldo por comisión sindical no puede ser mayor de 180 días.');" & vbNewLine
												Response.Write "return false;" & vbNewLine
											Response.Write "}" & vbNewLine
										Case 38
											Response.Write "if (GetDifferenceBetweenDates(parseInt(oForm.EmployeeDay.value), parseInt(oForm.EmployeeMonth.value), parseInt(oForm.EmployeeYear.value), parseInt(oForm.EmployeeEndDay.value), parseInt(oForm.EmployeeEndMonth.value), parseInt(oForm.EmployeeEndYear.value)) > 183) {" & vbNewLine
												Response.Write "alert('El período autorizado para prórroga de licencia sin goce de sueldo por otorgamiento de beca no puede ser mayor de 180 días.');" & vbNewLine
												Response.Write "return false;" & vbNewLine
											Response.Write "}" & vbNewLine
										Case 39
											Response.Write "if (GetDifferenceBetweenDates(parseInt(oForm.EmployeeDay.value), parseInt(oForm.EmployeeMonth.value), parseInt(oForm.EmployeeYear.value), parseInt(oForm.EmployeeEndDay.value), parseInt(oForm.EmployeeEndMonth.value), parseInt(oForm.EmployeeEndYear.value)) > 183) {" & vbNewLine
												Response.Write "alert('El período autorizado para prórroga de licencia sin goce de sueldo por ocupar cargo de elección popular o puesto de confianza fuera del Instituto no puede ser mayor de 180 días.');" & vbNewLine
												Response.Write "return false;" & vbNewLine
											Response.Write "}" & vbNewLine
										Case 40
											Response.Write "if (GetDifferenceBetweenDates(parseInt(oForm.EmployeeDay.value), parseInt(oForm.EmployeeMonth.value), parseInt(oForm.EmployeeYear.value), parseInt(oForm.EmployeeEndDay.value), parseInt(oForm.EmployeeEndMonth.value), parseInt(oForm.EmployeeEndYear.value)) > 183) {" & vbNewLine
												Response.Write "alert('El período autorizado para prórroga de licencia sin goce de sueldo por ocupar puesto de confianza dentro del Instituto fuera del Instituto no puede ser mayor de 180 días.');" & vbNewLine
												Response.Write "return false;" & vbNewLine
											Response.Write "}" & vbNewLine
										Case 41
											Response.Write "if (GetDifferenceBetweenDates(parseInt(oForm.EmployeeDay.value), parseInt(oForm.EmployeeMonth.value), parseInt(oForm.EmployeeYear.value), parseInt(oForm.EmployeeEndDay.value), parseInt(oForm.EmployeeEndMonth.value), parseInt(oForm.EmployeeEndYear.value)) > 183) {" & vbNewLine
												Response.Write "alert('El período autorizado para la prórroga de licencia sin goce de sueldo por asuntos particulares no puede ser mayor de 180 días.');" & vbNewLine
												Response.Write "oForm.EmployeeDay.focus();" & vbNewLine
												Response.Write "return false;" & vbNewLine
											Response.Write "}" & vbNewLine
										Case 44
											Response.Write "if (GetDifferenceBetweenDates(parseInt(oForm.EmployeeDay.value), parseInt(oForm.EmployeeMonth.value), parseInt(oForm.EmployeeYear.value), parseInt(oForm.EmployeeEndDay.value), parseInt(oForm.EmployeeEndMonth.value), parseInt(oForm.EmployeeEndYear.value)) > 183) {" & vbNewLine
												Response.Write "alert('El período autorizado para licencia sin goce de sueldo por comisión sindical no puede ser mayor de 180 días.');" & vbNewLine
												Response.Write "oForm.EmployeeDay.focus();" & vbNewLine
												Response.Write "return false;" & vbNewLine
											Response.Write "}" & vbNewLine
										Case 45
											Response.Write "if (GetDifferenceBetweenDates(parseInt(oForm.EmployeeDay.value), parseInt(oForm.EmployeeMonth.value), parseInt(oForm.EmployeeYear.value), parseInt(oForm.EmployeeEndDay.value), parseInt(oForm.EmployeeEndMonth.value), parseInt(oForm.EmployeeEndYear.value)) > 365) {" & vbNewLine
												Response.Write "alert('El período autorizado para licencia sin goce de sueldo por otorgamiento de beca no puede ser mayor de 1 año.');" & vbNewLine
												Response.Write "oForm.EmployeeDay.focus();" & vbNewLine
												Response.Write "return false;" & vbNewLine
											Response.Write "}" & vbNewLine
										Case 46
											Response.Write "if (GetDifferenceBetweenDates(parseInt(oForm.EmployeeDay.value), parseInt(oForm.EmployeeMonth.value), parseInt(oForm.EmployeeYear.value), parseInt(oForm.EmployeeEndDay.value), parseInt(oForm.EmployeeEndMonth.value), parseInt(oForm.EmployeeEndYear.value)) > 183) {" & vbNewLine
												Response.Write "alert('El período autorizado para licencia sin goce de sueldo por ocupar cargo de elección popular o puesto de confianza fuera del instituto no puede ser mayor de 180 días.');" & vbNewLine
												Response.Write "oForm.EmployeeDay.focus();" & vbNewLine
												Response.Write "return false;" & vbNewLine
											Response.Write "}" & vbNewLine
										Case 48
											Response.Write "if (GetDifferenceBetweenDates(parseInt(oForm.EmployeeDay.value), parseInt(oForm.EmployeeMonth.value), parseInt(oForm.EmployeeYear.value), parseInt(oForm.EmployeeEndDay.value), parseInt(oForm.EmployeeEndMonth.value), parseInt(oForm.EmployeeEndYear.value)) > 183) {" & vbNewLine
												Response.Write "alert('El período autorizado para licencia sin goce de sueldo por práctica de servicio social no puede ser mayor de 180 días.');" & vbNewLine
												Response.Write "oForm.EmployeeDay.focus();" & vbNewLine
												Response.Write "return false;" & vbNewLine
											Response.Write "}" & vbNewLine
									End Select
									If (InStr(1,",1,2,3,4,5,6,8,10,12,17,18,21,26,28,29,37,39,40,44,45,46,47,50,51,62,63,66,68,","," &lReasonID & ",", vbBinaryCompare) = 0) Then
										Response.Write "if ((parseInt('1' + oForm.EmployeeEndDay.value) - 100) * (parseInt('1' + oForm.EmployeeEndMonth.value) - 100) * parseInt(oForm.EmployeeEndYear.value) == 0 ) {" & vbNewLine
											Response.Write "alert('Favor de introducir la fecha final de la vigencia');" & vbNewLine
											Response.Write "oForm.EmployeeEndDay.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									End If
										Response.Write "if (((parseInt('1' + oForm.EmployeeEndDay.value) - 100) + ('1' + parseInt(oForm.EmployeeEndMonth.value) - 100) + parseInt(oForm.EmployeeEndYear.value)) > 0 ) {" & vbNewLine
											Response.Write "if ((parseInt('1' + oForm.EmployeeEndDay.value) - 100) * (parseInt('1' + oForm.EmployeeEndMonth.value) - 100) * parseInt(oForm.EmployeeEndYear.value) > 0 ) {" & vbNewLine
												Response.Write "if (((parseInt('1' + oForm.EmployeeDay.value) - 100) + ((parseInt('1' + oForm.EmployeeMonth.value) -100) * 100) + parseInt(oForm.EmployeeYear.value) * 10000) > ((parseInt('1' + oForm.EmployeeEndDay.value) - 100) + ((parseInt('1' + oForm.EmployeeEndMonth.value) - 100) * 100) + parseInt(oForm.EmployeeEndYear.value) * 10000)) {" & vbNewLine
													Response.Write "alert('Favor de verificar la vigencia del movimiento');" & vbNewLine
													Response.Write "oForm.EmployeeDay.focus();" & vbNewLine
													Response.Write "return false;" & vbNewLine
												Response.Write "}" & vbNewLine
											Response.Write "}" & vbNewLine
											Response.Write "else" & vbNewLine
											Response.Write "{" & vbNewLine
												Response.Write "alert('Favor de verificar la vigencia del movimiento');" & vbNewLine
												Response.Write "return false;" & vbNewLine
											Response.Write "}" & vbNewLine
										Response.Write "}" & vbNewLine
									Response.Write "if (parseInt(oForm.EmployeePayrollDate.value)==-1) {" & vbNewLine
										Response.Write "alert('No existen nóminas abiertas para el registro de movimientos.');" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "}" & vbNewLine
							End If
							If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",116,", vbBinaryCompare) > 0) Then
								Call DisplayEmployeeFormSection116a()
							End If
							If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",3,", vbBinaryCompare) > 0) Then
								Call DisplayEmployeeFormSection03()
							End If
							If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",4,", vbBinaryCompare) > 0) Then
								Call DisplayEmployeeFormSection04()
							End If
							If (lReasonID = 0) Or (lReasonID = 58) Or (lReasonID = 67) Then
								Response.Write "if (oForm) {" & vbNewLine
									Response.Write "if (oForm.EmployeeName.value == '') {" & vbNewLine
										Response.Write "alert('Favor de introducir el nombre del empleado.');" & vbNewLine
										Response.Write "oForm.EmployeeName.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "} else {" & vbNewLine
										Response.Write "oForm.EmployeeName.value = oForm.EmployeeName.value.toUpperCase();" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.EmployeeLastName.value == '') {" & vbNewLine
										Response.Write "alert('Favor de introducir el apellido paterno del empleado.');" & vbNewLine
										Response.Write "oForm.EmployeeLastName.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.EmployeeLastName.value != '') {" & vbNewLine
										Response.Write "oForm.EmployeeLastName.value = oForm.EmployeeLastName.value.toUpperCase();" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.EmployeeTypeID.value == '-1') {" & vbNewLine
										Response.Write "alert('Favor de seleccionar el tipo empleado.');" & vbNewLine
										Response.Write "oForm.EmployeeTypeID.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.RFC.value.length != 13) {" & vbNewLine
										Response.Write "alert('Favor de introducir el RFC del empleado.');" & vbNewLine
										Response.Write "oForm.RFC.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "} else {" & vbNewLine
										Response.Write "oForm.RFC.value = oForm.RFC.value.toUpperCase();" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.BirthDay.value == 0) {" & vbNewLine
											Response.Write "alert('Favor de introducir el día de nacimiento del empleado.');" & vbNewLine
											Response.Write "oForm.BirthDay.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.BirthYear.value == 0){" & vbNewLine
										Response.Write "alert('Favor de introducir el año de nacimiento del empleado.');" & vbNewLine
										Response.Write "oForm.BirthYear.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.CURP.value.length != 18) {" & vbNewLine
										Response.Write "alert('Favor de introducir el CURP del empleado.');" & vbNewLine
										Response.Write "oForm.CURP.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "} else {" & vbNewLine
										Response.Write "oForm.CURP.value = oForm.CURP.value.toUpperCase();" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.RFC.value.length == 13) {" & vbNewLine
										Response.Write "if (oForm.RFC.value.substr(8,2) != oForm.BirthDay.value) {" & vbNewLine
											Response.Write "alert('Favor verifique el día del RFC con el día de la fecha de nacimiento del empleado.');" & vbNewLine
											Response.Write "oForm.RFC.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.RFC.value.length == 13) {" & vbNewLine
										Response.Write "if (oForm.RFC.value.substr(6,2) != oForm.BirthMonth.value) {" & vbNewLine
											Response.Write "alert('Favor verifique el mes del RFC con el mes de la fecha de nacimiento del empleado.');" & vbNewLine
											Response.Write "oForm.RFC.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.RFC.value.length == 13) {" & vbNewLine
										Response.Write "if (oForm.RFC.value.substr(4,2) != oForm.BirthYear.value.substr(2,2)) {" & vbNewLine
											Response.Write "alert('Favor verifique el año del RFC con el año de la fecha de nacimiento del empleado.');" & vbNewLine
											Response.Write "oForm.RFC.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.RFC.value.substr(4,6) != oForm.CURP.value.substr(4,6)) {" & vbNewLine
										Response.Write "alert('Favor verifique el RFC con el CURP del empleado.');" & vbNewLine
										Response.Write "oForm.RFC.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.CURP.value.substr(10,1) != 'M' && oForm.CURP.value.substr(10,1) != 'H') {" & vbNewLine
										Response.Write "alert('Favor verifique el CURP del empleado.');" & vbNewLine
										Response.Write "oForm.CURP.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "}" & vbNewLine
							End If
							If ((lReasonID = 12) Or (lReasonID = 13) Or (lReasonID = 14) Or (lReasonID = 17) Or (lReasonID = 18) Or (lReasonID = 57)) And bActivate Then
								Response.Write "if (oForm) {" & vbNewLine
									Response.Write "if (oForm.EmployeeName.value == '') {" & vbNewLine
										Response.Write "alert('Favor de introducir el nombre del empleado.');" & vbNewLine
										Response.Write "oForm.EmployeeName.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.EmployeeLastName.value == '') {" & vbNewLine
										Response.Write "alert('Favor de introducir el apellido paterno.');" & vbNewLine
										Response.Write "oForm.EmployeeLastName.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.EmployeeTypeID.value == '-1') {" & vbNewLine
										Response.Write "alert('Favor de seleccionar el tipo empleado.');" & vbNewLine
										Response.Write "oForm.EmployeeTypeID.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.RFC.value.length != 13) {" & vbNewLine
										Response.Write "alert('Favor de introducir el RFC del empleado.');" & vbNewLine
										Response.Write "oForm.RFC.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.BirthDay.value == 0) {" & vbNewLine
										Response.Write "alert('Favor de introducir el día de nacimiento del empleado.');" & vbNewLine
										Response.Write "oForm.BirthDay.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.BirthYear.value == 0){" & vbNewLine
										Response.Write "alert('Favor de introducir el año de nacimiento del empleado.');" & vbNewLine
										Response.Write "oForm.BirthYear.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.CURP.value.length != 18) {" & vbNewLine
										Response.Write "alert('Favor de introducir el CURP del empleado.');" & vbNewLine
										Response.Write "oForm.CURP.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (parseInt(oForm.CountryIDCmb.value) == -1) {" & vbNewLine
										Response.Write "alert('Favor de indicar la nacionalidad del empleado.');" & vbNewLine
										Response.Write "oForm.CountryIDCmb.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.EmployeeAddress.value == '') {" & vbNewLine
										Response.Write "alert('Favor de introducir el domicilio del empleado.');" & vbNewLine
										Response.Write "oForm.EmployeeAddress.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.EmployeeCity.value == '') {" & vbNewLine
										Response.Write "alert('Favor de introducir la ciudad del domicilio del empleado.');" & vbNewLine
										Response.Write "oForm.EmployeeCity.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.EmployeeZipCode.value == '') {" & vbNewLine
										Response.Write "alert('Favor de introducir el código postal del domicilio del empleado.');" & vbNewLine
										Response.Write "oForm.EmployeeZipCode.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.RFC.value.length == 13) {" & vbNewLine
										Response.Write "if (oForm.RFC.value.substr(8,2) != oForm.BirthDay.value) {" & vbNewLine
											Response.Write "alert('Favor verifique el día del RFC con el día de la fecha de nacimiento del empleado.');" & vbNewLine
											Response.Write "oForm.RFC.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.RFC.value.length == 13) {" & vbNewLine
										Response.Write "if (oForm.RFC.value.substr(6,2) != oForm.BirthMonth.value) {" & vbNewLine
											Response.Write "alert('Favor verifique el mes del RFC con el mes de la fecha de nacimiento del empleado.');" & vbNewLine
											Response.Write "oForm.RFC.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.RFC.value.length == 13) {" & vbNewLine
										Response.Write "if (oForm.RFC.value.substr(4,2) != oForm.BirthYear.value.substr(2,2)) {" & vbNewLine
											Response.Write "alert('Favor verifique el año del RFC con el año de la fecha de nacimiento del empleado.');" & vbNewLine
											Response.Write "oForm.RFC.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "}" & vbNewLine
							End If
						Else
							Response.Write "if (oForm) {" & vbNewLine
								Response.Write "if (oForm.ConceptStartDate.value == '') {" & vbNewLine
									Response.Write "alert('Debe indicar el concepto a cancelar');" & vbNewLine
									Response.Write "return false;" & vbNewLine
								Response.Write "}" & vbNewLine
								Response.Write "if (GetDifferenceBetweenDates(parseInt(oForm.ConceptEndDay.value), parseInt(oForm.ConceptEndMonth.value), parseInt(oForm.ConceptEndYear.value), parseInt(oForm.ConceptStartDate.value.substr(6,2)), parseInt(oForm.ConceptStartDate.value.substr(4,2)), parseInt(oForm.ConceptStartDate.value.substr(0,4))) > 0) {" & vbNewLine
									Response.Write "alert ('La fecha para la cancelación del concepto es menor a su inicio');" & vbNewLine
									Response.Write "return false;" & vbNewLine
								Response.Write "}" & vbNewLine
							Response.Write "}" & vbNewLine
						End If
						Response.Write "return true;" & vbNewLine
					Response.Write "} // End of CheckEmployeeFields" & vbNewLine
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",5,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection05()
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",7,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection07()
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",130,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection130()
				End If
			Response.Write "//--></SCRIPT>" & vbNewLine
			If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",115,", vbBinaryCompare) > 0) Then
				Call DisplayEmployeeFormSection115a(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
			End If
			If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",120,", vbBinaryCompare) > 0) Then
				Call DisplayEmployeeFormSection120a(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
			End If
			If (lReasonID = EMPLOYEES_ADD_BENEFICIARIES) Or (lReasonID = EMPLOYEES_CREDITORS) Then
				Response.Write "<FORM NAME=""EmployeeBeneficiaryFrm"" ID=""EmployeeBeneficiaryFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckEmployeeBeneficiaryFields(this)"">"
			Else
				Response.Write "<FORM NAME=""EmployeeFrm"" ID=""EmployeeFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckEmployeeFields(this)"">"
			End If
				Select Case sURL
					Case "ServiceSheet"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""" & oRequest("SectionID").Item & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE=""" & lReasonID & """ />"
					Case Else
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_EMPLOYEE) & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE=""" & lReasonID & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AbsenceDate"" ID=""AbsenceDateHdn"" VALUE="""" />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OcurredDate"" ID=""OcurredDateHdn"" VALUE="""
							If (aEmployeeComponent(N_ID_EMPLOYEE) > -1) And (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) > -1) And (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) > 0) Then Response.Write aAbsenceComponent(N_OCURRED_DATE_ABSENCE)
						Response.Write """ />"
				End Select
				Select Case sURL
					Case "EmployeesAssignNumber", "EmployeesAssignTemporalNumber"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""EmployeesAssignNumber"" />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SaveEmployeesAssignNumber"" ID=""ActionHdn"" VALUE=""1"" />"
					Case "EmployeesMovements"
						If (InStr(1, ",101,", aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), vbBinaryCompare) > 0) Then
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""EmployeesMovements"" />"
						ElseIf (InStr(1, ",131,", aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), vbBinaryCompare) > 0) Then
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SaveDocumentsForLicenses"" ID=""SaveDocumentsForLicensesHdn"" VALUE=""1"" />"
						Else
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SaveEmployeesMovements"" ID=""ActionHdn"" VALUE=""1"" />"
							Select Case lReasonID
								Case EMPLOYEES_THIRD_CONCEPT, EMPLOYEES_BENEFICIARIES_DEBIT
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CreditChange"" ID=""CreditChangeHdn"" VALUE=""" & oRequest("CreditChange").Item & """ />"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CreditID"" ID=""CreditIDHdn"" VALUE=""" & oRequest("CreditID").Item & """ />"
								Case EMPLOYEES_ADD_BENEFICIARIES
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BeneficiaryChange"" ID=""BeneficiaryChangeHdn"" VALUE=""" & oRequest("BeneficiaryChange").Item & """ />"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BeneficiaryID"" ID=""BeneficiaryIDHdn"" VALUE=""" & oRequest("BeneficiaryID").Item & """ />"
									If CInt(oRequest("BeneficiaryChange").Item) Then
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BeneficiaryStartDate"" ID=""BeneficiaryStartDateHdn"" VALUE=""" & oRequest("BeneficiaryStartDate").Item & """ />"
									End If
								Case EMPLOYEES_CREDITORS
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BeneficiaryChange"" ID=""BeneficiaryChangeHdn"" VALUE=""" & oRequest("BeneficiaryChange").Item & """ />"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CreditorID"" ID=""CreditorIDHdn"" VALUE=""" & oRequest("CreditorID").Item & """ />"
									If CInt(oRequest("BeneficiaryChange").Item) Then
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CreditorStartDate"" ID=""CreditorStartDateHdn"" VALUE=""" & oRequest("CreditorStartDate").Item & """ />"
									End If
								Case EMPLOYEES_SUNDAYS, EMPLOYEES_EXTRAHOURS
									If CInt(oRequest("SundayChange").Item) Then
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SundayChange"" ID=""SundayChangeHdn"" VALUE=""" & oRequest("SundayChange").Item & """ />"
									End If
								Case EMPLOYEES_BANK_ACCOUNTS
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BankAccountChange"" ID=""BankAccountChangeHdn"" VALUE=""" & oRequest("BankAccountChange").Item & """ />"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AccountID"" ID=""AccountIDHdn"" VALUE=""" & oRequest("AccountID").Item & """ />"
								Case Else
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ModifyConcept"" ID=""ModifyConceptHdn"" VALUE=""" & oRequest("ModifyConcept").Item & """ />"
							End Select
						End If
					Case "EmployeesSafeSeparation"
					Case Else
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & sAction & """ />"
				End Select
				If (aEmployeeComponent(N_REASON_TYPE_ID_EMPLOYEE) >= 3 And aEmployeeComponent(N_REASON_TYPE_ID_EMPLOYEE) <= 6) Then
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobID"" ID=""JobIDHdn"" VALUE=""" & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & """ />"
				End If
				If lReasonID = 57 Then
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CompanyID"" ID=""CompanyIDHdn"" VALUE=""" & aEmployeeComponent(N_COMPANY_ID_EMPLOYEE) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobID"" ID=""JobIDHdn"" VALUE=""" & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ServiceID"" ID=""ServiceIDHdn"" VALUE=""" & aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PositionTypeID"" ID=""PositionTypeIDHdn"" VALUE=""" & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ClassificationID"" ID=""ClassificationIDHdn"" VALUE=""" & aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""GroupGradeLevelID"" ID=""GroupGradeLevelIDHdn"" VALUE=""" & aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IntegrationID"" ID=""IntegrationIDHdn"" VALUE=""" & aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JourneyID"" ID=""JourneyIDHdn"" VALUE=""" & aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""WorkingHours"" ID=""WorkingHoursHdn"" VALUE=""" & aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""LevelID"" ID=""LevelIDHdn"" VALUE=""" & aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StatusID"" ID=""StatusIDHdn"" VALUE=""" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PaymentCenterID"" ID=""PaymentCenterIDHdn"" VALUE=""" & aEmployeeComponent(N_PAYMENT_CENTER_ID_EMPLOYEE) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AntiquityID"" ID=""AntiquityIDHdn"" VALUE=""" & aEmployeeComponent(N_ANTIQUITY_EMPLOYEE) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Antiquity2ID"" ID=""Antiquity2IDHdn"" VALUE=""" & aEmployeeComponent(N_ANTIQUITY2_EMPLOYEE) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Antiquity3ID"" ID=""Antiquity3IDHdn"" VALUE=""" & aEmployeeComponent(N_ANTIQUITY3_EMPLOYEE) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Antiquity4ID"" ID=""Antiquity4IDHdn"" VALUE=""" & aEmployeeComponent(N_ANTIQUITY4_EMPLOYEE) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartDate"" ID=""StartDateHdn"" VALUE=""" & aEmployeeComponent(N_START_DATE_EMPLOYEE) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartDate2"" ID=""StartDate2Hdn"" VALUE=""" & aEmployeeComponent(N_START_DATE2_EMPLOYEE) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Active"" ID=""ActiveHdn"" VALUE=""" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & """ />"
				End If
				If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConsecutiveID"" ID=""ConsecutiveIDHdn"" VALUE="""" />"
				Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchRecordIFrame"" FRAMEBORDER=""0"" WIDTH=""0"" HEIGHT=""0""></IFRAME>"

				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",201,", vbBinaryCompare) > 0) Then
					Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
						Response.Write "<TR>"
							Response.Write "<TD VALIGN=""TOP"">"
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",101,", vbBinaryCompare) > 0) Then
					Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del empleado:&nbsp;</FONT></TD>"
							Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" VALUE=""" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & """ SIZE=""7"" MAXLENGTH=""7"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
					Response.Write "</TABLE>"
					Response.Write "<BR />"
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",120,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection120b(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",102,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection102(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",104,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection104(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",134,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection134(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",105,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection105(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",131,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection131(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",103,", vbBinaryCompare) > 0) Then
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Datos con los que se realizará la permuta</B></FONT>"
					Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
						Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
							Response.Write "<TR>"
								Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">No. del empleado:&nbsp;</FONT></TD>"
								Response.Write "<TD VALIGN=""TOP"">"
									Response.Write "<INPUT TYPE=""TEXT"" NAME=""EmployeeNumber2"" ID=""EmployeeNumber2Txt"" SIZE=""6"" MAXLENGTH=""6"" VALUE=""" & aEmployeeComponent(N_ID_EMPLOYEE_2) & """ CLASS=""TextFields"" onChange=""document.EmployeeFrm.EmployeeID.value='';"" />"
									Response.Write "<A HREF=""javascript: document.EmployeeFrm.EmployeeID.value=''; SearchRecord(document.EmployeeFrm.EmployeeNumber2.value, 'JobSwap&JobID1=" & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & "', 'SearchEmployeeNumber2IFrame', 'EmployeeFrm.EmployeeNumber2')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Buscar el número de empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A>&nbsp;"
								Response.Write "</TD>"
								Response.Write "<TD VALIGN=""TOP""><IMG SRC=""Images/Transparent.gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></TD>"
								Response.Write "<TD VALIGN=""TOP"">"
									Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeNumber2IFrame"" FRAMEBORDER=""0"" WIDTH=""300"" HEIGHT=""230""></IFRAME>"
								Response.Write "</TD>"
							Response.Write "</TR>"
						Response.Write "</TABLE>"
					Response.Write "<BR /><BR />"
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",106,", vbBinaryCompare) > 0) And (lReasonID <> 14) Then
					Call DisplayEmployeeFormSection106(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",107,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection107(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",109,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection109(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
				Else
					If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",108,", vbBinaryCompare) > 0) And (lReasonID <> 14) Then
						Call DisplayEmployeeFormSection108(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
					End If
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",136,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection136(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",125,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection125(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",133,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection133(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",129,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection129(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",135,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection135(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",111,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection111(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
				End If
				If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) And (aEmployeeComponent(N_STATUS_ID_EMPLOYEE) <> 0) And (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",131,", vbBinaryCompare) = 0) Then
					If aEmployeeComponent(N_STATUS_REASON_ID_EMPLOYEE) = 2 Or aEmployeeComponent(N_STATUS_REASON_ID_EMPLOYEE) = 3 Then
						If ((lReasonID > 0) Or (lReasonID = -64) Or (lReasonID = -75)) And (lReasonID <> 57) Then
							Response.Write "<FONT COLOR=""D20000"" FACE=""Arial"" SIZE=""2""><B>Razones de rechazo del movimiento</B></FONT>"
								Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
								Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
									Response.Write "<TR>"
									    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><TEXTAREA NAME=""Comments"" ID=""CommentsTxtArea"" ROWS=""5"" COLS=""60"" MAXLENGTH=""2000"" CLASS=""TextFields"">" & aEmployeeComponent(S_REASON_FOR_REJECTION_COMMENTS_EMPLOYEE) & "</TEXTAREA></FONT></TD>"
									    Response.Write "<TD></FONT></TD>"
									Response.Write "</TR>"
								Response.Write "</TABLE>"
								Response.Write "<BR /><BR />"
						End If
					End If
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",110,", vbBinaryCompare) > 0) Then
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Datos de la última plaza que tuvo</B></FONT>"
					Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
					If aEmployeeComponent(N_JOB_ID_EMPLOYEE) <> -1 Then
						Call DisplayJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
					Else
						Response.Write "<FONT FACE=""Arial"" SIZE=""3""><B>No tiene asignada ninguna plaza</B></FONT>"
					End If
					Response.Write "<BR /><BR />"
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",201,", vbBinaryCompare) > 0) Then
					Response.Write "</TD>"
					Response.Write "<TD VALIGN=""TOP"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
					Response.Write "<TD VALIGN=""TOP"">"
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",111,", vbBinaryCompare) > 0) Then
					If Len(sError)> 0 Then
						Response.Write "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
							Response.Write "<TR>"
							    Response.Write "<TD>"
							    Call DisplayErrorMessage("Mensaje del sistema", sError)
							    Response.Write "</TD>"
							Response.Write "</TR>"
						Response.Write "</TABLE>"
						Response.Write "<BR /><BR />"
					End If
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",112,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection112(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",117,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection117(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",113,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection113(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",114,", vbBinaryCompare) > 0) Then
					If Len(aEmployeeComponent(S_REASON_FOR_REJECTION_COMMENTS_EMPLOYEE)) > 0 Then
						Response.Write "<FONT COLOR=""D20000"" FACE=""Arial"" SIZE=""2""><B>Razones de rechazo del movimiento</B></FONT>"
						Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
						Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
							Response.Write "<TR>"
							    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_REASON_FOR_REJECTION_COMMENTS_EMPLOYEE) & "</FONT></TD>"
							    Response.Write "<TD></FONT></TD>"
							Response.Write "</TR>"
						Response.Write "</TABLE>"
					End If
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",115,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection115b(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",116,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection116b(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",118,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection118(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",121a,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection121a(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",121,", vbBinaryCompare) > 0) Then
					Call DisplayEmployeeFormSection121(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
				End If

				Response.Write "<BR />"
				If ((lReasonID > 0) Or (lReasonID = -64) Or (lReasonID = -75)) And (lReasonID <> 57) And (lReasonID <> 58) And (lReasonID <> EMPLOYEES_SERVICE_SHEET) And (aEmployeeComponent(N_ID_EMPLOYEE) > 0) Then
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Historial del empleado</B></FONT>"
					Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
					lErrorNumber = DisplayEmployeeHistoryList(oRequest, oADODBConnection, False, False, aEmployeeComponent, sErrorDescription)
					Response.Write "<BR />"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Historial de la plaza</B></FONT>"
					Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
					aJobComponent(N_ID_JOB) = aEmployeeComponent(N_JOB_ID_EMPLOYEE)
					lErrorNumber = DisplayJobHistoryList(oRequest, oADODBConnection, False, False, aJobComponent, sErrorDescription)
					Response.Write "<BR />"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Historial de ocupación de la plaza</B></FONT>"
					Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
					lErrorNumber = DisplayJobsHistoryListTable(oRequest, oADODBConnection, False, aJobComponent, sErrorDescription)
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Cheques cancelados</B></FONT>"
					Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
					aPaymentComponent(S_SORT_COLUMN_PAYMENT) = "Payments.EmployeeID"
					aPaymentComponent(N_EMPOYEE_ID_PAYMENT) = aEmployeeComponent(N_ID_EMPLOYEE)
					aPaymentComponent(S_QUERY_CONDITION_PAYMENT) =  "(Payments.StatusID Not In (-2,-1,1,2,3,4)) And Payments.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
					lErrorNumber = DisplayPaymentsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, False, False, aPaymentComponent, sErrorDescription)
					Response.Write "<BR /><BR />"
				End If
				If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",201,", vbBinaryCompare) > 0) Then
							Response.Write "</TD>"
						Response.Write "</TR>"
					Response.Write "</TABLE>"
				End If
				If (InStr(1, sURL, ",EmployeeManagement,", vbBinaryCompare) = 0) Then
					If (InStr(1, ",101,", aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), vbBinaryCompare) > 0) Then
						If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""" & sURL & """ ID=""AssignJobBtn"" VALUE=""Buscar Empleado"" CLASS=""Buttons"" />"
					ElseIf (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",119,", vbBinaryCompare) > 0) Then
						Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
						If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""AddBtn"" VALUE=""Modificar"" CLASS=""Buttons"" onClick=""bCheckReason=false""/>"
					ElseIf (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",5,", vbBinaryCompare) > 0) And lReasonID <> 57 Then
						'Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
						If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" onClick=""bCheckReason=false""/>"
					ElseIf (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",132,", vbBinaryCompare) > 0) Then
						Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
						If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" onClick=""bCheckReason=false""/>"
					ElseIf (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",128,", vbBinaryCompare) > 0) Then
						If (lReasonID = EMPLOYEES_SAFE_SEPARATION Or lReasonID=EMPLOYEES_ADD_SAFE_SEPARATION) Then
							If (aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 1) Then
								If False Then
									If lReasonID=EMPLOYEES_ADD_SAFE_SEPARATION Then
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesConceptsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=120) And (EndDate>=" & sDate & ")", "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
										If lErrorNumber = 0 Then
											If Not oRecordset.EOF Then
												If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS And bVisible Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" onClick=""bCheckReason=false""/>"
											End If
										End If
									Else
										If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS And bVisible Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" onClick=""bCheckReason=false""/>"
									End If
								End If
								If Len(oRequest("ModifyConcept").Item) > 0 Then
									If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS And bVisible Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""AddBtn"" VALUE=""Modificar"" CLASS=""Buttons"" onClick=""bCheckReason=false""/>"
								Else
									If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS And bVisible Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" onClick=""bCheckReason=false""/>"
								End If
							End If
						ElseIf (lReasonID = CANCEL_EMPLOYEES_CONCEPTS) Or (lReasonID = CANCEL_EMPLOYEES_C04) Then
							If (Len(oRequest("ModifyConcept").Item) > 0) Then
								If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS And bVisible Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Modificar"" CLASS=""Buttons"" onClick=""bCheckReason=false""/>"
							End If
						ElseIf (lReasonID = EMPLOYEES_SUNDAYS) Or (lReasonID = EMPLOYEES_EXTRAHOURS) Then
							If Len(oRequest("SundayChange").Item) > 0 Then
								If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS And bVisible Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Modificar"" CLASS=""Buttons"" onClick=""bCheckReason=false""/>"
							Else
								If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS And bVisible Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" onClick=""bCheckReason=false""/>"
							End If
						ElseIf (lReasonID = EMPLOYEES_ADD_BENEFICIARIES) Or (lReasonID = EMPLOYEES_CREDITORS) Then
							If Len(oRequest("BeneficiaryChange").Item) > 0 Then
								If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS And bVisible Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Modificar"" CLASS=""Buttons"" onClick=""bCheckReason=false""/>"
							Else
								If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS And bVisible Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" onClick=""bCheckReason=false""/>"
							End If
						ElseIf (lReasonID = EMPLOYEES_SERVICE_SHEET) Then
							If Len(oRequest("ServiceSheetChange").Item) > 0 Then
								If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS And bVisible Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""AddBtn"" VALUE=""Modificar"" CLASS=""Buttons"" onClick=""bCheckReason=false""/>"
							Else
								If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS And bVisible Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" onClick=""bCheckReason=false""/>"
							End If
						Else
							If Len(oRequest("CreditChange").Item) > 0 Then
								If Len(oRequest("ModifyConcept").Item) > 0 Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Payrolls Where (IsClosed<>1) And (IsActive_1=1)", "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
									If lErrorNumber = 0 Then
										If Not oRecordset.EOF Then
											bVisible = True
										Else
											bVisible = False
											Call DisplayErrorMessage("Mensaje del sistema", "No hay quincenas abiertas")
										End If
									End If							
									If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS And bVisible Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Modificar"" CLASS=""Buttons"" onClick=""bCheckReason=false""/>"
								Else
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Payrolls Where (IsClosed<>1) And (IsActive_1=1)", "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
									If lErrorNumber = 0 Then
										If Not oRecordset.EOF Then
											bVisible = True
										Else
											bVisible = False
											Call DisplayErrorMessage("Mensaje del sistema", "No hay quincenas abiertas")
										End If
									End If	
									If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS And bVisible Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Modificar"" CLASS=""Buttons"" onClick=""bCheckReason=false""/>"
								End If
							Else
								If Len(oRequest("ModifyConcept").Item) > 0 Then
									If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS And bVisible Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""AddBtn"" VALUE=""Modificar"" CLASS=""Buttons"" onClick=""bCheckReason=false""/>"
								Else
									If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS And bVisible Then 
										If lReasonID = -96 Then
											Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Cancelar concepto"" CLASS=""Buttons"" onClick=""bCheckReason=false""/>"
										Else
											Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" onClick=""bCheckReason=false""/>"
										End If
									End If
								End If
							End If
						End If
						If ((lReasonID <> CANCEL_EMPLOYEES_CONCEPTS) And (lReasonID <> CANCEL_EMPLOYEES_C04)) Or (Len(oRequest("ModifyConcept").Item) > 0) Then
							Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
						End If
						Select Case lReasonID
							Case -96
								Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""  Salir  "" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=EmployeesMovements&ReasonID="& lReasonID &"'"" />"
							Case EMPLOYEES_SERVICE_SHEET
								Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=ServiceSheet&ReasonID="& lReasonID & "&SectionID=" & iSectionID & "'"" />"
							Case Else
								Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=EmployeesMovements&ReasonID="& lReasonID &"'"" />"
						End Select
						Response.Write "<BR /><BR />"
					ElseIf (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",126,", vbBinaryCompare) > 0) Then
						Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
						If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""AddBtn"" VALUE=""Modificar"" CLASS=""Buttons"" onClick=""bCheckReason=false""/>"
					ElseIf (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",122,", vbBinaryCompare) > 0) Then
						Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Payrolls Where (IsClosed<>1) And (IsActive_7=1) And (PayrollTypeID=1)", "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""SaveEmployeesAdjustments"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" onClick=""bCheckReason=false""/>"
							Else
								sErrorDescription = "No hay quincenas abiertas"
								Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							End If
						End If
					Else
						If bActivate Then
							Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
							If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then
								Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""SaveChanges"" ID=""ModifyBtn"" VALUE=""         Guardar Cambios         "" CLASS=""Buttons"" />"
								Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""1"" />"
							End If
							If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ModificacionDePersonal & ",", vbBinaryCompare) > 0) Or (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_AsignacionDeNumeroDeEmpleado & ",", vbBinaryCompare) > 0) Then
								Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Register"" ID=""ModifyBtn"" VALUE=""Registrar para Validación"" CLASS=""Buttons"" />"
								Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""1"" />"
							End If
							If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ModificacionDePersonal & ",", vbBinaryCompare) > 0) Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""RemoveMotion"" ID=""ModifyBtn""    VALUE=""   Cancelar Movimiento   "" CLASS=""RedButtons"" />"
							Response.Write "<BR /><BR />"
							If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Validate"" ID=""ModifyBtn"" VALUE=""Registrar para Autorización"" CLASS=""Buttons"" />"
							Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""1"" />"
								Response.Write "<DIV NAME=""ApplyMovementDiv"" ID=""ApplyMovementDiv"" STYLE=""display: none"">"
									If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Authorization"" ID=""ModifyBtn"" VALUE=""     Aplicar Movimiento    "" CLASS=""RedButtons"" onClick=""bCheckReason=false""/>"
								Response.Write "</DIV>"
							If (lReasonID > 0) Or (lReasonID = -64) Or (lReasonID = -75) Then
								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
									Response.Write "ShowHideApplyMovementButton(document.EmployeeFrm.EmployeePayrollDate.value);" & vbNewLine
								Response.Write "//--></SCRIPT>" & vbNewLine
							End If
						Else
							Select Case aEmployeeComponent(N_STATUS_REASON_ID_EMPLOYEE)
								Case 1
									'If Not (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then
									'If Not (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then
										If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ModificacionDePersonal & ",", vbBinaryCompare) > 0) Then
											Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Register"" ID=""ModifyBtn"" VALUE=""  Registrar para Validación  "" CLASS=""Buttons"" />"
											Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""1"" />"
										End If
										If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then
											Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""RemoveMotion"" ID=""ModifyBtn""    VALUE=""   Cancelar Movimiento   "" CLASS=""RedButtons"" />"
											Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""1"" />"
										End If
							            'If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Validate"" ID=""ModifyBtn"" VALUE=""Registrar para Autorización"" CLASS=""Buttons"" />"
							            Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""1"" />"
									            Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Authorization"" ID=""ModifyBtn"" VALUE=""     Aplicar Movimiento    "" CLASS=""RedButtons"" onClick=""bCheckReason=false""/>"
                                        'End If
									'End If
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusEmployeesIDs1 From Reasons Where ReasonID=" & lReasonID, "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
									If lErrorNumber = 0 Then
										If Not oRecordset.EOF Then
											sStatusEmployeesIDs1 = CStr(oRecordset.Fields("StatusEmployeesIDs1").Value)
										End If
									End If
									If (InStr(1, sStatusEmployeesIDs1, aEmployeeComponent(N_STATUS_ID_EMPLOYEE), vbBinaryCompare) > 0) Then
										If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_AplicacionDeMovimientos & ",", vbBinaryCompare) > 0) Then
											Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""CancelMotion"" ID=""ModifyBtn"" VALUE=""Rechazar Movimiento"" CLASS=""Buttons"" onClick=""bCheckReason=true""/>"
											Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""1"" />"
										End If
										Response.Write "<DIV NAME=""ApplyMovementDiv"" ID=""ApplyMovementDiv"" STYLE=""display: none"">"
											If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_AplicacionDeMovimientos & ",", vbBinaryCompare) > 0) Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Authorization"" ID=""ModifyBtn"" VALUE=""    Aplicar Movimiento     "" CLASS=""RedButtons"" onClick=""bCheckReason=false""/>"
										Response.Write "</DIV>"
										If lReasonID > 0 Then
											Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
												Response.Write "ShowHideApplyMovementButton(document.EmployeeFrm.EmployeePayrollDate.value);" & vbNewLine
											Response.Write "//--></SCRIPT>" & vbNewLine
										End If
										Response.Write "<BR /><BR />"
									End If
								Case 2
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusEmployeesIDs1 From Reasons Where ReasonID=" & lReasonID, "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
									If lErrorNumber = 0 Then
										If Not oRecordset.EOF Then
											sStatusEmployeesIDs1 = "," & CStr(oRecordset.Fields("StatusEmployeesIDs1").Value) & ","
										End If
									End If
									If (InStr(1, sStatusEmployeeIDs1, aEmployeeComponent(N_STATUS_ID_EMPLOYEE), vbBinaryCompare) > 0) Then
										If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then
											Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""CancelMotion"" ID=""ModifyBtn"" VALUE=""Rechazar Movimiento"" CLASS=""Buttons"" onClick=""bCheckReason=true""/>"
											Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""1"" />"
										End If
										If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then
											Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Validate"" ID=""ModifyBtn"" VALUE=""Registrar para Autorización"" CLASS=""Buttons"" onClick=""bCheckReason=false""/>"
											Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""1"" />"
										End If
										Response.Write "<DIV NAME=""ApplyMovementDiv"" ID=""ApplyMovementDiv"" STYLE=""display: none"">"
											If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Authorization"" ID=""ModifyBtn"" VALUE=""    Aplicar Movimiento     "" CLASS=""RedButtons"" onClick=""bCheckReason=false""/>"
										Response.Write "</DIV>"
										If lReasonID > 0 Then
											Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
												Response.Write "ShowHideApplyMovementButton(document.EmployeeFrm.EmployeePayrollDate.value);" & vbNewLine
											Response.Write "//--></SCRIPT>" & vbNewLine
										End If
										Response.Write "<BR /><BR />"
									End If
								Case 3
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusEmployeesIDs1 From Reasons Where ReasonID=" & lReasonID, "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
									If lErrorNumber = 0 Then
										If Not oRecordset.EOF Then
											sStatusEmployeesIDs1 = CStr(oRecordset.Fields("StatusEmployeesIDs1").Value)
										End If
									End If
									If (InStr(1, sStatusEmployeesIDs1, aEmployeeComponent(N_STATUS_ID_EMPLOYEE), vbBinaryCompare) > 0) Then
										If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then
											Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""CancelMotion"" ID=""ModifyBtn"" VALUE=""Rechazar Movimiento"" CLASS=""Buttons"" onClick=""bCheckReason=true""/>"
											Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""1"" />"
										End If
										Response.Write "<DIV NAME=""ApplyMovementDiv"" ID=""ApplyMovementDiv"" STYLE=""display: none"">"
											If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Authorization"" ID=""ModifyBtn"" VALUE=""    Aplicar Movimiento     "" CLASS=""RedButtons"" onClick=""bCheckReason=false""/>"
										Response.Write "</DIV>"
										If lReasonID > 0 Then
											Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
												Response.Write "ShowHideApplyMovementButton(document.EmployeeFrm.EmployeePayrollDate.value);" & vbNewLine
											Response.Write "//--></SCRIPT>" & vbNewLine
										End If
										Response.Write "<BR /><BR />"
									End If
							End Select
						End If
					End If
				End If
			Response.Write "</FORM>"
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Select Case lReasonID
					Case EMPLOYEES_ADD_BENEFICIARIES
						If aEmployeeComponent(N_ID_EMPLOYEE) <> -1 Then
							Response.Write "ShowAmountFields(document.EmployeeBeneficiaryFrm.AlimonyTypeID.value, 'Concept');" & vbNewLine
						End If
					Case EMPLOYEES_CREDITORS
						If aEmployeeComponent(N_ID_EMPLOYEE) <> -1 Then
							Response.Write "ShowAmountFields(document.EmployeeBeneficiaryFrm.CreditorTypeID.value, 'Concept');" & vbNewLine
						End If
				End Select
			Response.Write "//--></SCRIPT>" & vbNewLine
		End If
	End If

	DisplayEmployeeForm = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeeFormSection03()
'************************************************************
'Purpose: To add JavaScript to validate amounts taken
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection03"

	Select Case lReasonID
		Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS, EMPLOYEES_NIGHTSHIFTS
			Response.Write "SelectAllItemsFromList(oForm.OcurredDates);" & vbNewLine
		Case Else
	End Select
	Response.Write "if (oForm) {" & vbNewLine
		Select Case lReasonID
			Case EMPLOYEES_BANK_ACCOUNTS, CANCEL_EMPLOYEES_CONCEPTS, CANCEL_EMPLOYEES_C04, EMPLOYEES_GRADE, EMPLOYEES_SERVICE_SHEET, CANCEL_EMPLOYEES_SSI
			Case Else
				Response.Write "oForm.ConceptAmount.value = oForm.ConceptAmount.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
				Response.Write "if (! CheckFloatValue(oForm.ConceptAmount, 'el monto del concepto', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "if (parseInt(oForm.EmployeePayrollDate.value)==-1) {" & vbNewLine
					Response.Write "alert('No existen nóminas abiertas para el registro de movimientos.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
		End Select
		Select Case lReasonID
			Case EMPLOYEES_BANK_ACCOUNTS, CANCEL_EMPLOYEES_CONCEPTS, CANCEL_EMPLOYEES_C04, EMPLOYEES_GRADE, EMPLOYEES_SERVICE_SHEET,CANCEL_EMPLOYEES_SSI
			Case EMPLOYEES_SPORTS
				Response.Write "if (oForm.ConceptAmount.value!=0) {" & vbNewLine
					Response.Write "alert('El monto de cuota deportivo debe ser 0.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
			Case EMPLOYEES_GLASSES
				Response.Write "if (oForm.ConceptAmount.value<=0) {" & vbNewLine
					Response.Write "alert('Favor de especificar el monto a pagar.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (oForm.ConceptAmount.value>1800) {" & vbNewLine
					Response.Write "alert('El valor del Importe para ayuda de anteojos no puede ser mayor a 1800 pesos.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
			Case Else
				Response.Write "if (oForm.ConceptAmount.value<=0) {" & vbNewLine
					Response.Write "alert('Favor de especificar el monto a pagar.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
		End Select
		Select Case lReasonID
			Case EMPLOYEES_THIRD_CONCEPT
				If StrComp(oRequest("CreditChange").Item, "1", vbBinaryCompare) <> 0 Then
					Response.Write "if (oForm.PaymentsNumber.value<=0) {" & vbNewLine
						Response.Write "alert('Favor de especificar la cantidad de pagos.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
				End If
			Case EMPLOYEES_BANK_ACCOUNTS
				Response.Write "if (oForm.Cheque.checked) {" & vbNewLine
					Response.Write "oForm.AccountNumber.value='.';" & vbNewLine
				Response.Write "} else {" & vbNewLine
					Response.Write "if (oForm.AccountNumber.value == null || oForm.AccountNumber.value.length == 0 || /^\s+$/.test(oForm.AccountNumber.value)) {" & vbNewLine
						Response.Write "alert('Favor de especificar correctamente el número de cuenta.');" & vbNewLine
						Response.Write "oForm.AccountNumber.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (parseInt(oForm.BankID.value)==-1) {" & vbNewLine
						Response.Write "alert('Seleccione el banco del que se registrará la cuenta.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (!VerifyAccountNumberLength(oForm.AccountNumber.value, oForm.BankID.value)) {" & vbNewLine
						Response.Write "alert('No coincide la longitud de la cuenta para el banco seleccionado.');" & vbNewLine
						Response.Write "oForm.AccountNumber.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (parseInt(oForm.BankID.value)==3) {" & vbNewLine
						Response.Write "if (oForm.Sucursal.value.length > 0 && oForm.Sucursal.value.length != 4) {" & vbNewLine
							Response.Write "alert('Favor de especificar un número de sucursal de 4 dígitos.');" & vbNewLine
							Response.Write "oForm.Sucursal.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (isNaN(parseInt(oForm.Sucursal.value))) {" & vbNewLine
							Response.Write "alert('Favor de especificar un valor númerico para el número de sucursal');" & vbNewLine
							Response.Write "oForm.Sucursal.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (parseInt(oForm.ConceptStartDate.value)==-1) {" & vbNewLine
					Response.Write "alert('No existen nóminas abiertas para el registro de cuentas bancarias.');" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
		End Select
		Select Case lReasonID
			Case -58, 14, EMPLOYEES_GRADE, EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_GLASSES, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_MONTHAWARD, EMPLOYEES_NIGHTSHIFTS, EMPLOYEES_CONCEPT_C3, EMPLOYEES_MOTHERAWARD, EMPLOYEES_ANUAL_AWARD, EMPLOYEES_NIGHTSHIFTS, EMPLOYEES_THIRD_CONCEPT, EMPLOYEES_FONAC_ADJUSTMENT, EMPLOYEES_ANTIQUITY_25_AND_30_YEARS
			Case EMPLOYEES_SERVICE_SHEET
				Response.Write "if (oForm.DocumentNumber1.value == '') {" & vbNewLine
					Response.Write "alert('Favor de introducir el número de documento.');" & vbNewLine
					Response.Write "oForm.DocumentNumber1.focus();" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (oForm.Authorizers.value == '') {" & vbNewLine
					Response.Write "alert('Favor de introducir los usuarios que deben autorizar el documento.');" & vbNewLine
					Response.Write "oForm.Authorizers.focus();" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
			Case EMPLOYEES_THIRD_CONCEPT
				Response.Write "if (((parseInt('1' + oForm.ConceptEndDay.value) - 100) + ('1' + parseInt(oForm.ConceptEndMonth.value) - 100) + parseInt(oForm.ConceptEndYear.value)) > 0 ) {" & vbNewLine
					Response.Write "if ((parseInt('1' + oForm.ConceptEndDay.value) - 100) * (parseInt('1' + oForm.ConceptEndMonth.value) - 100) * parseInt(oForm.ConceptEndYear.value) > 0 ) {" & vbNewLine
						Response.Write "if (parseInt(oForm.EmployeePayrollDate.value) > ((parseInt('1' + oForm.ConceptEndDay.value) - 100) + ((parseInt('1' + oForm.ConceptEndMonth.value) - 100) * 100) + parseInt(oForm.ConceptEndYear.value) * 10000)) {" & vbNewLine
							Response.Write "alert('Favor de verificar la vigencia del registro de crédito del empleado.');" & vbNewLine
							Response.Write "oForm.ConceptEndDay.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "else" & vbNewLine
					Response.Write "{" & vbNewLine
						Response.Write "alert('Favor de verificar la vigencia del movimiento');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
			Case EMPLOYEES_BANK_ACCOUNTS
				Response.Write "if (((parseInt('1' + oForm.ConceptEndDay.value) - 100) + ('1' + parseInt(oForm.ConceptEndMonth.value) - 100) + parseInt(oForm.ConceptEndYear.value)) > 0 ) {" & vbNewLine
					Response.Write "if ((parseInt('1' + oForm.ConceptEndDay.value) - 100) * (parseInt('1' + oForm.ConceptEndMonth.value) - 100) * parseInt(oForm.ConceptEndYear.value) > 0 ) {" & vbNewLine
						Response.Write "if (parseInt(oForm.ConceptStartDate.value) > ((parseInt('1' + oForm.ConceptEndDay.value) - 100) + ((parseInt('1' + oForm.ConceptEndMonth.value) - 100) * 100) + parseInt(oForm.ConceptEndYear.value) * 10000)) {" & vbNewLine
							Response.Write "alert('Favor de verificar la vigencia de los registros de cuentas bancarias');" & vbNewLine
							Response.Write "oForm.ConceptEndDay.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "else" & vbNewLine
					Response.Write "{" & vbNewLine
						Response.Write "alert('Favor de verificar la vigencia del movimiento');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
			Case Else
				Response.Write "if (((parseInt('1' + oForm.ConceptEndDay.value) - 100) + ('1' + parseInt(oForm.ConceptEndMonth.value) - 100) + parseInt(oForm.ConceptEndYear.value)) > 0 ) {" & vbNewLine
					Response.Write "if ((parseInt('1' + oForm.ConceptEndDay.value) - 100) * (parseInt('1' + oForm.ConceptEndMonth.value) - 100) * parseInt(oForm.ConceptEndYear.value) > 0 ) {" & vbNewLine
					If ((lReasonID = CANCEL_EMPLOYEES_CONCEPTS) Or (lReasonID = CANCEL_EMPLOYEES_C04)) And (Len(oRequest("ModifyConcept").Item) > 0) Then
						Response.Write "if (" & CLng(oRequest("ConceptStartDate").Item) & " > parseInt(parseInt(oForm.ConceptEndDay.value) + parseInt(oForm.ConceptEndMonth.value)*100 + parseInt(oForm.ConceptEndYear.value) * 10000)) {" & vbNewLine
							Response.Write "alert('Favor de verificar la vigencia del movimiento');" & vbNewLine
							Response.Write "oForm.ConceptEndDay.focus();" & vbNewLine
					Else
						Response.Write "if (((parseInt('1' + oForm.ConceptStartDay.value) - 100) + ((parseInt('1' + oForm.ConceptStartMonth.value) -100) * 100) + parseInt(oForm.ConceptStartYear.value) * 10000) > ((parseInt('1' + oForm.ConceptEndDay.value) - 100) + ((parseInt('1' + oForm.ConceptEndMonth.value) - 100) * 100) + parseInt(oForm.ConceptEndYear.value) * 10000)) {" & vbNewLine
							Response.Write "alert('Favor de verificar la vigencia del movimiento');" & vbNewLine
							Response.Write "oForm.ConceptStartDay.focus();" & vbNewLine
					End If
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "else" & vbNewLine
					Response.Write "{" & vbNewLine
						Response.Write "alert('Favor de verificar la vigencia del movimiento');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
		End Select
	Response.Write "}" & vbNewLine

End Function

Function DisplayEmployeeFormSection04()
'************************************************************
'Purpose: To add JavaScript to validate amounts taken
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection04"

	Response.Write "if (oForm.EmployeeAccessKey.value == '') {" & vbNewLine
		Response.Write "alert('Favor de introducir la clave de acceso del empleado.');" & vbNewLine
		Response.Write "oForm.EmployeeAccessKey.focus();" & vbNewLine
		Response.Write "return false;" & vbNewLine
	Response.Write "}" & vbNewLine
	Response.Write "if (oForm.AccessKeyChecked.value == '') {" & vbNewLine
		Response.Write "alert('Revise la disponibilidad de la clave de acceso para el empleado.');" & vbNewLine
		Response.Write "oForm.EmployeeAccessKey.focus();" & vbNewLine
		Response.Write "return false;" & vbNewLine
	Response.Write "}" & vbNewLine
	Response.Write "if (oForm.AccessKeyChecked.value == '0') {" & vbNewLine
		Response.Write "alert('La clave de acceso especificada ya está asignada a otro empleado. Favor de introducir otra.');" & vbNewLine
		Response.Write "oForm.EmployeeAccessKey.focus();" & vbNewLine
		Response.Write "return false;" & vbNewLine
	Response.Write "}" & vbNewLine
	Response.Write "if (oForm.EmployeePassword.value == '') {" & vbNewLine
		Response.Write "alert('Favor de introducir la contraseña del empleado.');" & vbNewLine
		Response.Write "oForm.EmployeePassword.focus();" & vbNewLine
		Response.Write "return false;" & vbNewLine
	Response.Write "}" & vbNewLine
	Response.Write "if (oForm.EmployeePassword.value != oForm.PasswordConfirmation.value) {" & vbNewLine
		Response.Write "alert('La contraseña y su confirmación no coinciden. Favor de introducirlas nuevamente.');" & vbNewLine
		Response.Write "oForm.EmployeePassword.value = '';" & vbNewLine
		Response.Write "oForm.PasswordConfirmation.value = '';" & vbNewLine
		Response.Write "oForm.EmployeePassword.focus();" & vbNewLine
		Response.Write "return false;" & vbNewLine
	Response.Write "}" & vbNewLine

End Function

Function DisplayEmployeeFormSection05()
'************************************************************
'Purpose: To add JavaScript to validate amounts taken
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection05"

	Response.Write "function GetEmployeeNumber(sEmployeeTypeID) {" & vbNewLine
		Response.Write "var bFound = false;" & vbNewLine
		Response.Write "if (sEmployeeTypeID == 6) {" & vbNewLine
			Response.Write "i=1;" & vbNewLine
			Response.Write "bFound = true;" & vbNewLine
		Response.Write "}" & vbNewLine
		Response.Write "else" & vbNewLine
			Response.Write "for (var i=0; i<asConsecutiveIDs.length-1; i++)" & vbNewLine
				Response.Write "if (asConsecutiveIDs[i][0] == sEmployeeTypeID) {" & vbNewLine
					Response.Write "bFound = true;" & vbNewLine
					Response.Write "break;" & vbNewLine
				Response.Write "}" & vbNewLine

		Response.Write "if (! bFound)" & vbNewLine
			Response.Write "i=0;" & vbNewLine

		Response.Write "document.EmployeeFrm.EmployeeID.value = parseInt(asConsecutiveIDs[i][1]) + 1;" & vbNewLine
		Response.Write "document.EmployeeFrm.EmployeeNumber.value = parseInt(asConsecutiveIDs[i][1])+ 1;" & vbNewLine
		Response.Write "document.EmployeeFrm.ConsecutiveID.value = '';" & vbNewLine
		Response.Write "if (asConsecutiveIDs[i][1] != '')" & vbNewLine
			Response.Write "document.EmployeeFrm.ConsecutiveID.value = asConsecutiveIDs[i][0];" & vbNewLine
	Response.Write "} // End of GetEmployeeNumber" & vbNewLine

	Response.Write " function ShowEmployeeField(sField, bShow) {" & vbNewLine
		Response.Write "var sForm = 'document.EmployeeFrm.';" & vbNewLine
		Response.Write "var oField = eval(sForm + sField);" & vbNewLine
		Response.Write "if (oField) {" & vbNewLine
			Response.Write "if (bShow) {" & vbNewLine
				Response.Write "oField.value = '';" & vbNewLine
				Response.Write "ShowDisplay(document.all['Has' + sField + 'Div']);" & vbNewLine
				Response.Write "HideDisplay(document.all['NotHas' + sField + 'Div']);" & vbNewLine
			Response.Write "} else {" & vbNewLine
				Response.Write "oField.value = '-1';" & vbNewLine
				Response.Write "HideDisplay(document.all['Has' + sField + 'Div']);" & vbNewLine
				Response.Write "ShowDisplay(document.all['NotHas' + sField + 'Div']);" & vbNewLine
			Response.Write "}" & vbNewLine
		Response.Write "}" & vbNewLine
	Response.Write "} // End of ShowEmployeeField" & vbNewLine
End Function

Function DisplayEmployeeFormSection07()
'************************************************************
'Purpose: To add JavaScript to validate amounts taken
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection07"
	Dim sAlimonyTypeIDs

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
			If False Then
				Response.Write "if (oForm.BeneficiaryNumber.value.length == 0) {" & vbNewLine
					Response.Write "alert('Favor de especificar el número del beneficiario.');" & vbNewLine
					Response.Write "oForm.BeneficiaryNumber.focus();" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
			End If
			Select Case lReasonID
				Case EMPLOYEES_ADD_BENEFICIARIES
					Response.Write "if (oForm.BeneficiaryNumber.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de especificar el número del beneficiario.');" & vbNewLine
						Response.Write "oForm.BeneficiaryNumber.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (!CheckIntegerValue(oForm.BeneficiaryNumber, 'el número de beneficiario de pensión alimenticia', N_MINIMUM_ONLY_FLAG, N_MAXIMUM_OPEN_FLAG, 0, 100)) {" & vbNewLine
						Response.Write "oForm.BeneficiaryNumber.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"  & vbNewLine
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
					Response.Write "if (parseInt(oForm.BeneficiaryPaymentCenterID.value)==-1) {" & vbNewLine
						Response.Write "alert('Seleccione el centro de pago en al que se enviara la pensión.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (oForm.ConceptAmount.value==0) {" & vbNewLine
						Response.Write "alert('Ingrese un monto válido para la pensión.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"  & vbNewLine
					Response.Write "else {" & vbNewLine
						Response.Write "if (IsAlimonyTypeForPercent(oForm.AlimonyTypeID.value)) {" & vbNewLine
							Response.Write "if (!CheckIntegerValue(oForm.ConceptAmount, 'el porcentaje de la pensión alimenticia', N_BOTH_FLAG, N_CLOSED_FLAG, 1, 100)) {" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
							Response.Write "else {" & vbNewLine
								Response.Write "if (!VerifyTotalAmountForPercent(oForm.AlimonyTypeID.value, oForm.ConceptAmount.value) && bCheckAmount) {" & vbNewLine
									Response.Write "alert('El porcentaje total de este tipo de pensión, registradas para el empleado, más ' + oForm.ConceptAmount.value + '% superan el 100%. No es posible registrarla con este monto.');" & vbNewLine
									Response.Write "return false;" & vbNewLine
								Response.Write "}"  & vbNewLine
							Response.Write "}" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "else {" & vbNewLine
							Response.Write "oForm.ConceptAmount.value = oForm.ConceptAmount.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "if (! CheckFloatValue(oForm.ConceptAmount, 'el monto de la pensión alimenticia', N_MINIMUM_ONLY_FLAG, N_MINIMUM_OPEN_FLAG, 0, 0)) {" & vbNewLine
									Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "}" & vbNewLine
				Case EMPLOYEES_CREDITORS
					Response.Write "if (oForm.CreditorNumber.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de especificar el número del acreedor.');" & vbNewLine
						Response.Write "oForm.CreditorNumber.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (oForm.CreditorNumber.value.length < 7) {" & vbNewLine
						Response.Write "alert('El número del beneficiario debe de ser por lo menos de 7 dígitos.');" & vbNewLine
						Response.Write "oForm.CreditorNumber.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (!CheckIntegerValue(oForm.CreditorNumber, 'el número del acreedor', N_MINIMUM_ONLY_FLAG, N_MAXIMUM_OPEN_FLAG, 0, 100)) {" & vbNewLine
						Response.Write "oForm.CreditorNumber.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"  & vbNewLine
					Response.Write "if (oForm.CreditorName.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de especificar el nombre del beneficiario.');" & vbNewLine
						Response.Write "oForm.CreditorName.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (oForm.CreditorLastName.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de especificar el apellido paterno.');" & vbNewLine
						Response.Write "oForm.CreditorLastName.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (parseInt(oForm.CreditorPaymentCenterID.value)==-1) {" & vbNewLine
						Response.Write "alert('Seleccione el centro de pago en al que se enviara la pensión.');" & vbNewLine
						Response.Write "oForm.CreditorPaymentCenterID.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (oForm.ConceptAmount.value==0) {" & vbNewLine
						Response.Write "alert('Ingrese un monto válido para la pensión.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}"  & vbNewLine
					Response.Write "else {" & vbNewLine
						Response.Write "if (IsAlimonyTypeForPercent(oForm.CreditorTypeID.value)) {" & vbNewLine
							Response.Write "if (!CheckIntegerValue(oForm.ConceptAmount, 'el porcentaje de la pensión alimenticia', N_BOTH_FLAG, N_CLOSED_FLAG, 1, 100)) {" & vbNewLine
									Response.Write "return false;" & vbNewLine
								Response.Write "}"  & vbNewLine
							Response.Write "else {" & vbNewLine
								Response.Write "if (!VerifyTotalAmountForPercent(oForm.CreditorTypeID.value, oForm.ConceptAmount.value) && bCheckAmount) {" & vbNewLine
									Response.Write "alert('El porcentaje total de este tipo de pensión, registradas para el empleado, más ' + oForm.ConceptAmount.value + '% superan el 100%. No es posible registrarla con este monto.');" & vbNewLine
									Response.Write "return false;" & vbNewLine
								Response.Write "}"  & vbNewLine
							Response.Write "}" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "else {" & vbNewLine
							Response.Write "oForm.ConceptAmount.value = oForm.ConceptAmount.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "if (! CheckFloatValue(oForm.ConceptAmount, 'el monto de la pensión alimenticia', N_MINIMUM_ONLY_FLAG, N_MINIMUM_OPEN_FLAG, 0, 0)) {" & vbNewLine
									Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "}" & vbNewLine
			End Select
		Response.Write "} // End of  if (oForm)" & vbNewLine
		Response.Write "return true;" & vbNewLine
	Response.Write "} // End of CheckEmployeeBeneficiaryFields" & vbNewLine
	Response.Write vbNewLine
	Response.Write "function IsAlimonyTypeForPercent(sValue) {" & vbNewLine
		'Response.Write "alert('function IsAlimonyTypeForPercent().' + sValue);" & vbNewLine
		Select Case lReasonID
			Case EMPLOYEES_ADD_BENEFICIARIES
				lErrorNumber = GetAlimonyTypesForPercent(sAlimonyTypeIDs, sErrorDescription)
			Case EMPLOYEES_CREDITORS
				lErrorNumber = GetCreditorTypesForPercent(sAlimonyTypeIDs, sErrorDescription)
		End Select
		If (lErrorNumber = L_ERR_NO_RECORDS) Then
			Response.Write "return false;" & vbNewLine
			sErrorDescription = ""
			lErrorNumber = 0
		Else
			sDisplayFormCaseOptions = Split(sAlimonyTypeIDs, "," , -1, vbBinaryCompare)
			Response.Write "switch (sValue) {" & vbNewLine
				For iIndex = 0 To UBound(sDisplayFormCaseOptions)
					Response.Write "case '" & CInt(sDisplayFormCaseOptions(iIndex)) & "':" & vbNewLine
				Next
					Response.Write "return true;" & vbNewLine
				Response.Write "default:" & vbNewLine
					Response.Write "return false;" & vbNewLine
			Response.Write "}" & vbNewLine
		End If
	Response.Write "} // End of IsAlimonyTypeForPercent" & vbNewLine
	Response.Write "function VerifyTotalAmountForPercent(lFlag, sValue) {" & vbNewLine
		Dim bExistRecords
		Dim lTotalAmount
		Dim asAlimonys
		Dim asTotalForAlimonys

		Select Case lReasonID
			Case EMPLOYEES_ADD_BENEFICIARIES
				lErrorNumber = GetAlimonyTypesTotalAmountForEmployee(lTotalAmount, sErrorDescription)
			Case EMPLOYEES_CREDITORS
				lErrorNumber = GetCreditorTypesTotalAmountForEmployee(lTotalAmount, sErrorDescription)
		End Select
		If (lErrorNumber = L_ERR_NO_RECORDS) Then
			Response.Write "return true;" & vbNewLine
			sErrorDescription = ""
			lErrorNumber = 0
		Else
			asAlimonys = Split(lTotalAmount, SECOND_LIST_SEPARATOR)
			Response.Write "switch (lFlag) {" & vbNewLine
			'Response.Write "alert('Ingresa a switch: ' + lFlag);" & vbNewLine
			For iIndex = 0 To UBound(asAlimonys) - 1
				asTotalForAlimonys = Split(asAlimonys(iIndex), LIST_SEPARATOR)
				Response.Write "case '" & asTotalForAlimonys(0) & "':" & vbNewLine
					Response.Write "if ((parseInt(sValue)+" & asTotalForAlimonys(1) & ")>100) {" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "else {" & vbNewLine
						Response.Write "return true;" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "break;" & vbNewLine
			Next
			Response.Write "default:" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of switch" & vbNewLine
		End If
	Response.Write "} // End of VerifyTotalAmountForPercent" & vbNewLine

End Function

Function DisplayEmployeeFormSection135(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 135
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection135"

	Dim sSucursal
	Call GetNameFromTable(oADODBConnection, "EmployeeAccount", aEmployeeComponent(N_ID_EMPLOYEE), "", "", sAccountNumber, sErrorDescription)
	sAccountNumber = Split(sAccountNumber, LIST_SEPARATOR)
	aEmployeeComponent(S_ACCOUNT_NUMBER_EMPLOYEE) = sAccountNumber(0)
	sSucursal = aEmployeeComponent(S_ACCOUNT_NUMBER_EMPLOYEE)(1)
	Response.Write "<SPAN NAME=""AccountNumberSpn"" ID=""AccountNumberSpn"""
		If StrComp(aEmployeeComponent(S_ACCOUNT_NUMBER_EMPLOYEE), ".", vbBinaryCompare) = 0 Then Response.Write "STYLE=""display: none"""
	Response.Write ">"
		Response.Write "Número de cuenta: <INPUT TYPE=""TEXT"" NAME=""AccountNumber"" ID=""AccountNumberTxt"""
			If Len(aEmployeeComponent(S_ACCOUNT_NUMBER_EMPLOYEE)) > 0 Then Response.Write "VALUE=""" & aEmployeeComponent(S_ACCOUNT_NUMBER_EMPLOYEE)
		Response.Write """ SIZE=""20"" MAXLENGTH=""20"" CLASS=""TextFields""/>"
	Response.Write "</SPAN>"
	Response.Write "<BR />"
	Response.Write "<SPAN NAME=""SucursalSpn"" ID=""SucursalSpn"" STYLE=""display: none"">"
		Response.Write "Sucursal: <INPUT TYPE=""TEXT"" NAME=""Sucursal"" ID=""SucursalTxt"" VALUE=""" & sSucursal & """ SIZE=""4"" MAXLENGTH=""4"" CLASS=""TextFields"" />"
	Response.Write "</SPAN>"
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Banco:&nbsp;</FONT></TD>"
			Response.Write "<TD><SELECT NAME=""BankID"" ID=""BankIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowAmountFields(this.value, 'SucursalSpn')"">"
				Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Banks", "BankID", "BankName", "Active = 1", "BankID", aEmployeeComponent(N_BANK_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
			Response.Write "</SELECT></TD>"
		Response.Write "</TR>"
		If Len(oRequest("BankAccountChange").Item) > 0 Then
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), -1, -1, -1) & "</FONT></TD>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptStartDate"" ID=""ConceptStartDateHdn"" VALUE=""" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & """ />"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de fin:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
				If aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 30000000 Then
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Indeterminada</FONT></TD>"
				Else
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE), -1, -1, -1) & "</FONT></TD>"
				End If
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptEndDate"" ID=""ConceptEndDateHdn"" VALUE=""" & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & """ />"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Aplica como:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
					Response.Write "<INPUT TYPE=""CHECKBOX"" NAME=""Cheque"" ID=""ChequeChk"" VALUE=""1"" onClick=""if (this.checked) {HideDisplay(document.all['AccountNumberSpn']); HideDisplay(document.all['SucursalSpn']);} else {ShowDisplay(document.all['AccountNumberSpn']); ShowAmountFields(document.EmployeeFrm.BankID.value, 'SucursalSpn');}"""
						If StrComp(aEmployeeComponent(S_ACCOUNT_NUMBER_EMPLOYEE), ".", vbBinaryCompare) = 0 Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write "> Cheque"
				Response.Write "</FONT></TD>"
			Response.Write "</TR>"
		Else
			Response.Write "<TR NAME=""ConceptStartDateDiv"" ID=""ConceptStartDateDiv"">"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Quincena de aplicación:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""ConceptStartDate"" ID=""ConceptStartDate"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And (IsActive_4=1) And (PayrollTypeID=1)", "PayrollID Desc", "", "No existen nóminas abiertas para el registro de cuentas;;;-1", sErrorDescription)
				Response.Write "</SELECT>&nbsp;"
				Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
			'	Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de fin:&nbsp;</FONT></TD>"
			'	Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE), "ConceptEnd", Year(Date())-1, Year(Date())+1, True, True) & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Aplica como:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
					Response.Write "<INPUT TYPE=""CHECKBOX"" NAME=""Cheque"" ID=""ChequeChk"" VALUE=""1"" onClick=""if (this.checked) {HideDisplay(document.all['AccountNumberSpn']); HideDisplay(document.all['SucursalSpn']);} else {ShowDisplay(document.all['AccountNumberSpn']); ShowAmountFields(document.EmployeeFrm.BankID.value, 'SucursalSpn');}"""
						If StrComp(aEmployeeComponent(S_ACCOUNT_NUMBER_EMPLOYEE), ".", vbBinaryCompare) = 0 Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write ">Cheque"
				Response.Write "</FONT></TD>"
			Response.Write "</TR>"
		End If
	Response.Write "</TABLE>"
End Function

Function DisplayEmployeeFormSection102(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 102
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection102"

	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Introduzca la información del reclamo:</B></FONT>"
	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		Response.Write "<TR NAME=""ConceptIDDiv"" ID=""ConceptIDDiv"">"
			Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Concepto:&nbsp;</FONT></TD>"
			Response.Write "<TD>"
				If aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1 Then
					Response.Write "<SELECT NAME=""ConceptID"" ID=""ConceptIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID>0)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT>"
				Else
					If aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 7 or aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 12 or aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 13 Then
						Response.Write "<SELECT NAME=""ConceptID"" ID=""ConceptIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID>0)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT>"
					Else
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptID"" ID=""ConceptIDHdn"" VALUE=""" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & """ />"
						Call GetNameFromTable(oADODBConnection, "Concepts", aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT>"
					End If
				End If
			Response.Write "</TD>"
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de omisión de pago:&nbsp;</FONT></TD>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_MISSING_DATE_EMPLOYEE), "Missing", Year(Date())-10, Year(Date()), True, True) & "</FONT></TD>"
		Response.Write "</TR>"
		Response.Write "<TR NAME=""PayrollDateDiv"" ID=""PayrollDateDiv"">"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Quincena de aplicación:&nbsp;</FONT></TD>"
			Response.Write "<TD><SELECT NAME=""EmployeePayrollDate"" ID=""EmployeePayrollDateCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""CheckStatusPayrolls()"">"
				If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And (IsActive_1=1) And (PayrollTypeID=1)", "PayrollID Desc", aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
				ElseIf (CInt(Request.Cookies("SIAP_SectionID")) = 2) Or (CInt(Request.Cookies("SIAP_SectionID")) = 7) Then
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And (IsActive_7=1) And (PayrollTypeID=1)", "PayrollID Desc", aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
				ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And (PayrollTypeID=1)", "PayrollID Desc", aPayrollRevisionComponent(N_DATE_PAYROLL_REVISION), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
				End If
			Response.Write "</SELECT></TD>"
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Monto:&nbsp;</FONT></TD>"
			Response.Write "<TD>"
					Response.Write "<INPUT TYPE=""TEXT"" NAME=""ConceptAmount"" ID=""ConceptAmountTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & FormatNumber(aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE), 2, True, False, True) & """ CLASS=""TextFields"" />&nbsp;"
					Response.Write "<SELECT NAME=""ConceptQttyID"" ID=""ConceptQttyIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowAmountFields(this.value, 'Concept');"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "QttyID=1", "QttyID", aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT>&nbsp;"
					Response.Write "<SPAN ID=""ConceptCurrencySpn""><SELECT NAME=""CurrencyID"" ID=""CurrencyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Currencies", "CurrencyID", "CurrencyName", "(CurrencyID=0) And (Active=1)", "CurrencyName", aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE), "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></SPAN>"
					Response.Write "<SPAN ID=""ConceptAppliesToSpn"" STYLE=""display: none""></SPAN>"
			Response.Write "</TD>"
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre del beneficiario:&nbsp;</FONT></TD>"
			Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""BeneficiaryName"" ID=""BeneficiaryNameTxt"" VALUE=""" & aEmployeeComponent(S_NAME_BENEFICIARY_EMPLOYEE) & """ SIZE=""50"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
		Response.Write "</TR>"
	Response.Write "</TABLE>"
End Function

Function DisplayEmployeeFormSection104(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 104
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection104"
	Dim oRecordset

	If Len(oRequest("ReasonID").Item) > 0 Then
		Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Información general</B></FONT>"
		Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""400"">&nbsp;</TD>"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""32""><A HREF=""UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & """><IMG SRC=""Images/MnLeftArrows.gif"" WIDTH=""32"" HEIGHT=""32"" ALT=""Prestación"" BORDER=""0"" /></A><BR /></TD>"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""290""><FONT FACE=""Arial"" SIZE=""2""><B><A HREF=""UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & lReasonID & """ CLASS=""SpecialLink"">Otro empleado</A></B><BR /></FONT>"
				'Response.Write "<DIV CLASS=""MenuOverflow""><FONT FACE=""Arial"" SIZE=""2"">Registre el movimiento para otro empleado.</FONT></DIV></TD>"
				Response.Write "<DIV CLASS=""MenuOverflow""><FONT FACE=""Arial"" SIZE=""2"">Registre el movimiento a un empleado diferente.</FONT></DIV></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	End If
	If (aEmployeeComponent(N_JOB_ID_EMPLOYEE) = -3) Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JobId From JobsHistoryList Where EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & " Order By EndDate Desc", "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		aJobComponent(N_ID_JOB) = oRecordset.Fields("JobID").Value
		aEmployeeComponent(N_JOB_ID_EMPLOYEE) = aJobComponent(N_ID_JOB)
	End If
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del empleado:&nbsp;</FONT></TD>"
			If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" VALUE=""" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""SpecialTextFields"" onFocus=""document.EmployeeFrm.EmployeeName.focus()"" /></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "</FONT><INPUT TYPE=""HIDDEN"" NAME=""EmployeeNumber"" ID=""EmployeeNumberHdn"" VALUE=""" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & """ /></TD>"
			End If
		Response.Write "</TR>"
		'Response.Write "<TR>"
		'	Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Clave de acceso:&nbsp;</FONT></TD>"
		'	Response.Write "<TD>"
		'		If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
		'			Response.Write "<INPUT TYPE=""TEXT"" NAME=""EmployeeAccessKey"" ID=""EmployeeAccessKeyTxt"" VALUE=""" & aEmployeeComponent(S_ACCESS_KEY_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""120"" CLASS=""TextFields"" />"
		'			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AccessKeyChecked"" ID=""AccessKeyCheckedHdn"" VALUE="""" />"
		'			Response.Write "<A HREF=""javascript: SearchRecord(document.EmployeeFrm.EmployeeAccessKey.value, 'EmployeeAccessKey', 'SearchEmployeeAccessKeyIFrame', 'EmployeeFrm.AccessKeyChecked')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Revisar disponibilidad de la clave de acceso"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A>&nbsp;"
		'			Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeAccessKeyIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""22""></IFRAME>"
		'		Else
		'			Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aEmployeeComponent(S_ACCESS_KEY_EMPLOYEE)) & "</FONT>"
		'			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeAccessKey"" ID=""EmployeeAccessKeyHdn"" VALUE=""" & aEmployeeComponent(S_ACCESS_KEY_EMPLOYEE) & """ />"
		'			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AccessKeyChecked"" ID=""AccessKeyCheckedHdn"" VALUE=""1"" />"
		'		End If
		'	Response.Write "</TD>"
		'Response.Write "</TR>"
		'Response.Write "<TR>"
		'	Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Contraseña:&nbsp;</FONT></TD>"
		'	Response.Write "<TD><INPUT TYPE=""PASSWORD"" NAME=""EmployeePassword"" ID=""EmployeePasswordPwd"" VALUE=""" & aEmployeeComponent(S_PASSWORD_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""120"" CLASS=""TextFields"" /></TD>"
		'Response.Write "</TR>"
		'Response.Write "<TR>"
		'	Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Confirmación:&nbsp;</FONT></TD>"
		'	Response.Write "<TD><INPUT TYPE=""PASSWORD"" NAME=""PasswordConfirmation"" ID=""PasswordConfirmationPwd"" VALUE=""" & aEmployeeComponent(S_PASSWORD_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""120"" CLASS=""TextFields"" /></TD>"
		'Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre(s):&nbsp;</FONT></TD>"
			Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeName"" ID=""EmployeeNameTxt"" VALUE=""" & aEmployeeComponent(S_NAME_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100""" & sReadOnly & "CLASS=""TextFields"" /></TD>"
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido paterno:&nbsp;</FONT></TD>"
			Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName"" ID=""EmployeeLastNameTxt"" VALUE=""" & aEmployeeComponent(S_LAST_NAME_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100""" & sReadOnly & "CLASS=""TextFields"" /></TD>"
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido materno:&nbsp;</FONT></TD>"
			Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName2"" ID=""EmployeeLastName2Txt"" VALUE=""" & aEmployeeComponent(S_LAST_NAME2_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100""" & sReadOnly & "CLASS=""TextFields"" /></TD>"
		Response.Write "</TR>"
	Response.Write "</TABLE>"
	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		If (aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 7 or aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 12 or aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 13) Then
			lErrorNumbeer = GetEmployeeSalary(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Salario mensual:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE)*2) & "</FONT></TD>"
			Response.Write "</TR>"
		End If
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de tabulador:&nbsp;</FONT></TD>"
			If Len(oRequest("Tab").Item) = 0 Then
				If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = -1) Then
					Response.Write "<TD><SELECT NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDCmb"" SIZE=""1"" CLASS=""Lists"""
					Response.Write " onChange=""GetEmployeeNumber(this.value)"""
					Response.Write ">"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EmployeeTypes", "EmployeeTypeID", "EmployeeTypeID As RecordID, EmployeeTypeName", "(Active=1)", "EmployeeTypeName", aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Else
'									Response.Write ">"
'									If (aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = 106) Or (aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = 156) Then
'										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EmployeeTypes", "EmployeeTypeID", "EmployeeTypeID As RecordID, EmployeeTypeName", "(Active=1) And (EmployeeTypeID In (0,1,2,3,4))", "EmployeeTypeName", aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
'									Else
'										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EmployeeTypes", "EmployeeTypeID", "EmployeeTypeID As RecordID, EmployeeTypeName", "(Active=1) And (EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ")", "EmployeeTypeName", aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
'									End If
					Call GetNameFromTable(oADODBConnection, "EmployeeTypes", aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & " " & sNames & "</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""HIDDEN"" NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDTxt"" VALUE=""" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & """ CLASS=""TextFields"" />"
				End If
			Else
				Call GetNameFromTable(oADODBConnection, "EmployeeTypes", aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & " " & sNames & "</FONT></TD>"
			End If
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Antigüedad ISSSTE:&nbsp;</FONT></TD>"
			If (CInt(aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE)) <> 6) And (CInt(aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE)) <> 7 And CInt(aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE)) <> 12 And CInt(aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE)) <> 13) Then
				lErrorNumber = CalculateEmployeeAntiquity(oADODBConnection, aEmployeeComponent, 0, sEmployeeDisplayFormAntiquity, lDisplayFormAntiquityYears, lDisplayFormAntiquityMonths, lDisplayFormAntiquityDays, sErrorDescription)
			Else
				lDisplayFormAntiquityYears = 0
				lDisplayFormAntiquityMonths = 0
				lDisplayFormAntiquityDays = 0
				sEmployeeDisplayFormAntiquity = "0 Años 0 Meses 0 Días"
			End If
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & sEmployeeDisplayFormAntiquity & "</TD>"
			Response.Write "<TD><INPUT TYPE=""HIDDEN"" NAME=""AntiquityYears"" ID=""AntiquityYearsTxt"" VALUE=""" & lDisplayFormAntiquityYears & """ CLASS=""TextFields"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AntiquityMonths"" ID=""AntiquityMonthsTxt"" VALUE=""" & lDisplayFormAntiquityMonths & """ CLASS=""TextFields"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AntiquityDays"" ID=""AntiquityDaysTxt"" VALUE=""" & lDisplayFormAntiquityDays & """ CLASS=""TextFields"" />"
			Response.Write "</TD>"
		Response.Write "</TR>"						
	Response.Write "</TABLE>"
	If (aEmployeeComponent(N_JOB_ID_EMPLOYEE) <> -1) And (bVisible = True) Then
		Response.Write "<BR />"
		Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Información de la plaza:</B></FONT>"
		Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Plaza:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & Right("000000" & aEmployeeComponent(N_JOB_ID_EMPLOYEE), Len("000000")) & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Empresa:&nbsp;</FONT></TD>"
				Call GetNameFromTable(oADODBConnection, "Companies", aJobComponent(N_COMPANY_ID_JOB), "", "", sNames, sErrorDescription)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & sNames & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Puesto:&nbsp;</FONT></TD>"
				Call GetNameFromTable(oADODBConnection, "Positions", aJobComponent(N_POSITION_ID_JOB), "", "", sNames, sErrorDescription)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & sNames & "</FONT></TD>"
			Response.Write "</TR>"
			If aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE) > -1 Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Clasificación:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE) & "</FONT></TD>"
				Response.Write "</TR>"
			End If
			If aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE) > -1 Then
				Response.Write "<TR>"
					Call GetNameFromTable(oADODBConnection, "GroupGradeLevels", aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Grupo, grado, nivel:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
			End If
			If aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE) > -1 Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Integración:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE)) & "</FONT></TD>"
				Response.Write "</TR>"
			End If
			Response.Write "<TR>"
				Call GetNameFromTable(oADODBConnection, "Levels", aEmployeeComponent(N_LEVEL_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nivel:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				'Call GetNameFromTable(oADODBConnection, "GroupGradeLevels", aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Jornada:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aJobComponent(D_WORKING_HOURS_JOB) & " horas</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de puesto:&nbsp;</FONT></TD>"
				Call GetNameFromTable(oADODBConnection, "PositionTypes", aJobComponent(N_POSITION_TYPE_ID_JOB), "", "", sNames, sErrorDescription)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & sNames & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Centro&nbsp;de&nbsp;Trabajo:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
				Call GetNameFromTable(oADODBConnection, "Areas", aJobComponent(N_AREA_ID_JOB), "", "", sNames, sErrorDescription)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & sNames & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Centro&nbsp;de&nbsp;pago:&nbsp;</FONT></TD>"
				Call GetNameFromTable(oADODBConnection, "Areas", aJobComponent(N_PAYMENT_CENTER_ID_JOB), "", "", sNames, sErrorDescription)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & sNames & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Servicio:&nbsp;</FONT></TD>"
				Call GetNameFromTable(oADODBConnection, "Services", aJobComponent(N_SERVICE_ID_JOB), "", "", sNames, sErrorDescription)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & sNames & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Turno:&nbsp;</FONT></TD>"
				Call GetNameFromTable(oADODBConnection, "Journeys", aJobComponent(N_JOURNEY_ID_JOB), "", "", sNames, sErrorDescription)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & sNames & "</FONT></TD>"
			Response.Write "</TR>"
'							Response.Write "<TR>"
'								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Horarios:&nbsp;</FONT></TD>"
'								Call GetNameFromTable(oADODBConnection, "Shifts", aJobComponent(N_SHIFT_ID_JOB), "", "", sNames, sErrorDescription)
'								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & sNames & "</FONT></TD>"
'							Response.Write "</TR>"
		Response.Write "</TABLE>"
		Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Vigencia de la plaza:</B></FONT>"
		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(aJobComponent(N_START_DATE_JOB))) & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de fin:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
				If CLng(aJobComponent(N_END_DATE_JOB)) = 30000000 Then
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Indefinida</FONT></TD>"
				Else
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(aJobComponent(N_END_DATE_JOB))) & "</FONT></TD>"
				End If
			Response.Write "</TR>"
		Response.Write "</TABLE>"
		Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Vigencia del último movimiento de la plaza:</B></FONT>"
		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"	
			Response.Write "<TR>"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JobDate, EndDate From JobsHistoryList Where JobID =" & aJobComponent(N_ID_JOB) & " order by 1 Desc", "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("JobDate").Value)) & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de fin:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
				If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Indefinida</FONT></TD>"
				Else
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & "</FONT></TD>"
				End If
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Estatus:&nbsp;</FONT></TD>"
				Call GetNameFromTable(oADODBConnection, "StatusJobs", aJobComponent(N_STATUS_ID_JOB), "", "", sNames, sErrorDescription)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & sNames & "</FONT></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
		'If lReasonID = 10 Then
		'	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
		'	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Historial de la ocupación de la plaza:</B></FONT>"
		'	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
		'	aJobComponent(N_ID_JOB) = aEmployeeComponent(N_JOB_ID_EMPLOYEE)
		'	lErrorNumber = DisplayJobsHistoryListTable(oRequest, oADODBConnection, False, aJobComponent, sErrorDescription)
		'End If
	End If
	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">RFC:&nbsp;</FONT></TD>"
			If bActivate And Not bReadOnly Then
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""RFC"" ID=""RFCTxt"" VALUE=""" & aEmployeeComponent(S_RFC_EMPLOYEE) & """ SIZE=""13"" MAXLENGTH=""13""" & sReadOnly & "CLASS=""TextFields"" /></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_RFC_EMPLOYEE) & "</FONT></TD>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RFC"" ID=""RFCHdn"" VALUE=""" & aEmployeeComponent(S_RFC_EMPLOYEE) & """ />"
			End If
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">CURP:&nbsp;</FONT></TD>"
			If bActivate And Not bReadOnly Then
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""CURP"" ID=""CURPTxt"" VALUE=""" & aEmployeeComponent(S_CURP_EMPLOYEE) & """ SIZE=""18"" MAXLENGTH=""18""" & sReadOnly & "CLASS=""TextFields"" />"
				Response.Write "<A HREF=""javascript: OpenNewWindow('http://consultas.curp.gob.mx/CurpSP/curp2.do?strCurp=' + document.EmployeeFrm.CURP.value + '&strTipo=B&entfija=DF&depfila=09020', null, 'CurpInversa', 400, 600, 'yes', 'yes');""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Consultar CURP"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_CURP_EMPLOYEE) & "</FONT></TD>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CURP"" ID=""CURPHdn"" VALUE=""" & aEmployeeComponent(S_CURP_EMPLOYEE) & """ />"
			End If
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. de seguro social:&nbsp;</FONT></TD>"
			If bActivate And Not bReadOnly Then
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""SocialSecurityNumber"" ID=""SocialSecurityNumberTxt"" VALUE=""" & aEmployeeComponent(S_SSN_EMPLOYEE) & """ SIZE=""11"" MAXLENGTH=""100""" & sReadOnly & "CLASS=""TextFields"" /></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_SSN_EMPLOYEE) & "</FONT></TD>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SocialSecurityNumber"" ID=""SocialSecurityNumberHdn"" VALUE=""" & aEmployeeComponent(S_SSN_EMPLOYEE) & """ />"
			End If
		Response.Write "</TR>"
		Response.Write "<TR NAME=""CountryDiv"" ID=""CountryDiv"">"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nacionalidad:&nbsp;</FONT></TD>"
				If bActivate And Not bReadOnly Then
					Response.Write "<TD><SELECT NAME=""CountryID"" ID=""CountryIDCmb"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Countries", "CountryID", "CountryID As RecordID, Nationality", "(Active=1)", "Nationality", aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Else
					Call GetNameFromTable(oADODBConnection, "Nationalities", aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CountryID"" ID=""CountryIDHdn"" VALUE=""" & aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE) & """ />"
					Response.Write "<TD>"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT>"
					Response.Write "</TD>"
				End If
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de nacimiento:&nbsp;</FONT></TD>"
			If bActivate And Not bReadOnly Then
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE), "Birth", N_FORM_START_YEAR, Year(Date()), True, True) & "</FONT></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE), -1, -1, -1) & "</FONT></TD>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BirthDate"" ID=""BirthDateHdn"" VALUE=""" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) & """ />"
			End If
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Lugar de Nacimiento:&nbsp;</FONT></TD>"
			If bActivate And Not bReadOnly Then
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""BirthPlace"" ID=""BirthPlaceTxt"" VALUE=""" & aEmployeeComponent(S_EMPLOYEE_BIRTHPLACE) & """ SIZE=""30"" MAXLENGTH=""100""" & sReadOnly & "CLASS=""TextFields"" /></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_EMPLOYEE_BIRTHPLACE) & "</FONT></TD>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BirthPlace"" ID=""BirthPlaceHdn"" VALUE=""" & aEmployeeComponent(S_EMPLOYEE_BIRTHPLACE) & """ />"
			End If
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Género:&nbsp;</FONT></TD>"
				If bActivate And Not bReadOnly Then
					Response.Write "<TD><SELECT NAME=""GenderID"" ID=""GenderIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Genders", "GenderID", "GenderID As RecordID, GenderName", "", "GenderName", aEmployeeComponent(N_GENDER_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Else
					Call GetNameFromTable(oADODBConnection, "Genders", aEmployeeComponent(N_GENDER_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""GenderID"" ID=""GenderIDHdn"" VALUE=""" & aEmployeeComponent(N_GENDER_ID_EMPLOYEE) & """ />"
					Response.Write "<TD>"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT>"
					Response.Write "</TD>"
				End If
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Estado civil:&nbsp;</FONT></TD>"
				If bActivate And Not bReadOnly Then
					Response.Write "<TD><SELECT NAME=""MaritalStatusID"" ID=""MaritalStatusIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "MaritalStatus", "MaritalStatusID", "MaritalStatusID As RecordID, MaritalStatusName", "", "MaritalStatusName", aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Else
					Call GetNameFromTable(oADODBConnection, "MaritalStatus", aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""MaritalStatusID"" ID=""MaritalStatusIDHdn"" VALUE=""" & aEmployeeComponent(N_GENDER_ID_EMPLOYEE) & """ />"
					Response.Write "<TD>"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT>"
					Response.Write "</TD>"
				End If
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Grupo Sanguíneo:&nbsp;</FONT></TD>"
			If bActivate And Not bReadOnly Then
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""BloodType"" ID=""BloodTypeTxt"" VALUE=""" & aEmployeeComponent(S_EMPLOYEE_BLOODTYPE) & """ SIZE=""6"" MAXLENGTH=""10""" & sReadOnly & "CLASS=""TextFields"" /></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_EMPLOYEE_BLOODTYPE) & "</FONT></TD>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BloodType"" ID=""BloodTypeHdn"" VALUE=""" & aEmployeeComponent(S_EMPLOYEE_BLOODTYPE) & """ />"
			End If
		Response.Write "</TR>"

		lErrorNumber = GetEmployeeStartDate(oADODBConnection, aEmployeeComponent, sErrorDescription)
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de ingreso al Instituto:&nbsp;</FONT></TD>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aEmployeeComponent(N_START_DATE_EMPLOYEE), -1, -1, -1) & "</FONT></TD>"
		Response.Write "</TR>"
	Response.Write "</TABLE>"
	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		If bActivate And Not bReadOnly Then
			Response.Write "<TR><TD COLSPAN=""2"">"
			Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Domicilio:<BR /></FONT>"
			Response.Write "<TEXTAREA NAME=""EmployeeAddress"" ID=""EmployeeAddressTxtArea"" ROWS=""5"" COLS=""50"" MAXLENGTH=""255"" CLASS=""TextFields"">" & aEmployeeComponent(S_ADDRESS_EMPLOYEE) & "</TEXTAREA>"
			Response.Write "</TD></TR>"
		Else
			Response.Write "<TR><TD>"
			Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Domicilio:<BR /></FONT>"
			Response.Write "<TD>"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_ADDRESS_EMPLOYEE) & "</FONT>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeAddress"" ID=""EmployeeAddressHdn"" VALUE=""" & aEmployeeComponent(S_ADDRESS_EMPLOYEE) & """ />"
			Response.Write "</TD></TR>"
		End If
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Delegación o Municipio:&nbsp;</FONT></TD>"
			If bActivate And Not bReadOnly Then
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeCity"" ID=""EmployeeCityTxt"" SIZE=""40"" MAXLENGTH=""255"" VALUE=""" & aEmployeeComponent(S_CITY_EMPLOYEE) & """ CLASS=""TextFields"" /></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_CITY_EMPLOYEE) & "</FONT></TD>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeCity"" ID=""EmployeeCityHdn"" VALUE=""" & aEmployeeComponent(S_CITY_EMPLOYEE) & """ />"
			End If
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Código Postal:&nbsp;</FONT></TD>"
			If bActivate And Not bReadOnly Then
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeZipCode"" ID=""EmployeeZipCodeTxt"" SIZE=""5"" MAXLENGTH=""5"" VALUE=""" & aEmployeeComponent(S_ZIP_CODE_EMPLOYEE) & """ CLASS=""TextFields"" /></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_ZIP_CODE_EMPLOYEE) & "</FONT></TD>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeZipCode"" ID=""EmployeeZipCodeHdn"" VALUE=""" & aEmployeeComponent(S_ZIP_CODE_EMPLOYEE) & """ />"
			End If
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Estado:&nbsp;</FONT></TD>"
				If bActivate And Not bReadOnly Then
					Response.Write "<TD><SELECT NAME=""StateID"" ID=""StateIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "States", "StateID", "StateCode, StateName", "", "StateName", aEmployeeComponent(N_ADDRESS_STATE_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Else
					Call GetNameFromTable(oADODBConnection, "States", aEmployeeComponent(N_ADDRESS_STATE_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StateID"" ID=""StateIDHdn"" VALUE=""" & aEmployeeComponent(N_ADDRESS_STATE_ID_EMPLOYEE) & """ />"
					Response.Write "<TD>"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT>"
					Response.Write "</TD>"
				End If
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Correo electrónico:&nbsp;</FONT></TD>"
			If bActivate And Not bReadOnly Then
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeEmail"" ID=""EmployeeEmailTxt"" VALUE=""" & aEmployeeComponent(S_EMAIL_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""120"" CLASS=""TextFields"" /></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_EMAIL_EMPLOYEE) & "</FONT></TD>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeEmail"" ID=""EmployeeEmailHdn"" VALUE=""" & aEmployeeComponent(S_EMAIL_EMPLOYEE) & """ />"
			End If
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Teléfono casa:&nbsp;</FONT></TD>"
			If bActivate And Not bReadOnly Then
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><INPUT TYPE=""TEXT"" NAME=""EmployeePhone"" ID=""EmployeePhoneTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & aEmployeeComponent(S_EMPLOYEE_PHONE_EMPLOYEE) & """ CLASS=""TextFields"" /></FONT></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_EMPLOYEE_PHONE_EMPLOYEE) & "</FONT></TD>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeePhone"" ID=""EmployeePhoneHdn"" VALUE=""" & aEmployeeComponent(S_EMPLOYEE_PHONE_EMPLOYEE) & """ />"
			End If
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Teléfono celular:&nbsp;</FONT></TD>"
			If bActivate And Not bReadOnly Then
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><INPUT TYPE=""TEXT"" NAME=""CellPhone"" ID=""CellPhoneTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & aEmployeeComponent(S_EMPLOYEE_CELLPHONE) & """ CLASS=""TextFields"" /></FONT></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_EMPLOYEE_CELLPHONE) & "</FONT></TD>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CellPhone"" ID=""CellPhoneHdn"" VALUE=""" & aEmployeeComponent(S_EMPLOYEE_CELLPHONE) & """ />"
			End If
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Teléfono oficina:&nbsp;</FONT></TD>"
			If bActivate And Not bReadOnly Then
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""OfficePhone"" ID=""OfficePhoneTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & aEmployeeComponent(S_OFFICE_PHONE_EMPLOYEE) & """ CLASS=""TextFields"" /></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_OFFICE_PHONE_EMPLOYEE) & "</FONT></TD>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OfficePhone"" ID=""OfficePhoneHdn"" VALUE=""" & aEmployeeComponent(S_OFFICE_PHONE_EMPLOYEE) & """ />"
			End If
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Ext. oficina:&nbsp;</FONT></TD>"
			If bActivate And Not bReadOnly Then
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""OfficeExt"" ID=""OfficeExtTxt"" SIZE=""5"" MAXLENGTH=""5"" VALUE=""" & aEmployeeComponent(S_EXT_OFFICE_EMPLOYEE) & """ CLASS=""TextFields"" /></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_EXT_OFFICE_EMPLOYEE) & "</FONT></TD>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OfficeExt"" ID=""OfficeExtHdn"" VALUE=""" & aEmployeeComponent(S_EXT_OFFICE_EMPLOYEE) & """ />"
			End If
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Clave de elector:&nbsp;</FONT></TD>"
			If bActivate And Not bReadOnly Then
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><INPUT TYPE=""TEXT"" NAME=""DocumentNumber1"" ID=""DocumentNumber1Txt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & aEmployeeComponent(S_DOCUMENT_NUMBER_1_EMPLOYEE) & """ CLASS=""TextFields"" /></FONT></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_DOCUMENT_NUMBER_1_EMPLOYEE) & "</FONT></TD>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DocumentNumber1"" ID=""DocumentNumber1Hdn"" VALUE=""" & aEmployeeComponent(S_DOCUMENT_NUMBER_1_EMPLOYEE) & """ />"
			End If
		Response.Write "</TR>"
		If bActivate And Not bReadOnly Then
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No.&nbsp;cartilla&nbsp;militar:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><INPUT TYPE=""TEXT"" NAME=""DocumentNumber3"" ID=""DocumentNumber3Txt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & aEmployeeComponent(S_DOCUMENT_NUMBER_3_EMPLOYEE) & """ CLASS=""TextFields"" /></FONT></TD>"
			Response.Write "</TR>"
		Else
			If aEmployeeComponent(N_GENDER_ID_EMPLOYEE) = 1 Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No.&nbsp;cartilla&nbsp;militar:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_DOCUMENT_NUMBER_3_EMPLOYEE) & "</FONT></TD>"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DocumentNumber3"" ID=""DocumentNumber3Hdn"" VALUE=""" & aEmployeeComponent(S_DOCUMENT_NUMBER_3_EMPLOYEE) & """ />"
				Response.Write "</TR>"
			End If
		End If
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Actividades:&nbsp;</FONT></TD>"
				If bActivate And Not bReadOnly Then
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><SELECT NAME=""EmployeeActivityID"" ID=""EmployeeActivityIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EmployeeActivities", "EmployeeActivityID", "EmployeeActivityName", "", "EmployeeActivityName", aEmployeeComponent(N_ACTIVITY_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></FONT></TD>"
				Else
					Call GetNameFromTable(oADODBConnection, "EmployeeActivities", aEmployeeComponent(N_ACTIVITY_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeActivityID"" ID=""EmployeeActivityIDHdn"" VALUE=""" & aEmployeeComponent(N_ACTIVITY_ID_EMPLOYEE) & """ />"
					Response.Write "<TD>"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT>"
					Response.Write "</TD>"
				End If
		Response.Write "</TR>"
	Response.Write "</TABLE>"
	Response.Write "<BR /><BR />"
	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Escolaridad</B></FONT>"
	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nivel máximo de estudios concluidos:&nbsp;</FONT></TD>"
				If bActivate And Not bReadOnly Then
					Response.Write "<TD><SELECT NAME=""SchoolarShipID"" ID=""SchoolarShipIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "SchoolarShips", "SchoolarShipID", "SchoolarShipName", "", "SchoolarShipID", aEmployeeComponent(N_EMPLOYEE_SCHOOLARSHIP_ID), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Else
					Call GetNameFromTable(oADODBConnection, "Schoolarships", aEmployeeComponent(N_EMPLOYEE_SCHOOLARSHIP_ID), "", "", sNames, sErrorDescription)
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SchoolarShipID"" ID=""SchoolarShipIDHdn"" VALUE=""" & aEmployeeComponent(N_EMPLOYEE_SCHOOLARSHIP_ID) & """ />"
					Response.Write "<TD>"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT>"
					Response.Write "</TD>"
				End If
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre de la Escuela:&nbsp;</FONT></TD>"
			If bActivate And Not bReadOnly Then
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""SchoolName"" ID=""SchoolNameTxt"" VALUE=""" & aEmployeeComponent(S_EMPLOYEE_SCHOOLNAME) & """ SIZE=""30"" MAXLENGTH=""120"" CLASS=""TextFields"" /></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_EMPLOYEE_SCHOOLNAME) & "</FONT></TD>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeEmail"" ID=""EmployeeEmailHdn"" VALUE=""" & aEmployeeComponent(S_EMPLOYEE_SCHOOLNAME) & """ />"
			End If
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;</FONT></TD>"
			If bActivate And Not bReadOnly Then
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_EMPLOYEE_SCHOOLARSHIP_DATE), "SchoolarShip", N_FORM_START_YEAR, Year(Date()), True, True) & "</FONT></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aEmployeeComponent(N_EMPLOYEE_SCHOOLARSHIP_DATE), -1, -1, -1) & "</FONT></TD>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SchoolarShipDate"" ID=""SchoolarShipDateHdn"" VALUE=""" & aEmployeeComponent(N_EMPLOYEE_SCHOOLARSHIP_DATE) & """ />"
			End If
		Response.Write "</TR>"
		Response.Write "<TR>"
		Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de fin:&nbsp;</FONT></TD>"
			If bActivate And Not bReadOnly Then
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_EMPLOYEE_SCHOOLARSHIP_DATE_END), "SchoolarShipEnd", N_FORM_START_YEAR, Year(Date()), True, True) & "</FONT></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aEmployeeComponent(N_EMPLOYEE_SCHOOLARSHIP_DATE_END), -1, -1, -1) & "</FONT></TD>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SchoolarShipDateEnd"" ID=""SchoolarShipDateEndHdn"" VALUE=""" & aEmployeeComponent(N_EMPLOYEE_SCHOOLARSHIP_DATE_END) & """ />"
			End If
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Especialidad:&nbsp;</FONT></TD>"
			If bActivate And Not bReadOnly Then
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""Specialism"" ID=""SpecialismTxt"" VALUE=""" & aEmployeeComponent(S_EMPLOYEE_SPECIALISM) & """ SIZE=""30"" MAXLENGTH=""120"" CLASS=""TextFields"" /></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_EMPLOYEE_SPECIALISM) & "</FONT></TD>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Specialism"" ID=""SpecialismHdn"" VALUE=""" & aEmployeeComponent(S_EMPLOYEE_SPECIALISM) & """ />"
			End If
		Response.Write "</TR>"
        Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cédula profesional:&nbsp;</FONT></TD>"
			If bActivate And Not bReadOnly Then
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><INPUT TYPE=""TEXT"" NAME=""DocumentNumber2"" ID=""DocumentNumber2Txt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & aEmployeeComponent(S_DOCUMENT_NUMBER_2_EMPLOYEE) & """ CLASS=""TextFields"" /></FONT></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_DOCUMENT_NUMBER_2_EMPLOYEE) & "</FONT></TD>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DocumentNumber2"" ID=""DocumentNumber2Hdn"" VALUE=""" & aEmployeeComponent(S_DOCUMENT_NUMBER_2_EMPLOYEE) & """ />"
			End If
		Response.Write "</TR>"
		Response.Write "<TR>"
		Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Idiomas:&nbsp;</FONT></TD>"
			If bActivate And Not bReadOnly Then
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""Languages"" ID=""LanguagesTxt"" VALUE=""" & aEmployeeComponent(S_EMPLOYEE_LANGUAGES) & """ SIZE=""10"" MAXLENGTH=""50"" CLASS=""TextFields"" /></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_EMPLOYEE_LANGUAGES) & "</FONT></TD>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Languages"" ID=""LanguagesHdn"" VALUE=""" & aEmployeeComponent(S_EMPLOYEE_LANGUAGES) & """ />"
			End If
		Response.Write "</TR>"
	Response.Write "</TABLE>"

	Response.Write "<BR /><BR />"
	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Beneficiarios de pago de defunción</B></FONT>"
	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		Response.Write "<TR>"
		Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre del Beneficiario 1:&nbsp;</FONT></TD>"
			If bActivate And Not bReadOnly Then
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""DeathBeneficiary"" ID=""DeathBeneficiaryTxt"" VALUE=""" & aEmployeeComponent(S_EMPLOYEE_DEATH_BENEFICIARY) & """ SIZE=""35"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_EMPLOYEE_DEATH_BENEFICIARY) & "</FONT></TD>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DeathBeneficiary"" ID=""DeathBeneficiaryHdn"" VALUE=""" & aEmployeeComponent(S_EMPLOYEE_DEATH_BENEFICIARY) & """ />"
			End If
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre del Beneficiario 2:&nbsp;</FONT></TD>"
			If bActivate And Not bReadOnly Then
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""DeathBeneficiary2"" ID=""DeathBeneficiary2Txt"" VALUE=""" & aEmployeeComponent(S_EMPLOYEE_DEATH_BENEFICIARY2) & """ SIZE=""35"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_EMPLOYEE_DEATH_BENEFICIARY2) & "</FONT></TD>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DeathBeneficiary2"" ID=""DeathBeneficiary2Hdn"" VALUE=""" & aEmployeeComponent(S_EMPLOYEE_DEATH_BENEFICIARY2) & """ />"
			End If
		Response.Write "</TR>"
	Response.Write "</TABLE>"
	Response.Write "<BR /><BR />"
End Function

Function DisplayEmployeeFormSection105(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 105
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection105"

	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Estatus actual del trabajador:</B></FONT>"
	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Estatus:&nbsp;</FONT></TD>"
				Call GetNameFromTable(oADODBConnection, "StatusEmployees", aEmployeeComponent(N_STATUS_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
			Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">El empleado está activo:&nbsp;</FONT></TD>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
				Response.Write "<INPUT TYPE=""RADIO"" NAME=""Active"" DISABLED ID=""ActiveRd"" VALUE=""1"" "
					If aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 1 Then Response.Write " CHECKED=""1"""
				Response.Write " />Sí"
				Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""1"" />"
				Response.Write "<INPUT TYPE=""RADIO"" NAME=""Active"" DISABLED ID=""ActiveRd"" VALUE=""0"""
					If aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0 Then Response.Write " CHECKED=""1"""
				Response.Write " />No"
			Response.Write "</FONT></TD>"
		Response.Write "</TR>"
	Response.Write "</TABLE>"
	Response.Write "<BR /><BR />"
End Function

Function DisplayEmployeeFormSection106(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 106
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection106"
	Dim iIndex
	Dim oRecordset

	'Condición retirada temporalmente
	'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Shifts.ShiftID, JourneyTypeID From Shifts, (SELECT ShiftID FROM EMPLOYEES WHERE EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") as empleado Where Shifts.ShiftID = empleado.ShiftID", "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Shifts.ShiftID, JourneyTypeID From Shifts", "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Horario del empleado:</B></FONT>"
	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		If B_ISSSTE Then
			If bActivate Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Horarios:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""ShiftID"" ID=""ShiftIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						If aJobComponent(N_ID_JOB) > 0 Then
							'Cambio temporal
							'Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Shifts", "ShiftID", "ShiftShortName, ShiftName", "JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB), "ShiftShortName", aEmployeeComponent(N_SHIFT_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Shifts", "ShiftID", "ShiftShortName, ShiftName", "", "ShiftShortName", aEmployeeComponent(N_SHIFT_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
						Else
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Shifts", "ShiftID", "ShiftShortName, ShiftName", "", "ShiftShortName", aEmployeeComponent(N_SHIFT_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
						End If
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
			Else
				Call GetNameFromTable(oADODBConnection, "Shifts",aEmployeeComponent(N_SHIFT_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ShiftID"" ID=""ShiftIDHdn"" VALUE=""" & aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) & """ />"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Horario:&nbsp;</FONT></TD>"
					Response.Write "<TD>"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT>"
					Response.Write "</TD>"
					Response.Write "<TD>"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(N_JOURNEY_TYPE_ID) & "</FONT>"
					Response.Write "</TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de Jornada:&nbsp;</FONT></TD>"
					Response.Write "<TD>"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & oRecordset.Fields("JourneyTypeID").Value & "</FONT>"
					Response.Write "</TD>"
				Response.Write "</TR>"
			End If
		Else
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Horario:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""> de </FONT><SELECT NAME=""StartHour1"" ID=""StartHour1Cmb"" SIZE=""1"" CLASS=""Lists"">"
					For iIndex = 0 To 2400 Step 30
						sNames = Right(("0000" & iIndex), 4)
						Response.Write "<OPTION VALUE=""" & iIndex & """"
						If iIndex = aEmployeeComponent(N_START_HOUR_1_EMPLOYEE) Then Response.Write " SELECTED=""1"""
						Response.Write ">" & Left(sNames, 2) & ":" & Right(sNames, 2) & "</OPTION>"
						If CInt(Right(("0000" & iIndex), 2)) = 30 Then iIndex = iIndex + 40
					Next
					Response.Write "</SELECT><FONT FACE=""Arial"" SIZE=""2""> a </FONT><SELECT NAME=""EndHour1"" ID=""EndHour1Cmb"" SIZE=""1"" CLASS=""Lists"">"
					For iIndex = 0 To 2400 Step 30
						sNames = Right(("0000" & iIndex), 4)
						Response.Write "<OPTION VALUE=""" & iIndex & """"
						If iIndex = aEmployeeComponent(N_END_HOUR_1_EMPLOYEE) Then Response.Write " SELECTED=""1"""
						Response.Write ">" & Left(sNames, 2) & ":" & Right(sNames, 2) & "</OPTION>"
						If CInt(Right(("0000" & iIndex), 2)) = 30 Then iIndex = iIndex + 40
					Next
					Response.Write "</SELECT><FONT FACE=""Arial"" SIZE=""2""> y de </FONT><SELECT NAME=""StartHour2"" ID=""StartHour2Cmb"" SIZE=""1"" CLASS=""Lists"">"
					For iIndex = 0 To 2400 Step 30
						sNames = Right(("0000" & iIndex), 4)
						Response.Write "<OPTION VALUE=""" & iIndex & """"
						If iIndex = aEmployeeComponent(N_START_HOUR_2_EMPLOYEE) Then Response.Write " SELECTED=""1"""
						Response.Write ">" & Left(sNames, 2) & ":" & Right(sNames, 2) & "</OPTION>"
						If CInt(Right(("0000" & iIndex), 2)) = 30 Then iIndex = iIndex + 40
					Next
					Response.Write "</SELECT><FONT FACE=""Arial"" SIZE=""2""> a </FONT><SELECT NAME=""EndHour2"" ID=""EndHour2Cmb"" SIZE=""1"" CLASS=""Lists"">"
					For iIndex = 0 To 2400 Step 30
						sNames = Right(("0000" & iIndex), 4)
						Response.Write "<OPTION VALUE=""" & iIndex & """"
						If iIndex = aEmployeeComponent(N_END_HOUR_2_EMPLOYEE) Then Response.Write " SELECTED=""1"""
						Response.Write ">" & Left(sNames, 2) & ":" & Right(sNames, 2) & "</OPTION>"
						If CInt(Right(("0000" & iIndex), 2)) = 30 Then iIndex = iIndex + 40
					Next
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
		End If
	Response.Write "</TABLE>"
	Response.Write "<BR /><BR />"
End Function

Function DisplayEmployeeFormSection107(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 107
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection107"
	Dim iIndex

	If lReasonID = 12 Then
		Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Horario del turno opcional / percepción adicional:</B></FONT>"
		Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Horario:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""> de </FONT><SELECT NAME=""StartHour3"" ID=""StartHour3Cmb"" SIZE=""1"" CLASS=""Lists"">"
					For iIndex = 0 To 2400 Step 30
						sNames = Right(("0000" & iIndex), 4)
						Response.Write "<OPTION VALUE=""" & iIndex & """"
						If iIndex = aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) Then Response.Write " SELECTED=""1"""
						Response.Write ">" & Left(sNames, 2) & ":" & Right(sNames, 2) & "</OPTION>"
						If CInt(Right(("0000" & iIndex), 2)) = 30 Then iIndex = iIndex + 40
					Next
					Response.Write "</SELECT><FONT FACE=""Arial"" SIZE=""2""> a </FONT><SELECT NAME=""EndHour3"" ID=""EndHour3Cmb"" SIZE=""1"" CLASS=""Lists"">"
					For iIndex = 0 To 2400 Step 30
						sNames = Right(("0000" & iIndex), 4)
						Response.Write "<OPTION VALUE=""" & iIndex & """"
						If iIndex = aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) Then Response.Write " SELECTED=""1"""
						Response.Write ">" & Left(sNames, 2) & ":" & Right(sNames, 2) & "</OPTION>"
						If CInt(Right(("0000" & iIndex), 2)) = 30 Then iIndex = iIndex + 40
					Next
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	Else
		If aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) = 1 Then
			Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Horario del turno opcional (base):</B></FONT>"
			Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				If (InStr(1,",1,2,3,4,5,6,8,10,37,38,39,40,41,43,44,45,46,47,48,62,63,66,", "," & lReasonID & ",", vbBinaryCompare) = 0) Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Horario:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""> de </FONT><SELECT NAME=""StartHour3"" ID=""StartHour3Cmb"" SIZE=""1"" CLASS=""Lists"">"
							For iIndex = 0 To 2400 Step 30
								sNames = Right(("0000" & iIndex), 4)
								Response.Write "<OPTION VALUE=""" & iIndex & """"
								If iIndex = aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) Then Response.Write " SELECTED=""1"""
								Response.Write ">" & Left(sNames, 2) & ":" & Right(sNames, 2) & "</OPTION>"
								If CInt(Right(("0000" & iIndex), 2)) = 30 Then iIndex = iIndex + 40
							Next
							Response.Write "</SELECT><FONT FACE=""Arial"" SIZE=""2""> a </FONT><SELECT NAME=""EndHour3"" ID=""EndHour3Cmb"" SIZE=""1"" CLASS=""Lists"">"
							For iIndex = 0 To 2400 Step 30
								sNames = Right(("0000" & iIndex), 4)
								Response.Write "<OPTION VALUE=""" & iIndex & """"
								If iIndex = aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) Then Response.Write " SELECTED=""1"""
								Response.Write ">" & Left(sNames, 2) & ":" & Right(sNames, 2) & "</OPTION>"
								If CInt(Right(("0000" & iIndex), 2)) = 30 Then iIndex = iIndex + 40
							Next
					Response.Write "</TR>"
				Response.Write "</TABLE>"
				Else
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartHour3"" ID=""StartHour3Hdn"" VALUE=""" & aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EndHour3"" ID=""EndHour3Hdn"" VALUE=""" & aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) & """ />"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Horario de:&nbsp;</FONT>"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & Left(aEmployeeComponent(N_START_HOUR_3_EMPLOYEE), 2) & ":" & Right(aEmployeeComponent(N_START_HOUR_3_EMPLOYEE), 2) & " a " & Left(aEmployeeComponent(N_END_HOUR_3_EMPLOYEE), 2) & ":" & Right(aEmployeeComponent(N_END_HOUR_3_EMPLOYEE), 2) & "</FONT>"
					Response.Write "</TD>"
				Response.Write "</TR>"
				Response.Write "</TABLE>"
				End If
			Response.Write "<BR /><BR />"
		ElseIf (aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) = 2) And (aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 0 Or aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 2 Or aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 3 Or aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 4) Then
			Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Horario de la percepción adicional (confianza):</B></FONT>"
			Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				If (InStr(1,",1,2,3,4,5,6,8,10,37,38,39,40,41,43,44,45,46,47,48,62,63,66,", "," & lReasonID & ",", vbBinaryCompare) = 0) Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Horario:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""> de </FONT><SELECT NAME=""StartHour3"" ID=""StartHour3Cmb"" SIZE=""1"" CLASS=""Lists"">"
						For iIndex = 0 To 2400 Step 30
							sNames = Right(("0000" & iIndex), 4)
							Response.Write "<OPTION VALUE=""" & iIndex & """"
							If iIndex = aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) Then Response.Write " SELECTED=""1"""
							Response.Write ">" & Left(sNames, 2) & ":" & Right(sNames, 2) & "</OPTION>"
							If CInt(Right(("0000" & iIndex), 2)) = 30 Then iIndex = iIndex + 40
						Next
						Response.Write "</SELECT><FONT FACE=""Arial"" SIZE=""2""> a </FONT><SELECT NAME=""EndHour3"" ID=""EndHour3Cmb"" SIZE=""1"" CLASS=""Lists"">"
						For iIndex = 0 To 2400 Step 30
							sNames = Right(("0000" & iIndex), 4)
							Response.Write "<OPTION VALUE=""" & iIndex & """"
							If iIndex = aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) Then Response.Write " SELECTED=""1"""
							Response.Write ">" & Left(sNames, 2) & ":" & Right(sNames, 2) & "</OPTION>"
							If CInt(Right(("0000" & iIndex), 2)) = 30 Then iIndex = iIndex + 40
						Next
					Response.Write "</TR>"
				Response.Write "</TABLE>"
			Else
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartHour3"" ID=""StartHour3Hdn"" VALUE=""" & aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EndHour3"" ID=""EndHour3Hdn"" VALUE=""" & aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) & """ />"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Horario de:&nbsp;</FONT>"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & Left(aEmployeeComponent(N_START_HOUR_3_EMPLOYEE), 2) & ":" & Right(aEmployeeComponent(N_START_HOUR_3_EMPLOYEE), 2) & " a " & Left(aEmployeeComponent(N_END_HOUR_3_EMPLOYEE), 2) & ":" & Right(aEmployeeComponent(N_END_HOUR_3_EMPLOYEE), 2) & "</FONT>"
					Response.Write "</TD>"
				Response.Write "</TR>"
				Response.Write "</TABLE>"
			End If
			Response.Write "<BR /><BR />"
		End If
	End If
End Function

Function DisplayEmployeeFormSection108(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 108
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection108"
	Dim oRecordset

	If aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) <> 7 and aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) <> 12 and aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) <> 13 Then
		Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Riesgos profesionales:</B></FONT>"
		Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptAmount, StartDate, EndDate From  EmployeesConceptsLKP Where (StartDate <= " & CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) & ") And (EndDate >= " & CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) & ") And (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID = 4) And (EndDate <> 0)", "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = oRecordset.Fields("conceptAmount").Value
			If bActivate Then
				If (InStr(1,",1,2,3,4,5,6,8,10,37,38,39,40,41,43,44,45,46,47,48,62,63,66,", "," & lReasonID & ",", vbBinaryCompare) = 0) Then
					Response.Write "<TR>"
						Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Nivel de riesgo:&nbsp;</FONT></TD>"
						Response.Write "<TD VALIGN=""TOP""><SELECT NAME=""RiskLevel"" ID=""RiskLevelCmb"" onChange=""SearchRecord(this.value + '&EmployeeID=' + EmployeeID.value, 'RiskLevel', 'RiskLevelIFrame', 'EmployeeFrm.RiskLevel');"" SIZE=""1"" CLASS=""Lists"">"
							If aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) >= 1 Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select RiskLevelID, RiskLevelName From RiskLevels where RiskLevelID > 0 Order By RiskLevelName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
							Else
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select RiskLevelID, RiskLevelName From RiskLevels Order By RiskLevelName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
							End If
							Do While Not oRecordset.EOF
							'Response.Write GenerateListOptionsFromQuery(oADODBConnection, "RiskLevels", "RiskLevelID", "RiskLevelName", "", "RiskLevelName", aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
								Response.Write "<OPTION VALUE=""" & CStr(oRecordset.Fields("RisklevelID").Value) & """"
								If aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = CInt(oRecordset.Fields("RisklevelID").Value) * 10 Then Response.Write " SELECTED"
								Response.Write ">" & CStr(oRecordset.Fields("RiskLevelName").Value) & "</OPTION>"
								oRecordset.MoveNext
							Loop
						Response.Write "</SELECT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD VALIGN=""TOP"" COLSPAN=""2""><IFRAME SRC=""SearchRecord.asp"" NAME=""RiskLevelIFrame"" FRAMEBORDER=""0"" WIDTH=""300"" HEIGHT=""100""></IFRAME></TD>"
					Response.Write "</TR>"
			Else
				Call GetNameFromTable(oADODBConnection, "RiskLevels", aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE), "", "", sNames, sErrorDescription)
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RiskLevel"" ID=""RiskLevelHdn"" VALUE=""" & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & """ />"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nivel de riesgo:&nbsp;</FONT></TD>"
					Response.Write "<TD>"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT>"
					Response.Write "</TD>"
				Response.Write "</TR>"
			End If
		Else
			Call GetNameFromTable(oADODBConnection, "RiskLevels", aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE), "", "", sNames, sErrorDescription)
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RiskLevel"" ID=""RiskLevelHdn"" VALUE=""" & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & """ />"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nivel de riesgo:&nbsp;</FONT></TD>"
				Response.Write "<TD>"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT>"
				Response.Write "</TD>"
			Response.Write "</TR>"
		End If
	Response.Write "</TABLE>"
	Response.Write "<BR /><BR />"
	End If
End Function

Function DisplayEmployeeFormSection109(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 109
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection109"

	Select Case aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE)
		Case 1
			aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 120
			lErrorNumber = GetEmployeeSpecificConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Prestaciones:</B></FONT>"
			Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Seguro de separación individualizado:&nbsp;</B></FONT></TD>"
					Response.Write "<TD>"
					Response.Write "</TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE>"
			Response.Write "<BR />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				If aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) > 0 Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;</FONT></TD>"
						Response.Write "<TD>"
							Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE)) & "</FONT>"
						Response.Write "</TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de fin:&nbsp;</FONT></TD>"
						Response.Write "<TD>"
							If aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 30000000 Then
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Indefinida</FONT>"
							Else
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE)) & "</FONT>"
							End If
						Response.Write "</TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Monto:&nbsp;</FONT></TD>"
						Response.Write "<TD>"
							Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE), 0, True, False, True) & "%</FONT>"
						Response.Write "</TD>"
					Response.Write "</TR>"
				Else
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No tiene registrada esta prestación.&nbsp;</FONT></TD>"
						Response.Write "<TD>"
						Response.Write "</TD>"
					Response.Write "</TR>"
				End If
			Response.Write "</TABLE>"
			Response.Write "<BR /><BR />"
			aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 87
			lErrorNumber = GetEmployeeSpecificConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><NOBR><B>Seguro adicional de separación individualizado</B>:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD>"
					Response.Write "</TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE>"
			Response.Write "<BR />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				If aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) > 0 Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;</FONT></TD>"
						Response.Write "<TD>"
							Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE)) & "</FONT>"
						Response.Write "</TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de fin:&nbsp;</FONT></TD>"
						Response.Write "<TD>"
							If aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 30000000 Then
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Indefinida</FONT>"
							Else
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE)) & "</FONT>"
							End If
						Response.Write "</TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Monto:&nbsp;</FONT></TD>"
						Response.Write "<TD>"
							Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE), 0, True, False, True) & "</FONT>"
							Call GetNameFromTable(oADODBConnection, "QttyNames", aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
							Response.Write CleanStringForHTML(sNames)
							'Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "(QttID=" & aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) & ")", "QttyID", aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</TD>"
					Response.Write "</TR>"
				Else
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No tiene registrada esta prestación.&nbsp;</FONT></TD>"
						Response.Write "<TD>"
						Response.Write "</TD>"
					Response.Write "</TR>"
				End If
			Response.Write "</TABLE>"
			Response.Write "<BR /><BR />"
		Case 0,2,3,4,5,6
			aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 4
			lErrorNumber = GetEmployeeSpecificConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Riesgos profesionales:</B></FONT>"
			Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				If aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) > 0 Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;</FONT></TD>"
						Response.Write "<TD>"
							Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE)) & "</FONT>"
						Response.Write "</TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de fin:&nbsp;</FONT></TD>"
						Response.Write "<TD>"
							If aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 30000000 Then
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Indefinida</FONT>"
							Else
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE)) & "</FONT>"
							End If
						Response.Write "</TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Monto:&nbsp;</FONT></TD>"
						Response.Write "<TD>"
							Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE), 0, True, False, True) & "%</FONT>"
						Response.Write "</TD>"
					Response.Write "</TR>"
				Else
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No tiene registrados riesgos profesionales:&nbsp;</FONT></TD>"
						Response.Write "<TD>"
						Response.Write "</TD>"
					Response.Write "</TR>"
				End If
			Response.Write "</TABLE>"
			Response.Write "<BR /><BR />"
	End Select
	If (aJobComponent(N_POSITION_TYPE_ID_JOB) = 1 Or aJobComponent(N_POSITION_TYPE_ID_JOB) = 2) And aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) <> 1 Then
		If aJobComponent(N_POSITION_TYPE_ID_JOB) = 1 Then
			aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 7
		Else
			aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 8
		End If
			lErrorNumber = GetEmployeeSpecificConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			If aJobComponent(N_POSITION_TYPE_ID_JOB) = 1 Then
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Turno opcional (base):&nbsp;</B></FONT>"
			Else
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Percepción adicional (confianza):&nbsp;</B></FONT>"
			End If
			Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				If aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) > 0 Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;</FONT></TD>"
						Response.Write "<TD>"
							Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE)) & "</FONT>"
						Response.Write "</TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de fin:&nbsp;</FONT></TD>"
						Response.Write "<TD>"
							If aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 30000000 Then
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Indefinida</FONT>"
							Else
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & DisplayNumericDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE)) & "</FONT>"
							End If
						Response.Write "</TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Monto:&nbsp;</FONT></TD>"
						Response.Write "<TD>"
							Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE), 2, True, False, True) & "%</FONT>"
						Response.Write "</TD>"
					Response.Write "</TR>"
				Else
					Response.Write "<TR>"
						If aJobComponent(N_POSITION_TYPE_ID_JOB) = 1 Then
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No tiene registrado el turno opcional (base):&nbsp;</FONT></TD>"
						Else
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No tiene registrada la percepción adicional (confianza):&nbsp;</FONT></TD>"
						End If
						Response.Write "<TD>"
						Response.Write "</TD>"
					Response.Write "</TR>"
				End If
			Response.Write "</TABLE>"
			Response.Write "<BR /><BR />"
	End If
End Function

Function DisplayEmployeeFormSection111(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 111
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection111"
	Dim sDropCause
	Dim sComments

	'If (aEmployeeComponent(N_STATUS_REASON_ID_EMPLOYEE) = 1) Or (aEmployeeComponent(N_STATUS_REASON_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = 155) Then
	'	If (lReasonID > 0) And (lReasonID <> 57) Then
	'		Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Observaciones del movimiento</B></FONT>"
	'			Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
	'			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
	'				Response.Write "<TR>"
	'				    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><TEXTAREA NAME=""EmployeeComments"" ID=""EmployeeCommentsTxtArea"" ROWS=""5"" COLS=""60"" MAXLENGTH=""500"" CLASS=""TextFields"">" & aEmployeeComponent(S_COMMENTS_EMPLOYEE) & "</TEXTAREA></FONT></TD>"
	'				    Response.Write "<TD></FONT></TD>"
	'				Response.Write "</TR>"
	'			Response.Write "</TABLE>"
	'			Response.Write "<BR /><BR />"
	'	End If
	'Else
	If lReasonID = 66 Then
		sDropCause = Mid(aEmployeeComponent(S_COMMENTS_EMPLOYEE),1,InStr(1,aEmployeeComponent(S_COMMENTS_EMPLOYEE),"Þ",vbBinaryCompare)-1)
		sComments = Mid(aEmployeeComponent(S_COMMENTS_EMPLOYEE),InStr(1,aEmployeeComponent(S_COMMENTS_EMPLOYEE),"Þ",vbBinaryCompare)+1)
		Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Causa de la baja:</B></FONT>"
		Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Motivo de la baja:&nbsp;</FONT></TD>"
			Response.Write "<TD><SELECT NAME=""DropReasonName"" ID=""DropReasonNameCmb"" SIZE=""1"" CLASS=""Lists"">"
				Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Reasons", "ReasonName", "ReasonName", "ReasonID in (1,2,4,63)", "ReasonID", sDropCause, "Ninguno;;;-1", sErrorDescription)
			Response.Write "</SELECT></TD>"
		Response.Write "</TR>"	
		Response.Write "<TR>"
			Response.Write "<TD ColSpan = 2><FONT FACE=""Arial"" SIZE=""2""><TEXTAREA NAME=""EmployeeComments"" ID=""EmployeeCommentsTxtArea"" ROWS=""5"" COLS=""60"" MAXLENGTH=""500"" CLASS=""TextFields"">" & sComments & "</TEXTAREA></FONT></TD>"
			Response.Write "<TD></FONT></TD>"
		Response.Write "</TR>"
		Response.Write "</TABLE>"
		Response.Write "<BR /><BR />"
	Else
		Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Observaciones del movimiento:</B></FONT>"
		Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				If lReasonID = 53 Then
					If Len(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE)) > 0 Then
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) & "</FONT></TD>"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeComments"" ID=""EmployeeCommentsHdn"" VALUE=""" & aEmployeeComponent(S_COMMENTS_EMPLOYEE) & """ />"
					Else
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No existen observaciones</FONT></TD>"
					End If
				Else
					If Len(aEmployeeComponent(S_COMMENTS_EMPLOYEE)) > 0 Then
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_COMMENTS_EMPLOYEE) & "</FONT></TD>"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeComments"" ID=""EmployeeCommentsHdn"" VALUE=""" & aEmployeeComponent(S_COMMENTS_EMPLOYEE) & """ />"
					Else
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No existen observaciones</FONT></TD>"
					End If
				End If
			    Response.Write "<TD></FONT></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
		Response.Write "<BR /><BR />"
	End If
End Function

Function DisplayEmployeeFormSection112(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 112
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection112"

	If aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) <> 30000000 And (lReasonID >= 36 And lReasonID <= 41) And (aEmployeeComponent(N_STATUS_REASON_ID_EMPLOYEE) < 1) And  (InStr(1, ",74,82,102,110,118,136,", "," & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ",", vbBinaryCompare) = 0) Then
		aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = AddDaysToSerialDate(aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE), 1)
		aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = AddDaysToSerialDate(aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE), 2)
	End If
	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Vigencia del movimiento</B></FONT>"
	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio de la vigencia:&nbsp;</FONT></TD>"
			'Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE), "Employee", Year(Date())-2, Year(Date())+7, True, True) & "</FONT></TD>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE), "Employee", Year(Date()) -2, Year(Date())+2, True, False) & "</FONT></TD>"
		Response.Write "</TR>"
		If lReasonID <> 57 Then
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de fin de la vigencia:&nbsp;</FONT></TD>"
				'Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE), "EmployeeEnd", Year(Date())-2, Year(Date())+50, True, True) & "</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE), "EmployeeEnd", Year(Date()) - 2, Year(Date())+2, True, True) & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR NAME=""PayrollDateDiv"" ID=""PayrollDateDiv"">"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Quincena de aplicación:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""EmployeePayrollDate"" ID=""EmployeePayrollDateCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""CheckStatusPayrolls()"">"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And (IsActive_1=1) And (PayrollTypeID IN (1,4))", "PayrollID Desc", aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
		End If
	Response.Write "</TABLE>"
	Response.Write "<BR /><BR />"
End Function

Function DisplayEmployeeFormSection113(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 113
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection113"

	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Vigencia del movimiento</B></FONT>"
	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
	If aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) <> 30000000 Then
		aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = AddDaysToSerialDate(aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE), 1)
	End If
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de reanudación&nbsp;</FONT></TD>"
			'Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE), "Employee", Year(Date())-3, Year(Date())+1, True, True) & "</FONT></TD>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE), "Employee", Year(Date()) - 2, Year(Date())+1, True, False) & "</FONT></TD>"
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""></FONT></TD>"
			Response.Write "<TD></TD>"
		Response.Write "</TR>"
		Response.Write "<TR NAME=""PayrollDateDiv"" ID=""PayrollDateDiv"">"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Quincena de aplicación:&nbsp;</FONT></TD>"
			Response.Write "<TD><SELECT NAME=""EmployeePayrollDate"" ID=""EmployeePayrollDateCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""CheckStatusPayrolls()"">"
				Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And (IsActive_1=1) And (PayrollTypeID=1)", "PayrollID Desc", aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
			Response.Write "</SELECT></TD>"
		Response.Write "</TR>"
	Response.Write "</TABLE>"
	Response.Write "<BR /><BR />"
End Function

Function DisplayEmployeeFormSection115a(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 115a
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection115a"
	Dim oRecordset

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeRequirementID From EmployeesRequirements Where ReasonID=" & lReasonID, "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Do While Not oRecordset.EOF
				asRequirements(CInt(oRecordset.Fields("EmployeeRequirementID").Value)) = 0
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesRequirementsFM1LKP Where EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & " And ReasonID=" & lReasonID, "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			sRequirements = ","
			If Not oRecordset.EOF Then
				Do While Not oRecordset.EOF
					sRequirements = sRequirements & CStr(oRecordset.Fields("EmployeeRequirementID").Value) & ","
					asRequirements(CInt(oRecordset.Fields("EmployeeRequirementID").Value)) = 1
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
			End If
		End If
	End If
End Function

Function DisplayEmployeeFormSection015b(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 015b
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection015b"
	Dim oRecordset

	Response.Write "<BR /><FONT FACE=""Arial"" SIZE=""2""><B>Documentos requeridos para la aplicación del movimiento</B></FONT>"
	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesRequirements Where ReasonID=" & lReasonID, "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Este movimiento no tiene requisitos</B></FONT>"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
					Response.Write "<TR>"
						Response.Write "<TD>"
						Response.Write "</TD>"
						Response.Write "<TD></TD>"
					Response.Write "</TR>"
			Response.Write "</TABLE>"
			Response.Write "<BR />"
		Else
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Do While Not oRecordset.EOF
					Response.Write "<TR>"
						Response.Write "<TD>"
						Response.Write "<INPUT TYPE=""CHECKBOX"" NAME=""" & CStr(oRecordset.Fields("EmployeeRequirementID").Value) & """ ID=""" & CStr(oRecordset.Fields("EmployeeRequirementID").Value) & "Chk"" Value=""" & asRequirements(CInt(oRecordset.Fields("EmployeeRequirementID").Value)) & """"
							If (InStr(1, sRequirements, "," & oRecordset.Fields("EmployeeRequirementID").Value & ",", vbBinaryCompare) > 0) Then Response.Write " CHECKED=""1"""
							Response.Write " />"
						Response.Write "</TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
							Response.Write CStr(oRecordset.Fields("EmployeeRequirementName").Value)
						Response.Write "</FONT></TD>"
					Response.Write "</TR>"
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
			Response.Write "</TABLE>"
			Response.Write "<BR /><BR />"
		End If
	End If
	Select Case lReasonID
		Case EMPLOYEES_SAFE_SEPARATION, EMPLOYEES_ADD_SAFE_SEPARATION, CANCEL_EMPLOYEES_SSI
			Response.Write "<IFRAME SRC=""BrowserFile.asp?EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & """ NAME=""EmployeeFilesIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""348""></IFRAME>"
		Case Else
	End Select
End Function

Function DisplayEmployeeFormSection116a()
'************************************************************
'Purpose: To add JavaScript to validate amounts taken
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection116a"

	Select Case lReasonID
		Case 14
			Response.Write "if (parseInt(oForm.CompanyID.value) == -1) {" & vbNewLine
				Response.Write "alert('Favor de indicar la empresa de la plaza.');" & vbNewLine
				Response.Write "oForm.CompanyID.focus();" & vbNewLine
				Response.Write "return false;" & vbNewLine
			Response.Write "}" & vbNewLine
			Response.Write "if ((parseInt(oForm.AreaID.value) == -1) || (parseInt(oForm.AreaID.value) == 0)) {" & vbNewLine
				Response.Write "alert('Favor de indicar el centro de trabajo de la plaza.');" & vbNewLine
				Response.Write "oForm.AreaID.focus();" & vbNewLine
				Response.Write "return false;" & vbNewLine
			Response.Write "}" & vbNewLine
			Response.Write "if (parseInt(oForm.PaymentCenterID.value) == -1) {" & vbNewLine
				Response.Write "alert('Favor de indicar el centro de pago de la plaza.');" & vbNewLine
				Response.Write "oForm.PaymentCenterID.focus();" & vbNewLine
				Response.Write "return false;" & vbNewLine
			Response.Write "}" & vbNewLine
			Response.Write "if (parseInt(oForm.ServiceID.value) == -1) {" & vbNewLine
				Response.Write "if (parseInt(oForm.EmployeeTypeID.value) != 7) {" & vbNewLine
					Response.Write "alert('Favor de indicar el servicio de la plaza.');" & vbNewLine
					Response.Write "oForm.ServiceID.focus();" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "}" & vbNewLine
			Response.Write "if (parseInt(oForm.JourneyID.value) == -1) {" & vbNewLine
				Response.Write "alert('Favor de indicar el turno de la plaza.');" & vbNewLine
				Response.Write "oForm.JourneyID.focus();" & vbNewLine
				Response.Write "return false;" & vbNewLine
			Response.Write "}" & vbNewLine
			Response.Write "if (parseInt(oForm.ShiftID.value) == -1) {" & vbNewLine
				Response.Write "alert('Favor de indicar el horario de la plaza.');" & vbNewLine
				Response.Write "oForm.ShiftID.focus();" & vbNewLine
				Response.Write "return false;" & vbNewLine
			Response.Write "}" & vbNewLine
	End Select
End Function

Function DisplayEmployeeFormSection116b(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 116b
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection116b"
	Dim sStatusJobIDs
	Dim oRecordset

	Select Case lReasonID
		Case 14
			Response.Write "<BR /><FONT FACE=""Arial"" SIZE=""2""><B>Complete los datos de la plaza que se creará</B></FONT>"
			Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD>" & "<IMG SRC=""Images/Transparent.gif"" WIDTH=""120"" HEIGHT=""1"" />" & "</TD>"
					Response.Write "<TD></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número de plaza:</FONT></TD>"
					Response.Write "<TD ALIGN=""LEFT""><FONT FACE=""Arial"" SIZE=""2"">" & RIGHT("000000" & aEmployeeComponent(N_ID_EMPLOYEE),6) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Puesto genérico:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						Call GetNameFromTable(oADODBConnection, "Positions", L_HONORARY_POSITION_ID, "", "", sNames, sErrorDescription)
						Response.Write sNames
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR NAME=""CompanyDiv"" ID=""CompanyDiv"">"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Empresa:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""CompanyID"" ID=""CompanyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Companies", "CompanyID", "CompanyShortName, CompanyName", "(ParentID>-1) And (Active=1)", "CompanyShortName", aJobComponent(N_COMPANY_ID_JOB), "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Centro de trabajo:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						Response.Write "<SELECT NAME=""AreaID"" ID=""AreaIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) = 0 Then
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(AreaID>-1) And (Active=1)", "AreaCode", aJobComponent(N_AREA_ID_JOB), "Ninguna;;;-1", sErrorDescription)
							Else
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(AreaID>-1) And (Active=1) And (Areas.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))", "AreaCode", aJobComponent(N_AREA_ID_JOB), "Ninguna;;;-1", sErrorDescription)
							End If
						Response.Write "</SELECT>"
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD>"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Centro de pago:&nbsp;</FONT>"
					Response.Write "</TD>"
					Response.Write "<TD>"
						Response.Write "<SELECT NAME=""PaymentCenterID"" ID=""PaymentCenterIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) = 0 Then
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(Active=1)", "AreaCode", aJobComponent(N_PAYMENT_CENTER_ID_JOB), "Ninguno;;;-1", sErrorDescription)
							Else
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(Active=1) And (Areas.AreaID In (" &aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))", "AreaCode", aJobComponent(N_PAYMENT_CENTER_ID_JOB), "Ninguno;;;-1", sErrorDescription)
							End If
						Response.Write "</SELECT><BR />"
					Response.Write "</TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Servicio:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""ServiceID"" ID=""ServiceIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Services", "ServiceID", "ServiceShortName, ServiceName", "", "ServiceShortName", aJobComponent(N_SERVICE_ID_JOB), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Turno:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""JourneyID"" ID=""JourneyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Journeys", "JourneyID", "JourneyShortName, JourneyName", "", "JourneyShortName", aJobComponent(N_JOURNEY_ID_JOB), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Horarios:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""ShiftID"" ID=""ShiftIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Shifts", "ShiftID", "ShiftShortName, ShiftName", "", "ShiftShortName", aJobComponent(N_SHIFT_ID_JOB), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Monto mensual:&nbsp;</FONT></TD>"
					Response.Write "<TD>"
						If (aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 7 or aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 12 or aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 13) Then
							Response.Write "<INPUT TYPE=""TEXT"" NAME=""ConceptAmount"" ID=""ConceptAmountTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & FormatNumber((aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) * 2), 2, True, False, True) & """ CLASS=""TextFields"" />&nbsp;"
						Else
							Response.Write "<INPUT TYPE=""TEXT"" NAME=""ConceptAmount"" ID=""ConceptAmountTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & FormatNumber(aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE), 2, True, False, True) & """ CLASS=""TextFields"" />&nbsp;"
						End If
							Response.Write "<SELECT NAME=""ConceptQttyID"" ID=""ConceptQttyIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowAmountFields(this.value, 'Concept');"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "QttyID=1", "QttyID", aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT>&nbsp;"
							Response.Write "<SPAN ID=""ConceptCurrencySpn""><SELECT NAME=""CurrencyID"" ID=""CurrencyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Currencies", "CurrencyID", "CurrencyName", "(CurrencyID=0) And (Active=1)", "CurrencyName", aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE), "Ninguna;;;-1", sErrorDescription)
							Response.Write "</SELECT></SPAN>"
					Response.Write "</TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE>"
			Response.Write "<BR /><BR />"
		Case 21
			Response.Write "<BR /><FONT FACE=""Arial"" SIZE=""2""><B>Nueva plaza a asignar misma adscripción:</B></FONT>"
			Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusJob2 From Reasons Where ReasonID=" & lReasonID, "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							sStatusJobIDs = CStr(oRecordset.Fields("StatusJob2").Value)
						End If
					End If
						Response.Write "<TD WIDTH=""17%""><FONT FACE=""Arial"" SIZE=""2"">Número de plaza:&nbsp;</FONT><INPUT TYPE=""TEXT"" NAME=""JobNumber"" ID=""JobNumberTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE="""" CLASS=""TextFields"" onChange=""document.EmployeeFrm.JobID.value='';"" />"
						Response.Write "<A HREF=""javascript: document.EmployeeFrm.JobID.value=''; SearchRecord(document.EmployeeFrm.JobNumber.value, 'JobNumber', 'SearchEmployeeNumberIFrame', 'EmployeeFrm.JobID')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Buscar el número de plaza"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><FONT FACE=""Arial"" SIZE=""2"">Seleccione la plaza:&nbsp;</FONT>&nbsp;</TD>"
						Response.Write "<TD>"
							Response.Write "<SELECT NAME=""JobID"" ID=""JobIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""SearchRecord(this.value, 'JobNumber', 'SearchEmployeeNumberIFrame', 'EmployeeFrm.JobNumber');"">"
								Response.Write "<OPTION VALUE=""""></OPTION>"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Jobs,(Select areaid, parentid from Areas) as areas,(Select zoneid from zones) as zones,(Select areaid from Jobs Where JobID = " & aJobComponent(N_ID_JOB) & ") as Plaza", "JobID", "JobNumber", "(Jobs.AreaID = Plaza.AreaID) And (Jobs.AreaID = areas.AreaID) And (Jobs.ZoneID = zones.ZoneID) And (Jobs.PositionID <> -1) And (Jobs.StatusID = 2)", "JobNumber", aEmployeeComponent(N_JOB_ID_HISTORY_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT>"
						Response.Write "</TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD VALIGN=""TOP""><IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""400""></IFRAME></TD>"
					Response.Write "</TR>"
				Response.Write "</TABLE>"
			Response.Write "<BR /><BR />"
		Case 13, 17
			Response.Write "<BR /><FONT FACE=""Arial"" SIZE=""2""><B>Nueva plaza a asignar:</B></FONT>"
			Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusJob2 From Reasons Where ReasonID=" & lReasonID, "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							sStatusJobIDs = CStr(oRecordset.Fields("StatusJob2").Value)
						End If
					End If
						Response.Write "<TD WIDTH=""17%""><FONT FACE=""Arial"" SIZE=""2"">Número de plaza:&nbsp;</FONT><INPUT TYPE=""TEXT"" NAME=""JobNumber"" ID=""JobNumberTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE="""" CLASS=""TextFields"" onChange=""document.EmployeeFrm.JobID.value='';"" />"
						If lReasonID = 13 Then
							Response.Write "<A HREF=""javascript: document.EmployeeFrm.JobID.value=''; SearchRecord(document.EmployeeFrm.JobNumber.value, 'JobNumber&StartYear='+document.EmployeeFrm.EmployeeYear.value+'&StartMonth='+document.EmployeeFrm.EmployeeMonth.value+'&StartDay='+document.EmployeeFrm.EmployeeDay.value+'&EndYear='+document.EmployeeFrm.EmployeeEndYear.value+'&EndMonth='+document.EmployeeFrm.EmployeeEndMonth.value+'&EndDay='+document.EmployeeFrm.EmployeeEndDay.value+'&lReasonID=13', 'SearchEmployeeNumberIFrame', 'EmployeeFrm.JobID')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Buscar el número de plaza"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><FONT FACE=""Arial"" SIZE=""2"">Seleccione la plaza:&nbsp;</FONT>&nbsp;</TD>"
						Else
							Response.Write "<A HREF=""javascript: document.EmployeeFrm.JobID.value=''; SearchRecord(document.EmployeeFrm.JobNumber.value, 'JobNumber', 'SearchEmployeeNumberIFrame', 'EmployeeFrm.JobID')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Buscar el número de plaza"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><FONT FACE=""Arial"" SIZE=""2"">Seleccione la plaza:&nbsp;</FONT>&nbsp;</TD>"
						End If
						Response.Write "<TD>"
						If lReasonID = 13 Then
							Response.Write "<SELECT NAME=""JobID"" ID=""JobIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""SearchRecord(this.value, 'JobNumber&IsCombo=1&StartYear='+document.EmployeeFrm.EmployeeYear.value+'&StartMonth='+document.EmployeeFrm.EmployeeMonth.value+'&StartDay='+document.EmployeeFrm.EmployeeDay.value+'&EndYear='+document.EmployeeFrm.EmployeeEndYear.value+'&EndMonth='+document.EmployeeFrm.EmployeeEndMonth.value+'&EndDay='+document.EmployeeFrm.EmployeeEndDay.value+'&lReasonID=13', 'SearchEmployeeNumberIFrame', 'EmployeeFrm.JobNumber');"">"
						Else
							Response.Write "<SELECT NAME=""JobID"" ID=""JobIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""SearchRecord(this.value, 'JobNumber', 'SearchEmployeeNumberIFrame', 'EmployeeFrm.JobNumber');"">"
						End If
								Response.Write "<OPTION VALUE=""""></OPTION>"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Jobs", "JobID", "JobNumber", "StatusID In (" & sStatusJobIDs & ") And (Jobs.PositionID <> -1)", "JobNumber", aEmployeeComponent(N_JOB_ID_HISTORY_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT>"
						Response.Write "</TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD VALIGN=""TOP""><IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""400""></IFRAME></TD>"
					Response.Write "</TR>"
				Response.Write "</TABLE>"
			Response.Write "<BR /><BR />"
		Case 50
			Response.Write "<BR /><FONT FACE=""Arial"" SIZE=""2""><B>Nueva plaza a asignar:</B></FONT>"
			Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusJob2 From Reasons Where ReasonID=" & lReasonID, "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							sStatusJobIDs = CStr(oRecordset.Fields("StatusJob2").Value)
						End If
					End If
						Response.Write "<TD WIDTH=""17%""><FONT FACE=""Arial"" SIZE=""2"">Número de plaza:&nbsp;</FONT><INPUT TYPE=""TEXT"" NAME=""JobNumber"" ID=""JobNumberTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE="""" CLASS=""TextFields"" onChange=""document.EmployeeFrm.JobID.value='';"" />"
						Response.Write "<A HREF=""javascript: document.EmployeeFrm.JobID.value=''; SearchRecord(document.EmployeeFrm.JobNumber.value, 'JobNumber', 'SearchEmployeeNumberIFrame', 'EmployeeFrm.JobID')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Buscar el número de plaza"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><FONT FACE=""Arial"" SIZE=""2"">Seleccione la plaza:&nbsp;</FONT>&nbsp;</TD>"
						Response.Write "<TD>"
							Response.Write "<SELECT NAME=""JobID"" ID=""JobIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""SearchRecord(this.value, 'JobNumber', 'SearchEmployeeNumberIFrame', 'EmployeeFrm.JobNumber');"">"
								Response.Write "<OPTION VALUE=""""></OPTION>"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Jobs, PositionTypes, Positions", "JobID", "JobNumber", "Jobs.StatusID In (" & sStatusJobIDs & ") And (Jobs.PositionID <> -1) And (Jobs.PositionID = Positions.PositionID) And (Positions.PositionTypeID = PositionTypes.PositionTypeID) And (PositionTypes.PositionTypeID = 1)", "JobNumber", aEmployeeComponent(N_JOB_ID_HISTORY_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT>"
						Response.Write "</TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD VALIGN=""TOP""><IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""400""></IFRAME></TD>"
					Response.Write "</TR>"
				Response.Write "</TABLE>"
			Response.Write "<BR /><BR />"
		Case 68
			Response.Write "<BR /><FONT FACE=""Arial"" SIZE=""2""><B>Nueva plaza a asignar:</B></FONT>"
			Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusJob2 From Reasons Where ReasonID=" & lReasonID, "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							sStatusJobIDs = CStr(oRecordset.Fields("StatusJob2").Value)
						End If
					End If
						Response.Write "<TD WIDTH=""17%""><FONT FACE=""Arial"" SIZE=""2"">Número de plaza:&nbsp;</FONT><INPUT TYPE=""TEXT"" NAME=""JobNumber"" ID=""JobNumberTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE="""" CLASS=""TextFields"" onChange=""document.EmployeeFrm.JobID.value='';"" />"
						Response.Write "<A HREF=""javascript: document.EmployeeFrm.JobID.value=''; SearchRecord(document.EmployeeFrm.JobNumber.value, 'JobNumber', 'SearchEmployeeNumberIFrame', 'EmployeeFrm.JobID')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Buscar el número de plaza"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><FONT FACE=""Arial"" SIZE=""2"">Seleccione la plaza:&nbsp;</FONT>&nbsp;</TD>"
						Response.Write "<TD>"
							Response.Write "<SELECT NAME=""JobID"" ID=""JobIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""SearchRecord(this.value, 'JobNumber', 'SearchEmployeeNumberIFrame', 'EmployeeFrm.JobNumber');"">"
								Response.Write "<OPTION VALUE=""""></OPTION>"
									If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) = 0 Then
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Jobs, Positions", "JobID", "JobNumber", "Jobs.StatusID In (" & sStatusJobIDs & ") And (Jobs.PositionID=Positions.PositionID) And (Positions.PositionTypeID=2) And (Jobs.PositionID <> -1)", "JobNumber", aEmployeeComponent(N_JOB_ID_HISTORY_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
									Else
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Jobs, Positions", "JobID", "JobNumber", "Jobs.StatusID In (" & sStatusJobIDs & ") And (Jobs.PositionID=Positions.PositionID) And (Positions.PositionTypeID=2) And (Jobs.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) And (Jobs.PositionID <> -1)", "JobNumber", aEmployeeComponent(N_JOB_ID_HISTORY_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
									End If
							Response.Write "</SELECT>"
						Response.Write "</TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD VALIGN=""TOP""><IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""400""></IFRAME></TD>"
					Response.Write "</TR>"
				Response.Write "</TABLE>"
			Response.Write "<BR /><BR />"
		Case Else
			Response.Write "<BR /><FONT FACE=""Arial"" SIZE=""2""><B>Nueva plaza a asignar:</B></FONT>"
				Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
				Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusJob2 From Reasons Where ReasonID=" & lReasonID, "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							sStatusJobIDs = CStr(oRecordset.Fields("StatusJob2").Value)
						End If
					End If
						Response.Write "<TD WIDTH=""17%""><FONT FACE=""Arial"" SIZE=""2"">Número de plaza:&nbsp;</FONT><INPUT TYPE=""TEXT"" NAME=""JobNumber"" ID=""JobNumberTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE="""" CLASS=""TextFields"" onChange=""document.EmployeeFrm.JobID.value='';"" />"
						Response.Write "<A HREF=""javascript: document.EmployeeFrm.JobID.value=''; SearchRecord(document.EmployeeFrm.JobNumber.value, 'JobNumber', 'SearchEmployeeNumberIFrame', 'EmployeeFrm.JobID')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Buscar el número de plaza"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><FONT FACE=""Arial"" SIZE=""2"">Seleccione la plaza:&nbsp;</FONT>&nbsp;</TD>"
						Response.Write "<TD>"
							Response.Write "<SELECT NAME=""JobID"" ID=""JobIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""SearchRecord(this.value, 'JobNumber', 'SearchEmployeeNumberIFrame', 'EmployeeFrm.JobNumber');"">"
								Response.Write "<OPTION VALUE=""""></OPTION>"
								If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) = 0 Then
									If (aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) <> -1) Then
										If ((CLng(aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE)) >= 0 And CLng(aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE)) <= 4) Or (CLng(aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE)) = 10)) Then
											Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Jobs, Positions", "JobID", "JobNumber", "Jobs.StatusID In (" & sStatusJobIDs & ") And (Jobs.PositionID=Positions.PositionID) And (Positions.EmployeeTypeID In (0,1,2,3,4,10) And (Jobs.PositionID <> -1))", "JobNumber", aEmployeeComponent(N_JOB_ID_HISTORY_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
										Else
											Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Jobs, Positions", "JobID", "JobNumber", "Jobs.StatusID In (" & sStatusJobIDs & ") And (Jobs.PositionID=Positions.PositionID) And (Positions.EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ") And (Jobs.PositionID <> -1)", "JobNumber", aEmployeeComponent(N_JOB_ID_HISTORY_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
										End If
									Else
										If ((CLng(aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE)) >= 0 And CLng(aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE)) <= 4) Or (CLng(aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE)) = 10)) Then
											Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Jobs, Positions", "JobID", "JobNumber", "Jobs.StatusID In (" & sStatusJobIDs & ") And (Jobs.PositionID=Positions.PositionID) And (Positions.EmployeeTypeID In (0,1,2,3,4,10) And (Jobs.PositionID <> -1))", "JobNumber", aEmployeeComponent(N_JOB_ID_HISTORY_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
										Else
											Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Jobs, Positions", "JobID", "JobNumber", "Jobs.StatusID In (" & sStatusJobIDs & ") And (Jobs.PositionID=Positions.PositionID) And (Jobs.PositionID <> -1)", "JobNumber", aEmployeeComponent(N_JOB_ID_HISTORY_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
										End If
									End If
								Else
									If (aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) <> -1) Then
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Jobs, Positions", "JobID", "JobNumber", "Jobs.StatusID In (" & sStatusJobIDs & ") And (Jobs.PositionID=Positions.PositionID) And (Positions.EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ") And (Jobs.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) And (Jobs.PositionID <> -1)", "JobNumber", aEmployeeComponent(N_JOB_ID_HISTORY_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
									Else
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Jobs, Positions", "JobID", "JobNumber", "Jobs.StatusID In (" & sStatusJobIDs & ") And (Jobs.PositionID=Positions.PositionID) And (Jobs.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) And (Jobs.PositionID <> -1)", "JobNumber", aEmployeeComponent(N_JOB_ID_HISTORY_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
									End If
								End If
							Response.Write "</SELECT>"
						Response.Write "</TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD VALIGN=""TOP""><IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""400""></IFRAME></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE>"
			Response.Write "<BR /><BR />"
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "if (document.EmployeeFrm.JobID.value !='') {" & vbNewLine
				Response.Write "SearchRecord(EmployeeFrm.JobID.value, 'JobNumber', 'SearchEmployeeNumberIFrame', 'EmployeeFrm.JobNumber');" & vbNewLine
			Response.Write "}" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
	End Select
End Function

Function DisplayEmployeeFormSection117(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 117
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection117"

	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Vigencia del movimiento</B></FONT>"
	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de baja:&nbsp;</FONT></TD>"
			If lreasonID = 62 Or lReasonID = 63 Then
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE), "Employee", Year(Date()) - 40, Year(Date())+1, True, True) & "</FONT></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE), "Employee", Year(Date()) - 20, Year(Date())+1, True, True) & "</FONT></TD>"
			End If
		Response.Write "</TR>"
		Response.Write "<TR NAME=""PayrollDateDiv"" ID=""PayrollDateDiv"">"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Quincena de aplicación:&nbsp;</FONT></TD>"
			Response.Write "<TD><SELECT NAME=""EmployeePayrollDate"" ID=""EmployeePayrollDateCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""CheckStatusPayrolls()"">"
				Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And (IsActive_1=1) And (PayrollTypeID=1)", "PayrollID Desc", aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
			Response.Write "</SELECT></TD>"
		Response.Write "</TR>"
	Response.Write "</TABLE>"
	Response.Write "<BR /><BR />"
End Function

Function DisplayEmployeeFormSection118(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 118
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection118"
	Dim oRecordset

	If lReasonID = 51 Then
		sErrorDescription = "No se pudo obtener la información de la plaza."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesHistoryList Where (bProcessed=2) And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID=" & lReasonID & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aJobComponent(N_AREA_ID_JOB) = CLng(oRecordset.Fields("AreaID").Value)
				aJobComponent(N_PAYMENT_CENTER_ID_JOB) = CLng(oRecordset.Fields("PaymentCenterID").Value)
			End If
		End If
	End If
	Response.Write "<BR /><FONT FACE=""Arial"" SIZE=""2""><B>Indique la nueva adscripción de la plaza:</B></FONT>"
		Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Centro de trabajo:&nbsp;</FONT></TD>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
				Response.Write "<SELECT NAME=""AreaID"" ID=""AreaIDCmb"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(AreaID>-1) And (Active=1)", "AreaCode", aJobComponent(N_AREA_ID_JOB), "Ninguna;;;-1", sErrorDescription)
				Response.Write "</SELECT>"
			Response.Write "</FONT></TD>"
		Response.Write "</TR>"
		Response.Write "<TR>"
				Response.Write "<TD>"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Centro de pago:&nbsp;</FONT>"
				Response.Write "</TD>"
				Response.Write "<TD>"
					Response.Write "<SELECT NAME=""PaymentCenterID"" ID=""PaymentCenterIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(AreaID>-1) And (Active=1)", "AreaCode", aJobComponent(N_PAYMENT_CENTER_ID_JOB), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT><BR />"
				Response.Write "</TD>"
			Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Servicio:&nbsp;</FONT></TD>"
			Response.Write "<TD><SELECT NAME=""ServiceID"" ID=""ServiceIDCmb"" SIZE=""1"" CLASS=""Lists"">"
				Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Services", "ServiceID", "ServiceShortName, ServiceName", "", "ServiceShortName", aEmployeeComponent(N_SERVICE_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
			Response.Write "</SELECT></TD>"
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Turno:&nbsp;</FONT></TD>"
			Response.Write "<TD><SELECT NAME=""JourneyID"" ID=""JourneyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
				Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Journeys", "JourneyID", "JourneyShortName, JourneyName", "", "JourneyShortName", aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
			Response.Write "</SELECT></TD>"
		Response.Write "</TR>"
'						Response.Write "<TR>"
'							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Horario:&nbsp;</FONT></TD>"
'							Response.Write "<TD><SELECT NAME=""newShiftID"" ID=""newShiftIDCmb"" SIZE=""1"" CLASS=""Lists"">"
'								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Shifts", "ShiftID", "ShiftShortName, ShiftName", "", "ShiftShortName", aEmployeeComponent(N_SHIFT_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
'							Response.Write "</SELECT></TD>"
'						Response.Write "</TR>"
	Response.Write "</TABLE>"
	Response.Write "<BR /><BR />"
End Function

Function DisplayEmployeeFormSection120a(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 120a
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection120a"

	Dim sAltDescription
	Dim sDescription

	Select Case sURL
		Case "EmployeesMovements"
			Select Case lReasonID
				Case EMPLOYEES_BANK_ACCOUNTS
					sAltDescription = "Cuenta bancaria"
					sDescription = "Registre una cuenta bancaria a un empleado diferente."
				Case EMPLOYEES_ADD_BENEFICIARIES
					sAltDescription = "Beneficiario de pensión alimenticia"
					sDescription = "Registre un(a) beneficiario(a) de pensión alimenticia a un empleado diferente."
				Case EMPLOYEES_CREDITORS
					sAltDescription = "Acreedores"
					sDescription = "Registre un(a) acreedor(a) de adeudos a un empleado diferente."
				Case -58
					sAltDescription = "Reclamo de pago"
					sDescription = "Registre un reclamo de pago por ajustes a un empleado diferente."
				Case CANCEL_EMPLOYEES_CONCEPTS, CANCEL_EMPLOYEES_C04
					sAltDescription = "Baja de prestaciones"
					sDescription = "De baja prestaciones de un empleado diferente."
				Case EMPLOYEES_THIRD_CONCEPT
					sAltDescription = "Registro en línea de Terceros"
					sDescription = "Registre un crédito a un empleado diferente."
				Case -96,-75,-64,1,2,3,4,5,6,8,10,12,13,14,17,18,21,26,28,29,30,31,32,33,34,37,38,39,40,41,43,44,45,46,47,48,50,51,53,57,58,62,63,66,68
					sAltDescription = "Movimientos de personal"
					sDescription = "Registre el movimiento a un empleado diferente."
				Case EMPLOYEES_GRADE
					sAltDescription = "Calificación de empleados"
					sDescription = "Registre la calificación a un empleado diferente."
				Case Else
					sAltDescription = "Prestación"
					sDescription = "Registre la prestación a un empleado diferente."
			End Select
			Call DisplayAnotherEmployeeForm(oRequest, oADODBConnection, sAction, "EmployeesMovements",400, lReasonID, sAltDescription, sDescription, sErrorDescription)
		Case "ServiceSheet"
			sAltDescription = "Hoja única de servicios"
			sDescription = "Registre la solicitud de una Hoja única de servicios a un empleado diferente."
			Call DisplayAnotherEmployeeForm(oRequest, oADODBConnection, "Main_ISSSTE.asp", "ServiceSheet", 400, EMPLOYEES_SERVICE_SHEET, sAltDescription, sDescription, sErrorDescription)
	End Select
End Function

Function DisplayEmployeeFormSection120b(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 120b
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection120b"

	If (lReasonID = -61 Or lReasonID = -62) And (aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) <> 1) Then
		Call DisplayErrorMessage("Mensaje del sistema", "El concepto sólo puede otorgarse a funcionarios")
	End If
	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Información general</B></FONT>"
	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del empleado:&nbsp;</FONT></TD>"
			If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" VALUE=""" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""SpecialTextFields"" onFocus=""document.EmployeeFrm.EmployeeName.focus()"" /></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "</FONT><INPUT TYPE=""HIDDEN"" NAME=""EmployeeNumber"" ID=""EmployeeNumberHdn"" VALUE=""" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & """ /></TD>"
			End If
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre(s):&nbsp;</FONT></TD>"
			Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeName"" ID=""EmployeeNameTxt"" VALUE=""" & aEmployeeComponent(S_NAME_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100""" & sReadOnly & "CLASS=""TextFields"" /></TD>"
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido paterno:&nbsp;</FONT></TD>"
			Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName"" ID=""EmployeeLastNameTxt"" VALUE=""" & aEmployeeComponent(S_LAST_NAME_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100""" & sReadOnly & "CLASS=""TextFields"" /></TD>"
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido materno:&nbsp;</FONT></TD>"
			Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName2"" ID=""EmployeeLastName2Txt"" VALUE=""" & aEmployeeComponent(S_LAST_NAME2_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100""" & sReadOnly & "CLASS=""TextFields"" /></TD>"
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Estatus:&nbsp;</FONT></TD>"
			Call GetNameFromTable(oADODBConnection, "StatusEmployees", aEmployeeComponent(N_STATUS_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
		Response.Write "</TR>"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartDate"" ID=""StartDateHdn"" VALUE=""" & aEmployeeComponent(N_START_DATE_EMPLOYEE) & """ />"
	Response.Write "</TABLE>"
	If (InStr(1, aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE), ",135,", vbBinaryCompare) > 0) Then
		Dim sAccountNumber
		Response.Write "<B>Tipo de pago: "
		Call GetNameFromTable(oADODBConnection, "EmployeeAccount", aEmployeeComponent(N_ID_EMPLOYEE), "", "", sAccountNumber, sErrorDescription)
		If StrComp(sAccountNumber, ".", vbBinaryCompare) = 0 Then
			Response.Write "Cheque"
		Else
			If Len(sAccountNumber) = 0 Then
				Response.Write "Ninguna"
			Else
				Response.Write "Depósito"
			End If
		End If
		Response.Write "</B><BR />"
	End If
	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de tabulador:&nbsp;</FONT></TD>"
			Response.Write "<TD><SELECT NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDCmb"" SIZE=""1"" CLASS=""Lists"""
				If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then 
					Response.Write " onChange=""GetEmployeeNumber(this.value)"""
					Response.Write ">"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EmployeeTypes", "EmployeeTypeID", "EmployeeTypeID As RecordID, EmployeeTypeName", "(Active=1)", "EmployeeTypeName", aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
				Else
					Response.Write ">"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EmployeeTypes", "EmployeeTypeID", "EmployeeTypeID As RecordID, EmployeeTypeName", "(Active=1) And (EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ")", "EmployeeTypeName", aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
				End If
			Response.Write "</SELECT></TD>"
		Response.Write "</TR>"
		If (CInt(Request.Cookies("SIAP_SectionID")) <> 2) And (lReasonID <> EMPLOYEES_BANK_ACCOUNTS) Then
			Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Antigüedad ISSSTE:&nbsp;</FONT></TD>"
			If (CInt(aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE)) <> 6) And (CInt(aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE)) <> 7 And CInt(aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE)) <> 12 And CInt(aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE)) <> 13) Then
				lErrorNumber = CalculateEmployeeAntiquity(oADODBConnection, aEmployeeComponent, 0, sEmployeeDisplayFormAntiquity, lDisplayFormAntiquityYears, lDisplayFormAntiquityMonths, lDisplayFormAntiquityDays, sErrorDescription)
			Else
				lDisplayFormAntiquityYears = 0
				lDisplayFormAntiquityMonths = 0
				lDisplayFormAntiquityDays = 0
				sEmployeeDisplayFormAntiquity = "0 Años 0 Meses 0 Días"
			End If
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & sEmployeeDisplayFormAntiquity & "</TD>"
			Response.Write "<TD><INPUT TYPE=""HIDDEN"" NAME=""AntiquityYears"" ID=""AntiquityYearsTxt"" VALUE=""" & lDisplayFormAntiquityYears & """ CLASS=""TextFields"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AntiquityMonths"" ID=""AntiquityMonthsTxt"" VALUE=""" & lDisplayFormAntiquityMonths & """ CLASS=""TextFields"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AntiquityDays"" ID=""AntiquityDaysTxt"" VALUE=""" & lDisplayFormAntiquityDays & """ CLASS=""TextFields"" />"
			Response.Write "</TD>"
			Response.Write "</TR>"
		End If
	Response.Write "</TABLE>"
	If aEmployeeComponent(N_JOB_ID_EMPLOYEE) <> -1 Then
		Response.Write "<BR />"
		Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Información de la plaza:</B></FONT>"
		Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Plaza:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & Right("000000" & aEmployeeComponent(N_JOB_ID_EMPLOYEE), Len("000000")) & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Puesto:&nbsp;</FONT></TD>"
				Call GetNameFromTable(oADODBConnection, "Positions", aJobComponent(N_POSITION_ID_JOB), "", "", sNames, sErrorDescription)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & sNames & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de puesto:&nbsp;</FONT></TD>"
				Call GetNameFromTable(oADODBConnection, "PositionTypes", aJobComponent(N_POSITION_TYPE_ID_JOB), "", "", sNames, sErrorDescription)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & sNames & "</FONT></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	Else
		Response.Write "<BR />"
		Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Información de la plaza:</B></FONT>"
		Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Plaza:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">El empleado no cuenta con una plaza asignada</FONT></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	End If
	Select Case lReasonID
		Case EMPLOYEES_SAFE_SEPARATION, EMPLOYEES_ADD_SAFE_SEPARATION
			Response.Write "<IFRAME SRC=""BrowserFile.asp?EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & """ NAME=""EmployeeFilesIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""348""></IFRAME>"
			Response.Write "<BR /><BR />"
		Case CANCEL_EMPLOYEES_CONCEPTS, CANCEL_EMPLOYEES_C04, CANCEL_EMPLOYEES_SSI
		Case Else
			Response.Write "<BR /><BR />"
	End Select
End Function

Function DisplayEmployeeFormSection121(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 121
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection121"

	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Información de los hijos:</B></FONT>"
	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			Response.Write"<TR colspan = 4>"                    
					lErrorNumber = DisplayEmployeeChildrenTable(oRequest, oADODBConnection, "EmployeesMovements", DISPLAY_NOTHING, True, False, aEmployeeComponent, sErrorDescription)	
			Response.Write "</TR>"	
	Response.Write "</TABLE>"
	Response.Write "<BR /><BR />"
End Function
Function DisplayEmployeeFormSection121a(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 121a
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection121a"

	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Información de los hijos:</B></FONT>"
	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
	
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><INPUT TYPE=""HIDDEN"" NAME=""ChildID"" ID=""ChildID"" SIZE=""5"" MAXLENGTH=""5"" VALUE="""& aEmployeeComponent(N_ID_CHILD_EMPLOYEE)  &""" CLASS=""TextFields"" /></FONT></TD>"  
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><INPUT TYPE=""TEXT"" NAME=""ChildName"" ID=""ChildNameTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE="""& aEmployeeComponent(S_NAME_CHILD_EMPLOYEE)  &""" CLASS=""TextFields"" /></FONT></TD>"  
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido Paterno:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><INPUT TYPE=""TEXT"" NAME=""ChildLastName"" ID=""ChildLastNameTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE="""& aEmployeeComponent(S_LAST_NAME_CHILD_EMPLOYEE) &""" CLASS=""TextFields"" /></FONT></TD>"
			Response.Write "<TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido materno:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><INPUT TYPE=""TEXT"" NAME=""ChildLastName2"" ID=""ChildLastName2Txt"" SIZE=""30"" MAXLENGTH=""100"" VALUE="""& aEmployeeComponent(S_LAST_NAME2_CHILD_EMPLOYEE) &""" CLASS=""TextFields"" /></FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de nacimiento:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_BIRTH_DATE_CHILD_EMPLOYEE), "ChildBirth", N_FORM_START_YEAR, Year(Date()), True, True) & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nivel máximo de estudios concluidos:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""ChildrenSchoolarShipID"" ID=""ChildrenSchoolarShipIDCmb"" SIZE=""1"" CLASS=""Lists"">"
				    Response.Write GenerateListOptionsFromQuery(oADODBConnection, "SchoolarShips", "SchoolarShipID", "SchoolarShipName", "", "SchoolarShipID", aEmployeeComponent(N_CHILD_LEVEL_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write"<TD>"
					If aEmployeeComponent(N_ID_CHILD_EMPLOYEE) = -1 Then
						If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""AddChild"" ID=""AddChildBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
					ElseIf Len(oRequest("Delete").Item) > 0 Then
						If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveConceptWngDiv']); ConceptFrm.Remove.focus()"" />"
					Else
						If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ModifyChildren"" ID=""ModifyChildrenBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
					End If
				Response.Write"</TD>"
				Response.Write "<TD>"
					Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons""/>"
					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
				Response.Write"</TD>"
			Response.Write "</TR>"
	Response.Write "</TABLE>"
	Response.Write "<BR /><BR />"
	Call DisplayWarningDiv("RemoveConceptWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
End Function

Function DisplayEmployeeFormSection125(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 125
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection125"
	Dim oRecordset

	Dim iIndex
	Dim sExtraURL
	Select Case lReasonID
		Case EMPLOYEES_GRADE
			sNames = "Calificación de empleados"
		Case EMPLOYEES_SERVICE_SHEET
			sNames = "Solicitud de hoja única de servicios"
		Case Else
			Call GetNameFromTable(oADODBConnection, "Reasons", lReasonID, "", "", sNames, "")
	End Select
	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Registro de: " & sNames & "</B></FONT>"
	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		Select Case lReasonID
			Case EMPLOYEES_SUNDAYS
		'If lReasonID = EMPLOYEES_SUNDAYS Then
				If Len(oRequest("SundayChange").Item) > 0 Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de incidencia:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), -1, -1, -1) & "</FONT></TD>"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptStartDate"" ID=""ConceptStartDateHdn"" VALUE=""" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & """ />"
					Response.Write "</TR>"
				Else
					aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = aEmployeeComponent(N_ID_EMPLOYEE)
					Response.Write "<TR>"
						Response.Write "<TD VALIGN=""TOP"">"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Seleccione solo días domingo del calendario.</B></FONT><BR /><BR />"
						Response.Write "<IFRAME SRC=""BrowserMonthForPayments.asp?FormName=EmployeeFrm&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&AbsenceDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & "&FromArrow=1&ReasonID=" & lReasonID & "&OnlySundays=1"" NAME=""BrowserMonthIFrame"" FRAMEBORDER=""0"" WIDTH=""340"" HEIGHT=""130""></IFRAME>"
						Response.Write "</TD>"
						Response.Write "<TD>&nbsp;&nbsp;&nbsp;</TD>"
						If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = 0) Then
								Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fechas seleccionadas:&nbsp;&nbsp;</FONT></TD>"
								Response.Write "<TD VALIGN=""TOP""><SELECT NAME=""OcurredDates"" ID=""OcurredDatesLst"" SIZE=""7"" MULTIPLE=""5"" CLASS=""Lists""></SELECT></TD>"
								Response.Write "<TD VALIGN=""BOTTOM""><A HREF=""javascript: RemoveSelectedItemsFromList(null, document.EmployeeFrm.OcurredDates)""><IMG SRC=""Images/BtnCrclDelete.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Eliminar"" BORDER=""0"" HSPACE=""5"" /></A></TD>"
						End If
					Response.Write "</TR>"
				End If
			Case EMPLOYEES_SERVICE_SHEET
				If Len(oRequest("ServiceSheetChange").Item) > 0 Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de solicitud:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aEmployeeComponent(N_EMPLOYEE_DOCUMENT_DATE), -1, -1, -1) & "</FONT></TD>"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DocumentDate"" ID=""DocumentDateHdn"" VALUE=""" & aEmployeeComponent(N_EMPLOYEE_DOCUMENT_DATE) & """ />"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Hora de solicitud:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayTimeFromSerialNumber(aEmployeeComponent(N_EMPLOYEE_DOCUMENT_TIME)) & "</FONT></TD>"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DocumentTime"" ID=""DocumentTimeHdn"" VALUE=""" & aEmployeeComponent(N_EMPLOYEE_DOCUMENT_TIME) & """ />"
					Response.Write "</TR>"
				Else
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de solicitud:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(Left(GetSerialNumberForDate(""), Len("00000000")), "DocumentDate", Year(Date())-1, Year(Date())+1, True, False) & "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Hora de solicitud:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayTimeCombosUsingSerial(aEmployeeComponent(N_EMPLOYEE_DOCUMENT_TIME), "Document", 0, 23, 1, False) & "</FONT></TD>"
					Response.Write "</TR>"
				End If
				If Len(oRequest("ServiceSheetChange").Item) > 0 Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de recepción:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aEmployeeComponent(N_EMPLOYEE_DOCUMENT_DATE_2), -1, -1, -1) & "</FONT></TD>"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Document2Date"" ID=""Document2DateHdn"" VALUE=""" & aEmployeeComponent(N_EMPLOYEE_DOCUMENT_DATE_2) & """ />"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Hora de recepción:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayTimeFromSerialNumber(aEmployeeComponent(N_EMPLOYEE_DOCUMENT_TIME_2)) & "</FONT></TD>"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Document2Time"" ID=""Document2TimeHdn"" VALUE=""" & aEmployeeComponent(N_EMPLOYEE_DOCUMENT_TIME_2) & """ />"
					Response.Write "</TR>"
				Else
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de recepción:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(Left(GetSerialNumberForDate(""), Len("00000000")), "Document2Date", Year(Date())-1, Year(Date())+1, True, False) & "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Hora de recepción:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayTimeCombosUsingSerial(aEmployeeComponent(N_EMPLOYEE_DOCUMENT_TIME_2), "Document2", 0, 23, 1, False) & "</FONT></TD>"
					Response.Write "</TR>"
				End If
			Case Else
				Select Case lReasonID
					Case EMPLOYEES_NIGHTSHIFTS
						If Len(oRequest("ModifyConcept").Item) > 0 Then
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de incidencia:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), -1, -1, -1) & "</FONT></TD>"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptStartDate"" ID=""ConceptStartDateHdn"" VALUE=""" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & """ />"
							Response.Write "</TR>"
						Else
							sExtraURL = "FormName=EmployeeFrm"
								Response.Write "<TR>"
									Response.Write "<TD VALIGN=""TOP"">"
										Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Seleccione solo días festivos.<BR />Estan marcados en rojo.</B></FONT><BR /><BR />"
										Response.Write "<IFRAME SRC=""BrowserMonthForPayments.asp?EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&FromArrow=1&OnlyHolidays=1&" & sExtraURL & """ NAME=""BrowserMonthIFrame"" FRAMEBORDER=""0"" WIDTH=""330"" HEIGHT=""130""></IFRAME>"
									Response.Write "</TD>"
									Response.Write "<TD>&nbsp;&nbsp;&nbsp;</TD>"
									Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fechas de las jornadas nocturnas:&nbsp;</FONT></TD>"
									Response.Write "<TD VALIGN=""TOP""><SELECT NAME=""OcurredDates"" ID=""OcurredDatesLst"" SIZE=""7"" MULTIPLE=""1""></SELECT></TD>"
									Response.Write "<TD VALIGN=""BOTTOM""><A HREF=""javascript: RemoveSelectedItemsFromList(null, document.EmployeeFrm.OcurredDates)""><IMG SRC=""Images/BtnCrclDelete.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Eliminar"" BORDER=""0"" HSPACE=""5"" /></A></TD>"
								Response.Write "</TR>"
						End If
					Case EMPLOYEES_EXTRAHOURS
						If Len(oRequest("SundayChange").Item) > 0 Then
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de incidencia:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), -1, -1, -1) & "</FONT></TD>"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptStartDate"" ID=""ConceptStartDateHdn"" VALUE=""" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & """ />"
							Response.Write "</TR>"
						Else
							sExtraURL = "FormName=EmployeeFrm"
								Response.Write "<TR>"
									Response.Write "<TD VALIGN=""TOP"">"
										Response.Write "<IFRAME SRC=""BrowserMonthForPayments.asp?FormName=EmployeeFrm&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&EmployeeDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & "&FromArrow=1&ReasonID=" & lReasonID & """ NAME=""BrowserMonthIFrame"" FRAMEBORDER=""0"" WIDTH=""330"" HEIGHT=""130""></IFRAME>"
									Response.Write "</TD>"
									Response.Write "<TD>&nbsp;&nbsp;&nbsp;</TD>"
									Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fechas de las horas extras:&nbsp;</FONT></TD>"
									Response.Write "<TD VALIGN=""TOP""><SELECT NAME=""OcurredDates"" ID=""OcurredDatesLst"" SIZE=""7"" MULTIPLE=""1""></SELECT></TD>"
									Response.Write "<TD VALIGN=""BOTTOM""><A HREF=""javascript: RemoveSelectedItemsFromList(null, document.EmployeeFrm.OcurredDates)""><IMG SRC=""Images/BtnCrclDelete.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Eliminar"" BORDER=""0"" HSPACE=""5"" /></A></TD>"
								Response.Write "</TR>"
						End If
					Case EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_GLASSES, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_MONTHAWARD, EMPLOYEES_NIGHTSHIFTS, EMPLOYEES_CONCEPT_C3, EMPLOYEES_MOTHERAWARD, EMPLOYEES_ANUAL_AWARD, EMPLOYEES_NIGHTSHIFTS, EMPLOYEES_THIRD_CONCEPT, EMPLOYEES_FONAC_CONCEPT, EMPLOYEES_FONAC_ADJUSTMENT, EMPLOYEES_ANTIQUITY_25_AND_30_YEARS
					Case EMPLOYEES_GRADE
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Año:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
							Response.Write "&nbsp;&nbsp;<TD><SELECT NAME=""YearID"" ID=""YearIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								For iIndex = (Year(Date()) - 2) To Year(Date()) + 2
									Response.Write "<OPTION VALUE=""" & iIndex & """>" & iIndex & "</OPTION>"
								Next
							Response.Write "</SELECT></TD>"
						Response.Write "</TR>"
					Case Else
						'lDisplayFormCurrentDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
						'If lReasonID = EMPLOYEES_FOR_RISK Then aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 4
						'If lReasonID = EMPLOYEES_ADDITIONALSHIFT Then aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 7
						'If lReasonID = EMPLOYEES_CONCEPT_08 Then aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 8
                        'If lReasonID = EMPLOYEES_HONORARIUM_CONCEPT Then aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 13
						'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StartDate, EndDate From EmployeesConceptsLKP Where EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & " And ConceptID = " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & " And StartDate < " & lDisplayFormCurrentDate & " And EndDate > " & lDisplayFormCurrentDate, "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						'aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = oRecordset.Fields("StartDate").Value
						'aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = oRecordset.Fields("StartDate").Value
						'aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = oRecordset.Fields("EndDate").Value
						If Len(oRequest("ModifyConcept").Item) > 0 Then
							If lReasonID <> EMPLOYEES_FOR_RISK Then
								Response.Write "<TR>"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), -1, -1, -1) & "</FONT></TD>"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptStartDate"" ID=""ConceptStartDateHdn"" VALUE=""" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & """ />"
								Response.Write "</TR>"
							Else
								Response.Write "<TR>"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE), -1, -1, -1) & "</FONT></TD>"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptStartDate"" ID=""ConceptStartDateHdn"" VALUE=""" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & """ />"
								Response.Write "</TR>"
							End If
						Else
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;</FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE), "ConceptStart", Year(Date())-1, Year(Date())+1, True, False) & "</FONT></TD>"
							Response.Write "</TR>"
						End If
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de fin de la vigencia:&nbsp;</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE), "ConceptEnd", Year(Date())-1, Year(Date())+3, True, True) & "</FONT></TD>"
						Response.Write "</TR>"
				End Select
		End Select
	Response.Write "</TABLE>"
	If lReasonID = EMPLOYEES_SUNDAYS Then Response.Write "<BR />"
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		Select Case lReasonID
			Case EMPLOYEES_THIRD_CONCEPT
				If Len(oRequest("CreditChange").Item) > 0 Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Q. Aplicación:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), -1, -1, -1) & "</FONT></TD>"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeePayrollDate"" ID=""EmployeePayrollDateHdn"" VALUE=""" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & """ />"
					Response.Write "</TR>"
				Else
					Response.Write "<TR NAME=""PayrollDateDiv"" ID=""PayrollDateDiv"">"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Quincena de aplicación:&nbsp;</FONT></TD>"
						Response.Write "<TD><SELECT NAME=""EmployeePayrollDate"" ID=""EmployeePayrollDateCmb"" SIZE=""1"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And (IsActive_7=1) And (PayrollTypeID=1)", "PayrollID Desc", aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
						Response.Write "</SELECT>&nbsp;"
						Response.Write "</TD>"
					Response.Write "</TR>"
				End If
			Case EMPLOYEES_GRADE
				Response.Write "<TR NAME=""PayrollDateDiv"" ID=""PayrollDateDiv"">"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Quincena a considerar:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""EmployeePayrollDate"" ID=""EmployeePayrollDateCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID=1) And (IsActive_7=1)", "PayrollID Desc", aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
					Response.Write "</SELECT>&nbsp;"
					Response.Write "</TD>"
				Response.Write "</TR>"
			Case EMPLOYEES_SERVICE_SHEET
			Case Else
				Response.Write "<TR NAME=""PayrollDateDiv"" ID=""PayrollDateDiv"">"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Quincena de aplicación:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""EmployeePayrollDate"" ID=""EmployeePayrollDateCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And (IsActive_7=1) And (PayrollTypeID=1)", "PayrollID Desc", aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
					Response.Write "</SELECT>&nbsp;"
					Response.Write "</TD>"
				Response.Write "</TR>"
		End Select
		Select Case lReasonID
			Case EMPLOYEES_THIRD_CONCEPT, EMPLOYEES_FONAC_CONCEPT
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de fin:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE), "ConceptEnd", Year(Date())-5, Year(Date())+5, True, True) & "</FONT></TD>"
				Response.Write "</TR>"
		End Select
		Select Case lReasonID
			Case EMPLOYEES_SERVICE_SHEET
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de hoja a generar:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""DocumentTypeID"" ID=""DocumentTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						If Len(oRequest("ServiceSheetChange").Item) > 0 Then
							Response.Write "<OPTION VALUE=0"
								If aEmployeeComponent(N_EMPLOYEE_DOCUMENT_TYPE) = 0 Then
									Response.Write " SELECTED=""1"""
								End If
							Response.Write " VALUE=0>Completa</OPTION>"
							Response.Write "<OPTION VALUE=1"
								If aEmployeeComponent(N_EMPLOYEE_DOCUMENT_TYPE) = 1 Then
									Response.Write " SELECTED=""1"""
								End If
							Response.Write " VALUE=1>Normal</OPTION>"
						Else
							Response.Write "<OPTION SELECTED VALUE=0>Completa</OPTION>"
							Response.Write "<OPTION VALUE=1>Normal</OPTION>"
						End If
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número de documento:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""DocumentNumber1"" ID=""DocumentNumber1Txt"" VALUE=""" & aEmployeeComponent(S_DOCUMENT_NUMBER_1_EMPLOYEE) & """ SIZE=""50"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
		End Select
		Response.Write "<TR NAME=""ConceptIDDiv"" ID=""ConceptIDDiv"">"
		Select Case lReasonID
			Case EMPLOYEES_THIRD_CONCEPT
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de crédito:&nbsp;</FONT></TD>"
			Case EMPLOYEES_GRADE
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Calificación:&nbsp;</FONT></TD>"
			Case EMPLOYEES_SERVICE_SHEET
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Usuarios que autorizan:&nbsp;</FONT></TD>"
			Case Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Concepto:&nbsp;</FONT></TD>"
		End Select
			Response.Write "<TD>"
				Select Case lReasonID
					Case EMPLOYEES_GRADE
					'If lReasonID = EMPLOYEES_GRADE Then
						Response.Write "<SELECT NAME=""EmployeeGrade"" ID=""EmployeeGradeCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write "<OPTION VALUE=""A"">A</OPTION>"
							Response.Write "<OPTION VALUE=""B"">B</OPTION>"
							Response.Write "<OPTION VALUE=""C"">C</OPTION>"
							Response.Write "<OPTION VALUE=""D"">D</OPTION>"
							Response.Write "<OPTION VALUE=""E"">E</OPTION>"
						Response.Write "</SELECT>&nbsp;&nbsp;"
					Case EMPLOYEES_SERVICE_SHEET
						Response.Write "&nbsp;&nbsp;<SELECT NAME=""Authorizers"" ID=""AuthorizersLst"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Users", "UserID", "UserLastName, UserName", "(UserID>=10)", "UserLastName, UserName", aEmployeeComponent(S_EMPLOYEE_AUTHORIZERS), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT>&nbsp;&nbsp;"
					Case Else
						If aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1 Then
							Response.Write "<SELECT NAME=""ConceptID"" ID=""ConceptIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							'Para asignar el ID del concepto
							Select Case lReasonID
								Case -89
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 100)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_SAFE_SEPARATION
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 120)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_ADD_SAFE_SEPARATION
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 87)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case 53
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 4)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_ANTIQUITIES
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 5)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_ADDITIONALSHIFT
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 7)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_GLASSES
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 24)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_ANTIQUITY_25_AND_30_YEARS
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 44)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)									
								Case EMPLOYEES_FAMILY_DEATH
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 45)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_PROFESSIONAL_DEGREE
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 46)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_MONTHAWARD
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 50)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_SPORTS_HELP
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 165)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_SPORTS
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 69)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_CARLOAN
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 74)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_CONCEPT_C3
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 94)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_EXTRAHOURS
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 10)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_SUNDAYS
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 17)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_BENEFICIARIES
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 70)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_BENEFICIARIES_DEBIT
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 86)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_CONCEPT_08
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 8)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_CHILDREN_SCHOOLARSHIPS
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 22)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_LICENSES
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 104)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_CONCEPT_16
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 19)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_NON_EXCENT
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 72)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_EXCENT
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 73)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_MOTHERAWARD
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 26)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_HELP_COMISSION
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 63)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_SAFEDOWN
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 67)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_ANUAL_AWARD
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 32)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_NIGHTSHIFTS
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 93)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_FONAC_CONCEPT
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 77)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_FONAC_ADJUSTMENT
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 76)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_CONCEPT_7S
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 146)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
								Case EMPLOYEES_THIRD_CONCEPT
									If CInt(oRequest("CreditChange").Item) = 0 Then
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "CreditTypes", "CreditTypeID", "CreditTypeShortName, CreditTypeName", "CreditTypeShortName <> '89'", "CreditTypeShortName, CreditTypeName", aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
									Else
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "CreditTypes", "CreditTypeID", "CreditTypeShortName, CreditTypeName", "CreditTypeID = " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE), "CreditTypeShortName, CreditTypeName", aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
									End If
								Case Else
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = 87)", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
							End Select
							Response.Write "</SELECT>"
						Else
							If lReasonID = EMPLOYEES_THIRD_CONCEPT Then
								Response.Write "<SELECT NAME=""ConceptID"" ID=""ConceptIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								If CInt(oRequest("CreditChange").Item) = 0 Then
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "CreditTypes", "CreditTypeID", "CreditTypeShortName, CreditTypeName", "", "CreditTypeShortName, CreditTypeName", aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
								Else
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "CreditTypes", "CreditTypeID", "CreditTypeShortName, CreditTypeName", "CreditTypeID = " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE), "CreditTypeShortName, CreditTypeName", aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
								End If
								Response.Write "</SELECT>"
							ElseIf (lReasonID = EMPLOYEES_SUNDAYS) Or (lReasonID = EMPLOYEES_EXTRAHOURS) Then
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptID"" ID=""ConceptIDHdn"" VALUE=""" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & """ />"
								Call GetNameFromTable(oADODBConnection, "Absences", aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT>"
							Else
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptID"" ID=""ConceptIDHdn"" VALUE=""" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & """ />"
								Call GetNameFromTable(oADODBConnection, "Concepts", aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT>"
							End If
						End If
				End Select
			Response.Write "</TD>"
		Response.Write "</TR>"
		Response.Write "<TR>"
			' Para mostrar la descripcion de las unidades para el concepto
			Select Case lReasonID
				Case EMPLOYEES_SAFE_SEPARATION, 53, EMPLOYEES_CONCEPT_08, EMPLOYEES_CONCEPT_7S
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Porcentaje:&nbsp;</FONT></TD>"
				Case -89, EMPLOYEES_ADDITIONALSHIFT, EMPLOYEES_GLASSES, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_LICENSES, EMPLOYEES_CONCEPT_16, EMPLOYEES_NON_EXCENT, EMPLOYEES_EXCENT, EMPLOYEES_ANUAL_AWARD, EMPLOYEES_MONTHAWARD, EMPLOYEES_SPORTS_HELP, EMPLOYEES_SPORTS, EMPLOYEES_CARLOAN, EMPLOYEES_CONCEPT_C3, EMPLOYEES_LICENSES, EMPLOYEES_CONCEPT_16
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Importe:&nbsp;</FONT></TD>"
				Case EMPLOYEES_ADD_SAFE_SEPARATION
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cantidad ($ o %):&nbsp;</FONT></TD>"
				Case EMPLOYEES_EXTRAHOURS
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. Horas:&nbsp;</FONT></TD>"
				Case EMPLOYEES_SUNDAYS
				Case EMPLOYEES_THIRD_CONCEPT
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Monto del Crédito:&nbsp;</FONT></TD>"
				Case EMPLOYEES_MOTHERAWARD, EMPLOYEES_HELP_COMISSION, EMPLOYEES_SAFEDOWN, EMPLOYEES_FONAC_CONCEPT, EMPLOYEES_NIGHTSHIFTS
					Response.Write "<TD COLSPAN=""2""></BR>"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>&nbsp;&nbsp;&nbsp;Concepto informado. No requiere registrar cantidad o porcentaje.</B></FONT></BR></BR>"
					Response.Write "</TD>"
				Case EMPLOYEES_GRADE, EMPLOYEES_SERVICE_SHEET
				Case Else
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cantidad:&nbsp;</FONT></TD>"
			End Select
			Response.Write "<TD>"
				' Para mostrar listas con los valores permitidos para algunos conceptos
				Select Case lReasonID
					Case EMPLOYEES_GRADE
					Case EMPLOYEES_SAFE_SEPARATION
						Response.Write "<SELECT NAME=""ConceptAmount"" ID=""ConceptAmountTxt"" SIZE=""1"" CLASS=""Lists"" />&nbsp;"
							Response.Write "<OPTION SELECTED VALUE=10>10</OPTION>"
							Response.Write "<OPTION VALUE=5>5</OPTION>"
							Response.Write "<OPTION VALUE=4>4</OPTION>"
							Response.Write "<OPTION VALUE=2>2</OPTION>"
						Response.Write "</SELECT>&nbsp;"
					Case 53
						Response.Write "<SELECT NAME=""RiskLevel"" ID=""RiskLevelCmb"" onChange=""SearchRecord(this.value + '&EmployeeID=' + EmployeeID.value, 'RiskLevel', 'RiskLevelIFrame', 'EmployeeFrm.RiskLevel');"" SIZE=""1"" CLASS=""Lists"">"
							If aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) >= 1 Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select RiskLevelID, RiskLevelName From RiskLevels where RiskLevelID > 0 Order By RiskLevelName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
							Else
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select RiskLevelID From RiskLevels Order By RiskLevelName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
							End If
							Do While Not oRecordset.EOF
								Response.Write "<OPTION VALUE=""" & CStr(oRecordset.Fields("RisklevelID").Value) & """"
								If aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = CInt(oRecordset.Fields("RisklevelID").Value) * 10 Then Response.Write " SELECTED"
								Response.Write ">" & (CInt(oRecordset.Fields("RiskLevelID").Value) * 10) & "</OPTION>"
								oRecordset.MoveNext
							Loop
						Response.Write "</SELECT>"
						Response.Write "</SELECT>&nbsp;"
					Case EMPLOYEES_FAMILY_DEATH
						Response.Write "<SELECT NAME=""ConceptAmount"" ID=""ConceptAmountTxt"" SIZE=""1"" CLASS=""Lists"" />&nbsp;"
							Response.Write "<OPTION VALUE=2800>2800</OPTION>"
						Response.Write "</SELECT>&nbsp;"
					Case EMPLOYEES_CONCEPT_08, EMPLOYEES_ADDITIONALSHIFT
						Response.Write "<SELECT NAME=""ConceptAmount"" ID=""ConceptAmountTxt"" SIZE=""1"" CLASS=""Lists"" />&nbsp;"
							Response.Write "<OPTION VALUE=46.1538>46.1538</OPTION>"
						Response.Write "</SELECT>&nbsp;"
					Case EMPLOYEES_MOTHERAWARD, EMPLOYEES_HELP_COMISSION, EMPLOYEES_SAFEDOWN, EMPLOYEES_FONAC_CONCEPT
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptAmount"" ID=""ConceptAmountHdn"" VALUE=""1"" />"
					Case EMPLOYEES_EXTRAHOURS
						Response.Write "<SELECT NAME=""ConceptAmount"" ID=""ConceptAmountTxt"" SIZE=""1"" CLASS=""Lists"" />&nbsp;"
							Response.Write "<OPTION VALUE=0.5>0.5</OPTION>"
							Response.Write "<OPTION VALUE=1>1.0</OPTION>"
							Response.Write "<OPTION VALUE=1.5>1.5</OPTION>"
							Response.Write "<OPTION VALUE=2>2.0</OPTION>"
							Response.Write "<OPTION VALUE=2.5>2.5</OPTION>"
							Response.Write "<OPTION VALUE=3>3.0</OPTION>"
						Response.Write "</SELECT>&nbsp;"
					Case EMPLOYEES_THIRD_CONCEPT
						If StrComp(oRequest("CreditChange").Item, 1, vbBinaryCompare) = 0 Then
							Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & "</FONT>&nbsp;&nbsp;"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptAmount"" ID=""ConceptAmountHdn"" VALUE=""" & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & """ />"
						Else
							Response.Write "<INPUT TYPE=""TEXT"" NAME=""ConceptAmount"" ID=""ContractNumberTxt"" SIZE=""20"" MAXLENGTH=""20"" CLASS=""TextFields"" " & "VALUE=""" & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & """/>&nbsp;"
						End If
					Case EMPLOYEES_CONCEPT_7S
						Response.Write "<SELECT NAME=""ConceptAmount"" ID=""ConceptAmountTxt"" SIZE=""1"" CLASS=""Lists"" />&nbsp;"
							Response.Write "<OPTION VALUE=1>1</OPTION>"
							Response.Write "<OPTION VALUE=2>2</OPTION>"
						Response.Write "</SELECT>&nbsp;"
					Case EMPLOYEES_SUNDAYS, EMPLOYEES_NIGHTSHIFTS
					Case EMPLOYEES_SERVICE_SHEET
					Case Else
						Response.Write "<INPUT TYPE=""TEXT"" NAME=""ConceptAmount"" ID=""ConceptAmountTxt"" SIZE=""20"" MAXLENGTH=""20"" CLASS=""TextFields"" " & "VALUE=""" & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & """/>&nbsp;"
				End Select
				' Bloque para mostrar selección de tipo de cantidad para el concepto
				Select Case lReasonID
					' Porcentaje (%)
					Case EMPLOYEES_SAFE_SEPARATION, 53, EMPLOYEES_CONCEPT_08, EMPLOYEES_ADDITIONALSHIFT, EMPLOYEES_CONCEPT_7S
						Response.Write "<SELECT NAME=""ConceptQttyID"" ID=""ConceptQttyIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowAmountFields(this.value, 'Concept');"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "QttyID=2", "QttyID", "2", "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT>&nbsp;"
						Response.Write "<SPAN ID=""ConceptCurrencySpn"">"
						Response.Write "</SPAN>"
						Response.Write "<SPAN ID=""ConceptAppliesToSpn"" STYLE=""display: none"">"
						Response.Write "</SPAN>"
					' Dinero ($)
					Case -89, EMPLOYEES_ANTIQUITIES, EMPLOYEES_GLASSES, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_LICENSES, EMPLOYEES_CONCEPT_16, EMPLOYEES_NON_EXCENT, EMPLOYEES_EXCENT, EMPLOYEES_ANUAL_AWARD, EMPLOYEES_MONTHAWARD, EMPLOYEES_SPORTS_HELP, EMPLOYEES_SPORTS, EMPLOYEES_CARLOAN, EMPLOYEES_CONCEPT_C3, EMPLOYEES_LICENSES, EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_BENEFICIARIES_DEBIT, EMPLOYEES_FONAC_ADJUSTMENT, EMPLOYEES_ANTIQUITY_25_AND_30_YEARS
						Response.Write "<SELECT NAME=""ConceptQttyID"" ID=""ConceptQttyIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowAmountFields(this.value, 'Concept');"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "QttyID=1", "QttyID", "2", "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT>&nbsp;"
						Response.Write "<SPAN ID=""ConceptCurrencySpn""><SELECT NAME=""CurrencyID"" ID=""CurrencyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Currencies", "CurrencyID", "CurrencyName", "(CurrencyID=0) And (Active=1)", "CurrencyName", aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE), "Ninguna;;;-1", sErrorDescription)
						Response.Write "</SELECT></SPAN>"
						Response.Write "<SPAN ID=""ConceptAppliesToSpn"" STYLE=""display: none"">"
						Response.Write "</SPAN>"
					' Porcentaje (%) y Dinero ($)
					Case EMPLOYEES_ADD_SAFE_SEPARATION
						'Response.Write "<SELECT NAME=""ConceptQttyID"" ID=""ConceptQttyIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowAmountFields(this.value, 'Concept');"">"
						Response.Write "<SELECT NAME=""ConceptQttyID"" ID=""ConceptQttyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "QttyID=1 or QttyID=2", "QttyID", aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT>&nbsp;"
						Response.Write "<SPAN ID=""ConceptCurrencySpn""><SELECT NAME=""CurrencyID"" ID=""CurrencyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Currencies", "CurrencyID", "CurrencyName", "(CurrencyID=0) And (Active=1)", "CurrencyName", aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE), "Ninguna;;;-1", sErrorDescription)
						Response.Write "</SELECT></SPAN>"
						Response.Write "<SPAN ID=""ConceptAppliesToSpn"" STYLE=""display: none""><SELECT NAME=""AppliesToID"" ID=""AppliesToIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID In (Select ConceptID From EmployeesConceptsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")))", "ConceptShortName, ConceptName", aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT></SPAN>"
					' Unidades
					Case EMPLOYEES_MOTHERAWARD, EMPLOYEES_HELP_COMISSION, EMPLOYEES_SAFEDOWN, EMPLOYEES_FONAC_CONCEPT
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptQttyID"" ID=""ConceptAmountHdn"" VALUE=""4"" />"
					Case EMPLOYEES_SUNDAYS, EMPLOYEES_NIGHTSHIFTS, EMPLOYEES_GRADE
					'Por horas
					Case EMPLOYEES_EXTRAHOURS
						Response.Write "<SELECT NAME=""ConceptQttyID"" ID=""ConceptQttyIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowAmountFields(this.value, 'Concept');"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "QttyID=6", "QttyID", aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT>&nbsp;"
					Case EMPLOYEES_THIRD_CONCEPT ' Para créditos de terceros
						Response.Write "<SELECT NAME=""ConceptQttyID"" ID=""ConceptQttyIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowAmountFields(this.value, 'Concept');"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "QttyID=1", "QttyID", aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT>&nbsp;"
						Response.Write "<SPAN ID=""ConceptCurrencySpn""><SELECT NAME=""CurrencyID"" ID=""CurrencyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Currencies", "CurrencyID", "CurrencyName", "(CurrencyID=0) And (Active=1)", "CurrencyName", aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE), "Ninguna;;;-1", sErrorDescription)
						Response.Write "</SELECT></SPAN>"
					' Todos los tipos
					Case EMPLOYEES_SERVICE_SHEET
					Case Else
						Response.Write "<SELECT NAME=""ConceptQttyID"" ID=""ConceptQttyIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowAmountFields(this.value, 'Concept');"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "", "QttyID", aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT>&nbsp;"
						Response.Write "<SPAN ID=""ConceptCurrencySpn""><SELECT NAME=""CurrencyID"" ID=""CurrencyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Currencies", "CurrencyID", "CurrencyName", "(CurrencyID=0) And (Active=1)", "CurrencyName", aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE), "Ninguna;;;-1", sErrorDescription)
						Response.Write "</SELECT></SPAN>"
						Response.Write "<SPAN ID=""ConceptAppliesToSpn"" STYLE=""display: none""><SELECT NAME=""AppliesToID"" ID=""AppliesToIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID In (Select ConceptID From EmployeesConceptsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")))", "ConceptShortName, ConceptName", aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT></SPAN>"
				End Select
			Response.Write "</TD>"
		Response.Write "</TR>"
		If lReasonID = 53 Then
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP"" COLSPAN=""2""><IFRAME SRC=""SearchRecord.asp"" NAME=""RiskLevelIFrame"" FRAMEBORDER=""0"" WIDTH=""300"" HEIGHT=""100""></IFRAME></TD>"
			Response.Write "</TR>"
		End If
		Select Case lReasonID
			Case EMPLOYEES_BENEFICIARIES_DEBIT
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Beneficiario:&nbsp;</FONT></TD>"
					'Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""PaymentsNumber"" ID=""PaymentsNumberTxt"" SIZE=""20"" MAXLENGTH=""20"" CLASS=""TextFields"" " & "VALUE=""" & aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE) & """/>&nbsp;"
					Response.Write "<TD><SELECT NAME=""BeneficiaryID"" ID=""BeneficiaryCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EmployeesBeneficiariesLKP", "BeneficiaryID", "BeneficiaryNumber, BeneficiaryName, BeneficiaryLastName", "(EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "BeneficiaryNumber, BeneficiaryName, BeneficiaryLastName", "", "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT>&nbsp;"
					Response.Write "</TD>"
				Response.Write "</TR>"
			Case EMPLOYEES_THIRD_CONCEPT ' Para créditos de terceros
				If StrComp(oRequest("CreditChange").Item, 1, vbbinaryCompare) = 0 Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cantidad de pagos:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE) & "</FONT></TD>"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PaymentsNumber"" ID=""PaymentsNumberHdn"" VALUE=""" & aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE) & """ />"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número de contrato:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_CREDIT_CONTRACT_NUMBER_EMPLOYEE) & "</FONT></TD>"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ContractNumber"" ID=""ContractNumberHdn"" VALUE=""" & aEmployeeComponent(S_CREDIT_CONTRACT_NUMBER_EMPLOYEE) & """ />"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número de cuenta:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_CREDIT_ACCOUNT_NUMBER_EMPLOYEE) & "</FONT></TD>"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AccountNumber"" ID=""AccountNumberHdn"" VALUE=""" & aEmployeeComponent(S_CREDIT_ACCOUNT_NUMBER_EMPLOYEE) & """ />"
					Response.Write "</TR>"
				Else
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cantidad de pagos:&nbsp;</FONT></TD>"
						Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""PaymentsNumber"" ID=""PaymentsNumberTxt"" SIZE=""20"" MAXLENGTH=""20"" CLASS=""TextFields"" " & "VALUE=""" & aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE) & """/>&nbsp;"
						Response.Write "<SELECT NAME=""PaymentQttyID"" ID=""PaymentQttyIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowAmountFields(this.value, 'Concept');"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "QttyID=17", "QttyID", aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT>&nbsp;"
						Response.Write "</TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número de contrato:&nbsp;</FONT></TD>"
						Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""ContractNumber"" ID=""ContractNumberTxt"" SIZE=""20"" MAXLENGTH=""20"" CLASS=""TextFields"" " & "VALUE=""" & aEmployeeComponent(S_CREDIT_CONTRACT_NUMBER_EMPLOYEE) & """/>&nbsp;"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número de cuenta:&nbsp;</FONT></TD>"
						Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""AccountNumber"" ID=""AccountNumberTxt"" SIZE=""20"" MAXLENGTH=""20"" CLASS=""TextFields"" " & "VALUE=""" & aEmployeeComponent(S_CREDIT_ACCOUNT_NUMBER_EMPLOYEE) & """/>&nbsp;"
					Response.Write "</TR>"
				End If
				If False Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Periodo:&nbsp;</FONT></TD>"
						Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""PeriodID"" ID=""PeriodIDTxt"" SIZE=""20"" MAXLENGTH=""20"" CLASS=""TextFields"" " & "VALUE=""" & aEmployeeComponent(N_CREDIT_PERIOD_ID_EMPLOYEE) & """/>&nbsp;"
					Response.Write "</TR>"
					Response.Write "<TR NAME=""PayrollDateDiv"" ID=""PayrollDateDiv"">"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Periodo:&nbsp;</FONT></TD>"
						Response.Write "<TD><SELECT NAME=""PeriodID"" ID=""PeriodID"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And (IsActive_1=1)", "PayrollID Desc", "", "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
						Response.Write "</SELECT>&nbsp;"
						Response.Write "</TD>"
					Response.Write "</TR>"
				End If
		End Select
		Select Case lReasonID
			Case EMPLOYEES_NIGHTSHIFTS, EMPLOYEES_GRADE
			Case Else
				Response.Write "<TR><TD COLSPAN=""2"">"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Observaciones:<BR /></FONT>"
					Response.Write "<TEXTAREA NAME=""ConceptComments"" ID=""ConceptCommentsTxtArea"" ROWS=""5"" COLS=""40"" MAXLENGTH=""2000"" CLASS=""TextFields"">" & aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) & "</TEXTAREA>"
				Response.Write "</TD></TR>"
		End Select
	Response.Write "</TABLE>"
	'Select Case lReasonID
	'	Case EMPLOYEES_SAFE_SEPARATION, EMPLOYEES_ADD_SAFE_SEPARATION, CANCEL_EMPLOYEES_SSI
	'		Response.Write "<IFRAME SRC=""BrowserFile.asp?EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & """ NAME=""EmployeeFilesIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""348""></IFRAME>"
	'	Case Else
	'End Select
	If False Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "ShowAmountFields(document.EmployeeFrm.ConceptQttyID.value, 'Concept');" & vbNewLine
			If Len(sURL) > 0 Then
				Response.Write "SendURLValuesToForm('" & sURL & "', document.ConceptFrm);" & vbNewLine
			End If
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If
	If False Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "ShowHideAbsencesFields(document.EmployeeFrm.AbsenceID.value);" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If
End Function

Function DisplayEmployeeFormSection129(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 129
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection129"

	Response.Write "<BR /><BR /><FONT FACE=""Arial"" SIZE=""2""><B>Registre la ausencia</B></FONT>"
	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
	Response.Write "<IFRAME SRC=""BrowserMonthForPayments.asp?EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & "&EmployeeDate=" & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & "&FromArrow=1"" NAME=""BrowserMonthIFrame"" FRAMEBORDER=""0"" WIDTH=""300"" HEIGHT=""112""></IFRAME>"
	'Response.Write "<FORM NAME=""AbsencesFrm"" ID=""AbsencesFrm"" ACTION=""" & sAction & """ METHOD=""GET"" onSubmit=""return CheckAbsenceFields(this)"">"
		sNames = oRequest("Action").Item
		If Len(sNames) = 0 Then sNames = "Absences"
		'Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & sNames & """ />"
		'Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & """ />"
		'Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Tab"" ID=""TabHdn"" VALUE=""4"" />"
		'Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OcurredDate"" ID=""OcurredDateHdn"" VALUE="""
		'	If (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) > -1) And (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) > -1) And (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) > 0) Then Response.Write aAbsenceComponent(N_OCURRED_DATE_ABSENCE)
		'Response.Write """ />"
		'If (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) > -1) And (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) > -1) And (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) > 0) Then
		'	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RegistrationDate"" ID=""RegistrationDateHdn"" VALUE=""" & aAbsenceComponent(N_REGISTRATION_DATE_ABSENCE) & """ />"
		'End If
		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			If StrComp(GetASPFileName(""), "Employees.asp", vbBinaryCompare) <> 0 Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. del empleado:&nbsp;</FONT></TD>"
					Response.Write "<TD>"
						Response.Write "<INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" SIZE=""20"" VALUE="""" CLASS=""TextFields"" onChange=""document.AbsencesFrm.EmployeeID.value='';"" />"
						Response.Write "<A HREF=""javascript: document.EmployeeFrm.EmployeeID.value=''; SearchRecord(document.AbsencesFrm.EmployeeNumber.value, 'EmployeeNumber', 'SearchEmployeeNumberIFrame', 'EmployeeFrm.EmployeeID')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A>&nbsp;"
						Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""22""></IFRAME>"
					Response.Write "</TD>"
				Response.Write "</TR>"
			End If
			If Not B_ISSSTE Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""AbsenceDate"" ID=""AbsenceDateTxt"" SIZE=""50"" VALUE="""" CLASS=""SpecialTextFields"" onFocus=""document.EmployeeFrm.AbsenceID.focus()"" /></TD>"
				Response.Write "</TR>"
			Else
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AbsenceDate"" ID=""AbsenceDateHdn"" VALUE="""" />"
			End If
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Concepto:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""AbsenceID"" ID=""AbsenceIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowHideAbsencesFields(this.value)"">"
					If Len(sAbsenceIDs) = 0 Then
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Absences", "AbsenceID", "AbsenceShortName, AbsenceName", "(Active=1)", "AbsenceShortName", aAbsenceComponent(N_ABSENCE_ID_ABSENCE), "Ninguno;;;-1", sErrorDescription)
					Else
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Absences", "AbsenceID", "AbsenceShortName, AbsenceName", "(Active=1) And (AbsenceID In (" & sAbsenceIDs & "))", "AbsenceShortName", sAbsenceIDs, "Ninguno;;;-1", sErrorDescription)
					End If
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR NAME=""AbsenceHoursDiv"" ID=""AbsenceHoursDiv"" STYLE=""display: none"">"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Horas:&nbsp;</NOBR></FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""AbsenceHours"" ID=""AbsenceHoursTxt"" SIZE=""2"" MAXLENGTH=""2"" VALUE=""" & aAbsenceComponent(N_HOURS_ABSENCE) & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			If Not B_ISSSTE Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>No. de oficio:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""DocumentNumber"" ID=""DocumentNumberTxt"" SIZE=""30"" MAXLENGTH=""50"" VALUE=""" & aAbsenceComponent(S_DOCUMENT_NUMBER_ABSENCE) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
			Else
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DocumentNumber"" ID=""DocumentNumberHdn"" VALUE=""."" />"
			End If
			Response.Write "<TR><TD COLSPAN=""2"">"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Observaciones:<BR /></FONT>"
				Response.Write "<TEXTAREA NAME=""Reasons"" ID=""ReasonsTxtArea"" ROWS=""5"" COLS=""40"" MAXLENGTH=""2000"" VALUE="""" CLASS=""TextFields"">" & aAbsenceComponent(S_REASONS_ABSENCE) & "</TEXTAREA>"
			Response.Write "</TD></TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Activo:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
					Response.Write "<INPUT TYPE=""RADIO"" NAME=""Active"" ID=""ActiveRd"" VALUE=""1"" "
						If aEmployeeComponent(N_ACTIVE_ABSENCE) = 1 Then Response.Write " CHECKED=""1"""
					Response.Write " />Sí"
					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""1"" />"
					Response.Write "<INPUT TYPE=""RADIO"" NAME=""Active"" ID=""ActiveRd"" VALUE=""0"""
						If aEmployeeComponent(N_ACTIVE_ABSENCE) = 0 Then Response.Write " CHECKED=""1"""
					Response.Write " />No"
				Response.Write "</FONT></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
		If B_ISSSTE Then
			If (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = 0) Then
				Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
					Response.Write "<TR>"
						Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fechas de las incidencias:&nbsp;</FONT></TD>"
						Response.Write "<TD VALIGN=""TOP""><SELECT NAME=""OcurredDates"" ID=""OcurredDatesLst"" SIZE=""5"" MULTIPLE=""1""></SELECT></TD>"
						Response.Write "<TD VALIGN=""BOTTOM""><A HREF=""javascript: RemoveSelectedItemsFromList(null, document.AbsencesFrm.OcurredDates)""><IMG SRC=""Images/BtnCrclDelete.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Eliminar"" BORDER=""0"" HSPACE=""5"" /></A></TD>"
					Response.Write "</TR>"
				Response.Write "</TABLE>"
			End If
		End If

		Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""340"" HEIGHT=""1"" /><BR /><BR />"

		If (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = 0) Then
			If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
		ElseIf Len(oRequest("Delete").Item) > 0 Then
			If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveAbsenceWngDiv']); AbsencesFrm.Remove.focus()"" />"
		Else
			If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
		End If
		Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
		Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=Users'"" />"
		Response.Write "<BR /><BR />"
		Call DisplayWarningDiv("RemoveAbsenceWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
	'Response.Write "</FORM>"
	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		Response.Write "ShowDisplay(document.all['AbsenceHoursDiv']);" & vbNewLine
	Response.Write "//--></SCRIPT>" & vbNewLine					
					
	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		Response.Write "ShowHideAbsencesFields(document.AbsencesFrm.AbsenceID.value);" & vbNewLine
	Response.Write "//--></SCRIPT>" & vbNewLine
End Function

Function DisplayEmployeeFormSection130()
'************************************************************
'Purpose: To add JavaScript to validate amounts taken
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection130"

	Response.Write "function CheckAbsenceFields(oForm) {" & vbNewLine
		If Len(oRequest("Delete").Item) = 0 Then
			Response.Write "if (oForm) {" & vbNewLine
				If Not B_ISSSTE Then
					Response.Write "if (oForm.DocumentNumber.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de introducir el número de folio.');" & vbNewLine
						Response.Write "oForm.DocumentNumber.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
				Else
					If StrComp(GetASPFileName(""), "Employees.asp", vbBinaryCompare) <> 0 Then
						Response.Write "if ((oForm.EmployeeID.value.length == 0) || (oForm.EmployeeID.value == '')) {" & vbNewLine
							Response.Write "alert('Favor de especificar el número de empleado.');" & vbNewLine
							Response.Write "oForm.EmployeeID.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					End If
					Response.Write "SelectAllItemsFromList(oForm.OcurredDates);" & vbNewLine
				End If
				Response.Write "if (!CheckIntegerValue(oForm.AbsenceHours, 'las horas', N_MINIMUM_ONLY_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
					Response.Write "return false;" & vbNewLine
			Response.Write "}" & vbNewLine
		End If
		Response.Write "return true;" & vbNewLine
	Response.Write "} // End of CheckAbsenceFields" & vbNewLine

	Response.Write "function ShowHideAbsencesFields(sValue) {" & vbNewLine
		Response.Write "var oForm = document.AbsencesFrm" & vbNewLine
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
End Function

Function DisplayEmployeeFormSection131(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 131
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection131"

	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Registre los datos del empleado:</B></FONT>"
	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del empleado:&nbsp;</FONT></TD>"
				If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" VALUE=""" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""SpecialTextFields"" onFocus=""document.EmployeeFrm.EmployeeName.focus()"" /></TD>"
				Else
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "</FONT><INPUT TYPE=""HIDDEN"" NAME=""EmployeeNumber"" ID=""EmployeeNumberHdn"" VALUE=""" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & """ /></TD>"
				End If
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. de oficio:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""DocumentForLicenseNumber"" ID=""DocumentForLicenseNumberTxt"" SIZE=""6"" MAXLENGTH=""25"" VALUE=""" & aEmployeeComponent(S_DOCUMENT_FOR_LICENSE_NUMBER_EMPLOYEE) & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			If Len(oRequest("ModifyDocumentLicense").Item) <> 0 Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. de oficio de cancelación:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""DocumentForCancelLicenseNumber"" ID=""DocumentForCancelLicenseNumberTxt"" SIZE=""6"" MAXLENGTH=""25"" VALUE=""" & aEmployeeComponent(S_DOCUMENT_FOR_CANCEL_LICENSE_NUMBER_EMPLOYEE) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
			Else
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DocumentForCancelLicenseNumber"" ID=""DocumentForCancelLicenseNumberHdn"" VALUE=""" & aEmployeeComponent(S_DOCUMENT_FOR_CANCEL_LICENSE_NUMBER_EMPLOYEE) & """ />"
			End If
			Response.Write "<TR>"
			    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha del documento:&nbsp;</FONT></TD>"
			    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_DATE_LICENSE_DOCUMENT_EMPLOYEE), "DocumentLicense", N_FORM_START_YEAR, Year(Date()), True, True) & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. de la solicitud:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""RequestNumber"" ID=""RequestNumberTxt"" SIZE=""10"" MAXLENGTH=""25"" VALUE=""" & aEmployeeComponent(S_REQUEST_NUMBER_EMPLOYEE) & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
			    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de Licencia:&nbsp;</FONT></TD>"
			    Response.Write "<TD><SELECT NAME=""LicenseSyndicateTypeID"" ID=""LicenseSyndicateTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
			        Response.Write GenerateListOptionsFromQuery(oADODBConnection, "LicenseSyndicateTypes", "LicenseSyndicateTypeID", "LicenseSyndicateTypeName", "", "LicenseSyndicateTypeID", aEmployeeComponent(N_SYNDICATE_TYPE_ID_LICENSE_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
			    Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
			    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio de la licencia:&nbsp;</FONT></TD>"
			    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_LICENSE_START_DATE_EMPLOYEE), "LicenseStart", Year(Date()) - 1, Year(Date()) + 1, True, True) & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
			    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de fin de la licencia:&nbsp;</FONT></TD>"
			    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_LICENSE_END_DATE_EMPLOYEE), "LicenseEnd", Year(Date()), Year(Date()) + 1, True, True) & "</FONT></TD>"
			Response.Write "</TR>"
			If Len(oRequest("ModifyDocumentLicense").Item) <> 0 Then
				Response.Write "<TR>"
				    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de cancelación de la licencia:&nbsp;</FONT></TD>"
				    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_CANCEL_DATE_LICENSE_DOCUMENT_EMPLOYEE), "LincenseCancel", N_FORM_START_YEAR, Year(Date()), True, True) & "</FONT></TD>"
				Response.Write "</TR>"
			Else
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""LincenseCancelDate"" ID=""LincenseCancelDateHdn"" VALUE=""" & aEmployeeComponent(N_CANCEL_DATE_LICENSE_DOCUMENT_EMPLOYEE) & """ />"
			End If
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Plantilla del documento:&nbsp;</FONT></TD>"
				Response.Write "<TD COLSPAN=""2""><SELECT NAME=""DocumentTemplate"" ID=""DocumentTemplateCmb"" SIZE=""1"" CLASS=""Lists"" >"
					If Len(oRequest("ModifyDocumentLicense").Item) = 0 Then
						Response.Write "<OPTION VALUE=""1. SNTISSSTE.htm"">1. SNTISSSTE.htm</OPTION>"
						Response.Write "<OPTION VALUE=""2. FSTSE.htm"">2. FSTSE.htm</OPTION>"
					Else
						Response.Write "<OPTION VALUE=""1. SNTISSSTE.htm"">1. SNTISSSTE.htm</OPTION>"
						Response.Write "<OPTION VALUE=""2. FSTSE.htm"">2. FSTSE.htm</OPTION>"
						Response.Write "<OPTION VALUE=""3. Cancela SNTISSSTE.htm"">3. Cancela SNTISSSTE.htm</OPTION>"
						Response.Write "<OPTION VALUE=""4. Cancela FSTSE.htm"">4. Cancela FSTSE.htm</OPTION>"
					End If
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
	Response.Write "</TABLE>"
	Response.Write "<BR /><BR />"
End Function

Function DisplayEmployeeFormSection133(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 133
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection133"

	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		If Len(oRequest("BeneficiaryChange").Item) > 0 Then
			Response.Write "var bCheckAmount = false;" & vbNewLine
		Else
			Response.Write "var bCheckAmount = true;" & vbNewLine
		End If
	Response.Write "//--></SCRIPT>" & vbNewLine
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		Select Case lReasonID
			Case EMPLOYEES_ADD_BENEFICIARIES
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Número del beneficiario:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""BeneficiaryNumber"" ID=""BeneficiaryNumberTxt"" VALUE="""
						If Len(aEmployeeComponent(S_NUMBER_BENEFICIARY_EMPLOYEE)) > 0 Then Response.Write aEmployeeComponent(S_NUMBER_BENEFICIARY_EMPLOYEE)
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
			Case EMPLOYEES_CREDITORS
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Número del acreedor:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""CreditorNumber"" ID=""CreditorNumberTxt"" VALUE="""
						If Len(aEmployeeComponent(S_NUMBER_CREDITOR_EMPLOYEE)) > 0 Then Response.Write aEmployeeComponent(S_NUMBER_CREDITOR_EMPLOYEE)
					Response.Write """ SIZE=""10"" MAXLENGTH=""10"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Nombre del acreedor:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""CreditorName"" ID=""CreditorNameTxt"" VALUE=""" & aEmployeeComponent(S_NAME_CREDITOR_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Apellido paterno:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""CreditorLastName"" ID=""CreditorLastNameTxt"" VALUE=""" & aEmployeeComponent(S_LAST_NAME_CREDITOR_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Apellido materno:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""CreditorLastName2"" ID=""CreditorLastName2Txt"" VALUE=""" & aEmployeeComponent(S_LAST_NAME2_CREDITOR_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
		End Select
		If Len(oRequest("BeneficiaryChange").Item) > 0 Then
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
				Select Case lReasonID
					Case EMPLOYEES_ADD_BENEFICIARIES
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aEmployeeComponent(N_START_DATE_BENEFICIARY_EMPLOYEE), -1, -1, -1) & "</FONT></TD>"
					Case EMPLOYEES_CREDITORS
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aEmployeeComponent(N_START_DATE_CREDITOR_EMPLOYEE), -1, -1, -1) & "</FONT></TD>"
				End Select
			Response.Write "</TR>"
		Else
			Select Case lReasonID
				Case EMPLOYEES_ADD_BENEFICIARIES
					Response.Write "<TR NAME=""PayrollDateDiv"" ID=""PayrollDateDiv"">"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Quincena de aplicación:&nbsp;</FONT></TD>"
						Response.Write "<TD><SELECT NAME=""BeneficiaryStartDate"" ID=""BeneficiaryStartDateCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And (IsActive_7=1) And (PayrollTypeID=1)", "PayrollID Desc", aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
						Response.Write "</SELECT>&nbsp;"
						Response.Write "</TD>"
					Response.Write "</TR>"
				Case EMPLOYEES_CREDITORS
					Response.Write "<TR NAME=""PayrollDateDiv"" ID=""PayrollDateDiv"">"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Quincena de aplicación:&nbsp;</FONT></TD>"
						Response.Write "<TD><SELECT NAME=""CreditorStartDate"" ID=""CreditorStartDateCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And (IsActive_7=1) And (PayrollTypeID=1)", "PayrollID Desc", aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
						Response.Write "</SELECT>&nbsp;"
						Response.Write "</TD>"
					Response.Write "</TR>"
			End Select
		End If
		Select Case lReasonID
			Case EMPLOYEES_ADD_BENEFICIARIES
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Fecha de término:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_END_DATE_BENEFICIARY_EMPLOYEE), "BeneficiaryEnd", N_FORM_START_YEAR, Year(Date()) + 10, True, True) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Centro de pago:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><SELECT NAME=""BeneficiaryPaymentCenterID"" ID=""BeneficiaryPaymentCenterIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas, Areas As PaymentCenters, Zones", "PaymentCenters.AreaID", "PaymentCenters.AreaCode, PaymentCenters.AreaName, Areas.AreaCode, Areas.AreaName, Zones.ZoneCode, Zones.ZoneName", "(Areas.AreaID=PaymentCenters.ParentID) And (PaymentCenters.Active=1) And (PaymentCenters.ParentID <> -1) And (Areas.AreaCode <> '00') And (PaymentCenters.ZoneID=Zones.ZoneID)", "Areas.AreaCode, PaymentCenters.AreaShortName", aEmployeeComponent(N_PAYMENT_CENTER_ID_BENEFICIARY_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Tipo de pensión:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><SELECT NAME=""AlimonyTypeID"" ID=""AlimonyTypeIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowAmountFields(this.value, 'Concept');"">"
					'Response.Write "<TD><SELECT NAME=""AlimonyTypeID"" ID=""AlimonyTypeIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowAmountFields(this.value, 'Concept');"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "AlimonyTypes", "AlimonyTypeID", "AlimonyTypeID As Temp1, AlimonyTypeName", "", "AlimonyTypeID", aEmployeeComponent(N_ALIMONY_TYPE_ID_BENEFICIARY_EMPLOYEE), "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><NOBR>Monto de la pensión alimenticia:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						Response.Write "<INPUT TYPE=""TEXT"" NAME=""ConceptAmount"" ID=""ConceptAmountTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & FormatNumber(aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE), 2, True, False, True) & """ CLASS=""TextFields"" />&nbsp;"
						Response.Write "<SELECT NAME=""ConceptQttyID"" ID=""ConceptQttyIDCmb"" SIZE=""1"" CLASS=""Lists""></SELECT>"
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--ShowAmountFields(document.EmployeeBeneficiaryFrm.AlimonyTypeID.value, 'Concept');//--></SCRIPT>"
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
			Case EMPLOYEES_CREDITORS
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Fecha de término:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_END_DATE_CREDITOR_EMPLOYEE), "CreditorEnd", N_FORM_START_YEAR, Year(Date()) + 10, True, True) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Centro de pago:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><SELECT NAME=""CreditorPaymentCenterID"" ID=""CreditorPaymentCenterIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas, Areas As PaymentCenters, Zones", "PaymentCenters.AreaID", "PaymentCenters.AreaCode, PaymentCenters.AreaName, Areas.AreaCode, Areas.AreaName, Zones.ZoneCode, Zones.ZoneName", "(Areas.AreaID=PaymentCenters.ParentID) And (PaymentCenters.Active=1) And (PaymentCenters.ParentID <> -1) And (Areas.AreaCode <> '00') And (PaymentCenters.ZoneID=Zones.ZoneID)", "Areas.AreaCode, PaymentCenters.AreaShortName", aEmployeeComponent(N_PAYMENT_CENTER_ID_CREDITOR_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Tipo de descuento:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><SELECT NAME=""CreditorTypeID"" ID=""CreditorTypeIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowAmountFields(this.value, 'Concept');"">"
					'Response.Write "<TD><SELECT NAME=""AlimonyTypeID"" ID=""AlimonyTypeIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowAmountFields(this.value, 'Concept');"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "CreditorsTypes", "CreditorTypeID", "CreditorTypeID As Temp1, CreditorTypeName", "", "CreditorTypeID", aEmployeeComponent(N_CREDITOR_TYPE_ID_EMPLOYEE), "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><NOBR>Monto del descuento para el trabajador:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						Response.Write "<INPUT TYPE=""TEXT"" NAME=""ConceptAmount"" ID=""ConceptAmountTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & FormatNumber(aEmployeeComponent(D_CREDITOR_AMOUNT_EMPLOYEE), 2, True, False, True) & """ CLASS=""TextFields"" />&nbsp;"
						Response.Write "<SELECT NAME=""ConceptQttyID"" ID=""ConceptQttyIDCmb"" SIZE=""1"" CLASS=""Lists""></SELECT>"
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--ShowAmountFields(document.EmployeeBeneficiaryFrm.CreditorTypeID.value, 'Concept');//--></SCRIPT>"
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
		End Select
			Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><NOBR>Monto mínimo:&nbsp;</NOBR></FONT></TD>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
				Response.Write "<INPUT TYPE=""TEXT"" NAME=""ConceptMin"" ID=""ConceptMinTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & FormatNumber(aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE), 2, True, False, True) & """ CLASS=""TextFields"" />&nbsp;"
				Response.Write "<SELECT NAME=""ConceptMinQttyID"" ID=""ConceptMinQttyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "(QttyID In (1))", "QttyID", aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
				Response.Write "</SELECT>"
			Response.Write "</FONT></TD>"
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><NOBR>Monto máximo:&nbsp;</NOBR></FONT></TD>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
				Response.Write "<INPUT TYPE=""TEXT"" NAME=""ConceptMax"" ID=""ConceptMaxTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & FormatNumber(aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE), 2, True, False, True) & """ CLASS=""TextFields"" />&nbsp;"
				Response.Write "<SELECT NAME=""ConceptMaxQttyID"" ID=""ConceptMaxQttyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "(QttyID In (1))", "QttyID", aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
				Response.Write "</SELECT>"
			Response.Write "</FONT></TD>"
		Response.Write "</TR>"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptTypeID"" ID=""ConceptTypeIDHdn"" VALUE=""" & aEmployeeComponent(N_CONCEPT_TYPE_ID_CONCEPT) & """ />"
		Select Case lReasonID
			Case EMPLOYEES_ADD_BENEFICIARIES
				Response.Write "<TR><TD COLSPAN=""2""><NOBR>"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Comentarios:<BR />"
					Response.Write "<TEXTAREA NAME=""BeneficiaryComments"" ID=""BeneficiaryCommentsTxtArea"" ROWS=""5"" COLS=""60"" MAXLENGTH=""2000"" CLASS=""TextFields"">" & aEmployeeComponent(S_COMMENTS_BENEFICIARY_EMPLOYEE) & "</TEXTAREA>"
				Response.Write "</NOBR></TD></TR>"
			Case EMPLOYEES_CREDITORS
				Response.Write "<TR><TD COLSPAN=""2""><NOBR>"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Comentarios:<BR />"
					Response.Write "<TEXTAREA NAME=""CreditorComments"" ID=""CreditorCommentsTxtArea"" ROWS=""5"" COLS=""60"" MAXLENGTH=""2000"" CLASS=""TextFields"">" & aEmployeeComponent(S_COMMENTS_CREDITOR_EMPLOYEE) & "</TEXTAREA>"
				Response.Write "</NOBR></TD></TR>"
		End Select
	Response.Write "</TABLE><BR />"
	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		If Len(sURL) > 0 Then
			Response.Write "SendURLValuesToForm('" & sURL & "', document.EmployeeBeneficiaryFrm);" & vbNewLine
		End If
	Response.Write "//--></SCRIPT>" & vbNewLine
End Function

Function DisplayEmployeeFormSection134(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 134
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection134"

	If lReasonID = 0 Then
		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""400"">&nbsp;</TD>"
						Response.Write "<TD VALIGN=""TOP"" WIDTH=""32""><A HREF=""UploadInfo.asp?Action=EmployeesAssignNumber&ReasonID=0""><IMG SRC=""Images/MnLeftArrows.gif"" WIDTH=""32"" HEIGHT=""32"" ALT=""Prestación"" BORDER=""0"" /></A><BR /></TD>"
						Response.Write "<TD VALIGN=""TOP"" WIDTH=""290""><FONT FACE=""Arial"" SIZE=""2""><B><A HREF=""UploadInfo.asp?Action=EmployeesAssignNumber&ReasonID=0"" CLASS=""SpecialLink"">Otro empleado</A></B><BR /></FONT>"
						Response.Write "<DIV CLASS=""MenuOverflow""><FONT FACE=""Arial"" SIZE=""2"">Registre el movimiento a un empleado diferente.</FONT></DIV></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	End If
	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Información general</B></FONT>"
	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del empleado:&nbsp;</FONT></TD>"
			If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" VALUE=""" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""SpecialTextFields"" onFocus=""document.EmployeeFrm.EmployeeName.focus()"" /></TD>"
			Else
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "</FONT><INPUT TYPE=""HIDDEN"" NAME=""EmployeeNumber"" ID=""EmployeeNumberHdn"" VALUE=""" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & """ /></TD>"
			End If
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre(s):&nbsp;</FONT></TD>"
			Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeName"" ID=""EmployeeNameTxt"" VALUE=""" & oRequest("sEmployeeName") & """ SIZE=""30"" MAXLENGTH=""100""" & sReadOnly & "CLASS=""TextFields"" /></TD>"
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido paterno:&nbsp;</FONT></TD>"
			Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName"" ID=""EmployeeLastNameTxt"" VALUE=""" & oRequest("sEmployeeLastName") & """ SIZE=""30"" MAXLENGTH=""100""" & sReadOnly & "CLASS=""TextFields"" /></TD>"
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido materno:&nbsp;</FONT></TD>"
			Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName2"" ID=""EmployeeLastName2Txt"" VALUE=""" & oRequest("sEmployeeLastName2") & """ SIZE=""30"" MAXLENGTH=""100""" & sReadOnly & "CLASS=""TextFields"" /></TD>"
		Response.Write "</TR>"
	Response.Write "</TABLE>"
	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de tabulador:&nbsp;</FONT></TD>"
			Response.Write "<TD><SELECT NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDCmb"" SIZE=""1"" CLASS=""Lists"""
				If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then 
					Response.Write " onChange=""GetEmployeeNumber(this.value)"""
					Response.Write ">"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EmployeeTypes", "EmployeeTypeID", "EmployeeTypeID As RecordID, EmployeeTypeName", "(Active=1)", "EmployeeTypeName", aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
				Else
					Response.Write ">"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EmployeeTypes", "EmployeeTypeID", "EmployeeTypeID As RecordID, EmployeeTypeName", "(Active=1) And (EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ")", "EmployeeTypeName", aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE), "Ninguno;;;-1", sErrorDescription)
				End If
			Response.Write "</SELECT></TD>"
		Response.Write "</TR>"
		If (lReasonID <> 58) Then
			Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Antigüedad ISSSTE:&nbsp;</FONT></TD>"
			Call GetNameFromTable(oADODBConnection, "Antiquities", aEmployeeComponent(N_ANTIQUITY_EMPLOYEE), "", "", sEmployeeDisplayFormAntiquity, sErrorDescription)
			Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeAntiquity"" ID=""EmployeeAntiquityTxt"" VALUE=""" & sEmployeeDisplayFormAntiquity & """ SIZE=""30"" MAXLENGTH=""100""" & sReadOnly & "CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
		End If
	Response.Write "</TABLE>"
	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">RFC:&nbsp;</FONT></TD>"
			Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""RFC"" ID=""RFCTxt"" VALUE=""" & oRequest("sRFC") & """ SIZE=""13"" MAXLENGTH=""13""" & sReadOnly & "CLASS=""TextFields"" /></TD>"
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">CURP:&nbsp;</FONT></TD>"
			Response.Write "<TD>"
				Response.Write "<INPUT TYPE=""TEXT"" NAME=""CURP"" ID=""CURPTxt"" VALUE=""" & oRequest("sCURP") & """ SIZE=""18"" MAXLENGTH=""18""" & sReadOnly & "CLASS=""TextFields"" />"
				Response.Write "<A HREF=""javascript: OpenNewWindow('http://consultas.curp.gob.mx/CurpSP/curp2.do?strCurp=' + document.EmployeeFrm.CURP.value + '&strTipo=B&entfija=DF&depfila=09020', null, 'CurpInversa', 400, 600, 'yes', 'yes');""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Consultar CURP"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A>"
			Response.Write "</TD>"
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de nacimiento:&nbsp;</FONT></TD>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE), "Birth", N_FORM_START_YEAR, Year(Date()), True, True) & "</FONT></TD>"
		Response.Write "</TR>"
	Response.Write "</TABLE>"
	Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR />"
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		Response.Write "<TR>"
			Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""1""><B>Nota:&nbsp;&nbsp;&nbsp;&nbsp;</B></FONT></TD>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">Es importante que verifique que el empleado de nuevo ingreso no se<BR />encuentre registrado en el Padrón Nacional de Inhabilitados por la<BR />Secretaría de la Función Pública.</FONT></TD>"
		Response.Write "</TR>"
	Response.Write "</TABLE>"
	Response.Write "<BR /><BR />"
End Function

Function DisplayEmployeeFormSection136(oRequest, oADODBConnection, sAction, sURL, lEmployeeID, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add JavaScript to display section 136
'         in the form
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFormSection136"

	Call GetNameFromTable(oADODBConnection, "Reasons", lReasonID, "", "", sNames, "")
	If ((lReasonID = CANCEL_EMPLOYEES_CONCEPTS) Or (lReasonID = CANCEL_EMPLOYEES_C04)) And (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1) Then
		Response.Write "<BR /><BR />"
		Response.Write "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD>"
					Call DisplayInstructionsMessage("Mensaje del sistema", "Seleccione una prestación en la opción 2 para cerrar sus efectos con la acción flecha curva o para cancelarla con la opción tache. Si la cancelación tiene efectos retroactivos requerira introducir revisión(des) de nómina(s) al empleado.")
				Response.Write "</TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
		Response.Write "<BR /><BR />"
	Else
		lErrorNumber = GetEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptAmount"" ID=""ConceptAmountHdn"" VALUE=""" & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & """ />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptActive"" ID=""ConceptActiveHdn"" VALUE=""" & aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) & """ />"
		Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Registro de: " & sNames & "</B></FONT>"
		Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			If (lReasonID = CANCEL_EMPLOYEES_CONCEPTS) Or (lReasonID = CANCEL_EMPLOYEES_C04) Or (lReasonID = CANCEL_EMPLOYEES_SSI) Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), -1, -1, -1) & "</FONT></TD>"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptStartDate"" ID=""ConceptStartDateHdn"" VALUE=""" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & """ />"
				Response.Write "</TR>"
			Else
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
					Response.Write "<TD><IFRAME SRC=""SearchRecord.asp"" NAME=""SearchConceptStartDateIFrame"" FRAMEBORDER=""0"" WIDTH=""600"" HEIGHT=""20""></IFRAME></TD>"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptStartDate"" ID=""ConceptStartDateHdn"" VALUE="""""" />"
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
			End If
			If ((lReasonID = CANCEL_EMPLOYEES_CONCEPTS) Or (lReasonID = CANCEL_EMPLOYEES_C04) Or (lReasonID = CANCEL_EMPLOYEES_SSI)) Then
				If Len(oRequest("Cancel").Item) > 0 Then
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Cancel"" ID=""CancelHdn"" VALUE=""" & CInt(oRequest("Cancel").Item) & """ />"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de fin:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
						If CLng(aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE)) = 30000000 Then
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & "A la fecha" & "</FONT></TD>"
						Else
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE), -1, -1, -1) & "</FONT></TD>"
						End If
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptEndDate"" ID=""ConceptStartDateHdn"" VALUE=""" & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & """ />"
					Response.Write "</TR>"
				Else
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de fin:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE), "ConceptEnd", Year(Date())-5, Year(Date())+2, True, True) & "</FONT></TD>"
					Response.Write "</TR>"
				End If
			End If
			If ((lReasonID = CANCEL_EMPLOYEES_CONCEPTS) Or (lReasonID = CANCEL_EMPLOYEES_C04)) Then
				Response.Write "<TR NAME=""ConceptIDDiv"" ID=""ConceptIDDiv"">"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Concepto:&nbsp;</FONT></TD>"
					Response.Write "<TD>"
						Response.Write "<SELECT NAME=""ConceptID"" ID=""ConceptIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID = " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ")", "ConceptShortName, ConceptName", "", "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT>"
					Response.Write "</TD>"
				Response.Write "</TR>"
			ElseIf lReasonID = -95 Then
				Response.Write "<TR NAME=""ConceptIDDiv"" ID=""ConceptIDDiv"">"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Concepto:&nbsp;</FONT></TD>"
					Response.Write "<TD>"
						Response.Write "<SELECT NAME=""ConceptID"" ID=""ConceptIDCmb"" SIZE=""1"" CLASS=""Lists"" onClick=""SearchRecord(this.value + '&lEmployeeID=' + EmployeeID.value, 'StartDateForConcept', 'SearchConceptStartDateIFrame', 'EmployeeFrm.ConceptStartDate')"">"
							lDisplayFormCurrentDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EmployeesConceptsLKP, Concepts", "EmployeesConceptsLKP.ConceptID", "EmployeesConceptsLKP.ConceptID,ConceptName", "(EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And ((EmployeesConceptsLKP.EndDate = 30000000) Or (EmployeesConceptsLKP.EndDate > " & lDisplayFormCurrentDate & ") And (EmployeesConceptsLKP.StartDate < " & lDisplayFormCurrentDate & "))	And (EmployeesConceptsLKP.ConceptID In (87,120)) And (EmployeesConceptsLKP.ConceptID = Concepts.ConceptID)", "EmployeesConceptsLKP.ConceptID,EmployeesConceptsLKP.EndDate", "", "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT>"
					Response.Write "</TD>"
				Response.Write "</TR>"
			Else
				Response.Write "<TR NAME=""ConceptIDDiv"" ID=""ConceptIDDiv"">"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Concepto:&nbsp;</FONT></TD>"
					Response.Write "<TD>"
						Response.Write "<SELECT NAME=""ConceptID"" ID=""ConceptIDCmb"" SIZE=""1"" CLASS=""Lists"" onClick=""SearchRecord(this.value + '&lEmployeeID=' + EmployeeID.value, 'StartDateForConcept', 'SearchConceptStartDateIFrame', 'EmployeeFrm.ConceptStartDate')"">"
							lDisplayFormCurrentDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EmployeesConceptsLKP, Concepts", "EmployeesConceptsLKP.ConceptID", "EmployeesConceptsLKP.ConceptID,ConceptName", "(EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And ((EmployeesConceptsLKP.EndDate = 30000000) Or (EmployeesConceptsLKP.EndDate > " & lDisplayFormCurrentDate & ") And (EmployeesConceptsLKP.StartDate < " & lDisplayFormCurrentDate & "))	And (EmployeesConceptsLKP.ConceptID In (4,7,8)) And (EmployeesConceptsLKP.ConceptID = Concepts.ConceptID)", "EmployeesConceptsLKP.ConceptID,EmployeesConceptsLKP.EndDate", "", "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT>"
					Response.Write "</TD>"
				Response.Write "</TR>"
			End If
			Response.Write "<TR NAME=""PayrollDateDiv"" ID=""PayrollDateDiv"">"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Quincena de aplicación:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""EmployeePayrollDate"" ID=""EmployeePayrollDateCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""CheckStatusPayrolls()"">"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And (IsActive_1=1) And (PayrollTypeID=1)", "PayrollID Desc", aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
				Response.Write "</SELECT>&nbsp;"
				Response.Write "</TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	End If
End Function
%>