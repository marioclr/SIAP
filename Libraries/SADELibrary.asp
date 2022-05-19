<%
Dim iSIAPSADEConnectionType
Dim oSIAPSADEADODBConnection
iSIAPSADEConnectionType = iConnectionType
Call CreateADODBConnection(SADE_DATABASE_PATH, SADE_DATABASE_USERNAME, SADE_DATABASE_PASSWORD, iSIAPSADEConnectionType, oSIAPSADEADODBConnection, "")

Function AssignEmployeesToCourse(oRequest, oADODBConnection, lCourseID, sEmployeeIDs, sSchoolarshipLevels, sEmployeesDuties, sEmployeeLocations, sErrorDescription)
'************************************************************
'Purpose: To assign the given employees to the course
'Inputs:  oRequest, oADODBConnection, lCourseID, sEmployeeIDs, sSchoolarshipLevels, sEmployeesDuties, sEmployeeLocations
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AssignEmployeesToCourse"
	Dim asEmployeeIDs
	Dim asSchoolarshipLevels
	Dim asEmployeesDuties
	Dim asEmployeeLocations
	Dim iIndex
	Dim lErrorNumber

	sErrorDescription = "No se pudieron registrar los empleados que tomarán el curso."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & SADE_PREFIX & "CursosEmpleadosLKP Where (ID_Curso=" & lCourseID & ")", "SADELibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	If lErrorNumber = 0 Then
		asEmployeeIDs = Split(sEmployeeIDs, ",")
		asSchoolarshipLevels = Split(sSchoolarshipLevels, LIST_SEPARATOR)
		asEmployeesDuties = Split(sEmployeesDuties, LIST_SEPARATOR)
		asEmployeeLocations = Split(sEmployeeLocations, LIST_SEPARATOR)
		For iIndex = 0 To UBound(asEmployeeIDs)
			sErrorDescription = "No se pudieron registrar los empleados que tomarán el curso."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & SADE_PREFIX & "CursosEmpleadosLKP (ID_Curso, ID_Empleado, SchoolarshipLevel, EmployeeDuties, EmployeeLocation, Calificacion, Aprobado) Values (" & lCourseID & ", " & asEmployeeIDs(iIndex) & ", '" & asSchoolarshipLevels(iIndex) & "', '" & asEmployeesDuties(iIndex) & "', '" & asEmployeeLocations(iIndex) & "', -1, 0)", "SADELibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		Next
	End If

	AssignEmployeesToCourse = lErrorNumber
	Err.number
End Function

Function AssignProfilesToCourse(oRequest, oADODBConnection, lCourseID, sProfileIDs, sErrorDescription)
'************************************************************
'Purpose: To assign the given profiles to the course
'Inputs:  oRequest, oADODBConnection, lCourseID, sProfileIDs
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AssignProfilesToCourse"
	Dim asProfileIDs
	Dim iIndex
	Dim lErrorNumber

	sErrorDescription = "No se pudieron registrar los empleados que tomarán el curso."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & SADE_PREFIX & "CursosPerfilesLKP Where (ID_Curso=" & lCourseID & ")", "SADELibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	If lErrorNumber = 0 Then
		asProfileIDs = Split(sProfileIDs, ",")
		For iIndex = 0 To UBound(asProfileIDs)
			sErrorDescription = "No se pudieron registrar los empleados que tomarán el curso."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & SADE_PREFIX & "CursosPerfilesLKP (ID_Curso, ID_Perfil) Values (" & lCourseID & ", " & asProfileIDs(iIndex) & ")", "SADELibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		Next
	End If

	AssignProfilesToCourse = lErrorNumber
	Err.number
End Function

Function DisplayCertificateForm(oRequest, oADODBConnection, lCourseID, lEmployeeID, sErrorDescription)
'************************************************************
'Purpose: To display the HTML Form to save the employee's grade
'Inputs:  oRequest, oADODBConnection, lCourseID, lEmployeeID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayCertificateForm"
	Dim oRecordset
	Dim lRecordID
	Dim lDate
	Dim lErrorNumber

	lRecordID = -1
	lDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	sErrorDescription = "No se pudo obtener la información del certificado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From " & SADE_PREFIX & "Constancias Where (ID_Usuario=" & lEmployeeID & ") And (ID_Curso=" & lCourseID & ")", "SADELibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			lRecordID = CLng(oRecordset.Fields("ID_Constancia").Value)
			lDate = CLng(oRecordset.Fields("Fecha_Impresion").Value)
		End If
		oRecordset.Close
	End If
	Response.Write "<FORM NAME=""GradeFrm"" ID=""GradeFrm"" ACTION=""SADE.asp"" METHOD=""GET"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""" & iSectionID & """ />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""2"" />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RecordID"" ID=""RecordIDHdn"" VALUE=""" & lRecordID & """ />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & lEmployeeID & """ />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CourseID"" ID=""CourseIDHdn"" VALUE=""" & lCourseID & """ />"

		Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Indique la fecha de emisión del certificado.<BR /><BR /></FONT>"
		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fecha:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(lDate, "Print", N_START_YEAR, Year(Date()), True, False) & "</FONT></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE><BR />"

		Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Imprimir"" CLASS=""Buttons"" />"
		Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
		Response.Write "<INPUT TYPE=""BUTTON"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?SectionID=" & oRequest("SectionID").Item & "&Step=2&CourseID=" & lCourseID & "'"" />"
	Response.Write "</FORM>"

	Set oRecordset = Nothing
	DisplayCertificateForm = lErrorNumber
	Err.number
End Function

Function DisplayCourseEmployeesTable(oRequest, oADODBConnection, lCourseID, bForCertificate, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the employees suscribed to the given course
'Inputs:  oRequest, oADODBConnection, lCourseID, bForCertificate, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayCourseEmployeesTable"
	Dim oRecordset
	Dim oGradesRecordset
	Dim sTemp
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sBoldBegin
	Dim sBoldEnd
	Dim sFontBegin
	Dim sFontEnd
	Dim lErrorNumber

	sCondition = ""
	If Not bAll Then sCondition = "And (Fecha_Final>=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ")"
	sErrorDescription = "No se pudo obtener la información de los cursos."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select " & SADE_PREFIX & "CursosEmpleadosLKP.*, ID_Perfil, Nombre_Perfil, EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, PositionShortName, PositionName, AreaShortName, AreaName From " & SADE_PREFIX & "CursosEmpleadosLKP, " & SADE_PREFIX & "Curso, " & SADE_PREFIX & "Perfiles, Employees, Jobs, Positions, Areas Where (" & SADE_PREFIX & "CursosEmpleadosLKP.ID_Curso=" & SADE_PREFIX & "Curso.ID_Curso) And (" & SADE_PREFIX & "Curso.MostrarEvaluaciones=" & SADE_PREFIX & "Perfiles.ID_Perfil) And (" & SADE_PREFIX & "CursosEmpleadosLKP.ID_Empleado=Employees.EmployeeID) And (Employees.JobID=Jobs.JobID) And (Jobs.PositionID=Positions.PositionID) And (Jobs.AreaID=Areas.AreaID) And (" & SADE_PREFIX & "CursosEmpleadosLKP.ID_Curso=" & lCourseID & ") Order By EmployeeLastName, EmployeeLastName2, EmployeeName", "SADELibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE WIDTH=""900"" BORDER="""
				If bForExport Then
					Response.Write "1"
				Else
					Response.Write "0"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = "No. del empleado,Nombre,Area,Puesto,Estatus,Calificación"
				asCellWidths = "100,150,175,175,100,100"
				asCellAlignments = ",,,,,RIGHT"
				If CLng(oRecordset.Fields("ID_Perfil").Value) > -1 Then
					asColumnsTitles = asColumnsTitles & ",Calif. global"
					asCellWidths = asCellWidths & ",100"
					asCellAlignments = asCellAlignments & ",RIGHT"
				End If
				If Not bForExport Then
					asColumnsTitles = asColumnsTitles & ",Acciones"
					asCellWidths = asCellWidths & ",100"
					asCellAlignments = asCellAlignments & ",CENTER"
				End If
				asColumnsTitles = Split(asColumnsTitles, ",", -1, vbBinaryCompare)
				asCellWidths = Split(asCellWidths, ",", -1, vbBinaryCompare)
				asCellAlignments = Split(asCellAlignments, ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				Do While Not oRecordset.EOF
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("ID_Empleado").Value), oRequest("EmployeeID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If

					sFontBegin = ""
					sFontEnd = ""
					If (CDbl(oRecordset.Fields("Calificacion").Value) > -1) And (CInt(oRecordset.Fields("Aprobado").Value) <> 1) Then
						sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
						sFontEnd = "</FONT>"
					End If

					sTemp = " "
					If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then sTemp = CStr(oRecordset.Fields("EmployeeLastName2").Value)
					Err.Clear
					sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & sTemp & ", " & CStr(oRecordset.Fields("EmployeeName").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("AreaShortName").Value) & ". " & CStr(oRecordset.Fields("AreaName").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
					 Select Case CInt(oRecordset.Fields("Aprobado").Value)
						Case 0
							sRowContents = sRowContents & "Reprobado"
						Case 1
							sRowContents = sRowContents & "Aprobado"
						Case 2
							sRowContents = sRowContents & "Desertó"
						Case 3
							sRowContents = sRowContents & "No deseó participar"
					End Select
					sRowContents = sRowContents & sBoldEnd & sFontEnd
					If CDbl(oRecordset.Fields("Calificacion").Value) > -1 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(oRecordset.Fields("Calificacion").Value), 2, True, False, True) & sBoldEnd & sFontEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & "No se ha registrado" & sBoldEnd & sFontEnd
					End If
					If CLng(oRecordset.Fields("ID_Perfil").Value) > -1 Then
						sErrorDescription = "No se pudo obtener la información de los cursos."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(Calificacion) As SumOfGrades, Count(Calificacion) As TotalCourses From " & SADE_PREFIX & "CursosEmpleadosLKP, " & SADE_PREFIX & "Curso Where (" & SADE_PREFIX & "CursosEmpleadosLKP.ID_Curso=" & SADE_PREFIX & "Curso.ID_Curso) And (MostrarEvaluaciones=" & CStr(oRecordset.Fields("ID_Perfil").Value) & ") And (ID_Empleado=" & CStr(oRecordset.Fields("ID_Empleado").Value) & ") And (Calificacion>-1)", "SADELibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oGradesRecordset)
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
							If lErrorNumber = 0 Then
								If Not oGradesRecordset.EOF Then
									sRowContents = sRowContents & FormatNumber((CDbl(oGradesRecordset.Fields("SumOfGrades").Value) / CDbl(oGradesRecordset.Fields("TotalCourses").Value)), 2, True, False, True)
								Else
									sRowContents = sRowContents & "<CENTER>---</CENTER>"
								End If
							Else
								sRowContents = sRowContents & "<CENTER>---</CENTER>"
							End If
						sRowContents = sRowContents & sBoldEnd & sFontEnd
					End If
					If Not bForExport Then
						sRowContents = sRowContents & TABLE_SEPARATOR
						If bForCertificate Then
							If (CDbl(oRecordset.Fields("Calificacion").Value) > -1) And (CDbl(oRecordset.Fields("Aprobado").Value) = 1) Then
								sRowContents = sRowContents & "<A HREF=""SADE.asp?SectionID=" & oRequest("SectionID").Item & "&Step=2&CourseID=" & CStr(oRecordset.Fields("ID_Curso").Value) & "&EmployeeID=" & CStr(oRecordset.Fields("ID_Empleado").Value) & """>"
									sRowContents = sRowContents & "<IMG SRC=""Images/IcnPrint.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Imprimir certificado"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
							End If
						Else
							If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
								sRowContents = sRowContents & "<A HREF=""SADE.asp?SectionID=" & oRequest("SectionID").Item & "&Step=2&CourseID=" & CStr(oRecordset.Fields("ID_Curso").Value) & "&EmployeeID=" & CStr(oRecordset.Fields("ID_Empleado").Value) & """>"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Registrar calificación"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
							End If
						End If
					End If

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos."
		End If
	End If

	Set oRecordset = Nothing
	DisplayCourseEmployeesTable = lErrorNumber
	Err.Clear
End Function

Function DisplayCourseProfilesForm(oRequest, oADODBConnection, lCourseID, sErrorDescription)
'************************************************************
'Purpose: To display the HTML Form for the courses registration
'Inputs:  oRequest, oADODBConnection, lCourseID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayCourseProfilesForm"
	Dim lErrorNumber
	Dim oRecordset

	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		Response.Write "var bAdd" & vbNewLine

		Response.Write "function CheckCourseRegistrationFields(oForm) {" & vbNewLine
			Response.Write "if (oForm) {" & vbNewLine
				Response.Write "SelectAllItemsFromList(oForm.ProfileIDs);" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "}" & vbNewLine
		Response.Write "} // End of CheckCourseRegistrationFields" & vbNewLine
	Response.Write "//--></SCRIPT>" & vbNewLine

	Response.Write "<FORM NAME=""CourseFrm"" ID=""CourseFrm"" ACTION=""SADE.asp"" METHOD=""GET"" onSubmit=""return CheckCourseRegistrationFields(this)"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""" & iSectionID & """ />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""2"" />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CourseID"" ID=""CourseIDHdn"" VALUE=""" & lCourseID & """ />"

		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
			Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Perfiles de puesto:<BR /></FONT>"
				Response.Write "<SELECT NAME=""ProfileID"" ID=""ProfileIDLst"" SIZE=""20"" MULTIPLE=""1"" CLASS=""Lists"">"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, SADE_PREFIX & "Perfiles", "ID_Perfil", "Nombre_Perfil", "(ID_Perfil Not In (Select ID_Perfil From " & SADE_PREFIX & "CursosPerfilesLKP Where ((ID_Curso=" & lCourseID & "))))", "Nombre_Perfil", "", "", sErrorDescription)
				Response.Write "</SELECT>"
			Response.Write "</TD>"
			Response.Write "<TD>"
				Response.Write "<A HREF=""javascript: DoNothing()"" onClick=""MoveItemsBetweenLists(['',''], document.CourseFrm.ProfileID, document.CourseFrm.ProfileIDs);""><IMG SRC=""Images/BtnCrclAdd.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Asignar el perfil al curso"" BORDER=""0"" /></A>"
				Response.Write "<BR /><BR />"
				Response.Write "<A HREF=""javascript: DoNothing()"" onClick=""MoveItemsBetweenLists(['',''], document.CourseFrm.ProfileIDs, document.CourseFrm.ProfileID);""><IMG SRC=""Images/BtnCrclAddLeft.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""No asignar el perfil al curso"" BORDER=""0"" /></A>"
			Response.Write "</TD>"
			Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Perfiles que tomarán el curso:<BR /></FONT>"
				Response.Write "<SELECT NAME=""ProfileIDs"" ID=""ProfileIDsLst"" SIZE=""20"" MULTIPLE=""1"" CLASS=""Lists"">"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, SADE_PREFIX & "Perfiles, " & SADE_PREFIX & "CursosPerfilesLKP", SADE_PREFIX & "Perfiles.ID_Perfil", "Nombre_Perfil", "(" & SADE_PREFIX & "Perfiles.ID_Perfil=" & SADE_PREFIX & "CursosPerfilesLKP.ID_Perfil) And (" & SADE_PREFIX & "CursosPerfilesLKP.ID_Curso=" & lCourseID & ")", "Nombre_Perfil", "", "", sErrorDescription)
				Response.Write "</SELECT>"
			Response.Write "</TD>"
		Response.Write "</TR></TABLE><BR />"

		Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Registrar"" CLASS=""Buttons"" />"
		Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
		Response.Write "<INPUT TYPE=""BUTTON"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?SectionID=" & oRequest("SectionID").Item & "'"" / id=1 name=1>"
	Response.Write "</FORM>"

	Set oRecordset = Nothing
	DisplayCourseProfilesForm = lErrorNumber
	Err.number
End Function

Function DisplayCourseRegistrationForm(oRequest, oADODBConnection, lCourseID, sErrorDescription)
'************************************************************
'Purpose: To display the HTML Form for the courses registration
'Inputs:  oRequest, oADODBConnection, lCourseID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayCourseRegistrationForm"
	Dim lErrorNumber
	Dim oRecordset

	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		Response.Write "var bAdd" & vbNewLine

		Response.Write "function AddEmployeeToCourse() {" & vbNewLine
			Response.Write "var oForm = document.CourseFrm;" & vbNewLine
			
			Response.Write "if (oForm) {" & vbNewLine
				Response.Write "if (oForm.EmployeeName.value == '.') {" & vbNewLine
					Response.Write "window.setTimeout('AddEmployeeToCourse()', 100);"
				Response.Write "} else {" & vbNewLine
					Response.Write "if (oForm.EmployeeName.value != '') {" & vbNewLine
						Response.Write "UnselectAllItemsFromList(oForm.EmployeeIDs);" & vbNewLine
						Response.Write "UnselectAllItemsFromList(oForm.SchoolarshipLevels);" & vbNewLine
						Response.Write "UnselectAllItemsFromList(oForm.EmployeesDuties);" & vbNewLine
						Response.Write "UnselectAllItemsFromList(oForm.EmployeeLocations);" & vbNewLine
						Response.Write "SelectListItemByValue(oForm.EmployeeID.value, false, oForm.EmployeeIDs);" & vbNewLine
						Response.Write "SelectSameItems(oForm.EmployeeIDs, oForm.SchoolarshipLevels);" & vbNewLine
						Response.Write "SelectSameItems(oForm.EmployeeIDs, oForm.EmployeesDuties);" & vbNewLine
						Response.Write "SelectSameItems(oForm.EmployeeIDs, oForm.EmployeeLocations);" & vbNewLine
						Response.Write "RemoveSelectedItemsFromList(null, document.CourseFrm.EmployeeIDs);" & vbNewLine
						Response.Write "RemoveSelectedItemsFromList(null, document.CourseFrm.SchoolarshipLevels);" & vbNewLine
						Response.Write "RemoveSelectedItemsFromList(null, document.CourseFrm.EmployeesDuties);" & vbNewLine
						Response.Write "RemoveSelectedItemsFromList(null, document.CourseFrm.EmployeeLocations);" & vbNewLine

						Response.Write "AddItemToList(oForm.EmployeeName.value, oForm.EmployeeID.value, null, oForm.EmployeeIDs);" & vbNewLine
						Response.Write "AddItemToList(oForm.SchoolarshipLevelTemp.value, oForm.SchoolarshipLevelTemp.value, null, oForm.SchoolarshipLevels);" & vbNewLine
						Response.Write "AddItemToList(oForm.EmployeeDutiesTemp.value, oForm.EmployeeDutiesTemp.value, null, oForm.EmployeesDuties);" & vbNewLine
						Response.Write "AddItemToList(oForm.EmployeeLocationTemp.value, oForm.EmployeeLocationTemp.value, null, oForm.EmployeeLocations);" & vbNewLine
						Response.Write "oForm.EmployeeName.value = '';" & vbNewLine
						Response.Write "oForm.EmployeeID.value = '';" & vbNewLine
						Response.Write "oForm.SchoolarshipLevelTemp.value = '';" & vbNewLine
						Response.Write "oForm.EmployeeDutiesTemp.value = '';" & vbNewLine
						Response.Write "oForm.EmployeeLocationTemp.value = '';" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "}" & vbNewLine
			Response.Write "oForm.EmployeeID.focus();" & vbNewLine
		Response.Write "} // End of AddEmployeeToCourse" & vbNewLine
		
		Response.Write "function CheckCourseRegistrationFields(oForm) {" & vbNewLine
			Response.Write "if (oForm) {" & vbNewLine
				Response.Write "SelectAllItemsFromList(oForm.EmployeeIDs);" & vbNewLine
				Response.Write "SelectAllItemsFromList(oForm.SchoolarshipLevels);" & vbNewLine
				Response.Write "SelectAllItemsFromList(oForm.EmployeesDuties);" & vbNewLine
				Response.Write "SelectAllItemsFromList(oForm.EmployeeLocations);" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "}" & vbNewLine
		Response.Write "} // End of CheckCourseRegistrationFields" & vbNewLine

		Response.Write "function SelectSameItemsForCourse(oList) {" & vbNewLine
			Response.Write "var oForm = document.CourseFrm;" & vbNewLine
			Response.Write "if (oForm) {" & vbNewLine
				Response.Write "SelectSameItems(oList, oForm.EmployeeIDs);" & vbNewLine
				Response.Write "SelectSameItems(oList, oForm.SchoolarshipLevels);" & vbNewLine
				Response.Write "SelectSameItems(oList, oForm.EmployeesDuties);" & vbNewLine
				Response.Write "SelectSameItems(oList, oForm.EmployeeLocations);" & vbNewLine
			Response.Write "}" & vbNewLine
		Response.Write "} // End of SelectSameItemsForCourse" & vbNewLine

		Response.Write "function SendEmployeeIDToIFrame() {" & vbNewLine
			Response.Write "var oForm = document.CourseFrm;" & vbNewLine
			Response.Write "var bCorrect = true;" & vbNewLine
			
			Response.Write "if (oForm.EmployeeID.value == '') {" & vbNewLine
				Response.Write "alert('Favor de indicar el número del empleado.');" & vbNewLine
				Response.Write "oForm.EmployeeID.focus();" & vbNewLine
				Response.Write "bCorrect = false;" & vbNewLine
			Response.Write "}" & vbNewLine
			Response.Write "if (oForm.SchoolarshipLevelTemp.value == '') {" & vbNewLine
				Response.Write "alert('Favor de indicar el nivel de estudios del empleado.');" & vbNewLine
				Response.Write "oForm.SchoolarshipLevelTemp.focus();" & vbNewLine
				Response.Write "bCorrect = false;" & vbNewLine
			Response.Write "}" & vbNewLine
			Response.Write "if (oForm.EmployeeDutiesTemp.value == '') {" & vbNewLine
				Response.Write "alert('Favor de indicar las actividades que desarrolla el empleado.');" & vbNewLine
				Response.Write "oForm.EmployeeDutiesTemp.focus();" & vbNewLine
				Response.Write "bCorrect = false;" & vbNewLine
			Response.Write "}" & vbNewLine
			Response.Write "if (oForm.EmployeeLocationTemp.value == '') {" & vbNewLine
				Response.Write "alert('Favor de indicar el lugar donde el empleado presta sus servicios.');" & vbNewLine
				Response.Write "oForm.EmployeeLocationTemp.focus();" & vbNewLine
				Response.Write "bCorrect = false;" & vbNewLine
			Response.Write "}" & vbNewLine
			Response.Write "if (bCorrect) {" & vbNewLine
				Response.Write "oForm.EmployeeName.value='.';" & vbNewLine
				Response.Write "SearchRecord(oForm.EmployeeID.value, 'EmployeesNameFromNumber', 'SearchRecordIFrame', 'CourseFrm.EmployeeName');" & vbNewLine
				Response.Write "AddEmployeeToCourse();" & vbNewLine
			Response.Write "}" & vbNewLine
		Response.Write "} // End of SendEmployeeIDToIFrame" & vbNewLine
	Response.Write "//--></SCRIPT>" & vbNewLine

	Response.Write "<FORM NAME=""CourseFrm"" ID=""CourseFrm"" ACTION=""SADE.asp"" METHOD=""POST"" onSubmit=""return CheckCourseRegistrationFields(this)"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""" & iSectionID & """ />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""2"" />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CourseID"" ID=""CourseIDHdn"" VALUE=""" & lCourseID & """ />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeName"" ID=""EmployeeNameHdn"" VALUE="""" />"

		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
			Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><B>Información del empleado:</B><BR />"
				Response.Write "Número del empleado: <INPUT TYPE=""TEXT"" NAME=""EmployeeID"" ID=""EmployeeIDTxt"" SIZE=""6"" MAXLENGTH=""6"" CLASS=""TextFields"" />&nbsp;"
				Response.Write "<A HREF=""javascript: SearchRecord(document.CourseFrm.EmployeeID.value, 'EmployeesInfo', 'EmployeeInfoIFrame', '');""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Obtener información"" BORDER=""0"" ALIGN=""ABSMIDDLE"" /></A><BR />"
				Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchRecordIFrame"" FRAMEBORDER=""0"" WIDTH=""250"" HEIGHT=""20""></IFRAME><BR />"
				Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
					Response.Write "<TR>"
						Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Nivel de estudios:&nbsp;</FONT></TD>"
						Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""SchoolarshipLevelTemp"" ID=""SchoolarshipLevelTempTxt"" SIZE=""20"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Actividades que dearrolla:&nbsp;</FONT></TD>"
						Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeDutiesTemp"" ID=""EmployeeDutiesTempTxt"" SIZE=""20"" MAXLENGTH=""255"" CLASS=""TextFields"" /></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Lugar donde presta sus servicios:&nbsp;</FONT></TD>"
						Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLocationTemp"" ID=""EmployeeLocationTempTxt"" SIZE=""20"" MAXLENGTH=""255"" CLASS=""TextFields"" /></TD>"
					Response.Write "</TR>"
				Response.Write "</TABLE><BR />"
				Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""EmployeeInfoIFrame"" FRAMEBORDER=""0"" WIDTH=""250"" HEIGHT=""300""></IFRAME>"
			Response.Write "</FONT></TD>"
			Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><BR /></FONT><A HREF=""javascript: SendEmployeeIDToIFrame();""><IMG SRC=""Images/BtnCrclAdd.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Inscribir al empleado en el curso"" BORDER=""0"" /></A>&nbsp;&nbsp;&nbsp;</TD>"
			Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><B>Empleados:</B><BR /></FONT><SELECT NAME=""EmployeeIDs"" ID=""EmployeeIDsLst"" SIZE=""20"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""SelectSameItemsForCourse(this);"">"
				Response.Write GenerateListOptionsFromQuery(oADODBConnection, SADE_PREFIX & "CursosEmpleadosLKP, Employees", "EmployeeID", "EmployeeName, EmployeeLastName, EmployeeLastName2", "(" & SADE_PREFIX & "CursosEmpleadosLKP.ID_Empleado=Employees.EmployeeID) And (" & SADE_PREFIX & "CursosEmpleadosLKP.ID_Curso=" & lCourseID & ")", "EmployeeName, EmployeeLastName, EmployeeLastName2", "", "", sErrorDescription)
			Response.Write "</SELECT>&nbsp;&nbsp;</TD>"
			Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><B>Niveles:</B><BR /></FONT><SELECT NAME=""SchoolarshipLevels"" ID=""SchoolarshipLevelsLst"" SIZE=""20"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""SelectSameItemsForCourse(this);"">"
				Response.Write GenerateListOptionsFromQuery(oADODBConnection, SADE_PREFIX & "CursosEmpleadosLKP, Employees", "SchoolarshipLevel", "SchoolarshipLevel As Temp1", "(" & SADE_PREFIX & "CursosEmpleadosLKP.ID_Empleado=Employees.EmployeeID) And (" & SADE_PREFIX & "CursosEmpleadosLKP.ID_Curso=" & lCourseID & ")", "EmployeeName, EmployeeLastName, EmployeeLastName2", "", "", sErrorDescription)
			Response.Write "</SELECT>&nbsp;&nbsp;</TD>"
			Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><B>Actividades:</B><BR /></FONT><SELECT NAME=""EmployeesDuties"" ID=""EmployeesDutiesLst"" SIZE=""20"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""SelectSameItemsForCourse(this);"">"
				Response.Write GenerateListOptionsFromQuery(oADODBConnection, SADE_PREFIX & "CursosEmpleadosLKP, Employees", "EmployeeDuties", "EmployeeDuties As Temp1", "(" & SADE_PREFIX & "CursosEmpleadosLKP.ID_Empleado=Employees.EmployeeID) And (" & SADE_PREFIX & "CursosEmpleadosLKP.ID_Curso=" & lCourseID & ")", "EmployeeName, EmployeeLastName, EmployeeLastName2", "", "", sErrorDescription)
			Response.Write "</SELECT>&nbsp;&nbsp;</TD>"
			Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><B>Lugares:</B><BR /></FONT><SELECT NAME=""EmployeeLocations"" ID=""EmployeeLocationsLst"" SIZE=""20"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""SelectSameItemsForCourse(this);"">"
				Response.Write GenerateListOptionsFromQuery(oADODBConnection, SADE_PREFIX & "CursosEmpleadosLKP, Employees", "EmployeeLocation", "EmployeeLocation As Temp1", "(" & SADE_PREFIX & "CursosEmpleadosLKP.ID_Empleado=Employees.EmployeeID) And (" & SADE_PREFIX & "CursosEmpleadosLKP.ID_Curso=" & lCourseID & ")", "EmployeeName, EmployeeLastName, EmployeeLastName2", "", "", sErrorDescription)
			Response.Write "</SELECT>&nbsp;&nbsp;</TD>"
			Response.Write "<TD VALIGN=""BOTTOM""><A HREF=""javascript: RemoveSelectedItemsFromList(null, document.CourseFrm.EmployeeIDs); RemoveSelectedItemsFromList(null, document.CourseFrm.SchoolarshipLevels); RemoveSelectedItemsFromList(null, document.CourseFrm.EmployeesDuties); RemoveSelectedItemsFromList(null, document.CourseFrm.EmployeeLocations);""><IMG SRC=""Images/BtnCrclDelete.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Cancelar la participación del empleado en el curso"" BORDER=""0"" /></A></TD>"
		Response.Write "</TR></TABLE><BR />"

		Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Registrar"" CLASS=""Buttons"" />"
		Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
		Response.Write "<INPUT TYPE=""BUTTON"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?SectionID=" & oRequest("SectionID").Item & "'"" />"
	Response.Write "</FORM>"

	Set oRecordset = Nothing
	DisplayCourseRegistrationForm = lErrorNumber
	Err.number
End Function

Function DisplayEmployeeCurriculum(oRequest, oADODBConnection, lEmployeeID, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display history of courses for the given employee
'Inputs:  oRequest, oADODBConnection, lEmployeeID, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeCurriculum"
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información del empleado en el curso."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeName, EmployeeLastName, EmployeeLastName2, EmployeeNumber, StartDate, Nombre_Curso, Descripcion, TiempoEstimado, Fecha_Inicio, Fecha_Final, Calificacion, Aprobado From Employees, " & SADE_PREFIX & "CursosEmpleadosLKP, " & SADE_PREFIX & "Curso, " & SADE_PREFIX & "CursosGruposLKP Where (Employees.EmployeeID=" & SADE_PREFIX & "CursosEmpleadosLKP.ID_Empleado) And (" & SADE_PREFIX & "CursosEmpleadosLKP.ID_Curso=" & SADE_PREFIX & "Curso.ID_Curso) And (" & SADE_PREFIX & "Curso.ID_Curso=" & SADE_PREFIX & "CursosGruposLKP.ID_Curso) And (Employees.EmployeeID=" & Right(("000000" & lEmployeeID), Len("000000")) & ") And (Calificacion>-1)", "SADELibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Response.Write "<FONT FACE=""Arial"" SIZE=""2"">"
			If Not oRecordset.EOF Then
				If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
					Response.Write "<B>CURRICULUM DE " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value)) & "</B><BR />"
				Else
					Response.Write "<B>CURRICULUM DE " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value)) & "</B><BR />"
				End If
				Response.Write "<B>No. de empleado:&nbsp;</B>" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & "<BR />"
				Response.Write "<B>Fecha de ingreso:&nbsp;</B>" & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1) & "<BR />"
				Response.Write "<BR /><BR />"
				Response.Write "<B>CURSOS QUE HA TOMADO</B><BR />"
				Do While Not oRecordset.EOF
					Response.Write "&nbsp;&nbsp;&nbsp;<B>" & CleanStringForHTML(CStr(oRecordset.Fields("Nombre_Curso").Value)) & "</B><BR />"
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Impartido del " & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("Fecha_Inicio").Value), -1, -1, -1) & " al " & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("Fecha_Final").Value), -1, -1, -1) & " por " & CleanStringForHTML(CStr(oRecordset.Fields("Descripcion").Value)) & ".<BR />"
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Duración: " & FormatNumber(CLng(oRecordset.Fields("TiempoEstimado").Value), 0, True, False, True) & " horas.<BR />"
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Calificación: " & FormatNumber(CDbl(oRecordset.Fields("Calificacion").Value), 2, True, False, True) & ".<BR />"
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Aprobado: " & DisplayYesNo(CInt(oRecordset.Fields("Aprobado").Value), True) & ".<BR />"
					Response.Write "<BR /><BR />"
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El empleado no ha tomado ningún curso o aún no se han registrado los resultados que obtuvo en ellos."
			End If
		Response.Write "</FONT>"
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	DisplayEmployeeCurriculum = lErrorNumber
	Err.number
End Function

Function DisplayGradesForm(oRequest, oADODBConnection, lCourseID, lEmployeeID, sErrorDescription)
'************************************************************
'Purpose: To display the HTML Form to save the employee's grade
'Inputs:  oRequest, oADODBConnection, lCourseID, lEmployeeID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayGradesForm"
	Dim oRecordset
	Dim dGrade
	Dim iApproved
	Dim sNames
	Dim lErrorNumber

	dGrade = ""
	iApproved = 0
	sErrorDescription = "No se pudo obtener la información del empleado en el curso."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From " & SADE_PREFIX & "CursosEmpleadosLKP Where (ID_Curso=" & lCourseID & ") And (ID_Empleado=" & lEmployeeID & ")", "SADELibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If CDbl(oRecordset.Fields("Calificacion").Value) > -1 Then dGrade = CDbl(oRecordset.Fields("Calificacion").Value)
			iApproved = CInt(oRecordset.Fields("Aprobado").Value)
		End If
		oRecordset.Close
	End If
	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		Response.Write "function CheckGradeForm(oForm) {" & vbNewLine
			Response.Write "if (oForm) {" & vbNewLine
				Response.Write "if (!CheckFloatValue(oForm.Grade, 'la calificación', N_MINIMUM_ONLY_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
					Response.Write "return false;" & vbNewLine

				Response.Write "return true;" & vbNewLine
			Response.Write "}" & vbNewLine
		Response.Write "} // End of CheckGradeForm" & vbNewLine
	Response.Write "//--></SCRIPT>" & vbNewLine

	Response.Write "<FORM NAME=""GradeFrm"" ID=""GradeFrm"" ACTION=""SADE.asp"" METHOD=""GET"" onSubmit=""return CheckGradeForm(this)"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""" & iSectionID & """ />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""2"" />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CourseID"" ID=""CourseIDHdn"" VALUE=""" & lCourseID & """ />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & lEmployeeID & """ />"

		Call GetNameFromTable(oADODBConnection, "Employees", lEmployeeID, "", "", sNames, "")
		Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Registre la calificación que obtuvo <B>" & sNames & "</B> en el curso.<BR /><BR /></FONT>"
		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Calificación:&nbsp;</FONT></TD>"
				Response.Write "<TD VALIGN=""TOP""><INPUT TYPE=""TEXT"" NAME=""Grade"" ID=""GradeTxt"" SIZE=""4"" MAXLENGTH=""4"" VALUE=""" & dGrade & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Estatus:&nbsp;</FONT></TD>"
				Response.Write "<TD VALIGN=""TOP""><SELECT NAME=""Approved"" ID=""ApprovedCmb"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write "<OPTION VALUE=""0"">Reprobado</OPTION>"
					Response.Write "<OPTION VALUE=""1"">Aprobado</OPTION>"
					Response.Write "<OPTION VALUE=""2"">Desertó</OPTION>"
					Response.Write "<OPTION VALUE=""3"">No deseó participar</OPTION>"
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE><BR />"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "SelectItemByValue('" & iApproved & "', false, document.GradeFrm.Approved);" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine

		Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Registrar"" CLASS=""Buttons"" />"
		Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
		Response.Write "<INPUT TYPE=""BUTTON"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?SectionID=" & oRequest("SectionID").Item & "&Step=2&CourseID=" & lCourseID & "'"" />"
	Response.Write "</FORM>"

	Set oRecordset = Nothing
	DisplayGradesForm = lErrorNumber
	Err.number
End Function

Function DoSADEActions(oRequest, iSectionID, sErrorDescription)
'************************************************************
'Purpose: To add, modify or delete an entry from the SADE DB
'Inputs:  oRequest, iSectionID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DoSADEActions"
	Dim sSchoolarshipLevels
	Dim sEmployeesDuties
	Dim sEmployeeLocations
	Dim oItem
	Dim lErrorNumber

	Select Case iSectionID
		Case 361
			If bAction Then
				If Len(oRequest("Add").Item) > 0 Then
					lErrorNumber = AddProfile(oRequest, oSIAPSADEADODBConnection, aProfileComponent, sErrorDescription)
					aProfileComponent(N_ID_PROFILE) = -1
				ElseIf Len(oRequest("Modify").Item) > 0 Then
					lErrorNumber = ModifyProfile(oRequest, oSIAPSADEADODBConnection, aProfileComponent, sErrorDescription)
				ElseIf Len(oRequest("Remove").Item) > 0 Then
					lErrorNumber = RemoveProfile(oRequest, oSIAPSADEADODBConnection, aProfileComponent, sErrorDescription)
					aProfileComponent(N_ID_PROFILE) = -1
				End If
			End If
		Case 362
			If bAction Then
				If Len(oRequest("Add").Item) > 0 Then
					lErrorNumber = AddCourse(oRequest, oSIAPSADEADODBConnection, aCourseComponent, sErrorDescription)
					aCourseComponent(N_ID_COURSE) = -1
				ElseIf Len(oRequest("Modify").Item) > 0 Then
					lErrorNumber = ModifyCourse(oRequest, oSIAPSADEADODBConnection, aCourseComponent, sErrorDescription)
				ElseIf Len(oRequest("Remove").Item) > 0 Then
					lErrorNumber = RemoveCourse(oRequest, oSIAPSADEADODBConnection, aCourseComponent, sErrorDescription)
					aCourseComponent(N_ID_COURSE) = -1
				End If
			End If
		Case 363
			If bAction Then
				If Len(oRequest("Modify").Item) > 0 Then
					For Each oItem In oRequest("SchoolarshipLevels")
						sSchoolarshipLevels = sSchoolarshipLevels & oItem & LIST_SEPARATOR
					Next
					If Len(sSchoolarshipLevels) > 0 Then sSchoolarshipLevels = Left(sSchoolarshipLevels, (Len(sSchoolarshipLevels) - Len(LIST_SEPARATOR)))
					For Each oItem In oRequest("EmployeesDuties")
						sEmployeesDuties = sEmployeesDuties & oItem & LIST_SEPARATOR
					Next
					For Each oItem In oRequest("EmployeeLocations")
						sEmployeeLocations = sEmployeeLocations & oItem & LIST_SEPARATOR
					Next
					If Len(sEmployeesDuties) > 0 Then sEmployeesDuties = Left(sEmployeesDuties, (Len(sEmployeesDuties) - Len(LIST_SEPARATOR)))
					lErrorNumber = AssignEmployeesToCourse(oRequest, oSIAPSADEADODBConnection, oRequest("CourseID").Item, Replace(oRequest("EmployeeIDs").Item, " ", ""), sSchoolarshipLevels, sEmployeesDuties, sEmployeeLocations, sErrorDescription)
				End If
			End If
		Case 365
			If bAction Then
				If Len(oRequest("Modify").Item) > 0 Then
					lErrorNumber = UpdateEmployeeGrade(oRequest, oSIAPSADEADODBConnection, oRequest("CourseID").Item, oRequest("EmployeeID").Item, sErrorDescription)
				End If
			End If
		Case 366
			If bAction Then
				If Len(oRequest("Modify").Item) > 0 Then
					lErrorNumber = UpdateEmployeeCertificate(oRequest, oSIAPSADEADODBConnection, oRequest("CourseID").Item, oRequest("EmployeeID").Item, sErrorDescription)
				End If
			End If
		Case 368
			If bAction Then
				If Len(oRequest("Modify").Item) > 0 Then
					lErrorNumber = AssignProfilesToCourse(oRequest, oSIAPSADEADODBConnection, oRequest("CourseID").Item, Replace(oRequest("ProfileIDs").Item, " ", ""), sErrorDescription)
				End If
			End If
	End Select

	DoSADEActions = lErrorNumber
	Err.number
End Function

Function PrintEmployeeCertificate(oRequest, oADODBConnection, lCourseID, lEmployeeID, sErrorDescription)
'************************************************************
'Purpose: To modify the employee's grade for the given course
'Inputs:  oRequest, oADODBConnection, lCourseID, lEmployeeID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "PrintEmployeeCertificate"
	Dim oRecordset
	Dim sFileContents
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información del certificado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select " & SADE_PREFIX & "Constancias.ID_Constancia, " & SADE_PREFIX & "Constancias.Fecha_Impresion, " & SADE_PREFIX & "Curso.Nombre_Curso, " & SADE_PREFIX & "Curso.Descripcion, " & SADE_PREFIX & "CursosGruposLKP.Fecha_Inicio, " & SADE_PREFIX & "CursosGruposLKP.Fecha_Final, EmployeeName, EmployeeLastName, EmployeeLastName2 From " & SADE_PREFIX & "Constancias, " & SADE_PREFIX & "Curso, " & SADE_PREFIX & "CursosEmpleadosLKP, " & SADE_PREFIX & "CursosGruposLKP, Employees Where (" & SADE_PREFIX & "Constancias.ID_Curso=" & SADE_PREFIX & "Curso.ID_Curso) And (" & SADE_PREFIX & "Constancias.ID_Curso=" & SADE_PREFIX & "CursosEmpleadosLKP.ID_Curso) And (" & SADE_PREFIX & "Constancias.ID_Usuario=" & SADE_PREFIX & "CursosEmpleadosLKP.ID_Empleado) And (" & SADE_PREFIX & "Constancias.ID_Curso=" & SADE_PREFIX & "CursosGruposLKP.ID_Curso) And (" & SADE_PREFIX & "Constancias.ID_Usuario=Employees.EmployeeID) And (" & SADE_PREFIX & "Constancias.ID_Usuario=" & lEmployeeID & ") And (" & SADE_PREFIX & "Constancias.ID_Curso=" & lCourseID & ")", "SADELibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sFileContents = GetFileContents(Server.MapPath(TEMPLATES_PHYSICAL_PATH & "Certificate.htm"), sErrorDescription)
			If Len(sFileContents) > 0 Then
				sFileContents = Replace(sFileContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
				If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
					sFileContents = Replace(sFileContents, "<EMPLOYEE_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value)))
				Else
					sFileContents = Replace(sFileContents, "<EMPLOYEE_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value)))
				End If
				sFileContents = Replace(sFileContents, "<COURSE_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("Nombre_Curso").Value)))
				sFileContents = Replace(sFileContents, "<PRINT_DATE />", DisplayDateFromSerialNumber(CLng(oRecordset.Fields("Fecha_Impresion").Value), -1, -1, -1))
				sFileContents = Replace(sFileContents, "<CERTIFICATE_NUMBER />", CLng(oRecordset.Fields("ID_Constancia").Value))
				sFileContents = Replace(sFileContents, "<COURSE_DESCRIPTION />", CleanStringForHTML(CStr(oRecordset.Fields("Descripcion").Value)))
				sFileContents = Replace(sFileContents, "<START_DATE />", DisplayDateFromSerialNumber(CLng(oRecordset.Fields("Fecha_Inicio").Value), -1, -1, -1))
				sFileContents = Replace(sFileContents, "<END_DATE />", DisplayDateFromSerialNumber(CLng(oRecordset.Fields("Fecha_Final").Value), -1, -1, -1))

				If FileExists(Server.MapPath(UPLOADED_PHYSICAL_PATH & "Certificate_" & aLoginComponent(N_USER_ID_LOGIN) & ".htm"), sErrorDescription) Then Call DeleteFile(Server.MapPath(UPLOADED_PHYSICAL_PATH & "Certificate_" & aLoginComponent(N_USER_ID_LOGIN) & ".htm"), sErrorDescription)
				lErrorNumber = SaveTextToFile(Server.MapPath(UPLOADED_PHYSICAL_PATH & "Certificate_" & aLoginComponent(N_USER_ID_LOGIN) & ".htm"), sFileContents, sErrorDescription)
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "OpenNewWindow('Uploaded Files\/Certificate_" & aLoginComponent(N_USER_ID_LOGIN) & ".htm" & "', '', 'ExportToWord', 640, 480, 'yes', 'yes')" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	PrintEmployeeCertificate = lErrorNumber
	Err.number
End Function

Function UpdateEmployeeCertificate(oRequest, oADODBConnection, lCourseID, lEmployeeID, sErrorDescription)
'************************************************************
'Purpose: To modify the employee's grade for the given course
'Inputs:  oRequest, oADODBConnection, lCourseID, lEmployeeID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UpdateEmployeeCertificate"
	Dim lRecordID
	Dim sDate
	Dim lErrorNumber

	If Len(oRequest("PrintYear").Item) > 0 Then
		sDate = CLng(oRequest("PrintYear").Item & Right(("0" & oRequest("PrintMonth").Item), Len("00")) & Right(("0" & oRequest("PrintDay").Item), Len("00")))
	ElseIf Len(oRequest("PrintDate").Item) > 0 Then
		sDate = CLng(oRequest("PrintDate").Item)
	Else
		sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
	End If
	If StrComp(oRequest("RecordID").Item, "-1", vbBinaryCompare) = 0 Then
		sErrorDescription = "No se pudo obtener un identificador para el nuevo curso."
		lErrorNumber = GetNewIDFromTable(oADODBConnection, SADE_PREFIX & "Constancias", "ID_Constancia", "", 1, lRecordID, sErrorDescription)
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo guardar la calificación del certificado del empleado en el curso."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & SADE_PREFIX & "Constancias (ID_Constancia, ID_Usuario, ID_Curso, Fecha_Impresion) Values (" & lRecordID & ", " & lEmployeeID & ", " & lCourseID & ", " & sDate & ")", "SADELibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	Else
		sErrorDescription = "No se pudo guardar la calificación del certificado del empleado en el curso."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update " & SADE_PREFIX & "Constancias Set Fecha_Impresion=" & sDate & " Where (ID_Constancia=" & oRequest("RecordID").Item & ")", "SADELibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	UpdateEmployeeCertificate = lErrorNumber
	Err.number
End Function

Function UpdateEmployeeGrade(oRequest, oADODBConnection, lCourseID, lEmployeeID, sErrorDescription)
'************************************************************
'Purpose: To modify the employee's grade for the given course
'Inputs:  oRequest, oADODBConnection, lCourseID, lEmployeeID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UpdateEmployeeGrade"
	Dim lErrorNumber

	sErrorDescription = "No se pudo modificar la calificación del empleado en el curso."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update " & SADE_PREFIX & "CursosEmpleadosLKP Set Calificacion=" & oRequest("Grade").Item & ", Aprobado=" & oRequest("Approved").Item & " Where (ID_Curso=" & lCourseID & ") And (ID_Empleado=" & lEmployeeID & ")", "SADELibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

	UpdateEmployeeGrade = lErrorNumber
	Err.number
End Function
%>