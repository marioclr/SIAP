<%
Function ModifyTaxInvertions(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To update the tax invertions table
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyTaxInvertions"
	Dim lErrorNumber

	ModifyTaxInvertions = lErrorNumber
	Err.Clear
End Function

Function DisplayEmploymentAllowancesTable(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the employment allowances table
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmploymentAllowancesTable"
	Dim lFirstID
	Dim lLastID
	Dim sCondition
	Dim oRecordset
	Dim asColumnsTitles
	Dim sRowContents
	Dim asRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	sCondition = " And (PeriodID=4)"
	If Len(oRequest("PeriodID").Item) > 0 Then sCondition = " And (PeriodID=" & oRequest("PeriodID").Item & ")"
	sErrorDescription = "No se pudo obtener el monto del concepto."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmploymentAllowances Where (EndDate=30000000) " & sCondition & " Order By InferiorLimit", "ReportsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "var lFirstID = 0;" & vbNewLine
				Response.Write "var lLastID = 0;" & vbNewLine

				Response.Write "function CheckTaxFields(oForm) {" & vbNewLine
					Response.Write "var oField = null;" & vbNewLine
					Response.Write "var dPrevField = 0;" & vbNewLine
					Response.Write "var dNextField = 0;" & vbNewLine

					Response.Write "if (oForm) {" & vbNewLine
						Response.Write "oField = eval('oForm.SuperiorLimit_' + lFirstID);" & vbNewLine
						Response.Write "dPrevField = eval('oForm.InferiorLimit_' + lFirstID + '.value');" & vbNewLine
						Response.Write "dPrevField = dPrevField.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "dPrevField = parseFloat(dPrevField);" & vbNewLine
						Response.Write "dNextField = eval('oForm.InferiorLimit_' + (lFirstID + 1) + '.value');" & vbNewLine
						Response.Write "dNextField = dNextField.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "dNextField = parseFloat(dNextField);" & vbNewLine
						Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if ((! isNaN(dPrevField)) && (! isNaN(dNextField))) {" & vbNewLine
							Response.Write "if (! CheckFloatValue(oField, 'el límite superior', N_BOTH_FLAG, N_OPEN_FLAG, dPrevField, dNextField))" & vbNewLine
								Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine

						Response.Write "oField = eval('oForm.AllowanceAmount_' + lFirstID);" & vbNewLine
						Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oField, 'la cuota fija', N_MINIMUM_ONLY_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine

						Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "oField = eval('oForm.PercentageForExcess_' + lFirstID);" & vbNewLine
						Response.Write "if (! CheckFloatValue(oField, 'el porcentaje excedente', N_BOTH_FLAG, N_OPEN_FLAG, 0, 100))" & vbNewLine
							Response.Write "return false;" & vbNewLine

						Response.Write "for (var i=lFirstID+1; i<lLastID; i++) {" & vbNewLine
							Response.Write "oField = eval('oForm.InferiorLimit_' + i);" & vbNewLine
							Response.Write "dPrevField = eval('oForm.SuperiorLimit_' + (i - 1) + '.value');" & vbNewLine
							Response.Write "dPrevField = dPrevField.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "dPrevField = parseFloat(dPrevField);" & vbNewLine
							Response.Write "dNextField = eval('oForm.SuperiorLimit_' + i + '.value');" & vbNewLine
							Response.Write "dNextField = dNextField.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "dNextField = parseFloat(dNextField);" & vbNewLine
							Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "if ((! isNaN(dPrevField)) && (! isNaN(dNextField))) {" & vbNewLine
								Response.Write "if (! CheckFloatValue(oField, 'el límite inferior', N_BOTH_FLAG, N_OPEN_FLAG, dPrevField, dNextField))" & vbNewLine
									Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine

							Response.Write "oField = eval('oForm.SuperiorLimit_' + i);" & vbNewLine
							Response.Write "dPrevField = eval('oForm.InferiorLimit_' + i + '.value');" & vbNewLine
							Response.Write "dPrevField = dPrevField.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "dPrevField = parseFloat(dPrevField);" & vbNewLine
							Response.Write "dNextField = eval('oForm.InferiorLimit_' + (i + 1) + '.value');" & vbNewLine
							Response.Write "dNextField = dNextField.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "dNextField = parseFloat(dNextField);" & vbNewLine
							Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "if ((! isNaN(dPrevField)) && (! isNaN(dNextField))) {" & vbNewLine
								Response.Write "if (! CheckFloatValue(oField, 'el límite superior', N_BOTH_FLAG, N_OPEN_FLAG, dPrevField, dNextField))" & vbNewLine
									Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine

							Response.Write "oField = eval('oForm.AllowanceAmount_' + i);" & vbNewLine
							Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "if (! CheckFloatValue(oField, 'la cuota fija', N_MINIMUM_ONLY_FLAG, N_OPEN_FLAG, 0, 0))" & vbNewLine
								Response.Write "return false;" & vbNewLine

							Response.Write "oField = eval('oForm.PercentageForExcess_' + i);" & vbNewLine
							Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "if (! CheckFloatValue(oField, 'el porcentaje excedente', N_BOTH_FLAG, N_OPEN_FLAG, 0, 100))" & vbNewLine
								Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine

						Response.Write "oField = eval('oForm.InferiorLimit_' + lLastID);" & vbNewLine
						Response.Write "dPrevField = eval('oForm.SuperiorLimit_' + (lLastID - 1) + '.value');" & vbNewLine
						Response.Write "dPrevField = dPrevField.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "dPrevField = parseFloat(dPrevField);" & vbNewLine
						Response.Write "dNextField = eval('oForm.SuperiorLimit_' + lLastID + '.value');" & vbNewLine
						Response.Write "dNextField = dNextField.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "dNextField = parseFloat(dNextField);" & vbNewLine
						Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if ((! isNaN(dPrevField)) && (! isNaN(dNextField))) {" & vbNewLine
							Response.Write "if (! CheckFloatValue(oField, 'el límite inferior', N_BOTH_FLAG, N_OPEN_FLAG, dPrevField, dNextField))" & vbNewLine
								Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine

						Response.Write "oField = eval('oForm.AllowanceAmount_' + lLastID);" & vbNewLine
						Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oField, 'la cuota fija', N_MINIMUM_ONLY_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine

						Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "oField = eval('oForm.PercentageForExcess_' + lLastID);" & vbNewLine
						Response.Write "if (! CheckFloatValue(oField, 'el porcentaje excedente', N_BOTH_FLAG, N_OPEN_FLAG, 0, 100))" & vbNewLine
							Response.Write "return false;" & vbNewLine

					Response.Write "}" & vbNewLine

					Response.Write "return true;" & vbNewLine
				Response.Write "} // End of CheckTaxFields" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine

			Response.Write "<FORM NAME=""TaxFrm"" ID=""TaxFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""POST"" onSubmit=""return CheckTaxFields(this)"">"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">"
					Response.Write "<B>Vigencia a partir del " & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1) & "</B><BR /><BR />"
					Response.Write "<INPUT TYPE=""RADIO"" NAME=""PeriodID"" ID=""PeriodIDRd"" VALUE=""4"""
						If StrComp(oRequest("PeriodID").Item, "8", vbBinaryCompare) <> 0 Then Response.Write " CHECKED=""1"""
					Response.Write " onClick=""window.location.href = '" & GetASPFileName("") & "?Action=EmploymentAllowances&PeriodID=4'"" />&nbsp;Mensual"
					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
					Response.Write "<INPUT TYPE=""RADIO"" NAME=""PeriodID"" ID=""PeriodIDRd"" VALUE=""8"""
						If StrComp(oRequest("PeriodID").Item, "3", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
					Response.Write " onClick=""window.location.href = '" & GetASPFileName("") & "?Action=EmploymentAllowances&PeriodID=3'"" />&nbsp;Quincenal<BR /><BR />"
				Response.Write "</FONT>"
				Response.Write "<TABLE BORDER="""
					If Not bForExport Then
						Response.Write "0"
					Else
						Response.Write "1"
					End If
					Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
					asColumnsTitles = Split("Límite inferior,Límite superior,Monto", ",", -1, vbBinaryCompare)
					asCellWidths = Split(",,", ",", -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
					Else
						If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
							lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
						Else
							lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
						End If
					End If
					asCellAlignments = Split(",,", ",", -1, vbBinaryCompare)
					lFirstID = CLng(oRecordset.Fields("AllowanceID").Value)
					Do While Not oRecordset.EOF
						sRowContents = "<INPUT TYPE=""TEXT"" NAME=""InferiorLimit_" & CStr(oRecordset.Fields("AllowanceID").Value) & """ ID=""InferiorLimit_" & CStr(oRecordset.Fields("AllowanceID").Value) & "Txt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & FormatNumber(CDbl(oRecordset.Fields("InferiorLimit").Value), 2, True, False, True) & """ CLASS=""TextFields"" />"
						If CDbl(oRecordset.Fields("SuperiorLimit").Value) < 1000000 Then
							sRowContents = sRowContents & TABLE_SEPARATOR & "<INPUT TYPE=""TEXT"" NAME=""SuperiorLimit_" & CStr(oRecordset.Fields("AllowanceID").Value) & """ ID=""SuperiorLimit_" & CStr(oRecordset.Fields("AllowanceID").Value) & "Txt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & FormatNumber(CDbl(oRecordset.Fields("SuperiorLimit").Value), 2, True, False, True) & """ CLASS=""TextFields"" />"
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & "---<INPUT TYPE=""HIDDEN"" NAME=""SuperiorLimit_" & CStr(oRecordset.Fields("AllowanceID").Value) & """ ID=""SuperiorLimit_" & CStr(oRecordset.Fields("AllowanceID").Value) & "Hdn"" VALUE=""" & CStr(oRecordset.Fields("SuperiorLimit").Value) & """ />"
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & "<INPUT TYPE=""TEXT"" NAME=""AllowanceAmount_" & CStr(oRecordset.Fields("AllowanceID").Value) & """ ID=""AllowanceAmount_" & CStr(oRecordset.Fields("AllowanceID").Value) & "Txt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & FormatNumber(CDbl(oRecordset.Fields("AllowanceAmount").Value), 2, True, False, True) & """ CLASS=""TextFields"" />"

						lLastID = CLng(oRecordset.Fields("AllowanceID").Value)
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						oRecordset.MoveNext
					Loop
				Response.Write "</TABLE><BR />"
				Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
				Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
				Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""RedButtons"" onClick=""window.location.href = '" & GetASPFileName("") & "?Action=TaxInvertions&PeriodID=" & oRequest("PeriodID").Item & "'"" />"
			Response.Write "</FORM>"

			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "lFirstID = " & lFirstID & ";" & vbNewLine
				Response.Write "lLastID = " & lLastID & ";" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	DisplayEmploymentAllowancesTable = lErrorNumber
	Err.Clear
End Function

Function DisplayTaxInvertionsTable(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the tax invertions table
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayTaxInvertionsTable"
	Dim lFirstID
	Dim lLastID
	Dim sCondition
	Dim oRecordset
	Dim asColumnsTitles
	Dim sRowContents
	Dim asRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	sCondition = " And (PeriodID=4)"
	If Len(oRequest("PeriodID").Item) > 0 Then sCondition = " And (PeriodID=" & oRequest("PeriodID").Item & ")"
	sErrorDescription = "No se pudo obtener el monto del concepto."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From TaxInvertions Where (EndDate=30000000) " & sCondition & " Order By InferiorLimit", "ReportsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "var lFirstID = 0;" & vbNewLine
				Response.Write "var lLastID = 0;" & vbNewLine

				Response.Write "function CheckTaxFields(oForm) {" & vbNewLine
					Response.Write "var oField = null;" & vbNewLine
					Response.Write "var dPrevField = 0;" & vbNewLine
					Response.Write "var dNextField = 0;" & vbNewLine

					Response.Write "if (oForm) {" & vbNewLine
						Response.Write "oField = eval('oForm.SuperiorLimit_' + lFirstID);" & vbNewLine
						Response.Write "dPrevField = eval('oForm.InferiorLimit_' + lFirstID + '.value');" & vbNewLine
						Response.Write "dPrevField = dPrevField.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "dPrevField = parseFloat(dPrevField);" & vbNewLine
						Response.Write "dNextField = eval('oForm.InferiorLimit_' + (lFirstID + 1) + '.value');" & vbNewLine
						Response.Write "dNextField = dNextField.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "dNextField = parseFloat(dNextField);" & vbNewLine
						Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if ((! isNaN(dPrevField)) && (! isNaN(dNextField))) {" & vbNewLine
							Response.Write "if (! CheckFloatValue(oField, 'el límite superior', N_BOTH_FLAG, N_OPEN_FLAG, dPrevField, dNextField))" & vbNewLine
								Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine

						Response.Write "oField = eval('oForm.InvertedTax_' + lFirstID);" & vbNewLine
						Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oField, 'la cuota fija', N_MINIMUM_ONLY_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine

						Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "oField = eval('oForm.InvertedRate_' + lFirstID);" & vbNewLine
						Response.Write "if (! CheckFloatValue(oField, 'el porcentaje excedente', N_BOTH_FLAG, N_OPEN_FLAG, 0, 100))" & vbNewLine
							Response.Write "return false;" & vbNewLine

						Response.Write "for (var i=lFirstID+1; i<lLastID; i++) {" & vbNewLine
							Response.Write "oField = eval('oForm.InferiorLimit_' + i);" & vbNewLine
							Response.Write "dPrevField = eval('oForm.SuperiorLimit_' + (i - 1) + '.value');" & vbNewLine
							Response.Write "dPrevField = dPrevField.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "dPrevField = parseFloat(dPrevField);" & vbNewLine
							Response.Write "dNextField = eval('oForm.SuperiorLimit_' + i + '.value');" & vbNewLine
							Response.Write "dNextField = dNextField.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "dNextField = parseFloat(dNextField);" & vbNewLine
							Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "if ((! isNaN(dPrevField)) && (! isNaN(dNextField))) {" & vbNewLine
								Response.Write "if (! CheckFloatValue(oField, 'el límite inferior', N_BOTH_FLAG, N_OPEN_FLAG, dPrevField, dNextField))" & vbNewLine
									Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine

							Response.Write "oField = eval('oForm.SuperiorLimit_' + i);" & vbNewLine
							Response.Write "dPrevField = eval('oForm.InferiorLimit_' + i + '.value');" & vbNewLine
							Response.Write "dPrevField = dPrevField.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "dPrevField = parseFloat(dPrevField);" & vbNewLine
							Response.Write "dNextField = eval('oForm.InferiorLimit_' + (i + 1) + '.value');" & vbNewLine
							Response.Write "dNextField = dNextField.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "dNextField = parseFloat(dNextField);" & vbNewLine
							Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "if ((! isNaN(dPrevField)) && (! isNaN(dNextField))) {" & vbNewLine
								Response.Write "if (! CheckFloatValue(oField, 'el límite superior', N_BOTH_FLAG, N_OPEN_FLAG, dPrevField, dNextField))" & vbNewLine
									Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine

							Response.Write "oField = eval('oForm.InvertedTax_' + i);" & vbNewLine
							Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "if (! CheckFloatValue(oField, 'la cuota fija', N_MINIMUM_ONLY_FLAG, N_OPEN_FLAG, 0, 0))" & vbNewLine
								Response.Write "return false;" & vbNewLine

							Response.Write "oField = eval('oForm.InvertedRate_' + i);" & vbNewLine
							Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "if (! CheckFloatValue(oField, 'el porcentaje excedente', N_BOTH_FLAG, N_OPEN_FLAG, 0, 100))" & vbNewLine
								Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine

						Response.Write "oField = eval('oForm.InferiorLimit_' + lLastID);" & vbNewLine
						Response.Write "dPrevField = eval('oForm.SuperiorLimit_' + (lLastID - 1) + '.value');" & vbNewLine
						Response.Write "dPrevField = dPrevField.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "dPrevField = parseFloat(dPrevField);" & vbNewLine
						Response.Write "dNextField = eval('oForm.SuperiorLimit_' + lLastID + '.value');" & vbNewLine
						Response.Write "dNextField = dNextField.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "dNextField = parseFloat(dNextField);" & vbNewLine
						Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if ((! isNaN(dPrevField)) && (! isNaN(dNextField))) {" & vbNewLine
							Response.Write "if (! CheckFloatValue(oField, 'el límite inferior', N_BOTH_FLAG, N_OPEN_FLAG, dPrevField, dNextField))" & vbNewLine
								Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine

						Response.Write "oField = eval('oForm.InvertedTax_' + lLastID);" & vbNewLine
						Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oField, 'la cuota fija', N_MINIMUM_ONLY_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine

						Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "oField = eval('oForm.InvertedRate_' + lLastID);" & vbNewLine
						Response.Write "if (! CheckFloatValue(oField, 'el porcentaje excedente', N_BOTH_FLAG, N_OPEN_FLAG, 0, 100))" & vbNewLine
							Response.Write "return false;" & vbNewLine

					Response.Write "}" & vbNewLine

					Response.Write "return true;" & vbNewLine
				Response.Write "} // End of CheckTaxFields" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine

			Response.Write "<FORM NAME=""TaxFrm"" ID=""TaxFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""POST"" onSubmit=""return CheckTaxFields(this)"">"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">"
					Response.Write "<B>Vigencia a partir del " & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1) & "</B><BR /><BR />"
					Response.Write "<INPUT TYPE=""RADIO"" NAME=""PeriodID"" ID=""PeriodIDRd"" VALUE=""4"""
						If StrComp(oRequest("PeriodID").Item, "8", vbBinaryCompare) <> 0 Then Response.Write " CHECKED=""1"""
					Response.Write " onClick=""window.location.href = '" & GetASPFileName("") & "?Action=TaxInvertions&PeriodID=4'"" />&nbsp;Mensual"
					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
					Response.Write "<INPUT TYPE=""RADIO"" NAME=""PeriodID"" ID=""PeriodIDRd"" VALUE=""8"""
						If StrComp(oRequest("PeriodID").Item, "8", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
					Response.Write " onClick=""window.location.href = '" & GetASPFileName("") & "?Action=TaxInvertions&PeriodID=8'"" />&nbsp;Anual<BR /><BR />"
				Response.Write "</FONT>"
				Response.Write "<TABLE BORDER="""
					If Not bForExport Then
						Response.Write "0"
					Else
						Response.Write "1"
					End If
					Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
					asColumnsTitles = Split("Límite inferior,Límite superior,Cuota fija,% excedente", ",", -1, vbBinaryCompare)
					asCellWidths = Split(",,,", ",", -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
					Else
						If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
							lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
						Else
							lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
						End If
					End If
					asCellAlignments = Split(",,,", ",", -1, vbBinaryCompare)
					lFirstID = CLng(oRecordset.Fields("TaxID").Value)
					Do While Not oRecordset.EOF
						sRowContents = "<INPUT TYPE=""TEXT"" NAME=""InferiorLimit_" & CStr(oRecordset.Fields("TaxID").Value) & """ ID=""InferiorLimit_" & CStr(oRecordset.Fields("TaxID").Value) & "Txt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & FormatNumber(CDbl(oRecordset.Fields("InferiorLimit").Value), 2, True, False, True) & """ CLASS=""TextFields"" />"
						If CDbl(oRecordset.Fields("SuperiorLimit").Value) < 1000000 Then
							sRowContents = sRowContents & TABLE_SEPARATOR & "<INPUT TYPE=""TEXT"" NAME=""SuperiorLimit_" & CStr(oRecordset.Fields("TaxID").Value) & """ ID=""SuperiorLimit_" & CStr(oRecordset.Fields("TaxID").Value) & "Txt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & FormatNumber(CDbl(oRecordset.Fields("SuperiorLimit").Value), 2, True, False, True) & """ CLASS=""TextFields"" />"
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & "---<INPUT TYPE=""HIDDEN"" NAME=""SuperiorLimit_" & CStr(oRecordset.Fields("TaxID").Value) & """ ID=""SuperiorLimit_" & CStr(oRecordset.Fields("TaxID").Value) & "Hdn"" VALUE=""" & CStr(oRecordset.Fields("SuperiorLimit").Value) & """ />"
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & "<INPUT TYPE=""TEXT"" NAME=""InvertedTax_" & CStr(oRecordset.Fields("TaxID").Value) & """ ID=""InvertedTax_" & CStr(oRecordset.Fields("TaxID").Value) & "Txt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & FormatNumber(CDbl(oRecordset.Fields("InvertedTax").Value), 2, True, False, True) & """ CLASS=""TextFields"" />"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<INPUT TYPE=""TEXT"" NAME=""InvertedRate_" & CStr(oRecordset.Fields("TaxID").Value) & """ ID=""InvertedRate_" & CStr(oRecordset.Fields("TaxID").Value) & "Txt"" SIZE=""5"" MAXLENGTH=""5"" VALUE=""" & FormatNumber((CDbl(oRecordset.Fields("InvertedRate").Value) * 100), 2, True, False, True) & """ CLASS=""TextFields"" />"

						lLastID = CLng(oRecordset.Fields("TaxID").Value)
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						oRecordset.MoveNext
					Loop
				Response.Write "</TABLE><BR />"
				Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
				Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
				Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""RedButtons"" onClick=""window.location.href = '" & GetASPFileName("") & "?Action=TaxInvertions&PeriodID=" & oRequest("PeriodID").Item & "'"" />"
			Response.Write "</FORM>"

			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "lFirstID = " & lFirstID & ";" & vbNewLine
				Response.Write "lLastID = " & lLastID & ";" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	DisplayTaxInvertionsTable = lErrorNumber
	Err.Clear
End Function

Function DisplayTaxLimitsTable(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the tax limits table
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayTaxLimitsTable"
	Dim lFirstID
	Dim lLastID
	Dim sCondition
	Dim oRecordset
	Dim asColumnsTitles
	Dim sRowContents
	Dim asRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	sCondition = " And (PeriodID=4)"
	If Len(oRequest("PeriodID").Item) > 0 Then sCondition = " And (PeriodID=" & oRequest("PeriodID").Item & ")"
	sErrorDescription = "No se pudo obtener el monto del concepto."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From TaxLimits Where (EndDate=30000000) " & sCondition & " Order By InferiorLimit", "ReportsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "var lFirstID = 0;" & vbNewLine
				Response.Write "var lLastID = 0;" & vbNewLine

				Response.Write "function CheckTaxFields(oForm) {" & vbNewLine
					Response.Write "var oField = null;" & vbNewLine
					Response.Write "var dPrevField = 0;" & vbNewLine
					Response.Write "var dNextField = 0;" & vbNewLine

					Response.Write "if (oForm) {" & vbNewLine
						Response.Write "oField = eval('oForm.SuperiorLimit_' + lFirstID);" & vbNewLine
						Response.Write "dPrevField = eval('oForm.InferiorLimit_' + lFirstID + '.value');" & vbNewLine
						Response.Write "dPrevField = dPrevField.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "dPrevField = parseFloat(dPrevField);" & vbNewLine
						Response.Write "dNextField = eval('oForm.InferiorLimit_' + (lFirstID + 1) + '.value');" & vbNewLine
						Response.Write "dNextField = dNextField.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "dNextField = parseFloat(dNextField);" & vbNewLine
						Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if ((! isNaN(dPrevField)) && (! isNaN(dNextField))) {" & vbNewLine
							Response.Write "if (! CheckFloatValue(oField, 'el límite superior', N_BOTH_FLAG, N_OPEN_FLAG, dPrevField, dNextField))" & vbNewLine
								Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine

						Response.Write "oField = eval('oForm.FixedAmount_' + lFirstID);" & vbNewLine
						Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oField, 'la cuota fija', N_MINIMUM_ONLY_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine

						Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "oField = eval('oForm.PercentageForExcess_' + lFirstID);" & vbNewLine
						Response.Write "if (! CheckFloatValue(oField, 'el porcentaje excedente', N_BOTH_FLAG, N_OPEN_FLAG, 0, 100))" & vbNewLine
							Response.Write "return false;" & vbNewLine

						Response.Write "for (var i=lFirstID+1; i<lLastID; i++) {" & vbNewLine
							Response.Write "oField = eval('oForm.InferiorLimit_' + i);" & vbNewLine
							Response.Write "dPrevField = eval('oForm.SuperiorLimit_' + (i - 1) + '.value');" & vbNewLine
							Response.Write "dPrevField = dPrevField.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "dPrevField = parseFloat(dPrevField);" & vbNewLine
							Response.Write "dNextField = eval('oForm.SuperiorLimit_' + i + '.value');" & vbNewLine
							Response.Write "dNextField = dNextField.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "dNextField = parseFloat(dNextField);" & vbNewLine
							Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "if ((! isNaN(dPrevField)) && (! isNaN(dNextField))) {" & vbNewLine
								Response.Write "if (! CheckFloatValue(oField, 'el límite inferior', N_BOTH_FLAG, N_OPEN_FLAG, dPrevField, dNextField))" & vbNewLine
									Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine

							Response.Write "oField = eval('oForm.SuperiorLimit_' + i);" & vbNewLine
							Response.Write "dPrevField = eval('oForm.InferiorLimit_' + i + '.value');" & vbNewLine
							Response.Write "dPrevField = dPrevField.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "dPrevField = parseFloat(dPrevField);" & vbNewLine
							Response.Write "dNextField = eval('oForm.InferiorLimit_' + (i + 1) + '.value');" & vbNewLine
							Response.Write "dNextField = dNextField.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "dNextField = parseFloat(dNextField);" & vbNewLine
							Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "if ((! isNaN(dPrevField)) && (! isNaN(dNextField))) {" & vbNewLine
								Response.Write "if (! CheckFloatValue(oField, 'el límite superior', N_BOTH_FLAG, N_OPEN_FLAG, dPrevField, dNextField))" & vbNewLine
									Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine

							Response.Write "oField = eval('oForm.FixedAmount_' + i);" & vbNewLine
							Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "if (! CheckFloatValue(oField, 'la cuota fija', N_MINIMUM_ONLY_FLAG, N_OPEN_FLAG, 0, 0))" & vbNewLine
								Response.Write "return false;" & vbNewLine

							Response.Write "oField = eval('oForm.PercentageForExcess_' + i);" & vbNewLine
							Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "if (! CheckFloatValue(oField, 'el porcentaje excedente', N_BOTH_FLAG, N_OPEN_FLAG, 0, 100))" & vbNewLine
								Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine

						Response.Write "oField = eval('oForm.InferiorLimit_' + lLastID);" & vbNewLine
						Response.Write "dPrevField = eval('oForm.SuperiorLimit_' + (lLastID - 1) + '.value');" & vbNewLine
						Response.Write "dPrevField = dPrevField.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "dPrevField = parseFloat(dPrevField);" & vbNewLine
						Response.Write "dNextField = eval('oForm.SuperiorLimit_' + lLastID + '.value');" & vbNewLine
						Response.Write "dNextField = dNextField.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "dNextField = parseFloat(dNextField);" & vbNewLine
						Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if ((! isNaN(dPrevField)) && (! isNaN(dNextField))) {" & vbNewLine
							Response.Write "if (! CheckFloatValue(oField, 'el límite inferior', N_BOTH_FLAG, N_OPEN_FLAG, dPrevField, dNextField))" & vbNewLine
								Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine

						Response.Write "oField = eval('oForm.FixedAmount_' + lLastID);" & vbNewLine
						Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oField, 'la cuota fija', N_MINIMUM_ONLY_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine

						Response.Write "oField.value = oField.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "oField = eval('oForm.PercentageForExcess_' + lLastID);" & vbNewLine
						Response.Write "if (! CheckFloatValue(oField, 'el porcentaje excedente', N_BOTH_FLAG, N_OPEN_FLAG, 0, 100))" & vbNewLine
							Response.Write "return false;" & vbNewLine

					Response.Write "}" & vbNewLine

					Response.Write "return true;" & vbNewLine
				Response.Write "} // End of CheckTaxFields" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine

			Response.Write "<FORM NAME=""TaxFrm"" ID=""TaxFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""POST"" onSubmit=""return CheckTaxFields(this)"">"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">"
					Response.Write "<B>Vigencia a partir del " & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1) & "</B><BR /><BR />"
					Response.Write "<INPUT TYPE=""RADIO"" NAME=""PeriodID"" ID=""PeriodIDRd"" VALUE=""4"""
						If StrComp(oRequest("PeriodID").Item, "8", vbBinaryCompare) <> 0 Then Response.Write " CHECKED=""1"""
					Response.Write " onClick=""window.location.href = '" & GetASPFileName("") & "?Action=TaxLimits&PeriodID=4'"" />&nbsp;Mensual"
					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
					Response.Write "<INPUT TYPE=""RADIO"" NAME=""PeriodID"" ID=""PeriodIDRd"" VALUE=""8"""
						If StrComp(oRequest("PeriodID").Item, "3", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
					Response.Write " onClick=""window.location.href = '" & GetASPFileName("") & "?Action=TaxLimits&PeriodID=3'"" />&nbsp;Quincenal<BR /><BR />"
				Response.Write "</FONT>"
				Response.Write "<TABLE BORDER="""
					If Not bForExport Then
						Response.Write "0"
					Else
						Response.Write "1"
					End If
					Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
					asColumnsTitles = Split("Límite inferior,Límite superior,Cuota fija,% excedente", ",", -1, vbBinaryCompare)
					asCellWidths = Split(",,,", ",", -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
					Else
						If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
							lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
						Else
							lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
						End If
					End If
					asCellAlignments = Split(",,,", ",", -1, vbBinaryCompare)
					lFirstID = CLng(oRecordset.Fields("TaxLimitID").Value)
					Do While Not oRecordset.EOF
						sRowContents = "<INPUT TYPE=""TEXT"" NAME=""InferiorLimit_" & CStr(oRecordset.Fields("TaxLimitID").Value) & """ ID=""InferiorLimit_" & CStr(oRecordset.Fields("TaxLimitID").Value) & "Txt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & FormatNumber(CDbl(oRecordset.Fields("InferiorLimit").Value), 2, True, False, True) & """ CLASS=""TextFields"" />"
						If CDbl(oRecordset.Fields("SuperiorLimit").Value) < 1000000 Then
							sRowContents = sRowContents & TABLE_SEPARATOR & "<INPUT TYPE=""TEXT"" NAME=""SuperiorLimit_" & CStr(oRecordset.Fields("TaxLimitID").Value) & """ ID=""SuperiorLimit_" & CStr(oRecordset.Fields("TaxLimitID").Value) & "Txt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & FormatNumber(CDbl(oRecordset.Fields("SuperiorLimit").Value), 2, True, False, True) & """ CLASS=""TextFields"" />"
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & "---<INPUT TYPE=""HIDDEN"" NAME=""SuperiorLimit_" & CStr(oRecordset.Fields("TaxLimitID").Value) & """ ID=""SuperiorLimit_" & CStr(oRecordset.Fields("TaxLimitID").Value) & "Hdn"" VALUE=""" & CStr(oRecordset.Fields("SuperiorLimit").Value) & """ />"
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & "<INPUT TYPE=""TEXT"" NAME=""FixedAmount_" & CStr(oRecordset.Fields("TaxLimitID").Value) & """ ID=""FixedAmount_" & CStr(oRecordset.Fields("TaxLimitID").Value) & "Txt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & FormatNumber(CDbl(oRecordset.Fields("FixedAmount").Value), 2, True, False, True) & """ CLASS=""TextFields"" />"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<INPUT TYPE=""TEXT"" NAME=""PercentageForExcess_" & CStr(oRecordset.Fields("TaxLimitID").Value) & """ ID=""PercentageForExcess_" & CStr(oRecordset.Fields("TaxLimitID").Value) & "Txt"" SIZE=""5"" MAXLENGTH=""5"" VALUE=""" & FormatNumber((CDbl(oRecordset.Fields("PercentageForExcess").Value) * 100), 2, True, False, True) & """ CLASS=""TextFields"" />"

						lLastID = CLng(oRecordset.Fields("TaxLimitID").Value)
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						oRecordset.MoveNext
					Loop
				Response.Write "</TABLE><BR />"
				Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
				Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
				Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""RedButtons"" onClick=""window.location.href = '" & GetASPFileName("") & "?Action=TaxInvertions&PeriodID=" & oRequest("PeriodID").Item & "'"" />"
			Response.Write "</FORM>"

			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "lFirstID = " & lFirstID & ";" & vbNewLine
				Response.Write "lLastID = " & lLastID & ";" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	DisplayTaxLimitsTable = lErrorNumber
	Err.Clear
End Function
%>