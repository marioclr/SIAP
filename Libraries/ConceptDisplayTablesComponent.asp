<%
Function DisplayConceptsTable(oRequest, oADODBConnection, lIDColumn, bUseLinks, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the concepts
'		  from the database in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, aConceptComponent
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayConceptsTable"
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

	iStatusID = aConceptComponent(N_STATUS_ID_CONCEPT)
	If iStatusID=0 Then
		sCondition = sCondition & " And (Concepts.StatusID<=" & iStatusID & ")"
	Else
		sCondition = sCondition & " And (Concepts.StatusID=" & iStatusID & ")"
	End If
	aConceptComponent(S_QUERY_CONDITION_CONCEPT) = aConceptComponent(S_QUERY_CONDITION_CONCEPT) & sCondition
	lErrorNumber = GetConcepts(oRequest, oADODBConnection, aConceptComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If Not bForExport Then Call DisplayIncrementalFetch(oRequest, CInt(oRequest("StartPage").Item), ROWS_CATALOG, oRecordset)
			Response.Write "<TABLE WIDTH=""550"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bUseLinks Then
					asColumnsTitles = Split("&nbsp;,Clave,Nombre,Fecha inicio,Fecha Fin,Acciones", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,80,270,90,90,80", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("&nbsp;,Clave,Nombre", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,80,350,90,90", ",", -1, vbBinaryCompare)
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
					If (StrComp(CStr(oRecordset.Fields("ConceptID").Value), oRequest("ConceptID").Item, vbBinaryCompare) = 0) And (CLng(oRecordset.Fields("StartDate").Value) = CLng(oRequest("StartDate").Item)) Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					If CInt(oRecordset.Fields("IsDeduction").Value) = 1 Then
						sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
						sFontEnd = "</FONT>"
					End If
					sRowContents = ""
					Select Case lIDColumn
						Case DISPLAY_RADIO_BUTTONS
							sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""" & CStr(oRecordset.Fields("ConceptID").Value) & """ />"
						Case DISPLAY_CHECKBOXES
							sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""ConceptID"" ID=""ConceptIDChk"" VALUE=""" & CStr(oRecordset.Fields("ConceptID").Value) & """ />"
						Case Else
							sRowContents = sRowContents & "&nbsp;"
					End Select
					sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
						sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&StartDate=" & CStr(oRecordset.Fields("StartDate").Value) & """"
					sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & sBoldEnd & sFontEnd & "</A>"
					sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
						sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptID=" & CStr(oRecordset.Fields("ConceptID").Value) & "&StartDate=" & CStr(oRecordset.Fields("StartDate").Value) & """"
					sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value)) & sBoldEnd & sFontEnd & "</A>"
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value))
					If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "A la fecha"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value))
					End If
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
	DisplayConceptsTable = lErrorNumber
	Err.Clear
End Function

Function DisplayConceptValuesTableSP(oRequest, oADODBConnection, iSelectedTab, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the ConceptValues for Concepts
'Inputs:  oRequest, oADODBConnection, iSelectedTab, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayConceptValuesTableSP"
	Dim sCondition
	Dim sCondition2
	Dim lPayrollID
	Dim lForPayrollID
	Dim lPayrollNumber
	Dim sDate
	Dim sStartDate
	Dim sEndDate
	Dim sFilePath
	Dim lReportID
	Dim sTemp
	Dim lCurrentID
	Dim dTotal
	Dim oRecordset
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asColumnsTitles
	Dim asCellWidths
	Dim asCellAlignments
	Dim sColumnsTitles
	Dim sCellWidths
	Dim sCellAlignments
	Dim lErrorNumber
	Dim sBoldBegin
	Dim sBoldEnd
	Dim sFontBegin
	Dim sFontEnd
	Dim bFirst
	Dim bContinue
	Dim sConceptIDs
	DIm iStatusID
	Dim sRecordIDs
	Dim sStartDateCondition

	Dim iLevelID
	Dim iAntiquityID
	Dim iAntiquityID2
	Dim iEconomicZoneID
	Dim iGroupGradeLevelID
	Dim iClassificationID
	Dim iIntegrationID
	Dim iPositionTypeID
	Dim sPositionTypeShortName
	Dim sPositionShortName
	Dim sLevelShortName
	Dim sWorkingHours
	Dim sPositionName
	Dim sGroupGradeLevelShortName
	Dim lCurrentPositionID
	Dim iCompanyID
	Dim sCompanyName

	Dim dConcept
	Dim lConcept_RecordID
	Dim lConcept_StartDate
	Dim bActiveConcept

	Dim dConcept_01
	Dim lConcept_01_RecordID
	Dim lConcept_01_StartDate
	Dim bActiveConcept_01

	Dim dConcept_03
	Dim lConcept_03_RecordID
	Dim lConcept_03_StartDate
	Dim bActiveConcept_03

	Dim dConcept_12
	Dim lConcept_12_RecordID
	Dim lConcept_12_StartDate
	Dim bActiveConcept_12

	Dim dConcept_35
	Dim lConcept_35_RecordID
	Dim lConcept_35_StartDate
	Dim bActiveConcept_35

	Dim dConcept_36
	Dim lConcept_36_RecordID
	Dim lConcept_36_StartDate
	Dim bActiveConcept_36

	Dim dConcept_48
	Dim lConcept_48_RecordID
	Dim lConcept_48_StartDate
	Dim bActiveConcept_48

	Dim dConcept_B2
	Dim lConcept_B2_RecordID
	Dim lConcept_B2_StartDate
	Dim bActiveConcept_B2

	Dim dConcept_Z3
	Dim lConcept_Z3_RecordID
	Dim lConcept_Z3_StartDate
	Dim bActiveConcept_Z3

	Dim dConcept_01_Z3
	Dim lConcept_01_Z3_RecordID
	Dim lConcept_01_Z3_StartDate
	Dim bActiveConcept_01_Z3

	Dim dConcept_03_Z3
	Dim lConcept_03_Z3_RecordID
	Dim lConcept_03_Z3_StartDate
	Dim bActiveConcept_03_Z3

	Dim dConcept_12_Z3
	Dim lConcept_12_Z3_RecordID
	Dim lConcept_12_Z3_StartDate
	Dim bActiveConcept_12_Z3

	Dim dConcept_35_Z3
	Dim lConcept_35_Z3_RecordID
	Dim lConcept_35_Z3_StartDate
	Dim bActiveConcept_35_Z3

	Dim dConcept_36_Z3
	Dim lConcept_36_Z3_RecordID
	Dim lConcept_36_Z3_StartDate
	Dim bActiveConcept_36_Z3

	Dim dConcept_48_Z3
	Dim lConcept_48_Z3_RecordID
	Dim lConcept_48_Z3_StartDate
	Dim bActiveConcept_48_Z3

	Dim dConcept_B2_Z3
	Dim lConcept_B2_Z3_RecordID
	Dim lConcept_B2_Z3_StartDate
	Dim bActiveConcept_B2_Z3

	sDate = Left(GetSerialNumberForDate(""), Len("00000000"))

	If aConceptComponent(N_STATUS_ID_CONCEPT) = 1 Then
		Call GetStartAndEndDatesFromURL("StartForValue1", "EndForValue", "ConceptsValues.StartDate", False, sCondition)
		sStartDateCondition = sCondition
	End If
	If (Len(oRequest("PositionID").Item) > 0) And (aConceptComponent(N_STATUS_ID_CONCEPT) > 0) Then
		If CInt(oRequest("PositionID").Item) <> -1 Then
			sCondition = sCondition & " And (Positions.PositionID In (" & oRequest("PositionID").Item & "))"
		End If
	Else
		If (aConceptComponent(N_STATUS_ID_CONCEPT) = 1) And (Not bForExport) Then
				sCondition = sCondition & " And (Positions.PositionID In (0))"
		End If
	End If
	'If (CInt(oRequest("StartForValueYear").Item) > 0) And (CInt(oRequest("StartForValueMonth").Item) > 0) And (CInt(oRequest("StartForValueDay").Item) > 0) And (CInt(oRequest("EndForValueYear").Item) > 0) And (CInt(oRequest("EndForValueYearMonth").Item) > 0) And (CInt(oRequest("EndForValueYearDay").Item) > 0) Then Call GetStartAndEndDatesFromURL("StartForValue", "EndForValue", "ConceptsValues.StartDate", True, sCondition)
	'If (CInt(oRequest("StartJobEndYear").Item) > 0) And (CInt(oRequest("StartJobEndMonth").Item) > 0) And (CInt(oRequest("StartJobEndDay").Item) > 0) And (CInt(oRequest("EndJobEndYear").Item) > 0) And (CInt(oRequest("EndJobEndMonth").Item) > 0) And (CInt(oRequest("EndJobEndDay").Item) > 0) Then Call GetStartAndEndDatesFromURL("StartJobEnd", "EndJobEnd", "Jobs.EndDate", False, sCondition)
	sErrorDescription = "No se pudieron obtener los montos pagados."
	iStatusID = aConceptComponent(N_STATUS_ID_CONCEPT)
	sCondition = sCondition & " And (ConceptsValues.StatusID=" & iStatusID & ")"
	If Len(oRequest("ConceptID").Item) > 0 Then
		lErrorNumber = GetConcept(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
		sStartDate = aConceptComponent(N_START_DATE_CONCEPT)
		sEndDate = aConceptComponent(N_END_DATE_CONCEPT)
		Select Case iSelectedTab
			Case 0
				If InStr(1, ",1,38,49,", "," & CStr(oRequest("ConceptID").Item) & ",", vbBinaryCompare) > 0 Then
					sCondition = sCondition & " And (ConceptID IN (1, 38, 49))"
				Else
					sCondition = sCondition & " And (ConceptID=" & CStr(oRequest("ConceptID").Item) & ")"
				End If
			Case 1, 2, 3, 4
				If InStr(1, ",1,3,", "," & CStr(oRequest("ConceptID").Item) & ",", vbBinaryCompare) > 0 Then
					sCondition = sCondition & " And (ConceptID IN (1, 3))"
				Else
					sCondition = sCondition & " And (ConceptID=" & CStr(oRequest("ConceptID").Item) & ")"
				End If
			Case 5
				If InStr(1, ",39,89,", "," & CStr(oRequest("ConceptID").Item) & ",", vbBinaryCompare) > 0 Then
					sCondition = sCondition & " And (ConceptID IN (39, 89))"
				Else
					sCondition = sCondition & " And (ConceptID=" & CStr(oRequest("ConceptID").Item) & ")"
				End If
			Case 6
				If InStr(1, ",14,", "," & CStr(oRequest("ConceptID").Item) & ",", vbBinaryCompare) > 0 Then
					sCondition = sCondition & " And (ConceptID IN (14))"
				Else
					sCondition = sCondition & " And (ConceptID=" & CStr(oRequest("ConceptID").Item) & ")"
				End If
			Case Else
				sCondition = sCondition & " And (ConceptID=" & CStr(oRequest("ConceptID").Item) & ")"
		End Select
	Else
		sStartDate = sDate
		sEndDate = sDate
		Select Case iSelectedTab
			Case 0
				sCondition = sCondition & " And (ConceptID IN (1, 38, 49))"
			Case 1, 2, 3, 4
				sCondition = sCondition & " And (ConceptID IN (1, 3))"
			Case 5
				sCondition = sCondition & " And (ConceptID IN (39, 89))"
			Case 6
				sCondition = sCondition & " And (ConceptID IN (14))"
		End Select
	End If
	'sCondition = sCondition & " And (((ConceptsValues.StartDate>=" & sStartDate & ") And (ConceptsValues.StartDate<=" & sEndDate & ")) Or ((ConceptsValues.EndDate>=" & sStartDate & ") And (ConceptsValues.EndDate<=" & sEndDate & ")) Or ((ConceptsValues.EndDate>=" & sStartDate & ") And (ConceptsValues.StartDate<=" & sEndDate & ")))"
	If Len(oRequest("ConceptID").Item) > 0 Then
		sCondition = sCondition & " And (((Companies.StartDate>="& sStartDate & ") And (Companies.StartDate<=" & sEndDate & ")) Or ((Companies.EndDate>=" & sStartDate & ") And (Companies.EndDate<=" & sEndDate & ")) Or ((Companies.EndDate>=" & sStartDate & ") And (Companies.StartDate<=" & sEndDate & ")))"
		sCondition = sCondition & " And (((Positions.StartDate>="& sStartDate & ") And (Positions.StartDate<=" & sEndDate & ")) Or ((Positions.EndDate>=" & sStartDate & ") And (Positions.EndDate<=" & sEndDate & ")) Or ((Positions.EndDate>=" & sStartDate & ") And (Positions.StartDate<=" & sEndDate & ")))"
		sCondition = sCondition & " And (((PositionTypes.StartDate>=" & sStartDate & ") And (PositionTypes.StartDate<=" & sEndDate & ")) Or ((PositionTypes.EndDate>=" & sStartDate & ") And (PositionTypes.EndDate<=" & sEndDate & ")) Or ((PositionTypes.EndDate>=" & sStartDate & ") And (PositionTypes.StartDate<=" & sEndDate & ")))"
		sCondition = sCondition & " And (((GroupGradeLevels.StartDate>=" & sStartDate & ") And (GroupGradeLevels.StartDate<=" & sEndDate & ")) Or ((GroupGradeLevels.EndDate>=" & sStartDate & ") And (GroupGradeLevels.EndDate<=" & sEndDate & ")) Or ((GroupGradeLevels.EndDate>=" & sStartDate & ") And (GroupGradeLevels.StartDate<=" & sEndDate & ")))"
		sCondition = sCondition & " And (((Levels.StartDate>=" & sStartDate & ") And (Levels.StartDate<=" & sEndDate & ")) Or ((Levels.EndDate>=" & sStartDate & ") And (Levels.EndDate<=" & sEndDate & ")) Or ((Levels.EndDate>=" & sStartDate & ") And (Levels.StartDate<=" & sEndDate & ")))"
		sCondition = sCondition & " And (ConceptsValues.EmployeeTypeID IN (-1, " & iSelectedTab & "))" & " And (Positions.EmployeeTypeID IN (-1, " & iSelectedTab & "))"
	Else
		sCondition = sCondition & " And (((Companies.StartDate>=ConceptsValues.StartDate) And (Companies.StartDate<=ConceptsValues.EndDate)) Or ((Companies.EndDate>=ConceptsValues.StartDate) And (Companies.EndDate<=ConceptsValues.EndDate)) Or ((Companies.EndDate>=ConceptsValues.StartDate) And (Companies.StartDate<=ConceptsValues.EndDate)))"
		sCondition = sCondition & " And (((Positions.StartDate>=ConceptsValues.StartDate) And (Positions.StartDate<=ConceptsValues.EndDate)) Or ((Positions.EndDate>=ConceptsValues.StartDate) And (Positions.EndDate<=ConceptsValues.EndDate)) Or ((Positions.EndDate>=ConceptsValues.StartDate) And (Positions.StartDate<=ConceptsValues.EndDate)))"
		sCondition = sCondition & " And (((PositionTypes.StartDate>=ConceptsValues.StartDate) And (PositionTypes.StartDate<=ConceptsValues.EndDate)) Or ((PositionTypes.EndDate>=ConceptsValues.StartDate) And (PositionTypes.EndDate<=ConceptsValues.EndDate)) Or ((PositionTypes.EndDate>=ConceptsValues.StartDate) And (PositionTypes.StartDate<=ConceptsValues.EndDate)))"
		sCondition = sCondition & " And (((GroupGradeLevels.StartDate>=ConceptsValues.StartDate) And (GroupGradeLevels.StartDate<=ConceptsValues.EndDate)) Or ((GroupGradeLevels.EndDate>=ConceptsValues.StartDate) And (GroupGradeLevels.EndDate<=ConceptsValues.EndDate)) Or ((GroupGradeLevels.EndDate>=ConceptsValues.StartDate) And (GroupGradeLevels.StartDate<=ConceptsValues.EndDate)))"
		sCondition = sCondition & " And (((Levels.StartDate>=ConceptsValues.StartDate) And (Levels.StartDate<=ConceptsValues.EndDate)) Or ((Levels.EndDate>=ConceptsValues.StartDate) And (Levels.EndDate<=ConceptsValues.EndDate)) Or ((Levels.EndDate>=ConceptsValues.StartDate) And (Levels.StartDate<=ConceptsValues.EndDate)))"
		sCondition = sCondition & " And (ConceptsValues.EmployeeTypeID IN (-1, " & iSelectedTab & "))" & " And (Positions.EmployeeTypeID IN (-1, " & iSelectedTab & "))"
	End If
	'sCondition = sCondition & " And (ConceptsValues.PositionID=2)"
	aConceptComponent(S_QUERY_CONDITION_CONCEPT) = sCondition

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptsValues.RecordID, ConceptsValues.ConceptID, ConceptsValues.ConceptAmount, ConceptsValues.StartDate, ConceptsValues.EndDate, ConceptsValues.StatusID, Positions.PositionID, ConceptsValues.LevelID, Levels.LevelShortName, ConceptsValues.EconomicZoneID, ConceptsValues.ClassificationID, ConceptsValues.IntegrationID, ConceptsValues.GroupGradeLevelID, GroupGradeLevels.GroupGradeLevelShortName, ConceptsValues.WorkingHours, ConceptsValues.AntiquityID, ConceptsValues.Antiquity2ID, Positions.PositionShortName, Positions.PositionName, PositionTypes.PositionTypeID, PositionTypes.PositionTypeShortName, PositionTypes.PositionTypeName, ConceptsValues.CompanyID, Companies.CompanyName From ConceptsValues, Positions, PositionTypes, GroupGradeLevels, Levels, Companies Where (ConceptsValues.PositionID=Positions.PositionID) And (Positions.PositionTypeID=PositionTypes.PositionTypeID) And (ConceptsValues.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (ConceptsValues.LevelID=Levels.LevelID) And (ConceptsValues.CompanyID=Companies.CompanyID)" & sCondition & " Order by PositionID, StartDate, LevelID, ClassificationID, IntegrationID, GroupGradeLevelID, WorkingHours, ConceptsValues.PositionTypeID, EconomicZoneID", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If iStatusID = 0 Then Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""sQuery"" ID=""sQueryHdn"" VALUE=""" & "Select ConceptsValues.RecordID, ConceptsValues.ConceptID, ConceptsValues.ConceptAmount, ConceptsValues.StartDate, ConceptsValues.EndDate, ConceptsValues.StatusID, Positions.PositionID, ConceptsValues.LevelID, Levels.LevelShortName, ConceptsValues.EconomicZoneID, ConceptsValues.ClassificationID, ConceptsValues.IntegrationID, ConceptsValues.GroupGradeLevelID, GroupGradeLevels.GroupGradeLevelShortName, ConceptsValues.WorkingHours, ConceptsValues.AntiquityID, ConceptsValues.Antiquity2ID, Positions.PositionShortName, Positions.PositionName, PositionTypes.PositionTypeID, PositionTypes.PositionTypeShortName, PositionTypes.PositionTypeName, ConceptsValues.CompanyID, Companies.CompanyName From ConceptsValues, Positions, PositionTypes, GroupGradeLevels, Levels, Companies Where (ConceptsValues.PositionID=Positions.PositionID) And (Positions.PositionTypeID=PositionTypes.PositionTypeID) And (ConceptsValues.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (ConceptsValues.LevelID=Levels.LevelID) And (ConceptsValues.CompanyID=Companies.CompanyID)" & sCondition & " Order by PositionID, StartDate, LevelID, ClassificationID, IntegrationID, GroupGradeLevelID, WorkingHours, ConceptsValues.PositionTypeID, EconomicZoneID" & """ />"
	sErrorDescription = "No se pudieron obtener los tabuladores."
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			bFirst = False
			bActiveConcept = False
			bActiveConcept_01 = False
			bActiveConcept_03 = False
			bActiveConcept_12 = False
			bActiveConcept_35 = False
			bActiveConcept_36 = False
			bActiveConcept_48 = False
			bActiveConcept_B2 = False
			bActiveConcept_Z3 = False
			bActiveConcept_01_Z3 = False
			bActiveConcept_03_Z3 = False
			bActiveConcept_12_Z3 = False
			bActiveConcept_35_Z3 = False
			bActiveConcept_36_Z3 = False
			bActiveConcept_48_Z3 = False
			bActiveConcept_B2_Z3 = False
			sRecordIDs = ""

			Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><TABLE BORDER="""
			If bForExport Then
				Response.Write "1"
			Else
				Response.Write "0"
			End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			If (Len(oRequest("ConceptID").Item) > 0) And (InStr(1, ",1,3,14,38,39,49,89,", "," & CStr(oRequest("ConceptID").Item) & ",", vbBinaryCompare) = 0) Then
				sColumnsTitles = "Tipo puesto,Código,Nivel,Jornada,Denominación del puesto,Fecha Inicio vigencia,Importe"
				sCellWidths = ",,,,,,,"
				sCellAlignments = "CENTER,RIGHT,RIGHT,RIGHT,LEFT,RIGHT,RIGHT"
			Else
				Select Case iSelectedTab
					Case 0
						asColumnsTitles = Split("<SPAN COLS=""8"">&nbsp;,<SPAN COLS=""4"">Zona 2,<SPAN COLS=""4"">Zona 3", ",", -1, vbBinaryCompare)
						asCellWidths = Split(",,,,,,,,,,,,,", ",", -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
						Else
							If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
								lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
							Else
								lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
							End If
						End If
						sColumnsTitles = "Compañía,Tipo puesto,Código,Nivel,Jornada,Denominación del puesto,Fecha Inicio vigencia,Fecha Fin vigencia,Sueldo,Asignación médica,Gastos de actualización,Total,Sueldo,Asignación médica,Gastos de actualización,Total"
						sCellWidths = ",,,,,,,,,,,,,"
						sCellAlignments = "CENTER,RIGHT,RIGHT,RIGHT,LEFT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT"
					Case 1
						sColumnsTitles = "Compañía,Zona Económica,Código,Denominación del puesto,Grupo grado nivel salarial,Clasificación,Integración,Fecha de inicio vigencia,Fecha Fin vigencia,Sueldo base,Compensación garantizada,Sueldo integrado"
						sCellWidths = ",,,,,,,,,,,,,"
						sCellAlignments = "CENTER,CENTER,LEFT,CENTER,CENTER,CENTER,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT"
					Case 2,4
						If iSelectedTab = 2 Then
							asColumnsTitles = Split("<SPAN COLS=""7"">&nbsp;,<SPAN COLS=""3"">Zona 2,<SPAN COLS=""3"">Zona 3", ",", -1, vbBinaryCompare)
							asCellWidths = Split(",,,,,,,,,,", ",", -1, vbBinaryCompare)
						Else
							asColumnsTitles = Split("<SPAN COLS=""6"">&nbsp;,<SPAN COLS=""3"">Zona 2,<SPAN COLS=""3"">Zona 3", ",", -1, vbBinaryCompare)
							asCellWidths = Split(",,,,,,,,,", ",", -1, vbBinaryCompare)
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
						If iSelectedTab = 2 Then
							sColumnsTitles = "Compañía,Tipo de puesto,Código,Nivel,Denominación del puesto,Fecha Inicio vigencia,Fecha Fin vigencia,Sueldo,Compensación garantizada,Total,Sueldo,Compensación garantizada,Total"
							sCellWidths = ",,,,,,,,,,"
							sCellAlignments = "CENTER,CENTER,CENTER,,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT"
						Else
							sColumnsTitles = "Compañía,Código,Nivel,Denominación del puesto,Fecha Inicio vigencia,Fecha Fin vigencia,Sueldo,Compensación garantizada,Total,Sueldo,Compensación garantizada,Total"
							sCellWidths = ",,,,,,,,,"
							sCellAlignments = "CENTER,CENTER,,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT"
						End If
					Case 3
						sColumnsTitles = "Compañía,Código,Nivel,Denominación del puesto,Fecha Inicio vigencia,Fecha Fin vigencia,Sueldo base,Compensación garantizada,Total mensual bruto"
						sCellWidths = ",,,,,,,,,"
						sCellAlignments = "LEFT,CENTER,LEFT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT"
					Case 5
						asColumnsTitles = Split("<SPAN COLS=""6"">&nbsp;,<SPAN COLS=""3"">Zona 2,<SPAN COLS=""3"">Zona 3", ",", -1, vbBinaryCompare)
						asCellWidths = Split(",,,,,,,,,", ",", -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
						Else
							If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
								lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
							Else
								lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
							End If
						End If
						sColumnsTitles = "Compañía,Código,Denominación del puesto,Nivel,Fecha Inicio vigencia,Fecha Fin vigencia,Beca,Complemento de beca,Total,Beca,Complemento de beca,Total"
						sCellWidths = ",,,,,,,,,"
						sCellAlignments = ",,,,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT"
					Case 6
						sColumnsTitles = "Compañía,Código,&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Denominación del puesto&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;,Nivel,Zona económica,Fecha Inicio vigencia,Fecha Fin vigencia,Beca"
						sCellWidths = ",,,,,"
						sCellAlignments = "LEFT,LEFT,,CENTER,CENTER,RIGHT"
				End Select
			End If
			'If (Not bForExport) And (iStatusID=0) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Or (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
			If (Not bForExport) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Or (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
				If CInt(Response.Cookies("SIAP_SectionID")) = 3 Then
					sColumnsTitles = sColumnsTitles & ",Acciones"
					sCellWidths = sCellWidths & ",80"
					sCellAlignments = sCellAlignments & ",CENTER"
				End If
			End If
			asColumnsTitles = Split(sColumnsTitles, ",", -1, vbBinaryCompare)
			asCellWidths = Split(sCellWidths, ",", -1, vbBinaryCompare)
			asCellAlignments = Split(sCellAlignments, ",", -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
			Else
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If
			End If
			lCurrentPositionID = -2
			dTotal = 0
			sBoldBegin = "<B>"
			sBoldEnd = "</B>"
			sFontBegin = ""
			sFontEnd = ""
			Do While Not oRecordset.EOF
				bContinue = False
				If (bFirst) And ((lCurrentPositionID <> CLng(oRecordset.Fields("PositionID").Value)) Or _
					(CLng(sStartDate) <> CLng(oRecordset.Fields("StartDate").Value)) Or _
					(CLng(iLevelID) <> CLng(oRecordset.Fields("LevelID").Value)) Or _
					(CLng(iClassificationID) <> CLng(oRecordset.Fields("ClassificationID").Value)) Or _
					(CLng(iIntegrationID) <> CLng(oRecordset.Fields("IntegrationID").Value)) Or _
					(CLng(iGroupGradeLevelID) <> CLng(oRecordset.Fields("GroupGradeLevelID").Value)) Or _
					(CSng(sWorkingHours) <> CSng(oRecordset.Fields("WorkingHours").Value)) Or _
					(CLng(iPositionTypeID) <> CLng(oRecordset.Fields("PositionTypeID").Value)) Or _
					(CInt(iAntiquityID) <> CInt(oRecordset.Fields("AntiquityID").Value)) Or _
					(CInt(iAntiquityID2) <> CInt(oRecordset.Fields("Antiquity2ID").Value))) _
				Then
					aConceptComponent(N_ID_CONCEPT) = CInt(oRecordset.Fields("ConceptID").Value)
					Select Case aConceptComponent(N_ID_CONCEPT)
						Case 1, 3, 14, 38, 39, 49, 89
							Select Case iSelectedTab
								Case 0
									sBoldBegin = ""
									sBoldEnd = ""
									If (StrComp(CStr(lConcept_01_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_35_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_48_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_01_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_35_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_48_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) _
									Then
										sBoldBegin = "<B>"
										sBoldEnd = "</B>"
									End If
									sFontBegin = ""
									sFontEnd = ""
									If (Not bActiveConcept_01) And (Not bActiveConcept_35) And (Not bActiveConcept_48) And (Not bActiveConcept_01_Z3) And (Not bActiveConcept_35_Z3) And (Not bActiveConcept_48_Z3) Then
										sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
										sFontEnd = "</FONT>"
									End If
									If iCompanyID = -1 Then
										sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
									Else
										sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(sCompanyName) & sBoldEnd & sFontEnd
									End If
									sRowContents =  sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionTypeShortName) & sBoldEnd & sFontEnd
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
									If iLevelID = -1 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
									Else
										If bForExport Then
											sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & "=T(""" & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & """)" & sBoldEnd & sFontEnd
										Else
											sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & sBoldEnd & sFontEnd
										End If
									End If
									If CSng(sWorkingHours) = -1 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sWorkingHours) & " Hrs." & sBoldEnd & sFontEnd
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionName) & sBoldEnd & sFontEnd
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
									If CLng(sEndDate) = 30000000 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
									End If
									If bActiveConcept_01 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=0&RecordID=" & lConcept_01_RecordID & "&ConceptID=1&ChangeEndDate=1&StartDate=" & lConcept_01_StartDate & """"
											sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
											sFontEnd = "</FONT>"
										Else
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=0&RecordID=" & lConcept_01_RecordID & "&ConceptID=1&StartDate=" & lConcept_01_StartDate & """"
											sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
											sFontEnd = "</FONT>"
										End If
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_01)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_01)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
									End If
									If Not IsEmpty(lConcept_01_RecordID) Then sRecordIDs = sRecordIDs & lConcept_01_RecordID & ","
									If bActiveConcept_35 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=0&RecordID=" & lConcept_35_RecordID & "&ConceptID=38&ChangeEndDate=1&StartDate=" & lConcept_35_StartDate & """"
											sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
											sFontEnd = "</FONT>"
										Else
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=0&RecordID=" & lConcept_35_RecordID & "&ConceptID=38&StartDate=" & lConcept_35_StartDate & """"
											sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
											sFontEnd = "</FONT>"
										End If
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_35)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_35)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
									End If
									If Not IsEmpty(lConcept_35_RecordID) Then sRecordIDs = sRecordIDs & lConcept_35_RecordID & ","
									If bActiveConcept_48 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=0&RecordID=" & lConcept_48_RecordID & "&ConceptID=49&ChangeEndDate=1&StartDate=" & lConcept_48_StartDate & """"
											sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
											sFontEnd = "</FONT>"
										Else
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=0&RecordID=" & lConcept_48_RecordID & "&ConceptID=49&StartDate=" & lConcept_48_StartDate & """"
											sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
											sFontEnd = "</FONT>"
										End If
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_48)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_48)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
									End If
									If Not IsEmpty(lConcept_48_RecordID) Then sRecordIDs = sRecordIDs & lConcept_48_RecordID & ","
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(Round((CDbl(dConcept_01) + CDbl(dConcept_35) + CDbl(dConcept_48))*2,2), 2, True, False, True)
									If bActiveConcept_01_Z3 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=0&RecordID=" & lConcept_01_Z3_RecordID & "&ConceptID=1&ChangeEndDate=1&StartDate=" & lConcept_01_Z3_StartDate & """"
											sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
											sFontEnd = "</FONT>"
										Else
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=0&RecordID=" & lConcept_01_Z3_RecordID & "&ConceptID=1&StartDate=" & lConcept_01_Z3_StartDate & """"
											sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
											sFontEnd = "</FONT>"
										End If
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_01_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_01_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
									End If
									If Not IsEmpty(lConcept_01_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_01_Z3_RecordID & ","
									If bActiveConcept_35_Z3 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=0&RecordID=" & lConcept_35_Z3_RecordID & "&ConceptID=38&ChangeEndDate=1&StartDate=" & lConcept_35_Z3_StartDate & """"
											sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
											sFontEnd = "</FONT>"
										Else
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=0&RecordID=" & lConcept_35_Z3_RecordID & "&ConceptID=38&StartDate=" & lConcept_35_Z3_StartDate & """"
											sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
											sFontEnd = "</FONT>"
										End If
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_35_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_35_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
									End If
									If Not IsEmpty(lConcept_35_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_35_Z3_RecordID & ","
									If bActiveConcept_48_Z3 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=0&RecordID=" & lConcept_48_Z3_RecordID & "&ConceptID=49&ChangeEndDate=1&StartDate=" & lConcept_48_Z3_StartDate & """"
											sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
											sFontEnd = "</FONT>"
										Else
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=0&RecordID=" & lConcept_48_Z3_RecordID & "&ConceptID=49&StartDate=" & lConcept_48_Z3_StartDate & """"
											sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
											sFontEnd = "</FONT>"
										End If
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_48_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_48_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
									End If
									If Not IsEmpty(lConcept_48_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_48_Z3_RecordID & ","
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(Round((CDbl(dConcept_01_Z3) + CDbl(dConcept_35_Z3) + CDbl(dConcept_48_Z3))*2,2), 2, True, False, True)
									If (Not bActiveConcept_01) And (Not bActiveConcept_35) And (Not bActiveConcept_48) And (Not bActiveConcept_01_Z3) And (Not bActiveConcept_35_Z3) And (Not bActiveConcept_48_Z3) Then
										If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
											If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
											sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
											If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
												sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
													sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
												sRowContents = sRowContents & "</A>&nbsp;"
											End If
										End If
									ElseIf (bActiveConcept_01) And (bActiveConcept_35) And (bActiveConcept_48) And (bActiveConcept_01_Z3) And (bActiveConcept_35_Z3) And (bActiveConcept_48_Z3) Then
										If (Not bForExport) And (iStatusID=1) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
											sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&ChangeEndDate=1"">"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar fecha de fin de vigencia"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
										End If
									End If
									asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If (dConcept_01 + dConcept_35 + dConcept_48 + lConcept_01_Z3_RecordID + dConcept_35_Z3 + dConcept_48_Z3) > 0 Then
										If bForExport Then
											lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
										Else
											lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
										End If
									End If
								Case 1
									sBoldBegin = ""
									sBoldEnd = ""
									If (StrComp(CStr(lConcept_01_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_03_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) _
									Then
										sBoldBegin = "<B>"
										sBoldEnd = "</B>"
									End If
									sFontBegin = ""
									sFontEnd = ""
									If (Not bActiveConcept_01) And (Not bActiveConcept_03) Then
										sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
										sFontEnd = "</FONT>"
									End If
									If iCompanyID = -1 Then
										sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
									Else
										sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(sCompanyName) & sBoldEnd & sFontEnd
									End If
									If iEconomicZoneID = 0 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(iEconomicZoneID) & sBoldEnd & sFontEnd
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionName) & sBoldEnd & sFontEnd
									If iGroupGradeLevelID = -1 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sGroupGradeLevelShortName) & sBoldEnd & sFontEnd
									End If
									If iClassificationID = -1 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(iClassificationID) & sBoldEnd & sFontEnd
									End If
									If iIntegrationID = -1 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(iIntegrationID) & sBoldEnd & sFontEnd
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
									If CLng(sEndDate) = 30000000 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
									End If
									If bActiveConcept_01 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=1&RecordID=" & lConcept_01_RecordID & "&ConceptID=1&ChangeEndDate=1&StartDate=" & lConcept_01_StartDate & """"
											sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
											sFontEnd = "</FONT>"
										Else
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=1&RecordID=" & lConcept_01_RecordID & "&ConceptID=1&StartDate=" & lConcept_01_StartDate & """"
											sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
											sFontEnd = "</FONT>"
										End If
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_01)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_01)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
									End If
									If Not IsEmpty(lConcept_01_RecordID) Then sRecordIDs = sRecordIDs & lConcept_01_RecordID & ","
									If bActiveConcept_03 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=1&RecordID=" & lConcept_03_RecordID & "&ConceptID=3&ChangeEndDate=1&StartDate=" & lConcept_03_StartDate & """"
											sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
											sFontEnd = "</FONT>"
										Else
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=1&RecordID=" & lConcept_03_RecordID & "&ConceptID=3&StartDate=" & lConcept_03_StartDate & """"
											sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
											sFontEnd = "</FONT>"
										End If
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_03)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_03)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
									End If
									If Not IsEmpty(lConcept_03_RecordID) Then sRecordIDs = sRecordIDs & lConcept_03_RecordID & ","
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(Round((CDbl(dConcept_01) + CDbl(dConcept_03))*2,2), 2, True, False, True)
									If (Not bActiveConcept_01) And (Not bActiveConcept_03) Then
										If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
											If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
											sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
											If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
												sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
													sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
												sRowContents = sRowContents & "</A>&nbsp;"
											End If
										End If
									End If
									asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If (dConcept_01 + dConcept_03) > 0 Then
										If bForExport Then
											lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
										Else
											lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
										End If
									End If
								Case 2,4
									sBoldBegin = ""
									sBoldEnd = ""
									If (StrComp(CStr(lConcept_01_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_03_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_01_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_03_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) _
									Then
										sBoldBegin = "<B>"
										sBoldEnd = "</B>"
									End If
									sFontBegin = ""
									sFontEnd = ""
									If (Not bActiveConcept_01) And (Not bActiveConcept_03) And (Not bActiveConcept_01_Z3) And (Not bActiveConcept_03_Z3) Then
										sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
										sFontEnd = "</FONT>"
									End If
									If iCompanyID = -1 Then
										sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
									Else
										sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(sCompanyName) & sBoldEnd & sFontEnd
									End If
									If iSelectedTab = 2 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionTypeShortName) & sBoldEnd & sFontEnd
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
									End If
									If iLevelID = -1 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
									Else
										If bForExport Then
											sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & "=T(""" & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & """)" & sBoldEnd & sFontEnd
										Else
											sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & sBoldEnd & sFontEnd
										End If
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionName) & sBoldEnd & sFontEnd
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
									If CLng(sEndDate) = 30000000 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
									End If
									If bActiveConcept_01 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_01_RecordID & "&ConceptID=1&ChangeEndDate=1&StartDate=" & lConcept_01_StartDate & """"
											sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
											sFontEnd = "</FONT>"
										Else
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_01_RecordID & "&ConceptID=1&StartDate=" & lConcept_01_StartDate & """"
											sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
											sFontEnd = "</FONT>"
										End If
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_01)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_01)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
									End If
									If Not IsEmpty(lConcept_01_RecordID) Then sRecordIDs = sRecordIDs & lConcept_01_RecordID & ","
									If bActiveConcept_03 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_03_RecordID & "&ConceptID=3&ChangeEndDate=1&StartDate=" & lConcept_03_StartDate & """"
											sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
											sFontEnd = "</FONT>"
										Else
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_03_RecordID & "&ConceptID=3&StartDate=" & lConcept_03_StartDate & """"
											sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
											sFontEnd = "</FONT>"
										End If
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_03)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_03)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
									End If
									If Not IsEmpty(lConcept_03_RecordID) Then sRecordIDs = sRecordIDs & lConcept_03_RecordID & ","
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(Round((CDbl(dConcept_01) + CDbl(dConcept_03))*2,2), 2, True, False, True)
									If bActiveConcept_01_Z3 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_01_Z3_RecordID & "&ConceptID=1&ChangeEndDate=1&StartDate=" & lConcept_01_Z3_StartDate & """"
											sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
											sFontEnd = "</FONT>"
										Else
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_01_Z3_RecordID & "&ConceptID=1&StartDate=" & lConcept_01_Z3_StartDate & """"
											sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
											sFontEnd = "</FONT>"
										End If
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_01_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_01_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
									End If
									If Not IsEmpty(lConcept_01_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_01_Z3_RecordID & ","
									If bActiveConcept_03_Z3 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_03_Z3_RecordID & "&ConceptID=3&ChangeEndDate=1&StartDate=" & lConcept_03_Z3_StartDate & """"
											sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
											sFontEnd = "</FONT>"
										Else
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_03_Z3_RecordID & "&ConceptID=3&StartDate=" & lConcept_03_Z3_StartDate & """"
											sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
											sFontEnd = "</FONT>"
										End If
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_03_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_03_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
									End If
									If Not IsEmpty(lConcept_03_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_03_Z3_RecordID & ","
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(Round((CDbl(dConcept_01_Z3) + CDbl(dConcept_03_Z3))*2,2), 2, True, False, True)
									If (Not bActiveConcept_01) And (Not bActiveConcept_03) And (Not bActiveConcept_01_Z3) And (Not bActiveConcept_03_Z3) Then
										If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
											If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
											sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
											If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
												sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
													sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
												sRowContents = sRowContents & "</A>&nbsp;"
											End If
										End If
									End If
									asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If (dConcept_01 + dConcept_03 + dConcept_01_Z3 + dConcept_03_Z3) > 0 Then
										If bForExport Then
											lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
										Else
											lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
										End If
									End If
								Case 3
									sBoldBegin = ""
									sBoldEnd = ""
									If (StrComp(CStr(lConcept_01_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_03_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) _
									Then
										sBoldBegin = "<B>"
										sBoldEnd = "</B>"
									End If
									sFontBegin = ""
									sFontEnd = ""
									If (Not bActiveConcept_01) And (Not bActiveConcept_03) Then
										sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
										sFontEnd = "</FONT>"
									End If
									If iCompanyID = -1 Then
										sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
									Else
										sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(sCompanyName) & sBoldEnd & sFontEnd
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
									If iLevelID = -1 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
									Else
										If bForExport Then
											sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & "=T(""" & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & """)" & sBoldEnd & sFontEnd
										Else
											sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & sBoldEnd & sFontEnd
										End If
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionName) & sBoldEnd & sFontEnd
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
									If CLng(sEndDate) = 30000000 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
									End If
									If bActiveConcept_01 Then
										sFontBegin = ""
										sFontEnd = ""
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=3&RecordID=" & lConcept_01_RecordID & "&ConceptID=1&ChangeEndDate=1&StartDate=" & lConcept_01_StartDate & """"
											sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
											sFontEnd = "</FONT>"
										Else
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=3&RecordID=" & lConcept_01_RecordID & "&ConceptID=1&StartDate=" & lConcept_01_StartDate & """"
											sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
											sFontEnd = "</FONT>"
										End If
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_01)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_01)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
									End If
									If Not IsEmpty(lConcept_01_RecordID) Then sRecordIDs = sRecordIDs & lConcept_01_RecordID & ","
									If bActiveConcept_03 Then
										sFontBegin = ""
										sFontEnd = ""
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=3&RecordID=" & lConcept_03_RecordID & "&ConceptID=3&ChangeEndDate=1&StartDate=" & lConcept_03_StartDate & """"
											sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
											sFontEnd = "</FONT>"
										Else
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=3&RecordID=" & lConcept_03_RecordID & "&ConceptID=3&StartDate=" & lConcept_03_StartDate & """"
											sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
											sFontEnd = "</FONT>"
										End If
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_03)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_03)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
									End If
									If Not IsEmpty(lConcept_03_RecordID) Then sRecordIDs = sRecordIDs & lConcept_03_RecordID & ","
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(Round((CDbl(dConcept_01) + CDbl(dConcept_03))*2,2), 2, True, False, True)
									If (Not bActiveConcept_01) And (Not bActiveConcept_03) Then
										If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
											If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
											sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
											If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
												sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
													sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
												sRowContents = sRowContents & "</A>&nbsp;"
											End If
										End If
									End If
									asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If (dConcept_01 + dConcept_03) > 0 Then
										If bForExport Then
											lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
										Else
											lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
										End If
									End If
								Case 5
									sBoldBegin = ""
									sBoldEnd = ""
									If (StrComp(CStr(lConcept_B2_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_36_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_B2_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
										(StrComp(CStr(lConcept_36_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) _
									Then
										sBoldBegin = "<B>"
										sBoldEnd = "</B>"
									End If
									sFontBegin = ""
									sFontEnd = ""
									If (Not bActiveConcept_B2) And (Not bActiveConcept_36) And (Not bActiveConcept_B2_Z3) And (Not bActiveConcept_36_Z3) Then
										sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
										sFontEnd = "</FONT>"
									End If
									If iCompanyID = -1 Then
										sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
									Else
										sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(sCompanyName) & sBoldEnd & sFontEnd
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionName) & sBoldEnd & sFontEnd
									If iLevelID = -1 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
									Else
										If bForExport Then
											sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & "=T(""" & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & """)" & sBoldEnd & sFontEnd
										Else
											sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & sBoldEnd & sFontEnd
										End If
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
									If CLng(sEndDate) = 30000000 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
									End If
									If bActiveConcept_B2 Then
										sFontBegin = ""
										sFontEnd = ""
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=5&RecordID=" & lConcept_B2_RecordID & "&ConceptID=89&ChangeEndDate=1&StartDate=" & lConcept_B2_StartDate & """"
											sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
											sFontEnd = "</FONT>"
										Else
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=5&RecordID=" & lConcept_B2_RecordID & "&ConceptID=89&StartDate=" & lConcept_B2_StartDate & """"
											sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
											sFontEnd = "</FONT>"
										End If
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_B2)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_B2)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
									End If
									If Not IsEmpty(lConcept_B2_RecordID) Then sRecordIDs = sRecordIDs & lConcept_B2_RecordID & ","
									If bActiveConcept_36 Then
										sFontBegin = ""
										sFontEnd = ""
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=5&RecordID=" & lConcept_36_RecordID & "&ConceptID=39&ChangeEndDate=1&StartDate=" & lConcept_36_StartDate & """"
											sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
											sFontEnd = "</FONT>"
										Else
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=5&RecordID=" & lConcept_36_RecordID & "&ConceptID=39&StartDate=" & lConcept_36_StartDate & """"
											sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
											sFontEnd = "</FONT>"
										End If
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_36)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_36)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
									End If
									If Not IsEmpty(lConcept_36_RecordID) Then sRecordIDs = sRecordIDs & lConcept_36_RecordID & ","
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(Round((CDbl(dConcept_B2) + CDbl(dConcept_36))*2,2), 2, True, False, True)
									If bActiveConcept_B2_Z3 Then
										sFontBegin = ""
										sFontEnd = ""
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=5&RecordID=" & lConcept_B2_Z3_RecordID & "&ConceptID=89&ChangeEndDate=1&StartDate=" & lConcept_B2_StartDate & """"
											sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
											sFontEnd = "</FONT>"
										Else
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=5&RecordID=" & lConcept_B2_Z3_RecordID & "&ConceptID=89&StartDate=" & lConcept_B2_StartDate & """"
											sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
											sFontEnd = "</FONT>"
										End If
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_B2_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_B2_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
									End If
									If Not IsEmpty(lConcept_B2_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_B2_Z3_RecordID & ","
									If bActiveConcept_36_Z3 Then
										sFontBegin = ""
										sFontEnd = ""
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=5&RecordID=" & lConcept_36_Z3_RecordID & "&ConceptID=39&ChangeEndDate=1&StartDate=" & lConcept_36_StartDate & """"
											sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
											sFontEnd = "</FONT>"
										Else
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=5&RecordID=" & lConcept_36_Z3_RecordID & "&ConceptID=39&StartDate=" & lConcept_36_StartDate & """"
											sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
											sFontEnd = "</FONT>"
										End If
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_36_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_36_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
									End If
									If Not IsEmpty(lConcept_36_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_36_Z3_RecordID & ","
									If InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
									'sRecordIDs=""
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(Round((CDbl(dConcept_B2_Z3) + CDbl(dConcept_36_Z3))*2,2), 2, True, False, True)
									If (Not bActiveConcept_B2) And (Not bActiveConcept_36) And (Not bActiveConcept_B2_Z3) And (Not bActiveConcept_36_Z3) Then
										If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
											sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
											If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
												sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
													sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
												sRowContents = sRowContents & "</A>&nbsp;"
											End If
										End If
									ElseIf (bActiveConcept_B2) And (bActiveConcept_36) And (bActiveConcept_B2_Z3) And (bActiveConcept_36_Z3) Then
										If (Not bForExport) And (iStatusID=1) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
											sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&ChangeEndDate=1"">"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar fecha de fin de vigencia"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
										End If
									End If
									asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If (dConcept_B2 + dConcept_36) > 0 Then
										If bForExport Then
											lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
										Else
											lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
										End If
									End If
								Case 6
									sBoldBegin = ""
									sBoldEnd = ""
									If (StrComp(CStr(lConcept_12_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Then
										sBoldBegin = "<B>"
										sBoldEnd = "</B>"
									End If
									sFontBegin = ""
									sFontEnd = ""
									If (Not bActiveConcept_12) Then
										sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
										sFontEnd = "</FONT>"
									End If
									If iCompanyID = -1 Then
										sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
									Else
										sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(sCompanyName) & sBoldEnd & sFontEnd
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionName) & sBoldEnd & sFontEnd
									If iLevelID = -1 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
									Else
										If bForExport Then
											sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & "=T(""" & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & """)" & sBoldEnd & sFontEnd
										Else
											sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & sBoldEnd & sFontEnd
										End If
									End If
									If iEconomicZoneID = 0 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(iEconomicZoneID) & sBoldEnd & sFontEnd
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
									If CLng(sEndDate) = 30000000 Then
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
									End If
									If bActiveConcept_12 Then
										sFontBegin = ""
										sFontEnd = ""
										sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
										If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=6&RecordID=" & lConcept_12_RecordID & "&ConceptID=14&ChangeEndDate=1&StartDate=" & lConcept_12_StartDate & """"
											sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
											sFontEnd = "</FONT>"
										Else
											sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=6&RecordID=" & lConcept_12_RecordID & "&ConceptID=14&StartDate=" & lConcept_12_StartDate & """"
											sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
											sFontEnd = "</FONT>"
										End If
										sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_12)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
									Else
										sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
										sFontEnd = "</FONT>"
										sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_12)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
									End If
									If Not IsEmpty(lConcept_12_RecordID) Then sRecordIDs = sRecordIDs & lConcept_12_RecordID & ","
									If Not bActiveConcept_12 Then
										If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
											If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
											sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
											If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
												sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
													sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
												sRowContents = sRowContents & "</A>&nbsp;"
											End If
										End If
									End If
									asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If dConcept_12 > 0 Then
										If bForExport Then
											lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
										Else
											lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
										End If
									End If
							End Select
						Case Else
							sBoldBegin = ""
							sBoldEnd = ""
							If (StrComp(CStr(lConcept_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Then
								sBoldBegin = "<B>"
								sBoldEnd = "</B>"
							End If
							sFontBegin = ""
							sFontEnd = ""
							If (Not bActiveConcept) Then
								sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
								sFontEnd = "</FONT>"
							End If
							If iCompanyID = -1 Then
								sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
							Else
								sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(sCompanyName) & sBoldEnd & sFontEnd
							End If
							If CInt(oRecordset.Fields("PositionTypeID").Value) = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PositionTypeShortName").Value)) & sBoldEnd & sFontEnd
							End If
							If CInt(oRecordset.Fields("PositionID").Value) = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(oRecordset.Fields("PositionShortName").Value) & sBoldEnd & sFontEnd
							End If
							If CInt(oRecordset.Fields("LevelID").Value) = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(oRecordset.Fields("LevelShortName").Value) & sBoldEnd & sFontEnd
							End If
							If CSng(oRecordset.Fields("WorkingHours").Value) = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("WorkingHours").Value)) & " Hrs." & sBoldEnd & sFontEnd
							End If
							If CInt(oRecordset.Fields("PositionID").Value) = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(oRecordset.Fields("PositionName").Value) & sBoldEnd & sFontEnd
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
							If CLng(sEndDate) = 30000000 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
							End If
							If bActiveConcept Then
								sFontBegin = ""
								sFontEnd = ""
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_RecordID & "&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&StartDate=" & lConcept_StartDate & """"
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_RecordID) Then sRecordIDs = sRecordIDs & lConcept_RecordID & ","
							End If
							If Not bActiveConcept Then
								If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
									If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
									sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
									If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
										sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
											sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
										sRowContents = sRowContents & "</A>&nbsp;"
									End If
								End If
							End If
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If dConcept > 0 Then
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
							End If
					End Select
					dConcept = 0
					dConcept_01 = 0
					dConcept_03 = 0
					dConcept_12 = 0
					dConcept_35 = 0
					dConcept_36 = 0
					dConcept_48 = 0
					dConcept_B2 = 0
					dConcept_Z3 = 0
					dConcept_01_Z3 = 0
					dConcept_03_Z3 = 0
					dConcept_12_Z3 = 0
					dConcept_35_Z3 = 0
					dConcept_36_Z3 = 0
					dConcept_48_Z3 = 0
					dConcept_B2_Z3 = 0
					lConcept_RecordID=0
					lConcept_01_RecordID=0
					lConcept_03_RecordID=0
					lConcept_12_RecordID=0
					lConcept_35_RecordID=0
					lConcept_36_RecordID=0
					lConcept_48_RecordID=0
					lConcept_B2_RecordID=0
					lConcept_Z3_RecordID=0
					lConcept_01_Z3_RecordID=0
					lConcept_03_Z3_RecordID=0
					lConcept_12_Z3_RecordID=0
					lConcept_35_Z3_RecordID=0
					lConcept_36_Z3_RecordID=0
					lConcept_48_Z3_RecordID=0
					lConcept_B2_Z3_RecordID=0
					bActiveConcept = False
					bActiveConcept_01 = False
					bActiveConcept_03 = False
					bActiveConcept_12 = False
					bActiveConcept_35 = False
					bActiveConcept_36 = False
					bActiveConcept_48 = False
					bActiveConcept_B2 = False
					bActiveConcept_Z3 = False
					bActiveConcept_01_Z3 = False
					bActiveConcept_03_Z3 = False
					bActiveConcept_12_Z3 = False
					bActiveConcept_35_Z3 = False
					bActiveConcept_36_Z3 = False
					bActiveConcept_48_Z3 = False
					bActiveConcept_B2_Z3 = False
					sRecordIDs = ""
				End If
				bFirst = True
				lCurrentPositionID = CLng(oRecordset.Fields("PositionID").Value)
				If CInt(oRecordset.Fields("EconomicZoneID").Value) = 3 Then
					Select Case CLng(oRecordset.Fields("ConceptID").Value)
						Case 1
							If Not bContinue Then
								If dConcept_01_Z3 > 0 Then
									bContinue = False
								Else
									dConcept_01_Z3 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_01_Z3_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_01_Z3_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_01_Z3 = True
									Else
										bActiveConcept_01_Z3 = False
									End If
								End If
							End If
						Case 3
							If Not bContinue Then
								If dConcept_03_Z3 > 0 Then
									bContinue = False
								Else
									dConcept_03_Z3 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_03_Z3_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_03_Z3_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_03_Z3 = True
									Else
										bActiveConcept_03_Z3 = False
									End If
								End If
							End If
						Case 14
							If Not bContinue Then
								If dConcept_12_Z3 > 0 Then
									bContinue = False
								Else
									dConcept_12_Z3 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_12_Z3_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_12_Z3_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_12_Z3 = True
									Else
										bActiveConcept_12_Z3 = False
									End If
								End If
							End If
						Case 38
							If Not bContinue Then
								If dConcept_35_Z3 > 0 Then
									bContinue = False
								Else
									dConcept_35_Z3 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_35_Z3_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_35_Z3_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_35_Z3 = True
									Else
										bActiveConcept_35_Z3 = False
									End If
								End If
							End If
						Case 39
							If Not bContinue Then
								If dConcept_36_Z3 > 0 Then
									bContinue = False
								Else
									dConcept_36_Z3 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_36_Z3_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_36_Z3_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_36_Z3 = True
									Else
										bActiveConcept_36_Z3 = False
									End If
								End If
							End If
						Case 49
							If Not bContinue Then
								If dConcept_48_Z3 > 0 Then
									bContinue = False
								Else
									dConcept_48_Z3 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_48_Z3_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_48_Z3_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_48_Z3 = True
									Else
										bActiveConcept_48_Z3 = False
									End If
								End If
							End If
						Case 89
							If Not bContinue Then
								If dConcept_B2_Z3 > 0 Then
									bContinue = False
								Else
									dConcept_B2_Z3 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_B2_Z3_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_B2_Z3_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_B2_Z3 = True
									Else
										bActiveConcept_B2_Z3 = False
									End If
								End If
							End If
						Case Else
							If Not bContinue Then
								If dConcept_Z3 > 0 Then
									bContinue = False
								Else
									dConcept_Z3 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_Z3_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_Z3_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_Z3 = True
									Else
										bActiveConcept_Z3 = False
									End If
								End If
							End If
					End Select
				Else
					Select Case CLng(oRecordset.Fields("ConceptID").Value)
						Case 1
							If Not bContinue Then
								If dConcept_01 > 0 Then
									bContinue = False
								Else
									dConcept_01 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_01_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_01_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_01 = True
									Else
										bActiveConcept_01 = False
									End If
								End If
							End If
						Case 3
							If Not bContinue Then
								If dConcept_03 > 0 Then
									bContinue = False
								Else
									dConcept_03 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_03_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_03_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_03 = True
									Else
										bActiveConcept_03 = False
									End If
								End If
							End If
						Case 14
							If Not bContinue Then
								If dConcept_12 > 0 Then
									bContinue = False
								Else
									dConcept_12 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_12_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_12_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_12 = True
									Else
										lConcept_12_RecordID = CLng(oRecordset.Fields("RecordID").Value)
										bActiveConcept_12 = False
									End If
								End If
							End If
						Case 38
							If Not bContinue Then
								If dConcept_35 > 0 Then
									bContinue = False
								Else
									dConcept_35 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_35_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_35_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_35 = True
									Else
										bActiveConcept_35 = False
									End If
								End If
							End If
						Case 39
							If Not bContinue Then
								If dConcept_36 > 0 Then
									bContinue = False
								Else
									dConcept_36 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_36_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_36_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_36 = True
									Else
										bActiveConcept_36 = False
									End If
								End If
							End If
						Case 49
							If Not bContinue Then
								If dConcept_48 > 0 Then
									bContinue = False
								Else
									dConcept_48 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_48_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_48_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_48 = True
									Else
										bActiveConcept_48 = False
									End If
								End If
							End If
						Case 89
							If Not bContinue Then
								If dConcept_B2 > 0 Then
									bContinue = False
								Else
									dConcept_B2 = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_B2_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_B2_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept_B2 = True
									Else
										bActiveConcept_B2 = False
									End If
								End If
							End If
						Case Else
							If Not bContinue Then
								If dConcept > 0 Then
									bContinue = False
								Else
									dConcept = CDbl(oRecordset.Fields("ConceptAmount").Value)
									lConcept_RecordID = CLng(oRecordset.Fields("RecordID").Value)
									If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And CInt(oRecordset.Fields("StatusID").Value) = 1 Then
										lConcept_StartDate = CLng(oRecordset.Fields("StartDate").Value)
										bActiveConcept = True
									Else
										bActiveConcept = False
									End If
								End If
							End If
					End Select
				End If
				aConceptComponent(N_ID_CONCEPT) = CInt(oRecordset.Fields("ConceptID").Value)
				iLevelID = CInt(oRecordset.Fields("LevelID").Value)
				iAntiquityID = CInt(oRecordset.Fields("AntiquityID").Value)
				iAntiquityID2 = CInt(oRecordset.Fields("Antiquity2ID").Value)
				iEconomicZoneID = CInt(oRecordset.Fields("EconomicZoneID").Value)
				iGroupGradeLevelID = CInt(oRecordset.Fields("GroupGradeLevelID").Value)
				iClassificationID = CInt(oRecordset.Fields("ClassificationID").Value)
				iIntegrationID = CInt(oRecordset.Fields("IntegrationID").Value)
				iPositionTypeID = CInt(oRecordset.Fields("PositionTypeID").Value)
				sPositionTypeShortName = CStr(oRecordset.Fields("PositionTypeShortName").Value)
				sPositionShortName = CStr(oRecordset.Fields("PositionShortName").Value)
				sLevelShortName = CStr(oRecordset.Fields("LevelShortName").Value)
				sWorkingHours = CStr(oRecordset.Fields("WorkingHours").Value)
				sPositionName = CStr(oRecordset.Fields("PositionName").Value)
				sStartDate = CStr(oRecordset.Fields("StartDate").Value)
				sEndDate = CStr(oRecordset.Fields("EndDate").Value)
				sGroupGradeLevelShortName = CStr(oRecordset.Fields("GroupGradeLevelShortName").Value)
				iCompanyID = CInt(oRecordset.Fields("CompanyID").Value)
				sCompanyName = CStr(oRecordset.Fields("CompanyName").Value)
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			Select Case aConceptComponent(N_ID_CONCEPT)
				Case 1, 3, 14, 38, 39, 49, 89
					Select Case iSelectedTab
						Case 0
							sBoldBegin = ""
							sBoldEnd = ""
							If (StrComp(CStr(lConcept_01_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_35_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_48_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_01_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_35_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_48_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) _
							Then
								sBoldBegin = "<B>"
								sBoldEnd = "</B>"
							End If
							sFontBegin = ""
							sFontEnd = ""
							If (Not bActiveConcept_01) And (Not bActiveConcept_35) And (Not bActiveConcept_48) And (Not bActiveConcept_01_Z3) And (Not bActiveConcept_35_Z3) And (Not bActiveConcept_48_Z3) Then
								sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
								sFontEnd = "</FONT>"
							End If
							If iCompanyID = -1 Then
								sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
							Else
								sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(sCompanyName) & sBoldEnd & sFontEnd
							End If
							sRowContents =  sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionTypeShortName) & sBoldEnd & sFontEnd
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
							If iLevelID = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
							Else
								If bForExport Then
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & "=T(""" & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & """)" & sBoldEnd & sFontEnd
								Else
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & sBoldEnd & sFontEnd
								End If
							End If
							If CSng(sWorkingHours) = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sWorkingHours) & " Hrs." & sBoldEnd & sFontEnd
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionName) & sBoldEnd & sFontEnd
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
							If CLng(sEndDate) = 30000000 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
							End If
							If bActiveConcept_01 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=0&RecordID=" & lConcept_01_RecordID & "&ConceptID=1&ChangeEndDate=1&StartDate=" & lConcept_01_StartDate & """"
									sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
									sFontEnd = "</FONT>"
								Else
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=0&RecordID=" & lConcept_01_RecordID & "&ConceptID=1&StartDate=" & lConcept_01_StartDate & """"
								End If
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_01)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_01)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
							End If
							If Not IsEmpty(lConcept_01_RecordID) Then sRecordIDs = sRecordIDs & lConcept_01_RecordID & ","
							If bActiveConcept_35 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=0&RecordID=" & lConcept_35_RecordID & "&ConceptID=38&ChangeEndDate=1&StartDate=" & lConcept_35_StartDate & """"
									sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
									sFontEnd = "</FONT>"
								Else
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=0&RecordID=" & lConcept_35_RecordID & "&ConceptID=38&StartDate=" & lConcept_35_StartDate & """"
									sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
									sFontEnd = "</FONT>"
								End If
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_35)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_35)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
							End If
							If Not IsEmpty(lConcept_35_RecordID) Then sRecordIDs = sRecordIDs & lConcept_35_RecordID & ","
							If bActiveConcept_48 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=0&RecordID=" & lConcept_48_RecordID & "&ConceptID=49&ChangeEndDate=1&StartDate=" & lConcept_48_StartDate & """"
									sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
									sFontEnd = "</FONT>"
								Else
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=0&RecordID=" & lConcept_48_RecordID & "&ConceptID=49&StartDate=" & lConcept_48_StartDate & """"
									sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
									sFontEnd = "</FONT>"
								End If
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_48)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_48)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
							End If
							If Not IsEmpty(lConcept_48_RecordID) Then sRecordIDs = sRecordIDs & lConcept_48_RecordID & ","
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(Round((CDbl(dConcept_01) + CDbl(dConcept_35) + CDbl(dConcept_48))*2,2), 2, True, False, True)
							If bActiveConcept_01_Z3 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=0&RecordID=" & lConcept_01_Z3_RecordID & "&ConceptID=1&ChangeEndDate=1&StartDate=" & lConcept_01_Z3_StartDate & """"
									sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
									sFontEnd = "</FONT>"
								Else
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=0&RecordID=" & lConcept_01_Z3_RecordID & "&ConceptID=1&StartDate=" & lConcept_01_Z3_StartDate & """"
									sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
									sFontEnd = "</FONT>"
								End If
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_01_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_01_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
							End If
							If Not IsEmpty(lConcept_01_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_01_Z3_RecordID & ","
							If bActiveConcept_35_Z3 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=0&RecordID=" & lConcept_35_Z3_RecordID & "&ConceptID=38&ChangeEndDate=1&StartDate=" & lConcept_35_Z3_StartDate & """"
									sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
									sFontEnd = "</FONT>"
								Else
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=0&RecordID=" & lConcept_35_Z3_RecordID & "&ConceptID=38&StartDate=" & lConcept_35_Z3_StartDate & """"
									sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
									sFontEnd = "</FONT>"
								End If
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_35_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_35_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
							End If
							If Not IsEmpty(lConcept_35_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_35_Z3_RecordID & ","
							If bActiveConcept_48_Z3 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=0&RecordID=" & lConcept_48_Z3_RecordID & "&ConceptID=49&ChangeEndDate=1&StartDate=" & lConcept_48_Z3_StartDate & """"
									sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
									sFontEnd = "</FONT>"
								Else
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=0&RecordID=" & lConcept_48_Z3_RecordID & "&ConceptID=49&StartDate=" & lConcept_48_Z3_StartDate & """"
									sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
									sFontEnd = "</FONT>"
								End If
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_48_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_48_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
							End If
							If Not IsEmpty(lConcept_48_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_48_Z3_RecordID & ","
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(Round((CDbl(dConcept_01_Z3) + CDbl(dConcept_35_Z3) + CDbl(dConcept_48_Z3))*2,2), 2, True, False, True)
							If (Not bActiveConcept_01) And (Not bActiveConcept_35) And (Not bActiveConcept_48) And (Not bActiveConcept_01_Z3) And (Not bActiveConcept_35_Z3) And (Not bActiveConcept_48_Z3) Then
								If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
									If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
									sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
									If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
										sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
											sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
										sRowContents = sRowContents & "</A>&nbsp;"
									End If
								End If
							ElseIf (bActiveConcept_01) And (bActiveConcept_35) And (bActiveConcept_48) And (bActiveConcept_01_Z3) And (bActiveConcept_35_Z3) And (bActiveConcept_48_Z3) Then
								If (Not bForExport) And (iStatusID=1) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
									sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&ChangeEndDate=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar fecha de fin de vigencia"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
								End If
							End If
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If (dConcept_01 + dConcept_35 + dConcept_48 + lConcept_01_Z3_RecordID + dConcept_35_Z3 + dConcept_48_Z3) > 0 Then
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
							End If
						Case 1
							sBoldBegin = ""
							sBoldEnd = ""
							If (StrComp(CStr(lConcept_01_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_03_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) _
							Then
								sBoldBegin = "<B>"
								sBoldEnd = "</B>"
							End If
							sFontBegin = ""
							sFontEnd = ""
							If (Not bActiveConcept_01) And (Not bActiveConcept_03) Then
								sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
								sFontEnd = "</FONT>"
							End If
							If iCompanyID = -1 Then
								sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
							Else
								sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(sCompanyName) & sBoldEnd & sFontEnd
							End If
							If iEconomicZoneID = 0 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(iEconomicZoneID) & sBoldEnd & sFontEnd
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionName) & sBoldEnd & sFontEnd
							If iGroupGradeLevelID = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sGroupGradeLevelShortName) & sBoldEnd & sFontEnd
							End If
							If iClassificationID = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(iClassificationID) & sBoldEnd & sFontEnd
							End If
							If iIntegrationID = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(iIntegrationID) & sBoldEnd & sFontEnd
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
							If CLng(sEndDate) = 30000000 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
							End If
							If bActiveConcept_01 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=1&RecordID=" & lConcept_01_RecordID & "&ConceptID=1&ChangeEndDate=1&StartDate=" & lConcept_01_StartDate & """"
									sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
									sFontEnd = "</FONT>"
								Else
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=1&RecordID=" & lConcept_01_RecordID & "&ConceptID=1&StartDate=" & lConcept_01_StartDate & """"
									sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
									sFontEnd = "</FONT>"
								End If
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_01)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_01)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
							End If
							If Not IsEmpty(lConcept_01_RecordID) Then sRecordIDs = sRecordIDs & lConcept_01_RecordID & ","
							If bActiveConcept_03 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=1&RecordID=" & lConcept_03_RecordID & "&ConceptID=3&ChangeEndDate=1&StartDate=" & lConcept_03_StartDate & """"
									sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
									sFontEnd = "</FONT>"
								Else
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=1&RecordID=" & lConcept_03_RecordID & "&ConceptID=3&StartDate=" & lConcept_03_StartDate & """"
									sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
									sFontEnd = "</FONT>"
								End If
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_03)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_03)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
							End If
							If Not IsEmpty(lConcept_03_RecordID) Then sRecordIDs = sRecordIDs & lConcept_03_RecordID & ","
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(Round((CDbl(dConcept_01) + CDbl(dConcept_03))*2,2), 2, True, False, True)
							If (Not bActiveConcept_01) And (Not bActiveConcept_03) Then
								If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
									If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
									sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
									If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
										sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
											sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
										sRowContents = sRowContents & "</A>&nbsp;"
									End If
								End If
							ElseIf (bActiveConcept_01) And (bActiveConcept_03) Then
								If (Not bForExport) And (iStatusID=1) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
									sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&ChangeEndDate=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar fecha de fin de vigencia"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
								End If
							End If
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If (dConcept_01 + dConcept_03) > 0 Then
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
							End If
						Case 2,4
							sBoldBegin = ""
							sBoldEnd = ""
							If (StrComp(CStr(lConcept_01_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_03_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_01_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_03_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) _
							Then
								sBoldBegin = "<B>"
								sBoldEnd = "</B>"
							End If
							sFontBegin = ""
							sFontEnd = ""
							If (Not bActiveConcept_01) And (Not bActiveConcept_03) And (Not bActiveConcept_01_Z3) And (Not bActiveConcept_03_Z3) Then
								sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
								sFontEnd = "</FONT>"
							End If
							If iCompanyID = -1 Then
								sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
							Else
								sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(sCompanyName) & sBoldEnd & sFontEnd
							End If
							If iSelectedTab = 2 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionTypeShortName) & sBoldEnd & sFontEnd
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
							End If
							If iLevelID = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
							Else
								If bForExport Then
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & "=T(""" & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & """)" & sBoldEnd & sFontEnd
								Else
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & sBoldEnd & sFontEnd
								End If
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionName) & sBoldEnd & sFontEnd
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
							If CLng(sEndDate) = 30000000 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
							End If
							If bActiveConcept_01 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_01_RecordID & "&ConceptID=1&ChangeEndDate=1&StartDate=" & lConcept_01_StartDate & """"
									sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
									sFontEnd = "</FONT>"
								Else
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_01_RecordID & "&ConceptID=1&StartDate=" & lConcept_01_StartDate & """"
									sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
									sFontEnd = "</FONT>"
								End If
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_01)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_01)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
							End If
							If Not IsEmpty(lConcept_01_RecordID) Then sRecordIDs = sRecordIDs & lConcept_01_RecordID & ","
							If bActiveConcept_03 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_03_RecordID & "&ConceptID=3&ChangeEndDate=1&StartDate=" & lConcept_03_StartDate & """"
									sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
									sFontEnd = "</FONT>"
								Else
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_03_RecordID & "&ConceptID=3&StartDate=" & lConcept_03_StartDate & """"
									sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
									sFontEnd = "</FONT>"
								End If
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_03)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_03)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
							End If
							If Not IsEmpty(lConcept_03_RecordID) Then sRecordIDs = sRecordIDs & lConcept_03_RecordID & ","
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(Round((CDbl(dConcept_01) + CDbl(dConcept_03))*2,2), 2, True, False, True)
							If bActiveConcept_01_Z3 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_01_Z3_RecordID & "&ConceptID=1&ChangeEndDate=1&StartDate=" & lConcept_01_Z3_StartDate & """"
									sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
									sFontEnd = "</FONT>"
								Else
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_01_Z3_RecordID & "&ConceptID=1&StartDate=" & lConcept_01_Z3_StartDate & """"
								End If
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_01_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_01_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
							End If
							If Not IsEmpty(lConcept_01_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_01_Z3_RecordID & ","
							If bActiveConcept_03_Z3 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_03_Z3_RecordID & "&ConceptID=3&ChangeEndDate=1&StartDate=" & lConcept_03_Z3_StartDate & """"
									sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
									sFontEnd = "</FONT>"
								Else
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_03_Z3_RecordID & "&ConceptID=3&StartDate=" & lConcept_03_Z3_StartDate & """"
									sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
									sFontEnd = "</FONT>"
								End If
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_03_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_03_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
							End If
							If Not IsEmpty(lConcept_03_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_03_Z3_RecordID & ","
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(Round((CDbl(dConcept_01_Z3) + CDbl(dConcept_03_Z3))*2,2), 2, True, False, True)
							If (Not bActiveConcept_01) And (Not bActiveConcept_03) And (Not bActiveConcept_01_Z3) And (Not bActiveConcept_03_Z3) Then
								If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
									If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
									sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
									If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
										sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
											sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
										sRowContents = sRowContents & "</A>&nbsp;"
									End If
								End If
							ElseIf (bActiveConcept_01) And (bActiveConcept_03) And (bActiveConcept_01_Z3) And (bActiveConcept_03_Z3) Then
								If (Not bForExport) And (iStatusID=1) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
									sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&ChangeEndDate=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar fecha de fin de vigencia"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
								End If
							End If
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If (dConcept_01 + dConcept_03 + dConcept_01_Z3 + dConcept_03_Z3) > 0 Then
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
							End If
						Case 3
							sBoldBegin = ""
							sBoldEnd = ""
							If (StrComp(CStr(lConcept_01_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_03_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) _
							Then
								sBoldBegin = "<B>"
								sBoldEnd = "</B>"
							End If
							sFontBegin = ""
							sFontEnd = ""
							If (Not bActiveConcept_01) And (Not bActiveConcept_03) Then
								sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
								sFontEnd = "</FONT>"
							End If
							If iCompanyID = -1 Then
								sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
							Else
								sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(sCompanyName) & sBoldEnd & sFontEnd
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
							If iLevelID = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
							Else
								If bForExport Then
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & "=T(""" & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & """)" & sBoldEnd & sFontEnd
								Else
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & sBoldEnd & sFontEnd
								End If
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionName) & sBoldEnd & sFontEnd
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
							If CLng(sEndDate) = 30000000 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
							End If
							If bActiveConcept_01 Then
								sFontBegin = ""
								sFontEnd = ""
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=3&RecordID=" & lConcept_01_RecordID & "&ConceptID=1&ChangeEndDate=1&StartDate=" & lConcept_01_StartDate & """"
									sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
									sFontEnd = "</FONT>"
								Else
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=3&RecordID=" & lConcept_01_RecordID & "&ConceptID=1&StartDate=" & lConcept_01_StartDate & """"
									sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
									sFontEnd = "</FONT>"
								End If
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_01)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_01)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
							End If
							If Not IsEmpty(lConcept_01_RecordID) Then sRecordIDs = sRecordIDs & lConcept_01_RecordID & ","
							If bActiveConcept_03 Then
								sFontBegin = ""
								sFontEnd = ""
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=3&RecordID=" & lConcept_03_RecordID & "&ConceptID=3&ChangeEndDate=1&StartDate=" & lConcept_03_StartDate & """"
									sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
									sFontEnd = "</FONT>"
								Else
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=3&RecordID=" & lConcept_03_RecordID & "&ConceptID=3&StartDate=" & lConcept_03_StartDate & """"
									sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
									sFontEnd = "</FONT>"
								End If
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_03)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_03)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
							End If
							If Not IsEmpty(lConcept_03_RecordID) Then sRecordIDs = sRecordIDs & lConcept_03_RecordID & ","
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(Round((CDbl(dConcept_01) + CDbl(dConcept_03))*2,2), 2, True, False, True)
							If (Not bActiveConcept_01) And (Not bActiveConcept_03) Then
								If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
									If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
									sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
									If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
										sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
											sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
										sRowContents = sRowContents & "</A>&nbsp;"
									End If
								End If
							ElseIf (bActiveConcept_01) And (bActiveConcept_03) Then
								If (Not bForExport) And (iStatusID=1) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
									sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&ChangeEndDate=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar fecha de fin de vigencia"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
								End If
							End If
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If (dConcept_01 + dConcept_03) > 0 Then
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
							End If
						Case 5
							sBoldBegin = ""
							sBoldEnd = ""
							If (StrComp(CStr(lConcept_B2_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_36_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_B2_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Or _
								(StrComp(CStr(lConcept_36_Z3_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) _
							Then
								sBoldBegin = "<B>"
								sBoldEnd = "</B>"
							End If
							sFontBegin = ""
							sFontEnd = ""
							If (Not bActiveConcept_B2) And (Not bActiveConcept_36) And (Not bActiveConcept_B2_Z3) And (Not bActiveConcept_36_Z3) Then
								sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
								sFontEnd = "</FONT>"
							End If
							If iCompanyID = -1 Then
								sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
							Else
								sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(sCompanyName) & sBoldEnd & sFontEnd
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionName) & sBoldEnd & sFontEnd
							If iLevelID = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
							Else
								If bForExport Then
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & "=T(""" & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & """)" & sBoldEnd & sFontEnd
								Else
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & sBoldEnd & sFontEnd
								End If
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
							If CLng(sEndDate) = 30000000 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
							End If
							If bActiveConcept_B2 Then
								sFontBegin = ""
								sFontEnd = ""
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=5&RecordID=" & lConcept_B2_RecordID & "&ConceptID=89&ChangeEndDate=1&StartDate=" & lConcept_B2_StartDate & """"
									sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
									sFontEnd = "</FONT>"
								Else
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=5&RecordID=" & lConcept_B2_RecordID & "&ConceptID=89&StartDate=" & lConcept_B2_StartDate & """"
									sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
									sFontEnd = "</FONT>"
								End If
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_B2)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_B2)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_B2_RecordID) Then sRecordIDs = sRecordIDs & lConcept_B2_RecordID & ","
							End If
							If bActiveConcept_36 Then
								sFontBegin = ""
								sFontEnd = ""
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=5&RecordID=" & lConcept_36_RecordID & "&ConceptID=39&ChangeEndDate=1&StartDate=" & lConcept_36_StartDate & """"
									sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
									sFontEnd = "</FONT>"
								Else
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=5&RecordID=" & lConcept_36_RecordID & "&ConceptID=39&StartDate=" & lConcept_36_StartDate & """"
									sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
									sFontEnd = "</FONT>"
								End If
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_36)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_36)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_36_RecordID) Then sRecordIDs = sRecordIDs & lConcept_36_RecordID & ","
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(Round((CDbl(dConcept_B2) + CDbl(dConcept_36))*2,2), 2, True, False, True)
							If bActiveConcept_B2_Z3 Then
								sFontBegin = ""
								sFontEnd = ""
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=5&RecordID=" & lConcept_B2_Z3_RecordID & "&ConceptID=89&ChangeEndDate=1&StartDate=" & lConcept_B2_StartDate & """"
									sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
									sFontEnd = "</FONT>"
								Else
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=5&RecordID=" & lConcept_B2_Z3_RecordID & "&ConceptID=89&StartDate=" & lConcept_B2_StartDate & """"
									sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
									sFontEnd = "</FONT>"
								End If
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_B2_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_B2_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_B2_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_B2_Z3_RecordID & ","
							End If
							If bActiveConcept_36_Z3 Then
								sFontBegin = ""
								sFontEnd = ""
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=5&RecordID=" & lConcept_36_Z3_RecordID & "&ConceptID=39&ChangeEndDate=1&StartDate=" & lConcept_36_StartDate & """"
									sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
									sFontEnd = "</FONT>"
								Else
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=5&RecordID=" & lConcept_36_Z3_RecordID & "&ConceptID=39&StartDate=" & lConcept_36_StartDate & """"
									sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
									sFontEnd = "</FONT>"
								End If
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_36_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_36_Z3)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
								If Not IsEmpty(lConcept_36_Z3_RecordID) Then sRecordIDs = sRecordIDs & lConcept_36_Z3_RecordID & ","
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(Round((CDbl(dConcept_B2_Z3) + CDbl(dConcept_36_Z3))*2,2), 2, True, False, True)
							If (Not bActiveConcept_B2) And (Not bActiveConcept_36) And (Not bActiveConcept_B2_Z3) And (Not bActiveConcept_36_Z3) Then
								If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
									If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
									sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
									If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
										sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
											sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
										sRowContents = sRowContents & "</A>&nbsp;"
									End If
								End If
							ElseIf (bActiveConcept_B2) And (bActiveConcept_36) And (bActiveConcept_B2_Z3) And (bActiveConcept_36_Z3) Then
								If (Not bForExport) And (iStatusID=1) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
									sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&ChangeEndDate=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar fecha de fin de vigencia"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
								End If
							End If
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If (dConcept_B2 + dConcept_36) > 0 Then
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
							End If
						Case 6
							sBoldBegin = ""
							sBoldEnd = ""
							If (StrComp(CStr(lConcept_12_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Then
								sBoldBegin = "<B>"
								sBoldEnd = "</B>"
							End If
							sFontBegin = ""
							sFontEnd = ""
							If (Not bActiveConcept_12) Then
								sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
								sFontEnd = "</FONT>"
							End If
							If iCompanyID = -1 Then
								sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
							Else
								sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(sCompanyName) & sBoldEnd & sFontEnd
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionShortName) & sBoldEnd & sFontEnd
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(sPositionName) & sBoldEnd & sFontEnd
							If iLevelID = -1 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
							Else
								If bForExport Then
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & "=T(""" & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & """)" & sBoldEnd & sFontEnd
								Else
									sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & sBoldEnd & sFontEnd
								End If
							End If
							If iEconomicZoneID = 0 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(iEconomicZoneID) & sBoldEnd & sFontEnd
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
							If CLng(sEndDate) = 30000000 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
							End If
							If bActiveConcept_12 Then
								sFontBegin = ""
								sFontEnd = ""
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
								If (CInt(Request.Cookies("SIAP_SectionID")) = 3) And ((aLoginComponent(N_PROFILE_ID_LOGIN) <> 4) And (aLoginComponent(N_PROFILE_ID_LOGIN) <> -1)) Then
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&EmployeeTypeID=6&RecordID=" & lConcept_12_RecordID & "&ConceptID=14&ChangeEndDate=1&StartDate=" & lConcept_12_StartDate & """"
									sFontBegin = "<FONT TITLE=""Cambiar fecha de fin"">"
									sFontEnd = "</FONT>"
								Else
									sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=6&RecordID=" & lConcept_12_RecordID & "&ConceptID=14&StartDate=" & lConcept_12_StartDate & """"
									sFontBegin = "<FONT TITLE=""Modificar información del concepto"">"
									sFontEnd = "</FONT>"
								End If
								sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_12)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
							Else
								sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
								sFontEnd = "</FONT>"
								sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept_12)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
							End If
							If Not IsEmpty(lConcept_12_RecordID) Then sRecordIDs = sRecordIDs & lConcept_12_RecordID & ","
							If Not bActiveConcept_12 Then
								If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
									If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
									sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
									If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
										sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
											sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
										sRowContents = sRowContents & "</A>&nbsp;"
									End If
								End If
							ElseIf bActiveConcept_12 Then
								If (Not bForExport) And (iStatusID=1) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
									sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&ChangeEndDate=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar fecha de fin de vigencia"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
								End If
							End If
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If dConcept_12 > 0 Then
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
							End If
					End Select
				Case Else
					sBoldBegin = ""
					sBoldEnd = ""
					If (StrComp(CStr(lConcept_RecordID), oRequest("RecordID").Item, vbBinaryCompare) = 0) Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sFontBegin = ""
					sFontEnd = ""
					If (Not bActiveConcept) Then
						sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
						sFontEnd = "</FONT>"
					End If
					If iCompanyID = -1 Then
						sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
					Else
						sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(sCompanyName) & sBoldEnd & sFontEnd
					End If
					If CInt(oRecordset.Fields("PositionTypeID").Value) = -1 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PositionTypeShortName").Value)) & sBoldEnd & sFontEnd
					End If
					If CInt(oRecordset.Fields("PositionID").Value) = -1 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(oRecordset.Fields("PositionShortName").Value) & sBoldEnd & sFontEnd
					End If
					If CInt(oRecordset.Fields("LevelID").Value) = -1 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(oRecordset.Fields("LevelShortName").Value) & sBoldEnd & sFontEnd
					End If
					If CSng(oRecordset.Fields("WorkingHours").Value) = -1 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("WorkingHours").Value)) & " Hrs." & sBoldEnd & sFontEnd
					End If
					If CInt(oRecordset.Fields("PositionID").Value) = -1 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CStr(oRecordset.Fields("PositionName").Value) & sBoldEnd & sFontEnd
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sStartDate)) & sBoldEnd & sFontEnd
					If CLng(sEndDate) = 30000000 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(sEndDate)) & sBoldEnd & sFontEnd
					End If
					If bActiveConcept Then
						sFontBegin = ""
						sFontEnd = ""
						sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
						sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=" & iSelectedTab & "&RecordID=" & lConcept_RecordID & "&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&StartDate=" & lConcept_StartDate & """"
						sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
					Else
						sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
						sFontEnd = "</FONT>"
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(Round(CDbl(dConcept)*2,2), 2, True, False, True) & sBoldEnd & sFontEnd
						If Not IsEmpty(lConcept_RecordID) Then sRecordIDs = sRecordIDs & lConcept_RecordID & ","
					End If
					If Not bActiveConcept Then
						If (Not bForExport) And (iStatusID=0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
							If  InStr(1, Right(sRecordIDs, 1), ",") > 0 Then sRecordIDs = Left(sRecordIDs, Len(sRecordIDs) -1)
							sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Remove=1"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;"
							If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
								sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & sRecordIDs & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&Apply=1"">"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;"
							End If
						End If
					End If
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If dConcept > 0 Then
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
					End If
			End Select
			Response.Write "</TABLE></DIV><BR /><BR />"
		Else
			If (iStatusID = 1) Then
				If (Len(sStartDateCondition) = 0) And (InStr(1, sCondition, "Positions.PositionID In (0)", vbBinaryCompare) > 0) Then
					lErrorNumber = L_ERR_NO_RECORDS
					sErrorDescription = "Seleccione un rango de fechas (por lo menos la fecha de inicio, si no se indica la fecha de fin se buscaran con fecha indefinida) o un puesto del filtro para poder consultar los tabuladores existentes."
				Else
					lErrorNumber = L_ERR_NO_RECORDS
					sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
				End If
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
			End If
		End If
	End If

	Set oRecordset = Nothing
	DisplayConceptValuesTableSP = lErrorNumber
	Err.Clear
End Function

Function DisplayConceptValuesTable(oRequest, oADODBConnection, iSelectedTab, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the ConceptValues for Concepts
'Inputs:  oRequest, oADODBConnection, iSelectedTab, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayConceptValuesTable"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim lPayrollNumber
	Dim sDate
	Dim sStartDate
	Dim sEndDate
	Dim sFilePath
	Dim lReportID
	Dim sTemp
	Dim lCurrentID
	Dim dTotal
	Dim oRecordset
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asColumnsTitles
	Dim asCellWidths
	Dim asCellAlignments
	Dim sColumnsTitles
	Dim sCellWidths
	Dim sCellAlignments
	Dim lErrorNumber
	Dim sBoldBegin
	Dim sBoldEnd
	Dim sFontBegin
	Dim sFontEnd
	Dim bFirst
	Dim bContinue
	Dim sConceptIDs
	Dim iStatusID
	Dim iConceptsValuesStatusID
	Dim sRecordIDs

	Dim iLevelID
	Dim iAntiquityID
	Dim iAntiquityID2
	Dim iEconomicZoneID
	Dim iGroupGradeLevelID
	Dim iClassificationID
	Dim iIntegrationID
	Dim iPositionTypeID
	Dim sPositionTypeShortName
	Dim sPositionShortName
	Dim sLevelShortName
	Dim sWorkingHours
	Dim sPositionName
	Dim sGroupGradeLevelShortName
	Dim lCurrentPositionID

	Dim dConcept
	Dim lConcept_RecordID
	Dim lConcept_StartDate
	Dim bActiveConcept
	Dim iRecordCounter

	sDate = Left(GetSerialNumberForDate(""), Len("00000000"))

	sErrorDescription = "No se pudieron obtener los montos pagados."
	iStatusID = aConceptComponent(N_STATUS_ID_CONCEPT)
	If iStatusID=0 Then
		sCondition = sCondition & " And (ConceptsValues.StatusID<=" & iStatusID & ")"
	Else
		sCondition = sCondition & " And (ConceptsValues.StatusID=" & iStatusID & ")"
	End If
	If Len(oRequest("ConceptID").Item) > 0 Then
		lErrorNumber = GetConcept(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
		sStartDate = aConceptComponent(N_START_DATE_CONCEPT)
		sEndDate = aConceptComponent(N_END_DATE_CONCEPT)
		sCondition = sCondition & " And (ConceptID=" & CStr(oRequest("ConceptID").Item) & ")"
	Else
		sStartDate = sDate
		sEndDate = sDate
		sCondition = sCondition & " And (ConceptID=-3)"
	End If
	sCondition = sCondition & " And (((Positions.StartDate>="& sStartDate & ") And (Positions.StartDate<=" & sEndDate & ")) Or ((Positions.EndDate>=" & sStartDate & ") And (Positions.EndDate<=" & sEndDate & ")) Or ((Positions.EndDate>=" & sStartDate & ") And (Positions.StartDate<=" & sEndDate & ")))"
	sCondition = sCondition & " And (((PositionTypes.StartDate>=" & sStartDate & ") And (PositionTypes.StartDate<=" & sEndDate & ")) Or ((PositionTypes.EndDate>=" & sStartDate & ") And (PositionTypes.EndDate<=" & sEndDate & ")) Or ((PositionTypes.EndDate>=" & sStartDate & ") And (PositionTypes.StartDate<=" & sEndDate & ")))"
	sCondition = sCondition & " And (((GroupGradeLevels.StartDate>=" & sStartDate & ") And (GroupGradeLevels.StartDate<=" & sEndDate & ")) Or ((GroupGradeLevels.EndDate>=" & sStartDate & ") And (GroupGradeLevels.EndDate<=" & sEndDate & ")) Or ((GroupGradeLevels.EndDate>=" & sStartDate & ") And (GroupGradeLevels.StartDate<=" & sEndDate & ")))"
	sCondition = sCondition & " And (((Levels.StartDate>=" & sStartDate & ") And (Levels.StartDate<=" & sEndDate & ")) Or ((Levels.EndDate>=" & sStartDate & ") And (Levels.EndDate<=" & sEndDate & ")) Or ((Levels.EndDate>=" & sStartDate & ") And (Levels.StartDate<=" & sEndDate & ")))"
	sCondition = sCondition & " And (ConceptsValues.EmployeeTypeID IN (-1, " & iSelectedTab & "))" & " And (Positions.EmployeeTypeID IN (-1, " & iSelectedTab & "))"
	aConceptComponent(S_QUERY_CONDITION_CONCEPT) = aConceptComponent(S_QUERY_CONDITION_CONCEPT) & sCondition

	lErrorNumber = GetConceptValues(oRequest, oADODBConnection, aConceptComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If Not bForExport Then Call DisplayIncrementalFetch(oRequest, CInt(oRequest("StartPage").Item), ROWS_REPORT, oRecordset)
			Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><TABLE BORDER="""
			If bForExport Then
				Response.Write "1"
			Else
				Response.Write "0"
			End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine

			sColumnsTitles = "Tipo tabulador,Tipo puesto,Código,Nivel,Jornada,Denominación del puesto,Zona económica,Fecha Inicio,Fecha Fin,Importe"
			sCellWidths = ",,,,,,,,"
			sCellAlignments = "CENTER,RIGHT,RIGHT,RIGHT,LEFT,RIGHT,RIGHT,RIGHT"
			If (Not bForExport) And (iStatusID=0) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Or (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
				sColumnsTitles = sColumnsTitles & ",Acciones"
				sCellWidths = sCellWidths & ",90"
				sCellAlignments = sCellAlignments & ",CENTER"
			End If
			asColumnsTitles = Split(sColumnsTitles, ",", -1, vbBinaryCompare)
			asCellWidths = Split(sCellWidths, ",", -1, vbBinaryCompare)
			asCellAlignments = Split(sCellAlignments, ",", -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
			Else
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If
			End If
			lCurrentPositionID = -2
			dTotal = 0
			sBoldBegin = "<B>"
			sBoldEnd = "</B>"
			sFontBegin = ""
			sFontEnd = ""
			iRecordCounter = 0
			Do While Not oRecordset.EOF
				bContinue = False
				sBoldBegin = ""
				sBoldEnd = ""
				If StrComp(CStr(oRecordset.Fields("RecordID").Value), oRequest("RecordID").Item, vbBinaryCompare) = 0 Then
					sBoldBegin = "<B>"
					sBoldEnd = "</B>"
				End If
				sFontBegin = ""
				sFontEnd = ""
				If Not (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And (CInt(oRecordset.Fields("StatusID").Value) = 1) Then
					sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
					sFontEnd = "</FONT>"
				End If
				aConceptComponent(N_ID_CONCEPT) = CInt(oRecordset.Fields("ConceptID").Value)
				If CInt(oRecordset.Fields("EmployeeTypeID").Value) = -1 Then
					sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
				Else
					sRowContents = sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeName").Value)) & sBoldEnd & sFontEnd
				End If
				If CInt(oRecordset.Fields("PositionTypeID").Value) = -1 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PositionTypeShortName").Value)) & sBoldEnd & sFontEnd
				End If
				If CInt(oRecordset.Fields("PositionID").Value) = -1 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value)) & sBoldEnd & sFontEnd
				End If
				If iLevelID = -1 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
				Else
					sLevelShortName = CStr(oRecordset.Fields("LevelShortName").Value)
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(Left(sLevelShortName, Len("00")) & "-" & Right(sLevelShortName, Len("0"))) & sBoldEnd & sFontEnd
				End If
				If CSng(oRecordset.Fields("WorkingHours").Value) = -1 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CSng(oRecordset.Fields("WorkingHours").Value)) & " Hrs." & sBoldEnd & sFontEnd
				End If
				If CInt(oRecordset.Fields("PositionID").Value) = -1 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todos") & sBoldEnd & sFontEnd
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value)) & sBoldEnd & sFontEnd
				End If
				If CInt(oRecordset.Fields("EconomicZoneID").Value) = 0 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("Todas") & sBoldEnd & sFontEnd
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CInt(oRecordset.Fields("EconomicZoneID").Value)) & sBoldEnd & sFontEnd
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)) & sBoldEnd & sFontEnd
				If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & sBoldEnd & sFontEnd
				End If
				If (CLng(oRecordset.Fields("EndDate").Value) > CLng(sDate)) And (CInt(oRecordset.Fields("StatusID").Value) = 1) Then
					sFontBegin = ""
					sFontEnd = ""
					sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
					sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&EmployeeTypeID=" & iSelectedTab & "&StartPage=" & CInt(oRequest("StartPage").Item) & "&RecordID=" & CLng(oRecordset.Fields("RecordID").Value) & "&ConceptID=" & CLng(oRecordset.Fields("ConceptID").Value) & "&StartDate=" & CLng(oRecordset.Fields("StartDate").Value) & """"
					sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & sBoldEnd & sFontEnd & "</A>"
				Else
					sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
					sFontEnd = "</FONT>"
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True) & sBoldEnd & sFontEnd
				End If
				If (Not bForExport) And (CInt(oRecordset.Fields("StatusID").Value) <= 0) And (B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
					sRowContents = sRowContents & TABLE_SEPARATOR 
					Select Case CInt(oRecordset.Fields("StatusID").Value)
						Case 0
							sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
						Case -1
							sRowContents = sRowContents & "<IMG SRC=""Images/IcnExclamationSmall.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Al agregar este registro se detectaron registros posteriores que serán ajustados al aplicar este registro"" BORDER=""0"" />"
							sRowContents = sRowContents & "&nbsp;"
						Case -2
							sRowContents = sRowContents & "<IMG SRC=""Images/IcnInformation.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Al agregar este registro se detectaron registros dentro de los efectos de este que se ajustaran al aplicar este registro"" BORDER=""0"" />"
							sRowContents = sRowContents & "&nbsp;"
						Case -3
							sRowContents = sRowContents & "<IMG SRC=""Images/IcnExclamationSmall.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Al agregar este registro se detectaron registros que cubren todo el periodo de este, los cuales se se ajustaran al aplicar este registro"" BORDER=""0"" />"
							sRowContents = sRowContents & "&nbsp;"
					End Select
					sRowContents = sRowContents & "&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & CLng(oRecordset.Fields("RecordID").Value) & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&Remove=1"">"
						sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
					sRowContents = sRowContents & "</A>&nbsp;"
					sRowContents = sRowContents & "&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptValuesAction=1&RecordID=" & CLng(oRecordset.Fields("RecordID").Value) & "&Tab=" & iSelectedTab & "&EmployeeTypeID=" & iSelectedTab & "&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&Apply=1"">"
						sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
					sRowContents = sRowContents & "</A>&nbsp;"
				End If
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				oRecordset.MoveNext
				iRecordCounter = iRecordCounter + 1
				If (Not bForExport) And (iRecordCounter >= ROWS_REPORT) Then Exit Do
				If Err.number <> 0 Then Exit Do
			Loop
			Response.Write "</TABLE></DIV><BR /><BR />"
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen tabuladores de pago registrados en el sistema para el concepto seleccionado."
		End If
	End If

	Set oRecordset = Nothing
	DisplayConceptValuesTable = lErrorNumber
	Err.Clear
End Function

Function DisplayConceptValuesTabs(oRequest, lConceptID, iSelectedTab, sErrorDescription)
'************************************************************
'Purpose: To display the tabs for the concepts HTML forms
'Inputs:  oRequest, lConceptID, iSelectedTab
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayConceptValuesTabs"
	Dim oRecordset
	Dim lErrorNumber

	Response.Write "<TABLE BORDER=""0"" WIDTH=""98%"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
		sErrorDescription = "No se pudieron obtener los tipos de tabuladores."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeTypes.EmployeeTypeID, EmployeeTypeName From EmployeeTypes Where (EmployeeTypes.EmployeeTypeID >= 0) And (EmployeeTypes.EmployeeTypeID <= 6)  Order By EmployeeTypes.EmployeeTypeID", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If iSelectedTab = -1 Then
				iSelectedTab = CInt(oRecordset.Fields("EmployeeTypeID").Value)
				aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) = iSelectedTab
			End If
			Do While Not oRecordset.EOF
				Response.Write "<TD BGCOLOR=""#"
					If iSelectedTab = CInt(oRecordset.Fields("EmployeeTypeID").Value) Then
						Response.Write S_MAIN_COLOR_FOR_GUI
					Else
						Response.Write "CCCCCC"
					End If
				Response.Write """ WIDTH=""5"" NAME=""TabContents" & CStr(oRecordset.Fields("EmployeeTypeID").Value) & "LfDiv"" ID=""TabContents" & CStr(oRecordset.Fields("EmployeeTypeID").Value) & "LfDiv""><IMG SRC=""Images/TbLf.gif"" WIDTH=""5"" HEIGHT=""21"" /></TD>"
				Response.Write "<TD BGCOLOR=""#"
					If iSelectedTab = CInt(oRecordset.Fields("EmployeeTypeID").Value) Then
						Response.Write S_MAIN_COLOR_FOR_GUI
					Else
						Response.Write "CCCCCC"
					End If
				Response.Write """ BACKGROUND=""Images/TbBg.gif"" WIDTH=""130"" ALIGN=""CENTER"" NAME=""TabContents" & CStr(oRecordset.Fields("EmployeeTypeID").Value) & "Div"" ID=""TabContents" & CStr(oRecordset.Fields("EmployeeTypeID").Value) & "Div""><NOBR><FONT FACE=""Arial"" SIZE=""2"">"
				Response.Write "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&ConceptID=" & lConceptID & "&Perceptions=1&Tab=" & CStr(oRecordset.Fields("EmployeeTypeID").Value) & """ CLASS=""TabLink""><DIV NAME=""TabText" & CStr(oRecordset.Fields("EmployeeTypeID").Value) & "Div"" ID=""TabText" & CStr(oRecordset.Fields("EmployeeTypeID").Value) & "Div"" STYLE=""color: #"
					If iSelectedTab = CInt(oRecordset.Fields("EmployeeTypeID").Value) Then
						Response.Write S_MENU_LINK_FOR_GUI
					Else
						Response.Write "000000"
					End If
				Response.Write ";""><B>&nbsp;&nbsp;&nbsp;" & CStr(oRecordset.Fields("EmployeeTypeName").Value) & "&nbsp;&nbsp;&nbsp;</B></DIV></A></FONT></NOBR></TD>"
				Response.Write "<TD BGCOLOR=""#"
					If iSelectedTab = CInt(oRecordset.Fields("EmployeeTypeID").Value) Then
						Response.Write S_MAIN_COLOR_FOR_GUI
					Else
						Response.Write "CCCCCC"
					End If
				Response.Write """ WIDTH=""5"" NAME=""TabContents" & CStr(oRecordset.Fields("EmployeeTypeID").Value) & "RgDiv"" ID=""TabContents" & CStr(oRecordset.Fields("EmployeeTypeID").Value) & "RgDiv""><IMG SRC=""Images/TbRg.gif"" WIDTH=""5"" HEIGHT=""21"" /></TD>"

				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
		End If
		Response.Write "<TD BACKGROUND=""Images/TbBgDot.gif"" WIDTH=""*""><IMG SRC=""Images/Transparent.gif"" WIDTH=""21"" HEIGHT=""21"" /></TD>"
	Response.Write "</TR></TABLE><BR />"

	Set oRecordset = Nothing
	DisplayConceptValuesTabs = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeeTypesTable(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the concepts
'		  from the database in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, aConceptComponent
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeTypesTable"
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

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeTypeID, EmployeeTypeShortName, EmployeeTypeName From EmployeeTypes Where (EmployeeTypeID>=0) And (EmployeeTypeID <= 6) And (Active = 1) Order By EmployeeTypeID", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE WIDTH=""450"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("&nbsp;,Id,Clave,Nombre", ",", -1, vbBinaryCompare)
				asCellWidths = Split("10,10,10,300", ",", -1, vbBinaryCompare)
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If

				asCellAlignments = Split(",,,", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sFontBegin = ""
					sFontEnd = ""
					sBoldBegin = ""
					sBoldEnd = ""
					sRowContents = TABLE_SEPARATOR & CStr(oRecordset.Fields("EmployeeTypeID").Value)
					sRowContents = sRowContents & TABLE_SEPARATOR & CStr(oRecordset.Fields("EmployeeTypeShortName").Value)
					sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Action=ConceptsValues&Perceptions=1&EmployeeTypeID=" & CStr(oRecordset.Fields("EmployeeTypeID").Value) & """>" & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeName").Value)) & sBoldEnd & sFontEnd & "</A>"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					oRecordset.MoveNext
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
	DisplayEmployeeTypesTable = lErrorNumber
	Err.Clear
End Function
%>