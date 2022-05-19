<%
Function GetPaperworksOwnersForUser(sOwnerIDs, sErrorDescription)
'************************************************************
'Purpose: To initialize the global variables using the URL
'Inputs:  oRequest
'Outputs: bAction, sCondition
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPaperworksOwnersForUser"
	Dim oItem
	Dim aItem
	Dim oRecordset
	Dim lErrorNumber

	sOwnerIDs = "-2"
	sErrorDescription = "No se pudieron obtener los permisos del usuario."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select OwnerID From UsersOwnersLKP Where (UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ")", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Do While Not oRecordset.EOF
			sOwnerIDs = sOwnerIDs & "," & CStr(oRecordset.Fields("OwnerID").Value)
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
		oRecordset.Close
	End If

	sErrorDescription = "No se pudieron obtener los permisos del usuario."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select OwnerID From UsersOwnersLKP Where (UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ")", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Do While Not oRecordset.EOF
			sOwnerIDs = sOwnerIDs & "," & CStr(oRecordset.Fields("OwnerID").Value)
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
		oRecordset.Close
	End If
	If InStr(1, sOwnerIDs & ",", ",-1,", vbBinaryCompare) = 0 Then
		sErrorDescription = "No se pudieron obtener los permisos del usuario."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select OwnerID From PaperworkOwners Where (ParentID In (" & sOwnerIDs & ")) And (OwnerID>-1)", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				sOwnerIDs = sOwnerIDs & "," & CStr(oRecordset.Fields("OwnerID").Value)
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
		End If
		sErrorDescription = "No se pudieron obtener los permisos del usuario."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select OwnerID From PaperworkOwners Where (ParentID In (" & sOwnerIDs & ")) And (OwnerID>-1)", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				sOwnerIDs = sOwnerIDs & "," & CStr(oRecordset.Fields("OwnerID").Value)
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
		End If
	End If

	GetPaperworksOwnersForUser = lErrorNumber
	Err.Clear
End Function

Function GetPaperworksURLValues(oRequest, bAction, bDisplayTable, sCondition)
'************************************************************
'Purpose: To initialize the global variables using the URL
'Inputs:  oRequest
'Outputs: bAction, sCondition
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPaperworksURLValues"
	Dim sOwnerIDs
	Dim oItem
	Dim aItem
	Dim oRecordset
	Dim lErrorNumber

	Call GetPaperworksOwnersForUser(sOwnerIDs, sErrorDescription)
	
	bAction = ((Len(oRequest("Add").Item) > 0) Or (Len(oRequest("Associate").Item) > 0) Or (Len(oRequest("Modify").Item) > 0) Or (Len(oRequest("Remove").Item) > 0) Or (Len(oRequest("DoClose").Item) > 0))
	bDisplayTable = ((Len(oRequest("DoSearch").Item) > 0) Or (Len(oRequest("Add").Item) > 0) Or (Len(oRequest("Change").Item) > 0) Or (Len(oRequest("Modify").Item) > 0) Or (Len(oRequest("Delete").Item) > 0))

	sCondition = ""
	If Len(oRequest("FilterStartNumber").Item) > 0 Then
		sCondition = sCondition & " And (PaperworkNumber>=" & Replace(oRequest("FilterStartNumber").Item, "´", "") & ")"
	End If
	If Len(oRequest("FilterEndNumber").Item) > 0 Then
		sCondition = sCondition & " And (PaperworkNumber<=" & Replace(oRequest("FilterEndNumber").Item, "´", "") & ")"
	End If
	If (InStr(1, oRequest, "StartStart", vbTextCompare) > 0) Or (InStr(1, oRequest, "EndStart", vbTextCompare) > 0) Then Call GetStartAndEndDatesFromURL("StartStart", "EndStart", "Paperworks.StartDate", False, sCondition)
	If Len(oRequest("FilterDocumentNumber").Item) > 0 Then
		sCondition = sCondition & " And (DocumentNumber Like ('" & S_WILD_CHAR & Replace(oRequest("FilterDocumentNumber").Item, "´", "") & S_WILD_CHAR & "'))"
	End If
	If Len(oRequest("SenderID").Item) > 0 Then
		sCondition = sCondition & " And (Paperworks.SenderID In (" & Replace(oRequest("SenderID").Item, ", ", ",") & "))"
	End If
	If Len(oRequest("FilterEmployeeID").Item) > 0 Then
		sCondition = sCondition & " And (Paperworks.OwnerID In (" & Replace(oRequest("FilterEmployeeID").Item, ", ", ",") & "))"
	End If
	If Len(oRequest("FilterDescription").Item) > 0 Then
		sCondition = sCondition & " And (Description Like ('" & S_WILD_CHAR & Replace(oRequest("FilterDescription").Item, "´", "") & S_WILD_CHAR & "'))"
	End If
	If Len(oRequest("FilterDocumentSubject").Item) > 0 Then
		sCondition = sCondition & " And (DocumentSubject Like ('" & S_WILD_CHAR & Replace(oRequest("FilterDocumentSubject").Item, "´", "") & S_WILD_CHAR & "'))"
	End If
	If Len(oRequest("FilterOwnerID").Item) > 0 Then
		If Len(oRequest("FullSearch").Item) > 0 Then
			sCondition = sCondition & " And (PaperworkOwners.ParentID=PaperworkOwners2.OwnerID) And (PaperworkOwners2.ParentID=PaperworkOwners1.OwnerID) And ((PaperworkOwners1.OwnerID In (" & Replace(oRequest("FilterOwnerID").Item, ", ", ",") & ")) Or (PaperworkOwners2.OwnerID In (" & Replace(oRequest("FilterOwnerID").Item, ", ", ",") & ")) Or (PaperworkOwners.OwnerID In (" & Replace(oRequest("FilterOwnerID").Item, ", ", ",") & ")))"
		Else
			sCondition = sCondition & " And (PaperworkOwners.OwnerID In (" & Replace(oRequest("FilterOwnerID").Item, ", ", ",") & "))"
		End If
	ElseIf InStr(1, "," & sOwnerIDs & ",", ",-1,", vbBinaryCompare) = 0 Then
		sCondition = sCondition & " And (PaperworkOwners.OwnerID In (" & sOwnerIDs & "))"
	End If
	If Len(oRequest("FilterPaperworkTypeID").Item) > 0 Then
		sCondition = sCondition & " And (Paperworks.PaperworkTypeID In (" & Replace(oRequest("FilterPaperworkTypeID").Item, ", ", ",") & "))"
	End If
	If Len(oRequest("SubjectTypeID").Item) > 0 Then
		sCondition = sCondition & " And (Paperworks.SubjectTypeID In (" & Replace(oRequest("SubjectTypeID").Item, ", ", ",") & "))"
	End If
	If Len(oRequest("FilterComments").Item) > 0 Then
		sCondition = sCondition & " And (Paperworks.Comments Like ('" & S_WILD_CHAR & Replace(oRequest("FilterComments").Item, "´", "") & S_WILD_CHAR & "'))"
	End If
	If Len(oRequest("FilterStatusID").Item) > 0 Then
		Select Case oRequest("FilterStatusID").Item
			Case "0"
				sCondition = sCondition & " And (Paperworks.StatusID In (" & Replace(oRequest("FilterStatusID").Item, ", ", ",") & ")) And (PaperworkOwnersLKP.EndDate=0)"
			Case "3"
				sCondition = sCondition & " And ((Paperworks.StatusID In (" & Replace(oRequest("FilterStatusID").Item, ", ", ",") & ")) Or (PaperworkOwnersLKP.EndDate<>0))"
			Case Else
				sCondition = sCondition & " And (Paperworks.StatusID In (" & Replace(oRequest("FilterStatusID").Item, ", ", ",") & "))"
		End Select
	End If
	If Len(oRequest("Closed").Item) > 0 Then
		If StrComp(oRequest("Closed").Item, "0", vbBinaryCompare) = 0 Then
			sCondition = sCondition & " And (PaperworkOwnersLKP.EndDate=0)"
		Else
			sCondition = sCondition & " And (PaperworkOwnersLKP.EndDate<>0)"
		End If
	End If
	If Len(oRequest("AreaID").Item) > 0 Then
		sCondition = sCondition & " And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (AreaPath Like '" & S_WILD_CHAR & oRequest("AreaID").Item & S_WILD_CHAR & "')"
	End If
	If (Len(sCondition) = 0) And (Len(oRequest("Remove").Item) = 0) And (Len(oRequest("PaperworkID").Item) > 0) Then sCondition = sCondition & " And (Paperworks.PaperworkID=" & oRequest("PaperworkID").Item & ")"

	GetPaperworksURLValues = Err.number
	Err.Clear
End Function

Function InitializeSupportComponent(oRequest, aCatalogComponent)
'************************************************************
'Purpose: To initialize the component for the employee support
'Inputs:  oRequest, aCatalogComponent
'Outputs: aCatalogComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeSupportComponent"

	Call InitializeCatalogComponent(oRequest, aCatalogComponent)
	aCatalogComponent(S_TABLE_NAME_CATALOG) = "Paperworks"
	aCatalogComponent(S_NAME_CATALOG) = "Trámites"
	aCatalogComponent(S_ORDER_CATALOG) = "PaperworkID"
	aCatalogComponent(N_NAME_CATALOG) = 1
	aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Número de folio,Fecha del documento,Documento,Procedencia,Empleado,Desc. Procedencia,Asunto,Tipo de trámite,Tipo de asunto,Fecha límite,Prioridad,Observaciones,Fecha de atención,Oficio de descargo,Estatus"
	aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
	aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "PaperworkID,PaperworkNumber,StartDate,DocumentNumber,SenderID,OwnerID,Description,DocumentSubject,PaperworkTypeID,SubjectTypeID,EstimatedDate,PriorityID,Comments,EndDate,DocClassification,StatusID"
	aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
	aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,0,0,1,1,1,0,1,0,0,0,1"
	aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
	aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,4,1,5,5,11,11,5,6,5,1,6,5,1,11,6"
	aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
	aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,10,0,100,100,6,2000,2000,0,100,0,0,2000,0,20,0"
	aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
	aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,15,0,0,0,0,15,0,15,15,0,15,0,0,0,15"
	aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
	aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,1," & N_START_YEAR & ",,,1,,,-1,-1," & N_START_YEAR & ",-1,," & N_START_YEAR & ",,-1"
	aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
	aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "10000000,10000000," & Year(Date()) & ",,,999999,,,-1,-1," & Year(Date()) + 1 & ",-1,,-1,,-1"
	aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
	aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",,-1,-1,,,1,-1,0,2,,0,,0"
	aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
	aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",,-1,-1,,,1,-1,0,2,,0,,0"
	aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
	aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞPaperworkSenders;,;SenderID;,;SenderID As RecordID, SenderName, EmployeeName, PositionName;,;(SenderID>-1);,;SenderID;,;;,;Ninguna;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞPaperworkTypes;,;PaperworkTypeID;,;PaperworkTypeID As RecordID, PaperworkTypeName;,;(Active=1);,;PaperworkTypeID;,;;,;Ninguno;;;-1ÞÞÞSubjectTypes;,;SubjectTypeID;,;SubjectTypeID As RecordID, SubjectTypeName;,;(Active=1);,;SubjectTypeID;,;;,;Ninguno;;;-1ÞÞÞÞÞÞPriorities;,;PriorityID;,;PriorityName;,;(Active=1);,;PriorityID;,;;,;Ninguna;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞStatusPaperworks;,;StatusID;,;StatusName;,;;,;StatusName;,;;,;Ninguno;;;-1"
	aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
	aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞ />&nbsp;<SPAN NAME=""SectionIDDiv"" ID=""SectionIDDiv""><SELECT NAME=""SectionID"" ID=""SectionIDCmb"" CLASS=""Lists"" onChange=""ChangePpwkNumber(this.value);"">" & GenerateListOptionsFromQuery(oADODBConnection, "PaperworkConsecutiveIDs", "CurrentID", "CurrentName", "", "OrderInList", "-1", "", "") & "</SELECT></SPAN><INPUT TYPE=""HIDDEN"" ÞÞÞÞÞÞÞÞÞ STYLE=""width: 0px"" /><INPUT TYPE=""TEXT"" NAME=""SenderName"" ID=""SenderNameTxt"" SIZE=""100"" VALUE="""" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.SenderName.value, 'PaperworkCatalogs&SenderIDs=1', 'SearchPpwkSendersIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Buscar procedencias"" BORDER=""0"" ALIGN=""ABSMIDDLE"" /></A><BR /><IFRAME SRC=""SearchRecord.asp"" NAME=""SearchPpwkSendersIFrame"" FRAMEBORDER=""0"" WIDTH=""1200"" HEIGHT=""26""></IFRAME><INPUT TYPE=""HIDDEN"" ÞÞÞ onChange=""SearchForRecord(this, 'EmployeeID&TableName=Employees&CodeField=EmployeeNumber', 'ControlFrm.EmployeeID');""ÞÞÞÞÞÞÞÞÞÞÞÞ STYLE=""width: 0px"" /><INPUT TYPE=""TEXT"" NAME=""SubjectTypeName"" ID=""SubjectTypeNameTxt"" SIZE=""100"" VALUE="""" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.SubjectTypeName.value, 'PaperworkCatalogs&SubjectTypeIDs=1&StartDate=' + document.CatalogFrm.StartDateYear.value + document.CatalogFrm.StartDateMonth.value + document.CatalogFrm.StartDateDay.value, 'SearchSubjectTypesIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Buscar tipos de asunto"" BORDER=""0"" ALIGN=""ABSMIDDLE"" /></A><BR /><IFRAME SRC=""SearchRecord.asp"" NAME=""SearchSubjectTypesIFrame"" FRAMEBORDER=""0"" WIDTH=""650"" HEIGHT=""26""></IFRAME><INPUT TYPE=""HIDDEN"" ÞÞÞ CheckEstimatedDate(); ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
	aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
	aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,8,9,16", ",")
	aCatalogComponent(N_ACTIVE_CATALOG) = -1
	aCatalogComponent(S_ADDITIONAL_FORM_SCRIPT_CATALOG) = "return CheckControlForm();"
	If Len(oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG))).Item) > 0 Then
		aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) = oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG))).Item
	End If
	aCatalogComponent(B_CHECK_FOR_DUPLICATED_CATALOG) = True
	aCatalogComponent(S_CANCEL_BUTTON_ACTION_CATALOG) = "window.location.href='EmployeeSupport.asp?Action=Paperworks&New=1'"
	lErrorNumber = GetConsecutiveID(oADODBConnection, 1061, aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1), sErrorDescription)

	InitializeSupportComponent = Err.number
	Err.Clear
End Function

Function AddPaperworkOwners(oRequest, oADODBConnection, lPaperworkID, sErrorDescription)
'************************************************************
'Purpose: To add the owners for the given paperwork
'Inputs:  oRequest, oADODBConnection, lPaperworkID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddPaperworkOwners"
	Dim asOwnerIDs
	Dim asActionIDs
	Dim asReportDates
	Dim asEndDates
	Dim asClosingNumber
	Dim iIndex
	Dim oItem
	Dim lErrorNumber

	sErrorDescription = "No se pudo eliminar la información del registro."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From PaperworkOwnersLKP Where (PaperworkID=" & lPaperworkID & ")", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	If lErrorNumber = 0 Then
		asOwnerIDs = Split(Replace(oRequest("OwnerIDs").Item, " ", ""), ",")
		asActionIDs = Split(Replace(oRequest("ActionIDs").Item, " ", ""), ",")
		asReportDates = Split(Replace(oRequest("ReportDates").Item, " ", ""), ",")
		asEndDates = Split(Replace(oRequest("EndDates").Item, " ", ""), ",")
		asClosingNumber = ""
		For Each oItem In oRequest("ClosingNumbers")
			asClosingNumber = asClosingNumber & oItem & LIST_SEPARATOR
		Next
		If Len(asClosingNumber) > 0 Then asClosingNumber = Left(asClosingNumber, (Len(asClosingNumber) - Len(LIST_SEPARATOR)))
		asClosingNumber = Split(asClosingNumber, LIST_SEPARATOR)
		If UBound(asClosingNumber) < UBound(asOwnerIDs) Then
			asClosingNumber = Join(asClosingNumber, LIST_SEPARATOR)
			asClosingNumber = Split(JoinLists(asClosingNumber, BuildList("0", LIST_SEPARATOR, UBound(asOwnerIDs) + 1), ","), LIST_SEPARATOR, -1, vbBinaryCompare)
		End If

		For iIndex = 0 To UBound(asOwnerIDs)
			If (Len(asOwnerIDs(iIndex)) > 0) And (Len(asActionIDs(iIndex)) > 0) And (Len(asReportDates(iIndex)) > 0) And (Len(asEndDates(iIndex)) > 0) Then
				sErrorDescription = "No se pudo agregar la información del registro."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PaperworkOwnersLKP (PaperworkID, OwnerID, PaperworkActionID, ReportDate, EndDate, ClosingNumber, Comments) Values (" & lPaperworkID & ", " & asOwnerIDs(iIndex) & ", " & asActionIDs(iIndex) & ", " & asReportDates(iIndex) & ", " & asEndDates(iIndex) & ", '" & asClosingNumber(iIndex) & "', '')", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
			End If
		Next
	End If

	AddPaperworkOwners = Err.number
	Err.Clear
End Function

Function DisplayClosePaperworksFrom(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the form to close the paperworks
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayClosePaperworksFrom"
	Dim sOwnerIDs
	Dim sCondition
	Dim iIndex
	Dim lErrorNumber

	sOwnerIDs = "-2"
	sCondition = ""
	sErrorDescription = "No se pudieron obtener los permisos del usuario."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select OwnerID From UsersOwnersLKP Where (UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ")", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Do While Not oRecordset.EOF
			sOwnerIDs = sOwnerIDs & "," & CStr(oRecordset.Fields("OwnerID").Value)
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
		oRecordset.Close
	End If
	If InStr(1, sOwnerIDs & ",", ",-1,", vbBinaryCompare) = 0 Then
		sErrorDescription = "No se pudieron obtener los permisos del usuario."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select OwnerID From PaperworkOwners Where (ParentID In (" & sOwnerIDs & ") And (OwnerID>-1))", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				sOwnerIDs = sOwnerIDs & "," & CStr(oRecordset.Fields("OwnerID").Value)
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
		End If
		sErrorDescription = "No se pudieron obtener los permisos del usuario."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select OwnerID From PaperworkOwners Where (ParentID In (" & sOwnerIDs & ") And (OwnerID>-1))", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				sOwnerIDs = sOwnerIDs & "," & CStr(oRecordset.Fields("OwnerID").Value)
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
		End If
		sCondition = "(OwnerID In (" & sOwnerIDs & ")) And (OwnerID>-1)"
	End If

	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		Response.Write "function CheckCloseFrm(oForm) {" & vbNewLine
			Response.Write "if (oForm) {" & vbNewLine
				Response.Write "if (oForm.PaperworkNumbers.options.length == 0) {" & vbNewLine
					Response.Write "alert('Favor de indicar los documentos a descargar');" & vbNewLine
					Response.Write "oForm.PaperworkNumberTemp.focus();" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine

				Response.Write "SelectAllItemsFromList(oForm.PaperworkNumbers);" & vbNewLine
				Response.Write "SelectAllItemsFromList(oForm.PaperworkYears);" & vbNewLine
				Response.Write "SelectAllItemsFromList(oForm.Owners);" & vbNewLine
				Response.Write "SelectAllItemsFromList(oForm.DocClassifications);" & vbNewLine
				Response.Write "SelectAllItemsFromList(oForm.Comments);" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "}" & vbNewLine
		Response.Write "} // End of CheckCloseFrm" & vbNewLine
	Response.Write "//--></SCRIPT>" & vbNewLine

	Response.Write "<FORM NAME=""CloseFrm"" ID=""CloseFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""POST"" onSubmit=""return CheckCloseFrm(this)"">"
		Response.Write "<TABLE WIDTH=""1000"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
			If Not bClosed Then
				Response.Write "<TD WIDTH=""1"" VALIGN=""TOP"">"
					Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
						Response.Write "<TR>"
							Response.Write "<TD WIDTH=""1""><FONT FACE=""Arial"" SIZE=""2"">No.&nbsp;de&nbsp;folio:&nbsp;</FONT></TD>"
							Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""PaperworkNumberTemp"" ID=""PaperworkNumberTempTxt"" SIZE=""10"" MAXLENGTH=""10"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Año:&nbsp;</FONT></TD>"
							Response.Write "<TD><SELECT NAME=""PaperworkYearTemp"" ID=""PaperworkYearTempCmb"" CLASS=""Lists"">"
								For iIndex = 2009 To Year(Date()) - 1
									Response.Write "<OPTION VALUE=""" & iIndex & """>" & iIndex & "</OPTION>"
								Next
								Response.Write "<OPTION VALUE=""" & Year(Date()) & """ SELECTED=""1"">" & Year(Date()) & "</OPTION>"
							Response.Write "</SELECT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Responsable:&nbsp;</FONT></TD>"
							Response.Write "<TD><SELECT NAME=""OwnerTemp"" ID=""OwnerTempCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "PaperworkOwners", "OwnerID", "OwnerID As RecordID, OwnerName, 'Empleado:' As Temp1, EmployeeID", sCondition, "OwnerID, OwnerName", "", "", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Oficio&nbsp;de&nbsp;descargo:&nbsp;</FONT></TD>"
							Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""DocClassificationTemp"" ID=""DocClassificationTempTxt"" SIZE=""30"" MAXLENGTH=""30"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Observaciones:&nbsp;</FONT></TD>"
							Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""CommentsTemp"" ID=""CommentsTempTxt"" SIZE=""50"" MAXLENGTH=""255"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
					Response.Write "</TABLE>"
				Response.Write "</TD>"
				Response.Write "<TD VALIGN=""TOP""><BR />"
					Response.Write "&nbsp;<A HREF=""javascript: AddPaperworkToClose()""><IMG SRC=""Images/BtnCrclAdd.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Turnar"" BORDER=""0"" /></A>&nbsp;"
					Response.Write "<BR /><BR />"
					Response.Write "&nbsp;<A HREF=""javascript: RemovePaperworkToClose()""><IMG SRC=""Images/BtnCrclDelete.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Remover"" BORDER=""0"" /></A>&nbsp;"
				Response.Write "</TD>"
			End If
			Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">No.&nbsp;de&nbsp;folio:<BR /></FONT>"
				Response.Write "<SELECT NAME=""PaperworkNumbers"" ID=""PaperworkNumbersLst"" SIZE=""3"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""SelectSameItemsForPaperworksToClose(this);""></SELECT>"
			Response.Write "</TD>"
			Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Año:<BR /></FONT>"
				Response.Write "<SELECT NAME=""PaperworkYears"" ID=""PaperworkYearsLst"" SIZE=""3"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""SelectSameItemsForPaperworksToClose(this);""></SELECT>"
			Response.Write "</TD>"
			Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Responsable:<BR /></FONT>"
				Response.Write "<SELECT NAME=""Owners"" ID=""OwnersLst"" SIZE=""3"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""SelectSameItemsForPaperworksToClose(this);""></SELECT>"
			Response.Write "</TD>"
			Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Oficio&nbsp;de&nbsp;descargo:<BR /></FONT>"
				Response.Write "<SELECT NAME=""DocClassifications"" ID=""DocClassificationsLst"" SIZE=""3"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""SelectSameItemsForPaperworksToClose(this);""></SELECT>"
			Response.Write "</TD>"
			Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Observaciones:<BR /></FONT>"
				Response.Write "<SELECT NAME=""Comments"" ID=""CommentsLst"" SIZE=""3"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""SelectSameItemsForPaperworksToClose(this);""></SELECT>"
			Response.Write "</TD>"
		Response.Write "</TR></TABLE><BR />"
		Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""DoClose"" ID=""DoCloseBtn"" VALUE=""Descargar"" CLASS=""Buttons"" />"
	Response.Write "</FORM><BR />"

	Response.Write "<FORM NAME=""CloseMultipleFrm"" ID=""CloseMultipleFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""POST"" onSubmit="""">"
		Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>O introduzca la información de los documentos a descargar, separando la información con tabuladores.</B><BR /></FONT>"
		Response.Write "<TABLE WIDTH=""1000"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
			Response.Write "<TD VALIGN=""TOP""><TEXTAREA NAME=""Ppwks"" ID=""PpwksTxtArea"" ROWS=""10"" COLS=""50"" CLASS=""TextFields""></TEXTAREA></TD>"
			Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><BR /><B>Formato:</B></FONT>"
				Response.Write "<FONT FACE=""Arial"" SIZE=""3""><PRE>"
					Response.Write "NO_FOLIO	AÑO	NO_RESPONSABLE	OFICIO_DESCARGO	OBSERVACIONES" & vbNewLine
					Response.Write "NO_FOLIO	AÑO	NO_RESPONSABLE	OFICIO_DESCARGO	OBSERVACIONES" & vbNewLine
					Response.Write "NO_FOLIO	AÑO	NO_RESPONSABLE	OFICIO_DESCARGO	OBSERVACIONES" & vbNewLine
				Response.Write "</PRE></FONT>"
			Response.Write "</TD>"
		Response.Write "</TR></TABLE><BR />"
		Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""DoClose"" ID=""DoCloseBtn"" VALUE=""Descargar"" CLASS=""Buttons"" />"
	Response.Write "</FORM>"

	DisplayClosePaperworksFrom = lErrorNumber
	Err.Clear
End Function

Function DisplayGuideSearchFrom(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the search form for the paperworks
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayGuideSearchFrom"
	Dim lErrorNumber

	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		Response.Write "function PrintGuide() {" & vbNewLine
			Response.Write "oForm = document.SearchFrm;" & vbNewLine
			Response.Write "if (oForm) {" & vbNewLine
				Response.Write "if (oForm.PaperworkNumber.value == '') {" & vbNewLine
					Response.Write "alert('Favor de especificar el número de folio');" & vbNewLine
					Response.Write "oForm.PaperworkNumber.focus();" & vbNewLine
				Response.Write "} else {" & vbNewLine
					Response.Write "OpenNewWindow('Export.asp?Action=Reports&Word=1&PaperworkID=' + oForm.PaperworkNumber.value + '&AddressID1=' + oForm.AddressID1.value + '&AddressID2=' + oForm.AddressID2.value + '&ReportID=1602&AccessKey=vac', '', 'ExportToExcel', 640, 480, 'yes', 'yes');" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "}" & vbNewLine
		Response.Write "} // End of PrintGuide" & vbNewLine
	Response.Write "//--></SCRIPT>" & vbNewLine
	Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""GET"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ForReport"" ID=""ForReportHdn"" VALUE=""1"" />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ForGuides"" ID=""ForGuidesHdn"" VALUE=""1"" />"
		Response.Write "<TABLE WIDTH=""700"" BORDER=""0"" CELLPADING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD WIDTH=""150""><FONT FACE=""Arial"" SIZE=""2"">Número&nbsp;de&nbsp;folio:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""PaperworkNumber"" ID=""PaperworkNumberTxt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & oRequest("PaperworkNumber").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Remitente:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""AddressID1"" ID=""AddressID1Cmb"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "PaperworkAddresses, States", "AddressID", "StateName, OwnerName, PositionName", "(PaperworkAddresses.StateID=States.StateID)", "StateName, AddressLevel", "", "", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Destinatario:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""AddressID2"" ID=""AddressID2Cmb"" SIZE=""1"" CLASS=""Lists"">"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "PaperworkAddresses, States", "AddressID", "StateName, OwnerName, PositionName", "(PaperworkAddresses.StateID=States.StateID)", "StateName, AddressLevel", "", "", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE><BR />"
		Response.Write "<INPUT TYPE=""BUTTON"" VALUE=""Imprimir"" CLASS=""Buttons"" onClick=""PrintGuide();"" />"
		Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
		Response.Write "<INPUT TYPE=""BUTTON"" VALUE=""Regresar"" CLASS=""RedButtons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=61';"" />"
	Response.Write "</FORM>"

	DisplayGuideSearchFrom = lErrorNumber
	Err.Clear
End Function

Function DisplayPaperworksSearchFrom(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the search form for the paperworks
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPaperworksSearchFrom"
	Dim sOwnerIDs
	Dim sCondition
	Dim oRecordset
	Dim lErrorNumber

	sCondition = ""
	Call GetPaperworksOwnersForUser(sOwnerIDs, sErrorDescription)
	If InStr(1, sOwnerIDs, "-1", vbBinaryCompare) = 0 Then sCondition = "(OwnerID In (" & sOwnerIDs & ")) And (OwnerID>-1)"

	Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""GET"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ForReport"" ID=""ForReportHdn"" VALUE=""" & oRequest("ForReport").Item & """ />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ForGuides"" ID=""ForGuidesHdn"" VALUE=""" & oRequest("ForGuides").Item & """ />"
		Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>BÚSQUEDA DE TRÁMITES</B><BR /></FONT>"
		Response.Write "<TABLE WIDTH=""700"" BORDER=""0"" CELLPADING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD WIDTH=""150""><FONT FACE=""Arial"" SIZE=""2"">Número&nbsp;de&nbsp;folio:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entre <INPUT TYPE=""TEXT"" NAME=""FilterStartNumber"" ID=""FilterStartNumberTxt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & oRequest("FilterStartNumber").Item & """ CLASS=""TextFields"" /> y <INPUT TYPE=""TEXT"" NAME=""FilterEndNumber"" ID=""FilterEndNumberTxt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & oRequest("FilterEndNumber").Item & """ CLASS=""TextFields"" /></FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha del documento:&nbsp;</FONT></TD>"
				Response.Write "<TD>"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("StartStartYear").Item), CInt(oRequest("StartStartMonth").Item), CInt(oRequest("StartStartDay").Item), "StartStartYear", "StartStartMonth", "StartStartDay", N_FORM_START_YEAR, Year(Date()), True, True)
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("EndStartYear").Item), CInt(oRequest("EndStartMonth").Item), CInt(oRequest("EndStartDay").Item), "EndStartYear", "EndStartMonth", "EndStartDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Documento:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""FilterDocumentNumber"" ID=""FilterDocumentNumberTxt"" SIZE=""30"" MAXLENGTH=""50"" VALUE=""" & oRequest("FilterDocumentNumber").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Procedencia:&nbsp;</FONT></TD>"
				Response.Write "<TD>"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SenderID"" ID=""SenderIDHdn"" />"
					Response.Write "<INPUT TYPE=""TEXT"" NAME=""SenderName"" ID=""SenderNameTxt"" SIZE=""100"" VALUE="""" />"
					Response.Write "<A HREF=""javascript: SearchRecord(document.SearchFrm.SenderName.value, 'PaperworkCatalogs&SenderIDs=1', 'SearchPpwkSendersIFrame', 'SearchFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Buscar procedencias"" BORDER=""0"" ALIGN=""ABSMIDDLE"" /></A><BR />"
					Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchPpwkSendersIFrame"" FRAMEBORDER=""0"" WIDTH=""1200"" HEIGHT=""26""></IFRAME>"
'						Response.Write "<OPTION VALUE="""">Todos</OPTION>"
'						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "PaperworkSenders", "SenderID", "SenderID As RecordID, SenderName, EmployeeName, PositionName", "", "SenderID", "", "", sErrorDescription)
'					Response.Write "</SELECT>"
				Response.Write "</TD>"
			Response.Write "</TR>"
'			Response.Write "<TR>"
'				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número&nbsp;de&nbsp;empleado:&nbsp;</FONT></TD>"
'				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""FilterEmployeeID"" ID=""FilterEmployeeIDTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE=""" & oRequest("FilterEmployeeID").Item & """ CLASS=""TextFields"" /></TD>"
'			Response.Write "</TR>"
'			Response.Write "<TR>"
'				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Desc. Procedencia:&nbsp;</FONT></TD>"
'				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""FilterDescription"" ID=""FilterDescriptionTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("FilterDocumentSubject").Item & """ CLASS=""TextFields"" /></TD>"
'			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Asunto:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""FilterDocumentSubject"" ID=""FilterDocumentSubjectTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("FilterDocumentSubject").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Responsable:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""FilterOwnerID"" ID=""FilterOwnerIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "PaperworkOwners", "OwnerID", "OwnerID As RecordID, OwnerName, EmployeeID", sCondition, "OwnerID", oRequest("FilterOwnerID").Item, "", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de trámite:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""FilterPaperworkTypeID"" ID=""FilterPaperworkTypeIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "PaperworkTypes", "PaperworkTypeID", "PaperworkTypeName", "", "PaperworkTypeName", oRequest("FilterPaperworkTypeID").Item, "Ninguno;;;-1", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Tipo de asunto:&nbsp;</FONT></TD>"
				Response.Write "<TD>"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SubjectTypeID"" ID=""SubjectTypeIDHdn"" />"
					Response.Write "<INPUT TYPE=""TEXT"" NAME=""SubjectTypeName"" ID=""SubjectTypeNameTxt"" SIZE=""100"" VALUE="""" />"
					Response.Write "<A HREF=""javascript: SearchRecord(document.SearchFrm.SubjectTypeName.value, 'PaperworkCatalogs&SubjectTypeIDs=1&StartDate=-1', 'SearchSubjectTypesIFrame', 'SearchFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Buscar tipos de asunto"" BORDER=""0"" ALIGN=""ABSMIDDLE"" /></A><BR />"
					Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchSubjectTypesIFrame"" FRAMEBORDER=""0"" WIDTH=""650"" HEIGHT=""26""></IFRAME>"
'						Response.Write "<OPTION VALUE="""">Todos</OPTION>"
'						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "SubjectTypes", "SubjectTypeID", "SubjectTypeID As RecordID, SubjectTypeName", "", "SubjectTypeID", oRequest("FilterSubjectTypeID").Item, "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT>"
				Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Observaciones:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""FilterComments"" ID=""FilterCommentsTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("FilterDocumentSubject").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Estatus:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""FilterStatusID"" ID=""FilterStatusIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "StatusPaperworks", "StatusID", "StatusName", "(StatusID In (0,3))", "StatusName", oRequest("FilterStatusID").Item, "Ninguno;;;-1", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE><BR />"
		Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""DoSearch"" ID=""DoSearchBtn"" VALUE=""Buscar Trámites"" CLASS=""Buttons"" />"
	Response.Write "</FORM>"
	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		If (Len(oRequest("StartStartYear").Item) = 0) And (Len(oRequest("StartStartMonth").Item) = 0) And (Len(oRequest("StartStartDay").Item) = 0) And (Len(oRequest("EndStartYear").Item) = 0) And (Len(oRequest("EndStartMonth").Item) = 0) And (Len(oRequest("EndStartDay").Item) = 0) Then
			Response.Write "SendURLValuesToForm('StartStartYear=" & Year(Date()) & "&StartStartMonth=01&StartStartDay=01&EndStartYear=" & Year(Date()) & "&EndStartMonth=12&EndStartDay=31', document.SearchFrm);" & vbNewLine
		End If
	Response.Write "//--></SCRIPT>" & vbNewLine

	DisplayPaperworksSearchFrom = lErrorNumber
	Err.Clear
End Function

Function DisplayOwnersInCatalogForm(oRequest, oADODBConnection, lPaperworkID, bClosed, sErrorDescription)
'************************************************************
'Purpose: To display the owners for the given paperworks
'		  using HTML lists
'Inputs:  oRequest, oADODBConnection, lPaperworkID, bClosed
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayOwnersInCatalogForm"
	Dim sOwnerIDs
	Dim sCondition
	Dim oRecordset
	Dim lErrorNumber

	sOwnerIDs = "-2"
	sCondition = ""
	sErrorDescription = "No se pudieron obtener los permisos del usuario."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select OwnerID From UsersOwnersLKP Where (UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ")", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Do While Not oRecordset.EOF
			sOwnerIDs = sOwnerIDs & "," & CStr(oRecordset.Fields("OwnerID").Value)
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
		oRecordset.Close
	End If
	If InStr(1, sOwnerIDs & ",", ",-1,", vbBinaryCompare) = 0 Then
		sErrorDescription = "No se pudieron obtener los permisos del usuario."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select OwnerID From PaperworkOwners Where (ParentID In (" & sOwnerIDs & ")) And (OwnerID>-1)", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				sOwnerIDs = sOwnerIDs & "," & CStr(oRecordset.Fields("OwnerID").Value)
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
		End If
		sErrorDescription = "No se pudieron obtener los permisos del usuario."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select OwnerID From PaperworkOwners Where (ParentID In (" & sOwnerIDs & ")) And (OwnerID>-1)", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				sOwnerIDs = sOwnerIDs & "," & CStr(oRecordset.Fields("OwnerID").Value)
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
		End If
		sCondition = " And (OwnerID In (" & sOwnerIDs & ")) And (OwnerID>-1)"
	End If

	Call DisplayTimeStamp(sCondition)
	aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = "<IMG SRC=""Images/DotBlue.gif"" WIDTH=""700"" HEIGHT=""1"" /><BR /><BR /><IFRAME SRC=""SearchRecord.asp"" NAME=""SearchPpwkOwnersIFrame"" FRAMEBORDER=""0"" WIDTH=""0"" HEIGHT=""0""></IFRAME>"
	aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<TABLE WIDTH=""800"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
		If Not bClosed Then
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<TD VALIGN=""TOP"">"
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<FONT FACE=""Arial"" SIZE=""2"">Responsable:&nbsp;<INPUT TYPE=""TEXT"" NAME=""OwnerIDToSearch"" ID=""OwnerIDToSearchTxt"" SIZE=""6"" MAXLENGTH=""4"" VALUE="""" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.OwnerIDToSearch.value, 'PaperworkCatalogs&OwnerIDs=1', 'SearchPpwkOwnersIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Buscar responsables"" BORDER=""0"" ALIGN=""ABSMIDDLE"" /></A><BR /></FONT>"
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<SELECT NAME=""OwnerIDTemp"" ID=""OwnerIDTempCmb"" SIZE=""1"" CLASS=""Lists"">"
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & GenerateListOptionsFromQuery(oADODBConnection, "PaperworkOwners", "OwnerID", "OwnerID As RecordID, OwnerName, 'Empleado:' As Temp1, EmployeeID", "(OwnerID>-1) And (OwnerID Not In (Select OwnerID From PaperworkOwnersLKP Where (PaperworkID=" & lPaperworkID & ")))" & sCondition, "OwnerID, OwnerName", "", "", sErrorDescription)
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "</SELECT><BR />"
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<FONT FACE=""Arial"" SIZE=""2"">Acción:<BR /></FONT>"
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<SELECT NAME=""ActionIDTemp"" ID=""ActionIDTempCmb"" SIZE=""1"" CLASS=""Lists"">"
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & GenerateListOptionsFromQuery(oADODBConnection, "PaperworkActions", "PaperworkActionID", "PaperworkActionShortName, PaperworkActionName", "(Active=1)", "PaperworkActionShortName, PaperworkActionName", "4", "", sErrorDescription)
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "</SELECT><BR />"
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "</TD>"
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<TD VALIGN=""TOP""><BR /><BR />"
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "&nbsp;<A HREF=""javascript: AddOwnerComment()""><IMG SRC=""Images/BtnCrclAdd.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Turnar"" BORDER=""0"" /></A>&nbsp;"
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<BR /><BR />"
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "&nbsp;<A HREF=""javascript: RemoveOwnerComment()""><IMG SRC=""Images/BtnCrclDelete.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Remover"" BORDER=""0"" /></A>&nbsp;"
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "</TD>"
		End If

		sErrorDescription = "No se pudieron obtener los responsables del documento."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaperworkOwnersLKP.*, OwnerName, PaperworkOwners.EmployeeID, '.' As EmployeeName, '.' As EmployeeLastName, '.' As EmployeeLastName2, PaperworkActionShortName, PaperworkActionName From PaperworkOwnersLKP, PaperworkOwners, PaperworkActions Where (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (PaperworkOwnersLKP.PaperworkActionID=PaperworkActions.PaperworkActionID) And (PaperworkOwnersLKP.PaperworkID=" & lPaperworkID & ") And (PaperworkOwners.OwnerID>-1) Order By ReportDate, PaperworkOwnersLKP.OwnerID", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Responsables:<BR /></FONT>"
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<SELECT NAME=""OwnerIDs"" ID=""OwnerIDsLst"" SIZE=""3"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""SelectSameItemsForOwners(this);"">"
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						'aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<OPTION VALUE=""" & CStr(oRecordset.Fields("OwnerID").Value) & """>" & CStr(oRecordset.Fields("OwnerID").Value) & " " & CStr(oRecordset.Fields("OwnerName").Value) & ". " & CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & "</OPTION>"
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<OPTION VALUE=""" & CStr(oRecordset.Fields("OwnerID").Value) & """>" & CStr(oRecordset.Fields("OwnerID").Value) & " " & CStr(oRecordset.Fields("OwnerName").Value) & ". EMPLEADO: " & CStr(oRecordset.Fields("EmployeeID").Value) & "</OPTION>"
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "</SELECT>"
		aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "</TD>"

		sErrorDescription = "No se pudieron obtener los responsables del documento."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaperworkOwnersLKP.*, OwnerName, PaperworkOwners.EmployeeID, '.' As EmployeeName, '.' As EmployeeLastName, '.' As EmployeeLastName2, PaperworkActionShortName, PaperworkActionName From PaperworkOwnersLKP, PaperworkOwners, PaperworkActions Where (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (PaperworkOwnersLKP.PaperworkActionID=PaperworkActions.PaperworkActionID) And (PaperworkOwnersLKP.PaperworkID=" & lPaperworkID & ") And (PaperworkOwners.OwnerID>-1) Order By ReportDate, PaperworkOwnersLKP.OwnerID", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Acciones:<BR /></FONT>"
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<SELECT NAME=""ActionIDs"" ID=""ActionIDsLst"" SIZE=""3"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""SelectSameItemsForOwners(this);"">"
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<OPTION VALUE=""" & CStr(oRecordset.Fields("PaperworkActionID").Value) & """>" & CStr(oRecordset.Fields("PaperworkActionShortName").Value) & ". " & CStr(oRecordset.Fields("PaperworkActionName").Value) & "</OPTION>"
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "</SELECT>"
		aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "</TD>"

		sErrorDescription = "No se pudieron obtener los responsables del documento."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaperworkOwnersLKP.*, OwnerName, PaperworkOwners.EmployeeID, '.' As EmployeeName, '.' As EmployeeLastName, '.' As EmployeeLastName2, PaperworkActionShortName, PaperworkActionName From PaperworkOwnersLKP, PaperworkOwners, PaperworkActions Where (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (PaperworkOwnersLKP.PaperworkActionID=PaperworkActions.PaperworkActionID) And (PaperworkOwnersLKP.PaperworkID=" & lPaperworkID & ") And (PaperworkOwners.OwnerID>-1) Order By ReportDate, PaperworkOwnersLKP.OwnerID", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fecha&nbsp;de&nbsp;turnado:<BR /></FONT>"
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<SELECT NAME=""ReportDates"" ID=""ReportDatesLst"" SIZE=""3"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""SelectSameItemsForOwners(this);"">"
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<OPTION VALUE=""" & CStr(oRecordset.Fields("ReportDate").Value) & """>" & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("ReportDate").Value), -1, -1, -1) & "</OPTION>"
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "</SELECT>"
		aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "</TD>"

		sErrorDescription = "No se pudieron obtener los responsables del documento."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaperworkOwnersLKP.*, OwnerName, PaperworkOwners.EmployeeID, '.' As EmployeeName, '.' As EmployeeLastName, '.' As EmployeeLastName2, PaperworkActionShortName, PaperworkActionName From PaperworkOwnersLKP, PaperworkOwners, PaperworkActions Where (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (PaperworkOwnersLKP.PaperworkActionID=PaperworkActions.PaperworkActionID) And (PaperworkOwnersLKP.PaperworkID=" & lPaperworkID & ") And (PaperworkOwners.OwnerID>-1) Order By ReportDate, PaperworkOwnersLKP.OwnerID", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fecha&nbsp;de&nbsp;cierre:<BR /></FONT>"
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<SELECT NAME=""EndDates"" ID=""EndDatesLst"" SIZE=""3"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""SelectSameItemsForOwners(this);"">"
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<OPTION VALUE=""" & oRecordset.Fields("EndDate").Value & """>"
							If CLng(oRecordset.Fields("EndDate").Value) = 0 Then
								aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "---"
							Else
								aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1)
							End If
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "</OPTION>"
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "</SELECT>"
		aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "</TD>"

		sErrorDescription = "No se pudieron obtener los responsables del documento."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaperworkOwnersLKP.*, OwnerName, PaperworkOwners.EmployeeID, '.' As EmployeeName, '.' As EmployeeLastName, '.' As EmployeeLastName2, PaperworkActionShortName, PaperworkActionName From PaperworkOwnersLKP, PaperworkOwners, PaperworkActions Where (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (PaperworkOwnersLKP.PaperworkActionID=PaperworkActions.PaperworkActionID) And (PaperworkOwnersLKP.PaperworkID=" & lPaperworkID & ") And (PaperworkOwners.OwnerID>-1) Order By ReportDate, PaperworkOwnersLKP.OwnerID", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Oficio&nbsp;de&nbsp;descargo:<BR /></FONT>"
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<SELECT NAME=""ClosingNumbers"" ID=""ClosingNumbersLst"" SIZE=""3"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""SelectSameItemsForOwners(this);"">"
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<OPTION VALUE=""" & oRecordset.Fields("ClosingNumber").Value & """>"
							aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & CStr(oRecordset.Fields("ClosingNumber").Value)
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "</OPTION>"
						oRecordset.MoveNext
						'If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "</SELECT>"
		aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "</TD>"

		sErrorDescription = "No se pudieron obtener los responsables del documento."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaperworkOwnersLKP.*, OwnerName, PaperworkOwners.EmployeeID, '.' As EmployeeName, '.' As EmployeeLastName, '.' As EmployeeLastName2, PaperworkActionShortName, PaperworkActionName From PaperworkOwnersLKP, PaperworkOwners, PaperworkActions Where (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (PaperworkOwnersLKP.PaperworkActionID=PaperworkActions.PaperworkActionID) And (PaperworkOwnersLKP.PaperworkID=" & lPaperworkID & ") And (PaperworkOwners.OwnerID>-1) Order By ReportDate, PaperworkOwnersLKP.OwnerID", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Asunto&nbsp;de&nbsp;descargo:<BR /></FONT>"
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<SELECT NAME=""OwnersComments"" ID=""OwnersCommentsLst"" SIZE=""3"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""SelectSameItemsForOwners(this);"">"
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<OPTION VALUE="""">"
							aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & CStr(oRecordset.Fields("Comments").Value)
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "</OPTION>"
						oRecordset.MoveNext
						'If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "</SELECT>"
		aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "</TD>"
	aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "</TR></TABLE><BR />"

'	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
'		Response.Write "document.all['ExtraHTMLForCatalogFrmDiv'].innerHTML = '" & sHTML & "';" & vbNewLine
'	Response.Write "//--></SCRIPT>" & vbNewLine

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayOwnersInCatalogForm = lErrorNumber
	Err.Clear
End Function

Function DisplayPaperworksForSupportTable(oRequest, oADODBConnection, bUseLinks, bForExport, sCondition, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the paperworks from
'		  the database in a table
'Inputs:  oRequest, oADODBConnection, bUseLinks, bForExport, sCondition
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPaperworksForSupportTable"
	Dim iIndex
	Dim sTables
	Dim oRecordset
	Dim sBoldBegin
	Dim sBoldEnd
	Dim iRecordCounter
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	sTables = ""
	If InStr(1, sCondition, "Jobs", vbBinaryCompare) > 0 Then sTables = sTables & ", Jobs, Areas"

	If Len(oRequest("FullSearch").Item) > 0 Then sTables = sTables & ", PaperworkOwners As PaperworkOwners2, PaperworkOwners As PaperworkOwners1"
	sErrorDescription = "No se pudo obtener la información del empleado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct Paperworks.*, PaperworkSenders.SenderID, SenderName, EmployeeName, PositionName, PaperworkOwners.OwnerID, PaperworkOwners.OwnerName, PaperworkOwners.EmployeeID, PaperworkOwnersLKP.EndDate, PaperworkTypeName, StatusName From Paperworks, PaperworkSenders, PaperworkOwnersLKP, PaperworkOwners, PaperworkTypes, StatusPaperworks" & sTables & " Where (Paperworks.SenderID=PaperworkSenders.SenderID) And (Paperworks.PaperworkID=PaperworkOwnersLKP.PaperworkID) And (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (Paperworks.PaperworkTypeID=PaperworkTypes.PaperworkTypeID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (PaperworkOwners.OwnerID>-1) " & sCondition & " Order By PaperworkNumber", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select Distinct Paperworks.*, PaperworkSenders.SenderID, SenderName, EmployeeName, PositionName, PaperworkOwners.OwnerID, PaperworkOwners.OwnerName, PaperworkOwners.EmployeeID, PaperworkOwnersLKP.EndDate, PaperworkTypeName, StatusName From Paperworks, PaperworkSenders, PaperworkOwnersLKP, PaperworkOwners, PaperworkTypes, StatusPaperworks" & sTables & " Where (Paperworks.SenderID=PaperworkSenders.SenderID) And (Paperworks.PaperworkID=PaperworkOwnersLKP.PaperworkID) And (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (Paperworks.PaperworkTypeID=PaperworkTypes.PaperworkTypeID) And (Paperworks.StatusID=StatusPaperworks.StatusID) And (PaperworkOwners.OwnerID>-1) " & sCondition & " Order By PaperworkNumber -->" & vbNewLine
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
				If bUseLinks And Not bForExport Then
					asColumnsTitles = Split("Acciones,No. de folio,Fecha del documento,Documento,Procedencia,Responsable,Asunto,Tipo de trámite,Estatus", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,100,200,100,100,200,200,100", ",", -1, vbBinaryCompare)
					asCellAlignments = Split("CENTER,,,,,,,,", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("No. de folio,Fecha del documento,Documento,Procedencia,Responsable,Asunto,Tipo de trámite,Estatus", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,200,100,100,200,200,100", ",", -1, vbBinaryCompare)
					asCellAlignments = Split(",,,,,,,", ",", -1, vbBinaryCompare)
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

				Do While Not oRecordset.EOF
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("PaperworkID").Value), oRequest("PaperworkID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sRowContents = ""
					If bUseLinks And Not bForExport Then
						sRowContents = sRowContents & "<NOBR>&nbsp;"
							If ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) And (Len(oRequest("ForReport").Item) = 0) Then
								sRowContents = sRowContents & "<A HREF=""EmployeeSupport.asp?PaperworkID=" & CStr(oRecordset.Fields("PaperworkID").Value) & "&Change=1"">"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
							End If

							sRowContents = sRowContents & "<A HREF=""javascript: OpenNewWindow('Export.asp?Action=Reports&Word=1&PaperworkID=" & CStr(oRecordset.Fields("PaperworkID").Value) & "&ReportID="
							If Len(oRequest("ForGuides").Item) > 0 Then
								sRowContents = sRowContents & "1602"
							Else
								sRowContents = sRowContents & "1600"
							End If
							sRowContents = sRowContents & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/IcnForm.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Imprimir"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"

							If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) And (Len(oRequest("ForReport").Item) = 0) Then
								sRowContents = sRowContents & "<A HREF=""EmployeeSupport.asp?PaperworkID=" & CStr(oRecordset.Fields("PaperworkID").Value) & "&Delete=1"">"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
							End If
						sRowContents = sRowContents & "&nbsp;</NOBR>"
						sRowContents = sRowContents & TABLE_SEPARATOR
					End If

					sRowContents = sRowContents & "<A"
						If (Not bForExport) And (Len(oRequest("ForReport").Item) = 0) Then sRowContents = sRowContents & " HREF=""EmployeeSupport.asp?PaperworkID=" & CStr(oRecordset.Fields("PaperworkID").Value) & "&Change=1"""
					sRowContents = sRowContents & ">" & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkNumber").Value)) & "</A>" & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & DisplayDateFromSerialNumber(CStr(oRecordset.Fields("StartDate").Value), -1, -1, -1) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("DocumentNumber").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("SenderID").Value) & ". " & CStr(oRecordset.Fields("SenderName").Value) & ". " & CStr(oRecordset.Fields("EmployeeName").Value) & " (" & CStr(oRecordset.Fields("PositionName").Value) & ")") & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("OwnerID").Value) & ". " & CStr(oRecordset.Fields("OwnerName").Value) & ". Empleado: " & CStr(oRecordset.Fields("EmployeeID").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CStr(oRecordset.Fields("DocumentSubject").Value) & sBoldEnd
					sRowContents = sRowContents & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkTypeName").Value)) & sBoldEnd
					If CLng(oRecordset.Fields("EndDate").Value) = 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value)) & sBoldEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & "Cerrado" & sBoldEnd
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
			Response.Write "</TABLE></DIV>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros que cumplan con los criterios de la búsqueda."
		End If
	End If

	DisplayPaperworksForSupportTable = lErrorNumber
	Err.Clear
End Function

Function PrintPaperwork(oRequest, oADODBConnection, lPaperworkID, lOwnerID, sErrorDescription)
'************************************************************
'Purpose: To print the paperwork information using a template
'Inputs:  oRequest, oADODBConnection, lPaperworkID, lOwnerID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "PrintPaperwork"
	Dim oRecordset
	Dim sContents
	Dim lEmployeeID
	Dim iIndex
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información del empleado."
'	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Paperworks.*, PaperworkSenders.SenderName As SenderAreaName, PaperworkSenders.PositionName As SenderPositionName, PaperworkSenders.EmployeeName As SenderName1, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, PaperworkTypeName, Owners.EmployeeName As OwnerName1, Owners.EmployeeLastName As OwnerLastName, Owners.EmployeeLastName2 As OwnerLastName2, PaperworkActionID From Paperworks, PaperworkSenders, Employees, PaperworkTypes, PaperworkOwnersLKP, PaperworkOwners, Employees As Owners Where (Paperworks.SenderID=PaperworkSenders.SenderID) And (Paperworks.OwnerID=Employees.EmployeeID) And (Paperworks.PaperworkTypeID=PaperworkTypes.PaperworkTypeID) And (Paperworks.PaperworkID=PaperworkOwnersLKP.PaperworkID) And (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (PaperworkOwners.EmployeeID=Owners.EmployeeID) And (Paperworks.PaperworkID=" & lPaperworkID & ") And (PaperworkOwners.OwnerID>-1) Order By PaperworkOwnersLKP.OwnerID", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Paperworks.*, PaperworkSenders.SenderName As SenderAreaName, PaperworkSenders.PositionName As SenderPositionName, PaperworkSenders.EmployeeName As SenderName1, PaperworkTypeName From Paperworks, PaperworkSenders, PaperworkTypes Where (Paperworks.SenderID=PaperworkSenders.SenderID) And (Paperworks.PaperworkTypeID=PaperworkTypes.PaperworkTypeID) And (Paperworks.PaperworkID=" & lPaperworkID & ")", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If FileExists(Server.MapPath(TEMPLATES_PHYSICAL_PATH & "EmployeeSupport.htm"), sErrorDescription) Then
				sContents = GetFileContents(Server.MapPath(TEMPLATES_PHYSICAL_PATH & "EmployeeSupport.htm"), sErrorDescription)
				If Len(sContents) > 0 Then
					sContents = Replace(sContents, "<PAPERWORK_NUMBER />", CleanStringForHTML(CStr(oRecordset.Fields("PaperworkNumber").Value)))
					sContents = Replace(sContents, "<START_DATE />", DisplayDateFromSerialNumber(CStr(oRecordset.Fields("StartDate").Value), -1, -1, -1))
					sContents = Replace(sContents, "<ESTIMATED_DATE />", DisplayDateFromSerialNumber(CStr(oRecordset.Fields("EstimatedDate").Value), -1, -1, -1))
					sContents = Replace(sContents, "<DOCUMENT_NUMBER />", CleanStringForHTML(CStr(oRecordset.Fields("DocumentNumber").Value)))
					sContents = Replace(sContents, "<EMPLOYEE_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("SenderName1").Value)))
					sContents = Replace(sContents, "<SENDER_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("SenderPositionName").Value) & ". " & CStr(oRecordset.Fields("SenderAreaName").Value)))
					sContents = Replace(sContents, "<SUBJECT />", CleanStringForHTML(CStr(oRecordset.Fields("DocumentSubject").Value)))
'					sContents = Replace(sContents, "<DESCRIPTION />", CleanStringForHTML(CStr(oRecordset.Fields("Description").Value)))
					sContents = Replace(sContents, "<COMMENTS />", CleanStringForHTML(CStr(oRecordset.Fields("Comments").Value)))
					sContents = Replace(sContents, "<PAPERWORK_TYPE_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("PaperworkTypeName").Value)))
					lEmployeeID = CLng(oRecordset.Fields("OwnerID").Value)
					oRecordset.Close

'					sErrorDescription = "No se pudo obtener la información del empleado."
'					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeName, EmployeeLastName, EmployeeLastName2 From Employees Where (EmployeeID=" & lEmployeeID & ")", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
'					If lErrorNumber = 0 Then
'						If Not oRecordset.EOF Then
'							sContents = Replace(sContents, "<EMPLOYEE_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value)))
'						Else
'							sContents = Replace(sContents, "<EMPLOYEE_NAME />", "----------")
'						End If
'					End If
'					oRecordset.Close

					sErrorDescription = "No se pudo obtener la información del empleado."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeName, EmployeeLastName, EmployeeLastName2, PaperworkActionID, PositionName From PaperworkOwnersLKP, PaperworkOwners, Employees, Jobs, Positions Where (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (PaperworkOwners.EmployeeID=Employees.EmployeeID) And (Employees.JobID=Jobs.JobID) And (Jobs.PositionID=Positions.PositionID) And (PaperworkOwnersLKP.PaperworkID=" & lPaperworkID & ") And (PaperworkOwners.OwnerID>-1) Order By PaperworkOwnersLKP.OwnerID", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							sContents = Replace(sContents, "<PAPERWORK_ACTION_ID_" & CStr(oRecordset.Fields("PaperworkActionID").Value) & " />", "<B>X</B>")
							'Do While Not oRecordset.EOF
								sContents = Replace(sContents, "<OWNER_NAME_1 />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & "<BR /><BR />" & CStr(oRecordset.Fields("PositionName").Value)))
								oRecordset.MoveNext
								If Not oRecordset.EOF Then
									sContents = Replace(sContents, "<OWNER_NAME_2 />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & " (" & CStr(oRecordset.Fields("PositionName").Value) & ")"))
									oRecordset.MoveNext
									If Not oRecordset.EOF Then
										sContents = Replace(sContents, "<OWNER_NAME_3 />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & " (" & CStr(oRecordset.Fields("PositionName").Value) & ")"))
										oRecordset.MoveNext
									End If
								End If
							'Loop
						End If
						oRecordset.Close
					End If
					For iIndex = 0 To 100
						sContents = Replace(sContents, "<PAPERWORK_ACTION_ID_" & iIndex & " />", "&nbsp;")
					Next

					Response.Write sContents
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
	PrintPaperwork = lErrorNumber
	Err.Clear
End Function

Function PrintPaperworkGuide(oRequest, oADODBConnection, lPaperworkID, lAddressID1, lAddressID2, sErrorDescription)
'************************************************************
'Purpose: To print the paperwork information using a template
'Inputs:  oRequest, oADODBConnection, lPaperworkID, lAddressID1, lAddressID2
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "PrintPaperworkGuide"
	Dim oRecordset
	Dim sContents
	Dim sTemp
	Dim asZones
	Dim sNames
	Dim lErrorNumber

	If FileExists(Server.MapPath(TEMPLATES_PHYSICAL_PATH & "EmployeeSupportGuide.htm"), sErrorDescription) Then
		sContents = GetFileContents(Server.MapPath(TEMPLATES_PHYSICAL_PATH & "EmployeeSupportGuide.htm"), sErrorDescription)
		If Len(sContents) > 0 Then
			sContents = Replace(sContents, "<CURRENT_DATE />", Right(("0" & Day(Date())), Len("00")) & "/" & Right(("0" & Month(Date())), Len("00")) & "/" & Year(Date()))
			sContents = Replace(sContents, "<PAPERWORK_NUMBER />", lPaperworkID)

			sErrorDescription = "No se pudo obtener la información del remitente."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaperworkAddresses.*, StateName From PaperworkAddresses, States Where (PaperworkAddresses.StateID=States.StateID) And (AddressID=" & lAddressID1 & ")", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sContents = Replace(sContents, "<OWNER_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("OwnerName").Value)))
					sContents = Replace(sContents, "<POSITION_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value)))
					sContents = Replace(sContents, "<OWNER_ADDRESS />", CleanStringForHTML(CStr(oRecordset.Fields("OwnerAddress").Value)))
					sContents = Replace(sContents, "<OWNER_ADDRESS2 />", CleanStringForHTML(CStr(oRecordset.Fields("OwnerAddress2").Value)))
					sContents = Replace(sContents, "<OWNER_CITY />", CleanStringForHTML(CStr(oRecordset.Fields("OwnerCity").Value)))
					sContents = Replace(sContents, "<OWNER_ZIP_CODE />", CleanStringForHTML(CStr(oRecordset.Fields("OwnerZipCode").Value)))
					sContents = Replace(sContents, "<STATE_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("StateName").Value)))
					sTemp = ""
					sTemp = CStr(oRecordset.Fields("OwnerPhone").Value)
					sContents = Replace(sContents, "<OWNER_PHONE />", CleanStringForHTML(sTemp))
					oRecordset.Close
				End If
			End If

			sErrorDescription = "No se pudo obtener la información del destinatario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaperworkAddresses.*, StateName From PaperworkAddresses, States Where (PaperworkAddresses.StateID=States.StateID) And (AddressID=" & lAddressID2 & ")", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sContents = Replace(sContents, "<OWNER_NAME_2 />", CleanStringForHTML(CStr(oRecordset.Fields("OwnerName").Value)))
					sContents = Replace(sContents, "<POSITION_NAME_2 />", CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value)))
					sContents = Replace(sContents, "<OWNER_ADDRESS_2 />", CleanStringForHTML(CStr(oRecordset.Fields("OwnerAddress").Value)))
					sContents = Replace(sContents, "<OWNER_ADDRESS2_2 />", CleanStringForHTML(CStr(oRecordset.Fields("OwnerAddress2").Value)))
					sContents = Replace(sContents, "<OWNER_CITY_2 />", CleanStringForHTML(CStr(oRecordset.Fields("OwnerCity").Value)))
					sContents = Replace(sContents, "<OWNER_ZIP_CODE_2 />", CleanStringForHTML(CStr(oRecordset.Fields("OwnerZipCode").Value)))
					sContents = Replace(sContents, "<STATE_NAME_2 />", CleanStringForHTML(CStr(oRecordset.Fields("StateName").Value)))
					sTemp = ""
					sTemp = CStr(oRecordset.Fields("OwnerPhone").Value)
					sContents = Replace(sContents, "<OWNER_PHONE_2 />", CleanStringForHTML(sTemp))
					oRecordset.Close
				End If
			End If
			Response.Write sContents
		End If
	End If

	Set oRecordset = Nothing
	PrintPaperworkGuide = lErrorNumber
	Err.Clear
End Function

Function PrintPaperworkList(oRequest, oADODBConnection, lListID, sErrorDescription)
'************************************************************
'Purpose: To print the paperwork list
'Inputs:  oRequest, oADODBConnection, lListID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "PrintPaperworkList"
	Dim sPaperworkIDs
	Dim sListNumber
	Dim sSenderName
	Dim sRecipientName
	Dim iCounter
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información del empleado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ListNumber, SenderName, RecipientName, PaperworkIDs From PaperworkLists Where (ListID=" & lListID & ")", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sPaperworkIDs = CStr(oRecordset.Fields("PaperworkIDs").Value)
			sListNumber = CStr(oRecordset.Fields("ListNumber").Value)
			sSenderName = CStr(oRecordset.Fields("SenderName").Value)
			sRecipientName = CStr(oRecordset.Fields("RecipientName").Value)
			oRecordset.Close
			
			sErrorDescription = "No se pudo obtener la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaperworkNumber, PaperworkSenders.SenderID, SenderName, PositionName, EmployeeName, Paperworks.StartDate, DocumentSubject From Paperworks, PaperworkSenders Where (Paperworks.SenderID=PaperworkSenders.SenderID) And (Paperworks.PaperworkNumber In (" & sPaperworkIDs & ")) Order By PaperworkNumber", "EmployeeSupportLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>"
						Response.Write "No. de lista: " & CleanStringForHTML(sListNumber) & "<BR />"
						Response.Write "Procedencia: " & CleanStringForHTML(sSenderName) & "<BR />"
						Response.Write "Dirigido a: " & CleanStringForHTML(sRecipientName) & "<BR />"
					Response.Write "</B></FONT><BR /><BR />"
					Response.Write "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
						asColumnsTitles = Split("Consecutivo,Folio,Procedencia,Fecha,Observaciones", ",", -1, vbBinaryCompare)
						asCellWidths = Split("100,100,100,100,400", ",", -1, vbBinaryCompare)
						asCellAlignments = Split(",,,,", ",", -1, vbBinaryCompare)
						lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)

						iCounter = 1
						Do While Not oRecordset.EOF
							sRowContents = iCounter
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaperworkNumber").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SenderName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value) & " (Puesto: " & CStr(oRecordset.Fields("EmployeeName").Value) & ")")
							sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
							'sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, -1)
							If CLng(oRecordset.Fields("SenderID").Value) = 3853 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("DocumentSubject").Value))
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
							End If
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)

							iCounter = iCounter + 1
							oRecordset.MoveNext
							If Err.number <> 0 Then Exit Do
						Loop
					Response.Write "</TABLE>" & vbNewLine
				Else
					lErrorNumber = L_ERR_NO_RECORDS
					sErrorDescription = "No existen registros que cumplan con los criterios de la búsqueda."
				End If
			End If
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros que cumplan con los criterios de la búsqueda."
		End If
	End If

	Set oRecordset = Nothing
	PrintPaperworkList = lErrorNumber
	Err.Clear
End Function
%>