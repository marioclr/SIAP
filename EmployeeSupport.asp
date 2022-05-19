<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
%>
<!-- #include file="Libraries/GlobalVariables.asp" -->
<!-- #include file="Libraries/LoginComponent.asp" -->
<!-- #include file="Libraries/CatalogsLib.asp" -->
<!-- #include file="Libraries/CatalogComponent.asp" -->
<!-- #include file="Libraries/EmployeeSupportLib.asp" -->
<%
Dim bAction
Dim bDisplayTable
Dim sCondition
Dim oRecordset
Dim asIDs
Dim iIndex
Dim sNames
Dim sOwnerIDs
Dim sError
Dim bClose
Dim sIDs

If B_ISSSTE Then
Else
	If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And (N_CATALOGS_PERMISSIONS) Then
	Else
		Response.Redirect "AccessDenied.asp?Permission=" & N_CATALOGS_PERMISSIONS
	End If
End If

Call GetPaperworksURLValues(oRequest, bAction, bDisplayTable, sCondition)
Call InitializeSupportComponent(oRequest, aCatalogComponent)
If bAction Then
	Call InitializeValuesForCatalogComponent(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
	lErrorNumber = DoPaperworkAction(oRequest, oADODBConnection, aCatalogComponent, sCondition, sErrorDescription)
	sError = sErrorDescription
End If

If Len(oRequest("ForGuides").Item) > 0 Then
	aHeaderComponent(S_TITLE_NAME_HEADER) = "Impresión de guías"
	Response.Cookies("SoS_SectionID") = 1642
ElseIf Len(oRequest("ForReport").Item) > 0 Then
	aHeaderComponent(S_TITLE_NAME_HEADER) = "Generación de volantes"
	Response.Cookies("SoS_SectionID") = 1641
ElseIf (Len(oRequest("Search").Item) > 0) Or (Len(oRequest("DoSearch").Item) > 0) Then
	aHeaderComponent(S_TITLE_NAME_HEADER) = "Búsqueda de documentos"
	Response.Cookies("SoS_SectionID") = 1061
ElseIf (Len(oRequest("Close").Item) > 0) Or (Len(oRequest("DoClose").Item) > 0) Then
	aHeaderComponent(S_TITLE_NAME_HEADER) = "Descargo de documentos"
	Response.Cookies("SoS_SectionID") = 1061
ElseIf Len(oRequest("PaperworkID").Item) > 0 Then
	aHeaderComponent(S_TITLE_NAME_HEADER) = "Seguimiento de documentos"
	Response.Cookies("SoS_SectionID") = 1061
ElseIf Len(oRequest("Owners").Item) > 0 Then
	aHeaderComponent(S_TITLE_NAME_HEADER) = "Permisos de los usuarios para ver responsables"
	Response.Cookies("SoS_SectionID") = 1061
Else
	aHeaderComponent(S_TITLE_NAME_HEADER) = "Registro de documentos"
	Response.Cookies("SoS_SectionID") = 1061
End If
Select Case CInt(Request.Cookies("SIAP_SectionID"))
	Case 1
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
	Case 2
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
	Case 3
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case 4
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
	Case 5
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = BUDGET_TOOLBAR
	Case 6
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = REPORTS_TOOLBAR
	Case Else
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
End Select
bWaitMessage = True

asIDs = ""
sErrorDescription = "No se pudo obtener la información de la ventanilla única."
lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CurrentID, LastID From PaperworkConsecutiveIDs Where (CurrentID>-1) Order By OrderInList", "EmployeeSupport.asp", "_root", 000, sErrorDescription, oRecordset)
If lErrorNumber = 0 Then
	Do While Not oRecordset.EOF
		asIDs = asIDs & CStr(oRecordset.Fields("CurrentID").Value) & "," & CStr(oRecordset.Fields("LastID").Value) & ";"
		oRecordset.MoveNext
	Loop
	oRecordset.Close
	If Len(asIDs) > 0 Then asIDs = Left(asIDs, (Len(asIDs) - Len(";")))
	asIDs = Split(asIDs, ";")
End If
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
		<SCRIPT LANGUAGE="JavaScript"><!--
			var lEstimatedDate = 30000000;
			var asSubjectTypes = new Array(<%
				Call GenerateJavaScriptArrayFromQuery(oADODBConnection, "SubjectTypes", "SubjectTypeID", "DaysForAttention", "(DaysForAttention>0)", "SubjectTypeID", "")
			%>['-2', '0']);
			var asSectionNumbers = new Array(<%
				For iIndex = 0 To UBound(asIDs)
					asIDs(iIndex) = Split(asIDs(iIndex), ",")
					Response.Write "['" & asIDs(iIndex)(0) & "', '"
						If lErrorNumber = 0 Then
							sErrorDescription = "No se pudo obtener la información de la ventanilla única."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(PaperworkNumber) From Paperworks Where (PaperworkNumber>=" & asIDs(iIndex)(0) & ") And (PaperworkNumber<=" & asIDs(iIndex)(1) & ") And (StartDate>" & Year(Date()) & "0000)", "EmployeeSupport.asp", "_root", 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									If Not IsNull(oRecordset.Fields(0).Value) Then
										Response.Write CLng(oRecordset.Fields(0).Value) + 1
									Else
										Response.Write asIDs(iIndex)(0)
									End If
								Else
									Response.Write asIDs(iIndex)(0)
								End If
							Else
								Response.Write asIDs(iIndex)(0)
							End If
						End If
					Response.Write "'], "
				Next
			%>['', '']);

			<%sOwnerIDs = ",-2,"
			sErrorDescription = "No se pudo obtener la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select OwnerID From UsersOwnersLKP Where (UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ")", "EmployeeSupport.asp", "_root", 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				Do While Not oRecordset.EOF
					sOwnerIDs = sOwnerIDs & CStr(oRecordset.Fields("OwnerID").Value) & ","
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				oRecordset.Close
			End If
			If InStr(1, sOwnerIDs & ",", ",-1,", vbBinaryCompare) = 0 Then
				sErrorDescription = "No se pudieron obtener los permisos del usuario."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select OwnerID From PaperworkOwners Where (ParentID In (-2" & sOwnerIDs & "-2))", "EmployeeSupportLib.asp", "_root", 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						sOwnerIDs = sOwnerIDs & CStr(oRecordset.Fields("OwnerID").Value) & ","
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
				sErrorDescription = "No se pudieron obtener los permisos del usuario."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select OwnerID From PaperworkOwners Where (ParentID In (-2" & sOwnerIDs & "-2))", "EmployeeSupportLib.asp", "_root", 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						sOwnerIDs = sOwnerIDs & CStr(oRecordset.Fields("OwnerID").Value) & ","
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
			End If%>
			var sOwnerIDs = '<%Response.Write sOwnerIDs%>';

			function ChangePpwkNumber(sSectionID) {
				var oForm = document.CatalogFrm;

				if (oForm) {
					oForm.PaperworkNumber.value = '';
					for (var i=0; i<asSectionNumbers.length; i++) {
						if (asSectionNumbers[i][0] == sSectionID) {
							oForm.PaperworkNumber.value = asSectionNumbers[i][1];
							break;
						}
					}
				}
			} // End of ChangePpwkNumber

			function CheckControlForm() {
				var oForm = document.CatalogFrm;
				var oControlForm = document.ControlFrm;

				if (oForm) {
					SelectAllItemsFromList(oForm.OwnerIDs);
					SelectAllItemsFromList(oForm.ActionIDs);
					SelectAllItemsFromList(oForm.ReportDates);
					SelectAllItemsFromList(oForm.EndDates);
					SelectAllItemsFromList(oForm.ClosingNumbers);
					SelectAllItemsFromList(oForm.OwnersComments);

					<% If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0)) = -1 Then %>
						if (oForm.Description.value == '') {
							oForm.Description.value = '.';
						}
						if ((oForm.SenderID.value == '') || (oForm.SenderID.value == '-1')) {
							alert('Favor de especificar la procedencia.');
							oForm.SenderName.focus();
							return false;
						}
						if ((oForm.SubjectTypeID.value == '') || (oForm.SubjectTypeID.value == '-1')) {
							alert('Favor de especificar el tipo de asunto.');
							oForm.SubjectTypeName.focus();
							return false;
						}
						if (parseInt(oForm.EstimatedDateYear.value + oForm.EstimatedDateMonth.value + oForm.EstimatedDateDay.value) > 0) {
							if (parseInt(oForm.StartDateYear.value + oForm.StartDateMonth.value + oForm.StartDateDay.value) > parseInt(oForm.EstimatedDateYear.value + oForm.EstimatedDateMonth.value + oForm.EstimatedDateDay.value)) {
								alert('La fecha límite no puede ser anterior a la fecha del documento.');
								oForm.EstimatedDateDay.focus();
								return false;
							}
						}
					<% Else %>
						if (oForm.DESCRIPTION.value == '') {
							oForm.DESCRIPTION.value = '.';
						}
						if ((oForm.SENDERID.value == '') || (oForm.SENDERID.value == '-1')) {
							alert('Favor de especificar la procedencia.');
						    //oForm.SenderName.focus();
							return false;
						}
						if ((oForm.SUBJECTTYPEID.value == '') || (oForm.SUBJECTTYPEID.value == '-1')) {
							alert('Favor de especificar el tipo de asunto.');
							//oForm.SubjectTypeName.focus();
							return false;
						}
						if (parseInt(oForm.ESTIMATEDDATEYear.value + oForm.ESTIMATEDDATEMonth.value + oForm.ESTIMATEDDATEDay.value) > 0) {
							if (parseInt(oForm.STARTDATEYear.value + oForm.STARTDATEMonth.value + oForm.STARTDATEDay.value) > parseInt(oForm.ESTIMATEDDATEYear.value + oForm.ESTIMATEDDATEMonth.value + oForm.ESTIMATEDDATEDay.value)) {
								alert('La fecha límite no puede ser anterior a la fecha del documento.');
								oForm.ESTIMATEDDATEDay.focus();
								return false;
							}
						}
					<% End If %>
/*
					if ((oForm.OwnerID.value != '') && (oControlForm.EmployeeID.value == '-1')) {
						alert('El número del empleado es inválido');
						oForm.OwnerID.focus();
						return false;
					} else {
						if (oForm.OwnerID.value == '') {
							alert('Favor de especificar el número de empleado');
							oForm.OwnerID.focus();
							return false;
						} else {
							oForm.OwnerID.value = oControlForm.EmployeeID.value;
						}
					}
*/
					if (oForm.OwnerIDs.length == 0) {
						alert('Favor de turnar el documento.');
						oForm.OwnerIDToSearch.focus();
						return false;
					}
					return true;
				}
				return false;
			} // End of CheckControlForm
			
			function AddOwnerComment() {
				var oForm = document.CatalogFrm;
				var bEmpty = true;
				if (oForm) {
					for (var i=0; i<oForm.OwnerIDs.options.length; i++) {
						if ((oForm.OwnerIDs.options[i].value == oForm.OwnerIDTemp.value) && (oForm.ActionIDs.options[i].value == oForm.ActionIDTemp.value)) {
							bEmpty = false;
							break;
						}
					}
					if (bEmpty) {
						AddItemToList(GetSelectedText(oForm.OwnerIDTemp), oForm.OwnerIDTemp.value, null, oForm.OwnerIDs);
						AddItemToList(GetSelectedText(oForm.ActionIDTemp), oForm.ActionIDTemp.value, null, oForm.ActionIDs);
						AddItemToList('<%Response.Write DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, -1)%>', '<%Response.Write Left(GetSerialNumberForDate(""), Len("00000000"))%>', null, oForm.ReportDates);
						AddItemToList('---', 0, null, oForm.EndDates);
						AddItemToList('', '', null, oForm.ClosingNumbers);
						AddItemToList('', '', null, oForm.OwnersComments);
						ResizeOwnerComments();
						oForm.OwnerIDTemp.options[0].selected = true;
						oForm.OwnerIDTemp.focus();
					}
				}
			} // End of AddOwnerComment

			function AddPaperworkToClose() {
				var bCorrect = true;
				var oForm = document.CloseFrm;
				if (oForm) {
					if (oForm.PaperworkNumberTemp.value == '') {
						alert('Favor de especificar el número de folio del documento a cerrar');
						oForm.PaperworkNumberTemp.focus();
						bCorrect = false;
					} else {
						if (! CheckIntegerValue(oForm.PaperworkNumberTemp, 'el folio del documento a cerrar', N_MINIMUM_ONLY_FLAG, N_OPEN_FLAG, 0, 0))
							bCorrect = false;
					}
					if (oForm.DocClassificationTemp.value == '') {
						alert('Favor de especificar el número del oficio de descargo');
						oForm.DocClassificationTemp.focus();
						bCorrect = false;
					}
					if (bCorrect) {
						AddItemToList(oForm.PaperworkNumberTemp.value, oForm.PaperworkNumberTemp.value, null, oForm.PaperworkNumbers);
						AddItemToList(oForm.PaperworkYearTemp.value, oForm.PaperworkYearTemp.value, null, oForm.PaperworkYears);
						AddItemToList(GetSelectedText(oForm.OwnerTemp), oForm.OwnerTemp.value, null, oForm.Owners);
						AddItemToList(oForm.DocClassificationTemp.value, oForm.DocClassificationTemp.value, null, oForm.DocClassifications);
						AddItemToList(oForm.CommentsTemp.value, oForm.CommentsTemp.value, null, oForm.Comments);
						ResizePaperworksToClose();
						oForm.PaperworkNumberTemp.value = '';
						//oForm.OwnerTemp.value = '';
						oForm.DocClassificationTemp.value = '';
						oForm.CommentsTemp.value = '';
						oForm.PaperworkNumberTemp.focus();
					}
				}
			} // End of AddPaperworkToClose

			function RemoveOwnerComment() {
				var oForm = document.CatalogFrm;
				var i = 0;
				var oRegExp = null;
				if (oForm) {
					for (i=0; i<oForm.OwnerIDs.options.length; i++) {
						if (oForm.OwnerIDs.options[i].selected) {
							oRegExp = eval('/,' + oForm.OwnerIDs.options[i].value + ',/gi');
							<%If InStr(1, "," & sOwnerIDs & ",", ",-1,", vbBinaryCompare) = 0 Then%>
							if (sOwnerIDs.search(oRegExp) != -1) {
							<%Else%>
							if (true) {
							<%End If%>
								oForm.OwnerIDs.options[i] = null;
								oForm.ActionIDs.options[i] = null;
								oForm.ReportDates.options[i] = null;
								oForm.EndDates.options[i] = null;
								oForm.ClosingNumbers.options[i] = null;
								oForm.OwnersComments.options[i] = null;
								i--;
							}
						}
					}
					ResizeOwnerComments();
					oForm.OwnerIDTemp.focus();
				}
			} // End of RemoveOwnerComment

			function RemovePaperworkToClose() {
				var oForm = document.CloseFrm;
				if (oForm) {
					RemoveSelectedItemsFromList(null, oForm.PaperworkNumbers);
					RemoveSelectedItemsFromList(null, oForm.PaperworkYears);
					RemoveSelectedItemsFromList(null, oForm.Owners);
					RemoveSelectedItemsFromList(null, oForm.DocClassifications);
					RemoveSelectedItemsFromList(null, oForm.Comments);
					ResizePaperworksToClose();
					oForm.PaperworkNumberTemp.focus();
				}
			} // End of RemovePaperworkToClose

			function ResizeOwnerComments() {
				var oForm = document.CatalogFrm;
				if (oForm) {
					if (oForm.OwnerIDs.options.length > 3) {
						oForm.OwnerIDs.size = oForm.OwnerIDs.options.length;
						oForm.ActionIDs.size = oForm.ActionIDs.options.length;
						oForm.ReportDates.size = oForm.ReportDates.options.length;
						oForm.EndDates.size = oForm.EndDates.options.length;
						oForm.ClosingNumbers.size = oForm.ClosingNumbers.options.length;
						oForm.OwnersComments.size = oForm.OwnersComments.options.length;
					} else {
						oForm.OwnerIDs.size = 3;
						oForm.ActionIDs.size = 3;
						oForm.ReportDates.size = 3;
						oForm.EndDates.size = 3;
						oForm.ClosingNumbers.size = 3;
						oForm.OwnersComments.size = 3;
					}
				}
			} // End of ResizeOwnerComments

			function ResizePaperworksToClose() {
				var oForm = document.CloseFrm;
				if (oForm) {
					if (oForm.PaperworkNumbers.options.length > 3) {
						oForm.PaperworkNumbers.size = oForm.PaperworkNumbers.options.length;
						oForm.PaperworkYears.size = oForm.PaperworkYears.options.length;
						oForm.DocClassifications.size = oForm.DocClassifications.options.length;
						oForm.Comments.size = oForm.Comments.options.length;
						oForm.Owners.size = oForm.Owners.options.length;
					} else {
						oForm.PaperworkNumbers.size = 3;
						oForm.PaperworkYears.size = 3;
						oForm.DocClassifications.size = 3;
						oForm.Comments.size = 3;
						oForm.Owners.size = 3;
					}
				}
			} // End of ResizePaperworksToClose

			function SelectSameItemsForOwners(oList) {
				var oForm = document.CatalogFrm;
				if (oForm) {
					SelectSameItems(oList, oForm.OwnerIDs);
					SelectSameItems(oList, oForm.ActionIDs);
					SelectSameItems(oList, oForm.ReportDates);
					SelectSameItems(oList, oForm.EndDates);
					SelectSameItems(oList, oForm.ClosingNumbers);
					SelectSameItems(oList, oForm.OwnersComments);
				}
			} // End of SelectSameItemsForOwners

			function SelectSameItemsForPaperworksToClose(oList) {
				var oForm = document.CloseFrm;
				if (oForm) {
					SelectSameItems(oList, oForm.PaperworkNumbers);
					SelectSameItems(oList, oForm.PaperworkYears);
					SelectSameItems(oList, oForm.Owners);
					SelectSameItems(oList, oForm.DocClassifications);
					SelectSameItems(oList, oForm.Comments);
				}
			} // End of SelectSameItemsForPaperworksToClose

			function CheckEstimatedDate() {
				var oForm = document.CatalogFrm;

				if (oForm) {
					if (lEstimatedDate < parseInt(oForm.EstimatedDateYear.value + oForm.EstimatedDateMonth.value + oForm.EstimatedDateDay.value)) {
						alert('La fecha límite no puede ser posterior a ' + ('' + lEstimatedDate).substr(6, 2) + '/' + ('' + lEstimatedDate).substr(4, 2) + '/' + ('' + lEstimatedDate).substr(0, 4));
						SetDateCombos(('' + lEstimatedDate).substr(0, 4), ('' + lEstimatedDate).substr(4, 2), ('' + lEstimatedDate).substr(6, 2), oForm.EstimatedDateYear, oForm.EstimatedDateMonth, oForm.EstimatedDateDay);
					}
				}
			} // End of CheckEstimatedDate
		//--></SCRIPT>
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%If ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS) = N_ADD_PERMISSIONS) And (Len(oRequest("ForReport").Item) = 0) And (Len(oRequest("Owners").Item) = 0) Then
			aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Agregar documento",_
					  "",_
					  "", GetASPFileName("") & "?New=1", (CInt(Request.Cookies("SIAP_SectionID")) = 2)),_
				Array("Exportar a Excel",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=Paperworks&Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(RemoveParameterFromURLString(oRequest, "Action")) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (((Len(oRequest("Search").Item) > 0) Or (Len(oRequest("DoSearch").Item) > 0) Or (Len(oRequest("ForReport").Item) > 0)) And (Len(oRequest("ForGuides").Item) = 0) And bDisplayTable)),_
				Array("Imprimir",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=Reports&Word=1&PaperworkID=" & oRequest("PaperworkID").Item & "&ReportID=1600&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (Len(oRequest("PaperworkID").Item) > 0))_
			)
			aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 810
			aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
			aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 183
		End If%>
		<!-- #include file="_Header.asp" -->
		<%Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> > "
			Select Case CInt(Request.Cookies("SIAP_SectionID"))
				Case 1
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > "
				Case 2
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > "
				Case 3
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > "
				Case 4
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > "
				Case 5
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=5"">Presupuesto</A> > "
				Case 6
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=6"">Departamento técnico</A> > "
				Case Else
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > "
			End Select
			Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > "
			If Len(oRequest("ForGuides").Item) > 0 Then
				Response.Write "<B>Impresión de guías</B>"
			ElseIf Len(oRequest("ForReport").Item) > 0 Then
				Response.Write "<B>Generación de volantes</B>"
			ElseIf (Len(oRequest("Search").Item) > 0) Or (Len(oRequest("DoSearch").Item) > 0) Then
				Response.Write "<B>Búsqueda de documentos</B>"
			ElseIf (Len(oRequest("Assign").Item) > 0) Or (Len(oRequest("DoAssign").Item) > 0) Then
				Response.Write "<B>Asignación de documentos</B>"
			ElseIf (Len(oRequest("Close").Item) > 0) Or (Len(oRequest("DoClose").Item) > 0) Then
				Response.Write "<B>Descargo de documentos</B>"
			ElseIf Len(oRequest("PaperworkID").Item) > 0 Then
				Response.Write "<B>Seguimiento de documentos</B>"
			ElseIf Len(oRequest("Owners").Item) > 0 Then
				If Len(oRequest("UserID").Item) > 0 Then
					Call GetNameFromTable(oADODBConnection, "Users", CStr(oRequest("UserID").Item), "", "", sNames, "")
					Response.Write "<B>Permisos de los usuarios para ver responsables: " & sNames & "</B>"
				Else
					Response.Write "<B>Permisos de los usuarios para ver responsables</B>"
				End If
			Else
				Response.Write "<B>Registro de documentos</B>"
			End If
		Response.Write "<BR /><BR />"
		'sIDs = GetOwnerHierarchy(oRequest, oADODBConnection, 1300, sErrorDescription)
		'If (InStr(Left(sIDs, 1), ",") > 0) Then
		'    sIDs = Right(sIDs, Len(sIDs) -1)
		'End If
		'Response.Write sIDs
		If Len(sError) > 0 Then
			Call DisplayErrorMessage("Mensaje del sistema", sError)
		ElseIf lErrorNumber <> 0 Then
			Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
			lErrorNumber = 0
			sErrorDescription = ""
		ElseIf Len(oRequest("Error").Item) > 0 Then
			Call DisplayErrorMessage("Mensaje del sistema", CStr(oRequest("Error").Item))
		End If
		Response.Write "<BR />"
		If (Len(oRequest("Search").Item) > 0) Or (Len(oRequest("DoSearch").Item) > 0) Or (Len(oRequest("ForReport").Item) > 0) Then
			If Len(oRequest("ForGuides").Item) > 0 Then
				lErrorNumber = DisplayGuideSearchFrom(oRequest, oADODBConnection, sErrorDescription)
			Else
				lErrorNumber = DisplayPaperworksSearchFrom(oRequest, oADODBConnection, sErrorDescription)
				If (lErrorNumber = 0) And bDisplayTable Then
					Response.Write "<IMG SRC=""Images/DotBlue.gif"" WIDTH=""700"" HEIGHT=""1"" /><BR /><BR />"
					lErrorNumber = DisplayPaperworksForSupportTable(oRequest, oADODBConnection, True, False, sCondition, sErrorDescription)
				End If
				If lErrorNumber <> 0 Then
					Response.Write "<BR />"
					Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
					lErrorNumber = 0
					sErrorDescription = ""
				End If
			End If
		ElseIf (Len(oRequest("Assign").Item) > 0) Or (Len(oRequest("DoAssign").Item) > 0) Then
			lErrorNumber = DisplayAssignPaperworksFrom(oRequest, oADODBConnection, sErrorDescription)
			If lErrorNumber <> 0 Then
				Response.Write "<BR />"
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
				lErrorNumber = 0
				sErrorDescription = ""
			End If
		ElseIf (Len(oRequest("Close").Item) > 0) Or (Len(oRequest("DoClose").Item) > 0) Then
			lErrorNumber = DisplayClosePaperworksFrom(oRequest, oADODBConnection, sErrorDescription)
			If lErrorNumber <> 0 Then
				Response.Write "<BR />"
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
				lErrorNumber = 0
				sErrorDescription = ""
			End If
		ElseIf Len(oRequest("Owners").Item) > 0 Then
			Response.Write "<FORM NAME=""OwnersFrm"" ID=""OwnersFrm"" ACTION=""EmployeeSupport.asp"" METHOD=""POST"">"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Owners"" ID=""OwnersHdn"" VALUE=""1"" />"
				If Len(oRequest("UserID").Item) = 0 Then
					If Len(sError) = 0 Then Call DisplayInstructionsMessage("Permisos de los usuarios para ver responsables", "<B>Paso 1. Seleccione al usuario</B><BR />Paso 2. Seleccione los responsables que podrá ver")
					Response.Write "<BR /><BR />"
					Response.Write "Usuario:&nbsp;"
					Response.Write "<SELECT NAME=""UserID"" ID=""UserIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						sErrorDescription = "No se pudo obtener la información de los registros."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select UserID, UserName, UserLastName, UserPermissions4, ProfileName From Users, UserProfiles Where (Users.ProfileID=UserProfiles.ProfileID) And (UserID>9) And ((InStr(',' || UserPermissions2 || ',', ',175,') > 0 Or InStr(',' || UserPermissions2 || ',', ',256,') > 0 Or InStr(',' || UserPermissions2 || ',', ',312,') > 0  Or InStr(',' || UserPermissions2 || ',', ',412,') > 0 Or InStr(',' || UserPermissions2 || ',', ',509,') > 0 Or InStr(',' || UserPermissions2 || ',', ',603,') > 0 Or InStr(',' || UserPermissions2 || ',', ',800,') > 0) Or (UserPermissions4=-1)) Order By UserLastName, UserName", "EmployeeSupport.asp", "_root", 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							Do While Not oRecordset.EOF
								If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or _
									InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_VentanillaUnica & ",", vbBinaryCompare) > 0 Or _
									InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_VentanillaUnica & ",", vbBinaryCompare) > 0 Or _
									InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_VentanillaUnica & ",", vbBinaryCompare) > 0 Or _
									InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_VentanillaUnica & ",", vbBinaryCompare) > 0 Or _
									InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_05_VentanillaUnica & ",", vbBinaryCompare) > 0 Or _
									InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_VentanillaUnica & ",", vbBinaryCompare) > 0 Or _
									InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_VentanillaUnica & ",", vbBinaryCompare) > 0) Then
									Response.Write "<OPTION VALUE=""" & CStr(oRecordset.Fields("UserID").Value) & """>" & CStr(oRecordset.Fields("UserLastName").Value) & ", " & CStr(oRecordset.Fields("UserName").Value) & " (" & CStr(oRecordset.Fields("ProfileName").Value) & ")</OPTION>"
								End If
								oRecordset.MoveNext
								If Err.number <> 0 Then Exit Do
							Loop
						End If
					Response.Write "</SELECT><BR /><BR />"
					Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Continue"" ID=""ContinueBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
				Else
					If Len(sError) = 0 Then Call DisplayInstructionsMessage("Permisos de los usuarios para ver responsables", "Paso 1. Seleccione al usuario<BR /><B>Paso 2. Seleccione los responsables que podrá ver</B>")
					Response.Write "<BR /><BR />"
					Response.Write "Responsables:<BR />&nbsp;&nbsp;&nbsp;"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserID"" ID=""UserIDHdn"" VALUE=""" & oRequest("UserID").Item & """ />"
					Response.Write "<SELECT NAME=""OwnerID"" ID=""OwnerIDCmb"" SIZE=""10"" MULTIPLE=""1"" CLASS=""Lists"">"
						sErrorDescription = "No se pudo obtener la información de los registros."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select OwnerID From UsersOwnersLKP Where (UserID=" & oRequest("UserID").Item & ")", "EmployeeSupport.asp", "_root", 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							sNames = "-2"
							Do While Not oRecordset.EOF
								sNames = sNames & "," & CStr(oRecordset.Fields("OwnerID").Value)
								oRecordset.MoveNext
								If Err.number <> 0 Then Exit Do
							Loop
							oRecordset.Close
						End If
						Response.Write "<OPTION VALUE=""-1"""
							If InStr(1, sNames, ",-1", vbBinaryCompare) > 1 Then Response.Write " SELECTED=""1"""
						Response.Write ">Todos</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "PaperworkOwners", "OwnerID", "OwnerID As RecordID, OwnerName, 'Empleado:' As Temp1, EmployeeID", sCondition, "OwnerID, OwnerName", sNames, "", sErrorDescription)
					Response.Write "</SELECT><BR /><BR />"
					Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Associate"" ID=""AssociateBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
					Response.Write "<INPUT TYPE=""BUTTON"" VALUE=""Seleccionar Otro Usuario"" CLASS=""Buttons"" onClick=""window.location.href = 'EmployeeSupport.asp?Owners=1';"" />"
				End If
			Response.Write "</FORM>"
		Else
			aCatalogComponent(S_URL_PARAMETERS_CATALOG) = ""
			If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0)) > -1 Then
				aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,4,1,5,6,11,11,5,6,6,1,6,5,1,11,6"
				aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
				aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
				aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
				aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) = "1,2,3,4,5,6,7,8,9,10,11,12,13,14,15"

				If (StrComp("8", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or _
					InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or _
					InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or _
					InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0) Then
					aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) = "1,4,5,6,10"
				End If
				If (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_ModificarResponsableDeDocumento & ",", vbBinaryCompare) > 0) Then
					aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) = "1,5,6,10"
				End If
				lErrorNumber = GetCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
				If Not (InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_ModificarResponsableDeDocumento & ",", vbBinaryCompare) > 0) Then
					aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(4) = "PaperworkSenders;,;SenderID;,;SenderID As RecordID, SenderName, EmployeeName, PositionName;,;(SenderID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(4) & ");,;SenderID;,;;,;Ninguna;;;-1"
				End If
			End If
			lErrorNumber = DisplayOwnersInCatalogForm(oRequest, oADODBConnection, CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0)), (Len(Trim(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(14))) > 0),  sErrorDescription)
			If lErrorNumber <> 0 Then
				Response.Write "<BR />"
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
				lErrorNumber = 0
				sErrorDescription = ""
			End If
			If Len(Trim(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(14))) > 0 Then
				Call DisplayInstructionsMessage("Documento cerrado", "<B>Fecha de atención:</B> " & DisplayDateFromSerialNumber(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(13), -1, -1, -1) & "<BR /><B>Oficio de descargo:</B> " & CleanStringForHTML(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(14)) & "<BR />")
			End If
			lErrorNumber = DisplayCatalogForm(oRequest, oADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				If Len(oRequest("New").Item) = 0 Then
					'Response.Write "HideDisplay(document.all['SectionIDDiv']);" & vbNewLine
				Else
					Response.Write "ChangePpwkNumber(document.CatalogFrm.SectionID.value);" & vbNewLine
				End If
'				If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0)) > -1 Then
'					Call GetNameFromTable(oADODBConnection, "PaperworkSenders", aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(4), "", "", sNames, sErrorDescription)
'					Response.Write "document.CatalogFrm.SenderName.value = '" & sNames & "';" & vbNewLine
'					Call GetNameFromTable(oADODBConnection, "SubjectTypes", aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(9), "", "", sNames, sErrorDescription)
'					Response.Write "document.CatalogFrm.SubjectTypeName.value = '" & sNames & "';" & vbNewLine
'				End If
				If Len(Trim(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(14))) > 0 Then
					Response.Write "HideDisplay(document.CatalogFrm.Modify);" & vbNewLine
				End If
			Response.Write "//--></SCRIPT>" & vbNewLine
			If lErrorNumber <> 0 Then
				Response.Write "<BR />"
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
				lErrorNumber = 0
				sErrorDescription = ""
			End If

			If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0)) > -1 Then
                bClose = False
                If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(13)) > 0 Then
                    bClose = True
                End If
				Response.Write "<IMG SRC=""Images/DotBlue.gif"" WIDTH=""700"" HEIGHT=""1"" /><BR /><BR />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>SOPORTE DOCUMENTAL</B><BR /><BR /></FONT>"
                Response.Write "<IFRAME SRC=""BrowserFile.asp?PaperworkID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) & "&PaperworkNumber=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & "&StartDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2) & "&isClose=" & bClose & """ NAME=""EmployeeFilesIFrame"" FRAMEBORDER=""0"" WIDTH=""420"" HEIGHT=""500""></IFRAME>"
                Response.Write "&nbsp;&nbsp;&nbsp;"
				Response.Write "<IFRAME SRC=""HistoryList.asp?Action=Paperworks&PaperworkID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) & "&OwnerID=" & aLoginComponent(S_EMPLOYEE_NUMBER_LOGIN) & """ NAME=""HistoryListIFrame"" FRAMEBORDER=""0"" WIDTH=""720"" HEIGHT=""500""></IFRAME>"
			End If

			Response.Write "<FORM NAME=""ControlFrm"" ID=""ControlFrm"" onSubmit=""return false"">"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""-1"" />"
			Response.Write "</FORM>"
			Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchRecordIFrame"" FRAMEBORDER=""0"" WIDTH=""0"" HEIGHT=""0""></IFRAME>"

			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG))) > -1 Then
					Response.Write "document.ControlFrm.EmployeeID.value='" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(5) & "';" & vbNewLine
					Response.Write "ResizeOwnerComments();" & vbNewLine
				Else
					Response.Write "if (document.all['CatalogFrm_EndDateDiv']) {" & vbNewLine
						Response.Write "HideDisplay(document.all['CatalogFrm_EndDateDiv']);" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (document.all['CatalogFrm_DocClassificationDiv']) {" & vbNewLine
						Response.Write "HideDisplay(document.all['CatalogFrm_DocClassificationDiv']);" & vbNewLine
					Response.Write "}" & vbNewLine
				End If
				Response.Write "if (document.all['CatalogFrm_EndDateDiv']) {" & vbNewLine
					Response.Write "HideDisplay(document.all['CatalogFrm_EndDateDiv']);" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (document.all['CatalogFrm_StatusIDDiv']) {" & vbNewLine
					Response.Write "HideDisplay(document.all['CatalogFrm_StatusIDDiv']);" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
		End If
		%>
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>