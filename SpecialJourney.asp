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
<!-- #include file="Libraries/EmployeeComponent.asp" -->
<!-- #include file="Libraries/EmployeeSpecialJourneysLib.asp" -->
<!-- #include file="Libraries/EmployeeSpecialJourneyComponent.asp" -->
<!-- #include file="Libraries/JobComponent.asp" -->
<!-- #include file="Libraries/PayrollComponent.asp" -->
<!-- #include file="Libraries/QueriesLib.asp" -->
<!-- #include file="Libraries/UploadInfoLibrary.asp" -->
<%
Dim iSelectedTab
Dim sNames
Dim bShowForm
Dim bAction
Dim bError
Dim sCondition
Dim sAction
Dim sFilter
Dim bFilter

Dim iSectionID
Dim sSubSectionID
Dim iStep
Dim sFileName
Dim lReasonID
Dim iSpecialJourneyType

Dim sAltDescription
Dim sDescription
Dim iStarPage

sFilter = ""
sAltDescription = "Guardias"
sDescription = "Registre una guardia a un empleado diferente."

iStep = 1
If Len(oRequest("Step").Item) > 0 Then iStep = CInt(oRequest("Step").Item)
If CInt(oRequest("SectionID").Item) > 0 Then Response.Cookies("SIAP_SectionID") = CInt(oRequest("SectionID").Item)
If CInt(oRequest("SubSectionID").Item) > 0 Then Response.Cookies("SIAP_SubSectionID") = CInt(oRequest("SubSectionID").Item)
If CInt(oRequest("SpecialJourneyType").Item) > 0 Then iSpecialJourneyType = CInt(oRequest("SpecialJourneyType").Item)
iStarPage = CInt(oRequest("StartPage").Item)
sAction = oRequest("Action").Item

sFileName = Server.MapPath(UPLOADED_PHYSICAL_PATH & oRequest("Load").Item & "_" & aLoginComponent(N_USER_ID_LOGIN) & ".txt")
If Len(oRequest("RawData").Item) > 0 Then
	lErrorNumber = SaveTextToFile(sFileName, oRequest("RawData").Item, sErrorDescription)
	Response.Redirect "SpecialJourney.asp?Action=SpecialJourneyType=" & iSpecialJourneyType & "&Step=" & iStep
End If

Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
Call InitializeSpecialJourneyComponent(oRequest, aSpecialJourneyComponent)
Call GetSpecialJourneysURLValues(oRequest, iSelectedTab, bAction, aSpecialJourneyComponent(S_QUERY_CONDITION_SPECIAL_JOURNEY))

aJobComponent(B_SEND_TO_IFRAME_JOB) = True
bShowForm = True

bError = False
If bAction Then
	lErrorNumber = DoSpecialJourneysAction(oRequest, oADODBConnection, oRequest("Action").Item, iSpecialJourneyType, sErrorDescription)
	bError = (lErrorNumber <> 0)
	If bError Then
		Response.Redirect "SpecialJourney.asp?SpecialJourneyType=" & oRequest("SpecialJourneyType").Item & "&Success=1&ErrorDescription=" & sErrorDescription
	Else
		Response.Redirect "SpecialJourney.asp?SpecialJourneyType=" & oRequest("SpecialJourneyType").Item & "&Success=0"
	End If
End If

Select Case iGlobalSectionID
	Case 1
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
	Case 2
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
	Case 3
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case 4
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
	Case Else
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
End Select
aHeaderComponent(S_TITLE_NAME_HEADER) = "Guardias y Suplencias"
bWaitMessage = True
Response.Cookies("SoS_SectionID") = 191
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/Export.js"></SCRIPT>
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" >
		<%If Len(oRequest("ReadOnly").Item) = 0 Then
			aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Agregar un nuevo registro",_
					  "",_
					  "", "SpecialJourney.asp?RecordID=-1&New=1", N_ADD_PERMISSIONS),_
				Array("<LINE />",_
					  "",_
					  "", "", ((Len(oRequest("New").Item) = 0) And ((True) Or (Len(oRequest("RecordID").Item) > 0)))),_
				Array("Exportar a Excel",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=SpecialJourney&Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (True)),_
				Array("Exportar registro a Excel",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=SpecialJourney&Excel=1&RecordID=" & oRequest("RecordID").Item & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (True)),_
				Array("Exportar registro a Word",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=SpecialJourney&Word=1&RecordID=" & oRequest("RecordID").Item & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (True)),_
				Array("Exportar a Excel los registros mostrados",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=SpecialJourney&Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") &"&StartPage="& oRequest("StartPage") &"&StartForValueDay="& oRequest("StartForValueDay") &"&StartForValueMonth="& oRequest("StartForValueMonth") &"&StartForValueYear="& oRequest("StartForValueYear") &"&EndForValueDay="& oRequest("EndForValueDay") &"&EndForValueMonth="& oRequest("EndForValueMonth") &"&EndForValueYear="& oRequest("EndForValueYear") &"&PositionShortName=" & oRequest("PositionShortName") & "&GroupGradeLevelID=" & oRequest("GroupGradeLevelID") & "&ApplyFilter=++Filtrar++ " & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')",(Len(oRequest("ApplyFilter").Item)>0))_
				)
		Else
			aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Exportar a Excel",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=SpecialJourney&Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (True)),_
				Array("Exportar a Word",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=SpecialJourney&Word=1&PositionID=" & oRequest("PositionID").Item & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (True))_
			)
		End If
		aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 793
		aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
		aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 200
		%>
		<!-- #include file="_Header.asp" -->
		<%Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> > "
		If B_ISSSTE Then
            If CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
                If CInt(Request.Cookies("SIAP_SubSectionID")) = 423 Then ' Guardias
                    If iSpecialJourneyType = 1 Then
					    Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=423"">Guardias</A> > <B>Registro de información para internos</B><BR /><BR />"
                    Else
                        Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=423"">Guardias</A> > <B>Registro de información para externos</B><BR /><BR />"
                    End If
                ElseIf CInt(Request.Cookies("SIAP_SubSectionID")) = 424 Then ' Suplencias
                    If iSpecialJourneyType = 1 Then
					    Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=424"">Suplencias</A> > <B>Registro de información para internos</B><BR /><BR />"
                    Else
                        Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=424"">Suplencias</A> > <B>Registro de información para externos</B><BR /><BR />"
                    End If
				ElseIf CInt(Request.Cookies("SIAP_SubSectionID")) = 427 Then ' Personal externo
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > <B>Registro de Personal externo</B><BR /><BR />"
				ElseIf CInt(Request.Cookies("SIAP_SubSectionID")) = 428 Then ' Beneficiario(a)s de pensión
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > <B>Registro de beneficiario(a)s de pensión</B><BR /><BR />"
                End If
			ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
                If CInt(Request.Cookies("SIAP_SubSectionID")) = 423 Then ' Guardias
                    If iSpecialJourneyType = 1 Then
					    Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=731"">Empleados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=423"">Guardias</A> > <B>Registro de información para internos</B><BR /><BR />"
                    Else
                        Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=731"">Empleados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=423"">Guardias</A> > <B>Registro de información para externos</B><BR /><BR />"
                    End If
                ElseIf CInt(Request.Cookies("SIAP_SubSectionID")) = 424 Then ' Suplencias
                    If iSpecialJourneyType = 1 Then
					    Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=731"">Empleados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=424"">Suplencias</A> > <B>Registro de información para internos</B><BR /><BR />"
                    Else
                        Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=731"">Empleados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=424"">Suplencias</A> > <B>Registro de información para externos</B><BR /><BR />"
                    End If
				ElseIf CInt(Request.Cookies("SIAP_SubSectionID")) = 427 Then ' Personal externo
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=731"">Empleados</A> > <B>Registro de Personal externo</B><BR /><BR />"
				ElseIf CInt(Request.Cookies("SIAP_SubSectionID")) = 428 Then ' Beneficiario(a)s de pensión
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=731"">Empleados</A> > <B>Registro de beneficiario(a)s de pensión</B><BR /><BR />"
                End If

                    'If CInt(Request.Cookies("SIAP_SubSectionID")) = 4231 Then
					'    Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=731"">Empleados</A> > <B>Registro de información para internos</B><BR /><BR />"
                    'Else
                    '    Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=731"">Empleados</A> > <B>Registro de información para externos</B><BR /><BR />"
                    'End If
			Else
            End If
		Else
			Response.Write "<A HREF=""HumanResources.asp"">Recursos Humanos</A> > "
		End If
        Response.Write "<BR /><BR />"
            If iStep <= 1 Then
				If Len(oRequest("Success").Item) > 0 Then
					If CInt(oRequest("Success").Item) = 1 Then
						Call DisplayErrorMessage("Error al realizar la operación", CStr(oRequest("ErrorDescription").Item))
					Else
						Call DisplayErrorMessage("Confirmación", "La operación fué ejecutada exitosamente.")
					End If
				End If
				Select Case CInt(Request.Cookies("SIAP_SubSectionID"))
					Case 423, 424, 425, 426
						Response.Write "<IMG SRC=""Images/Crcl1.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: ShowDisplay(document.all['SpecialJourneyFormDiv']); if(document.all['SpecialJourneyUploadDiv'] != null) { HideDisplay(document.all['SpecialJourneyUploadDiv']) }; if(document.all['SpecialJourneyValidateDiv'] != null) { HideDisplay(document.all['SpecialJourneyValidateDiv']) }; if(document.all['SpecialJourneyTableDiv'] != null) { HideDisplay(document.all['SpecialJourneyTableDiv']) };"">Deseo registrar la información en línea</A><BR /><BR />"
						Response.Write "<DIV NAME=""SpecialJourneyFormDiv"" ID=""SpecialJourneyFormDiv"">"
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							'Response.Write "function CheckSpecialJourneyFields(oForm) {" & vbNewLine
							'        Response.Write "if (oForm) {" & vbNewLine
							'            If B_ISSSTE Then
							'	            Response.Write "if (oForm.EmployeeID.value.length == 0) {" & vbNewLine
							'		            Response.Write "alert('Favor de introducir el número de empleado.');" & vbNewLine
							'		            Response.Write "oForm.EmployeeID.focus();" & vbNewLine
							'		            Response.Write "return false;" & vbNewLine
							'	            Response.Write "}" & vbNewLine
							'            End If
							'        Response.Write "}" & vbNewLine
							'Response.Write "} // End of CheckSpecialJourneyFields" & vbNewLine
						Response.Write "//--></SCRIPT>" & vbNewLine

						If iSpecialJourneyType = 1 Then ' Internos
							If Len(oRequest("EmployeeID").Item) > 0 Then
								lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
								If lErrorNumber = 0 Then
									If VerifyRequerimentsForEmployeesSpecialJourneys(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
										Call DisplayInternalSpecialJourneyForm(oRequest, oADODBConnection, aSpecialJourneyComponent, iSpecialJourneyType, sErrorDescription)
									Else
										lErrorNumber = -1
										Call DisplayErrorMessage("Advertencia", "Este empleado no cumple con las especificaciones de la matríz de guardias y suplencias.")
										Call DisplayAnotherJourneyForm(oRequest, oADODBConnection, "SpecialJourney.asp", -1, 10, iSpecialJourneyType, sAltDescription, sDescription, sErrorDescription)
									End If
								Else
									Call DisplayErrorMessage("Advertencia", "Este empleado no se encuentra registrado en el sistema.")
									Call DisplayAnotherJourneyForm(oRequest, oADODBConnection, "SpecialJourney.asp", -1, 10, iSpecialJourneyType, sAltDescription, sDescription, sErrorDescription)
								End If
							Else
								Call DisplayInternalSpecialJourneyForm(oRequest, oADODBConnection, aSpecialJourneyComponent, iSpecialJourneyType, sErrorDescription)
							End If
						ElseIf iSpecialJourneyType = 2 Then ' Externos
							If Len(oRequest("RFC").Item) > 0 Then
								lErrorNumber = CheckExistencyOfExternalEmployee(aSpecialJourneyComponent, sErrorDescription)
								If lErrorNumber = 0 Then
									If VerifyRequerimentsForExternalSpecialJourneys(oADODBConnection, aSpecialJourneyComponent, sErrorDescription) Then
										Call DisplayExternalSpecialJourneyForm(oRequest, oADODBConnection, aSpecialJourneyComponent, iSpecialJourneyType, sErrorDescription)
									Else
										lErrorNumber = -1
										Call DisplayErrorMessage("Advertencia", "Este colaborador externo no cumple con los requerimientos para registro de guardias.")
										Call DisplayAnotherJourneyForm(oRequest, oADODBConnection, "SpecialJourney.asp", -1, 10, iSpecialJourneyType, sAltDescription, sDescription, sErrorDescription)
									End If
								Else
									Call DisplayErrorMessage("Advertencia", "Este colaborador externo no ha sido registrado aún, por lo que se requiere sea registrado en el padrón.")
									Call DisplayAnotherJourneyForm(oRequest, oADODBConnection, "SpecialJourney.asp", -1, 10, iSpecialJourneyType, sAltDescription, sDescription, sErrorDescription)
								End If
							Else
								Call DisplayExternalSpecialJourneyForm(oRequest, oADODBConnection, aSpecialJourneyComponent, iSpecialJourneyType, sErrorDescription)
							End If
						Else
							'If bSearchForm Then
							If True Then
								lErrorNumber = Display423SearchForm(oRequest, oADODBConnection, iSpecialJourneyType, sErrorDescription)
							ElseIf Len(oRequest("DoSearch").Item) > 0 Then
								If iSectionID = 425 Then 'RQ
									If StrComp(oRequest("Internal").Item, "1", vbBinaryCompare) = 0 Then
										aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("2,3,4,5,6,10,17,18,19,28,31", ",")
									Else
										aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("3,4,5,6,10,17,18,19,28,31", ",")
									End If
								End If
								If Len(aCatalogComponent(S_QUERY_CONDITION_CATALOG)) = 0 Then aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " And (AppliedDate In (Select PayrollID From Payrolls Where IsActive_5<>0))"
								lErrorNumber = DisplayCatalogsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, True, aCatalogComponent, sErrorDescription)
								aCatalogComponent(S_QUERY_CONDITION_CATALOG) = ""
								If lErrorNumber = L_ERR_NO_RECORDS Then
									Call DisplayErrorMessage("Búsqueda vacía", "No existen registros que cumplan con los criterios de la búsqueda")
									Response.Write "<BR />"
									lErrorNumber = Display423SearchForm(oRequest, oADODBConnection, sErrorDescription)
								End If
							End If
						End If

						Response.Write "<BR /><BR />"
						Response.Write "</DIV>"

						Response.Write "<IMG SRC=""Images/Crcl2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['SpecialJourneyFormDiv'] != null) { HideDisplay(document.all['SpecialJourneyFormDiv']) }; ShowDisplay(document.all['SpecialJourneyUploadDiv']); if(document.all['SpecialJourneyValidateDiv'] != null) { HideDisplay(document.all['SpecialJourneyValidateDiv']) }; if(document.all['SpecialJourneyTableDiv'] != null) { HideDisplay(document.all['SpecialJourneyTableDiv']) };"">Deseo subir la información a través de un archivo</A><BR /><BR />"
						Response.Write "<DIV NAME=""SpecialJourneyUploadDiv"" ID=""SpecialJourneyUploadDiv"" STYLE=""display: none"">"
							Response.Write "<FORM NAME=""SpecialJourneyUploadFrm"" ID=""SpecialJourneyUploadFrm"" METHOD=""POST"" onSubmit=""return true"">"

								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""2"" />"
								'Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ActionHdn"" VALUE=""" & lReasonID & """ />"
								'Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Success"" ID=""ActionHdn"" VALUE=""" & lSuccess & """ />"
								'Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & lEmployeeID & """ />"
								'Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ErrorDescription"" ID=""EmployeeIDHdn"" VALUE=""" & sError & """ />"

								Call DisplayInstructionsMessage("Paso 1. Introduzca el archivo a utilizar", "<BLOCKQUOTE><OL>" & _
																	"<LI>Abra el documento de Excel con la información que desea subir.</LI>" & _
																	"<LI>Copie únicamente las celdas que contienen la información deseada.</LI>" & _
																	"<LI>Pegue dicha información en la caja de texto.</LI>" & _
																	"<LI><B>O seleccione el archivo de texto que contiene la información a subir.</B></LI>" & _
																"</OL></BLOCKQUOTE>")
								Response.Write "<BR />"
								Response.Write "<B>Para este concepto se requiere: </B> RFC, Puesto, Adscripción, Servicio, Nivel/subnivel, Horas laboradas, Folio de autorización, Fecha de inicio, Fecha de fin, Turno, Días/horas reportadas, Movimiento, Motivo, Quincena de aplicación, y Comentarios(opcional)."
								Response.Write "<BR />"
								Response.Write "<BR />"

								Response.Write "<TEXTAREA NAME=""RawData"" ID=""RawDataTxtArea"" ROWS=""10"" COLS=""119"" CLASS=""TextFields"" onChange=""bReady = (this.value != '')""></TEXTAREA><BR /><BR />"
								Response.Write "<INPUT TYPE=""SUBMIT"" VALUE=""Continuar"" CLASS=""Buttons"" />"

							Response.Write "</FORM>"
							Response.Write "<IFRAME SRC=""BrowserFileForInfo.asp?Action=" & oRequest("Action").Item & "&UserID=" & aLoginComponent(N_USER_ID_LOGIN) & """ NAME=""UploadInfoIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""60""></IFRAME>"
							Response.Write "<BR />"
						Response.Write "</DIV>"

						Response.Write "<IMG SRC=""Images/Crcl3.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['SpecialJourneyFormDiv'] != null) { HideDisplay(document.all['SpecialJourneyFormDiv']) }; if(document.all['SpecialJourneyUploadDiv'] != null) { HideDisplay(document.all['SpecialJourneyUploadDiv']) }; ShowDisplay(document.all['SpecialJourneyValidateDiv']); if(document.all['SpecialJourneyTableDiv'] != null) { HideDisplay(document.all['SpecialJourneyTableDiv']) };""><FONT FACE=""Arial"" SIZE=""2"">Registros en proceso de aplicación</FONT></A><BR /><BR />"
						Response.Write "<DIV NAME=""SpecialJourneyValidateDiv"" ID=""SpecialJourneyValidateDiv"" STYLE=""display: none"">"
							Response.Write "<FORM NAME=""SpecialJourneyValidateFrm"" ID=""SpecialJourneyValidateFrm"" METHOD=""POST"" onSubmit=""return bReady"">"
							'Response.Write "<IFRAME SRC=""BrowserFileForInfo.asp?Action=SpecialJourney&UserID=" & aLoginComponent(N_USER_ID_LOGIN) & """ NAME=""UploadInfoIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""60""></IFRAME>"

								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & sAction & """ />"

								If iSpecialJourneyType = 1 Then ' Internos
									If Len(oRequest("EmployeeID").Item) > 0 Then
										If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) And (Request.Cookies("SIAP_SectionID") <> 7) Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""AuthorizationFile"" ID=""ModifyBtn"" VALUE=""Aplicar registros en proceso"" CLASS=""Buttons""/>"
									Else
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_EMPLOYEE) & """ />"
										If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) And (Request.Cookies("SIAP_SectionID") <> 7) Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""AuthorizationFile"" ID=""ModifyBtn"" VALUE=""Aplicar registros para este empleado"" CLASS=""Buttons""/>"
									End If
								ElseIf iSpecialJourneyType = 2 Then ' Externos
									If Len(oRequest("RFC").Item) = 0 Then
										If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) And (Request.Cookies("SIAP_SectionID") <> 7) Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""AuthorizationFile"" ID=""ModifyBtn"" VALUE=""Aplicar registros en proceso"" CLASS=""Buttons""/>"
									Else
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_EMPLOYEE) & """ />"
										If (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ValidacionDeMovimientos & ",", vbBinaryCompare) > 0) And (Request.Cookies("SIAP_SectionID") <> 7) Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""AuthorizationFile"" ID=""ModifyBtn"" VALUE=""Aplicar registros para este empleado"" CLASS=""Buttons""/>"
									End If
								End If
								Response.Write "<BR /><BR />"

								aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ACTIVE) = 0
								If (CInt(oRequest("RowsType").Item) = 1) Then
									iStarPage = 0
								End If
								If iSpecialJourneyType = 1 Then
									lErrorNumber = DisplayInternalSpecialJourneyTable(oRequest, oADODBConnection, False, iStarPage, aSpecialJourneyComponent, iSpecialJourneyType, sErrorDescription)
								Else
									lErrorNumber = DisplayExternalSpecialJourneyTable(oRequest, oADODBConnection, False, iStarPage, aSpecialJourneyComponent, iSpecialJourneyType, sErrorDescription)
								End If
								If lErrorNumber <> 0 Then
									Response.Write "<BR />"
									Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
									lErrorNumber = 0
									sErrorDescription = ""
								End If

							Response.Write "</FORM>"
						Response.Write "</DIV>"

						Response.Write "<IMG SRC=""Images/Crcl4.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['SpecialJourneyFormDiv'] != null) { HideDisplay(document.all['SpecialJourneyFormDiv']) }; if(document.all['SpecialJourneyUploadDiv'] != null) { HideDisplay(document.all['SpecialJourneyUploadDiv']) }; if(document.all['SpecialJourneyValidateDiv'] != null) { HideDisplay(document.all['SpecialJourneyValidateDiv']) }; ShowDisplay(document.all['SpecialJourneyTableDiv']);""><FONT FACE=""Arial"" SIZE=""2"">Historial de Registros existentes</FONT></A><BR /><BR />"
						Response.Write "<DIV NAME=""SpecialJourneyTableDiv"" ID=""SpecialJourneyTableDiv"" STYLE=""display: none"">"
							Response.Write "<FORM NAME=""SpecialJourneyTableFrm"" ID=""SpecialJourneyTableFrm"" METHOD=""POST"" onSubmit=""return bReady"">"

								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & sAction & """ />"
								aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ACTIVE) = 1
								If (Len(oRequest("StartPage").Item) > 0) Then
									If (CInt(oRequest("RowsType").Item) = 0) Then
										iStarPage = 0
									Else
 										iStarPage = CInt(oRequest("StartPage").Item)
									End If
								End If
								If iSpecialJourneyType = 1 Then
									lErrorNumber = DisplayInternalSpecialJourneyTable(oRequest, oADODBConnection, False, iStarPage, aSpecialJourneyComponent, iSpecialJourneyType, sErrorDescription)
								Else
									lErrorNumber = DisplayExternalSpecialJourneyTable(oRequest, oADODBConnection, False, iStarPage, aSpecialJourneyComponent, iSpecialJourneyType, sErrorDescription)
								End If
								If lErrorNumber <> 0 Then
									Response.Write "<BR />"
									Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
									lErrorNumber = 0
									sErrorDescription = ""
								End If

							Response.Write "</FORM>"
						Response.Write "</DIV>"

						If (Len(oRequest("StartPage").Item) > 0) Then
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							If CInt(oRequest("RowsType").Item) = 0 Then
								Response.Write "if(document.all['SpecialJourneyFormDiv'] != null) { HideDisplay(document.all['SpecialJourneyFormDiv']) }; if(document.all['SpecialJourneyUploadDiv'] != null) { HideDisplay(document.all['SpecialJourneyUploadDiv']) }; ShowDisplay(document.all['SpecialJourneyValidateDiv']); if(document.all['SpecialJourneyTableDiv'] != null) { HideDisplay(document.all['SpecialJourneyTableDiv']) };" & vbNewLine
							Else
								Response.Write "if(document.all['SpecialJourneyFormDiv'] != null) { HideDisplay(document.all['SpecialJourneyFormDiv']) }; if(document.all['SpecialJourneyUploadDiv'] != null) { HideDisplay(document.all['SpecialJourneyUploadDiv']) }; if(document.all['SpecialJourneyValidateDiv'] != null) { HideDisplay(document.all['SpecialJourneyValidateDiv']) }; ShowDisplay(document.all['SpecialJourneyTableDiv']);" & vbNewLine
							End If
							Response.Write "//--></SCRIPT>" & vbNewLine
						End If

					Case 427
						Response.Write "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0""><TR>" & vbNewLine
							Response.Write "<TD WIDTH=""700"" VALIGN=""TOP"">"
								Response.Write "<IMG SRC=""Images/Crcl1.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: ShowDisplay(document.all['ConceptInfoFormDiv']); if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) };""><FONT FACE=""Arial"" SIZE=""2"">Consulta los registros activos para el personal externo</FONT></A><BR /><BR />"
								Response.Write "<DIV NAME=""ConceptInfoFormDiv"" ID=""ConceptInfoFormDiv"" STYLE=""display: none"">"
									Response.Write "<FORM NAME=""ConceptInfoFrm"" ID=""ConceptInfoFrm"" ACTION=""SpecialJourney.asp"" METHOD=""POST"">"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""SpecialJourney"" />"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartPage"" ID=""StartPageHdn"" VALUE=""1"" />"
										Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" /><FONT FACE=""Arial"" SIZE=""2""> Mostrar empleados con RFC: </FONT>"
										Response.Write "<INPUT TYPE=""TEXT"" NAME=""RFC"" ID=""RFCTxt"" SIZE=""13"" MAXLENGTH=""13"" VALUE=""" & CleanStringForHTML(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_RFC)) & """ CLASS=""TextFields"" />"
										Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""SUBMIT"" VALUE=""Consultar registros"" CLASS=""Buttons""><BR />"
										'Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""800"" HEIGHT=""1"" /><BR />"
									Response.Write "</FORM>"
								If Len(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_RFC)) > 0 Then aSpecialJourneyComponent(S_QUERY_CONDITION_SPECIAL_JOURNEY) = aSpecialJourneyComponent(S_QUERY_CONDITION_SPECIAL_JOURNEY) & " And (ExternalSpecialJourneys.RFC like '%" & aSpecialJourneyComponent(S_SPECIAL_JOURNEY_RFC) & "%')"
								Response.Write "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
									Response.Write "<TR><TD>" & vbNewLine
										aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ACTIVE) = 1
										lErrorNumber = DisplayExternalEmployeeTable(oRequest, oADODBConnection, DISPLAY_NOTHING, False, aSpecialJourneyComponent, sErrorDescription)
										If lErrorNumber <> 0 Then
											Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
											lErrorNumber = 0
											sErrorDescription = ""
										End If
									Response.Write "</TD></TR>" & vbNewLine
								Response.Write "</TABLE>" & vbNewLine
							Response.Write "</DIV>"

							Response.Write "<BR />"
							Response.Write "<IMG SRC=""Images/Crcl2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['ConceptInfoFormDiv'] != null) { HideDisplay(document.all['ConceptInfoFormDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; ShowDisplay(document.all['UploadValidateInfoFormDiv']);""><FONT FACE=""Arial"" SIZE=""2"">Consulta los registros que estan en proceso para el personal externo</FONT></A><BR /><BR />"
							Response.Write "<DIV NAME=""UploadValidateInfoFormDiv"" ID=""UploadValidateInfoFormDiv"" STYLE=""display: none"">"
								Response.Write "<FORM NAME=""ConceptInfoFrm"" ID=""ConceptInfoFrm"" ACTION=""UploadInfo.asp"" METHOD=""POST"">"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDHdn"" VALUE=""" & oRequest("EmployeeTypeID").Item & """ />"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartPage"" ID=""StartPageHdn"" VALUE=""1"" />"
									Response.Write "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
										Response.Write "<TR>"
											If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<TR><TD><INPUT TYPE=""SUBMIT"" NAME=""AuthorizationFile"" ID=""ModifyBtn"" VALUE=""Aplicar Movimientos Seleccionados"" CLASS=""Buttons""/></TD></TR>"
										Response.Write "</TR>"
									Response.Write "</TABLE>"
									Response.Write "<BR />"
									aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ACTIVE) = 0
									aSpecialJourneyComponent(S_QUERY_CONDITION_SPECIAL_JOURNEY) = ""
									lErrorNumber = DisplayExternalEmployeeTable(oRequest, oADODBConnection, DISPLAY_NOTHING, False, aSpecialJourneyComponent, sErrorDescription)
									If lErrorNumber <> 0 Then
										Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
										lErrorNumber = 0
										sErrorDescription = ""
									End If
								Response.Write "</FORM>"
							Response.Write "</DIV>"
						Response.Write "</TD>"
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							Response.Write "if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; ShowDisplay(document.all['ConceptInfoFormDiv']);" & vbNewLine
						Response.Write "//--></SCRIPT>" & vbNewLine

						Response.Write "<TD>&nbsp;</TD>"
						Response.Write "<TD BGCOLOR=""" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""1"" ><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
						Response.Write "<TD>&nbsp;</TD>"
						Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">" & vbNewLine

						aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) = aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EXTERNAL_ID)
						If iGlobalSectionID = 4 Then lErrorNumber = DisplayExternalJourneyForm(oRequest, oADODBConnection, GetASPFileName(""), aSpecialJourneyComponent, sErrorDescription)
						If False Then
							If lErrorNumber <> 0 Then
								Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
								lErrorNumber = 0
								Response.Write "<BR />"
							End If
							If Not IsEmpty(lSuccess) Then
								If lSuccess = 1 Then
									Call DisplayErrorMessage("Confirmación", "La operación fué ejecutada exitosamente.")
								Else
									Call DisplayErrorMessage("Error al realizar la operación", CStr(oRequest("sErrorDescription").Item))
								End If
							End If
						End If
						Response.Write "</FONT></TD>" & vbNewLine
						Response.Write "</TR></TABLE>" & vbNewLine
					Case 428
						Call DisplayErrorMessage("Advertencia", "Aquí registrar beneficiario(a)s de pensión!!!")
				End Select
            Else
                Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 1. </B>Introduzca el archivo a utilizar.<BR /><BR />"
				Select Case iStep
					Case 2
						Call DisplayEmployeesSpecialJourneysColumns(lReasonID, sFileName, sErrorDescription)
					Case 3
						Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Identifique las columnas del archivo.<BR /><BR />"
						lErrorNumber = UploadEmployeesSpecialJourneysFile(oADODBConnection, EMPLOYEES_EXTRAHOURS, sFileName, sErrorDescription)
						Response.Write "<BR />"
						If lErrorNumber = 0 Then
							Call DisplayErrorMessage("Confirmación", "Las incidencias fueron registradas con éxito.")
						Else
							Call DisplayErrorMessage("Error al registrar las incidencias", sErrorDescription)
							lErrorNumber = 0
							sErrorDescription = ""
						End If
				End Select
            End If
%>
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>