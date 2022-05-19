<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Server.ScriptTimeout = 72000
%>
<!-- #include file="Libraries/GlobalVariables.asp" -->
<!-- #include file="Libraries/LoginComponent.asp" -->
<!-- #include file="Libraries/GraphComponent.asp" -->
<!-- #include file="Libraries/AbsenceComponent.asp" -->
<!-- #include file="Libraries/AlimonyTypeComponent.asp" -->
<!-- #include file="Libraries/EmployeeComponent.asp" -->
<!-- #include file="Libraries/EmployeeSupportLib.asp" -->
<!-- #include file="Libraries/ReportsLib.asp" -->
<!-- #include file="Libraries/ReportComponent.asp" -->
<!-- #include file="Libraries/ZIPLibrary.asp" -->
<%
Dim oItem
Dim iIndex
Dim aTemplateValue
Dim sNames
Dim bOnlyExport
Dim asPeriods
Dim band

If B_ISSSTE Then
Else
	If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REPORTS_PERMISSIONS) = N_REPORTS_PERMISSIONS Then
	Else
		Response.Redirect "AccessDenied.asp?Permission=" & N_REPORTS_PERMISSIONS
	End If
End If

Call InitializeReportsComponent(oRequest, aReportsComponent)


Response.Cookies("SoS_SectionID") = 1000 + iGlobalSectionID
Select Case iGlobalSectionID
	Case 1
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Personal"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
	Case 19
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Reportes"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
	Case 2
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Prestaciones"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
	Case 21
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Terceros institucionales"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
	Case 22
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Prestaciones e incidencias"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
	Case 23
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Registro de terceros"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
	Case 3
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Desarrollo Humano"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case 33
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Plantillas de personal"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case 34
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Selección de personal"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case 35
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Capacitación"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case 36
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Planeación de recursos humanos"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case 4
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Informática"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
	Case 42
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Empleados"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
	Case 49
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Reportes"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
	Case 5
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Presupuesto"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = BUDGET_TOOLBAR
	Case 6
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Departamento Técnico"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = REPORTS_TOOLBAR
	Case 64
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Reportes"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = REPORTS_TOOLBAR
	Case 7
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Desconcentrados"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
	Case 71
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Empleados"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
	Case 72
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Nóminas"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
	Case 73
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Reportes"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
End Select
aHeaderComponent(S_TITLE_NAME_HEADER) = GetReportNameByConstant(aReportsComponent(N_ID_REPORTS))
bWaitMessage = True

Select Case aReportsComponent(N_ID_REPORTS)
	Case LOGS_HISTORY_REPORTS, ISSSTE_1001_REPORTS, ISSSTE_1002_REPORTS, ISSSTE_1004_REPORTS, ISSSTE_1007_REPORTS, ISSSTE_1009_REPORTS, ISSSTE_1012_REPORTS, ISSSTE_1013_REPORTS, ISSSTE_1014_REPORTS, ISSSTE_1015_REPORTS, ISSSTE_1018_REPORTS, ISSSTE_1019_REPORTS, ISSSTE_1020_REPORTS, ISSSTE_1021_REPORTS, ISSSTE_1022_REPORTS, ISSSTE_1024_REPORTS, ISSSTE_1027_REPORTS, ISSSTE_1028_REPORTS, ISSSTE_1031_REPORTS, ISSSTE_1032_REPORTS, ISSSTE_1033_REPORTS, ISSSTE_1034_REPORTS, ISSSTE_1035_REPORTS, ISSSTE_1101_REPORTS, ISSSTE_1102_REPORTS, ISSSTE_1103_REPORTS, ISSSTE_1104_REPORTS, ISSSTE_1105_REPORTS, ISSSTE_1106_REPORTS, ISSSTE_1107_REPORTS, ISSSTE_1108_REPORTS, ISSSTE_1109_REPORTS, ISSSTE_1110_REPORTS, ISSSTE_1111_REPORTS, ISSSTE_1112_REPORTS, ISSSTE_1113_REPORTS, ISSSTE_1114_REPORTS, ISSSTE_1115_REPORTS, ISSSTE_1116_REPORTS, ISSSTE_1117_REPORTS, ISSSTE_1118_REPORTS, ISSSTE_1119_REPORTS, ISSSTE_1152_REPORTS, ISSSTE_1200_REPORTS, ISSSTE_1201_REPORTS, ISSSTE_1202_REPORTS, ISSSTE_1203_REPORTS, ISSSTE_1204_REPORTS, ISSSTE_1205_REPORTS, ISSSTE_1206_REPORTS, ISSSTE_1207_REPORTS, ISSSTE_1208_REPORTS, ISSSTE_1222_REPORTS, ISSSTE_1334_REPORTS, ISSSTE_1335_REPORTS, ISSSTE_1336_REPORTS, ISSSTE_1337_REPORTS, ISSSTE_1339_REPORTS, ISSSTE_1354_REPORTS, ISSSTE_1356_REPORTS, ISSSTE_1401_REPORTS, ISSSTE_1411_REPORTS, ISSSTE_1412_REPORTS, ISSSTE_1413_REPORTS, ISSSTE_1417_REPORTS, ISSSTE_1420_REPORTS, ISSSTE_1421_REPORTS, ISSSTE_1422_REPORTS, ISSSTE_1423_REPORTS, ISSSTE_1424_REPORTS, ISSSTE_1425_REPORTS, ISSSTE_1427_REPORTS, ISSSTE_1428_REPORTS, ISSSTE_1429_REPORTS, ISSSTE_1430_REPORTS, ISSSTE_1471_REPORTS, ISSSTE_1472_REPORTS, ISSSTE_1473_REPORTS, ISSSTE_1490_REPORTS, ISSSTE_1493_REPORTS, ISSSTE_1494_REPORTS, ISSSTE_1499_REPORTS, ISSSTE_1581_REPORTS, ISSSTE_1582_REPORTS, ISSSTE_1583_REPORTS, ISSSTE_1584_REPORTS, ISSSTE_1605_REPORTS, ISSSTE_1606_REPORTS, ISSSTE_1702_REPORTS, ISSSTE_1703_REPORTS, ISSSTE_1704_REPORTS, ISSSTE_2420_REPORTS, ISSSTE_2421_REPORTS, ISSSTE_2422_REPORTS, ISSSTE_2423_REPORTS, ISSSTE_2427_REPORTS, ISSSTE_2428_REPORTS, ISSSTE_2429_REPORTS, ISSSTE_2430_REPORTS, ISSSTE_1209_REPORTS, ISSSTE_1431_REPORTS, ISSSTE_2431_REPORTS, ISSSTE_1432_REPORTS, ISSSTE_2432_REPORTS
		If aReportsComponent(N_STEP_REPORTS) > 1 Then aReportsComponent(B_READY_REPORTS) = True
	Case AREAS_COUNT_REPORTS, EMPLOYEES_COUNT_REPORTS, JOBS_COUNT_REPORTS, AREAS_LIST_REPORTS, EMPLOYEES_LIST_REPORTS, JOBS_LIST_REPORTS, SPECIAL_JOBS_LIST_REPORTS, ISSSTE_1364_REPORTS, ISSSTE_1503_REPORTS, ISSSTE_1504_REPORTS, ISSSTE_1609_REPORTS, ISSSTE_1701_REPORTS, JOBS_LIST_BY_MODIFY_DATE
		If aReportsComponent(N_STEP_REPORTS) > 2 Then aReportsComponent(B_READY_REPORTS) = True
	Case ISSSTE_1561_REPORTS, ISSSTE_1562_REPORTS, ISSSTE_1563_REPORTS, ISSSTE_1571_REPORTS
		If aReportsComponent(N_STEP_REPORTS) > 2 Then
			aReportsComponent(B_READY_REPORTS) = True
			bOnlyExport = True
		End If
	Case ISSSTE_1003_REPORTS, ISSSTE_1005_REPORTS, ISSSTE_1006_REPORTS, ISSSTE_1010_REPORTS, ISSSTE_1011_REPORTS, ISSSTE_1026_REPORTS, ISSSTE_1029_REPORTS, ISSSTE_1030_REPORTS, ISSSTE_1100_REPORTS, ISSSTE_1151_REPORTS, ISSSTE_1153_REPORTS, ISSSTE_1154_REPORTS, ISSSTE_1155_REPORTS, ISSSTE_1157_REPORTS, ISSSTE_1211_REPORTS, ISSSTE_1311_REPORTS, ISSSTE_1338_REPORTS, ISSSTE_1339_REPORTS, ISSSTE_1340_REPORTS, ISSSTE_1371_REPORTS, ISSSTE_1372_REPORTS, ISSSTE_1373_REPORTS, ISSSTE_1374_REPORTS, ISSSTE_1404_REPORTS, ISSSTE_1414_REPORTS, ISSSTE_1415_REPORTS, ISSSTE_1416_REPORTS, ISSSTE_1426_REPORTS, ISSSTE_1433_REPORTS, ISSSTE_1434_REPORTS, ISSSTE_1470_REPORTS, ISSSTE_1474_REPORTS, ISSSTE_1475_REPORTS, ISSSTE_1476_REPORTS, ISSSTE_1477_REPORTS, ISSSTE_1478_REPORTS, ISSSTE_1491_REPORTS, ISSSTE_1492_REPORTS, ISSSTE_1603_REPORTS, ISSSTE_1604_REPORTS, ISSSTE_1607_REPORTS, ISSSTE_1608_REPORTS, ISSSTE_1610_REPORTS, ISSSTE_1611_REPORTS, ISSSTE_1612_REPORTS, ISSSTE_2426_REPORTS, ISSSTE_1431_REPORTS, ISSSTE_2431_REPORTS
		If aReportsComponent(N_STEP_REPORTS) > 1 Then
			aReportsComponent(B_READY_REPORTS) = False
			aReportsComponent(B_HIDE_CONTINUE_REPORTS) = True
		End If
	Case ISSSTE_1400_REPORTS
		If aReportsComponent(N_STEP_REPORTS) > 3 Then
			aReportsComponent(B_READY_REPORTS) = False
			aReportsComponent(B_HIDE_CONTINUE_REPORTS) = True
		End If
	Case ISSSTE_1403_REPORTS
		If aReportsComponent(N_STEP_REPORTS) > 3 Then
			aReportsComponent(B_READY_REPORTS) = False
			aReportsComponent(B_HIDE_CONTINUE_REPORTS) = True
		End If
	Case ISSSTE_1613_REPORTS
		If aReportsComponent(N_STEP_REPORTS) > 1 Then aReportsComponent(B_READY_REPORTS) = True
End Select
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/Export.js"></SCRIPT>
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%If bOnlyExport Then
			aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Exportar a Excel",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=Reports&Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True),_
				Array("Exportar a Word",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=Reports&Word=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
			)
			aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 793
			aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
			aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 200
		ElseIf aReportsComponent(B_READY_REPORTS) Then
			aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Guardar el reporte",_
					  "",_
					  "", "javascript: document.SaveReportFrm.submit();", True),_
				Array("Modificar el reporte",_
					  "",_
					  "", "javascript: document.ModifyReportFrm.submit();", True),_
				Array("Exportar a Excel",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=Reports&Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (InStr(1, ",1001,1613,1007,1207,1208,1490,", aReportsComponent(N_ID_REPORTS), vbBinaryCompare) = 0)),_
				Array("Exportar a Excel",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=Reports&Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & oRequest & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (InStr(1, ",1613,", aReportsComponent(N_ID_REPORTS), vbBinaryCompare) > 0)),_
				Array("Exportar a Word",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=Reports&Word=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (InStr(1, ",1001,1007,1207,1208,1490,", aReportsComponent(N_ID_REPORTS), vbBinaryCompare) > 0)),_
				Array("Imprimir",_
					  "",_
					  "", "javascript: SendReportToPrint('ReportDiv', '" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "')", False)_
			)
			aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 793
			aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
			aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 200
		End If%>
		<!-- #include file="_Header.asp" -->
		<%If aReportsComponent(N_ID_REPORTS) = 0 Then
			If B_ISSSTE Then
				Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> > " & GetReportPathByConstant(iGlobalSectionID, aReportsComponent(N_ID_REPORTS)) & "<BR /><BR />"
			Else
				Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> > <B>Reportes</B><BR /><BR /><BR />"
			End If
			aMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Reportes guardados",_
					  "Reportes generados por otros usuarios a partir de plantillas y filtros.",_
					  "Images/MnLeftArrows.gif", "SavedReport.asp?ReportType=1", True),_
				Array("Historial de entradas al sistema",_
					  "Obtenga un conteo de las entradas de los usuarios en el sistema, así como el detalle de las entradas por cada usuario.",_
					  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=" & LOGS_HISTORY_REPORTS, (InStr(1, aLoginComponent(S_ACCESS_KEY_LOGIN), "vac", vbBinaryCompare) = 1)),_
				Array("<LINE />",_
					  "",_
					  "", "", False)_
			)
			aMenuComponent(B_USE_DIV_MENU) = True
			Response.Write "<TABLE WIDTH=""900"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Call DisplayMenuInThreeSmallColumns(aMenuComponent)
			Response.Write "</TABLE>"
		Else
			If B_ISSSTE Then
				Select Case aReportsComponent(N_ID_REPORTS)
					Case ISSSTE_1028_REPORTS
						Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> ><A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> ><A HREF=""Main_ISSSTE.asp?SectionID=491"">Ejercicio bimestral del SAR</A> > <B> Reporte de cifras del bimestre </B>"
					Case ISSSTE_1031_REPORTS
						Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> ><A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> ><A HREF=""Main_ISSSTE.asp?SectionID=491"">Ejercicio bimestral del SAR</A> > <B> Reportes de movimientos del bimestre </B>"
					Case ISSSTE_1032_REPORTS
						Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> ><A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> ><A HREF=""Main_ISSSTE.asp?SectionID=491"">Ejercicio bimestral del SAR</A> > <B> Reportes de disperción por UA </B>"
					Case ISSSTE_1033_REPORTS
						Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> ><A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> ><A HREF=""Main_ISSSTE.asp?SectionID=491"">Ejercicio bimestral del SAR</A> > <B> Reporte de aportaciones </B>"
					Case ISSSTE_1034_REPORTS
						Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> ><A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> ><A HREF=""Main_ISSSTE.asp?SectionID=491"">Ejercicio bimestral del SAR</A> > <B> Control y distribución de comprobantes de abono en cuenta </B>"
					Case ISSSTE_1035_REPORTS
						Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> ><A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> ><A HREF=""Main_ISSSTE.asp?SectionID=491"">Ejercicio bimestral del SAR</A> > <B> Generar proceso de altas, bajas y cambios </B>"
					Case Else 
						Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> > " & GetReportPathByConstant(iGlobalSectionID, aReportsComponent(N_ID_REPORTS)) & "<BR /><BR />"
				End Select 
			Else
				Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> > <A HREF=""Reports.asp"">Reportes</A>"
				If Len(oRequest("Saved").Item) > 0 Then Response.Write " > <A HREF=""SavedReport.asp?ReportType=" & oRequest("ReportType").Item & """>Reportes guardados</A>"
				Response.Write " > <B>" & GetReportNameByConstant(aReportsComponent(N_ID_REPORTS)) & "</B><BR /><BR />"
			End If

			Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><FORM NAME=""ReportFrm"" ID=""ReportFrm"" ACTION=""Reports.asp"" METHOD=""POST"" onSubmit="""
				Select Case aReportsComponent(N_ID_REPORTS)
					Case AREAS_COUNT_REPORTS, EMPLOYEES_COUNT_REPORTS, JOBS_COUNT_REPORTS, AREAS_LIST_REPORTS, EMPLOYEES_LIST_REPORTS, JOBS_LIST_REPORTS, ISSSTE_1364_REPORTS, ISSSTE_1504_REPORTS, ISSSTE_1701_REPORTS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								Response.Write "if (document.ReportFrm.Template.options.length > 0) SelectAllItemsFromList(document.ReportFrm.Template); else {alert('Es necesario que incluya al menos un campo en el reporte'); return false;}"
						End Select
					Case ISSSTE_1155_REPORTS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								Response.Write "if (this.Limit1.value != '') {this.Limit1.value = this.Limit1.value.replace(/" & NUMERIC_SEPARATOR & "/gi, ''); "
								Response.Write "if (! CheckFloatValue(this.Limit1, 'el límite inferior', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0)) {return false;}} "
								Response.Write "if (this.Limit2.value != '') {this.Limit2.value = this.Limit2.value.replace(/" & NUMERIC_SEPARATOR & "/gi, ''); "
								Response.Write "if (! CheckFloatValue(this.Limit2, 'el límite superior', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0)) {return false;}} return true;"
						End Select
					Case ISSSTE_1561_REPORTS, ISSSTE_1562_REPORTS, ISSSTE_1563_REPORTS, ISSSTE_1571_REPORTS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								Response.Write "return CheckRadioSelection(document.ReportFrm.RecordID);"
						End Select
					Case ISSSTE_1203_REPORTS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								Response.Write "return VerifyEmployeeNumber();"
						End Select
					Case Else
				End Select
			Response.Write """>"
				Select Case aReportsComponent(N_ID_REPORTS)
					Case LOGS_HISTORY_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_LOG_DATE_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 0 Then Response.Write "<DONT_EXPORT>"
									lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 0 Then Response.Write "</DONT_EXPORT>"
								If lErrorNumber = 0 Then
									lErrorNumber = DisplayLogCount(oRequest, False, sErrorDescription)
									Response.Write "<BR />"
									lErrorNumber = DisplayLogHistoryList(oRequest, False, sErrorDescription)
								End If
						End Select
					Case AREAS_COUNT_REPORTS
						sFlags = L_AREA_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_TYPE_FLAGS & "," & L_CONFINE_TYPE_FLAGS & "," & L_CENTER_TYPE_FLAGS & "," & L_CENTER_SUBTYPE_FLAGS & "," & L_ATTENTION_LEVEL_FLAGS & "," & L_ZONE_FLAGS & "," & L_ECONOMIC_ZONE_FLAGS & "," & L_AREA_STATUS_FLAGS & "," & L_AREA_ACTIVE_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								sFlags = L_AREA_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_TYPE_FLAGS & "," & L_CONFINE_TYPE_FLAGS & "," & L_CENTER_TYPE_FLAGS & "," & L_CENTER_SUBTYPE_FLAGS & "," & L_ATTENTION_LEVEL_FLAGS & "," & L_ZONE_FLAGS & "," & L_ECONOMIC_ZONE_FLAGS & "," & L_AREA_STATUS_FLAGS & "," & L_AREA_ACTIVE_FLAGS
								lErrorNumber = DisplayReportTemplateForm(sFlags, sErrorDescription)
							Case 2
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 3
								If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 0 Then Response.Write "<DONT_EXPORT>"
									lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 0 Then Response.Write "</DONT_EXPORT>"
								If lErrorNumber = 0 Then
									lErrorNumber = DisplayAreasCount(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case EMPLOYEES_COUNT_REPORTS
						sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_NAME_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_JOURNEY_FLAGS & "," & L_SHIFT_FLAGS & "," & L_LEVEL_FLAGS & "," & L_EMPLOYEE_STATUS_FLAGS & "," & L_SOCIAL_SECURITY_NUMBER_FLAGS & "," & L_EMPLOYEE_BIRTH_FLAGS & "," & L_EMPLOYEE_RFC_FLAGS & "," & L_EMPLOYEE_CURP_FLAGS & "," & L_EMPLOYEE_GENDER_FLAGS & "," & L_EMPLOYEE_ACTIVE_FLAGS & "," & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								sFlags = L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_JOURNEY_FLAGS & "," & L_SHIFT_FLAGS & "," & L_LEVEL_FLAGS & "," & L_EMPLOYEE_STATUS_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_EMPLOYEE_GENDER_FLAGS & "," & L_EMPLOYEE_ACTIVE_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS
								lErrorNumber = DisplayReportTemplateForm(sFlags, sErrorDescription)
							Case 2
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 3
								If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 0 Then Response.Write "<DONT_EXPORT>"
									lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 0 Then Response.Write "</DONT_EXPORT>"
								If lErrorNumber = 0 Then
									lErrorNumber = DisplayEmployeesCount(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case JOBS_COUNT_REPORTS
						sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_NAME_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_JOURNEY_FLAGS & "," & L_SHIFT_FLAGS & "," & L_LEVEL_FLAGS & "," & L_EMPLOYEE_STATUS_FLAGS & "," & L_SOCIAL_SECURITY_NUMBER_FLAGS & "," & L_EMPLOYEE_BIRTH_FLAGS & "," & L_EMPLOYEE_RFC_FLAGS & "," & L_EMPLOYEE_CURP_FLAGS & "," & L_EMPLOYEE_GENDER_FLAGS & "," & L_EMPLOYEE_ACTIVE_FLAGS & "," & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS & "," & L_JOB_STATUS_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								sFlags = L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_JOURNEY_FLAGS & "," & L_SHIFT_FLAGS & "," & L_LEVEL_FLAGS & "," & L_EMPLOYEE_STATUS_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_EMPLOYEE_GENDER_FLAGS & "," & L_EMPLOYEE_ACTIVE_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS & "," & L_JOB_STATUS_FLAGS
								lErrorNumber = DisplayReportTemplateForm(sFlags, sErrorDescription)
							Case 2
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 3
								If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 0 Then Response.Write "<DONT_EXPORT>"
									lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 0 Then Response.Write "</DONT_EXPORT>"
								If lErrorNumber = 0 Then
									lErrorNumber = DisplayJobsCount(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case AREAS_LIST_REPORTS
						sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_NAME_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_JOURNEY_FLAGS & "," & L_SHIFT_FLAGS & "," & L_LEVEL_FLAGS & "," & L_EMPLOYEE_STATUS_FLAGS & "," & L_SOCIAL_SECURITY_NUMBER_FLAGS & "," & L_EMPLOYEE_BIRTH_FLAGS & "," & L_EMPLOYEE_RFC_FLAGS & "," & L_EMPLOYEE_CURP_FLAGS & "," & L_EMPLOYEE_GENDER_FLAGS & "," & L_EMPLOYEE_ACTIVE_FLAGS & "," & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_NAME_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_JOURNEY_FLAGS & "," & L_SHIFT_FLAGS & "," & L_LEVEL_FLAGS & "," & L_EMPLOYEE_STATUS_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_EMPLOYEE_EMAIL_FLAGS & "," & L_SOCIAL_SECURITY_NUMBER_FLAGS & "," & L_EMPLOYEE_BIRTH_FLAGS & "," & L_EMPLOYEE_RFC_FLAGS & "," & L_EMPLOYEE_CURP_FLAGS & "," & L_EMPLOYEE_GENDER_FLAGS & "," & L_EMPLOYEE_ACTIVE_FLAGS & "," & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS
								lErrorNumber = DisplayReportTemplateForm(sFlags, sErrorDescription)
							Case 2
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 3
								If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 0 Then Response.Write "<DONT_EXPORT>"
									lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 0 Then Response.Write "</DONT_EXPORT>"
								If lErrorNumber = 0 Then
									lErrorNumber = DisplayAreasList(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case EMPLOYEES_LIST_REPORTS
						sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_NAME_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_JOURNEY_FLAGS & "," & L_SHIFT_FLAGS & "," & L_LEVEL_FLAGS & "," & L_EMPLOYEE_STATUS_FLAGS & "," & L_EMPLOYEE_START_DATE_FLAGS & "," & L_SOCIAL_SECURITY_NUMBER_FLAGS & "," & L_EMPLOYEE_BIRTH_FLAGS & "," & L_EMPLOYEE_RFC_FLAGS & "," & L_EMPLOYEE_CURP_FLAGS & "," & L_EMPLOYEE_GENDER_FLAGS & "," & L_EMPLOYEE_ACTIVE_FLAGS & "," & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_NAME_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_JOURNEY_FLAGS & "," & L_SHIFT_FLAGS & "," & L_LEVEL_FLAGS & "," & L_EMPLOYEE_STATUS_FLAGS & "," & L_EMPLOYEE_START_DATE_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_EMPLOYEE_EMAIL_FLAGS & "," & L_SOCIAL_SECURITY_NUMBER_FLAGS & "," & L_EMPLOYEE_BIRTH_FLAGS & "," & L_EMPLOYEE_RFC_FLAGS & "," & L_EMPLOYEE_CURP_FLAGS & "," & L_EMPLOYEE_GENDER_FLAGS & "," & L_EMPLOYEE_ACTIVE_FLAGS & "," & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS
								lErrorNumber = DisplayReportTemplateForm(sFlags, sErrorDescription)
							Case 2
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 3
								If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 0 Then Response.Write "<DONT_EXPORT>"
									lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 0 Then Response.Write "</DONT_EXPORT>"
								If lErrorNumber = 0 Then
									lErrorNumber = DisplayEmployeesList(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case JOBS_LIST_REPORTS, SPECIAL_JOBS_LIST_REPORTS, JOBS_LIST_BY_MODIFY_DATE
						If Len(oRequest("Flags").Item) = 0 Then
							sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_NAME_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_JOURNEY_FLAGS & "," & L_SHIFT_FLAGS & "," & L_LEVEL_FLAGS & "," & L_EMPLOYEE_STATUS_FLAGS & "," & L_SOCIAL_SECURITY_NUMBER_FLAGS & "," & L_EMPLOYEE_BIRTH_FLAGS & "," & L_EMPLOYEE_RFC_FLAGS & "," & L_EMPLOYEE_CURP_FLAGS & "," & L_EMPLOYEE_GENDER_FLAGS & "," & L_EMPLOYEE_ACTIVE_FLAGS & "," & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS & "," & L_JOB_STATUS_FLAGS
						Else
							sFlags = oRequest("Flags").Item
						End If
						If (aReportsComponent(N_ID_REPORTS) = SPECIAL_JOBS_LIST_REPORTS Or (aReportsComponent(N_ID_REPORTS) = JOBS_LIST_BY_MODIFY_DATE)) And (aReportsComponent(N_STEP_REPORTS) = 1) Then aReportsComponent(N_STEP_REPORTS) = 2
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_NAME_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_JOURNEY_FLAGS & "," & L_SHIFT_FLAGS & "," & L_LEVEL_FLAGS & "," & L_EMPLOYEE_STATUS_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_EMPLOYEE_EMAIL_FLAGS & "," & L_SOCIAL_SECURITY_NUMBER_FLAGS & "," & L_EMPLOYEE_BIRTH_FLAGS & "," & L_EMPLOYEE_RFC_FLAGS & "," & L_EMPLOYEE_CURP_FLAGS & "," & L_EMPLOYEE_GENDER_FLAGS & "," & L_EMPLOYEE_ACTIVE_FLAGS & "," & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS & "," & L_JOB_STATUS_FLAGS
								lErrorNumber = DisplayReportTemplateForm(sFlags, sErrorDescription)
							Case 2
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""MovedEmployees"" ID=""MovedEmployeesHdn"" VALUE=""" & oRequest("MovedEmployees").Item & """ />"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobsOwners"" ID=""JobsOwnersHdn"" VALUE=""" & oRequest("JobsOwners").Item & """ />"
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobStatusID"" ID=""JobStatusIDHdn"" VALUE=""" & oRequest("JobStatusID").Item & """ />"
								If Len(oRequest("Flags").Item) > 0 Then
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Flags"" ID=""FlagsHdn"" VALUE=""" & oRequest("Flags").Item & """ />"
								End If
							Case 3
								If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 0 Then Response.Write "<DONT_EXPORT>"
									lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 0 Then Response.Write "</DONT_EXPORT>"
								If lErrorNumber = 0 Then
									lErrorNumber = DisplayJobsList(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1001_REPORTS
						sFlags = L_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_CHECK_CONCEPTS_EMPLOYEES_FLAGS & "," & L_BANK_FLAGS & "," & L_REPORT_TITLE_FLAGS
						asTitles = Split("Título 1;;;Título 2;;;Título 3", LIST_SEPARATOR)
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1001(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1002_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<B>Nota: </B>Esta nómina será comparada contra la nómina anterior.<BR />"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1002(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1003_REPORTS, ISSSTE_1470_REPORTS, ISSSTE_4701_REPORTS
						If aReportsComponent(N_ID_REPORTS) = ISSSTE_4701_REPORTS Then
							sFlags = L_CANCELL_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_CHECK_CONCEPTS_ALL_FLAGS & "," & L_BANK_FLAGS & "," & L_TOTAL_PAYMENT_FLAGS & "," & L_HAS_ALIMONY_FLAGS & "," & L_HAS_CREDITS_FLAGS & "," & L_CHECK_NUMBER_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Else
							sFlags = L_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_CHECK_CONCEPTS_ALL_FLAGS & "," & L_BANK_FLAGS & "," & L_TOTAL_PAYMENT_FLAGS & "," & L_HAS_ALIMONY_FLAGS & "," & L_HAS_CREDITS_FLAGS & "," & L_CHECK_NUMBER_FLAGS & "," & L_ZIP_WARNING_FLAGS
						End If
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<BR />"
								Response.Write "<IMG SRC=""Images/Crcl.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fecha de emisión de la nómina:&nbsp;</FONT>"
								Response.Write DisplayDateCombos(CInt(oRequest("PayrollIssueYear").Item), CInt(oRequest("PayrollIssueMonth").Item), CInt(oRequest("PayrollIssueDay").Item), "PayrollIssueYear", "PayrollIssueMonth", "PayrollIssueDay", N_FORM_START_YEAR, Year(Date())+1, True, True)
'								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
'									If Len(oRequest("PayrollIssueYear").Item) = 0 Then Response.Write "document.ReportFrm.PayrollIssueYear.value = " & Year(Date()) & ";" & vbNewLine
'								Response.Write "//--></SCRIPT>" & vbNewLine
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									Select Case aReportsComponent(N_ID_REPORTS)
										Case ISSSTE_4701_REPORTS
											lErrorNumber = BuildReports1003Cancel(oRequest, oADODBConnection, False, sErrorDescription)
										Case Else
											lErrorNumber = BuildReports1003(oRequest, oADODBConnection, False, sErrorDescription)
									End Select
                                End If
						End Select
					Case ISSSTE_1004_REPORTS
						sFlags = L_MONTHS_FLAGS & "," & L_YEARS_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_CHECK_CONCEPTS_EMPLOYEES_FLAGS & "," & L_BANK_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1004(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1005_REPORTS
						sFlags = L_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_CHECK_CONCEPTS_FLAGS & "," & L_BANK_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1005(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1006_REPORTS, ISSSTE_4702_REPORTS
						If aReportsComponent(N_ID_REPORTS) = ISSSTE_4702_REPORTS Then
							sFlags = L_CANCELL_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_ZONE_FOR_PAYMENT_CENTER_FLAGS & "," & L_CHECK_CONCEPTS_ALL_FLAGS & "," & L_BANK_FLAGS & "," & L_ZIP_WARNING_FLAGS & "," & L_REPORT_TITLE_FLAGS
						Else
							sFlags = L_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_ZONE_FOR_PAYMENT_CENTER_FLAGS & "," & L_CHECK_CONCEPTS_ALL_FLAGS & "," & L_BANK_FLAGS & "," & L_ZIP_WARNING_FLAGS & "," & L_REPORT_TITLE_FLAGS
						End If
						'sFlags = L_CLOSED_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_ZONE_FOR_PAYMENT_CENTER_FLAGS & "," & L_CHECK_CONCEPTS_ALL_FLAGS & "," & L_BANK_FLAGS & "," & L_ZIP_WARNING_FLAGS & "," & L_REPORT_TITLE_FLAGS
						asTitles = Split("Título A;;;Título B;;;Título C", LIST_SEPARATOR)
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									Select Case aReportsComponent(N_ID_REPORTS)
										Case ISSSTE_4702_REPORTS
											lErrorNumber = BuildReport1006Cancel(oRequest, oADODBConnection, True, sErrorDescription)
										Case Else
											lErrorNumber = BuildReport1006(oRequest, oADODBConnection, True, sErrorDescription)
									End Select
								End If
						End Select
					Case ISSSTE_1007_REPORTS
						sFlags = L_DONT_CLOSE_DIV_FLAGS & "," & L_PAYROLL_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATE_TYPE_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)

								Response.Write "<IMG SRC=""Images/Crcl8.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Tipo de quincena:<BR /></FONT>"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""-2,-1,0"""
									If InStr(1, ",124,155,", "," & oRequest("ConceptID").Item & ",", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
								Response.Write " />&nbsp;Normal<BR />"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""124""" '69 y 89
									If StrComp(oRequest("ConceptID").Item, "124", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
								Response.Write " />&nbsp;Pensión alimenticia<BR />"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""155""" '155
									If StrComp(oRequest("ConceptID").Item, "155", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
								Response.Write " />&nbsp;Acreedores<BR />"
								Response.Write "</DIV>"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1007(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1008_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_MONTHS_FLAGS & "," & L_YEARS_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1008(oRequest, oADODBConnection, False, False, sErrorDescription)
								End If
						End Select	
					Case ISSSTE_1009_REPORTS
						sFlags = L_MONTHS_FLAGS & "," & L_YEARS_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_CHECK_CONCEPTS_EMPLOYEES_FLAGS & "," & L_BANK_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1009(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1010_REPORTS
						sFlags = L_PAYROLL_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_CHECK_CONCEPTS_EMPLOYEES_FLAGS & "," & L_BANK_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1010(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1011_REPORTS
						sFlags = L_PAYROLL_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_CHECK_CONCEPTS_EMPLOYEES_FLAGS & "," & L_BANK_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1011(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1012_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_ORDINARY_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1012(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1013_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_ORDINARY_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1013(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1014_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_ORDINARY_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1014(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1015_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_ORDINARY_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1015(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1016_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_ORDINARY_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1016(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1017_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_ORDINARY_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1017(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1018_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_ORDINARY_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1018(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1019_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_ORDINARY_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1019(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1020_REPORTS
						sFlags = L_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_BANK_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1020(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1021_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_ORDINARY_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1021(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1022_REPORTS
						sFlags = L_TOTAL_PAYMENT_FLAGS & "," & L_PAYROLL_FLAGS & "," & L_COMPANY_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1022(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1023_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_MONTHS_FLAGS & "," & L_YEARS_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1023(oRequest, oADODBConnection, False, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1024_REPORTS
						sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_PAYROLL_FLAGS & "," & L_COMPANY_FLAGS & "," & L_CONCEPT_1_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1024(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1025_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_MONTHS_FLAGS & "," & L_YEARS_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1025(oRequest, oADODBConnection, False, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1026_REPORTS
						sFlags = L_ABSENCE_APPLIED_DATE_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_ABSENCE_ID_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_ABSENCE_ACTIVE_FLAGS & "," & L_USER_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<IMG SRC=""Images/Crcl10.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fechas de vigencia de las incidencias:<BR /></FONT>"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
								Response.Write DisplayDateCombos(CInt(oRequest("OcurredDateYear").Item), CInt(oRequest("OcurredDateMonth").Item), CInt(oRequest("OcurredDateDay").Item), "OcurredDateYear", "OcurredDateMonth", "OcurredDateDay", N_FORM_START_YEAR, Year(Date()), True, True)
								Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
								Response.Write DisplayDateCombos(CInt(oRequest("EndDateYear").Item), CInt(oRequest("EndDateMonth").Item), CInt(oRequest("EndDateDay").Item), "EndDateYear", "EndDateMonth", "EndDateDay", N_FORM_START_YEAR, Year(Date()), True, True)
								Response.Write "<BR /><BR />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Agrupado por centro de pago:&nbsp;</FONT><INPUT TYPE=""CHECKBOX"" NAME=""ForPayments"" ID=""ForPaymentsTxt"" SIZE=""2"" MAXLENGTH=""2"" VALUE=""1"" />"
								Response.Write "</DIV>"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									If Len(oRequest("ForPayments").Item) > 0 Then
										lErrorNumber = BuildReport1026(oRequest, oADODBConnection, sErrorDescription)
									Else
										lErrorNumber = BuildReport1126(oRequest, oADODBConnection, sErrorDescription)
									End If
								End If
						End Select
					Case ISSSTE_1027_REPORTS
						sFlags = L_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_CHECK_CONCEPTS_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1027(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1028_REPORTS
						lErrorNumber = BuildReport1028(oRequest, oADODBConnection, sErrorDescription)
					Case ISSSTE_1029_REPORTS
						sFlags = L_ABSENCE_ID_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_ABSENCE_ACTIVE_FLAGS & "," & L_ABSENCE_APPLIED_DATE_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1029(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1030_REPORTS
						sFlags = L_ABSENCE_APPLIED_DATE_FLAGS & "," & L_ABSENCE_ID_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_ABSENCE_ACTIVE_FLAGS & "," & L_USER_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<IMG SRC=""Images/Crcl9.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fechas de vigencia de las incidencias:<BR /></FONT>"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
								Response.Write DisplayDateCombos(CInt(oRequest("OcurredDateYear").Item), CInt(oRequest("OcurredDateMonth").Item), CInt(oRequest("OcurredDateDay").Item), "OcurredDateYear", "OcurredDateMonth", "OcurredDateDay", N_FORM_START_YEAR, Year(Date()), True, True)
								Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
								Response.Write DisplayDateCombos(CInt(oRequest("EndDateYear").Item), CInt(oRequest("EndDateMonth").Item), CInt(oRequest("EndDateDay").Item), "EndDateYear", "EndDateMonth", "EndDateDay", N_FORM_START_YEAR, Year(Date()), True, True)
								Response.Write "<BR /><BR />"
								Response.Write "</DIV>"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1030(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1031_REPORTS
						sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_MOVEMENT_TYPE
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = BuildReport1031(oRequest, oADODBConnection, sErrorDescription)
						End Select
					Case ISSSTE_1032_REPORTS
						lErrorNumber = BuildReport1032(oRequest, oADODBConnection, sErrorDescription)
					Case ISSSTE_1033_REPORTS
						lErrorNumber = BuildReport1033(oRequest, oADODBConnection, sErrorDescription)
					Case ISSSTE_1034_REPORTS
						lErrorNumber = BuildReport1034(oRequest, oADODBConnection, sErrorDescription)
					Case ISSSTE_1035_REPORTS
						lErrorNumber = BuildReport1035(oRequest, oADODBConnection, sErrorDescription)
					Case ISSSTE_1100_REPORTS
						sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_NAME_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_JOURNEY_FLAGS & "," & L_SHIFT_FLAGS & "," & L_LEVEL_FLAGS & "," & L_EMPLOYEE_STATUS_FLAGS & "," & L_SOCIAL_SECURITY_NUMBER_FLAGS & "," & L_EMPLOYEE_BIRTH_FLAGS & "," & L_EMPLOYEE_RFC_FLAGS & "," & L_EMPLOYEE_CURP_FLAGS & "," & L_EMPLOYEE_GENDER_FLAGS & "," & L_EMPLOYEE_ACTIVE_FLAGS & "," & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1100(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1101_REPORTS, ISSSTE_1113_REPORTS
						sFlags = L_YEARS_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1101(oRequest, oADODBConnection, oRequest("YearID").Item, -1, 1, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1102_REPORTS
						sFlags = L_DATE_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & L_USER_FLAGS & "," & L_ADJUSTMENT_APPLIED_DATE_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1102(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1103_REPORTS
						sFlags = L_USER_FLAGS & "," & L_PAYROLL_FLAGS & "," & L_DATE_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_JOB_NUMBER_FLAGS & "," & L_EMPLOYEE_REASON_ID_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1103(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1104_REPORTS
						sFlags = L_PAYROLL_FLAGS & "," & L_USER_FLAGS & "," & L_DATE_FLAGS & "," & L_EMPLOYEE_REASON_ID_FLAGS 
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1104(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1105_REPORTS
						sFlags = L_USER_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_PAYROLL_FLAGS & "," & L_DATE_FLAGS & "," & L_EMPLOYEE_REASON_ID_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1105(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1106_REPORTS
						sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_PAYROLL_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1106(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1107_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAYROLL_FLAGS & "," & L_CONCEPT_1_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1200(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1108_REPORTS
						sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_CONCEPT_ID_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_USER_FLAGS & "," & L_CONCEPTS_APPLIED_DATE_FLAGS & "," & L_ZIP_WARNING_FLAGS & "," & L_CONCEPT_ACTIVE_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<IMG SRC=""Images/Crcl10.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fechas de vigencia de los conceptos:<BR /></FONT>"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
								Response.Write DisplayDateCombos(CInt(oRequest("StartDateYear").Item), CInt(oRequest("StartDateMonth").Item), CInt(oRequest("StartDateDay").Item), "StartDateYear", "StartDateMonth", "StartDateDay", N_FORM_START_YEAR, Year(Date()), True, True)
								Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
								Response.Write DisplayDateCombos(CInt(oRequest("EndDateYear").Item), CInt(oRequest("EndDateMonth").Item), CInt(oRequest("EndDateDay").Item), "EndDateYear", "EndDateMonth", "EndDateDay", N_FORM_START_YEAR, Year(Date()), True, True)
								Response.Write "<BR /><BR />"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1108(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1109_REPORTS
						If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
							sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_DATE_FLAGS & "," & L_EMPLOYEE_REASON_ID_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Else
							sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_NUMBER7_FLAGS & "," & L_DATE_FLAGS & "," & L_EMPLOYEE_REASON_ID_FLAGS & "," & L_ZIP_WARNING_FLAGS
						End If
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1109(oRequest, oADODBConnection, False, Null, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1110_REPORTS
						If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
							sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_DATE_FLAGS
						Else
							sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_NUMBER7_FLAGS & "," & L_DATE_FLAGS
						End If
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1110(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1111_REPORTS
						sFlags = L_COMPANY_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_JOURNEY_FLAGS & "," & L_SHIFT_FLAGS & "," & L_LEVEL_FLAGS & "," & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS & "," & L_JOB_STATUS_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1111(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1112_REPORTS
						sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_EMPLOYEE_STATUS_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1112(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1115_REPORTS
						If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
							sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_DATE_FLAGS
						Else
							sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_NUMBER7_FLAGS & "," & L_DATE_FLAGS
						End If
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1115(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1116_REPORTS, ISSSTE_1204_REPORTS, ISSSTE_1702_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_EMPLOYEE_NUMBER1_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<IMG SRC=""Images/Crcl2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Antigüedad hasta el día: </FONT>"
								Response.Write DisplayDateCombosUsingSerial(CLng(Left(GetSerialNumberForDate(""), Len("00000000"))), "Employee", N_FORM_START_YEAR, Year(Date()), True, False)
								Response.Write "<BR /><BR />"
							Case 2
								'lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1116(oRequest, oADODBConnection, False, Null, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1117_REPORTS, ISSSTE_1205_REPORTS, ISSSTE_1703_REPORTS
						sFlags = L_DONT_CLOSE_FILTER_DIV_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<BR /><IMG SRC=""Images/Crcl9.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Antigüedad hasta el día: </FONT>"
								Response.Write DisplayDateCombosUsingSerial(CLng(Left(GetSerialNumberForDate(""), Len("00000000"))), "Employee", N_FORM_START_YEAR, Year(Date()), True, False)
								Response.Write "<BR /><BR />"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Antigüedad hasta el día: </B>"
									If (Len(oRequest("EmployeeYear").Item) > 0) And (Len(oRequest("EmployeeMonth").Item) > 0) And (Len(oRequest("EmployeeDay").Item) > 0) Then
										Response.Write DisplayDateFromSerialNumber(CLng(oRequest("EmployeeYear").Item & oRequest("EmployeeMonth").Item & oRequest("EmployeeDay").Item), -1, -1, -1)
									Else
										Response.Write DisplayDateFromSerialNumber("", -1, -1, -1)
									End If
								Response.Write "</FONT><BR /></DIV><BR /><BR />"
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1117(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1118_REPORTS, ISSSTE_1206_REPORTS, ISSSTE_1704_REPORTS
						sFlags = L_DONT_CLOSE_FILTER_DIV_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)

								Response.Write "<BR /><IMG SRC=""Images/Crcl.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fecha para el cálculo:&nbsp;</FONT>"
								Response.Write DisplayDateCombos(Year(Date()), Month(Date()), Day(Date()), "PayrollYear", "PayrollMonth", "PayrollDay", N_FORM_START_YEAR, Year(Date()), True, False)
								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
									If Len(oRequest("PayrollYear").Item) = 0 Then Response.Write "document.ReportFrm.PayrollYear.value = " & Year(Date()) & ";" & vbNewLine
									If Len(oRequest("PayrollMonth").Item) = 0 Then Response.Write "document.ReportFrm.PayrollMonth.value = '10';" & vbNewLine
									If Len(oRequest("PayrollDay").Item) = 0 Then Response.Write "document.ReportFrm.PayrollDay.value = '01';" & vbNewLine
								Response.Write "//--></SCRIPT>" & vbNewLine
							Case 2
								sFlags = L_DONT_CLOSE_FILTER_DIV_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								Response.Write "<B>Fecha de cálculo: " & DisplayDateFromSerialNumber(oRequest("PayrollYear").Item & oRequest("PayrollMonth").Item & oRequest("PayrollDay").Item, -1, -1, -1) & "</B>"
								Response.Write "</DIV><BR />"
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1118(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1119_REPORTS
						sFlags = L_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<IMG SRC=""Images/Crcl10.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								'Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fechas de vigencia de los conceptos:<BR /></FONT>"
								'Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
								'Response.Write DisplayDateCombos(CInt(oRequest("StartDateYear").Item), CInt(oRequest("StartDateMonth").Item), CInt(oRequest("StartDateDay").Item), "StartDateYear", "StartDateMonth", "StartDateDay", N_FORM_START_YEAR, Year(Date()), True, True)
								'Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
								'Response.Write DisplayDateCombos(CInt(oRequest("EndDateYear").Item), CInt(oRequest("EndDateMonth").Item), CInt(oRequest("EndDateDay").Item), "EndDateYear", "EndDateMonth", "EndDateDay", N_FORM_START_YEAR, Year(Date()), True, True)
								'Response.Write "<BR /><BR />"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1119(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
                    Case ISSSTE_1120_REPORTS
                        sFlags = L_YEARS_FLAGS & "," & L_CONCEPT_2_FLAGS & "," & L_COMPANY_FLAGS & "," & L_CONCENTRATE_CONCEPTS_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
                            Case 1
                                lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
                            Case 2
                                lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1120(oRequest, oADODBConnection, sErrorDescription)
								End If
                        End Select
                    Case ISSSTE_1151_REPORTS, ISSSTE_1152_REPORTS, ISSSTE_1153_REPORTS, ISSSTE_1154_REPORTS, ISSSTE_1155_REPORTS, ISSSTE_1157_REPORTS
						If aReportsComponent(N_ID_REPORTS) = ISSSTE_1151_REPORTS Then
							sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_CONCEPT_ID_FLAGS & "," & L_EMPLOYEE_NAME_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_EMPLOYEE_STATUS_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Else
							sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_NAME_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_EMPLOYEE_STATUS_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS & "," & L_ZIP_WARNING_FLAGS
						End If
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next

								Response.Write "<TABLE BGCOLOR=""#" & S_WIDGET_FRAME_FOR_GUI & """ WIDTH=""98%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""1""><TR><TD>"
									Response.Write "<TABLE BGCOLOR=""#FFFFFF"" WIDTH=""100%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""3""><TR><TD><FONT FACE=""Arial"" SIZE=""2"">"
										Response.Write "<FONT FACE=""Arial"" SIZE=""2"" COLOR=""#" & S_INSTRUCTIONS_FOR_GUI & """><B>SELECCIONE EL AÑO A PROCESAR</B><BR /><BR /></FONT>"
										Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B>Año a procesar:</B>&nbsp;<SELECT NAME=""YearID"" ID=""YearIDCmb"" SIZE=""1"" CLASS=""Lists"">"
											For iIndex = 2009 To Year(Date())
												Response.Write "<OPTION VALUE=""" & iIndex & """"
													If Month(Date()) > 3 Then
														If iIndex = Year(Date()) Then Response.Write " SELECTED=""1"""
													Else
														If iIndex = Year(Date()) - 1 Then Response.Write " SELECTED=""1"""
													End If
												Response.Write ">" & iIndex & "</OPTION>"
											Next
										Response.Write "</SELECT><BR /><BR />"
										If aReportsComponent(N_ID_REPORTS) = ISSSTE_1155_REPORTS Then
											Response.Write "<FONT FACE=""Arial"" SIZE=""2"" COLOR=""#" & S_INSTRUCTIONS_FOR_GUI & """><B>SELECCIONE LOS LÍMITES PARA LA EXCLUSIÓN DE AJUSTES</B><BR /><BR /></FONT>"
											Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B>Importe máximo a descontar:</B>&nbsp;<INPUT TYPE=""TEXT"" NAME=""MaxDiscount"" ID=""MaxDiscountTxt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & oRequest("MaxDiscount").Item & """ /><BR />"
											'Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B>Excluir ingresos menores a:</B>&nbsp;<INPUT TYPE=""TEXT"" NAME=""Limit1"" ID=""Limit1Txt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & oRequest("Limit1").Item & """ /><BR />"
											Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<B>Excluir ingresos mayores a:</B>&nbsp;<INPUT TYPE=""TEXT"" NAME=""Limit2"" ID=""Limit2Txt"" SIZE=""10"" MAXLENGTH=""10"" VALUE="""
												If Len(oRequest("Limit2").Item) > 0 Then
													Response.Write oRequest("Limit2").Item
												Else
													Response.Write "400,000"
												End If
											Response.Write """ /><BR /><BR />"
										End If
									Response.Write "</FONT></TD></TR></TABLE>"
								Response.Write "</TD></TR></TABLE><BR />"

								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								Response.Write "<B>Año procesado: " & oRequest("YearID").Item & "</B><BR />"
								If aReportsComponent(N_ID_REPORTS) = ISSSTE_1155_REPORTS Then
									Response.Write "<B>Se descontó como máximo " & FormatNumber(oRequest("MaxDiscount").Item, 2, True, False, True) & "</B><BR />"
									'Response.Write "<B>Se excluyeron los ingresos menores a " & FormatNumber(oRequest("Limit1").Item, 2, True, False, True) & "</B><BR />"
									Response.Write "<B>Se excluyeron los ingresos mayores a " & FormatNumber(oRequest("Limit2").Item, 2, True, False, True) & "</B><BR />"
								End If
								Response.Write "<BR />"
								If lErrorNumber = 0 Then
									Select Case aReportsComponent(N_ID_REPORTS)
										Case ISSSTE_1151_REPORTS
											lErrorNumber = BuildReport1151(oRequest, oADODBConnection, sErrorDescription)
										Case ISSSTE_1152_REPORTS
											lErrorNumber = BuildReport1152(oRequest, oADODBConnection, sErrorDescription)
										Case ISSSTE_1153_REPORTS
											lErrorNumber = BuildReport1153(oRequest, oADODBConnection, False, sErrorDescription)
										Case ISSSTE_1154_REPORTS
											lErrorNumber = BuildReport1154(oRequest, oADODBConnection, sErrorDescription)
										Case ISSSTE_1155_REPORTS
											lErrorNumber = BuildReport1153(oRequest, oADODBConnection, True, sErrorDescription)
											If lErrorNumber = 0 Then
												Call DisplayErrorMessage("Ajustes aplicados", "Los ajustes a los impuestos fueron registrados con éxito.")
											End If
										Case ISSSTE_1157_REPORTS
											lErrorNumber = BuildReport1157(oRequest, oADODBConnection, sErrorDescription)
									End Select
								End If
						End Select
					Case ISSSTE_1200_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAYROLL_FLAGS & "," & L_CONCEPT_1_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1200(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1201_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAYROLL_FLAGS & "," & L_CONCEPT_1_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Agrupado por centro de pago:&nbsp;</FONT><INPUT TYPE=""CHECKBOX"" NAME=""ForPayments"" ID=""ForPaymentsTxt"" SIZE=""2"" MAXLENGTH=""2"" VALUE=""1"" />"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									If Len(oRequest("ForPayments").Item) > 0 Then
										lErrorNumber = BuildReport1201b(oRequest, oADODBConnection, sErrorDescription)
									Else
										lErrorNumber = BuildReport1201(oRequest, oADODBConnection, sErrorDescription)
									End If
								End If
						End Select
					Case ISSSTE_1202_REPORTS
						sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_CREDITS_TYPES_ID_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_USER_FLAGS & "," & L_CREDITS_APPLIED_DATE_FLAGS & "," & L_CREDITS_ACTIVE_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)

								Response.Write "<IMG SRC=""Images/Crcl9.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fechas de vigencia de los créditos:<BR /></FONT>"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
								Response.Write DisplayDateCombos(CInt(oRequest("CreditStartYear").Item), CInt(oRequest("CreditStartMonth").Item), CInt(oRequest("CreditStartDay").Item), "CreditStartYear", "CreditStartMonth", "CreditStartDay", N_FORM_START_YEAR, Year(Date()), True, True)
								Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
								Response.Write DisplayDateCombos(CInt(oRequest("CreditEndYear").Item), CInt(oRequest("CreditEndMonth").Item), CInt(oRequest("CreditEndDay").Item), "CreditEndYear", "CreditEndMonth", "CreditEndDay", N_FORM_START_YEAR, Year(Date()), True, True)
								Response.Write "<BR /><BR />"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1202(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1203_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_SERVICES_SHEET_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)

								Response.Write "<IMG SRC=""Images/Crcl.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fechas de vigencia de las hojas únicas de servicio:<BR /></FONT>"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
								Response.Write DisplayDateCombos(CInt(oRequest("DocumentStartYear").Item), CInt(oRequest("DocumentStartMonth").Item), CInt(oRequest("DocumentStartDay").Item), "DocumentStartYear", "DocumentStartMonth", "DocumentStartDay", N_FORM_START_YEAR, Year(Date()), True, True)
								Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
								Response.Write DisplayDateCombos(CInt(oRequest("DocumentEndYear").Item), CInt(oRequest("DocumentEndMonth").Item), CInt(oRequest("DocumentEndDay").Item), "DocumentEndYear", "DocumentEndMonth", "DocumentEndDay", N_FORM_START_YEAR, Year(Date()), True, True)
								Response.Write "<BR /><BR />"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									Select Case oRequest("ServicesSheetTypeID").Item
										Case "A"
											lErrorNumber = BuildReport1203(oRequest, oADODBConnection, sErrorDescription)
										Case "B"
											lErrorNumber = BuildReport1203b(oRequest, oADODBConnection, sErrorDescription)
										Case "C"
											lErrorNumber = BuildReport1203c(oRequest, oADODBConnection, sErrorDescription)
									End Select
								End If
						End Select
					Case ISSSTE_1207_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_EMPLOYEE_NUMBER1_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<IMG SRC=""Images/Crcl2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">No. de oficio:<BR /></FONT>"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""DocumentNumber"" ID=""DocumentNumberTxt"" VALUE=""" & oRequest("DocumentNumber").Item & """ SIZE=""10"" MAXLENGTH=""10"" CLASS=""TextFields"" />"
								Response.Write "<BR /><BR />"

								Response.Write "<IMG SRC=""Images/Crcl3.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Observaciones:<BR /></FONT>"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<TEXTAREA NAME=""Comments"" ID=""CommentsTxtArea"" ROWS=""6"" COLS=""50"" MAXLENGTH=""255"" CLASS=""TextFields"">" & oRequest("Comments").Item & "</TEXTAREA>"
								Response.Write "<BR /><BR />"

								Response.Write "<IMG SRC=""Images/Crcl4.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Propósito:<BR /></FONT>"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<TEXTAREA NAME=""Purpose"" ID=""PurposeTxtArea"" ROWS=""6"" COLS=""50"" MAXLENGTH=""255"" CLASS=""TextFields"">" & oRequest("Purpose").Item & "</TEXTAREA>"
								Response.Write "<BR /><BR />"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1207(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1208_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_EMPLOYEE_NUMBER1_FLAGS & "," & L_CONCEPT_1_FLAGS & "," & L_DATE_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1208(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1209_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_USER_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)

								Response.Write "<IMG SRC=""Images/Crcl5.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fechas de registro de las revisiones:<BR /></FONT>"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
								Response.Write DisplayDateCombos(CInt(oRequest("AddDateStartYear").Item), CInt(oRequest("AddDateStartMonth").Item), CInt(oRequest("AddDateStartDay").Item), "AddDateStartYear", "AddDateStartMonth", "AddDateStartDay", N_FORM_START_YEAR, Year(Date()), True, True)
								Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
								Response.Write DisplayDateCombos(CInt(oRequest("AddDateEndYear").Item), CInt(oRequest("AddDateEndMonth").Item), CInt(oRequest("AddDateEndDay").Item), "AddDateEndYear", "AddDateEndMonth", "AddDateEndDay", N_FORM_START_YEAR, Year(Date()), True, True)
								Response.Write "<BR /><BR />"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1209(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1210_REPORTS
						sFlags = L_DONT_CLOSE_DIV_FLAGS & "," & L_ABSENCE_APPLIED_DATE_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_ABSENCE_ACTIVE_FLAGS & "," & L_USER_FLAGS & "," & L_EXTRAHOURS_AND_SUNDAYS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<IMG SRC=""Images/Crcl10.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fechas de vigencia de las incidencias:<BR /></FONT>"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
								Response.Write DisplayDateCombos(CInt(oRequest("OcurredDateYear").Item), CInt(oRequest("OcurredDateMonth").Item), CInt(oRequest("OcurredDateDay").Item), "OcurredDateYear", "OcurredDateMonth", "OcurredDateDay", N_FORM_START_YEAR, Year(Date()), True, True)
								Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
								Response.Write DisplayDateCombos(CInt(oRequest("EndDateYear").Item), CInt(oRequest("EndDateMonth").Item), CInt(oRequest("EndDateDay").Item), "EndDateYear", "EndDateMonth", "EndDateDay", N_FORM_START_YEAR, Year(Date()), True, True)
								Response.Write "<BR /><BR />"
								Response.Write "</DIV>"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1210(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1211_REPORTS
						sFlags = L_YEARS_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TotalByArea"" ID=""TotalByAreaHdn"" VALUE=""1"" />"
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1211(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1221_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_CREDITS_TYPES_ID_FLAGS & "," & S_CREDITS_UPLOADED_FILE_NAME
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1221(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1222_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & S_CREDITS_UPLOADED_FILE_NAME
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1222(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1223_REPORTS
						sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_AREA_FLAGS & "," & L_CONCEPTS_APPLIED_DATE_FLAGS 
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1223(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1224_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS  & "," & L_CONCEPTS_APPLIED_DATE_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								
								Response.Write "<IMG SRC=""Images/Crcl5.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fechas de vigencia de las pensiones:<BR /></FONT>"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
								Response.Write DisplayDateCombos(CInt(oRequest("BeneficiaryStartYear").Item), CInt(oRequest("BeneficiaryStartMonth").Item), CInt(oRequest("BeneficiaryStartDay").Item), "BeneficiaryStartYear", "BeneficiaryStartMonth", "BeneficiaryStartDay", N_FORM_START_YEAR, Year(Date()), True, True)
								Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
								Response.Write DisplayDateCombos(CInt(oRequest("BeneficiaryEndYear").Item), CInt(oRequest("BeneficiaryEndMonth").Item), CInt(oRequest("BeneficiaryEndDay").Item), "BeneficiaryEndYear", "BeneficiaryEndMonth", "BeneficiaryEndDay", N_FORM_START_YEAR, Year(Date()), True, True)
								Response.Write "<BR /><BR />"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1224(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1225_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1225(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1311_REPORTS
						sFlags = L_PAYROLL_FLAGS& "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_NAME_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_JOURNEY_FLAGS & "," & L_SHIFT_FLAGS & "," & L_LEVEL_FLAGS & "," & L_EMPLOYEE_STATUS_FLAGS & "," & L_EMPLOYEE_START_DATE_FLAGS & "," & L_SOCIAL_SECURITY_NUMBER_FLAGS & "," & L_EMPLOYEE_BIRTH_FLAGS & "," & L_EMPLOYEE_RFC_FLAGS & "," & L_EMPLOYEE_CURP_FLAGS & "," & L_EMPLOYEE_GENDER_FLAGS & "," & L_EMPLOYEE_ACTIVE_FLAGS & "," & L_JOB_NUMBER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_JOB_TYPE_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1311(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1334_REPORTS
						sFlags = L_PAYROLL_FLAGS & "," & L_GENERATING_AREAS_FLAGS & "," & L_MEDICAL_AREAS_TYPES_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1334(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1335_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_EMPLOYEE_TYPE1_FLAGS & "," & L_CONCEPTS_VALUES_STATUS_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1335(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1336_REPORTS
						sFlags = L_GENERATING_AREAS_FLAGS & "," & L_AREA_CODE_FLAGS & "," & L_AREA_TYPE_FLAGS & "," & L_CONFINE_TYPE_FLAGS & "," & L_CENTER_TYPE_FLAGS & "," & L_CENTER_SUBTYPE_FLAGS & "," & L_ATTENTION_LEVEL_FLAGS & "," & L_ECONOMIC_ZONE_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1336(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1337_REPORTS
						sFlags = L_GENERATING_AREAS_FLAGS & "," & L_AREA_CODE_FLAGS & "," & L_AREA_TYPE_FLAGS & "," & L_CONFINE_TYPE_FLAGS & "," & L_CENTER_TYPE_FLAGS & "," & L_ATTENTION_LEVEL_FLAGS & "," & L_ECONOMIC_ZONE_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1337(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1338_REPORTS
						sFlags = L_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_CHECK_CONCEPTS_ALL_FLAGS & "," & L_BANK_FLAGS & "," & L_TOTAL_PAYMENT_FLAGS & "," & L_HAS_ALIMONY_FLAGS & "," & L_HAS_CREDITS_FLAGS & "," & L_CHECK_NUMBER_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<BR />"
								Response.Write "<IMG SRC=""Images/Crcl.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fecha de emisión de la nómina:&nbsp;</FONT>"
								Response.Write DisplayDateCombos(CInt(oRequest("PayrollIssueYear").Item), CInt(oRequest("PayrollIssueMonth").Item), CInt(oRequest("PayrollIssueDay").Item), "PayrollIssueYear", "PayrollIssueMonth", "PayrollIssueDay", Year(Date()), Year(Date())+1, True, True)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReports1003(oRequest, oADODBConnection, True, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1339_REPORTS, ISSSTE_1340_REPORTS
						sFlags = L_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_CHECK_CONCEPTS_ALL_FLAGS & "," & L_BANK_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								If aReportsComponent(N_ID_REPORTS) = ISSSTE_1339_REPORTS Then
									Response.Write "<IMG SRC=""Images/Crcl.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
									Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Formato:<BR /></FONT>"
									Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""RADIO"" NAME=""LongReport"" ID=""LongReportRd"" VALUE="""" CHECKED=""1"" /> Formato corto<BR />"
									Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""RADIO"" NAME=""LongReport"" ID=""LongReportRd"" VALUE=""1"" /> Formato largo<BR />"
								Else
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Cancelled"" ID=""CancelledHdn"" VALUE=""1"" />"
								End If
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1339(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1354_REPORTS
						sFlags = L_NO_DIV_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<IMG SRC=""Images/Crcl1.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Tipo de registro:<BR /></FONT>"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""Kardex5TypeID"" ID=""Kardex5TypeIDLst"" SIZE=""1"" CLASS=""Lists"">"
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Kardex5Types", "Kardex5TypeID", "Kardex5TypeName", "", "Kardex5TypeName", oRequest("Kardex5TypeID").Item, "", sErrorDescription)
								Response.Write "</SELECT><BR /><BR />"
							Case 2
								lErrorNumber = BuildReport1354(oRequest, oADODBConnection, False, sErrorDescription)
						End Select
					Case ISSSTE_1356_REPORTS
						sFlags = L_NO_DIV_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<IMG SRC=""Images/Crcl1.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Registro de:<BR /></FONT>"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""KardexChangeTypeID"" ID=""KardexChangeTypeIDLst"" SIZE=""1"" CLASS=""Lists"">"
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "KardexChangeTypes", "KardexChangeTypeID", "KardexChangeTypeName", "", "KardexChangeTypeName", oRequest("KardexChangeTypeID").Item, "", sErrorDescription)
								Response.Write "</SELECT><BR /><BR />"
							Case 2
								lErrorNumber = BuildReport1356(oRequest, oADODBConnection, False, sErrorDescription)
						End Select
					Case ISSSTE_1364_REPORTS
						sFlags = L_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_COURSE_NAME_FLAGS & "," & L_COURSE_DIPLOMA_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_NAME_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS & "," & L_COURSE_NAME_FLAGS & "," & L_COURSE_DIPLOMA_FLAGS & "," & L_COURSE_LOCATION_FLAGS & "," & L_COURSE_DURATION_FLAGS & "," & L_COURSE_PARTICIPANTS_FLAGS & "," & L_COURSE_DATES_FLAGS & "," & L_COURSE_GRADE_FLAGS
								lErrorNumber = DisplayReportTemplateForm(sFlags, sErrorDescription)
							Case 2
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 3
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1364(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1371_REPORTS
						sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_PAYROLL_FLAGS & "," & L_ZONE_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1371(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1372_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAYROLL_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1372(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1373_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAYROLL_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1373(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1374_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAYROLL_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1374(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1400_REPORTS
						'sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_ZONE_FLAGS & "," & L_BANK_FLAGS & "," & L_CHECK_CONCEPTS_FLAGS
						sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & ","& L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_ZONE_FLAGS & "," & L_CHECK_CONCEPTS_EMPLOYEES_FLAGS & "," & L_BANK_FLAGS &"," & L_PAYMENT_TYPE_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
							sFlags = L_NO_DIV_FLAGS & "," & L_PAYROLL_FLAGS
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PayrollID"" ID=""PayrollIDHdn"" VALUE=""" & oRequest("PayrollID").Item & """ />"
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 3
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									Call DisplayURLParametersAsHiddenValues(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(oRequest, "ReportID"), "ReportType"), "ReportStep"), "FromSectionID"), "PayrollCode"), "PayrollCLC"),"Memorandum"),"PayrollDescription"),"FileCLC"),"CancelYearCLC"),"PayrollTypeIDCLC"),"YearCLC"),"MonthCLC"),"QuarterCLC"))
									Response.Write "<TABLE>"
                                        Response.Write"<TR>"
									    If (Len(oRequest("PayrollCode").Item) <> 0) Then
										    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Bimestre: </B></TD><TD colspan=""5""><INPUT TYPE=""Text"" NAME=""PayrollCode"" ID=""PayrollCodeTxt"" SIZE=""13"" MAXLENGTH=""13"" VALUE=""" & oRequest("PayrollCode").Item & """ CLASS=""TextFields"" />"
									    Else
										    asPeriods = Split("0,1,1,2,2,3,3,4,4,5,5,6,6",",")
										    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Bimestre: </B></TD><TD colspan=""5""><INPUT TYPE=""Text"" NAME=""PayrollCode"" ID=""PayrollCodeTxt"" SIZE=""13"" MAXLENGTH=""13"" VALUE=""" & Mid(oRequest("PayrollID").Item,1,4) & "0" & asPeriods(Mid(oRequest("PayrollID").Item,5,2)) & """ CLASS=""TextFields"" />"
									    End If
									    Response.Write "</TD>"
                                    Response.Write "<TR>"
                                        Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>CLC: </B></TD>"
                                        Response.Write "<TD colspan=""5""><INPUT TYPE=""Text"" NAME=""PayrollCLC"" ID=""PayrollCLCTxt"" SIZE=""30"" MAXLENGTH=""30"" VALUE=""" & oRequest("PayrollCLC").Item & """ CLASS=""TextFields"" /> </TD>"
                                    Response.Write "</TR>"
                                    Response.Write "<TR>"
                                        Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Descripción de la nómina: </B></TD>"
                                        Response.Write "<TD colspan=""5""><INPUT TYPE=""Text"" NAME=""PayrollDescription"" ID=""PayrollDescriptionTxt"" SIZE=""30"" MAXLENGTH=""150"" VALUE="""& oRequest("PayrollDescription").Item &""" CLASS=""TextFields"" /></TD>"
                                    Response.Write "</TR>"
                                    Response.Write "<TR>"
                                        Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Año de la nómina: </B></TD>"
                                        Response.Write "<TD><INPUT TYPE=""Text"" NAME=""YearCLC"" ID=""YearCLCTxt"" SIZE=""4"" MAXLENGTH=""4"" VALUE="""& Mid(oRequest("PayrollID").Item,1,4) &""" CLASS=""TextFields"" /></TD>"
                                    'Response.Write "</TR>"
                                    'Response.Write "<TR>"
                                        Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Mes: </B></TD>"
                                        Response.Write "<TD><INPUT TYPE=""Text"" NAME=""MonthCLC"" ID=""MonthCLCTxt"" SIZE=""2"" MAXLENGTH=""2"" VALUE="""& Mid(oRequest("PayrollID").Item,5,2) &""" CLASS=""TextFields"" /></TD>"
                                    'Response.Write "</TR>"
                                    'Response.Write "<TR>"
                                        Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Quincena: </B></TD>"
                                        Response.Write "<TD><INPUT TYPE=""Text"" NAME=""QuarterCLC"" ID=""QuarterCLCTxt"" SIZE=""2"" MAXLENGTH=""2"" VALUE="""& oRequest("QuarterCLC").Item &""" CLASS=""TextFields"" /></TD>"
                                    Response.Write "</TR>"
                                    Response.Write "<TR>"
                                        Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Memorandum: </B></TD>"
                                        Response.Write "<TD colspan=""5""><INPUT TYPE=""Text"" NAME=""Memorandum"" ID=""MemorandumTxt"" SIZE=""30"" MAXLENGTH=""150"" VALUE="""& oRequest("Memorandum").Item &""" CLASS=""TextFields"" /></TD>"
                                    Response.Write "</TR>"
                                    Response.Write "<TR>"
                                        Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Archivo: </B></TD>"
                                        Response.Write "<TD colspan=""5""><INPUT TYPE=""Text"" NAME=""FileCLC"" ID=""FileCLCTxt"" SIZE=""30"" MAXLENGTH=""150"" VALUE="""& oRequest("FileCLC").Item &""" CLASS=""TextFields"" /></TD>"
                                    Response.Write "</TR>"
									Response.Write "<TR>"
                                        Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Año Cancelación: </B></TD>"
                                        Response.Write "<TD><INPUT TYPE=""Text"" NAME=""CancelYearCLC"" ID=""CancelYearCLCTxt"" SIZE=""4"" MAXLENGTH=""4"" VALUE="""
                                        If len(oRequest("CancelYearCLC").Item)>0 Then
                                            Response.Write ""& oRequest("CancelYearCLC").Item &""
                                        Else 
                                            Response.Write ""& Mid(oRequest("PayrollID").Item,1,4) &""  
                                        End If
                                         Response.Write""" CLASS=""TextFields"" /></TD>"
                                    Response.Write "</TR>"
                                    Response.Write "<TR>"
                                        Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Tipo de nómina:</B></TD>"
                                        Response.Write "<TD colspan=""5""><SELECT NAME=""PayrollTypeIDCLC"" ID=""PayrollTypeIDCLCTxt"" SIZE=""1"" CLASS=""Lists"">"
                                        Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					                    Response.Write GenerateListOptionsFromQuery(oADODBConnection, "PayrollTypes", "PayrollTypeID", "PayrollTypeName", "", "PayrollTypeName", "", "", sErrorDescription)
				                    Response.Write "</SELECT></TD>"
                                    Response.Write "</TR>"                                 
                                    Response.Write "</TABLE>"
									
								End If
							Case 4
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									Call DisplayURLParametersAsHiddenValues(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(oRequest, "ReportID"), "ReportType"), "ReportStep"), "FromSectionID"), "PayrollCode"), "PayrollCLC"),"Memorandum"),"PayrollDescription"),"FileCLC"),"CancelYearCLC"),"PayrollTypeIDCLC"),"YearCLC"),"MonthCLC"),"QuarterCLC"))
									'Call DisplayURLParametersAsHiddenValues(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(oRequest, "ReportID"), "ReportType"), "ReportStep"), "FromSectionID"))
									Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
										Response.Write "<TR>"
											Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Bimestre:&nbsp;</B></TD>"
											Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & oRequest("PayrollCode").Item & "</TD>"
										Response.Write "</TR>"
										Response.Write "<TR>"
											Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>CLC:&nbsp;</B></TD>"
											Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & oRequest("PayrollCLC").Item & "</TD>"
										Response.Write "</TR>"
									Response.Write "</TABLE>"
									'lErrorNumber = BuildReport14001(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
						If (lErrorNumber = 0) And (aReportsComponent(N_STEP_REPORTS) > 2) Then
							Call GetNameFromTable(oADODBConnection, "Payrolls", oRequest("PayrollID").Item, "", "", sNames, sErrorDescription)
							Response.Write "<BR /><B>CLCs para la quincena " & CleanStringForHTML(sNames) & "</B>:"
							lErrorNumber = BuildReport1400b1(oRequest, oADODBConnection, CLng(oRequest("PayrollID").Item), sErrorDescription)
                        End If
					Case ISSSTE_1401_REPORTS
						sFlags = L_PAYROLL_FLAGS & "," & L_CONCEPT_ID_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_CHECK_CONCEPTS_ALL_FLAGS & "," & L_BANK_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1401(oRequest, oADODBConnection, "", sErrorDescription)
								End If
						End Select
					Case ISSSTE_1402_REPORTS
						sFlags = L_PAYROLL_FLAGS & "," & L_CONCEPT_ID_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_CHECK_CONCEPTS_EMPLOYEES_FLAGS & "," & L_BANK_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TotalByArea"" ID=""TotalByAreaHdn"" VALUE=""1"" />"
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1401(oRequest, oADODBConnection, "", sErrorDescription)
								End If
						End Select
					Case ISSSTE_1403_REPORTS
						sFlags = L_PAYROLL_FLAGS &","& L_MONTHS_FLAGS &"," &L_QUARTER_FLAGS &","& L_YEARS_FLAGS &","&  L_STATES_FLAGS &"," & L_PAYMENT_TYPE_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TotalByArea"" ID=""TotalByAreaHdn"" VALUE=""1"" />"
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1403(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1404_REPORTS
						sFlags = L_PAYROLL_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_CHECK_CONCEPTS_EMPLOYEES_FLAGS & "," & L_BANK_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TotalByArea"" ID=""TotalByAreaHdn"" VALUE=""1"" />"
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1404(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1411_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1411(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1412_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1401(oRequest, oADODBConnection, "56,76,77", sErrorDescription)
								End If
						End Select
					Case ISSSTE_1413_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1413(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1414_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1414(oRequest, oADODBConnection, True, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1415_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1415(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1416_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1414(oRequest, oADODBConnection, False, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1417_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1417(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1420_REPORTS, ISSSTE_2420_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_ONE_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""18,34"" "
									If StrComp(oRequest("ConceptID").Item, "96", vbBinaryCompare) <> 0 Then Response.Write " CHECKED=""1"""
								Response.Write " /> Guardias y suplencias<BR />"
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""96"" "
									If StrComp(oRequest("ConceptID").Item, "96", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
								Response.Write " /> Guardias PROVAC<BR />"
							Case 2
								sFlags = L_NO_DIV_FLAGS & "," & L_ONE_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYROLL_FLAGS & "," & L_CONCEPT_FLAGS
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1420(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1421_REPORTS, ISSSTE_2421_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""18"" "
									If (Len(oRequest("ConceptID").Item) = 0) Or (StrComp(oRequest("ConceptID").Item, "18", vbBinaryCompare) = 0) Then Response.Write " CHECKED=""1"""
								Response.Write " /> Guardias<BR />"
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""34"" "
									If StrComp(oRequest("ConceptID").Item, "34", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
								Response.Write " /> Suplencias<BR />"
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""96"" "
									If StrComp(oRequest("ConceptID").Item, "96", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
								Response.Write " /> Guardias PROVAC<BR />"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1421(oRequest, oADODBConnection, False, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1422_REPORTS, ISSSTE_2422_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""18"" "
									If (Len(oRequest("ConceptID").Item) = 0) Or (StrComp(oRequest("ConceptID").Item, "18", vbBinaryCompare) = 0) Then Response.Write " CHECKED=""1"""
								Response.Write " /> Guardias<BR />"
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""34"" "
									If StrComp(oRequest("ConceptID").Item, "34", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
								Response.Write " /> Suplencias<BR />"
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""96"" "
									If StrComp(oRequest("ConceptID").Item, "96", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
								Response.Write " /> Guardias PROVAC<BR />"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1422(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1423_REPORTS, ISSSTE_2423_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""18"" "
									If (Len(oRequest("ConceptID").Item) = 0) Or (StrComp(oRequest("ConceptID").Item, "18", vbBinaryCompare) = 0) Then Response.Write " CHECKED=""1"""
								Response.Write " /> Guardias<BR />"
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""34"" "
									If StrComp(oRequest("ConceptID").Item, "34", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
								Response.Write " /> Suplencias<BR />"
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""96"" "
									If StrComp(oRequest("ConceptID").Item, "96", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
								Response.Write " /> Guardias PROVAC<BR />"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1421(oRequest, oADODBConnection, True, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1424_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""External"" ID=""ExternalHdn"" VALUE=""" & oRequest("External").Item & """ />"
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1424(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1425_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""External"" ID=""ExternalHdn"" VALUE=""" & oRequest("External").Item & """ />"
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1425(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1426_REPORTS, ISSSTE_2426_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYROLL_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""18"" "
									If (Len(oRequest("ConceptID").Item) = 0) Or (StrComp(oRequest("ConceptID").Item, "18", vbBinaryCompare) = 0) Then Response.Write " CHECKED=""1"""
								Response.Write " /> Guardias<BR />"
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""34"" "
									If StrComp(oRequest("ConceptID").Item, "34", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
								Response.Write " /> Suplencias<BR />"
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""96"" "
									If StrComp(oRequest("ConceptID").Item, "96", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
								Response.Write " /> Guardias PROVAC<BR />"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1426(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1427_REPORTS, ISSSTE_2427_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""18"" "
									If (Len(oRequest("ConceptID").Item) = 0) Or (StrComp(oRequest("ConceptID").Item, "18", vbBinaryCompare) = 0) Then Response.Write " CHECKED=""1"""
								Response.Write " /> Guardias<BR />"
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""34"" "
									If StrComp(oRequest("ConceptID").Item, "34", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
								Response.Write " /> Suplencias<BR />"
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""96"" "
									If StrComp(oRequest("ConceptID").Item, "96", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
								Response.Write " /> Guardias PROVAC<BR />"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1427(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1428_REPORTS, ISSSTE_2428_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""18"" "
									If (Len(oRequest("ConceptID").Item) = 0) Or (StrComp(oRequest("ConceptID").Item, "18", vbBinaryCompare) = 0) Then Response.Write " CHECKED=""1"""
								Response.Write " /> Guardias<BR />"
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""34"" "
									If StrComp(oRequest("ConceptID").Item, "34", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
								Response.Write " /> Suplencias<BR />"
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""96"" "
									If StrComp(oRequest("ConceptID").Item, "96", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
								Response.Write " /> Guardias PROVAC<BR />"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1428(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1429_REPORTS, ISSSTE_2429_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""18"" "
									If (Len(oRequest("ConceptID").Item) = 0) Or (StrComp(oRequest("ConceptID").Item, "18", vbBinaryCompare) = 0) Then Response.Write " CHECKED=""1"""
								Response.Write " /> Guardias<BR />"
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""34"" "
									If StrComp(oRequest("ConceptID").Item, "34", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
								Response.Write " /> Suplencias<BR />"
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""96"" "
									If StrComp(oRequest("ConceptID").Item, "96", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
								Response.Write " /> Guardias PROVAC<BR />"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1429(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1430_REPORTS, ISSSTE_2430_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYROLL_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""18"" "
									If (Len(oRequest("ConceptID").Item) = 0) Or (StrComp(oRequest("ConceptID").Item, "18", vbBinaryCompare) = 0) Then Response.Write " CHECKED=""1"""
								Response.Write " /> Guardias<BR />"
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""34"" "
									If StrComp(oRequest("ConceptID").Item, "34", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
								Response.Write " /> Suplencias<BR />"
								Response.Write "<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""96"" "
									If StrComp(oRequest("ConceptID").Item, "96", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
								Response.Write " /> Guardias PROVAC<BR />"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1430(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1431_REPORTS, ISSSTE_2431_REPORTS
						sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_BANK_FLAGS & "," & L_CONCEPTS_APPLIED_DATE_FLAGS & "," & L_USER_FLAGS  & "," & L_BANK_ACCOUNTS_ACTIVE_FLAGS  & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<BR />"
								Response.Write "<IMG SRC=""Images/Crcl6.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"

								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fechas de vigencia de las cuentas bancarias:<BR /></FONT>"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
								Response.Write DisplayDateCombos(CInt(oRequest("AccountStartDateYear").Item), CInt(oRequest("AccountStartDateMonth").Item), CInt(oRequest("AccountStartDateDay").Item), "AccountStartDateYear", "AccountStartDateMonth", "AccountStartDateDay", N_FORM_START_YEAR, Year(Date()), True, True)
								Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
								Response.Write DisplayDateCombos(CInt(oRequest("AccountEndDateYear").Item), CInt(oRequest("AccountEndDateMonth").Item), CInt(oRequest("AccountEndDateDay").Item), "AccountEndDateYear", "AccountEndDateMonth", "AccountEndDateDay", N_FORM_START_YEAR, Year(Date()), True, True)
								Response.Write "<BR /><BR />"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1431(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1432_REPORTS, ISSSTE_2432_REPORTS
						sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_ZONE_FLAGS_FOR_EMPLOYEES & "," & L_BANK_FLAGS '& "," & L_CONCEPTS_APPLIED_DATE_FLAGS  ' L_ZONE_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<BR />"
								Response.Write "<IMG SRC=""Images/Crcl6.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Fechas de registro de las cuentas bancarias:<BR /></FONT>"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
								Response.Write DisplayDateCombos(CInt(oRequest("RegistrationStartDateYear").Item), CInt(oRequest("RegistrationStartDateMonth").Item), CInt(oRequest("RegistrationStartDateDay").Item), "RegistrationStartDateYear", "RegistrationStartDateMonth", "RegistrationStartDateDay", N_FORM_START_YEAR, Year(Date()), True, True)
								Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
								Response.Write DisplayDateCombos(CInt(oRequest("RegistrationEndDateYear").Item), CInt(oRequest("RegistrationEndDateMonth").Item), CInt(oRequest("RegistrationEndDateDay").Item), "RegistrationEndDateYear", "RegistrationEndDateMonth", "RegistrationEndDateDay", N_FORM_START_YEAR, Year(Date()), True, True)
								Response.Write "<BR /><BR />"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1432(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1433_REPORTS
						sFlags = L_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_BANK_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1433(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1434_REPORTS
						sFlags = L_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_BANK_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1434(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1435_REPORTS
						sFlags = L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_CONCEPT_ID_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1435(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1471_REPORTS
						sFlags = L_PAYROLL_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_ONE_BANK_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1471(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1114_REPORTS, ISSSTE_1472_REPORTS
						sFlags = L_DONT_CLOSE_FILTER_DIV_FLAGS & "," & L_DONT_CLOSE_DIV_FLAGS & "," & L_PAYROLL_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_ZONE_FLAGS & "," & L_BANK_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<IMG SRC=""Images/Crcl7.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Tipo de pago:&nbsp;</FONT>"
								Response.Write "<SELECT NAME=""PaymentType"" ID=""PaymentTypeCmb"" SIZE=""1"" CLASS=""Lists"">"
									Response.Write "<OPTION VALUE="""">Todos</OPTION>"
									Response.Write "<OPTION VALUE=""0"">Cheques</OPTION>"
									Response.Write "<OPTION VALUE=""1"">Depósitos</OPTION>"
								Response.Write "</SELECT><BR /><BR /></DIV>"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">"
									Response.Write "<B>Tipo de pago:</B><BR />"
									Select Case oRequest("PaymentType").Item
										Case ""
											Response.Write "-Todos"
										Case "0"
											Response.Write "-Cheques"
										Case "1"
											Response.Write "-Depósitos"
									End Select
								Response.Write "</FONT></DIV>"
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1472(oRequest, oADODBConnection, -1, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1473_REPORTS
						sFlags = L_PAYROLL_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_ZONE_FLAGS & "," & L_BANK_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1472(oRequest, oADODBConnection, 1, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1474_REPORTS
						sFlags = L_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_ZONE_FLAGS & "," & L_ISSSTE_ONE_BANK_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<BR />"
								Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
									Response.Write "<TR>"
										Response.Write "<TD><IMG SRC=""Images/Crcl8.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" /></TD>"
										Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de generación (Serfín):&nbsp;</FONT></TD>"
										Response.Write "<TD>" & DisplayDateCombos(CInt(oRequest("FileYear").Item), CInt(oRequest("FileMonth").Item), CInt(oRequest("FileDay").Item), "FileYear", "FileMonth", "FileDay", N_FORM_START_YEAR, Year(Date()), True, False) & "</TD>"
									Response.Write "</TR>"
									Response.Write "<TR>"
										Response.Write "<TD><IMG SRC=""Images/Crcl9.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" /></TD>"
										Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de depósito (Banamex,SPEI , Banorte y Serfín):&nbsp;</FONT></TD>"
										Response.Write "<TD>" & DisplayDateCombos(CInt(oRequest("PayrollDepositYear").Item), CInt(oRequest("PayrollDepositMonth").Item), CInt(oRequest("PayrollDepositDay").Item), "PayrollDepositYear", "PayrollDepositMonth", "PayrollDepositDay", N_FORM_START_YEAR, Year(Date()), True, False) & "</TD>"
									Response.Write "</TR>"

									Response.Write "<TR>"
										Response.Write "<TD><IMG SRC=""Images/Crcl10.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" /></TD>"
										Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. secuencial:&nbsp;</FONT></TD>"
										Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""FileNumber"" ID=""FileNumberHdn"" SIZE=""4"" MAXLENGTH=""4"" VALUE="""
											If Len(oRequest("FileNumber").Item) = 0 Then
												Response.Write "0001"
											Else
												Response.Write oRequest("FileNumber").Item
											End If
										Response.Write """ CLASS=""TextFields"" /></TD>"
									Response.Write "</TR>"

									Response.Write "<TR>"
										Response.Write "<TD><IMG SRC=""Images/Crcl11.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" /></TD>"
										Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. del cliente:&nbsp;</FONT></TD>"
										Response.Write "<TD>"
                                            Response.Write "<SELECT NAME=""Field02"" ID=""Field02Hdn"" SIZE=""1"" CLASS=""Lists"" >"
                                                Response.Write "<option value=""000059667242"">000059667242</option>"
                                                Response.Write "<option value=""000058364781"">000058364781</option>"  
                                            Response.Write "<\SELECT>"
                                        Response.Write "<\TD>"
                                        'Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""Field02"" ID=""Field02Hdn"" SIZE=""12"" MAXLENGTH=""12"" VALUE="""
										'	If Len(oRequest("Field02").Item) = 0 Then												
                                        '          Response.Write"000059667242"                                                
									    '		Else
										'  		  Response.Write oRequest("Field02").Item
										'	End If
										'Response.Write """ CLASS=""TextFields"" /></TD>"
									Response.Write "</TR>"

									Response.Write "<TR>"
										Response.Write "<TD><IMG SRC=""Images/Crcl12.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" /></TD>"
										Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Sucursal:&nbsp;</FONT></TD>"
										Response.Write "<TD>"
                                            Response.Write "<SELECT NAME=""Field03"" ID=""Field03Hdn"" SIZE=""1"" CLASS=""Lists"" >"
                                                Response.Write "<option value=""0100"">0100</option>"
                                                Response.Write "<option value=""0224"">0224</option>"  
                                            Response.Write "<\SELECT>"
                                        Response.Write "<\TD>"
                                        'Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""Field03"" ID=""Field03Hdn"" SIZE=""12"" MAXLENGTH=""12"" VALUE="""
										'	If Len(oRequest("Field03").Item) = 0 Then
										'		Response.Write "0100"
										'	Else
										'		Response.Write oRequest("Field03").Item
										'	End If
										'Response.Write """ CLASS=""TextFields"" /></TD>"
									Response.Write "</TR>"

									Response.Write "<TR>"
										Response.Write "<TD><IMG SRC=""Images/Crcl13.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" /></TD>"
										Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cuenta de cargo:&nbsp;</FONT></TD>"
										Response.Write "<TD>"
                                            Response.Write "<SELECT NAME=""Field04"" ID=""Field04Hdn"" SIZE=""1"" CLASS=""Lists"" >"
                                                Response.Write "<option value=""00000000000000978469"">00000000000000978469</option>"
                                                Response.Write "<option value=""00000000000007708668"">00000000000007708668</option>"
                                                Response.Write "<option value=""00000000000004160460"">00000000000004160460</option>"
                                            Response.Write "<\SELECT>"
                                        Response.Write "<\TD>"
                                        'Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""Field04"" ID=""Field04Hdn"" SIZE=""20"" MAXLENGTH=""20"" VALUE="""
										'	If Len(oRequest("Field04").Item) = 0 Then
										'		Response.Write "00000000000007708668"
										'	Else
										'		Response.Write oRequest("Field04").Item
										'	End If
										'Response.Write """ CLASS=""TextFields"" /></TD>"
									Response.Write "</TR>"

									Response.Write "<TR>"
										Response.Write "<TD><IMG SRC=""Images/Crcl14.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" /></TD>"
										Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Versión del layout (Banamex):&nbsp;</FONT></TD>"
										Response.Write "<TD><SELECT NAME=""LayoutType"" ID=""LayoutTypeCmb"" SIZE=""1"">"
											Response.Write "<OPTION VALUE=""B"">B</OPTION>"
											Response.Write "<OPTION VALUE=""C"">C</OPTION>"
										Response.Write "</SELECT></TD>"
									Response.Write "</TR>"

									Response.Write "<TR>"
										Response.Write "<TD><IMG SRC=""Images/Crcl15.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" /></TD>"
										Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. de contrato (BBVA Bancomer, Pagel):&nbsp;</FONT></TD>"
										Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""Field05"" ID=""Field05Hdn"" SIZE=""12"" MAXLENGTH=""12"" VALUE="""
											If Len(oRequest("Field05").Item) = 0 Then
												Response.Write "000000000000"
											Else
												Response.Write oRequest("Field05").Item
											End If
										Response.Write """ CLASS=""TextFields"" /></TD>"
									Response.Write "</TR>"

									Response.Write "<TR>"
										Response.Write "<TD><IMG SRC=""Images/Crcl16.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" /></TD>"
										Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Emisora (Banorte):&nbsp;</FONT></TD>"
										Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""Field01"" ID=""Field01Hdn"" SIZE=""5"" MAXLENGTH=""5"" VALUE="""
											If Len(oRequest("Field01").Item) = 0 Then
												Response.Write "44951"
											Else
												Response.Write oRequest("Field01").Item
											End If
										Response.Write """ CLASS=""TextFields"" /></TD>"
									Response.Write "</TR>"
								
									'Response.Write "<TR>"
									'	Response.Write "<TD><IMG SRC=""Images/Crcl17.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" /></TD>"
									'	Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. contrato B.E. (SPEI):&nbsp;</FONT></TD>"
									'	Response.Write "<TD>"
                                    '        Response.Write "<SELECT NAME=""Field06"" ID=""Field06Hdn"" SIZE=""1"" CLASS=""Lists"" >"
                                    '            Response.Write "<option value=""000059667242"">000059667242</option>"
                                    '            Response.Write "<option value=""000058364781"">000058364781</option>"  
                                    '        Response.Write "<\SELECT>"
                                    '    Response.Write "<\TD>"
                                    '    'Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""Field06"" ID=""Field06Hdn"" SIZE=""15"" MAXLENGTH=""5"" VALUE="""
									'	'	If Len(oRequest("Field06").Item) = 0 Then
									'	'		Response.Write "000058364781"
									'	'	Else
									'	'		Response.Write oRequest("Field06").Item
									'	'	End If
									'	'Response.Write """ CLASS=""TextFields"" /></TD>"
									'Response.Write "</TR>"

									'Response.Write "<TR>"
									'	Response.Write "<TD><IMG SRC=""Images/Crcl18.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" /></TD>"
									'	Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Sucursal de cargo (SPEI):&nbsp;</FONT></TD>"
									'	Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""SpeiBranch"" ID=""SpeiBranchHdn"" SIZE=""15"" MAXLENGTH=""5"" VALUE="""
									'		If Len(oRequest("SpeiBranch").Item) = 0 Then
									'			Response.Write "0100"
									'		Else
									'			Response.Write oRequest("SpeiBranch").Item
									'		End If
									'	Response.Write """ CLASS=""TextFields"" /></TD>"
									'Response.Write "</TR>"

									'Response.Write "<TR>"
									'	Response.Write "<TD><IMG SRC=""Images/Crcl19.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" /></TD>"
									'	Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de aplicación (SPEI):&nbsp;</FONT></TD>"
									'	Response.Write "<TD>" & DisplayDateCombos(CInt(oRequest("ApplicationYear").Item), CInt(oRequest("ApplicationMonth").Item), CInt(oRequest("ApplicationDay").Item), "PayrollIssueYear", "PayrollIssueMonth", "PayrollIssueDay", N_FORM_START_YEAR, Year(Date()), True, False) & "</TD>"
									'Response.Write "</TR>"

									'Response.Write "<TR>"
									'	Response.Write "<TD><IMG SRC=""Images/Crcl20.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" /></TD>"
									'	Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Referencia bajo valor (SPEI):&nbsp;</FONT></TD>"
									'	Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""SpeiRef"" ID=""SpeiRefHdn"" SIZE=""15"" MAXLENGTH=""5"" VALUE="""
									'		If Len(oRequest("SpeiRef").Item) = 0 Then
									'			Response.Write "0260612"
									'		Else
									'			Response.Write oRequest("SpeiRef").Item
									'		End If
									'	Response.Write """ CLASS=""TextFields"" /></TD>"
									'Response.Write "</TR>"
								Response.Write "</TABLE>"
								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
									If Len(oRequest("PayrollIssueYear").Item) = 0 Then Response.Write "document.ReportFrm.PayrollIssueYear.value = " & Year(Date()) & ";" & vbNewLine
									If Len(oRequest("FileYear").Item) = 0 Then Response.Write "document.ReportFrm.FileYear.value = " & Year(Date()) & ";" & vbNewLine
									Response.Write "AddItemToList('SERFÍN. HONORARIOS', '-17', null, document.ReportFrm.BankID);" & vbNewLine
									Response.Write "AddItemToList('BBVA BANCOMER. Baja California', '-1', null, document.ReportFrm.BankID);" & vbNewLine
									Response.Write "AddItemToList('BBVA BANCOMER. Pagel', '-2', null, document.ReportFrm.BankID);" & vbNewLine
								Response.Write "//--></SCRIPT>" & vbNewLine
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1474(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1475_REPORTS
						sFlags = L_PAYROLL_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_ZONE_FLAGS & "," & L_ISSSTE_ONE_BANK_FLAGS & "," & L_ONLY_CHECK_CONCEPTS_EMPLOYEES_FLAGS & "," & L_CHECK_NUMBER_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1475(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1476_REPORTS
						sFlags = L_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_ZONE_FLAGS & "," & L_BANK_FLAGS & "," & L_CHECK_CONCEPTS_ALL_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1476(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1477_REPORTS
						sFlags = L_MONTHS_FLAGS & "," & L_YEARS_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_ZONE_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1477(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1478_REPORTS
						sFlags = L_MONTHS_FLAGS & "," & L_YEARS_FLAGS & "," & L_COMPANY_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_ZONE_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)

								Response.Write "<IMG SRC=""Images/Crcl7.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Impuesto:&nbsp;</FONT>"
								Response.Write "<INPUT TYPE=""TEXT"" NAME=""TaxAmount"" ID=""TaxAmountTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE=""2.5"" CLASS=""TextFields"" />&nbsp;%<BR />"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1478(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1490_REPORTS, ISSSTE_4703_REPORTS
						If aReportsComponent(N_ID_REPORTS) = ISSSTE_4703_REPORTS Then
							sFlags = L_DONT_CLOSE_DIV_FLAGS & "," & L_CANCELL_PAYROLL_FLAGS & "," & L_COMPANY_FLAGS & "," & L_REPORT_TYPE_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_ZONE_FOR_PAYMENT_CENTER_FLAGS & "," & L_BANK_FLAGS & "," & L_STATE_TYPE_FLAGS
                        Else
							sFlags = L_DONT_CLOSE_DIV_FLAGS & "," & L_PAYROLL_FLAGS & "," & L_COMPANY_FLAGS & "," & L_REPORT_TYPE_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_ZONE_FOR_PAYMENT_CENTER_FLAGS & "," & L_BANK_FLAGS & "," & L_STATE_TYPE_FLAGS
						End If
                        Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next

								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)

								Response.Write "<IMG SRC=""Images/Crcl10.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Tipo de quincena:<BR /></FONT>"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""-2,-1,0"""
									If InStr(1, ",124,155,", "," & oRequest("ConceptID").Item & ",", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
								Response.Write " />&nbsp;Normal<BR />"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""124""" '69 y 89
									If StrComp(oRequest("ConceptID").Item, "124", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
								Response.Write " />&nbsp;Pensión alimenticia<BR />"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""RADIO"" NAME=""ConceptID"" ID=""ConceptIDRd"" VALUE=""155""" '155
									If StrComp(oRequest("ConceptID").Item, "155", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
								Response.Write " />&nbsp;Acreedores<BR /><BR />"
								Response.Write "</DIV>"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									Select Case aReportsComponent(N_ID_REPORTS)
										Case ISSSTE_4703_REPORTS
											lErrorNumber = BuildReport1490Cancel(oRequest, oADODBConnection, False, sErrorDescription)
										Case Else
											lErrorNumber = BuildReport1490(oRequest, oADODBConnection, False, sErrorDescription)
									End Select
                                End If
						End Select
					Case ISSSTE_1491_REPORTS
						sFlags = L_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_THIRD_CONCEPTS2_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1491(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1492_REPORTS
						sFlags = L_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_THIRD_CONCEPTS_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1492(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1493_REPORTS
						sFlags = L_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_AREA_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_STATES_FLAGS & "," & L_THIRD_CONCEPTS_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1401(oRequest, oADODBConnection, oRequest("ConceptID").Item, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1494_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_CLOSED_PAYROLL_FLAGS & "," & L_MEMORY_CONCEPT_ID_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1494(oRequest, oADODBConnection, oRequest("ConceptID").Item, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1495_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_AUDIT_TYPE_ID_FLAGS & "," & L_AUDIT_OPERATION_TYPE_ID_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1495(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1499_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1499(oRequest, oADODBConnection, oRequest("EmployeeNumber").Item, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1503_REPORTS
						sFlags = L_BUDGET_ORIGINAL_POSITION_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_COMPANY_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_LEVEL_FLAGS & "," & L_ECONOMIC_ZONE_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								For Each oItem In oRequest
									If InStr(1, oItem, "P_", vbBinaryCompare) = 1 Then
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""" & oItem & """ ID=""" & oItem & "Hdn"" VALUE=""" & oRequest(oItem).Item & """ />"
									End If
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CompanyID"" ID=""CompanyIDHdn"" VALUE=""" & Replace(oRequest("CompanyID").Item, " ", "") & """ />"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDHdn"" VALUE=""" & Replace(oRequest("EmployeeTypeID").Item, " ", "") & """ />"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ClassificationID"" ID=""ClassificationIDHdn"" VALUE=""" & Replace(oRequest("ClassificationID").Item, " ", "") & """ />"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""GroupGradeLevelID"" ID=""GroupGradeLevelIDHdn"" VALUE=""" & Replace(oRequest("GroupGradeLevelID").Item, " ", "") & """ />"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IntegrationID"" ID=""IntegrationIDHdn"" VALUE=""" & Replace(oRequest("IntegrationID").Item, " ", "") & """ />"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""LevelID"" ID=""LevelIDHdn"" VALUE=""" & Replace(oRequest("LevelID").Item, " ", "") & """ />"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EconomicZoneID"" ID=""EconomicZoneIDHdn"" VALUE=""" & Replace(oRequest("EconomicZoneID").Item, " ", "") & """ />"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetPositionID"" ID=""BudgetPositionIDHdn"" VALUE=""" & Replace(oRequest("BudgetPositionID").Item, " ", "") & """ />"
								Response.Write "<TABLE BORER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
									Response.Write "<TD VALIGN=""TOP"">"
										lErrorNumber = DisplayReport1503Positions(oRequest, oADODBConnection, False, sErrorDescription)
										If lErrorNumber <> 0 Then
											Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
										End If
									Response.Write "</TD>"
									If lErrorNumber = 0 Then
										Response.Write "<TD>&nbsp;</TD>"
										Response.Write "<TD BGCOLOR=""#" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""1"" VALIGN=""TOP""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
										Response.Write "<TD>&nbsp;</TD>"
										Response.Write "<TD WIDTH=""100%"" VALIGN=""TOP"">"
											lErrorNumber = DisplayReport1503Parameters(oRequest, oADODBConnection, False, sErrorDescription)
										Response.Write "</TD>"
									Else
										lErrorNumber = 0
										sErrorDescription = ""
										aReportsComponent(B_HIDE_CONTINUE_REPORTS) = True
									End If
								Response.Write "</TR></TABLE>"
							Case 3
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1503(oRequest, oADODBConnection, aReportsComponent(N_ID_REPORTS), False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1504_REPORTS, ISSSTE_1701_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_BUDGET_AREA_FLAGS & "," & L_BUDGET_COMPANIES_FLAGS & "," & L_BUDGET_PROGRAM_DUTY_FLAGS & "," & L_BUDGET_FUND_FLAGS & "," & L_BUDGET_DUTY_FLAGS & "," & L_BUDGET_ACTIVE_DUTY_FLAGS & "," & L_BUDGET_SPECIFIC_DUTY_FLAGS & "," & L_BUDGET_PROGRAM_FLAGS & "," & L_BUDGET_REGION_FLAGS & "," & L_BUDGET_UR_FLAGS & "," & L_BUDGET_CT_FLAGS & "," & L_BUDGET_AUX_FLAGS & "," & L_BUDGET_LOCATION_FLAGS & "," & L_BUDGET_BUDGET1_FLAGS & "," & L_BUDGET_BUDGET2_FLAGS & "," & L_BUDGET_BUDGET3_FLAGS & "," & L_BUDGET_CONFINE_TYPE_FLAGS & "," & L_BUDGET_ACTIVITY1_FLAGS & "," & L_BUDGET_ACTIVITY2_FLAGS & "," & L_BUDGET_PROCESS_FLAGS & "," & L_BUDGET_YEAR_FLAGS & "," & L_BUDGET_MONTH_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								sFlags = L_NO_DIV_FLAGS & "," & L_BUDGET_AREA_FLAGS & "," & L_BUDGET_PROGRAM_DUTY_FLAGS & "," & L_BUDGET_FUND_FLAGS & "," & L_BUDGET_DUTY_FLAGS & "," & L_BUDGET_ACTIVE_DUTY_FLAGS & "," & L_BUDGET_SPECIFIC_DUTY_FLAGS & "," & L_BUDGET_PROGRAM_FLAGS & "," & L_BUDGET_REGION_FLAGS & "," & L_BUDGET_UR_FLAGS & "," & L_BUDGET_CT_FLAGS & "," & L_BUDGET_AUX_FLAGS & "," & L_BUDGET_LOCATION_FLAGS & "," & L_BUDGET_BUDGET1_FLAGS & "," & L_BUDGET_BUDGET2_FLAGS & "," & L_BUDGET_BUDGET3_FLAGS & "," & L_BUDGET_CONFINE_TYPE_FLAGS & "," & L_BUDGET_ACTIVITY1_FLAGS & "," & L_BUDGET_ACTIVITY2_FLAGS & "," & L_BUDGET_PROCESS_FLAGS & "," & L_BUDGET_YEAR_FLAGS
								lErrorNumber = DisplayReportTemplateForm(sFlags, sErrorDescription)
							Case 2
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								Response.Write "<DIV CLASS=""ReportFilter"">"
									lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "</DIV>"
							Case 3
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1504(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1561_REPORTS
						sFlags = L_BUDGET_ORIGINAL_POSITION_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_COMPANY_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_LEVEL_FLAGS & "," & L_ECONOMIC_ZONE_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								lErrorNumber = DisplayReport1503Saved(oRequest, oADODBConnection, sErrorDescription)
								If lErrorNumber <> 0 Then aReportsComponent(B_HIDE_CONTINUE_REPORTS) = True
							Case 2
								aReportComponent(N_ID_REPORT) = CLng(oRequest("RecordID").Item)
								lErrorNumber = GetReport(oRequest, oADODBConnection, aReportComponent, sErrorDescription)
								If lErrorNumber = 0 Then
									Call DisplayURLParametersAsHiddenValues(RemoveParameterFromURLString(RemoveParameterFromURLString(aReportComponent(S_PARAMETERS_REPORT), "ReportID"), "ReportStep"))
								End If
							Case 3
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									'lErrorNumber = BuildReport1561(oRequest, oADODBConnection, False, sErrorDescription)
									lErrorNumber = BuildReport1503(oRequest, oADODBConnection, aReportsComponent(N_ID_REPORTS), False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1562_REPORTS
						sFlags = L_BUDGET_ORIGINAL_POSITION_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_COMPANY_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_LEVEL_FLAGS & "," & L_ECONOMIC_ZONE_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								lErrorNumber = DisplayReport1503Saved(oRequest, oADODBConnection, sErrorDescription)
								If lErrorNumber <> 0 Then aReportsComponent(B_HIDE_CONTINUE_REPORTS) = True
							Case 2
								aReportComponent(N_ID_REPORT) = CLng(oRequest("RecordID").Item)
								lErrorNumber = GetReport(oRequest, oADODBConnection, aReportComponent, sErrorDescription)
								If lErrorNumber = 0 Then
									Call DisplayURLParametersAsHiddenValues(RemoveParameterFromURLString(RemoveParameterFromURLString(aReportComponent(S_PARAMETERS_REPORT), "ReportID"), "ReportStep"))
								End If
							Case 3
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									'lErrorNumber = BuildReport1561(oRequest, oADODBConnection, False, sErrorDescription)
									lErrorNumber = BuildReport1503(oRequest, oADODBConnection, aReportsComponent(N_ID_REPORTS), False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1563_REPORTS
						sFlags = L_BUDGET_ORIGINAL_POSITION_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_COMPANY_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_LEVEL_FLAGS & "," & L_ECONOMIC_ZONE_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								lErrorNumber = DisplayReport1503Saved(oRequest, oADODBConnection, sErrorDescription)
								If lErrorNumber <> 0 Then aReportsComponent(B_HIDE_CONTINUE_REPORTS) = True
							Case 2
								aReportComponent(N_ID_REPORT) = CLng(oRequest("RecordID").Item)
								lErrorNumber = GetReport(oRequest, oADODBConnection, aReportComponent, sErrorDescription)
								If lErrorNumber = 0 Then
									Call DisplayURLParametersAsHiddenValues(RemoveParameterFromURLString(RemoveParameterFromURLString(aReportComponent(S_PARAMETERS_REPORT), "ReportID"), "ReportStep"))
								End If
							Case 3
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1503(oRequest, oADODBConnection, aReportsComponent(N_ID_REPORTS), False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1571_REPORTS
						sFlags = L_BUDGET_ORIGINAL_POSITION_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_COMPANY_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_LEVEL_FLAGS & "," & L_ECONOMIC_ZONE_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								lErrorNumber = DisplayReport1503Saved(oRequest, oADODBConnection, sErrorDescription)
								If lErrorNumber <> 0 Then aReportsComponent(B_HIDE_CONTINUE_REPORTS) = True
							Case 2
								aReportComponent(N_ID_REPORT) = CLng(oRequest("RecordID").Item)
								lErrorNumber = GetReport(oRequest, oADODBConnection, aReportComponent, sErrorDescription)
								If lErrorNumber = 0 Then
									Call DisplayURLParametersAsHiddenValues(RemoveParameterFromURLString(RemoveParameterFromURLString(aReportComponent(S_PARAMETERS_REPORT), "ReportID"), "ReportStep"))
								End If
							Case 3
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									'lErrorNumber = BuildReport1561(oRequest, oADODBConnection, False, sErrorDescription)
									lErrorNumber = BuildReport1503(oRequest, oADODBConnection, aReportsComponent(N_ID_REPORTS), False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1581_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAYROLL_FLAGS & "," & L_COMPANY_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1581(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1582_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAYROLL_FLAGS & "," & L_COMPANY_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1582(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1583_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_YEARS_FLAGS & "," & L_COMPANY_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)

								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
									If Len(oRequest("YearID").Item) = 0 Then
										Response.Write "document.ReportFrm.YearID.value = " & Year(Date()) & ";" & vbNewLine
									End If
								Response.Write "//--></SCRIPT>" & vbNewLine
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1583(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1584_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAYROLL_FLAGS & "," & L_COMPANY_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1584(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1603_REPORTS
						sFlags = L_EMPLOYEE_NUMBER_FLAGS & "," & L_PAPERWORK_NUMBER_FLAGS & "," & L_PAPERWORK_START_DATE_FLAGS & "," & L_PAPERWORK_ESTIMATED_DATE_FLAGS & "," & L_PAPERWORK_END_DATE_FLAGS & "," & L_PAPERWORK_DOCUMENT_NUMBER_FLAGS & "," & L_PAPERWORK_TYPE_FLAGS & "," & L_PAPERWORK_OWNER_FLAGS & "," & L_PAPERWORK_STATUS_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1603(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1604_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAPERWORK_START_DATE_FLAGS & "," & L_PAPERWORK_ESTIMATED_DATE_FLAGS & "," & L_PAPERWORK_TYPE_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)

								Response.Write "<IMG SRC=""Images/Crcl4.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Incluir:<BR /></FONT>"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""Include"" ID=""IncludeChk"" VALUE=""1"" CHECKED=""1"" />&nbsp;Asuntos abiertos y defasados<BR />"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""Include"" ID=""IncludeChk"" VALUE=""2"" CHECKED=""1"" />&nbsp;Asuntos cerrados pero defasados<BR />"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""CHECKBOX"" NAME=""Include"" ID=""IncludeChk"" VALUE=""3"" CHECKED=""1"" />&nbsp;Asuntos resueltos (cerrados a tiempo)<BR />"
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1604(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1605_REPORTS
						sFlags = L_DATE_FLAGS & "," & L_DOCUMENT_FOR_LICENSE_NUMBER_FLAGS & "," & L_DOCUMENT_REQUEST_NUMBER_FLAGS & "," & L_DOCUMENT_FOR_CANCEL_LICENSE_NUMBER_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
									Response.Write "SendURLValuesToForm('"
									If Month(Date()) = 1 Then
										Response.Write "StartYear=" & Year(Date()) - 1 & "&StartMonth=12"
									Else
										Response.Write "StartYear=" & Year(Date()) & "&StartMonth=" & Right(("0" & Month(Date()) - 1), Len("00"))
									End If
									Response.Write "&StartDay=01', document.ReportFrm);" & vbNewLine
								Response.Write "//--></SCRIPT>" & vbNewLine
							Case 2
								If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 0 Then Response.Write "<DONT_EXPORT>"
									lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 0 Then Response.Write "</DONT_EXPORT>"
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1605(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1606_REPORTS
						sFlags = L_DATE_FLAGS & "," & L_DOCUMENT_FOR_LICENSE_NUMBER_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
									Response.Write "SendURLValuesToForm('"
									If Month(Date()) = 1 Then
										Response.Write "StartYear=" & Year(Date()) - 1 & "&StartMonth=12"
									Else
										Response.Write "StartYear=" & Year(Date()) & "&StartMonth=" & Right(("0" & Month(Date()) - 1), Len("00"))
									End If
									Response.Write "&StartDay=01', document.ReportFrm);" & vbNewLine
								Response.Write "//--></SCRIPT>" & vbNewLine
							Case 2
								If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 0 Then Response.Write "<DONT_EXPORT>"
									lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 0 Then Response.Write "</DONT_EXPORT>"
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1606(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1607_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAPERWORK_NUMBER_FLAGS & "," & L_PAPERWORK_START_DATE_FLAGS & "," & L_PAPERWORK_ESTIMATED_DATE_FLAGS & "," & L_PAPERWORK_END_DATE_FLAGS & "," & L_PAPERWORK_DOCUMENT_NUMBER_FLAGS & "," & L_PAPERWORK_TYPE_FLAGS & "," & L_PAPERWORK_OWNER_FLAGS & "," & L_PAPERWORK_STATUS_FLAGS & "," & L_PAPERWORK_PRIORITY_FLAGS & "," & L_PAPERWORK_OWNERS_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<IMG SRC=""Images/Crcl11.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Pendiente de descargo:<BR /></FONT>"
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""Closed"" ID=""ClosedLst"" SIZE=""1"" CLASS=""Lists"">"
									Response.Write "<OPTION VALUE="""""
										If Len(oRequest("Closed").Item) = 0 Then Response.Write " SELECTED=""1"""
									Response.Write ">Todos</OPTION>"
									Response.Write "<OPTION VALUE=""1"""
										If StrComp(oRequest("Closed").Item, "1", vbBinaryCompare) = 0 Then Response.Write " SELECTED=""1"""
									Response.Write ">No</OPTION>"
									Response.Write "<OPTION VALUE=""0"""
										If StrComp(oRequest("Closed").Item, "0", vbBinaryCompare) = 0 Then Response.Write " SELECTED=""1"""
									Response.Write ">Sí</OPTION>"
								Response.Write "</SELECT><BR /><BR />"

								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
									Response.Write "SendURLValuesToForm('"
									If Len(oRequest("YearID").Item) = 0 Then
										If Month(Date()) = 1 Then
											Response.Write "YearID=" & Year(Date()) - 1
										Else
											Response.Write "YearID=" & Year(Date())
										End If
									Else
										Response.Write "YearID=" & oRequest("YearID").Item
									End If
									Response.Write "', document.ReportFrm);" & vbNewLine
								Response.Write "//--></SCRIPT>" & vbNewLine
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1607(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1608_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAPERWORK_FOLIO_NUMBER_FLAGS & "," & L_PAPERWORK_START_DATE_FLAGS & "," & L_PAPERWORK_DOCUMENT_NUMBER_FLAGS & "," & L_PAPERWORK_TYPE_FLAGS & "," & L_PAPERWORK_OWNERS_FLAGS & "," & L_PAPERWORK_STATUS_FLAGS & "," & L_PAPERWORK_SUBJECT_TYPES & "," & L_PAPERWORK_PRIORITY_FLAGS & "," & L_PAPERWORK_ESTIMATED_DATE_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1608(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1609_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAPERWORK_FOLIO_NUMBER_FLAGS & "," & L_PAPERWORK_START_DATE_FLAGS & "," & L_PAPERWORK_DOCUMENT_NUMBER_FLAGS & "," & L_PAPERWORK_TYPE_FLAGS & "," & L_PAPERWORK_OWNERS_FLAGS & "," & L_PAPERWORK_STATUS_FLAGS & "," & L_PAPERWORK_SUBJECT_TYPES & "," & L_PAPERWORK_PRIORITY_FLAGS & "," & L_PAPERWORK_ESTIMATED_DATE_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								aTemplateValue = Split(Replace(oRequest("Template").Item, " ", "", 1, -1, vbBinaryCompare), ",", -1, vbBinaryCompare)
								For iIndex = 0 To UBound(aTemplateValue)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Template"" ID=""TemplateHdn"" VALUE=""" & aTemplateValue(iIndex) & """ />"
								Next
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
							Case 2
								lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								'lErrorNumber = DisplayReport1609Table(oRequest, oADODBConnection, False, "", sErrorDescription)
								lErrorNumber = DisplayReport1609TableFull(oRequest, oADODBConnection, False, "", sErrorDescription)
								Call DisplayURLParametersAsHiddenValues(RemoveParameterFromURLString(RemoveParameterFromURLString(RemoveParameterFromURLString(oRequest, "ReportID"), "ReportType"), "ReportStep"))
                            Case 3
								'lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If lErrorNumber = 0 Then
									'lErrorNumber = BuildReport1609(oRequest, oADODBConnection, sErrorDescription)
									lErrorNumber = BuildReport1609Full(oRequest, oADODBConnection, sErrorDescription)
								End If
						End Select
					Case ISSSTE_1610_REPORTS, ISSSTE_1611_REPORTS, ISSSTE_1612_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAPERWORK_START_DATE_FLAGS & "," & L_PAPERWORK_DOCUMENT_NUMBER_FLAGS & "," & L_PAPERWORK_OWNER_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								If aReportsComponent(N_ID_REPORTS) = ISSSTE_1610_REPORTS Then
									Response.Write "<IMG SRC=""Images/Crcl2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
									Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Mostrar los asuntos a nivel de:<BR /></FONT>"
									Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""PaperworkLevelID"" ID=""PaperworkLevelIDLst"" SIZE=""1"" CLASS=""Lists"">"
										Response.Write "<OPTION VALUE=""1"""
											If StrComp(oRequest("PaperworkLevelID").Item, "1", vbBinaryCompare) = 0 Then Response.Write " SELECTED=""1"""
										Response.Write ">Subdirección</OPTION>"
										Response.Write "<OPTION VALUE=""2"""
											If StrComp(oRequest("PaperworkLevelID").Item, "2", vbBinaryCompare) = 0 Then Response.Write " SELECTED=""1"""
										Response.Write ">Jefatura de servicio</OPTION>"
										Response.Write "<OPTION VALUE=""3"""
											If StrComp(oRequest("PaperworkLevelID").Item, "3", vbBinaryCompare) = 0 Then Response.Write " SELECTED=""1"""
										Response.Write ">Jefatura de depto</OPTION>"
									Response.Write "</SELECT><BR /><BR />"
								ElseIf aReportsComponent(N_ID_REPORTS) = ISSSTE_1612_REPORTS Then
									Response.Write "<IMG SRC=""Images/Crcl2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
									Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Pendiente de descargo:<BR /></FONT>"
									Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""Closed"" ID=""ClosedLst"" SIZE=""1"" CLASS=""Lists"">"
										Response.Write "<OPTION VALUE="""""
											If Len(oRequest("Closed").Item) = 0 Then Response.Write " SELECTED=""1"""
										Response.Write ">Todos</OPTION>"
										Response.Write "<OPTION VALUE=""1"""
											If StrComp(oRequest("Closed").Item, "1", vbBinaryCompare) = 0 Then Response.Write " SELECTED=""1"""
										Response.Write ">No</OPTION>"
										Response.Write "<OPTION VALUE=""0"""
											If StrComp(oRequest("Closed").Item, "0", vbBinaryCompare) = 0 Then Response.Write " SELECTED=""1"""
										Response.Write ">Sí</OPTION>"
									Response.Write "</SELECT><BR /><BR />"

									Response.Write "<IMG SRC=""Images/Crcl3.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
									Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Mostrar los asuntos a nivel de:<BR /></FONT>"
									Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""PaperworkLevelID"" ID=""PaperworkLevelIDLst"" SIZE=""1"" CLASS=""Lists"">"
										Response.Write "<OPTION VALUE=""1"""
											If StrComp(oRequest("PaperworkLevelID").Item, "1", vbBinaryCompare) = 0 Then Response.Write " SELECTED=""1"""
										Response.Write ">Subdirección</OPTION>"
										Response.Write "<OPTION VALUE=""2"""
											If StrComp(oRequest("PaperworkLevelID").Item, "2", vbBinaryCompare) = 0 Then Response.Write " SELECTED=""1"""
										Response.Write ">Jefatura de servicio</OPTION>"
										Response.Write "<OPTION VALUE=""3"""
											If StrComp(oRequest("PaperworkLevelID").Item, "3", vbBinaryCompare) = 0 Then Response.Write " SELECTED=""1"""
										Response.Write ">Jefatura de depto</OPTION>"
									Response.Write "</SELECT><BR /><BR />"
								End If
								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
									Response.Write "SendURLValuesToForm('"
									If Len(oRequest("YearID").Item) = 0 Then
										If Month(Date()) = 1 Then
											Response.Write "YearID=" & Year(Date()) - 1
										Else
											Response.Write "YearID=" & Year(Date())
										End If
									Else
										Response.Write "YearID=" & oRequest("YearID").Item
									End If
									Response.Write "', document.ReportFrm);" & vbNewLine
								Response.Write "//--></SCRIPT>" & vbNewLine
							Case 2
								If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 0 Then Response.Write "<DONT_EXPORT>"
									lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 0 Then Response.Write "</DONT_EXPORT>"
								If lErrorNumber = 0 Then
									Select Case aReportsComponent(N_ID_REPORTS)
										Case ISSSTE_1610_REPORTS
											lErrorNumber = BuildReport1610(oRequest, oADODBConnection, sErrorDescription)
										Case ISSSTE_1611_REPORTS
											lErrorNumber = BuildReport1611(oRequest, oADODBConnection, sErrorDescription)
										Case ISSSTE_1612_REPORTS
											lErrorNumber = BuildReport1612(oRequest, oADODBConnection, sErrorDescription)
									End Select
								End If
						End Select
					Case ISSSTE_1613_REPORTS
						sFlags = L_NO_DIV_FLAGS & "," & L_PAPERWORK_START_DATE_FLAGS & "," & L_PAPERWORK_OWNERS_FLAGS & "," & L_ZIP_WARNING_FLAGS
						Select Case aReportsComponent(N_STEP_REPORTS)
							Case 1
								lErrorNumber = DisplayReportFilter(sFlags, sErrorDescription)
								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
									Response.Write "SendURLValuesToForm('"
									If Len(oRequest("YearID").Item) = 0 Then
										If Month(Date()) = 1 Then
											Response.Write "YearID=" & Year(Date()) - 1
										Else
											Response.Write "YearID=" & Year(Date())
										End If
									Else
										Response.Write "YearID=" & oRequest("YearID").Item
									End If
									Response.Write "', document.ReportFrm);" & vbNewLine
								Response.Write "//--></SCRIPT>" & vbNewLine
							Case 2
								If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 0 Then Response.Write "<DONT_EXPORT>"
									lErrorNumber = DisplayFilterInformation(oRequest, sFlags, False, "", sErrorDescription)
								If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 0 Then Response.Write "</DONT_EXPORT>"
								If lErrorNumber = 0 Then
									lErrorNumber = BuildReport1613(oRequest, oADODBConnection, False, sErrorDescription)
								End If
						End Select
				End Select
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReportID"" ID=""ReportIDHdn"" VALUE=""" & aReportsComponent(N_ID_REPORTS) & """ /><BR />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReportStep"" ID=""ReportStepHdn"" VALUE=""" & (aReportsComponent(N_STEP_REPORTS) + 1) & """ /><BR />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FromSectionID"" ID=""FromSectionIDHdn"" VALUE=""" & oRequest("FromSectionID").Item & """ />"
				Response.Write "<DONT_EXPORT><INPUT TYPE=""BUTTON"" NAME=""Back"" ID=""BackBtn"" VALUE=""Regresar"" onClick="""
					If aReportsComponent(N_STEP_REPORTS) > 1 Then
						'Response.Write "document.ModifyReportFrm.ReportStep.value='" & aReportsComponent(N_STEP_REPORTS) - 1 & "'; document.ModifyReportFrm.submit();"
                        If aReportsComponent(N_STEP_REPORTS) = 3 Then
                            If aReportsComponent(N_ID_REPORTS) = ISSSTE_1609_REPORTS Then
                                Select Case CInt(Request.Cookies("SIAP_SectionID"))
                                    Case 1
                                        Response.Write "window.location.href='Main_ISSSTE.asp?SectionID=61';"
                                    Case Else
                                        Response.Write "window.location.href='Main_ISSSTE.asp?SectionID=61';"
                                End Select
                            Else
						        Response.Write "document.ModifyReportFrm.ReportStep.value='" & aReportsComponent(N_STEP_REPORTS) - 1 & "'; document.ModifyReportFrm.submit();"
                            End If
                        Else
                            Response.Write "document.ModifyReportFrm.ReportStep.value='" & aReportsComponent(N_STEP_REPORTS) - 1 & "'; document.ModifyReportFrm.submit();"
                        End If
					Else
						Select Case CInt(Request.Cookies("SIAP_SectionID"))
							Case 1
								Select Case aReportsComponent(N_ID_REPORTS)
									Case 1101
										Response.Write "window.location.href='Main_ISSSTE.asp?SectionID=1';"
									Case 1151, 1152, 1153, 1154, 1155, 1157
										Response.Write "window.location.href='Main_ISSSTE.asp?SectionID=15';"
									Case Else
										Response.Write "window.location.href='Main_ISSSTE.asp?SectionID=19';"
								End Select
							Case 2
								Select Case aReportsComponent(N_ID_REPORTS)
									Case 1151
										Response.Write "window.location.href='Main_ISSSTE.asp?SectionID=291';"
									Case 1203, 1204, 1205, 1206, 1207, 1208
										Response.Write "window.location.href='Main_ISSSTE.asp?SectionID=25';"
									Case 1221, 1222, 1225
										Response.Write "window.location.href='Main_ISSSTE.asp?SectionID=21';"
									Case 1223, 1224
										Response.Write "window.location.href='Main_ISSSTE.asp?SectionID=23';"
									Case 1603, 1604, 1607, 1608, 1609, 1610, 1611, 1612
										Response.Write "window.location.href='Main_ISSSTE.asp?SectionID=61';"
									Case Else
										Response.Write "window.location.href='Main_ISSSTE.asp?SectionID=24';"
								End Select
							Case 3
								Response.Write "window.location.href='Main_ISSSTE.asp?SectionID=34';"
							Case 4
								Select Case aReportsComponent(N_ID_REPORTS)
									Case 1470, 1471, 1472, 1473, 1474
										Response.Write "window.location.href='Main_ISSSTE.asp?SectionID=47';"
									Case 2420, 2421, 2422, 2423, 2426, 2427, 2428, 2429, 2430, 1431, 2431, 1432, 2432
										Response.Write "window.location.href='Main_ISSSTE.asp?SectionID=42';"
									Case Else
										Response.Write "window.location.href='Main_ISSSTE.asp?SectionID=49';"
								End Select
							Case 5
								Select Case aReportsComponent(N_ID_REPORTS)
									Case 1503
										Response.Write "window.location.href='Main_ISSSTE.asp?SectionID=53';"
									Case 1504
										Response.Write "window.location.href='Main_ISSSTE.asp?SectionID=5';"
									Case 1561, 1562, 1563
										Response.Write "window.location.href='Main_ISSSTE.asp?SectionID=56';"
									Case Else
										Response.Write "window.location.href='Main_ISSSTE.asp?SectionID=58';"
								End Select
							Case 6
								Response.Write "window.location.href='Main_ISSSTE.asp?SectionID=64';"
							Case 7
								If CInt(oRequest("SubSectionID").Item) = 1 Then
									Response.Write "window.location.href='Main_ISSSTE.asp?SectionID=713&SubSectionID=1';"
								ElseIf CInt(oRequest("SubSectionID").Item) = 4 Then
									Response.Write "window.location.href='Main_ISSSTE.asp?SectionID=733&SubSectionID=4';"
								Else
									Response.Write "window.location.href='Main_ISSSTE.asp?SectionID=73';"
								End If
							Case Else
								Response.Write "window.history.go(-1);"
						End Select
					End If
					Response.Write """ CLASS=""Buttons"" /></DONT_EXPORT>"
				If Not aReportsComponent(B_READY_REPORTS) And Not aReportsComponent(B_HIDE_CONTINUE_REPORTS) Then
					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""200"" HEIGHT=""1"" />"
					Response.Write "<SPAN NAME=""ContinueSpn"" ID=""ContinueSpn""><INPUT TYPE=""SUBMIT"" VALUE=""Continuar"" CLASS=""Buttons"" /></SPAN>"
				End If
			Response.Write "</FORM></DIV>"

			If lErrorNumber <> 0 Then
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
			ElseIf ((aReportsComponent(N_ID_REPORTS) = ISSSTE_1561_REPORTS) Or (aReportsComponent(N_ID_REPORTS) = ISSSTE_1562_REPORTS) Or (aReportsComponent(N_ID_REPORTS) = ISSSTE_1563_REPORTS) Or (aReportsComponent(N_ID_REPORTS) = ISSSTE_1571_REPORTS)) And (aReportsComponent(N_STEP_REPORTS) = 2) Then
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "document.ReportFrm.submit();" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
			End If
		End If
		sNames = GetSerialNumberForDate("")
		If aReportsComponent(B_READY_REPORTS) Then
			Response.Write "<FORM NAME=""ExportFrm"" ID=""ExportFrm"" ACTION=""Export.asp?TempFile=" & sNames & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & """ METHOD=""POST"" TARGET=""ExportToExcel"">"
				Call DisplayURLParametersAsHiddenValues("Action=Reports&Excel=1&ColorIndex=" & iColorIndex & "&TempFile=" & sNames & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&" & RemoveEmptyParametersFromURLString(oRequest))
			Response.Write "</FORM>"
			Response.Write "<FORM NAME=""SaveReportFrm"" ID=""SaveReportFrm"" ACTION=""SavedReport.asp"" METHOD=""POST"">"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""New"" ID=""NewHdn"" VALUE=""1"" />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReportType"" ID=""ReportTypeHdn"" VALUE=""1"" />"
				Call DisplayURLParametersAsHiddenValues(RemoveParameterFromURLString(oRequest, "ReportType"))
			Response.Write "</FORM>"
		End If
		Response.Write "<FORM NAME=""ModifyReportFrm"" ID=""ModifyReportFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""POST"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReportStep"" ID=""ReportStepHdn"" VALUE=""1"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReportType"" ID=""ReportTypeHdn"" VALUE=""1"" />"
			Select Case aReportsComponent(N_ID_REPORTS)
				Case ISSSTE_1609_REPORTS
				Case Else
					Call DisplayURLParametersAsHiddenValues(RemoveParameterFromURLString(RemoveParameterFromURLString(oRequest, "ReportType"), "ReportStep"))
			End Select
		Response.Write "</FORM>"
		%>
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>