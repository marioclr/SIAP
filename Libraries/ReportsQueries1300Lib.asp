<%
Function BuildReport1311(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de plantilla de personal
'         Carpeta 1. DOCUMENTACIÓN ENTREGADA POR JEFATURA DE SERVICIOS DE DESARROLLO HUMANO
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1311"
	Dim sHeaderContents
	Dim oRecordset
	Dim lCurrentID
	Dim sContents
	Dim sRowContents
	Dim lErrorNumber
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim sFullFile
	Dim sSourceFolderPath
	Dim lPayrollID
	Dim lForPayrollID
	Dim sCondition
	Dim iIndex
	Dim sPreviousZone
	Dim sTables

	oStartDate = Now()
	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	sTables = ""
	If ((InStr(1, sCondition, "=Companies.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(Companies.", vbBinaryCompare) > 0)) Then sTables = ", Companies"
	If ((InStr(1, sCondition, "=StatusEmployees.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(StatusEmployees.", vbBinaryCompare) > 0)) Then sTables = sTables & ", StatusEmployees"
	If ((InStr(1, sCondition, "=Genders.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(Genders.", vbBinaryCompare) > 0)) Then sTables = sTables & ", Genders"
	If ((InStr(1, sCondition, "=JobTypes.", vbBinaryCompare) > 0) Or (InStr(1, sCondition, "(JobTypes.", vbBinaryCompare) > 0)) Then sTables = sTables & ", JobTypes"
	If lErrorNumber = 0 Then
		If lErrorNumber = 0 Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
			sFilePath = sFilePath & "\"
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
			sFullFile = sFilePath & "PLANTILLA_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			Response.Flush()
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudieron obtener los registros de los empleados."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Employees.RFC, Employees.CURP, Employees.StartDate, PositionTypes.PositionTypeShortName, EmployeeTypes.EmployeeTypeShortName, Journeys.JourneyName, Shifts.ShiftShortName, Shifts.ShiftName, Shifts.StartHour1, Shifts.EndHour1, Shifts.StartHour2, Shifts.EndHour2, Shifts.WorkingHours, Employees.SocialSecurityNumber, ZoneTypes.ZoneTypeName, Positions.PositionShortName, Levels.LevelShortName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryListForPayroll.JobID, EmployeesHistoryListForPayroll.IntegrationID, EmployeesHistoryListForPayroll.ClassificationID, PaymentCenters.AreaCode, PaymentCenters.AreaName, Services.ServiceShortName, Services.ServiceName, Areas.AreaName, Areas.AreaCode, EmployeesHistoryListForPayroll.ZoneID, Zones.ZoneName, ParentZones.ZoneName As ParentZoneName, Entidades.ZoneName As EntidadesZoneName, GeneratingAreas.GeneratingAreaName, Payroll_" & lPayrollID & ".ConceptID, Payroll_" & lPayrollID & ".ConceptAmount From Payroll_" & lPayrollID & ", Employees, EmployeesHistoryListForPayroll, Areas, EmployeeTypes, GeneratingAreas, GroupGradeLevels, Journeys, Levels, Areas As PaymentCenters, Positions, PositionTypes, Services, Shifts, Zones, Zones As ParentZones, Zones As Entidades, ZoneTypes" & sTables & " Where (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.ShiftID=Shifts.ShiftID) And (EmployeesHistoryListForPayroll.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesHistoryListForPayroll.ServiceID=Services.ServiceID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (Zones.ParentID=ParentZones.ZoneID) And (ParentZones.ParentID=Entidades.ZoneID) And (Areas.GeneratingAreaID=GeneratingAreas.GeneratingAreaID) And (EmployeesHistoryListForPayroll.JourneyID=Journeys.JourneyID) And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (GeneratingAreas.StartDate<=" & lForPayrollID & ") And (GeneratingAreas.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (Journeys.StartDate<=" & lForPayrollID & ") And (Journeys.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Services.StartDate<=" & lForPayrollID & ") And (Services.EndDate>=" & lForPayrollID & ") And (Shifts.StartDate<=" & lForPayrollID & ") And (Shifts.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Payroll_" & lPayrollID & ".RecordDate=" & lPayrollID & ") " & sCondition & " Order By EmployeesHistoryListForPayroll.ZoneID, Entidades.ZoneName, EmployeesHistoryListForPayroll.EmployeeID, Payroll_" & lPayrollID & ".ConceptID", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select Employees.EmployeeID, Employees.EmployeeNumber, Employees.EmployeeLastName, Employees.EmployeeLastName2, Employees.EmployeeName, Employees.RFC, Employees.CURP, Employees.StartDate, PositionTypes.PositionTypeShortName, EmployeeTypes.EmployeeTypeShortName, Journeys.JourneyName, Shifts.ShiftShortName, Shifts.ShiftName, Shifts.StartHour1, Shifts.EndHour1, Shifts.StartHour2, Shifts.EndHour2, Shifts.WorkingHours, Employees.SocialSecurityNumber, ZoneTypes.ZoneTypeName, Positions.PositionShortName, Levels.LevelShortName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryListForPayroll.JobID, EmployeesHistoryListForPayroll.IntegrationID, EmployeesHistoryListForPayroll.ClassificationID, PaymentCenters.AreaCode, PaymentCenters.AreaName, Services.ServiceShortName, Services.ServiceName, Areas.AreaName, Areas.AreaCode, EmployeesHistoryListForPayroll.ZoneID, Zones.ZoneName, ParentZones.ZoneName As ParentZoneName, Entidades.ZoneName As EntidadesZoneName, GeneratingAreas.GeneratingAreaName, Payroll_" & lPayrollID & ".ConceptID, Payroll_" & lPayrollID & ".ConceptAmount From Payroll_" & lPayrollID & ", Employees, EmployeesHistoryListForPayroll, Areas, EmployeeTypes, GeneratingAreas, GroupGradeLevels, Journeys, Levels, Areas As PaymentCenters, Positions, PositionTypes, Services, Shifts, Zones, Zones As ParentZones, Zones As Entidades, ZoneTypes" & sTables & " Where (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.ShiftID=Shifts.ShiftID) And (EmployeesHistoryListForPayroll.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesHistoryListForPayroll.ServiceID=Services.ServiceID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (Zones.ParentID=ParentZones.ZoneID) And (ParentZones.ParentID=Entidades.ZoneID) And (Areas.GeneratingAreaID=GeneratingAreas.GeneratingAreaID) And (EmployeesHistoryListForPayroll.JourneyID=Journeys.JourneyID) And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (EmployeeTypes.StartDate<=" & lForPayrollID & ") And (EmployeeTypes.EndDate>=" & lForPayrollID & ") And (GeneratingAreas.StartDate<=" & lForPayrollID & ") And (GeneratingAreas.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (Journeys.StartDate<=" & lForPayrollID & ") And (Journeys.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") And (Services.StartDate<=" & lForPayrollID & ") And (Services.EndDate>=" & lForPayrollID & ") And (Shifts.StartDate<=" & lForPayrollID & ") And (Shifts.EndDate>=" & lForPayrollID & ") And (Zones.StartDate<=" & lForPayrollID & ") And (Zones.EndDate>=" & lForPayrollID & ") And (Payroll_" & lPayrollID & ".RecordDate=" & lPayrollID & ") " & sCondition & " Order By EmployeesHistoryListForPayroll.ZoneID, Entidades.ZoneName, EmployeesHistoryListForPayroll.EmployeeID, Payroll_" & lPayrollID & ".ConceptID -->" & vbNewLine
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						sPreviousZone = ""
						lCurrentID = -2
						Do While Not oRecordset.EOF
							If StrComp(sPreviousZone, CStr(oRecordset.Fields("EntidadesZoneName").Value), vbBinaryCompare) <> 0 Then
								If Len(sPreviousZone) > 0 Then
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									lErrorNumber = AppendTextToFile(sFullFile, sRowContents, sErrorDescription)
									lCurrentID = -2
									lErrorNumber = AppendTextToFile(sDocumentName, "</TABLE>", sErrorDescription)
								End If
								sDocumentName = sFilePath & "PLANTILLA_" & CStr(oRecordset.Fields("EntidadesZoneName").Value) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"
								sHeaderContents = Replace(sHeaderContents, "<PAYROLL_DATE />", DisplayNumericDateFromSerialNumber(lForPayrollID))
								sRowContents = "<TABLE BORDER=""1"">"
								sRowContents = sRowContents & "<TR><TD COLSPAN=""140"" ALIGN=""CENTER"">REPORTE DE PLANTILLA DE PERSONAL DE LA NÓMINA: " & DisplayNumericDateFromSerialNumber(lForPayrollID) & "</TD></TR>"
								sRowContents = sRowContents & "<TR><TD>NUM EMPLEADO</TD><TD>APELLIDO PATERNO</TD><TD>APELLIDO MATERNO</TD><TD>NOMBRE(S)</TD><TD>RFC</TD><TD>CURP</TD><TD>FECHA ALTA</TD><TD>IMPUTACION</TD><TD>FECHA PAGO</TD><TD>TIPO PTO</TD><TD>TAB</TD><TD>TURNO</TD><TD>CODIGO DE HORARIO</TD><TD>DENOMINACION</TD><TD>HORA DE ENTRADA</TD><TD>HORA DE SALIDA</TD><TD>HORA DE ENTRADA</TD><TD>HORA DE SALIDA</TD><TD>NUM. HORAS LABORADAS</TD><TD>NUM. DE SEGURIDAD SOCIAL</TD><TD>PLAZA</TD><TD>ZONA ECO</TD><TD>PUESTO</TD><TD>NIVEL</TD><TD>GRUPO, GRADO Y NIV SAL</TD><TD>INTEGRA SALARIAL</TD><TD>CLASIFICACION</TD><TD>DENOMINACION</TD><TD>CENTRO PAGO</TD><TD>SERVICIO</TD><TD>DENOMINACION</TD><TD>ADSCRIPCION</TD><TD>TRONCAL</TD><TD>CENTRO DE TRABAJO</TD><TD>POBLACION</TD><TD>MUNICIPIO</TD><TD>ENTIDAD</TD><TD>AREA GENERADORA</TD><TD>DIV_GEO</TD><TD>SUELDO</TD><TD>PREV.SOC.AYU.MULTIPLE</TD><TD>COMPENS. GARANTIZADA</TD><TD>COMPENSACION POR RIESGOS PROFES</TD><TD>COMPENSACION POR ANTIGÜEDAD</TD><TD>QUINQUENIOS</TD><TD>TURNO OPCIONAL</TD><TD>PERCEP ADICIONAL</TD><TD>HORAS EXTRAS</TD><TD>HORAS EXTRAS DOBLES</TD><TD>HORAS EXTRAS TRIPLES</TD><TD>DESPENSA</TD><TD>BECA PASANTES INTERINOS Y PREGRADO</TD><TD>AYUDA RENTA BECARIO</TD><TD>PRIMA DOMINICAL EXCENTA</TD><TD>PRIMA DOMINICAL GRAVABLE</TD><TD>REMUN. GUARDIAS</TD><TD>DEVOLUCION DEDUCCIONES INDEBIDAS</TD><TD>PRIMA VACACIONAL EXCENTA</TD><TD>PRIMA VACACIONAL GRAVABLE</TD><TD>BECA HIJOS DE TRABAJA</TD><TD>AYUDA DE ANTEOJOS</TD><TD>PREMIO ANIVERSARIO</TD><TD>PREMIO 10 DE MAYO</TD><TD>PREMIO POR ANTIGÜEDAD</TD><TD>ESTÍMULO ADICIONAL</TD><TD>PREMIO MONEDA DE ORO</TD><TD>ESTIM.PROD.CAL.PERS.MED</TD><TD>MATERIAL DIDACTICO</TD><TD>REMUN.SUPLENCIAS</TD><TD>BONO DE REYES</TD><TD>AYUDA TRANSPORTE</TD><TD>AYUDA COMPRA DE UTILES</TD><TD>ASIG. MEDICA</TD><TD>COMPLEMENTO DE BECA MEDICOS RESIDENTES</TD><TD>ESTÍMULO ASISTENCIA</TD><TD>ESTÍMULO PUNTUALIDAD</TD><TD>ESTÍMULO DESEMPEÑO</TD><TD>ESTÍMULO MÉRITO RELEVANTE</TD><TD>PREM.ANTIG.25 Y 30 AÑOS</TD><TD>AYUDA MUERTE FAM.1er.G</TD><TD>AYUDA IMPRES.TESIS</TD><TD>APOYO DESAR.CAPACITAC</TD><TD>CREDITO AL SALARIO</TD><TD>A.G.A.</TD><TD>PREMIO TRAB. DEL MES</TD><TD>BECA MEDIC.RESIDENTES</TD><TD>AJUSTE RESIDENTES</TD><TD>PAGO RETIRO</TD><TD>AJUSTE CALENDARIO</TD><TD>PREM.ESTIM.RECOMPENSA</TD><TD>GUARDIAS PROVAC</TD><TD>SUPLENCIAS PROVAC</TD><TD>DEVOL.NO GRAVABLES</TD><TD>GRATIFICACION MES DE BECA</TD><TD>REZAGO QUIRÚRGICO</TD><TD>AHORRO SOLIDARIO</TD><TD>PERCEPCIONES</TD><TD>INASISTENCIAS</TD><TD>SERVICIO MEDICO</TD><TD>FONDO DE PRESTACIONES</TD><TD>I.S.P.T</TD><TD>CUOTA SINDICAL SNTISSSTE</TD><TD>SEGURO HIPOTECARIO</TD><TD>CREDITO HIPOTECARIO</TD><TD>SEGURO HIPOTECARIO AVA</TD><TD>PRESTAMO PERSONAL</TD><TD>COMISION AUXILIO</TD><TD>CREDITO FOVISSSTE</TD><TD>SEGURO VIDA HGO 1</TD><TD>SEGURO VIDA HGO 2</TD><TD>SEGURO INSTITUCIONAL</TD><TD>SEGURO DEL RETIRO</TD><TD>CUOTA DEPORTIVO</TD><TD>PENSION ALIMENTICIA</TD><TD>RETARDOS</TD><TD>REINT. SUELDOS COBROS INDEBIDOS</TD><TD>OTRAS DEDUCCIONES</TD><TD>PRESTAMO AUTOMOVIL SERVIDORES PUBLICOS SUPERIORES</TD><TD>SEGURO COMERCIAL AMER</TD><TD>AJUSTES FONAC</TD><TD>FONDO AHORRO CAPITALIZ</TD><TD>PRESTAMO AUTOMOVIL MANDOS MEDIOS</TD><TD>APORTACIONES VOLUNTARIAS AL SAR</TD><TD>SERV.GASTOS FUN.NASER</TD><TD>SEGURO AUTO PROVINCIAL</TD><TD>SEGURO DAÑOS FOVISSSTE</TD><TD>SEGURO DE VIDA AIG</TD><TD>CREDITO AHORRA YA</TD><TD>ADEUDO PENSION ALIMENT</TD><TD>ISR SALARIOS NOM EXTRAOR</TD><TD>COBRO PUESTO ANTERIOR</TD><TD>CUOTA SINDICAL SINDICATO INDEPENDIENTE</TD><TD>SEGURO DE SALUD</TD><TD>SEGURO DE SALUD DE LOS PENSIONISTAS</TD><TD>SEGURO DE INVALIDEZ Y VIDA</TD><TD>SERVICIOS SOCIALES Y CULTURALES</TD><TD>SEGURO DE RETIRO, CESANTIA EN EDAD AVANZADA Y VEJEZ</TD><TD>DEDUCCIONES</TD><TD>LIQUIDO</TD><TD>OBSERVACIONES</TD></TR>"
								lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
								If Len(sPreviousZone) = 0 Then lErrorNumber = AppendTextToFile(sFullFile, sRowContents, sErrorDescription)
								sPreviousZone = CStr(oRecordset.Fields("EntidadesZoneName").Value)
							End If

							If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
								If lCurrentID > -2 Then
									lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
									lErrorNumber = AppendTextToFile(sFullFile, sRowContents, sErrorDescription)
								End If
								sRowContents = "<TR>"
									sRowContents = sRowContents & "<TD>=T(""" & CStr(oRecordset.Fields("EmployeeNumber").Value) & """)</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("EmployeeLastName").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>"
										If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
											sRowContents = sRowContents & CStr(oRecordset.Fields("EmployeeLastName2").Value)
										Else
											sRowContents = sRowContents & " "
										End If
									sRowContents = sRowContents & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("EmployeeName").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>"
										sRowContents = sRowContents & CStr(oRecordset.Fields("RFC").Value)
										Err.Clear
									sRowContents = sRowContents & "</TD>"
									sRowContents = sRowContents & "<TD>"
										sRowContents = sRowContents & CStr(oRecordset.Fields("CURP").Value)
										Err.Clear
									sRowContents = sRowContents & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("StartDate").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & DisplayNumericDateFromSerialNumber(lForPayrollID) & "</TD>"
									sRowContents = sRowContents & "<TD>" & DisplayNumericDateFromSerialNumber(lPayrollID) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("PositionTypeShortName").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("EmployeeTypeShortName").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("JourneyName").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("ShiftShortName").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("ShiftName").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("StartHour1").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("EndHour1").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("StartHour2").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("EndHour2").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("WorkingHours").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>"
										sRowContents = sRowContents & CStr(oRecordset.Fields("SocialSecurityNumber").Value)
										Err.Clear
									sRowContents = sRowContents & "</TD>"
									sRowContents = sRowContents & "<TD>=T(""" & CStr(oRecordset.Fields("JobID").Value) & """)</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("ZoneTypeName").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("PositionShortName").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("LevelShortName").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("GroupGradeLevelShortName").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("IntegrationID").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("ClassificationID").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("PaymentCenterShortName").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("PaymentCenterName").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("ServiceShortName").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("ServiceName").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("AreaName").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("AreaCode").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("ZoneName").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("ParentZoneName").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("ParentZoneName").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("EntidadesZoneName").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("GeneratingAreaName").Value) & "</TD>"
									sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("EntidadesZoneName").Value) & "</TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_1 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_2 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_3 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_4 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_5 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_6 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_7 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_8 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_9 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_10 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_12 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_14 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_15 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_16 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_17 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_18 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_19 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_20 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_21 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_22 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_23 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_24 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_25 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_26 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_27 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_29 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_31 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_32 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_33 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_34 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_35 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_36 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_37 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_38 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_39 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_40 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_41 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_42 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_43 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_44 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_45 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_46 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_47 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_48 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_49 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_50 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_89 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_90 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_91 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_92 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_94 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_96 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_97 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_100 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_107 /></TD>"
									sRowContents = sRowContents & "<TD>0</TD>"
									sRowContents = sRowContents & "<TD>0</TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_-1 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_50 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_53 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_54 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_55 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_56 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_57 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_58 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_59 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_61 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_63 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_64 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_65 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_66 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_67 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_51 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_69 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_70 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_71 /></TD>"
									sRowContents = sRowContents & "<TD>0</TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_73 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_74 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_75 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_76 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_77 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_78 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_79 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_80 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_81 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_83 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_84 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_85 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_86 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_62 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_95 /></TD>"
									sRowContents = sRowContents & "<TD>0</TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_116 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_117 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_118 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_23 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_51 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_-2 /></TD>"
									sRowContents = sRowContents & "<TD><CONCEPT_0 /></TD>"
									sRowContents = sRowContents & "<TD></TD>"
								sRowContents = sRowContents & "</TR>"
								lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
							End If
							sRowContents = Replace(sRowContents, "<CONCEPT_" & CStr(oRecordset.Fields("ConceptID").Value) & " />", CDbl(oRecordset.Fields("ConceptAmount").Value))
						    oRecordset.MoveNext
							If Err.number <> 0 Then Exit Do
						Loop
						oRecordset.Close
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						lErrorNumber = AppendTextToFile(sDocumentName, "</TABLE>", sErrorDescription)
						lErrorNumber = AppendTextToFile(sFullFile, sRowContents, sErrorDescription)
						lErrorNumber = AppendTextToFile(sFullFile, "</TABLE>", sErrorDescription)
					End If
				End If

				lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudieron guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
				If lErrorNumber = 0 Then
					lErrorNumber = DeleteFolder(sFilePath, sErrorDescription)
				End If
				oEndDate = Now()
				If (lErrorNumber = 0) And B_USE_SMTP Then
					If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
				End If
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
				oZonesRecordset.Close
			End If
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1311 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1334(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte personal médico en contacto con el paciente UNIMED 03
'         Carpeta 1. DOCUMENTACIÓN ENTREGADA POR JEFATURA DE SERVICIOS DE DESARROLLO HUMANO
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1334"
	Dim sHeaderContents
	Dim oRecordset
	Dim oAreasRecordset
	Dim oMedicalAreasRecordset
	Dim oZonesRecordset
	Dim oRecordset1
	Dim oRecordset2
	Dim sRowContents
	Dim lErrorNumber
	Dim asTotalColumn
	Dim iIndex
	Dim iBeginColumn
	Dim iEndColumn
	Dim lZoneID
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim sSourceFolderPath
	Dim lPayrollID
	Dim lForPayrollID
	Dim sCondition
	Dim lTotalCenter
	Dim lAreaID
	Dim sServiceIDs
	Dim sPositionIDs
	Dim iURPrevious
	Dim iURActual
	Dim iColumnNumber
	Dim sColumnNumber
	Dim asColumnNumber
	Dim asTotalColumnNumber
	Dim sStateName
	Dim iMedicalAreasTypeID
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim sBoldBegin
	Dim sBoldEnd
	Dim lSubTotal
	Dim lTotal1
	Dim lTotal2
	Dim lTotal3
	
	sColumnNumber = ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
	asColumnNumber = Split(sColumnNumber, ",", -1, vbBinaryCompare)
	asTotalColumnNumber = Split(sColumnNumber, ",", -1, vbBinaryCompare)
	sBoldBegin = "<B>"
	sBoldEnd = "</B>"
	lSubTotal = 0

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)

	oStartDate = Now()
	If (InStr(1, aReportTitle(L_MEDICAL_AREAS_TYPES_FLAGS), "03", vbBinaryCompare) > 0) Then
		iMedicalAreasTypeID = 3
		iBeginColumn = 1
		iEndColumn = 11
	Else
		iMedicalAreasTypeID = 4
		iBeginColumn = 12
		iEndColumn = 20
	End If
	lTotalCenters = 0

	If iMedicalAreasTypeID = 3 Then
		sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1334.htm"), sErrorDescription)
		If (Len(sHeaderContents) > 0) Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
			sFilePath = sFilePath & "\"
			sSourceFolderPath  = Server.MapPath(TEMPLATES_PATH & "Images")
			sSourceFolderPath = sSourceFolderPath & "\"
			sErrorDescription "Error al copiar el logo a la carpeta destino"
			'lErrorNumber = CopyFolder(sSourceFolderPath, sFilePath, sErrorDescription)
			If lErrorNumber = 0 Then
				sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
				sDocumentName = sFilePath & "Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"
				sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME, 1, -1, vbBinaryCompare)
				sHeaderContents = Replace(sHeaderContents, "<PAYROLL_DATE />", DisplayNumericDateFromSerialNumber(lForPayrollID))
				sHeaderContents = Replace(sHeaderContents, "<MEDICAL_AREAS_TYPE_NAME />", aReportTitle(L_MEDICAL_AREAS_TYPES_FLAGS))
				lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
				sRowContents = "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
					sRowContents = sRowContents & "<TR BGCOLOR=""#99CCFF"">"
						sRowContents = sRowContents & "<TD>UR</TD>"
						sRowContents = sRowContents & "<TD>CT</TD>"
						sRowContents = sRowContents & "<TD>AUX</TD>"
						sRowContents = sRowContents & "<TD>DENOMINACIÓN</TD>"
						sRowContents = sRowContents & "<TD>POBLACIÓN</TD>"
						sRowContents = sRowContents & "<TD>MÉDICO GENERAL O FAMILIAR</TD>"
						sRowContents = sRowContents & "<TD>GINECO-OBSTETRIA</TD>"
						sRowContents = sRowContents & "<TD>PEDIATRA</TD>"
						sRowContents = sRowContents & "<TD>ODONTOLOGO</TD>"
						sRowContents = sRowContents & "<TD>CIRUJANO</TD>"
						sRowContents = sRowContents & "<TD>INTERNISTA</TD>"
						sRowContents = sRowContents & "<TD>OTRAS ESPECIALIDADES</TD>"
						sRowContents = sRowContents & "<TD>OTRAS LABORES</TD>"
						sRowContents = sRowContents & "<TD>RESIDENTES</TD>"
						sRowContents = sRowContents & "<TD>INTERNOS</TD>"
						sRowContents = sRowContents & "<TD>PASANTES</TD>"
						sRowContents = sRowContents & "<TD>TOTAL</TD>"
					sRowContents = sRowContents & "</TR>"
				sRowContents = sRowContents & "</TABLE>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
				Response.Flush()
			End If
		End If
		sErrorDescription = "No se pudieron obtener las áreas Generadoras."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaID, AreaTypeID, AreaLevelTypeID, CenterTypeID, AttentionLevelID, AreaCode, AreaShortName, AreaName, URCTAUX, Zones.ZoneName, States.ZoneName As StateName From Areas, Zones, Zones As ParentZones, Zones As States Where (AreaID > -1) And (Areas.EndDate=30000000) And (Areas.ZoneID > 0) And (AreaTypeID = 2) And (CenterTypeID < 10) And (Areas.ZoneID = Zones.ZoneID) And (ParentZones.ZoneID = Zones.ParentID) And (States.ZoneID = ParentZones.ParentID) " & sCondition & " Order By States.ZoneID, URCTAUX", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oAreasRecordset)
		If lErrorNumber = 0 Then
			If Not oAreasRecordset.EOF Then
				iURPrevious = CInt(Left(CleanStringForHTML(CStr(oAreasRecordset.Fields("URCTAUX").Value)),Len("000")))
				Do While Not oAreasRecordset.EOF
					lTotalCenter = 0
					lAreaID = CLng(oAreasRecordset.Fields("AreaID").Value)
					sStateName = CleanStringForHTML(CStr(oAreasRecordset.Fields("StateName").Value))
					sRowContents = "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0""><TR>"
						sRowContents = sRowContents & "<TD>" & "=T(""" & Left(CleanStringForHTML(CStr(oAreasRecordset.Fields("URCTAUX").Value)),Len("000")) & """)" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "=T(""" & Mid(CleanStringForHTML(CStr(oAreasRecordset.Fields("URCTAUX").Value)),Len("00000"),Len("000")) & """)" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "=T(""" & Right(CleanStringForHTML(CStr(oAreasRecordset.Fields("URCTAUX").Value)),Len("00")) & """)" & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oAreasRecordset.Fields("AreaName").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oAreasRecordset.Fields("ZoneName").Value)) & "</TD>"
						For iIndex = iBeginColumn To iEndColumn
							asColumnNumber(iIndex) = 0
						Next
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Areas.AreaID, ColumnNumber, Count(*) As Total From Areas, EmployeesHistoryListForPayroll, MedicalAreas Where (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.AreaID=" & lAreaID & ") And (EmployeesHistoryListForPayroll.PositionID=MedicalAreas.PositionID) And (EmployeesHistoryListForPayroll.ServiceID=MedicalAreas.ServiceID) And (MedicalAreas.ServiceID<>-1) And (EmployeesHistoryListForPayroll.ServiceID<>-1) And (EmployeesHistoryListForPayroll.PositionID<>-1) And (MedicalAreas.PositionID>0) And (MedicalAreas.MedicalAreasTypeID=" & iMedicalAreasTypeID & ") Group By Areas.AreaID, ColumnNumber", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset1)
						Do While Not oRecordset1.EOF
							iColumnNumber = CLng(oRecordset1.Fields("ColumnNumber").Value)
							asColumnNumber(iColumnNumber) = asColumnNumber(iColumnNumber) + CLng(oRecordset1.Fields("Total").Value)
							oRecordset1.MoveNext
							If Err.number <> 0 Then Exit Do
						Loop
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Areas.AreaID, ColumnNumber, Count(*) As Total From Areas, EmployeesHistoryListForPayroll, MedicalAreas Where (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.AreaID=" & lAreaID & ") And (EmployeesHistoryListForPayroll.PositionID=MedicalAreas.PositionID) And (MedicalAreas.ServiceID=-1) And (EmployeesHistoryListForPayroll.ServiceID<>-1) And (EmployeesHistoryListForPayroll.PositionID<>-1) And (MedicalAreas.PositionID>0) And (MedicalAreas.MedicalAreasTypeID=" & iMedicalAreasTypeID & ") Group By Areas.AreaID, ColumnNumber", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset2)
						Do While Not oRecordset2.EOF
							iColumnNumber = CLng(oRecordset2.Fields("ColumnNumber").Value)
							asColumnNumber(iColumnNumber) = asColumnNumber(iColumnNumber) + CLng(oRecordset2.Fields("Total").Value)
							oRecordset2.MoveNext
							If Err.number <> 0 Then Exit Do
						Loop
						For iIndex = iBeginColumn To iEndColumn
							sRowContents = sRowContents & "<TD>" & asColumnNumber(iIndex) & "</TD>"
							lTotalCenter = lTotalCenter + CInt(asColumnNumber(iIndex))
							asTotalColumnNumber(iIndex) = asTotalColumnNumber(iIndex) + CLng(asColumnNumber(iIndex))
						Next
						sRowContents = sRowContents & "<TD>" & lTotalCenter & "</TD>"
					sRowContents = sRowContents & "</TR>"
					If lTotalCenter<>0 Then
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					End If
					oAreasRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
					iURActual = CInt(Left(CleanStringForHTML(CStr(oAreasRecordset.Fields("URCTAUX").Value)),Len("000")))
					If iURPrevious <> iURActual Then
						sRowContents = "</TABLE>"
						sRowContents = sRowContents & "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
							sRowContents = sRowContents & "<TR>"
								sRowContents = sRowContents & "<TD COLSPAN=""5"">" & sBoldBegin & sStateName & sBoldEnd & "</TD>"
								lTotalCenter = 0
								For iIndex = iBeginColumn To iEndColumn
									sRowContents = sRowContents & "<TD>" & asTotalColumnNumber(iIndex) & "</TD>"
									lTotalCenter = lTotalCenter + CDbl(asTotalColumnNumber(iIndex))
									asTotalColumnNumber(iIndex) = asTotalColumnNumber(iIndex) + CLng(asTotalColumnNumber(iIndex))
								Next
								For iIndex = iBeginColumn To iEndColumn
									asTotalColumnNumber(iIndex) = 0
								Next
								sRowContents = sRowContents & "<TD>" & lTotalCenter & "</TD>"
								sRowContents = sRowContents & "</TR>"
								sRowContents = sRowContents & "<TR>"
								For iIndex = iBeginColumn To iEndColumn + 5
									sRowContents = sRowContents & "<TD></TD>"
								Next
							sRowContents = sRowContents & "</TR>"
						sRowContents = sRowContents & "</TABLE>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						iURPrevious = CInt(Left(CleanStringForHTML(CStr(oAreasRecordset.Fields("URCTAUX").Value)),Len("000")))
					End If
				Loop
				sRowContents = "</TABLE>"
				sRowContents = sRowContents & "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
					sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD COLSPAN=""5"">" & sBoldBegin & sStateName & sBoldEnd & "</TD>"
						lTotalCenter = 0
						For iIndex = iBeginColumn To iEndColumn
							sRowContents = sRowContents & "<TD>" & sBoldBegin & asTotalColumnNumber(iIndex) & sBoldEnd & "</TD>"
							lTotalCenter = lTotalCenter + CDbl(asTotalColumnNumber(iIndex))
							asTotalColumnNumber(iIndex) = asTotalColumnNumber(iIndex) + CLng(asTotalColumnNumber(iIndex))
						Next
						For iIndex = iBeginColumn To iEndColumn
							asTotalColumnNumber(iIndex) = 0
						Next
						sRowContents = sRowContents & "<TD>" & sBoldBegin & lTotalCenter & sBoldEnd & "</TD>"
						sRowContents = sRowContents & "</TR>"
						sRowContents = sRowContents & "<TR>"
						For iIndex = iBeginColumn To iEndColumn + 5
							sRowContents = sRowContents & "<TD></TD>"
						Next
					sRowContents = sRowContents & "</TR>"
				sRowContents = sRowContents & "</TABLE>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				oAreasRecordset.Close
				lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudieron guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
				If lErrorNumber = 0 Then
					lErrorNumber = DeleteFolder(sFilePath, sErrorDescription)
				End If
				oEndDate = Now()
				If (lErrorNumber = 0) And B_USE_SMTP Then
					If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
				End If
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
				oZonesRecordset.Close
			End If
		End If
	Else
		sHeaderContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1334.htm"), sErrorDescription)
		If (Len(sHeaderContents) > 0) Then
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
			sFilePath = sFilePath & "\"
			sSourceFolderPath  = Server.MapPath(TEMPLATES_PATH & "Images")
			sSourceFolderPath = sSourceFolderPath & "\"
			sErrorDescription "Error al copiar el logo a la carpeta destino"
			'lErrorNumber = CopyFolder(sSourceFolderPath, sFilePath, sErrorDescription)
			If lErrorNumber = 0 Then
				sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
				sDocumentName = sFilePath & "Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"
				sHeaderContents = Replace(sHeaderContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME, 1, -1, vbBinaryCompare)
				sHeaderContents = Replace(sHeaderContents, "<PAYROLL_DATE />", DisplayNumericDateFromSerialNumber(lForPayrollID))
				sHeaderContents = Replace(sHeaderContents, "<MEDICAL_AREAS_TYPE_NAME />", aReportTitle(L_MEDICAL_AREAS_TYPES_FLAGS))
				lErrorNumber = SaveTextToFile(sDocumentName, sHeaderContents, sErrorDescription)
				sRowContents = "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
				sRowContents = sRowContents & "<TR BGCOLOR=""#99CCFF"">"
				sRowContents = sRowContents & "<TD>UR</TD><TD>CT</TD><TD>AUX</TD><TD>DENOMINACIÓN</TD><TD>POBLACIÓN</TD><TD>GENERAL</TD><TD>ESPECIALISTA</TD><TD>AUXILIAR</TD><TD>PASANTE</TD><TD>TOTAL</TD><TD>LABORATORISTAS</TD><TD>RAYOS X</TD><TD>OTROS</TD><TD>TOTAL</TD><TD>ADMVO.</TD><TD>SERVS.GRALES.</TD><TD>TOTAL</TD>"
				sRowContents = sRowContents & "</TR>"
				sRowContents = sRowContents & "</TABLE>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
				Response.Flush()
			End If
		End If
		sErrorDescription = "No se pudieron obtener las áreas Generadoras."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaID, AreaTypeID, AreaLevelTypeID, CenterTypeID, AttentionLevelID, AreaCode, AreaShortName, AreaName, URCTAUX, Zones.ZoneName, States.ZoneName As StateName From Areas, Zones, Zones As ParentZones, Zones As States Where (AreaID > -1) And (Areas.EndDate=30000000) And (Areas.ZoneID > 0) And (AreaTypeID = 2) And (CenterTypeID < 10) And (Areas.ZoneID = Zones.ZoneID) And (ParentZones.ZoneID = Zones.ParentID) And (States.ZoneID = ParentZones.ParentID) " & sCondition & " Order By States.ZoneID, URCTAUX", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oAreasRecordset)
		If lErrorNumber = 0 Then
			If Not oAreasRecordset.EOF Then
				iURPrevious = CInt(Left(CleanStringForHTML(CStr(oAreasRecordset.Fields("URCTAUX").Value)),Len("000")))
				lTotal1 = 0
				lTotal2 = 0
				lTotal3 = 0
				Do While Not oAreasRecordset.EOF
					lTotalCenter = 0
					lSubTotal = 0
					lAreaID = CLng(oAreasRecordset.Fields("AreaID").Value)
					sStateName = CleanStringForHTML(CStr(oAreasRecordset.Fields("StateName").Value))
					sRowContents = "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0""><TR>"
					sRowContents = sRowContents & "<TD>" & "=T(""" & Left(CleanStringForHTML(CStr(oAreasRecordset.Fields("URCTAUX").Value)),Len("000")) & """)" & "</TD>"
					sRowContents = sRowContents & "<TD>" & "=T(""" & Mid(CleanStringForHTML(CStr(oAreasRecordset.Fields("URCTAUX").Value)),Len("00000"),Len("000")) & """)" & "</TD>"
					sRowContents = sRowContents & "<TD>" & "=T(""" & Right(CleanStringForHTML(CStr(oAreasRecordset.Fields("URCTAUX").Value)),Len("00")) & """)" & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oAreasRecordset.Fields("AreaName").Value)) & "</TD>"
					sRowContents = sRowContents & "<TD>" & CleanStringForHTML(CStr(oAreasRecordset.Fields("ZoneName").Value)) & "</TD>"
					For iIndex = iBeginColumn To iEndColumn
						asColumnNumber(iIndex) = 0
					Next
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Areas.AreaID, ColumnNumber, Count(*) As Total From Areas, EmployeesHistoryListForPayroll, MedicalAreas Where (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.AreaID=" & lAreaID & ") And (EmployeesHistoryListForPayroll.PositionID=MedicalAreas.PositionID) And (EmployeesHistoryListForPayroll.ServiceID=MedicalAreas.ServiceID) And (MedicalAreas.ServiceID<>-1) And (EmployeesHistoryListForPayroll.ServiceID<>-1) And (EmployeesHistoryListForPayroll.PositionID<>-1) And (MedicalAreas.PositionID>0) And (MedicalAreas.MedicalAreasTypeID=" & iMedicalAreasTypeID & ") Group By Areas.AreaID, ColumnNumber", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset1)
					Do While Not oRecordset1.EOF
						iColumnNumber = CLng(oRecordset1.Fields("ColumnNumber").Value)
						asColumnNumber(iColumnNumber) = asColumnNumber(iColumnNumber) + CLng(oRecordset1.Fields("Total").Value)
						oRecordset1.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Areas.AreaID, ColumnNumber, Count(*) As Total From Areas, EmployeesHistoryListForPayroll, MedicalAreas Where (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.AreaID=" & lAreaID & ") And (EmployeesHistoryListForPayroll.PositionID=MedicalAreas.PositionID) And (MedicalAreas.ServiceID=-1) And (EmployeesHistoryListForPayroll.ServiceID<>-1) And (EmployeesHistoryListForPayroll.PositionID<>-1) And (MedicalAreas.PositionID>0) And (MedicalAreas.MedicalAreasTypeID=" & iMedicalAreasTypeID & ") Group By Areas.AreaID, ColumnNumber", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset2)
					Do While Not oRecordset2.EOF
						iColumnNumber = CLng(oRecordset2.Fields("ColumnNumber").Value)
						asColumnNumber(iColumnNumber) = asColumnNumber(iColumnNumber) + CLng(oRecordset2.Fields("Total").Value)
						oRecordset2.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					For iIndex = iBeginColumn To iEndColumn
						sRowContents = sRowContents & "<TD>" & asColumnNumber(iIndex) & "</TD>"
						lTotalCenter = lTotalCenter + CInt(asColumnNumber(iIndex))
						lSubTotal = lSubTotal + CInt(asColumnNumber(iIndex))
						If iIndex = 15 Or iIndex = 18 Or iIndex = 20 Then
							sRowContents = sRowContents & "<TD>" & lSubTotal & "</TD>"
							lSubTotal = 0
						End If
						asTotalColumnNumber(iIndex) = asTotalColumnNumber(iIndex) + CLng(asColumnNumber(iIndex))
					Next
					'sRowContents = sRowContents & "<TD>" & lTotalCenter & "</TD>"
					sRowContents = sRowContents & "</TR>"
					If lTotalCenter<>0 Then
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
					End If
					oAreasRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
					iURActual = CInt(Left(CleanStringForHTML(CStr(oAreasRecordset.Fields("URCTAUX").Value)),Len("000")))
					If iURPrevious <> iURActual Then
						sRowContents = "</TABLE>"
						sRowContents = sRowContents & "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
						sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD COLSPAN=""5"">" & sBoldBegin & sStateName & sBoldEnd & "</TD>"
						lTotalCenter = 0
						lSubTotal = 0
						For iIndex = iBeginColumn To iEndColumn
							sRowContents = sRowContents & "<TD>" & sBoldBegin & asTotalColumnNumber(iIndex) & sBoldEnd & "</TD>"
							lTotalCenter = lTotalCenter + CLng(asTotalColumnNumber(iIndex))
							asTotalColumnNumber(iIndex) = asTotalColumnNumber(iIndex) + CLng(asTotalColumnNumber(iIndex))
							lSubTotal = lSubTotal + asTotalColumnNumber(iIndex)
							If iIndex = 15 Or iIndex = 18 Or iIndex = 20 Then
								sRowContents = sRowContents & "<TD>" & sBoldBegin & lSubTotal/2 & sBoldEnd & "</TD>"
								lSubTotal = 0
							End If
						Next
						For iIndex = iBeginColumn To iEndColumn
							asTotalColumnNumber(iIndex) = 0
						Next
						sRowContents = sRowContents & "<TD>" & lTotalCenter & "</TD>"
						sRowContents = sRowContents & "</TR>"
						sRowContents = sRowContents & "<TR>"
						For iIndex = iBeginColumn To iEndColumn + 5
							sRowContents = sRowContents & "<TD></TD>"
						Next
						sRowContents = sRowContents & "</TR>"
						sRowContents = sRowContents & "</TABLE>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						iURPrevious = CInt(Left(CleanStringForHTML(CStr(oAreasRecordset.Fields("URCTAUX").Value)),Len("000")))
					End If
				Loop
				sRowContents = "</TABLE>"
				sRowContents = sRowContents & "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
				sRowContents = sRowContents & "<TR>"
				sRowContents = sRowContents & "<TD COLSPAN=""5"">" & sBoldBegin & sStateName & sBoldEnd & "</TD>"
				lTotalCenter = 0
				lSubTotal = 0
				For iIndex = iBeginColumn To iEndColumn
					sRowContents = sRowContents & "<TD>" & sBoldBegin & asTotalColumnNumber(iIndex) & sBoldEnd & "</TD>"
					lTotalCenter = lTotalCenter + CDbl(asTotalColumnNumber(iIndex))
					asTotalColumnNumber(iIndex) = asTotalColumnNumber(iIndex) + CLng(asTotalColumnNumber(iIndex))
					lSubTotal = lSubTotal + asTotalColumnNumber(iIndex)
					If iIndex = 15 Or iIndex = 18 Or iIndex = 20 Then
						sRowContents = sRowContents & "<TD>" & sBoldBegin & lSubTotal/2 & sBoldEnd & "</TD>"
						lSubTotal = 0
					End If
				Next
				sRowContents = sRowContents & "</TR>"
				sRowContents = sRowContents & "<TR>"
				For iIndex = iBeginColumn To iEndColumn + 5
					sRowContents = sRowContents & "<TD></TD>"
				Next
				sRowContents = sRowContents & "</TR>"
				sRowContents = sRowContents & "</TABLE>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				oAreasRecordset.Close
				lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudieron guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
				If lErrorNumber = 0 Then
					lErrorNumber = DeleteFolder(sFilePath, sErrorDescription)
				End If
				oEndDate = Now()
				If (lErrorNumber = 0) And B_USE_SMTP Then
					If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
				End If
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
				oZonesRecordset.Close
			End If
		End If
	End If	

	Set oRecordset = Nothing
	BuildReport1334 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1335(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Reporte de tabuladores
'         Carpeta 1. DOCUMENTACIÓN ENTREGADA POR JEFATURA DE SERVICIOS DE DESARROLLO HUMANO
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1335"
	Dim oRecordset
	Dim oEmployeeTypesRecordset
	Dim sCondition
	Dim asConditions
	Dim asTitles
	Dim iIndex
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim iEmployeeTypeID
	Dim lStatusID
	lStatusID = aReportTitle(L_CONCEPTS_VALUES_STATUS_FLAGS)

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	If (InStr(1, sCondition, "EmployeeTypes.", vbBinaryCompare) > 0) Then
		sCondition = Replace(sCondition, "EmployeeTypes.", "")
	End If

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeTypeID, EmployeeTypeShortName, EmployeeTypeName From EmployeeTypes Where (Active = 1)" & sCondition & " Order By EmployeeTypeID", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oEmployeeTypesRecordset)
	If lErrorNumber = 0 Then
		If Not oEmployeeTypesRecordset.EOF Then
			Call GetNameFromTable(oADODBConnection, "StatusConceptsValues", lStatusID, "", "", sNames, sErrorDescription)
			Response.Write "<B>TABULADORES CON ESTATUS " & UCase(sNames) & "</B>"
			Response.Write "<BR /><BR />"
			Do While Not oEmployeeTypesRecordset.EOF
				iEmployeeTypeID = CInt(oEmployeeTypesRecordset.Fields("EmployeeTypeID").Value)
				Select Case iEmployeeTypeID
					Case 0 'M. Médica, paramédica y grupos afines
						sCondition = "(ConceptsValues1.PositionID=Positions.PositionID) And (Positions.EmployeeTypeID=0)"
						sCondition = sCondition & " And (ConceptsValues1.PositionTypeID=PositionTypes.PositionTypeID)"
						sCondition = sCondition & " And (ConceptsValues1.StatusID=" & lStatusID & ") And (ConceptsValues1.EmployeeTypeID=0) And (ConceptsValues1.EconomicZoneID=2) And (ConceptsValues1.ConceptID=1)"
						sCondition = sCondition & " And (ConceptsValues2.StatusID=" & lStatusID & ") And (ConceptsValues2.EmployeeTypeID=0) And (ConceptsValues2.EconomicZoneID=2) And (ConceptsValues2.ConceptID=38)"
						sCondition = sCondition & " And (ConceptsValues3.StatusID=" & lStatusID & ") And (ConceptsValues3.EmployeeTypeID=0) And (ConceptsValues3.EconomicZoneID=2) And (ConceptsValues3.ConceptID=49)"
						sCondition = sCondition & " And (ConceptsValues4.StatusID=" & lStatusID & ") And (ConceptsValues4.EmployeeTypeID=0) And (ConceptsValues4.EconomicZoneID=3) And (ConceptsValues4.ConceptID=1)"
						sCondition = sCondition & " And (ConceptsValues5.StatusID=" & lStatusID & ") And (ConceptsValues5.EmployeeTypeID=0) And (ConceptsValues5.EconomicZoneID=3) And (ConceptsValues5.ConceptID=38)"
						sCondition = sCondition & " And (ConceptsValues6.StatusID=" & lStatusID & ") And (ConceptsValues6.EmployeeTypeID=0) And (ConceptsValues6.EconomicZoneID=3) And (ConceptsValues6.ConceptID=49)"
						sCondition = sCondition & " And (ConceptsValues1.WorkingHours=ConceptsValues2.WorkingHours) And (ConceptsValues1.LevelID=ConceptsValues2.LevelID) And (ConceptsValues1.PositionID=ConceptsValues2.PositionID)"
						sCondition = sCondition & " And (ConceptsValues1.WorkingHours=ConceptsValues3.WorkingHours) And (ConceptsValues1.LevelID=ConceptsValues3.LevelID) And (ConceptsValues1.PositionID=ConceptsValues3.PositionID)"
						sCondition = sCondition & " And (ConceptsValues1.WorkingHours=ConceptsValues4.WorkingHours) And (ConceptsValues1.LevelID=ConceptsValues4.LevelID) And (ConceptsValues1.PositionID=ConceptsValues4.PositionID)"
						sCondition = sCondition & " And (ConceptsValues1.WorkingHours=ConceptsValues5.WorkingHours) And (ConceptsValues1.LevelID=ConceptsValues5.LevelID) And (ConceptsValues1.PositionID=ConceptsValues5.PositionID)"
						sCondition = sCondition & " And (ConceptsValues1.WorkingHours=ConceptsValues6.WorkingHours) And (ConceptsValues1.LevelID=ConceptsValues6.LevelID) And (ConceptsValues1.PositionID=ConceptsValues6.PositionID)"
						sCondition = sCondition & " And (ConceptsValues2.WorkingHours=ConceptsValues3.WorkingHours) And (ConceptsValues2.LevelID=ConceptsValues3.LevelID) And (ConceptsValues2.PositionID=ConceptsValues3.PositionID)"
						sCondition = sCondition & " And (ConceptsValues2.WorkingHours=ConceptsValues4.WorkingHours) And (ConceptsValues2.LevelID=ConceptsValues4.LevelID) And (ConceptsValues2.PositionID=ConceptsValues4.PositionID)"
						sCondition = sCondition & " And (ConceptsValues2.WorkingHours=ConceptsValues5.WorkingHours) And (ConceptsValues2.LevelID=ConceptsValues5.LevelID) And (ConceptsValues2.PositionID=ConceptsValues5.PositionID)"
						sCondition = sCondition & " And (ConceptsValues2.WorkingHours=ConceptsValues6.WorkingHours) And (ConceptsValues2.LevelID=ConceptsValues6.LevelID) And (ConceptsValues2.PositionID=ConceptsValues6.PositionID)"
						sCondition = sCondition & " And (ConceptsValues3.WorkingHours=ConceptsValues4.WorkingHours) And (ConceptsValues3.LevelID=ConceptsValues4.LevelID) And (ConceptsValues3.PositionID=ConceptsValues4.PositionID)"
						sCondition = sCondition & " And (ConceptsValues3.WorkingHours=ConceptsValues5.WorkingHours) And (ConceptsValues3.LevelID=ConceptsValues5.LevelID) And (ConceptsValues3.PositionID=ConceptsValues5.PositionID)"
						sCondition = sCondition & " And (ConceptsValues3.WorkingHours=ConceptsValues6.WorkingHours) And (ConceptsValues3.LevelID=ConceptsValues6.LevelID) And (ConceptsValues3.PositionID=ConceptsValues6.PositionID)"
						sCondition = sCondition & " And (ConceptsValues4.WorkingHours=ConceptsValues5.WorkingHours) And (ConceptsValues4.LevelID=ConceptsValues5.LevelID) And (ConceptsValues4.PositionID=ConceptsValues5.PositionID)"
						sCondition = sCondition & " And (ConceptsValues4.WorkingHours=ConceptsValues6.WorkingHours) And (ConceptsValues4.LevelID=ConceptsValues6.LevelID) And (ConceptsValues4.PositionID=ConceptsValues6.PositionID)"
						sCondition = sCondition & " And (ConceptsValues5.WorkingHours=ConceptsValues6.WorkingHours) And (ConceptsValues5.LevelID=ConceptsValues6.LevelID) And (ConceptsValues5.PositionID=ConceptsValues6.PositionID)"
						asTitles = Split(",CÓDIGO,NIVEL,JORNADA,DENOMINACIÓN,SUELDO,ASIGNACIÓN MÉDICA,AYUDA PARA GASTOS DE ACTUALIZACIÓN,TOTAL,SUELDO,ASIGNACIÓN MÉDICA,AYUDA PARA GASTOS DE ACTUALIZACIÓN,TOTAL", ",")
						sErrorDescription = "No se pudieron obtener los tabuladores del puesto."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct ConceptsValues1.RecordID, PositionTypeShortName, Positions.PositionShortName, Positions.PositionName, ConceptsValues1.LevelID, ConceptsValues1.WorkingHours, ConceptsValues1.ConceptAmount*2 As Sueldo_Z2, ConceptsValues2.ConceptAmount*2 As AsignacionMedica_Z2, ConceptsValues3.ConceptAmount*2 As AyudaGastos_Z2, ConceptsValues4.ConceptAmount*2 As Sueldo_Z3, ConceptsValues5.ConceptAmount*2 As AsignacionMedica_Z3, ConceptsValues6.ConceptAmount*2 As AyudaGastos_Z3 From ConceptsValues As ConceptsValues1, ConceptsValues As ConceptsValues2, ConceptsValues As ConceptsValues3, ConceptsValues As ConceptsValues4, ConceptsValues As ConceptsValues5, ConceptsValues As ConceptsValues6, Positions, PositionTypes Where " & sCondition & " Order By PositionTypeShortName, Positions.PositionShortName, Positions.PositionName, ConceptsValues1.LevelID, ConceptsValues1.WorkingHours", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								Response.Write "<TABLE BORDER="""
								If Not bForExport Then
									Response.Write "0"
								Else
									Response.Write "1"
								End If
								Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
								asColumnsTitles = Split("<SPAN COLS=""5"">&nbsp;,<SPAN COLS=""4"">Zona 2,<SPAN COLS=""4"">Zona 3", ",", -1, vbBinaryCompare)
								asCellWidths = Split(",,,,,,,,,,,,", ",", -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
								Else
									If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
										lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									Else
										lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									End If
								End If

								asColumnsTitles = Split("Tipo puesto,Código,Nivel,Jornada,Denominación,Sueldo,Asignación médica,Gastos de actualización,Total,Sueldo,Asignación médica,Gastos de actualización,Total", ",", -1, vbBinaryCompare)
								asCellWidths = Split(",,,,,,,,,,,,", ",", -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
								Else
									If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
										lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									Else
										lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									End If
								End If
								asCellAlignments = Split("CENTER,RIGHT,RIGHT,RIGHT,LEFT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
								Do While Not oRecordset.EOF
									sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("PositionTypeShortName").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(Right(("000" & CStr(oRecordset.Fields("LevelID").Value)), Len("000"))) & "&nbsp;"
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("WorkingHours").Value)) & " Hrs."
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Sueldo_Z2").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("AsignacionMedica_Z2").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("AyudaGastos_Z2").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Sueldo_Z2").Value)+CDbl(oRecordset.Fields("AsignacionMedica_Z2").Value) + CDbl(oRecordset.Fields("AyudaGastos_Z2").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Sueldo_Z3").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("AsignacionMedica_Z3").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("AyudaGastos_Z3").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Sueldo_Z3").Value)+CDbl(oRecordset.Fields("AsignacionMedica_Z3").Value)+CDbl(oRecordset.Fields("AyudaGastos_Z3").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR
									asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If bForExport Then
										lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
									Else
										lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
									End If
									oRecordset.MoveNext
									If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
								Loop
								Response.Write "</TABLE><BR /><BR />"
							Else
								lErrorNumber = L_ERR_NO_RECORDS
								sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
							End If
						End If
					Case 1 'F. Funcionario
						sErrorDescription = "No se pudieron obtener los tabuladores del puesto."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct PositionShortName, PositionName, GroupGradeLevelShortName, ConceptsValues_1.ClassificationID, ConceptsValues_1.IntegrationID, ConceptsValues_1.StartDate, (ConceptsValues_1.ConceptAmount*2) As Sueldo, (ConceptsValues_3.ConceptAmount*2) As CompensacionGarantizada From ConceptsValues As ConceptsValues_1, ConceptsValues As ConceptsValues_3, Positions, GroupGradeLevels Where (ConceptsValues_1.GroupGradeLevelID=ConceptsValues_3.GroupGradeLevelID) And (ConceptsValues_1.ClassificationID=ConceptsValues_3.ClassificationID) And (ConceptsValues_1.IntegrationID=ConceptsValues_3.IntegrationID) And (ConceptsValues_1.GroupGradeLevelID=Positions.GroupGradeLevelID) And (ConceptsValues_1.ClassificationID=Positions.ClassificationID) And (ConceptsValues_1.IntegrationID=Positions.IntegrationID) And (ConceptsValues_3.GroupGradeLevelID=Positions.GroupGradeLevelID) And (ConceptsValues_3.ClassificationID=Positions.ClassificationID) And (ConceptsValues_3.IntegrationID=Positions.IntegrationID) And (Positions.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (ConceptsValues_1.ConceptID=1) And (ConceptsValues_3.ConceptID=3) And (Positions.EmployeeTypeID=1) And (ConceptsValues_1.StatusID=" & lStatusID & ") And (ConceptsValues_3.StatusID=" & lStatusID & ") Order By PositionShortName, PositionName, GroupGradeLevelShortName, ConceptsValues_1.ClassificationID, ConceptsValues_1.IntegrationID", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								Response.Write "<TABLE BORDER="""
								If Not bForExport Then
									Response.Write "0"
								Else
									Response.Write "1"
								End If
								Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
								asColumnsTitles = Split("Código,Denominación,Grupo-grado-nivel salarial,Clasificación,Integración,Sueldo base,Compensación garantizada,Sueldo integrado", ",", -1, vbBinaryCompare)
								asCellWidths = Split("100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
								Else
									If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
										lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									Else
										lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									End If
								End If
								asCellAlignments = Split(",,,,,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
								Do While Not oRecordset.EOF
									sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CStr(oRecordset.Fields("GroupGradeLevelShortName").Value)
									sRowContents = sRowContents & TABLE_SEPARATOR & CStr(oRecordset.Fields("ClassificationID").Value)
									sRowContents = sRowContents & TABLE_SEPARATOR & CStr(oRecordset.Fields("IntegrationID").Value)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Sueldo").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("CompensacionGarantizada").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Sueldo").Value) + CDbl(oRecordset.Fields("CompensacionGarantizada").Value), 2, True, False, True)
									asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If bForExport Then
										lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
									Else
										lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
									End If
									oRecordset.MoveNext
									If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
								Loop
								Response.Write "</TABLE><BR /><BR />"
							Else
								lErrorNumber = L_ERR_NO_RECORDS
								sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
							End If
						End If
					Case 2 'O. Operativo
						sErrorDescription = "No se pudieron obtener los tabuladores del puesto."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct PositionShortName, PositionName, PositionTypeName, ConceptsValues_1Z2.LevelID, ConceptsValues_1Z2.StartDate, (ConceptsValues_1Z2.ConceptAmount*2) As Sueldo_Z2, (ConceptsValues_3Z2.ConceptAmount*2) As Compensacion_Z2, (ConceptsValues_1Z3.ConceptAmount*2) As Sueldo_Z3, (ConceptsValues_3Z3.ConceptAmount*2) As Compensacion_Z3 From ConceptsValues As ConceptsValues_1Z2, ConceptsValues As ConceptsValues_3Z2, ConceptsValues As ConceptsValues_1Z3, ConceptsValues As ConceptsValues_3Z3, Positions, PositionTypes Where (ConceptsValues_1Z2.LevelID=Positions.LevelID) And (ConceptsValues_1Z2.PositionTypeID=Positions.PositionTypeID) And (ConceptsValues_3Z2.LevelID=Positions.LevelID) And (ConceptsValues_3Z2.PositionTypeID=Positions.PositionTypeID) And (ConceptsValues_1Z3.LevelID=Positions.LevelID) And (ConceptsValues_1Z3.PositionTypeID=Positions.PositionTypeID) And (ConceptsValues_3Z3.LevelID=Positions.LevelID) And (ConceptsValues_3Z3.PositionTypeID=Positions.PositionTypeID) And (ConceptsValues_1Z2.StartDate=ConceptsValues_3Z2.StartDate) And (ConceptsValues_3Z3.StartDate=ConceptsValues_3Z3.StartDate) And (Positions.PositionTypeID=PositionTypes.PositionTypeID) And (ConceptsValues_1Z2.ConceptID=1) And (ConceptsValues_3Z2.ConceptID=3) And (ConceptsValues_1Z3.ConceptID=1) And (ConceptsValues_3Z3.ConceptID=3) And (ConceptsValues_1Z2.EconomicZoneID=2) And (ConceptsValues_3Z2.EconomicZoneID=2) And (ConceptsValues_1Z3.EconomicZoneID=3) And (ConceptsValues_3Z3.EconomicZoneID=3) And (Positions.EmployeeTypeID=2) And (ConceptsValues_1Z2.StatusID=" & lStatusID & ") And (ConceptsValues_3Z2.StatusID=" & lStatusID & ") And (ConceptsValues_1Z3.StatusID=" & lStatusID & ") And (ConceptsValues_3Z3.StatusID=" & lStatusID & ") Order By PositionShortName, PositionName, PositionTypeName, ConceptsValues_1Z2.LevelID", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								Response.Write "<TABLE BORDER="""
								If Not bForExport Then
									Response.Write "0"
								Else
									Response.Write "1"
								End If
								Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
								asColumnsTitles = Split("Tipo de puesto,Código,Nivel,Denominación,Sueldo,Compensación garantizada,Total zona 2,Sueldo,Compensación garantizada,Total zona 3", ",", -1, vbBinaryCompare)
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
								asCellAlignments = Split("LEFT,CENTER,LEFT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
								Do While Not oRecordset.EOF
									sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("PositionTypeShortName").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(Right(("000" & CStr(oRecordset.Fields("LevelID").Value)), Len("000"))) & "&nbsp;"
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Sueldo_Z2").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Compensacion_Z2").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Sueldo_Z2").Value)+CDbl(oRecordset.Fields("Compensacion_Z2").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Sueldo_Z3").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Compensacion_Z3").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Sueldo_Z3").Value)+CDbl(oRecordset.Fields("Compensacion_Z3").Value), 2, True, False, True)
									asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If bForExport Then
										lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
									Else
										lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
									End If
									oRecordset.MoveNext
									If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
								Loop
								Response.Write "</TABLE><BR /><BR />"
							Else
								lErrorNumber = L_ERR_NO_RECORDS
								sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
							End If
						End If
					Case 3 'A. Alta responsabilidad
						sErrorDescription = "No se pudieron obtener los tabuladores del puesto."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct PositionShortName, PositionName, ConceptsValues_1.LevelID, ConceptsValues_1.StartDate, (ConceptsValues_1.ConceptAmount*2) As Sueldo, (ConceptsValues_3.ConceptAmount*2) As Compensacion From ConceptsValues As ConceptsValues_1, ConceptsValues As ConceptsValues_3, Positions Where (ConceptsValues_1.LevelID=Positions.LevelID) And (ConceptsValues_3.LevelID=Positions.LevelID) And (ConceptsValues_1.StartDate=ConceptsValues_3.StartDate) And (ConceptsValues_1.ConceptID=1) And (ConceptsValues_3.ConceptID=3) And (Positions.EmployeeTypeID=3) And (ConceptsValues_1.StatusID=" & lStatusID & ") And (ConceptsValues_3.StatusID=" & lStatusID & ") Order By PositionShortName, PositionName, ConceptsValues_1.LevelID", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								Response.Write "<TABLE BORDER="""
								If Not bForExport Then
									Response.Write "0"
								Else
									Response.Write "1"
								End If
								Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
								asColumnsTitles = Split("Código,Denominación,Nivel,Fecha inicio,Sueldo base,Compensación garantizada,Total mensual bruto", ",", -1, vbBinaryCompare)
								asCellWidths = Split(",,,,,,", ",", -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
								Else
									If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
										lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									Else
										lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									End If
								End If
								asCellAlignments = Split(",,,,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
								Do While Not oRecordset.EOF
									sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(Right(("000" & CStr(oRecordset.Fields("LevelID").Value)), Len("000"))) & "&nbsp;"
									sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CStr(oRecordset.Fields("StartDate").Value), -1, -1, -1)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Sueldo").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Compensacion").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Sueldo").Value) + CDbl(oRecordset.Fields("Compensacion").Value), 2, True, False, True)
									asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If bForExport Then
										lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
									Else
										lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
									End If
									oRecordset.MoveNext
									If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
								Loop
								Response.Write "</TABLE><BR /><BR />"
							Else
								lErrorNumber = L_ERR_NO_RECORDS
								sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
							End If
						End If
					Case 4 'E. Enlace
						sErrorDescription = "No se pudieron obtener los tabuladores del puesto."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct PositionShortName, PositionName, ConceptsValues_1Z2.LevelID, ConceptsValues_1Z2.StartDate, (ConceptsValues_1Z2.ConceptAmount*2) As Sueldo_Z2, (ConceptsValues_3Z2.ConceptAmount*2) As Compensacion_Z2, (ConceptsValues_1Z3.ConceptAmount*2) As Sueldo_Z3, (ConceptsValues_3Z3.ConceptAmount*2) As Compensacion_Z3 From ConceptsValues As ConceptsValues_1Z2, ConceptsValues As ConceptsValues_3Z2, ConceptsValues As ConceptsValues_1Z3, ConceptsValues As ConceptsValues_3Z3, Positions Where (ConceptsValues_1Z2.LevelID=Positions.LevelID) And (ConceptsValues_3Z2.LevelID=Positions.LevelID) And (ConceptsValues_1Z3.LevelID=Positions.LevelID) And (ConceptsValues_3Z3.LevelID=Positions.LevelID) And (ConceptsValues_1Z2.StartDate=ConceptsValues_3Z2.StartDate) And (ConceptsValues_3Z3.StartDate=ConceptsValues_3Z3.StartDate) And (ConceptsValues_1Z2.ConceptID=1) And (ConceptsValues_3Z2.ConceptID=3) And (ConceptsValues_1Z3.ConceptID=1) And (ConceptsValues_3Z3.ConceptID=3) And (ConceptsValues_1Z2.EconomicZoneID=2) And (ConceptsValues_3Z2.EconomicZoneID=2) And (ConceptsValues_1Z3.EconomicZoneID=3) And (ConceptsValues_3Z3.EconomicZoneID=3) And (Positions.EmployeeTypeID=4) And (ConceptsValues_1Z2.StatusID=" & lStatusID & ") And (ConceptsValues_3Z2.StatusID=" & lStatusID & ") And (ConceptsValues_1Z3.StatusID=" & lStatusID & ") And (ConceptsValues_3Z3.StatusID=" & lStatusID & ") Order By PositionShortName, PositionName, ConceptsValues_1Z2.LevelID", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								Response.Write "<TABLE BORDER="""
								If Not bForExport Then
									Response.Write "0"
								Else
									Response.Write "1"
								End If
								Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
								asColumnsTitles = Split("Código,Denominación,Nivel,Fecha inicio,Sueldo,Compensación garantizada,Total zona 2,Sueldo,Compensación garantizada,Total zona 3", ",", -1, vbBinaryCompare)
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
								asCellAlignments = Split(",,,,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
								Do While Not oRecordset.EOF
									sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(Right(("000" & CStr(oRecordset.Fields("LevelID").Value)), Len("000"))) & "&nbsp;"
									sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Sueldo_Z2").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Compensacion_Z2").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Sueldo_Z2").Value) + CDbl(oRecordset.Fields("Compensacion_Z2").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Sueldo_Z3").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Compensacion_Z3").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Sueldo_Z3").Value) + CDbl(oRecordset.Fields("Compensacion_Z3").Value), 2, True, False, True)
									asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If bForExport Then
										lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
									Else
										lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
									End If
									oRecordset.MoveNext
									If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
								Loop
								Response.Write "</TABLE><BR /><BR />"
							Else
								lErrorNumber = L_ERR_NO_RECORDS
								sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
							End If
						End If
					Case 5 'R. Residente
						sErrorDescription = "No se pudieron obtener los tabuladores del puesto."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct PositionShortName, PositionName, ConceptsValues_36Z2.LevelID, ConceptsValues_36Z2.StartDate, (ConceptsValues_36Z2.ConceptAmount*2) As Complemento_Z2, (ConceptsValues_B2Z2.ConceptAmount*2) As Beca_Z2, (ConceptsValues_36Z3.ConceptAmount*2) As Complemento_Z3, (ConceptsValues_B2Z3.ConceptAmount*2) As Beca_Z3 From ConceptsValues As ConceptsValues_36Z2, ConceptsValues As ConceptsValues_B2Z2, ConceptsValues As ConceptsValues_36Z3, ConceptsValues As ConceptsValues_B2Z3, Positions Where (ConceptsValues_36Z2.LevelID=Positions.LevelID) And (ConceptsValues_B2Z2.LevelID=Positions.LevelID) And (ConceptsValues_36Z3.LevelID=Positions.LevelID) And (ConceptsValues_B2Z3.LevelID=Positions.LevelID) And (ConceptsValues_36Z2.StartDate=ConceptsValues_B2Z2.StartDate) And (ConceptsValues_36Z3.StartDate=ConceptsValues_B2Z3.StartDate) And (ConceptsValues_36Z2.ConceptID=39) And (ConceptsValues_B2Z2.ConceptID=89) And (ConceptsValues_36Z3.ConceptID=39) And (ConceptsValues_B2Z3.ConceptID=89) And (ConceptsValues_36Z2.EconomicZoneID=2) And (ConceptsValues_B2Z2.EconomicZoneID=2) And (ConceptsValues_36Z3.EconomicZoneID=3) And (ConceptsValues_B2Z3.EconomicZoneID=3) And (Positions.EmployeeTypeID=5) And (ConceptsValues_36Z2.StatusID=" & lStatusID & ") And (ConceptsValues_B2Z2.StatusID=" & lStatusID & ") And (ConceptsValues_36Z3.StatusID=" & lStatusID & ") And (ConceptsValues_B2Z3.StatusID=" & lStatusID & ") Order By PositionShortName, PositionName, ConceptsValues_36Z2.LevelID", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								Response.Write "<TABLE BORDER="""
								If Not bForExport Then
									Response.Write "0"
								Else
									Response.Write "1"
								End If
								Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
								asColumnsTitles = Split("<SPAN COLS=""4"">&nbsp;,<SPAN COLS=""3"">Zona 2,<SPAN COLS=""3"">Zona 3", ",", -1, vbBinaryCompare)
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
									
								asColumnsTitles = Split("Código,Denominación,Nivel,Fecha inicio,Beca,Complementeo de beca,Total,Beca,Complementeo de beca,Total", ",", -1, vbBinaryCompare)
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
								asCellAlignments = Split("LEFT,CENTER,LEFT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
								Do While Not oRecordset.EOF
									sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(Right(("000" & CStr(oRecordset.Fields("LevelID").Value)), Len("000"))) & "&nbsp;"
									sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Beca_Z2").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Complemento_Z2").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Beca_Z2").Value) + CDbl(oRecordset.Fields("Complemento_Z2").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Beca_Z3").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Complemento_Z3").Value), 2, True, False, True)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Beca_Z3").Value) + CDbl(oRecordset.Fields("Complemento_Z3").Value), 2, True, False, True)
									asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If bForExport Then
										lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
									Else
										lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
									End If
									oRecordset.MoveNext
									If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
								Loop
								Response.Write "</TABLE><BR /><BR />"
							Else
								lErrorNumber = L_ERR_NO_RECORDS
								sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
							End If
						End If
					Case 6 'B. Becario
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct PositionShortName, PositionName, ConceptsValues_12.LevelID, ConceptsValues_12.EconomicZoneID, ConceptsValues_12.StartDate, (ConceptsValues_12.ConceptAmount*2) As Sueldo From ConceptsValues As ConceptsValues_12, Positions Where (ConceptsValues_12.LevelID=Positions.LevelID) And (ConceptsValues_12.ConceptID=14) And (Positions.EmployeeTypeID=6) And (ConceptsValues_12.StatusID=" & lStatusID & ") Order By PositionShortName, PositionName, ConceptsValues_12.LevelID, ConceptsValues_12.EconomicZoneID", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								Response.Write "<TABLE BORDER="""
								If Not bForExport Then
									Response.Write "0"
								Else
									Response.Write "1"
								End If
								Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
								asColumnsTitles = Split("Código,Denominación,Nivel,Fecha inicio,Beca", ",", -1, vbBinaryCompare)
								asCellWidths = Split(",,,,", ",", -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
								Else
									If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
										lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									Else
										lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									End If
								End If
								asCellAlignments = Split(",,,,RIGHT", ",", -1, vbBinaryCompare)
								Do While Not oRecordset.EOF
									sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(Right(("000" & CStr(oRecordset.Fields("LevelID").Value)), Len("000"))) & "&nbsp;"
									'sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EconomicZoneID").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
									sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Sueldo").Value), 2, True, False, True)
									asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
									If bForExport Then
										lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
									Else
										lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
									End If
									oRecordset.MoveNext
									If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
								Loop
								Response.Write "</TABLE><BR /><BR />"
							Else
								lErrorNumber = L_ERR_NO_RECORDS
								sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
							End If
						End If
				End Select
				oEmployeeTypesRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1335 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1336(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Reporte de tabuladores
'         Carpeta 1. DOCUMENTACIÓN ENTREGADA POR JEFATURA DE SERVICIOS DE DESARROLLO HUMANO
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1336"
	Dim oRecordset
	Dim sCondition
	Dim asConditions
	Dim asTitles
	Dim iIndex
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim iEmployeeTypeID

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)

	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaCode, AreaShortName, AreaName, URCTAUX, CompanyName, AreaTypeName, ConfineTypeName, AreaLevelTypeName, CenterTypeName, CenterSubTypeName, AttentionLevelName, AreaAddress, AreaCity, AreaZip, Zones.ZoneName, EconomicZoneName, CashierOfficeName, GeneratingAreaName, Areas.StartDate, Areas.ZoneID, Areas.EndDate, Areas.FinishDate, ParentZones.ZoneName As Municipio, States.ZoneName As StateName From Areas, AreaLevelTypes, AreaTypes, AttentionLevels, CashierOffices, CenterSubtypes, CenterTypes, Companies, ConfineTypes, EconomicZones, GeneratingAreas, Zones, Zones As ParentZones, Zones As States Where Areas.AreaLevelTypeID = AreaLevelTypes.AreaLevelTypeID And Areas.AreaTypeID = AreaTypes.AreaTypeID And Areas.AttentionLevelID = AttentionLevels.AttentionLevelID And Areas.CashierOfficeID = CashierOffices.CashierOfficeID And Areas.CenterSubtypeID = CenterSubtypes.CenterSubtypeID And Areas.CenterTypeID = CenterTypes.CenterTypeID And Areas.CompanyID = Companies.CompanyID And Areas.ConfineTypeID = ConfineTypes.ConfineTypeID And Areas.EconomicZoneID = EconomicZones.EconomicZoneID And Areas.GeneratingAreaID = GeneratingAreas.GeneratingAreaID And Areas.ZoneID = Zones.ZoneID And Areas.AreaLevelTypeID = 2 And Zones.ParentID = ParentZones.ZoneID And ParentZones.ParentID = States.ZoneID " & sCondition & " Order By AreaShortName, AreaName", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
				Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("Clave,Nombre,UR-CT-AUX,Empresa,Tipo de área,Ámbito del área,Tipo de centro de trabajo, Subtipo de centro de trabajo,Nivel de atención,Dirección,Ciudad,C.P.,Entidad,Municipio,Población,Zona económica,Área generadora,Pagaduría SIPE,Fecha de inicio,Fecha de término,Fecha inhabilitado", ",", -1, vbBinaryCompare)
				asCellWidths = Split(",,,,,,,,,,,,,,,,,,,,", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If
				asCellAlignments = Split(",,,,,,,,,,,,,,,,,,,,", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sRowContents = "&nbsp;" & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("URCTAUX").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CompanyName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ConfineTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CenterTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CenterSubTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AttentionLevelName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaAddress").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaCity").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaZip").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("StateName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Municipio").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ZoneName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EconomicZoneName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("GeneratingAreaName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CashierOfficeName").Value))
					If CLng(oRecordset.Fields("StartDate").Value) = 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "-"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value))
					End If
					If CLng(oRecordset.Fields("EndDate").Value)= 30000000 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "-"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value))
					End If
					If CLng(oRecordset.Fields("EndDate").Value)= 30000000 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "-"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("FinishDate").Value))
					End If
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				Response.Write "</TABLE><BR /><BR />"
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
			End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1336 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1337(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Reporte de tabuladores
'         Carpeta 1. DOCUMENTACIÓN ENTREGADA POR JEFATURA DE SERVICIOS DE DESARROLLO HUMANO
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1337"
	Dim oRecordset
	Dim sCondition
	Dim asConditions
	Dim asTitles
	Dim iIndex
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim iEmployeeTypeID

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)

	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaCode, AreaShortName, AreaName, URCTAUX, CompanyName, AreaTypeName, ConfineTypeName, AreaLevelTypeName, CenterTypeName, CenterSubTypeName, AttentionLevelName, AreaAddress, AreaCity, AreaZip, Zones.ZoneName, EconomicZoneName, CashierOfficeName, GeneratingAreaName, Areas.StartDate, Areas.ZoneID, Areas.EndDate, Areas.FinishDate, ParentZones.ZoneName As Municipio, States.ZoneName As StateName From Areas, AreaLevelTypes, AreaTypes, AttentionLevels, CashierOffices, CenterSubtypes, CenterTypes, Companies, ConfineTypes, EconomicZones, GeneratingAreas, Zones, Zones As ParentZones, Zones As States Where Areas.AreaLevelTypeID = AreaLevelTypes.AreaLevelTypeID And Areas.AreaTypeID = AreaTypes.AreaTypeID And Areas.AttentionLevelID = AttentionLevels.AttentionLevelID And Areas.CashierOfficeID = CashierOffices.CashierOfficeID And Areas.CenterSubtypeID = CenterSubtypes.CenterSubtypeID And Areas.CenterTypeID = CenterTypes.CenterTypeID And Areas.CompanyID = Companies.CompanyID And Areas.ConfineTypeID = ConfineTypes.ConfineTypeID And Areas.EconomicZoneID = EconomicZones.EconomicZoneID And Areas.GeneratingAreaID = GeneratingAreas.GeneratingAreaID And Areas.ZoneID = Zones.ZoneID And Areas.AreaLevelTypeID = 2 And Zones.ParentID = ParentZones.ZoneID And ParentZones.ParentID = States.ZoneID " & sCondition & " Order By AreaShortName, AreaName", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
				Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("Clave Centro de pago, Nombre centro de pago,Clave centro de trabajo,Nombre centro de trabajo,UR-CT-AUX,Empresa,Tipo de área,Ámbito del área,Tipo de centro de trabajo, Subtipo de centro de trabajo,Nivel de atención,Dirección,Ciudad,C.P.,Entidad,Municipio,Población,Zona económica,Área generadora,Pagaduría SIPE,Fecha de inicio,Fecha de término,Fecha inhabilitado", ",", -1, vbBinaryCompare)
				asCellWidths = Split(",,,,,,,,,,,,,,,,,,,,", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If
				asCellAlignments = Split(",,,,,,,,,,,,,,,,,,,,", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sRowContents = "&nbsp;" & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;" & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("URCTAUX").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CompanyName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ConfineTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CenterTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CenterSubTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AttentionLevelName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaAddress").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaCity").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaZip").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("StateName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Municipio").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ZoneName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EconomicZoneName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("GeneratingAreaName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CashierOfficeName").Value))
					If CLng(oRecordset.Fields("StartDate").Value) = 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "-"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value))
					End If
					If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "-"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value))
					End If
					If CLng(oRecordset.Fields("EndDate").Value)= 30000000 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "-"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("FinishDate").Value))
					End If
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				Response.Write "</TABLE><BR /><BR />"
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
			End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1337 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1339(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the SICAD file
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1339"
	Dim lType
    Dim sQueryBegin
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim lPayrollNumber
	Dim sDate
	Dim sFilePath
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim sTemp
	Dim sCurrentID
	Dim dTotal01
	Dim dTotalS7
	Dim adTotal
	Dim asTemp
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	lType = 1
	If Len(oRequest("LongReport").Item) > 0 Then
		lType = 2
	ElseIf Len(oRequest("Cancelled").Item) > 0 Then
		lType = 0
	End If
    sQueryBegin = ""
	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
    If InStr(1, sCondition, " And (EmployeesHistoryList.JobID=Jobs.JobID)", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Jobs"
    sCondition = Replace(sCondition, "EmployeesHistoryList", "EmployeesHistoryListForPayroll")
	sCondition = Replace(Replace(Replace(sCondition, "Banks.", "BankAccounts."), "Companies.", "EmployeesHistoryListForPayroll."), "EmployeeTypes.", "EmployeesHistoryListForPayroll.")
	lPayrollNumber = (CInt(Left(lForPayrollID, Len("0000"))) * 100) + CInt(GetPayrollNumber(lForPayrollID))
	oStartDate = Now()

	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
		sCondition = sCondition & " And ((EmployeesHistoryListForPayroll.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (EmployeesHistoryListForPayroll.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
	End If

	sErrorDescription = "No se pudieron obtener los montos pagados."
	If StrComp(oRequest("CheckConceptID").Item, "69", vbBinaryCompare) = 0 Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, CashierOfficeShortName, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeNumber, RFC, CURP, SocialSecurityNumber, BeneficiaryName As EmployeeName, BeneficiaryLastName As EmployeeLastName, BeneficiaryLastName2 As EmployeeLastName2, EmployeesHistoryListForPayroll.PositionTypeID, EmployeesHistoryListForPayroll.LevelID, ParentZones.ZoneCode, PaymentCenters.AreaShortName, AccountNumber, PositionShortName, Concepts.ConceptID, Concepts.ConceptShortName, Concepts.IsDeduction, Payroll_" & lPayrollID & ".ConceptAmount, Payrolls.PayrollTypeID, ForPayrollDate From EmployeesBeneficiariesLKP, Payrolls, Payroll_" & lPayrollID & ", Concepts, BankAccounts, Employees, EmployeesHistoryListForPayroll, Areas, Areas As PaymentCenters, CashierOffices, Zones, Zones As Zones2, Zones As ParentZones, Positions Where (Payrolls.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Areas.CashierOfficeID=CashierOffices.CashierOfficeID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") " & sCondition & " Order By EmployeesHistoryListForPayroll.EmployeeNumber, RecordDate, Concepts.IsDeduction, ConceptShortName", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, CashierOfficeShortName, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeNumber, RFC, CURP, SocialSecurityNumber, BeneficiaryName As EmployeeName, BeneficiaryLastName As EmployeeLastName, BeneficiaryLastName2 As EmployeeLastName2, EmployeesHistoryListForPayroll.PositionTypeID, EmployeesHistoryListForPayroll.LevelID, ParentZones.ZoneCode, PaymentCenters.AreaShortName, AccountNumber, PositionShortName, Concepts.ConceptID, Concepts.ConceptShortName, Concepts.IsDeduction, Payroll_" & lPayrollID & ".ConceptAmount, Payrolls.PayrollTypeID, ForPayrollDate From EmployeesBeneficiariesLKP, Payrolls, Payroll_" & lPayrollID & ", Concepts, BankAccounts, Employees, EmployeesHistoryListForPayroll, Areas, Areas As PaymentCenters, CashierOffices, Zones, Zones As Zones2, Zones As ParentZones, Positions Where (Payrolls.PayrollID=" & lPayrollID & ") And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Areas.CashierOfficeID=CashierOffices.CashierOfficeID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") " & sCondition & " Order By EmployeesHistoryListForPayroll.EmployeeNumber, RecordDate, Concepts.IsDeduction, ConceptShortName -->" & vbNewLine
	ElseIf StrComp(oRequest("CheckConceptID").Item, "155", vbBinaryCompare) = 0 Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesCreditorsLKP.CreditorNumber As EmployeeID, CashierOfficeShortName, EmployeesCreditorsLKP.CreditorNumber As EmployeeNumber, RFC, CURP, SocialSecurityNumber, CreditorName As EmployeeName, CreditorLastName As EmployeeLastName, CreditorLastName2 As EmployeeLastName2, EmployeesHistoryListForPayroll.PositionTypeID, EmployeesHistoryListForPayroll.LevelID, ParentZones.ZoneCode, PaymentCenters.AreaShortName, AccountNumber, PositionShortName, Concepts.ConceptID, Concepts.ConceptShortName, Concepts.IsDeduction, Payroll_" & lPayrollID & ".ConceptAmount, Payrolls.PayrollTypeID, ForPayrollDate From EmployeesCreditorsLKP, Payrolls, Payroll_" & lPayrollID & ", Concepts, BankAccounts, Employees, EmployeesHistoryListForPayroll, Areas, Areas As PaymentCenters, CashierOffices, Zones, Zones As Zones2, Zones As ParentZones, Positions Where (Payrolls.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Areas.CashierOfficeID=CashierOffices.CashierOfficeID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") " & sCondition & " Order By EmployeesHistoryListForPayroll.EmployeeNumber, RecordDate, Concepts.IsDeduction, ConceptShortName", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select EmployeesCreditorsLKP.CreditorNumber As EmployeeID, CashierOfficeShortName, EmployeesCreditorsLKP.CreditorNumber As EmployeeNumber, RFC, CURP, SocialSecurityNumber, CreditorName As EmployeeName, CreditorLastName As EmployeeLastName, CreditorLastName2 As EmployeeLastName2, EmployeesHistoryListForPayroll.PositionTypeID, EmployeesHistoryListForPayroll.LevelID, ParentZones.ZoneCode, PaymentCenters.AreaShortName, AccountNumber, PositionShortName, Concepts.ConceptID, Concepts.ConceptShortName, Concepts.IsDeduction, Payroll_" & lPayrollID & ".ConceptAmount, Payrolls.PayrollTypeID, ForPayrollDate From EmployeesCreditorsLKP, Payrolls, Payroll_" & lPayrollID & ", Concepts, BankAccounts, Employees, EmployeesHistoryListForPayroll, Areas, Areas As PaymentCenters, CashierOffices, Zones, Zones As Zones2, Zones As ParentZones, Positions Where (Payrolls.PayrollID=" & lPayrollID & ") And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Areas.CashierOfficeID=CashierOffices.CashierOfficeID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") " & sCondition & " Order By EmployeesHistoryListForPayroll.EmployeeNumber, RecordDate, Concepts.IsDeduction, ConceptShortName -->" & vbNewLine
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryListForPayroll.EmployeeID, CashierOfficeShortName, EmployeesHistoryListForPayroll.EmployeeNumber, RFC, CURP, SocialSecurityNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, EmployeesHistoryListForPayroll.PositionTypeID, EmployeesHistoryListForPayroll.LevelID, ParentZones.ZoneCode, PaymentCenters.AreaShortName, AccountNumber, PositionShortName, Concepts.ConceptID, Concepts.ConceptShortName, Concepts.IsDeduction, Payroll_" & lPayrollID & ".ConceptAmount, Payrolls.PayrollTypeID, ForPayrollDate From Payrolls, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, Areas, Areas As PaymentCenters, CashierOffices, Zones, Zones As Zones2, Zones As ParentZones, Positions" & sQueryBegin & " Where (Payrolls.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Areas.CashierOfficeID=CashierOffices.CashierOfficeID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID>0) " & Replace(sCondition, "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Order By EmployeesHistoryListForPayroll.EmployeeNumber, RecordDate, Concepts.IsDeduction, ConceptShortName", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryListForPayroll.EmployeeID, CashierOfficeShortName, EmployeesHistoryListForPayroll.EmployeeNumber, RFC, CURP, SocialSecurityNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, EmployeesHistoryListForPayroll.PositionTypeID, EmployeesHistoryListForPayroll.LevelID, ParentZones.ZoneCode, PaymentCenters.AreaShortName, AccountNumber, PositionShortName, Concepts.ConceptID, Concepts.ConceptShortName, Concepts.IsDeduction, Payroll_" & lPayrollID & ".ConceptAmount, Payrolls.PayrollTypeID, ForPayrollDate From Payrolls, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, Areas, Areas As PaymentCenters, CashierOffices, Zones, Zones As Zones2, Zones As ParentZones, Positions" & sQueryBegin & " Where (Payrolls.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Areas.CashierOfficeID=CashierOffices.CashierOfficeID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID>0) " & Replace(sCondition, "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Order By EmployeesHistoryListForPayroll.EmployeeNumber, RecordDate, Concepts.IsDeduction, ConceptShortName -->" & vbNewLine
	End If
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".txt"
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(Replace(sFilePath, ".txt", ".zip")) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			sFilePath = Server.MapPath(sFilePath)
			Response.Flush()

			sCurrentID = ""
			adTotal = Split("0,0,0,0,0,0,0,0,0", ",")
			dTotal01 = 0
			dTotalS7 = 0
			adTotal(0) = 0
			adTotal(1) = 0
			adTotal(2) = 0
			adTotal(3) = 0
			adTotal(4) = 0
			adTotal(5) = 0
			adTotal(6) = 0
			adTotal(7) = 0
			adTotal(8) = 0
			Do While Not oRecordset.EOF
				If StrComp(sCurrentID, CStr(oRecordset.Fields("EmployeeID").Value), vbBinaryCompare) <> 0 Then
					If Len(sCurrentID) > 0 Then
						If ((dTotal01 / 100) * 0.0275) <> dTotalS7 Then dTotal01 = FormatNumber((dTotalS7 / 0.0275), 2, True, False, False)
						Select Case lType
							Case 2
								sRowContents = Replace(sRowContents, "<CONCEPT_00 />", Right(("000000000000" & Int(adTotal(0) * 100)), Len("000000000000")))
								sRowContents = Replace(sRowContents, "<CONCEPT_01 />", Right(("000000000000" & (dTotal01 * 100)), Len("000000000000")))
								sRowContents = Replace(sRowContents, "<CONCEPT_60 />", "000000")
								sRowContents = Replace(sRowContents, "<CONCEPT_62 />", "0")
								sRowContents = Replace(sRowContents, "<CONCEPT_85 />", "000000")
							Case Else
								sRowContents = Replace(sRowContents, "<CONCEPT_00 />", Right(("000000000000" & Int(adTotal(0) * 100)), Len("000000000000")))
								sRowContents = Replace(sRowContents, "<CONCEPT_01 />", Right(("000000000000" & (dTotal01 * 100)), Len("000000000000")))
								sRowContents = Replace(sRowContents, "<CONCEPT_60 />", "0000000")
								sRowContents = Replace(sRowContents, "<CONCEPT_62 />", "0000000")
								sRowContents = Replace(sRowContents, "<CONCEPT_85 />", "0000000")
						End Select
						sRowContents = Replace(sRowContents, "<CONCEPT_69 />", "0000000")

						sRowContents = Replace(sRowContents, "<CONCEPT_S8 />", "0")
						sRowContents = Replace(sRowContents, "<CONCEPT_1S />", "0")
						sRowContents = Replace(sRowContents, "<CONCEPT_04 />", "0")
						sRowContents = Replace(sRowContents, "<CONCEPT_59 />", "0")
						sRowContents = Replace(sRowContents, "<CONCEPT_?? />", "0000")
						sRowContents = Replace(sRowContents, "<CONCEPT_51 />", Right(("000000" & Int(adTotal(3) * 100)), Len("000000")))
						sRowContents = Replace(sRowContents, "<CONCEPT_52 />", Right(("000000" & Int(adTotal(4) * 100)), Len("000000")))
						If adTotal(6) > 0 Then sRowContents = Replace(sRowContents, "<HAS_55 />", "07")
						If adTotal(7) > 0 Then sRowContents = Replace(sRowContents, "<HAS_55 />", "09")
						sRowContents = Replace(sRowContents, "<HAS_55 />", "00")
						sRowContents = Replace(sRowContents, "<CONCEPT_55 />", Right(("000000" & ((adTotal(6) + adTotal(7)) * 100)), Len("000000")))
						If adTotal(5) > 0 Then
							sRowContents = Replace(sRowContents, "<HAS_56 />", "06")
						Else
							sRowContents = Replace(sRowContents, "<HAS_56 />", "00")
						End If
						sRowContents = Replace(sRowContents, "<CONCEPT_56 />", Right(("000000" & (adTotal(5) * 100)), Len("000000")))

						sRowContents = Replace(sRowContents, "<SUMANDO />", Right(("0000000" & Int(adTotal(2) * 100)), Len("0000000")))
						sRowContents = Replace(sRowContents, "<TOTAL />", Right(("000000000000" & Int((adTotal(0) - adTotal(1) + adTotal(8)) * 100)), Len("000000")))
						lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
						dTotal01 = 0
						dTotalS7 = 0
						adTotal(0) = 0
						adTotal(1) = 0
						adTotal(2) = 0
						adTotal(3) = 0
						adTotal(4) = 0
						adTotal(5) = 0
						adTotal(6) = 0
						adTotal(7) = 0
					End If
					Select Case lType 'Clave del ramo
						Case 0 'Cancelaciones
							sRowContents = "00023"
						Case 2 'Largo
							sRowContents = "023"
						Case Else 'Corto
							sRowContents = "00023"
					End Select
					Select Case lType 'Clave de la pagaduría
						Case 0
							sRowContents = sRowContents & Right(("000000" & CStr(oRecordset.Fields("CashierOfficeShortName").Value)), Len("000000"))
						Case 2
							sRowContents = sRowContents & Right(("000000" & CStr(oRecordset.Fields("CashierOfficeShortName").Value)), Len("000000"))
						Case Else
							sRowContents = sRowContents & Right(("00000" & CStr(oRecordset.Fields("CashierOfficeShortName").Value)), Len("00000"))
					End Select
					sRowContents = sRowContents & "000000000" 'Right(("000000000" & CStr(oRecordset.Fields("EmployeeNumber").Value)), Len("000000000")) 'Número ISSSTE	'xxxxxxxxxxxxx
					sRowContents = sRowContents & Left((CStr(oRecordset.Fields("RFC").Value) & "0000000000000"), Len("0000000000000")) 'RFC
					Select Case lType
						Case 1
							sTemp = ""
							sTemp = CStr(oRecordset.Fields("CURP").Value)
							Err.Clear
							sRowContents = sRowContents & Left((sTemp & "000000000000000000"), Len("000000000000000000")) 'CURP
							sTemp = ""
							sTemp = CStr(oRecordset.Fields("SocialSecurityNumber").Value)
							Err.Clear
							sRowContents = sRowContents & Left((sTemp & "00000000000"), Len("00000000000")) 'No. Seguro Social
							sTemp = " "
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then sTemp = CStr(oRecordset.Fields("EmployeeLastName2").Value) & " "
							Err.Clear
							sRowContents = sRowContents & Left((CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & sTemp & CStr(oRecordset.Fields("EmployeeName").Value) & "                                        "), Len("                                        ")) 'Nombre del trabajador
							sRowContents = sRowContents & "<CONCEPT_01 />" 'Sueldo base
							If StrComp(CStr(oRecordset.Fields("AccountNumber").Value), ".", vbBinaryCompare) = 0 Then 'CLABE
								sRowContents = sRowContents & "                              "
							Else
								asTemp = Split(CStr(oRecordset.Fields("AccountNumber").Value), LIST_SEPARATOR)
								sRowContents = sRowContents & Left((asTemp(0) & "                              "), Len("                              "))
							End If
							Select Case CInt(oRecordset.Fields("PayrollTypeID").Value) 'Tipo de nómina
								Case 0
									sRowContents = sRowContents & "3"
								Case 2, 3, 4
									sRowContents = sRowContents & "2"
								Case Else
									sRowContents = sRowContents & "1"
							End Select
							sRowContents = sRowContents & lPayrollNumber 'Periodo I
							sRowContents = sRowContents & lPayrollNumber 'Periodo F
							sRowContents = sRowContents & "<CONCEPT_85 />" 'Crédito Adicional
							sRowContents = sRowContents & "<CONCEPT_60 />" 'Crédito Personal
							sRowContents = sRowContents & "<CONCEPT_62 />" 'Crédito FOVISSSTE
							sRowContents = sRowContents & "<CONCEPT_69 />" 'Pensión alimenticia
							sRowContents = sRowContents & "<CONCEPT_00 />" 'Percepción Total
							sRowContents = sRowContents & "1" 'Tipo de registro
						Case 2
							sRowContents = sRowContents & "<CONCEPT_01 />" 'Sueldo base
							sTemp = " "
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then sTemp = CStr(oRecordset.Fields("EmployeeLastName2").Value) & " "
							Err.Clear
							sRowContents = sRowContents & Left((CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & sTemp & CStr(oRecordset.Fields("EmployeeName").Value) & "                                        "), Len("                                        ")) 'Nombre del trabajador
							If CLng(oRecordset.Fields("LevelID").Value) = -1 Then
								sTemp = "000"
							Else
								sTemp = Right(("000" & CStr(oRecordset.Fields("LevelID").Value)), Len("000"))
							End If
							sRowContents = sRowContents & Right(("0000000000" & CStr(oRecordset.Fields("AreaShortName").Value)), Len("0000000000")) & Right(("000000" & CStr(oRecordset.Fields("EmployeeNumber").Value)), Len("000000")) & Right(("       " & CStr(oRecordset.Fields("PositionShortName").Value)), Len("       ")) & Left(sTemp, Len("00")) & " " & Right(sTemp, Len("0")) & "   " 'Clave de cobro del trabajador
							Select Case CInt(oRecordset.Fields("PositionTypeID").Value) 'Tipo de nombramiento
								Case 1
									sRowContents = sRowContents & "1" 'Base o planta
								Case 2
									sRowContents = sRowContents & "2" 'Confianza o supernumerario
								Case 3
									sRowContents = sRowContents & "5" 'Lista de yaua eventual honorarios
								Case 100
									sRowContents = sRowContents & "3" 'Interino o provisional
								Case 200
									sRowContents = sRowContents & "4" 'Lista de raya o base
								Case Else
									sRowContents = sRowContents & "6" 'Otros
							End Select
							Select Case CInt(oRecordset.Fields("PayrollTypeID").Value) 'Tipo de nómina
								Case 0
									sRowContents = sRowContents & "3"
								Case 2, 3, 4
									sRowContents = sRowContents & "2"
								Case Else
									sRowContents = sRowContents & "1"
							End Select
							sRowContents = sRowContents & " " 'Filler
							sRowContents = sRowContents & GetPayrollStartDate(lForPayrollID) & lForPayrollID 'Fecha
							sRowContents = sRowContents & "<CONCEPT_S8 /><CONCEPT_1S /><CONCEPT_04 /><CONCEPT_59 /><CONCEPT_62 /><CONCEPT_?? />" 'Apor Or
							sRowContents = sRowContents & "<CONCEPT_51 />" 'Ser Med
							sRowContents = sRowContents & "<CONCEPT_52 />" 'Fon Prest
							sRowContents = sRowContents & "<CONCEPT_85 />" 'Otros
							sRowContents = sRowContents & "0000" 'Filler
							sRowContents = sRowContents & "<CONCEPT_60 />" 'PCP
							sRowContents = sRowContents & "000000" 'ASM
							sRowContents = sRowContents & "<HAS_56 />" 'TH
							sRowContents = sRowContents & "<CONCEPT_56 />" 'IH
							sRowContents = sRowContents & "<HAS_55 />" 'TSH
							sRowContents = sRowContents & "<CONCEPT_55 />" 'ISH
							sTemp = ""
							sTemp = CStr(oRecordset.Fields("CURP").Value)
							Err.Clear
							sRowContents = sRowContents & Left((sTemp & "000000000000000000"), Len("000000000000000000")) 'CURP
							sRowContents = sRowContents & " " 'Filler
							sRowContents = sRowContents & "<SUMANDO />" 'Sumando	'xxxxxxxxxxxxx
							sRowContents = sRowContents & lPayrollNumber 'Rep_Quin
							sRowContents = sRowContents & lForPayrollID 'Fecha_Rep
							sRowContents = sRowContents & Right(("00" & CStr(oRecordset.Fields("ZoneCode").Value)), Len("00")) 'Entidad
							sTemp = ""
							sTemp = CStr(oRecordset.Fields("SocialSecurityNumber").Value)
							Err.Clear
							sRowContents = sRowContents & Left((sTemp & "00000000000"), Len("00000000000")) 'No. Seguro Social
							sRowContents = sRowContents & "000000000000" '"<CONCEPT_00 />" 'Sal_Sar
							sRowContents = sRowContents & "   " 'Filler
							sRowContents = sRowContents & "1" 'Tipo_Reg
					End Select

					sCurrentID = CStr(oRecordset.Fields("EmployeeID").Value)
				End If
				Select Case CLng(oRecordset.Fields("ConceptID").Value)
					Case 1, 5, 7, 8, 47, 89
						dTotal01 = dTotal01 + CDbl(oRecordset.Fields("ConceptAmount").Value)
					Case 2, 3, 12, 36
					Case 4
						dTotal01 = dTotal01 + CDbl(oRecordset.Fields("ConceptAmount").Value)
						sRowContents = Replace(sRowContents, "<CONCEPT_04 />", "1")
					Case 6
						dTotal01 = dTotal01 + CDbl(oRecordset.Fields("ConceptAmount").Value)
					Case 23
						sRowContents = Replace(sRowContents, "<CONCEPT_1S />", "1")
						adTotal(2) = adTotal(2) + CDbl(oRecordset.Fields("ConceptAmount").Value)
						adTotal(4) = adTotal(4) + CDbl(oRecordset.Fields("ConceptAmount").Value)
					Case 51
						adTotal(2) = adTotal(2) + CDbl(oRecordset.Fields("ConceptAmount").Value)
						adTotal(4) = adTotal(4) + CDbl(oRecordset.Fields("ConceptAmount").Value)
					Case 55, 63, 67, 68, 77
					Case 53
						'adTotal(3) = adTotal(3) + CDbl(oRecordset.Fields("ConceptAmount").Value)
					Case 54
						'sRowContents = Replace(sRowContents, "<CONCEPT_52 />", Right(("000000" & Int(CDbl(oRecordset.Fields("ConceptAmount").Value) * 100)), Len("000000")))
					Case 57
						adTotal(6) = adTotal(6) + CDbl(oRecordset.Fields("ConceptAmount").Value)
					Case 58
						adTotal(5) = adTotal(5) + CDbl(oRecordset.Fields("ConceptAmount").Value)
					Case 59
						adTotal(7) = adTotal(7) + CDbl(oRecordset.Fields("ConceptAmount").Value)
						adTotal(2) = adTotal(2) + CDbl(oRecordset.Fields("ConceptAmount").Value)
					Case 60
						sRowContents = Replace(sRowContents, "<CONCEPT_59 />", "1")
					Case 61 'Préstamo personal
						Select Case lType
							Case 2
								sRowContents = Replace(sRowContents, "<CONCEPT_60 />", Right(("000000" & Int(CDbl(oRecordset.Fields("ConceptAmount").Value) * 100)), Len("000000")))
							Case Else
								sRowContents = Replace(sRowContents, "<CONCEPT_60 />", Right(("0000000" & Int(CDbl(oRecordset.Fields("ConceptAmount").Value) * 100)), Len("0000000")))
						End Select
						adTotal(2) = adTotal(2) + CDbl(oRecordset.Fields("ConceptAmount").Value)
					Case 64 'Crédito FOVISSSTE
						adTotal(5) = adTotal(5) + CDbl(oRecordset.Fields("ConceptAmount").Value)
						Select Case lType
							Case 2
								sRowContents = Replace(sRowContents, "<CONCEPT_62 />", "1")
							Case Else
								sRowContents = Replace(sRowContents, "<CONCEPT_62 />", Right(("0000000" & Int(CDbl(oRecordset.Fields("ConceptAmount").Value) * 100)), Len("0000000")))
						End Select
						adTotal(2) = adTotal(2) + CDbl(oRecordset.Fields("ConceptAmount").Value)
					Case 66 'Seguro de vida Met Life II
						adTotal(5) = adTotal(5) + CDbl(oRecordset.Fields("ConceptAmount").Value)
					Case 70 'Pensión alimenticia
						sRowContents = Replace(sRowContents, "<CONCEPT_69 />", Right(("0000000" & Int(CDbl(oRecordset.Fields("ConceptAmount").Value) * 100)), Len("0000000")))
					Case 74, 78, 85, 105, 109, 126
					Case 82 'Crédito adicional
						Select Case lType
							Case 2
								sRowContents = Replace(sRowContents, "<CONCEPT_85 />", Right(("000000" & Int(CDbl(oRecordset.Fields("ConceptAmount").Value) * 100)), Len("000000")))
							Case Else
								sRowContents = Replace(sRowContents, "<CONCEPT_85 />", Right(("0000000" & Int(CDbl(oRecordset.Fields("ConceptAmount").Value) * 100)), Len("0000000")))
						End Select
						adTotal(2) = adTotal(2) + CDbl(oRecordset.Fields("ConceptAmount").Value)
					Case 116
						dTotalS7 = CDbl(oRecordset.Fields("ConceptAmount").Value)
						adTotal(2) = adTotal(2) + CDbl(oRecordset.Fields("ConceptAmount").Value)
						adTotal(3) = adTotal(3) + CDbl(oRecordset.Fields("ConceptAmount").Value)
					Case 117
						sRowContents = Replace(sRowContents, "<CONCEPT_S8 />", "1")
						adTotal(2) = adTotal(2) + CDbl(oRecordset.Fields("ConceptAmount").Value)
						adTotal(3) = adTotal(3) + CDbl(oRecordset.Fields("ConceptAmount").Value)
					Case 118
						adTotal(2) = adTotal(2) + CDbl(oRecordset.Fields("ConceptAmount").Value)
						adTotal(4) = adTotal(4) + CDbl(oRecordset.Fields("ConceptAmount").Value)
					Case 119
						adTotal(2) = adTotal(2) + CDbl(oRecordset.Fields("ConceptAmount").Value)
					Case 20000
						sRowContents = Replace(sRowContents, "<CONCEPT_?? />", Right(("0000" & Int(CDbl(oRecordset.Fields("ConceptAmount").Value) * 100)), Len("0000")))
					Case Else
				End Select
				Select Case CStr(oRecordset.Fields("ConceptShortName").Value)
					Case "01", "02", "03", "04", "05", "06", "07", "08", "10", "11", "12", "1S", "33", "35", "48", "B2", "B3", "E2", "E3"
						adTotal(0) = adTotal(0) + CDbl(oRecordset.Fields("ConceptAmount").Value)
					Case "55", "56", "57", "58", "59", "60", "61", "62", "63", "64", "65", "66", "73", "78", "81", "83", "85", "86", "87", "88", "89", "ET", "MT", "PP", "SH", "SR", "54", "CS"
						adTotal(8) = adTotal(8) + CDbl(oRecordset.Fields("ConceptAmount").Value)
						adTotal(1) = adTotal(1) + CDbl(oRecordset.Fields("ConceptAmount").Value)
					Case Else
						If CInt(oRecordset.Fields("IsDeduction").Value) Then
							adTotal(1) = adTotal(1) + CDbl(oRecordset.Fields("ConceptAmount").Value)
						End If
				End Select
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
			If ((dTotal01 / 100) * 0.0275) <> dTotalS7 Then dTotal01 = FormatNumber((dTotalS7 / 0.0275), 2, True, False, False)
			Select Case lType
				Case 2
					sRowContents = Replace(sRowContents, "<CONCEPT_00 />", Right(("000000000000" & Int(adTotal(0) * 100)), Len("000000000000")))
					sRowContents = Replace(sRowContents, "<CONCEPT_01 />", Right(("000000000000" & (dTotal01 * 100)), Len("000000000000")))
					sRowContents = Replace(sRowContents, "<CONCEPT_60 />", "000000")
					sRowContents = Replace(sRowContents, "<CONCEPT_62 />", "0")
					sRowContents = Replace(sRowContents, "<CONCEPT_85 />", "000000")
				Case Else
					sRowContents = Replace(sRowContents, "<CONCEPT_00 />", Right(("000000000000" & Int(adTotal(0) * 100)), Len("000000000000")))
					sRowContents = Replace(sRowContents, "<CONCEPT_01 />", Right(("000000000000" & (dTotal01 * 100)), Len("000000000000")))
					sRowContents = Replace(sRowContents, "<CONCEPT_60 />", "0000000")
					sRowContents = Replace(sRowContents, "<CONCEPT_62 />", "0000000")
					sRowContents = Replace(sRowContents, "<CONCEPT_85 />", "0000000")
			End Select
			sRowContents = Replace(sRowContents, "<CONCEPT_69 />", "0000000")
			sRowContents = Replace(sRowContents, "<CONCEPT_S8 />", "0")
			sRowContents = Replace(sRowContents, "<CONCEPT_1S />", "0")
			sRowContents = Replace(sRowContents, "<CONCEPT_04 />", "0")
			sRowContents = Replace(sRowContents, "<CONCEPT_59 />", "0")
			sRowContents = Replace(sRowContents, "<CONCEPT_?? />", "0000")
			sRowContents = Replace(sRowContents, "<CONCEPT_51 />", Right(("000000" & Int(adTotal(3) * 100)), Len("000000")))
			sRowContents = Replace(sRowContents, "<CONCEPT_52 />", Right(("000000" & Int(adTotal(4) * 100)), Len("000000")))
			If adTotal(6) > 0 Then sRowContents = Replace(sRowContents, "<HAS_55 />", "07")
			If adTotal(7) > 0 Then sRowContents = Replace(sRowContents, "<HAS_55 />", "09")
			sRowContents = Replace(sRowContents, "<HAS_55 />", "00")
			sRowContents = Replace(sRowContents, "<CONCEPT_55 />", Right(("000000" & ((adTotal(6) + adTotal(7)) * 100)), Len("000000")))
			If adTotal(5) > 0 Then
				sRowContents = Replace(sRowContents, "<HAS_56 />", "06")
			Else
				sRowContents = Replace(sRowContents, "<HAS_56 />", "00")
			End If
			sRowContents = Replace(sRowContents, "<CONCEPT_56 />", Right(("000000" & (adTotal(5) * 100)), Len("000000")))

			sRowContents = Replace(sRowContents, "<SUMANDO />", Right(("0000000" & Int(adTotal(2) * 100)), Len("0000000")))
			sRowContents = Replace(sRowContents, "<TOTAL />", Right(("000000000000" & Int((adTotal(0) - adTotal(1) + adTotal(8)) * 100)), Len("000000")))
			lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
		End If
		If FileExists(sFilePath, sErrorDescription) Then
			lErrorNumber = ZipFile(sFilePath, Replace(sFilePath, ".txt", ".zip"), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudieron guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				lErrorNumber = DeleteFile(sFilePath, sErrorDescription)
			End If
			oEndDate = Now()
			If (lErrorNumber = 0) And B_USE_SMTP Then
				If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(Replace(sFilePath, ".txt", ".zip"), CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
			End If
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1339 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1354(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the report form for the EmployeesKardex5 table
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1354"
	Dim iKardex5TypeID
	Dim sFileContents
	Dim lCounter
	Dim oDate
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim sCurrentRecords
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	iKardex5TypeID = CInt(oRequest("Kardex5TypeID").Item)
	sFileContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1354.htm"), sErrorDescription)
	If Len(sFileContents) > 0 Then
		If iKardex5TypeID = 0 Then
			sFileContents = Replace(sFileContents, "<TITLE />", "CANDIDATOS TIPO CRONOLÓGICO")
			sFileContents = Replace(sFileContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME, 1, -1, vbBinaryCompare)
			Response.Write sFileContents

			Response.Write "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0""><TR>" & vbNewLine
				Response.Write "<TD VALIGN=""TOP"">"
					sErrorDescription = "No se pudo obtener la información de los registros de la bolsa de trabajo."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesKardex5 Where (Kardex5TypeID=0) And (Kardex5OriginID=0) Order By EmployeeLastName, EmployeeLastName2, EmployeeName, StartDate", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							Response.Write "<TABLE BORDER="""
								If Not bForExport Then
									Response.Write "0"
								Else
									Response.Write "1"
								End If
							Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
								asColumnsTitles = Split("<SPAN COLS=""9"" />LISTADO DEL INSTITUTO", ",", -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
								Else
									If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
										lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									Else
										lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									End If
								End If
								asColumnsTitles = Split("No,Nombre,<SPAN COLS=""5"" />Fecha de registro,Nominación,Motivo de baja", ",", -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
								Else
									If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
										lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									Else
										lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									End If
								End If

								lCounter = 1
								sCurrentRecords = ""
								Do While Not oRecordset.EOF
									sRowContents = lCounter
									If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
										sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & ", " & CStr(oRecordset.Fields("EmployeeName").Value))
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value))
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""5"" />" & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Nomination").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Reasons").Value))
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
							sErrorDescription = "No existen registos en la base de datos que cumplan con los criterios del filtro."
						End If
					End If
				Response.Write "</TD>" & vbNewLine
				Response.Write "<TD>&nbsp;</TD>" & vbNewLine
				Response.Write "<TD VALIGN=""TOP"">"
					sErrorDescription = "No se pudo obtener la información de los registros de la bolsa de trabajo."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesKardex5 Where (Kardex5TypeID=0) And (Kardex5OriginID=1) Order By EmployeeLastName, EmployeeLastName2, EmployeeName, StartDate", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							Response.Write "<TABLE BORDER="""
								If Not bForExport Then
									Response.Write "0"
								Else
									Response.Write "1"
								End If
							Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
								asColumnsTitles = Split("<SPAN COLS=""9"" />LISTADO DEL SINDICATO", ",", -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
								Else
									If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
										lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									Else
										lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									End If
								End If
								asColumnsTitles = Split("No,Nombre,<SPAN COLS=""5"" />Fecha de registro,Nominación,Motivo de baja", ",", -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
								Else
									If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
										lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									Else
										lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									End If
								End If

								lCounter = 1
								sCurrentRecords = ""
								Do While Not oRecordset.EOF
									sRowContents = lCounter
									If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
										sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & ", " & CStr(oRecordset.Fields("EmployeeName").Value))
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value))
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""5"" />" & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Nomination").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Reasons").Value))
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
							sErrorDescription = "No existen registos en la base de datos que cumplan con los criterios del filtro."
						End If
					End If
				Response.Write "</TD>" & vbNewLine
			Response.Write "</TR></TABLE>" & vbNewLine
		Else
			sFileContents = Replace(sFileContents, "<TITLE />", "BOLSA DE TRABAJO TIPO PUNTUACIÓN")
			sFileContents = Replace(sFileContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME, 1, -1, vbBinaryCompare)
			Response.Write sFileContents

			Response.Write "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0""><TR>" & vbNewLine
				Response.Write "<TD VALIGN=""TOP"">"
					sErrorDescription = "No se pudo obtener la información de los registros de la bolsa de trabajo."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesKardex5.*, EmployeeName, EmployeeLastName, EmployeeLastName2, SchoolarshipName From EmployeesKardex5, Schoolarships Where (EmployeesKardex5.SchoolarshipID=Schoolarships.SchoolarshipID) And (Kardex5TypeID=1) And (Kardex5OriginID=0) Order By EmployeeLastName, EmployeeLastName2, EmployeeName, StartDate", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							Response.Write "<TABLE BORDER="""
								If Not bForExport Then
									Response.Write "0"
								Else
									Response.Write "1"
								End If
							Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
								asColumnsTitles = Split("<SPAN COLS=""9"" />LISTADO DEL INSTITUTO", ",", -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
								Else
									If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
										lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									Else
										lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									End If
								End If
								asColumnsTitles = Split("No,Nombre,E,P,TS,TR,TOTAL,Nominación,Motivo de baja", ",", -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
								Else
									If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
										lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									Else
										lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									End If
								End If

								lCounter = 1
								sCurrentRecords = ""
								Do While Not oRecordset.EOF
									sRowContents = lCounter
									If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
										sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & ", " & CStr(oRecordset.Fields("EmployeeName").Value))
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value))
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SchoolarshipName").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Relationship").Value))
									oDate = Date()
									sRowContents = sRowContents & TABLE_SEPARATOR & DateDiff("m", GetDateFromSerialNumber(CStr(oRecordset.Fields("ServiceYears").Value)), oDate) & "&nbsp;meses"
									oDate = Date()
									sRowContents = sRowContents & TABLE_SEPARATOR & DateDiff("m", GetDateFromSerialNumber(CStr(oRecordset.Fields("KardexYears").Value)), oDate) & "&nbsp;meses"
									sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Nomination").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Reasons").Value))
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
							sErrorDescription = "No existen registos en la base de datos que cumplan con los criterios del filtro."
						End If
					End If
				Response.Write "</TD>" & vbNewLine
				Response.Write "<TD>&nbsp;</TD>" & vbNewLine
				Response.Write "<TD VALIGN=""TOP"">"
					sErrorDescription = "No se pudo obtener la información de los registros de la bolsa de trabajo."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesKardex5.*, SchoolarshipName From EmployeesKardex5, Schoolarships Where (EmployeesKardex5.SchoolarshipID=Schoolarships.SchoolarshipID) And (Kardex5TypeID=1) And (Kardex5OriginID=1) Order By EmployeeLastName, EmployeeLastName2, EmployeeName, StartDate", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							Response.Write "<TABLE BORDER="""
								If Not bForExport Then
									Response.Write "0"
								Else
									Response.Write "1"
								End If
							Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
								asColumnsTitles = Split("<SPAN COLS=""9"" />LISTADO DEL SINDICATO", ",", -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
								Else
									If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
										lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									Else
										lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									End If
								End If
								asColumnsTitles = Split("No,Nombre,E,P,TS,TR,TOTAL,Nominación,Motivo de baja", ",", -1, vbBinaryCompare)
								If bForExport Then
									lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
								Else
									If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
										lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									Else
										lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
									End If
								End If

								lCounter = 1
								sCurrentRecords = ""
								Do While Not oRecordset.EOF
									sRowContents = lCounter
									If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
										sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & ", " & CStr(oRecordset.Fields("EmployeeName").Value))
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value))
									End If
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("SchoolarshipName").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Relationship").Value))
									oDate = Date()
									sRowContents = sRowContents & TABLE_SEPARATOR & DateDiff("m", GetDateFromSerialNumber(CStr(oRecordset.Fields("ServiceYears").Value)), oDate) & "&nbsp;meses"
									oDate = Date()
									sRowContents = sRowContents & TABLE_SEPARATOR & DateDiff("m", GetDateFromSerialNumber(CStr(oRecordset.Fields("KardexYears").Value)), oDate) & "&nbsp;meses"
									sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Nomination").Value))
									sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Reasons").Value))
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
							sErrorDescription = "No existen registos en la base de datos que cumplan con los criterios del filtro."
						End If
					End If
				Response.Write "</TD>" & vbNewLine
			Response.Write "</TR></TABLE>" & vbNewLine
			Response.Write "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR><TD COLSPAN=""19""><FONT FACE=""Arial"" SIZE=""2"">E = Escolaridad&nbsp;&nbsp;&nbsp;P = Parentesco&nbsp;&nbsp;&nbsp;TS = Tiempo de servicio en el Instituto&nbsp;&nbsp;&nbsp;TR = Tiempo en el registro de bolsa de trabajo&nbsp;&nbsp;&nbsp;</FONT></TD></TR></TABLE>" & vbNewLine
		End If
	End If
	sFileContents = GetFileContents(Server.MapPath("Templates\FooterForReport_1354.htm"), sErrorDescription)
	sFileContents = Replace(sFileContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME, 1, -1, vbBinaryCompare)
	Response.Write sFileContents

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1354 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1356(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the report form for the EmployeesKardex4 table
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1356"
	Dim iKardexChangeTypeID
	Dim sFileContents
	Dim sNames
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim sCurrentRecords
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	iKardexChangeTypeID = CInt(oRequest("KardexChangeTypeID").Item)
	sFileContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1356.htm"), sErrorDescription)
	If Len(sFileContents) > 0 Then
		Call GetNameFromTable(oADODBConnection, "KardexChangeTypes", iKardexChangeTypeID, "", "", sNames, "")
		sFileContents = Replace(sFileContents, "<TITLE />", UCase(sNames))
		sFileContents = Replace(sFileContents, "<CURRENT_DATE />", asMonthNames_es(Month(Date())) & "/" & Year(Date()), 1, -1, vbBinaryCompare)
		sFileContents = Replace(sFileContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME, 1, -1, vbBinaryCompare)
		Response.Write sFileContents

		sErrorDescription = "No se pudo obtener la información de los registros de escalafón."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesKardex4.*, EmployeeName, EmployeeLastName, EmployeeLastName2, Employees.WorkingHours, Positions.PositionShortName, Positions.PositionName, Journeys.JourneyShortName, Journeys.JourneyName, ShiftName, Services.ServiceName, NewPositions.PositionShortName As NewPositionShortName, NewPositions.PositionName As NewPositionName, NewServices.ServiceName As NewServiceName, NewJourneys.JourneyShortName As NewJourneyShortName, NewJourneys.JourneyName As NewJourneyName, Areas.AreaName, NewAreas.AreaName As NewAreaName From EmployeesKardex4, Employees, Jobs, Positions, Journeys, Shifts, Services, Services As NewServices, Positions As NewPositions, Journeys As NewJourneys, Areas, Areas As NewAreas Where (EmployeesKardex4.EmployeeID=Employees.EmployeeID) And (Employees.JobID=Jobs.JobID) And (Jobs.PositionID=Positions.PositionID) And (Employees.JourneyID=Journeys.JourneyID) And (Employees.ShiftID=Shifts.ShiftID) And (Employees.ServiceID=Services.ServiceID) And (EmployeesKardex4.PositionID=NewPositions.PositionID) And (EmployeesKardex4.ServiceID=NewServices.ServiceID) And (EmployeesKardex4.JourneyID=NewJourneys.JourneyID) And (Jobs.AreaID=Areas.AreaID) And (EmployeesKardex4.AreaID=NewAreas.AreaID) And (KardexChangeTypeID=" & iKardexChangeTypeID & ") Order By EmployeesKardex4.StartDate, EmployeeNumber", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				Response.Write "<TABLE BORDER="""
					If bForExport Then
						Response.Write "1"
					Else
						Response.Write "0"
					End If
				Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine

					sCurrentRecords = ""
					Do While Not oRecordset.EOF
						sRowContents = "No. REGISTRO GENERAL:<BR />&nbsp;&nbsp;&nbsp;" & CleanStringForHTML(CStr(oRecordset.Fields("KardexNumber1").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & "No. REGISTRO INDIVIDUAL:<BR />&nbsp;&nbsp;&nbsp;" & CleanStringForHTML(CStr(oRecordset.Fields("KardexNumber2").Value))
						If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
							sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""4"" />NOMBRE:<BR />&nbsp;&nbsp;&nbsp;" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & ", " & CStr(oRecordset.Fields("EmployeeName").Value))
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""4"" />NOMBRE:<BR />&nbsp;&nbsp;&nbsp;" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value))
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & "NO. DE PLAZA:<BR />&nbsp;&nbsp;&nbsp;" & CleanStringForHTML(CStr(oRecordset.Fields("JobID").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""2"" />No. EMPLEADO:<BR />&nbsp;&nbsp;&nbsp;" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeID").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>Observaciones</CENTER>" & CleanStringForHTML(CStr(oRecordset.Fields("Comments").Value))
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "<NO_LINE />", "", "", sErrorDescription)
						End If

						If (iKardexChangeTypeID = 0) Or (iKardexChangeTypeID = 1) Then
							sRowContents = "<CENTER>PUESTO QUE OCUPA:</CENTER>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>JORNADA</CENTER>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>CLAVE</CENTER>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>PUESTO<BR />SOLICITADO</CENTER>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>CLAVE</CENTER>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>SERVICIO SOLICITADO</CENTER>"
						Else
							sRowContents = "<CENTER>PUESTO DENOMINACIÓN</CENTER>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>CLAVE</CENTER>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>RANGO</CENTER>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>JORNADA</CENTER>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>ESPECIALIDAD</CENTER>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>TURNO ACTUAL</CENTER>"
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>TURNO SOLICITADO</CENTER>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>FECHA DE REGISTRO</CENTER>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>FECHA DE RESOLUCIÓN</CENTER>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "<NO_LINE />", "", "", sErrorDescription)
						End If

						If (iKardexChangeTypeID = 0) Or (iKardexChangeTypeID = 1) Then
							sRowContents = "<CENTER>" & CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value)) & "</CENTER>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>" & CleanStringForHTML(CStr(oRecordset.Fields("WorkingHours").Value)) & "</CENTER>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>" & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value)) & "</CENTER>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>" & CleanStringForHTML(CStr(oRecordset.Fields("NewPositionName").Value)) & "</CENTER>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>" & CleanStringForHTML(CStr(oRecordset.Fields("NewPositionShortName").Value)) & "</CENTER>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>" & CleanStringForHTML(CStr(oRecordset.Fields("NewServiceName").Value)) & "</CENTER>"
						Else
							sRowContents = "<CENTER>" & CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value)) & "</CENTER>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>" & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value)) & "</CENTER>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>???</CENTER>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>" & CleanStringForHTML(CStr(oRecordset.Fields("WorkingHours").Value)) & "</CENTER>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>" & CleanStringForHTML(CStr(oRecordset.Fields("ServiceName").Value)) & "</CENTER>"
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("JourneyShortName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("JourneyName").Value)) & "<BR />" & CleanStringForHTML(CStr(oRecordset.Fields("ShiftName").Value))
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>" & CleanStringForHTML(CStr(oRecordset.Fields("NewJourneyShortName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("NewJourneyName").Value)) & "</CENTER>"
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
						sRowContents = sRowContents & TABLE_SEPARATOR
						If CLng(oRecordset.Fields("EndDate").Value) = 0 Then
							sRowContents = sRowContents & "&nbsp;"
						Else
							sRowContents = sRowContents & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1)
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "<NO_LINE />", "", "", sErrorDescription)
						End If

						If (iKardexChangeTypeID = 0) Or (iKardexChangeTypeID = 1) Then
							sRowContents = "<SPAN COLS=""3"" />ESPECIALIDAD<BR />" & CleanStringForHTML(CStr(oRecordset.Fields("ServiceName").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""2"" />ADSCRIPCIÓN ACTUAL<BR />" & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""2"" />ADSCRIPCIÓN SOLICITADA<BR />" & CleanStringForHTML(CStr(oRecordset.Fields("NewAreaName").Value))
						Else
							sRowContents = "<SPAN COLS=""4"" />ADSCRIPCIÓN ACTUAL<BR />" & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""3"" />ADSCRIPCIÓN SOLICITADA<BR />" & CleanStringForHTML(CStr(oRecordset.Fields("NewAreaName").Value))
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
						sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<B>No. DICT. SME-" & CleanStringForHTML(CStr(oRecordset.Fields("DocumentNumber").Value)) & "</B>"
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If

						Response.Write "<TR><TD COLSPAN=""10"">&nbsp;</TD></TR>" & vbNewLine
						lCounter = lCounter + 1
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
				Response.Write "</TABLE>" & vbNewLine
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registos en la base de datos que cumplan con los criterios del filtro."
			End If
		End If
	End If
	sFileContents = GetFileContents(Server.MapPath("Templates\FooterForReport_1356.htm"), sErrorDescription)
	sFileContents = Replace(sFileContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME, 1, -1, vbBinaryCompare)
	Response.Write sFileContents

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1356 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1364(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Reporte de Desarrollo humano
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1364"
	Dim sCondition
	Dim sFieldNames
	Dim sTableNames
	Dim sJoinCondition
	Dim sSortFields
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim sCurrentRecords
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	Call GetDBFieldsNames(oRequest, 0, sCondition, sFieldNames, sTableNames, sJoinCondition, sSortFields)
	sTableNames = Replace(sTableNames, "PositionTypes", "Employees")
	If (InStr(1, " " & sTableNames & ",", " SADE_Curso,", vbBinaryCompare) = 0) Then sTableNames = "SADE_Curso, " & sTableNames
	If (InStr(1, " " & sTableNames & ",", " Employees,", vbBinaryCompare) > 0) Or (InStr(1, " " & sTableNames & ",", " Jobs,", vbBinaryCompare) > 0) Or (InStr(1, " " & sTableNames & ",", " Areas,", vbBinaryCompare) > 0) Or (InStr(1, " " & sTableNames & ",", " Companies,", vbBinaryCompare) > 0) Then
		If (InStr(1, " " & sTableNames & ",", " Employees,", vbBinaryCompare) = 0) Then sTableNames = "Employees, " & sTableNames
		If (InStr(1, " " & sTableNames & ",", " Jobs,", vbBinaryCompare) = 0) Then sTableNames = "Jobs, " & sTableNames
		If (InStr(1, " " & sTableNames & ",", " Areas,", vbBinaryCompare) = 0) Then sTableNames = "Areas, " & sTableNames
		If (InStr(1, " " & sTableNames & ",", " Companies,", vbBinaryCompare) = 0) Then sTableNames = "Companies, " & sTableNames
	End If
	If ((InStr(1, sJoinCondition, "=SADE_CursosEmpleadosLKP.ID_Curso)", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(SADE_CursosEmpleadosLKP.ID_Curso", vbBinaryCompare) = 0)) And (InStr(1, " " & sTableNames & ",", " Employees,", vbBinaryCompare) > 0) Then
		sJoinCondition = sJoinCondition & " And (SADE_Curso.ID_Curso=SADE_CursosEmpleadosLKP.ID_Curso)"
		If (InStr(1, " " & sTableNames & ",", " SADE_CursosEmpleadosLKP,", vbBinaryCompare) = 0) Then sTableNames = "SADE_CursosEmpleadosLKP, " & sTableNames
	End If
	If ((InStr(1, sJoinCondition, "=Employees.EmployeeID)", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(Employees.EmployeeID", vbBinaryCompare) = 0)) And (InStr(1, " " & sTableNames & ",", " Employees,", vbBinaryCompare) > 0) Then sJoinCondition = sJoinCondition & " And (SADE_CursosEmpleadosLKP.ID_Empleado=Employees.EmployeeID)"
	If ((InStr(1, sJoinCondition, "=Jobs.JobID)", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(Jobs.JobID", vbBinaryCompare) = 0)) And (InStr(1, " " & sTableNames & ",", " Jobs,", vbBinaryCompare) > 0) Then sJoinCondition = sJoinCondition & " And (Employees.JobID=Jobs.JobID)"
	If ((InStr(1, sJoinCondition, "=Areas.AreaID)", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(Areas.AreaID", vbBinaryCompare) = 0)) And (InStr(1, " " & sTableNames & ",", " Areas,", vbBinaryCompare) > 0) Then sJoinCondition = sJoinCondition & " And (Jobs.AreaID=Areas.AreaID)"
	If ((InStr(1, sJoinCondition, "=Companies.CompanyID)", vbBinaryCompare) = 0) Or (InStr(1, sJoinCondition, "(Companies.CompanyID", vbBinaryCompare) = 0)) And (InStr(1, " " & sTableNames & ",", " Companies,", vbBinaryCompare) > 0) Then sJoinCondition = sJoinCondition & " And (Employees.CompanyID=Companies.CompanyID)"
	sJoinCondition = Replace(sJoinCondition, " And (Positions.PositionTypeID=PositionTypes.PositionTypeID)", "")
	sJoinCondition = Replace(sJoinCondition, "PositionTypes.PositionTypeID", "Employees.PositionTypeID")
	sCondition = Replace(sCondition, "PositionTypes.PositionTypeID", "Employees.PositionTypeID")

	sErrorDescription = "No se pudo obtener la información de los cursos de capacitación para desarrollo humano."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct SADE_Curso.ID_Curso " & sFieldNames & " From " & sTableNames & " Where (SADE_Curso.ID_Curso>-1) " & sCondition & sJoinCondition & " Order By " & sSortFields, "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split(BuildTableTemplateHeader(oRequest, 0, ""), ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				sCurrentRecords = ""
				Do While Not oRecordset.EOF
					lErrorNumber = BuildRowData(oRecordset, 1, -1, sRowContents)
					If StrComp(sRowContents, sCurrentRecords, vbBinaryCompare) <> 0 Then
						sCurrentRecords = sRowContents
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
					End If
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registos en la base de datos que cumplan con los criterios del filtro."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	BuildReport1364 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1371(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de plantilla de personal
'         Carpeta 1. DOCUMENTACIÓN ENTREGADA POR JEFATURA DE SERVICIOS DE DESARROLLO HUMANO
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1371"
	Dim sHeaderContents
	Dim oRecordset
	Dim oPayrollRecordset
	Dim sContents
	Dim sRowContents
	Dim lErrorNumber
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim sSourceFolderPath
	Dim lPayrollID
	Dim lForPayrollID
	Dim sCondition
	Dim iConcept
	Dim lEmployeeID
	Dim dConcept01
	Dim dConcept03
	Dim sPreviousZone
	Dim iCounter
	iCounter = 0

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	oStartDate = Now()
	sDate = GetSerialNumberForDate("")
	sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
	lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
	sFilePath = sFilePath & "\"
	sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
	Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
	Response.Flush()
	If lErrorNumber = 0 Then
		sErrorDescription = "No se pudieron obtener los registros de los empleados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryListForPayroll.EmployeeID, CashierOfficeShortName, EmployeesHistoryListForPayroll.EmployeeNumber, RFC, CURP, SocialSecurityNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, EmployeesHistoryListForPayroll.PositionTypeID, EmployeesHistoryListForPayroll.LevelID, ParentZones.ZoneCode, PaymentCenters.AreaShortName, AccountNumber, PositionShortName, Concepts.ConceptID, Concepts.ConceptShortName, Concepts.IsDeduction, Payroll_" & lPayrollID & ".ConceptAmount, Payrolls.PayrollTypeID, ForPayrollDate From Payrolls, Payroll_" & lPayrollID & ", Concepts, Employees, EmployeesHistoryListForPayroll, Areas, Areas As PaymentCenters, CashierOffices, Zones, Zones As Zones2, Zones As ParentZones, Positions Where (Payrolls.PayrollID=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (Areas.CashierOfficeID=CashierOffices.CashierOfficeID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID>0) " & Replace(sCondition, "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Order By EmployeesHistoryListForPayroll.EmployeeNumber, RecordDate, Concepts.IsDeduction, ConceptShortName.ZoneID, Zones.ZoneName, ParentZones.ZoneName As ParentZoneName, Entidades.ZoneName As EntidadesZoneName, GeneratingAreas.GeneratingAreaName", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sPreviousZone = ""
				Do While Not oRecordset.EOF
					If StrComp(sPreviousZone, CStr(oRecordset.Fields("EntidadesZoneName").Value), vbBinaryCompare) <> 0 Then
						If Len(sPreviousZone) > 0 Then
							lErrorNumber = AppendTextToFile(sDocumentName, "</TABLE>", sErrorDescription)
						End If
						sDocumentName = sFilePath & "RUSP_" & CStr(oRecordset.Fields("EntidadesZoneName").Value) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"
						sHeaderContents = Replace(sHeaderContents, "<PAYROLL_DATE />", DisplayNumericDateFromSerialNumber(lForPayrollID))
						sRowContents = "<TABLE BORDER=""1"">"
						sRowContents = sRowContents & "<TR>"
						sRowContents = sRowContents & "<TD>" & "AREA GEN" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "CENTRAB" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "DEN_CENT" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "POBLACION" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "MUNICIPIO" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "SERVICIO" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "DEN_SERV" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "NUM_PLAZA" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "CLAVE PUESTO" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "RAMO" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "UNID RESP" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "CONSEC UNICO" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "CONSEC_JEFE" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "DEN_PUESTO" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "PUESTO PRESUP" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "NIVEL TAB" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "SUB NIV T" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "GPO GDO NIV" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "CLASIF" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "ZONA_ECO" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "SUELDO_MES" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "COMPENSA_MES" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "ENTIDAD PLAZA" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "PAIS PLAZA" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "TIPO PLAZA" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "TIPO PUES ESTRAT" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "TIPO FUNCION PUESTO" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "TIPO PERSONAL" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "CODIGO PUESTO RHNET" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "ESTAT PLAZA" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "RFC" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "CURP" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "NOMBRE" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "1APELLIDO" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "2APELLIDO" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "FECHA NACIMIENTO" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "SEXO" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "ENTIDAD NACIMIENTO" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "PAIS NACIMIENTO" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "E MAIL INSTITUCIONAL" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "INSTIT SEG SOC" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "NUMERO SEGURIDAD SOCIAL" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "CLAVE PRESUP SEP" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "NIVEL TAB PLAZA" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "TIPO CONTRATACION" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "DECLARA PATRIM" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "MOTIVO DEC PAT" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "NUM_EMPLEADO" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "FECHA INGRESO ADM PUB FED" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "FECHA INGRESO SER PROF CARR" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "FECHA INGRESO REING INSTITUTO" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "FECHA ALTA ULTIMO PUESTO" & "</TD>"
						sRowContents = sRowContents & "<TD>" & "FECHA OBLIGACION A DECLAR PATRIM" & "</TD>"
						sRowContents = sRowContents & "</TR>"
						lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
						sPreviousZone = CStr(oRecordset.Fields("EntidadesZoneName").Value)
					End If
					lEmployeeID = CLng(oRecordset.Fields("EmployeeID").Value)
					sErrorDescription = "No se pudieron obtener los registros de los empleados."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptID, ConceptAmount From Payroll_" & lPayrollID & " Where ConceptID In (1,2) And EmployeeID=" & lEmployeeID, "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oPayrollRecordset)
					If lErrorNumber = 0 Then
						If Not oPayrollRecordset.EOF Then
							Do While Not oPayrollRecordset.EOF
								iConcept = CInt(oPayrollRecordset.Fields("ConceptID").Value)
								Select Case iConcept
									Case 1
									    dConcept01 = CDbl(oPayrollRecordset.Fields("ConceptAmount").Value)
									Case 2
									    dConcept03 = CDbl(oPayrollRecordset.Fields("ConceptAmount").Value)
								End Select
								oPayrollRecordset.MoveNext
								If Err.number <> 0 Then Exit Do
							Loop
						End If
					End If
					iCounter = iCounter + 1
					sRowContents = "<TR>"
						sRowContents = sRowContents & "<TD>" & CleanSringForHTML(CStr(oRecordset.Fields("GeneratingAreaName").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanSringForHTML(CStr(oRecordset.Fields("AreaCode").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanSringForHTML(CStr(oRecordset.Fields("AreaName").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanSringForHTML(CStr(oRecordset.Fields("ZoneName").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanSringForHTML(CStr(oRecordset.Fields("ParentZoneName").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanSringForHTML(CStr(oRecordset.Fields("ServiceShortName").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanSringForHTML(CStr(oRecordset.Fields("ServiceName").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>=T(""" & CleanSringForHTML(CStr(oRecordset.Fields("JobNumber").Value)) & """)</TD>"
						sRowContents = sRowContents & "<TD>" & CleanSringForHTML(CStr(oRecordset.Fields("PositionShortName").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>51</TD>"
						sRowContents = sRowContents & "<TD>GYN</TD>"
						sRowContents = sRowContents & "<TD>" & iCounter & "</TD>"
						sRowContents = sRowContents & "<TD></TD>"
						sRowContents = sRowContents & "<TD>" & CleanSringForHTML(CStr(oRecordset.Fields("PositionName").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanSringForHTML(CStr(oRecordset.Fields("PositionShortName").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>=T(""" & CleanSringForHTML(Left(CStr(oRecordset.Fields("LevelShortName").Value), Len("00"))) & """)</TD>"
						sRowContents = sRowContents & "<TD>=T(""" & CleanSringForHTML(Right(CStr(oRecordset.Fields("LevelShortName").Value), Len("0"))) & """)</TD>"
						sRowContents = sRowContents & "<TD>" & CleanSringForHTML(CStr(oRecordset.Fields("GroupGradeLevelShortName").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanSringForHTML(CStr(oRecordset.Fields("ClassificationID").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanSringForHTML(CStr(oRecordset.Fields("ZoneTypeName").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & asConceptAmount(1) & "</TD>"
						asConceptAmount(1) = 0
						sRowContents = sRowContents & "<TD>" & asConceptAmount(3) & "</TD>"
						asConceptAmount(3) = 0
						sRowContents = sRowContents & "<TD>" & CleanSringForHTML(CStr(oRecordset.Fields("EntidadesZoneName").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>700</TD>"
						sRowContents = sRowContents & "<TD>1</TD>"
						sRowContents = sRowContents & "<TD>3</TD>"
						sRowContents = sRowContents & "<TD>4</TD>"
						sRowContents = sRowContents & "<TD>4</TD>"
						sRowContents = sRowContents & "<TD></TD>"
						sRowContents = sRowContents & "<TD>1</TD>"
						sRowContents = sRowContents & "<TD>"
							sRowContents = sRowContents & CleanSringForHTML(CStr(oRecordset.Fields("RFC").Value))
						sRowContents = sRowContents & "</TD>"
						sRowContents = sRowContents & "<TD>"
							sRowContents = sRowContents & CleanSringForHTML(CStr(oRecordset.Fields("CURP").Value))
						sRowContents = sRowContents & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanSringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanSringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>"
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
								sRowContents = sRowContents & CleanSringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value))
							Else
								sRowContents = sRowContents & " "
							End If
						sRowContents = sRowContents & "</TD>"
						sRowContents = sRowContents & "<TD>" & DisplayShortDateFromSerialNumber(CLng(oRecordset.Fields("BirthDate").Value), -1, -1 ,-1) & "</TD>"
						sRowContents = sRowContents & "<TD>" & CleanSringForHTML(CStr(oRecordset.Fields("GenderID").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD></TD>"
						sRowContents = sRowContents & "<TD>700</TD>"
						sRowContents = sRowContents & "<TD>"
							sRowContents = sRowContents & CleanSringForHTML(CStr(oRecordset.Fields("EmployeeEmail").Value))
						sRowContents = sRowContents & "</TD>"
						sRowContents = sRowContents & "<TD>1</TD>"
						sRowContents = sRowContents & "<TD>"
							sRowContents = sRowContents & CleanSringForHTML(CStr(oRecordset.Fields("SocialSecurityNumber").Value))
						sRowContents = sRowContents & "</TD>"
						sRowContents = sRowContents & "<TD></TD>"
						sRowContents = sRowContents & "<TD>" & CleanSringForHTML(CStr(oRecordset.Fields("LevelShortName").Value)) & "</TD>"
						sRowContents = sRowContents & "<TD>2</TD>"
						sRowContents = sRowContents & "<TD>N</TD>"
						sRowContents = sRowContents & "<TD></TD>"
						sRowContents = sRowContents & "<TD>=T(""" & CleanSringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & """)</TD>"
						sRowContents = sRowContents & "<TD>" & DisplayShortDateFromSerialNumber(CLng(oRecordset.Fields("StartDate2").Value), -1, -1 ,-1) & "</TD>"
						sRowContents = sRowContents & "<TD></TD>"
						sRowContents = sRowContents & "<TD>" & DisplayShortDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1 ,-1) & "</TD>"
						sRowContents = sRowContents & "<TD></TD>"
						sRowContents = sRowContents & "<TD></TD>"
					sRowContents = sRowContents & "</TR>"
					lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
				    oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				oRecordset.Close
				sRowContents = "</TABLE>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
			End If
		End If

		lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
		If lErrorNumber = 0 Then
			Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
			sErrorDescription = "No se pudo guardar la información del reporte."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			lErrorNumber = DeleteFolder(sFilePath, sErrorDescription)
		End If
		oEndDate = Now()
		If (lErrorNumber = 0) And B_USE_SMTP Then
			If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
		End If
	Else
		lErrorNumber = L_ERR_NO_RECORDS
		sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
		oZonesRecordset.Close
	End If

	Set oRecordset = Nothing
	BuildReport1371 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1372(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: Reporte de registros de cuentas bancarias
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1372"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim sTemp
	Dim oRecordset
	Dim sDate
	Dim sFileName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	oStartDate = Now()
	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)

	sErrorDescription = "No se pudieron eliminar los registros del repositorio temporal."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesRUSP Where (PayrollDate=" & lPayrollID & ")", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	If lErrorNumber = 0 Then
		sErrorDescription = "No se pudieron agregar los registros del repositorio temporal."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesRUSP Select Distinct EmployeesHistoryListForPayroll.EmployeeID, " & lPayrollID & " As PayrollDate, -1 As BranchID, '000' As URID, PositionID, -1 As BossPositionID, '.' As PositionName, '.' As BudgetCode, EmployeesHistoryListForPayroll.LevelID As EmployeeTypeCode, Areas.EconomicZoneID As EconomicZoneID, 0 As Concept01, 0 As Concept03, -1 As StateID, 700 As CountryID, 1 As PositionTypeID, -1 As PositionTypeID2, -EmployeesHistoryListForPayroll.EmployeeTypeID As PositionFunctionID, EmployeesHistoryListForPayroll.EmployeeTypeID, 'NULL' As PositionCode, 1 As StatusID, -1 As BirthStateID, 700 As BirthCountryID, 1 As CompanyID, 'NULL' As SEPCode, EmployeesHistoryListForPayroll.LevelID As LevelCode, -EmployeesHistoryListForPayroll.PositionTypeID As ContactTypeID, 0 As DoTax, 0 As TaxReasonID, Employees.StartDate As StartDate1, Employees.StartDate As StartDate2, Employees.StartDate As StartDate3, Employees.StartDate As StartDate4, 0 As TaxDate, 0 As EndDate, -1 As EndReasonID, 0 As bPublicService, 2 As MovementTypeID From Employees, EmployeesHistoryListForPayroll, Areas, GroupGradeLevels Where (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (EmployeesHistoryListForPayroll.GroupGradeLevelID=-1)", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		Response.Write vbNewLine & "<!-- Query: Insert Into EmployeesRUSP Select Distinct EmployeesHistoryListForPayroll.EmployeeID, " & lPayrollID & " As PayrollDate, -1 As BranchID, '000' As URID, PositionID, -1 As BossPositionID, '.' As PositionName, '.' As BudgetCode, EmployeesHistoryListForPayroll.LevelID As EmployeeTypeCode, Areas.EconomicZoneID As EconomicZoneID, 0 As Concept01, 0 As Concept03, -1 As StateID, 700 As CountryID, 1 As PositionTypeID, -1 As PositionTypeID2, -EmployeesHistoryListForPayroll.EmployeeTypeID As PositionFunctionID, EmployeesHistoryListForPayroll.EmployeeTypeID, 'NULL' As PositionCode, 1 As StatusID, -1 As BirthStateID, 700 As BirthCountryID, 1 As CompanyID, 'NULL' As SEPCode, EmployeesHistoryListForPayroll.LevelID As LevelCode, -EmployeesHistoryListForPayroll.PositionTypeID As ContactTypeID, 0 As DoTax, 0 As TaxReasonID, Employees.StartDate As StartDate1, Employees.StartDate As StartDate2, Employees.StartDate As StartDate3, Employees.StartDate As StartDate4, 0 As TaxDate, 0 As EndDate, -1 As EndReasonID, 0 As bPublicService, 2 As MovementTypeID From Employees, EmployeesHistoryListForPayroll, Areas Where (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (EmployeesHistoryListForPayroll.GroupGradeLevelID=-1) -->" & vbNewLine

		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesRUSP Select Distinct EmployeesHistoryListForPayroll.EmployeeID, " & lPayrollID & " As PayrollDate, -1 As BranchID, '000' As URID, PositionID, -1 As BossPositionID, '.' As PositionName, '.' As BudgetCode, GroupGradeLevelShortName As EmployeeTypeCode, Areas.EconomicZoneID As EconomicZoneID, 0 As Concept01, 0 As Concept03, -1 As StateID, 700 As CountryID, 1 As PositionTypeID, -1 As PositionTypeID2, -EmployeesHistoryListForPayroll.EmployeeTypeID As PositionFunctionID, EmployeesHistoryListForPayroll.EmployeeTypeID, 'NULL' As PositionCode, 1 As StatusID, -1 As BirthStateID, 700 As BirthCountryID, 1 As CompanyID, 'NULL' As SEPCode, GroupGradeLevelShortName As LevelCode, -EmployeesHistoryListForPayroll.PositionTypeID As ContactTypeID, 0 As DoTax, 0 As TaxReasonID, Employees.StartDate As StartDate1, Employees.StartDate As StartDate2, Employees.StartDate As StartDate3, Employees.StartDate As StartDate4, 0 As TaxDate, 0 As EndDate, -1 As EndReasonID, 0 As bPublicService, 2 As MovementTypeID From Employees, EmployeesHistoryListForPayroll, Areas, GroupGradeLevels Where (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (GroupGradeLevels.StartDate<=" & lForPayrollID & ") And (GroupGradeLevels.EndDate>=" & lForPayrollID & ") And (EmployeesHistoryListForPayroll.GroupGradeLevelID<>-1)", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		Response.Write vbNewLine & "<!-- Query: Insert Into EmployeesRUSP Select Distinct EmployeesHistoryListForPayroll.EmployeeID, " & lPayrollID & " As PayrollDate, -1 As BranchID, '000' As URID, PositionID, -1 As BossPositionID, '.' As PositionName, '.' As BudgetCode, GroupGradeLevelShortName As EmployeeTypeCode, Areas.EconomicZoneID As EconomicZoneID, 0 As Concept01, 0 As Concept03, -1 As StateID, 700 As CountryID, 1 As PositionTypeID, -1 As PositionTypeID2, -EmployeesHistoryListForPayroll.EmployeeTypeID As PositionFunctionID, EmployeesHistoryListForPayroll.EmployeeTypeID, 'NULL' As PositionCode, 1 As StatusID, -1 As BirthStateID, 700 As BirthCountryID, 1 As CompanyID, 'NULL' As SEPCode, GroupGradeLevelShortName As LevelCode, -EmployeesHistoryListForPayroll.PositionTypeID As ContactTypeID, 0 As DoTax, 0 As TaxReasonID, Employees.StartDate As StartDate1, Employees.StartDate As StartDate2, Employees.StartDate As StartDate3, Employees.StartDate As StartDate4, 0 As TaxDate, 0 As EndDate, -1 As EndReasonID, 0 As bPublicService, 2 As MovementTypeID From Employees, EmployeesHistoryListForPayroll, Areas Where (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (EmployeesHistoryListForPayroll.GroupGradeLevelID<>-1) -->" & vbNewLine
	End If

	If (iConnectionType = ACCESS) Or (iConnectionType = ACCESS_DSN) Then
		sErrorDescription = "No se pudieron modificar los registros del repositorio temporal."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & " Where (ConceptID In (1,14,89)) And (RecordDate=" & lForPayrollID & ") Group By EmployeeID", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				sErrorDescription = "No se pudieron modificar los registros del repositorio temporal."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesRUSP Set Concept01=" & CStr(oRecordset.Fields("TotalAmount").Value) & " Where (EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & ") And (PayrollDate=" & lPayrollID & ")", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
		End If

		sErrorDescription = "No se pudieron modificar los registros del repositorio temporal."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & " Where (ConceptID In (3,15,39,47)) And (RecordDate=" & lForPayrollID & ") Group By EmployeeID", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				sErrorDescription = "No se pudieron modificar los registros del repositorio temporal."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesRUSP Set Concept03=" & CStr(oRecordset.Fields("TotalAmount").Value) & " Where (EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & ") And (PayrollDate=" & lPayrollID & ")", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
		End If
	Else
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron modificar los registros del repositorio temporal."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesRUSP Set EmployeesRUSP.Concept01=(EmployeesRUSP.Concept01 + Payroll_" & lPayrollID & ".ConceptAmount) From EmployeesRUSP, Payroll_" & lPayrollID & " Where (EmployeesRUSP.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (EmployeesRUSP.PayrollDate=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=1) And (Payroll_" & lPayrollID & ".RecordDate=" & lForPayrollID & ")", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron modificar los registros del repositorio temporal."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesRUSP Set EmployeesRUSP.Concept01=(EmployeesRUSP.Concept01 + Payroll_" & lPayrollID & ".ConceptAmount) From EmployeesRUSP, Payroll_" & lPayrollID & " Where (EmployeesRUSP.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (EmployeesRUSP.PayrollDate=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=14) And (Payroll_" & lPayrollID & ".RecordDate=" & lForPayrollID & ")", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron modificar los registros del repositorio temporal."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesRUSP Set EmployeesRUSP.Concept01=(EmployeesRUSP.Concept01 + Payroll_" & lPayrollID & ".ConceptAmount) From EmployeesRUSP, Payroll_" & lPayrollID & " Where (EmployeesRUSP.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (EmployeesRUSP.PayrollDate=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=89) And (Payroll_" & lPayrollID & ".RecordDate=" & lForPayrollID & ")", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If

		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron modificar los registros del repositorio temporal."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesRUSP Set EmployeesRUSP.Concept03=(EmployeesRUSP.Concept03 + Payroll_" & lPayrollID & ".ConceptAmount) From EmployeesRUSP, Payroll_" & lPayrollID & " Where (EmployeesRUSP.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (EmployeesRUSP.PayrollDate=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=3) And (Payroll_" & lPayrollID & ".RecordDate=" & lForPayrollID & ")", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron modificar los registros del repositorio temporal."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesRUSP Set EmployeesRUSP.Concept03=(EmployeesRUSP.Concept03 + Payroll_" & lPayrollID & ".ConceptAmount) From EmployeesRUSP, Payroll_" & lPayrollID & " Where (EmployeesRUSP.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (EmployeesRUSP.PayrollDate=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=15) And (Payroll_" & lPayrollID & ".RecordDate=" & lForPayrollID & ")", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron modificar los registros del repositorio temporal."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesRUSP Set EmployeesRUSP.Concept03=(EmployeesRUSP.Concept03 + Payroll_" & lPayrollID & ".ConceptAmount) From EmployeesRUSP, Payroll_" & lPayrollID & " Where (EmployeesRUSP.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (EmployeesRUSP.PayrollDate=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=39) And (Payroll_" & lPayrollID & ".RecordDate=" & lForPayrollID & ")", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron modificar los registros del repositorio temporal."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesRUSP Set EmployeesRUSP.Concept03=(EmployeesRUSP.Concept03 + Payroll_" & lPayrollID & ".ConceptAmount) From EmployeesRUSP, Payroll_" & lPayrollID & " Where (EmployeesRUSP.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (EmployeesRUSP.PayrollDate=" & lPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID=47) And (Payroll_" & lPayrollID & ".RecordDate=" & lForPayrollID & ")", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron modificar los registros del repositorio temporal."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesRUSP Set EmployeesRUSP.PositionID=Jobs.PositionID From EmployeesRUSP, Employees, Jobs Where (EmployeesRUSP.EmployeeID=Employees.EmployeeID) And (Employees.JobID=Jobs.JobID) And (EmployeesRUSP.PayrollDate=" & lPayrollID & ")", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron modificar los registros del repositorio temporal."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesRUSP Set EmployeesRUSP.BossPositionID=ParentJobs.PositionID From EmployeesRUSP, Employees, Jobs, PositionsHierarchy, Jobs As ParentJobs Where (EmployeesRUSP.EmployeeID=Employees.EmployeeID) And (Employees.JobID=Jobs.JobID) And (Jobs.JobID=PositionsHierarchy.JobID) And (PositionsHierarchy.JobID=ParentJobs.JobID) And (EmployeesRUSP.PayrollDate=" & lPayrollID & ")", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	If lErrorNumber = 0 Then
		sErrorDescription = "No se pudieron modificar los registros del repositorio temporal."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesRUSP Set Concept01=Concept01*2, Concept03=Concept03*2 Where (PayrollDate=" & lPayrollID & ")", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	If lErrorNumber = 0 Then
		sErrorDescription = "No se pudieron modificar los registros del repositorio temporal."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesRUSP Set PositionFunctionID=1 Where (PayrollDate=" & lPayrollID & ") And (PositionFunctionID In (-1))", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If
	If lErrorNumber = 0 Then
		sErrorDescription = "No se pudieron modificar los registros del repositorio temporal."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesRUSP Set PositionFunctionID=2 Where (PayrollDate=" & lPayrollID & ") And (PositionFunctionID In (-3,-4))", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If
	If lErrorNumber = 0 Then
		sErrorDescription = "No se pudieron modificar los registros del repositorio temporal."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesRUSP Set PositionFunctionID=3 Where (PayrollDate=" & lPayrollID & ") And (PositionFunctionID In (-2,-7))", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If
	If lErrorNumber = 0 Then
		sErrorDescription = "No se pudieron modificar los registros del repositorio temporal."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesRUSP Set PositionFunctionID=4 Where (PayrollDate=" & lPayrollID & ") And (PositionFunctionID In (0,-5,-6))", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	If lErrorNumber = 0 Then
		sErrorDescription = "No se pudieron modificar los registros del repositorio temporal."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesRUSP Set ContactTypeID=1 Where (PayrollDate=" & lPayrollID & ") And (ContactTypeID In (-2))", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If
	If lErrorNumber = 0 Then
		sErrorDescription = "No se pudieron modificar los registros del repositorio temporal."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesRUSP Set ContactTypeID=2 Where (PayrollDate=" & lPayrollID & ") And (ContactTypeID In (-1))", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If
	If lErrorNumber = 0 Then
		sErrorDescription = "No se pudieron modificar los registros del repositorio temporal."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesRUSP Set ContactTypeID=3 Where (PayrollDate=" & lPayrollID & ") And (ContactTypeID In (-4,-5))", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If
	If lErrorNumber = 0 Then
		sErrorDescription = "No se pudieron modificar los registros del repositorio temporal."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesRUSP Set ContactTypeID=4 Where (PayrollDate=" & lPayrollID & ") And (ContactTypeID In (-3))", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	If lErrorNumber = 0 Then
		sErrorDescription = "No se pudieron modificar los registros del repositorio temporal."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesRUSP Set DoTax=1, TaxReasonID=1 Where (PayrollDate=" & lPayrollID & ") And (EmployeeTypeID In (1))", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	If lErrorNumber = 0 Then
		sErrorDescription = "No se pudieron obtener los registros de los empleados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesRUSP.*, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, EmployeeEmail, SocialSecurityNumber, BirthDate, RFC, CURP, GenderID, PositionShortName, PositionLongName, Zones1.ZoneCode From EmployeesRUSP, Employees, EmployeesHistoryListForPayroll, Positions, Areas, Zones As Zones1, Zones As Zones2, Zones As Zones3 Where (EmployeesRUSP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones3.ZoneID) And (Zones3.ParentID=Zones2.ZoneID) And (Zones2.ParentID=Zones1.ZoneID) And (EmployeesRUSP.PayrollDate=" & lPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones3.StartDate<=" & lForPayrollID & ") And (Zones3.EndDate>=" & lForPayrollID & ") Order By EmployeesHistoryListForPayroll.EmployeeNumber", "ReportsQueries1300Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select EmployeesRUSP.*, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, EmployeeEmail, SocialSecurityNumber, BirthDate, RFC, CURP, GenderID, PositionShortName, PositionLongName, Zones1.ZoneCode From EmployeesRUSP, Employees, EmployeesHistoryListForPayroll, Positions, Areas, Zones As Zones1, Zones As Zones2, Zones As Zones3 Where (EmployeesRUSP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones3.ZoneID) And (Zones3.ParentID=Zones2.ZoneID) And (Zones2.ParentID=Zones1.ZoneID) And (EmployeesRUSP.PayrollDate=" & lPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Zones3.StartDate<=" & lForPayrollID & ") And (Zones3.EndDate>=" & lForPayrollID & ") Order By EmployeesHistoryListForPayroll.EmployeeNumber -->" & vbNewLine
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sDate = GetSerialNumberForDate("")
				sFileName = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls")
				Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Replace(sFileName, ".xls", ".zip") & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
				Response.Flush()

'				lErrorNumber = AppendTextToFile(sFileName, "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">", sErrorDescription)
					asColumnsTitles = "Ramo,Unidad responsable,Consecutivo único del puesto,Consecutivo puesto del jefe,Nombre del puesto,Código presupuestal,Nivel tabular autorizado,Zona económica,Sueldo base tabular,Compensación garantizada,Entidad federativa de la plaza,País de la plaza,Tipo de plaza,Tipo de puesto estratégico,Tipo de función del puesto,Tipo de personal,Código de puesto RHNet,Estatus ocupacional,RFC-SP,CURP,Nombre(s),Primer apellido,Segundo apellido,Fecha de nacimiento,Sexo,Entidad federativa de nacimiento,País de nacimiento,e-mail,Institución de seguridad social,Númeo de seguridad social,Clave SEP presupuestal,Nivel tabular pagado,Tipo de contratación,Declaración patrimonial,Motivo de obligación declaración patrimonial,Número de empleado,Ingreso a la APF,Ingreso al SPC,Ingreso a la institución,Alta al último puesto,Obligación a presentar declaración patrimonial"
'					asColumnsTitles = Split(asColumnsTitles, ",", -1, vbBinaryCompare)
'					asCellWidths = Split("100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
'					lErrorNumber = AppendTextToFile(sFileName, GetTableHeaderPlainText(asColumnsTitles, True, ""), sErrorDescription)
					lErrorNumber = AppendTextToFile(sFileName, Replace(asColumnsTitles, ",", vbTab), sErrorDescription)

					lCurrentID = -2
					bDisplay = False
'					asCellAlignments = Split(",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,", ",", -1, vbBinaryCompare)
					Do While Not oRecordset.EOF
						sRowContents = "=T(""" & CleanStringForHTML(Right(("00" & CStr(oRecordset.Fields("BranchID").Value)), Len("00"))) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(Right(("000" & CStr(oRecordset.Fields("URID").Value)), Len("000"))) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(Right(("0000000" & CStr(oRecordset.Fields("PositionID").Value)), Len("0000000"))) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(Right(("0000000" & CStr(oRecordset.Fields("BossPositionID").Value)), Len("0000000"))) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionLongName").Value))
						'sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("BudgetCode").Value)) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeCode").Value)) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("EconomicZoneID").Value)) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & FormatNumber(CDbl(oRecordset.Fields("Concept01").Value), 2, True, False, False) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & FormatNumber(CDbl(oRecordset.Fields("Concept03").Value), 2, True, False, False) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(Right(("00" & CStr(oRecordset.Fields("ZoneCode").Value)), Len("00"))) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(Right(("000" & CStr(oRecordset.Fields("CountryID").Value)), Len("000"))) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("PositionTypeID").Value)) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR '& "=T(""" & CleanStringForHTML(Right(("00" & CStr(oRecordset.Fields("PositionTypeID2").Value)), Len("00"))) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionFunctionID").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(Right(("0" & CStr(oRecordset.Fields("EmployeeTypeID").Value)), Len("0"))) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionCode").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("StatusID").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR
							sTemp = ""
							sTemp = CStr(oRecordset.Fields("RFC").Value)
							Err.Clear
							sRowContents = sRowContents & CleanStringForHTML(Right(("             " & sTemp), Len("             ")))
						sRowContents = sRowContents & TABLE_SEPARATOR
							sTemp = ""
							sTemp = CStr(oRecordset.Fields("CURP").Value)
							Err.Clear
							sRowContents = sRowContents & CleanStringForHTML(Right(("                  " & sTemp), Len("                  ")))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(Trim(Replace(CStr(oRecordset.Fields("EmployeeName").Value), "Ñ", "#")))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(Trim(Replace(CStr(oRecordset.Fields("EmployeeLastName").Value), "Ñ", "#")))
						sRowContents = sRowContents & TABLE_SEPARATOR
							sTemp = " "
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then sTemp = CStr(oRecordset.Fields("EmployeeLastName2").Value)
							Err.Clear
							If Len(sTemp) > 0 Then
								sRowContents = sRowContents & CleanStringForHTML(Trim(Replace(sTemp, "Ñ", "#")))
							Else
								sRowContents = sRowContents & "NULL"
							End If
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("BirthDate").Value)) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(Replace(CStr(oRecordset.Fields("GenderID").Value), "0", "2")) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T("""
							sTemp = ""
							sTemp = CStr(oRecordset.Fields("CURP").Value)
							sTemp = sTemp & "                  "
							Err.Clear
							Select Case Mid(sTemp, 12, 2)
								Case "AS"
									sRowContents = sRowContents & "01"
								Case "BC"
									sRowContents = sRowContents & "02"
								Case "BS"
									sRowContents = sRowContents & "03"
								Case "CC"
									sRowContents = sRowContents & "04"
								Case "CH"
									sRowContents = sRowContents & "08"
								Case "CL"
									sRowContents = sRowContents & "05"
								Case "CM"
									sRowContents = sRowContents & "06"
								Case "CS"
									sRowContents = sRowContents & "07"
								Case "DF"
									sRowContents = sRowContents & "09"
								Case "DG"
									sRowContents = sRowContents & "10"
								Case "GR"
									sRowContents = sRowContents & "12"
								Case "GT"
									sRowContents = sRowContents & "11"
								Case "HG"
									sRowContents = sRowContents & "13"
								Case "JC"
									sRowContents = sRowContents & "14"
								Case "MC"
									sRowContents = sRowContents & "15"
								Case "MN"
									sRowContents = sRowContents & "16"
								Case "MS"
									sRowContents = sRowContents & "17"
								Case "NE"
									sRowContents = sRowContents & "33"
								Case "NL"
									sRowContents = sRowContents & "19"
								Case "NT"
									sRowContents = sRowContents & "18"
								Case "OC"
									sRowContents = sRowContents & "20"
								Case "PL"
									sRowContents = sRowContents & "21"
								Case "QR"
									sRowContents = sRowContents & "23"
								Case "QT"
									sRowContents = sRowContents & "22"
								Case "SL"
									sRowContents = sRowContents & "25"
								Case "SP"
									sRowContents = sRowContents & "24"
								Case "SR"
									sRowContents = sRowContents & "26"
								Case "TC"
									sRowContents = sRowContents & "27"
								Case "TL"
									sRowContents = sRowContents & "29"
								Case "TS"
									sRowContents = sRowContents & "28"
								Case "VZ"
									sRowContents = sRowContents & "30"
								Case "YN"
									sRowContents = sRowContents & "31"
								Case "ZS"
									sRowContents = sRowContents & "32"
								Case Else
									sRowContents = sRowContents & "00"
							End Select
						sRowContents = sRowContents & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T("""
							Select Case Mid(sTemp, 12, 2)
								Case "NE"
									sRowContents = sRowContents & "600"
								Case Else
									sRowContents = sRowContents & "700"
							End Select
						sRowContents = sRowContents & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T("""
							sTemp = ""
							sTemp = CStr(oRecordset.Fields("EmployeeEmail").Value)
							Err.Clear
							If Len(sTemp) > 0 Then
								sRowContents = sRowContents & CleanStringForHTML(sTemp)
							Else
								sRowContents = sRowContents & "siapisssteweb@issste.gob.mx"
							End If
						sRowContents = sRowContents & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("CompanyID").Value)) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T("""
							sTemp = ""
							sTemp = CStr(oRecordset.Fields("SocialSecurityNumber").Value)
							Err.Clear
							If Len(sTemp) = 11 Then
								sRowContents = sRowContents & CleanStringForHTML(sTemp)
							Else
								sRowContents = sRowContents & "NULL"
							End If
						sRowContents = sRowContents & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("SEPCode").Value)) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("LevelCode").Value)) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("ContactTypeID").Value)) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T("""
							If CInt(oRecordset.Fields("DoTax").Value) = 0 Then
								sRowContents = sRowContents & "N"
							Else
								sRowContents = sRowContents & "S"
							End If
						sRowContents = sRowContents & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T("""
							If CInt(oRecordset.Fields("DoTax").Value) = 0 Then
								sRowContents = sRowContents & "NULL"
							Else
								sRowContents = sRowContents & CleanStringForHTML(Right(("0" & CStr(oRecordset.Fields("TaxReasonID").Value)), Len("0")))
							End If
						sRowContents = sRowContents & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate2").Value)) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T("""
							If CLng(oRecordset.Fields("StartDate3").Value) = 0 Then
								sRowContents = sRowContents & "NULL"
							Else
								sRowContents = sRowContents & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate3").Value))
							End If
						sRowContents = sRowContents & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate1").Value)) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate4").Value)) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T("""
							If CInt(oRecordset.Fields("DoTax").Value) = 0 Then
								sRowContents = sRowContents & "NULL"
							Else
								sRowContents = sRowContents & DisplayNumericDateFromSerialNumber(CStr(oRecordset.Fields("TaxDate").Value))
							End If
						sRowContents = sRowContents & """)"

'						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
'						lErrorNumber = AppendTextToFile(sFileName, GetTableRowText(asRowContents, True, ""), sErrorDescription)
						lErrorNumber = AppendTextToFile(sFileName, Replace(sRowContents, TABLE_SEPARATOR, vbTab), sErrorDescription)

						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
'				lErrorNumber = AppendTextToFile(sFileName, "</TABLE>", sErrorDescription)

				lErrorNumber = ZipFolder(sFileName, Replace(sFileName, ".xls", ".zip"), sErrorDescription)
				If lErrorNumber = 0 Then
					Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
					sErrorDescription = "No se pudieron guardar la información del reporte."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
				If lErrorNumber = 0 Then
					lErrorNumber = DeleteFile(sFileName, sErrorDescription)
				End If
				oEndDate = Now()
				If (lErrorNumber = 0) And B_USE_SMTP Then
					If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(Replace(sFileName, ".xls", ".zip"), CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
				End If
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros en el sistema que cumplan con los criterios del filtro."
			End If
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1372 = lErrorNumber
	Err.Clear
End Function
%>