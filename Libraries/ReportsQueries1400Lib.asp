<%
Function BuildReport1400(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To save the CLCs into the employees payroll
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1400"
	Dim sCondition
	Dim sCondition1
	Dim sQuery
	Dim sQueryT
	Dim sTables
	Dim lPayrollID
	Dim lForPayrollID
	Dim oRecordset
	Dim asPeriods
	Dim lPayrollCLC
	Dim lPayrollCode
	Dim iIndex
    Dim lPCLCs

	Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
	sQueryT = "Update PayrollsCLCs Set PayrollCode = ' ', PayrollCLC = ' ', FilterParameters = ' ' Where PayrollCLC Like 'T_%'"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQueryT, "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	sTables = "EmployeesHistoryListForPayroll, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, Zones As ZonesForPaymentCenter, Zones As ZonesForPaymentCenter2, Zones As ParentZonesForPaymentCenter"
	sCondition1 = "  And (EmployeesHistoryListForPayroll.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=ZonesForPaymentCenter.ZoneID) And (ZonesForPaymentCenter.ParentID=ZonesForPaymentCenter2.ZoneID) And (ZonesForPaymentCenter2.ParentID=ParentZonesForPaymentCenter.ZoneID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (PaymentCenters.StartDate<=" & lForPayrollID & ") And (PaymentCenters.EndDate>=" & lForPayrollID & ")"
	sCondition = Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.")
	sCondition = Replace(sCondition, "Zones.ZonePath", "ZonesForPaymentCenter.ZonePath")
	sCondition = Replace(sCondition, "EmployeeTypes.", "EmployeesHistoryListForPayroll.")
	sCondition = Replace(sCondition, "Companies.", "EmployeesHistoryListForPayroll.")
	sCondition = Replace(sCondition, "Banks.", "EmployeesHistoryListForPayroll.")
	sCondition = Replace(sCondition, "ParentZones.", "ParentZonesForPaymentCenter.")
	If StrComp(oRequest("CheckConceptID").Item, "69", vbBinaryCompare) = 0 Then
		sTables = "EmployeesHistoryListForPayroll, Zones As ZonesForPaymentCenter"
		sCondition1 = " And (EmployeesHistoryListForPayroll.ZoneID=ZonesForPaymentCenter.ZoneID) "
		sCondition = sCondition & " And (EmployeesHistoryListForPayroll.EmployeeID In (Select EmployeeID From EmployeesBeneficiariesLKP Where (StartDate <= "& oRequest("PayrollID").Item &") And (EndDate >= " & oRequest("PayrollID").Item & ")))"
	End If
	sQuery = ""
	If Len(oRequest("PayrollCode").Item) > 0 Then 
		lPayrollCode = oRequest("PayrollCode").Item
	Else
		If StrComp(oRequest("ReportStep").Item, "3", vbBinaryCompare) = 0 Then
			asPeriods = Split("0,1,1,2,2,3,3,4,4,5,5,6,6",",")
			lPayrollCode = Mid(oRequest("PayrollID").Item,1,4) & "0" & asPeriods(Mid(oRequest("PayrollID").Item,5,2))
		End If
	End If
	sQuery = sQuery & "PayrollCode='" & lPayrollCode & "', "
	If Len(oRequest("PayrollCLC").Item) > 0 Then 
		lPayrollCLC = oRequest("PayrollCLC").Item
	Else
		lPayrollCLC = "T_" & oRequest("PayrollID").Item
	End If
	sQueryT = "Select Count(*) Total From PayrollsCLCs Where PayrollCLC = '" & lPayrollCLC &"'"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQueryT, "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If oRecordset.Fields("Total").Value = 0 Then
		sQuery = sQuery & "PayrollCLC='" & lPayrollCLC & "', "
		If Len(sQuery) > 0 Then
			sErrorDescription = "No se pudo obtener la información de la nómina."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update PayrollsCLCs Set " & sQuery & "FilterParameters='" & Replace(sCondition, "'", """") & "' Where (PayrollID=" & lPayrollID & ") And (EmployeeID In (Select EmployeesHistoryListForPayroll.EmployeeID From " & sTables & " Where (PayrollID=" & lPayrollID & ")" & sCondition1 & sCondition & "))", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			Response.Write vbNewLine & "<!-- Query: Update PayrollsCLCs Set " & sQuery & "FilterParameters='" & Replace(sCondition, "'", """") & "' Where (PayrollID=" & lPayrollID & ") And (EmployeeID In (Select EmployeesHistoryListForPayroll.EmployeeID From EmployeesHistoryListForPayroll, Areas As PaymentCenters, Zones Where (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (PayrollID=" & lPayrollID & ") " & sCondition & ")) -->" & vbNewLine
			If lErrorNumber = 0 Then
				sQuery = "Select Count(PayrollCLC) As Regs From PayrollsCLCs Where (PayrollCLC = '" & lPayrollCLC & "') And (PayrollCode = '" & lPayrollCode & "')"
				sErrorDescription = "No se encontraron registros con los criterios indicados"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If CLng(oRecordset.Fields("Regs").Value) = 0 Then
						sErrorDescription = "No se encontraron registros con los criterios indicados"
						lErrorNumber = -1
                    'Else
                    'sQuery = "select payrollclc,payrollcode from payrollsCLCs where payrollID = "&lPayrollID&" group by payrollCLC,payrollcode;"
                    'lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
                    'lPCLCs = oRecordset.GetRows()
                    'For iIndex = 0 To UBound(lPCLCs,2)
                        'lErrorNumber =  prc_actualiza_clc(oRequest, oADODBConnection,lPCLCs(0,iIndex), lPCLCs(1,iIndex), lPayrollID, sErrorDescription)
                    'Next
					End If
				End If
			End If
		End If
	End If

	BuildReport1400 = lErrorNumber
	Err.Clear
End Function


Function BuildReport1400b(oRequest, oADODBConnection, lPayrollID, sErrorDescription)
'************************************************************
'Purpose: To display the CLCs for the given payroll
'Inputs:  oRequest, oADODBConnection, lPayrollID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1400b"
	Dim alConcepts
	Dim asCLCs
	Dim asCondition
	Dim asPeriods
	Dim sBanks
	Dim sCompanies
	Dim iIndex
	Dim jIndex
	Dim sCondition
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim sQuery
	Dim sBankNAme
	Dim sCompanyName
	Dim lPeriod
	Dim bFirstRow
    Dim lPayrollCLC
    Dim lPayrollCode

	bFirstRow =  True
	asPeriods = Split("0,1,1,2,2,3,3,4,4,5,5,6,6",",")
	If StrComp(oRequest("ReportStep").Item, "4", vbBinaryCompare) = 0 Then
		sCondition = " And (PayrollCLC='" & oRequest("PayrollCLC").Item & "')"
		lPeriod = CLng(oRequest("PayrollCode").Item)
	Else
		lPeriod = Mid(lPayrollID, 1, 4) & "0" & asPeriods(CInt(Mid(lPayrollID, 5, 2)))
	End If
	asCLCs = ""
	sErrorDescription = "No se pudieron obtener las CLCs generadas para la nómina especificada."
	If aReportsComponent(N_STEP_REPORTS) = 3 Then
		sQuery = "Select PayrollCLC, PayrollCode, FilterParameters, Count(PayrollCLC) As Regs From PayrollsCLCs Where (PayrollCode='" & lPeriod & "') And (PayrollCLC Like 'T_" & S_WILD_CHAR & "') Group By PayrollCLC, PayrollCode, FilterParameters Order By PayrollCLC"
	Else
		sQuery = "Select PayrollCLC, PayrollCode, FilterParameters, Count(PayrollCLC) As Regs From PayrollsCLCs Where (PayrollCode='" & lPeriod & "') Group By PayrollCLC, PayrollCode, FilterParameters Order By PayrollCLC"
	End If
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: " & sQuery & " -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			asCLCs = oRecordset.GetRows()
			Response.Write "<TABLE WIDTH=""1100"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			asColumnsTitles = Split("CLC,Unidad,Qna,Banco,F-Pago,CP,Registros,Percepciones,Deducciones,Líquido,", ",", -1, vbBinaryCompare)
			asCellWidths = Split("100,100,100,100,100,100,100,100,100,100,50,", ",", -1, vbBinaryCompare)
			If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
				lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			Else
				lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			End If
			asCellAlignments = Split(",,CENTER,,RIGHT,CENTER,RIGHT,RIGHT,RIGHT,RIGHT,CENTER", ",", -1, vbBinaryCompare)
				For iIndex = 0 To UBound(asCLCs,2)
                    lPayrollCLC = asCLCs(0,iIndex)
                    lPayrollCode = asCLCs(1,iIndex)
                    'lErrorNumber =  prc_actualiza_clc(oRequest, oADODBConnection, lPayrollCLC, lPayrollCode, sErrorDescription)
                    If lErrorNumber = 0 Then 
					    sQuery = "Select ConceptName, Sum(ConceptAmount) Importe From Payroll_" & lPayrollID & " Pr, Concepts C Where EmployeeID In (Select EmployeeID From PayrollsCLCs Where PayrollCLC = '" & asCLCs(0,iIndex) & "') And Pr.ConceptID In (-2,-1,0) And (Pr.ConceptID = C.ConceptID) Group By ConceptName Order By ConceptName Desc"
					    lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					    alConcepts = oRecordset.GetRows()

					    sQuery = "Select Distinct BankID From EmployeesHistoryListForPayroll EHLFP, (Select EmployeeID From PayrollsCLCs Where PayrollCLC = '" & asCLCs(0,iIndex) & "') Pr Where (EHLFP.EmployeeID = Pr.EmployeeID) And (PayrollID = " & lPayrollID & ")"
					    lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					    sBanks =""
					    Do While Not oRecordset.EOF
						    Call GetNameFromTable(oADODBConnection, "Banks", oRecordset.Fields("BankID").Value, "", "", sBankName, "")
						    sBanks = sBanks & "," & sBankName
						    oRecordset.MoveNext
					    Loop
					    sQuery = "Select Distinct CompanyID From EmployeesHistoryListForPayroll EHL, (Select EmployeeID From PayrollsCLCs Where (PayrollCLC = '" & asCLCs(0,iIndex) & "')) Pr Where (EHl.EmployeeID = Pr.EmployeeID) And (PayrollID = " & lPayrollID & ")"
					    lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					    sCompanies = ""
					    Do While Not oRecordset.EOF
						    Call GetNameFromTable(oADODBConnection, "Companies", oRecordset.Fields("CompanyID").Value, "", "", sCompanyName, "")
						    sCompanies = sCompanies & "," & sCompanyName
						    oRecordset.MoveNext
					    Loop
					    sRowContents = asCLCs(0,iIndex)
					    sRowContents = sRowContents & TABLE_SEPARATOR & sCompanies
					    sRowContents = sRowContents & TABLE_SEPARATOR & asCLCs(1,iIndex)
					    sRowContents = sRowContents & TABLE_SEPARATOR & sBanks

					    If InStr(1,asCLCs(2,iIndex),"AccountNumber",vbBinaryCompare) = 0 Then
						    sRowContents = sRowContents & TABLE_SEPARATOR & "Depósito y cheque"
					    Else
						    If InStr(1,asCLCs(2,iIndex),"AccountNumber<>",vbBinaryCompare) > 0 Then
							    sRowContents = sRowContents & TABLE_SEPARATOR & "Depósito bancario"
						    ElseIf InStr(1,asCLCs(2,iIndex),"AccountNumber=",vbBinaryCompare) > 0 Then
							    sRowContents = sRowContents & TABLE_SEPARATOR & "Cheque"
						    End If
					    End If
					    If InStr(1,asCLCs(2,iIndex),",9,",vbBinaryCompare) = 0 Then
						    sRowContents = sRowContents & TABLE_SEPARATOR & "Foráneo"
					    Else
						    sRowContents = sRowContents & TABLE_SEPARATOR & "Local"
					    End If
					    sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(asCLCs(3,iIndex)), 0, True, True, True)
					    sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(alConcepts(1,0)), 2, True, True, True)
					    sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(alConcepts(1,2)), 2, True, True, True)
					    sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(alConcepts(1,1)), 2, True, True, True)
                    
					    If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_DELETE_PERMISSIONS) = N_DELETE_PERMISSIONS Then
						    sRowContents = sRowContents & TABLE_SEPARATOR  & "<A HREF=""Catalogs.asp?Action=PayrollsClcs&PayrollCLC=" & CStr(asCLCs(0,iIndex)) & "&PayrollID=" & lPayrollID & "&Remove=1"">"
						    sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
					    End If
					    asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					    lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
                    End If
				Next
			Response.Write "</TABLE>" & vbNewLine
		End If
	End If
   End Function

   
Function BuildReport1400b1(oRequest, oADODBConnection, lPayrollID, sErrorDescription)
'************************************************************
'Purpose: To display the CLCs for the given payroll
'Inputs:  oRequest, oADODBConnection, lPayrollID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1400b"
	Dim alConcepts
	Dim asCLCs
	Dim asCondition
	Dim asPeriods
	Dim sBanks
	Dim sCompanies
	Dim iIndex
	Dim jIndex
	Dim sCondition
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim sQuery
	Dim sBankNAme
	Dim sCompanyName
	Dim lPeriod
	Dim bFirstRow
    Dim lPayrollCLC
    Dim lPayrollCode
    'Variables request
    Dim asEmployeeNumbers
    Dim asCompanyID
    Dim asAreaID
    Dim asSubAreaID
    Dim asEmployeeTypeID
    Dim asPaymentCenterID
    Dim asZoneID
    Dim asBankID
    Dim asCheckConcept
    Dim sExtraConditions
    Dim scad
    Dim sExtraTables
    Dim a
    Dim lPayrollTypeID
    Dim lPayrollDescription
    Dim lMemorandum
    Dim lPayrollFile
    Dim lCancelDate
    Dim lPayrollYear
    Dim lPayrollMonth
    Dim lFortNightly

	bFirstRow =  True
	asPeriods = Split("0,1,1,2,2,3,3,4,4,5,5,6,6",",")
    lPayrollTypeID =oRequest("PayrollTypeID").Item
	If StrComp(oRequest("ReportStep").Item, "4", vbBinaryCompare) = 0 Then
		sCondition = " And (PayrollCLC='" & oRequest("PayrollCLC").Item & "')"
		lPeriod = oRequest("PayrollCode").Item
        lPayrollCLC = oRequest("PayrollCLC").Item
        lPayrollDescription = oRequest("PayrollDescription").Item
        lMemorandum = oRequest("Memorandum").Item
        lPayrollFile = oRequest("FileCLC").Item
        lCancelDate = oRequest("CancelYearCLC").Item
        lPayrollYear = oRequest("YearCLC").Item
        lPayrollMonth = oRequest("MonthCLC").Item
        lFortNightly = oRequest("QuarterCLC").Item
        lPayrollTypeID =oRequest("PayrollTypeIDCLC").Item
        
	Else
		lPeriod = Mid(lPayrollID, 1, 4) & "0" & asPeriods(CInt(Mid(lPayrollID, 5, 2)))
	End If
	asCLCs = ""
    sExtraConditions = ""
    sExtraTables = ""   

   If len(oRequest("EmployeeNumbers").Item)>0 Then
         asEmployeeNumbers = split(GetParameterFromURLString(oRequest, "EmployeeNumbers"),"%0D%0A")
         scad = asEmployeeNumbers(0)
         For iIndex = 1 To UBound(asEmployeeNumbers)     
            scad = scad&","&asEmployeeNumbers(iIndex)
         Next
         sExtraConditions = sExtraConditions & " And EHL.EmployeeID In ("& scad &")" 
    End If

    If len(oRequest("EmployeeTypeID").Item)>0 Then
         asEmployeeTypeID = split(GetParameterFromURLString(oRequest, "EmployeeTypeID"),",")
         scad = asEmployeeTypeID(0)
         For iIndex = 1 To UBound(asEmployeeTypeID)     
            scad = scad&","&asEmployeeTypeID(iIndex)
         Next
         sExtraConditions = sExtraConditions & " And EHL.EmployeeTypeID In ("& scad &")" 
    End If
    If len(oRequest("CompanyID").Item)>0 Then
         asCompanyID = split(GetParameterFromURLString(oRequest, "CompanyID"),",")        
         sExtraConditions = sExtraConditions & "  And ((EHL.CompanyID = "&asCompanyID(0)&") "
         'scad = asCompanyID(0)
         'scad = ""
         sCompanies = ""
         'If UBound(asCompanyID) =0 Then
             'Call GetNameFromTable(oADODBConnection, "Companies", asCompanyID(0), "", "", sCompanyName, "")
			 'sCompanies = sCompanies & "," & sCompanyName
             'sExtraConditions = sExtraConditions & " Or (EHL.CompanyID = ("& scad &")"
         'Else
             For iIndex = 1 To UBound(asCompanyID)+1     
            'scad = scad&","&asCompanyID(iIndex)
                sExtraConditions = sExtraConditions & " Or (EHL.CompanyID = "&asCompanyID(iIndex)&")"      
			    Call GetNameFromTable(oADODBConnection, "Companies", asCompanyID(iIndex-1), "", "", sCompanyName, "")
			    sCompanies = sCompanies & "," & sCompanyName
            Next
            sExtraConditions = sExtraConditions &")"      
         'End If
          
    End If

    If len(oRequest("AreaID").Item)>0 Then
        sExtraTables = sExtraTables & " ,Areas A"
        If Len(oRequest("SubAreaID").Item)>1 Then
            asAreaID = split(GetParameterFromURLString(oRequest, "SubAreaID"),"%2C")
        Else
            asAreaID = split(GetParameterFromURLString(oRequest, "AreaID"),",")
        End If
        sExtraConditions = sExtraConditions & "  And (( A.AreaPath like'%,"& asAreaID(0)&",%') "
        For iIndex = 1 To UBound(asAreaID)           
           sExtraConditions = sExtraConditions & " Or ( A.AreaPath like'%,"& asAreaID(iIndex)&",%')"      
        Next
        sExtraConditions = sExtraConditions & ") "
    End If
    If len(oRequest("ZoneID").Item)>0 Then
        sExtraTables = sExtraTables & " ,Zones Z"
        asZoneID = split(GetParameterFromURLString(oRequest, "ZoneID"),",")       
        sExtraConditions = sExtraConditions & " And(( Z.ZonePath like'%,"& asZoneID(0)&",%')"
        For iIndex = 1 To UBound(asZoneID)   
           sExtraConditions = sExtraConditions & " Or ( Z.ZonePath like'%,"& asZoneID(iIndex)&",%')"      
        Next
        sExtraConditions = sExtraConditions & " )"
    End If
    If len(oRequest("BankID").Item)>0 Then
         sBanks =""
         asBankID = split(GetParameterFromURLString(oRequest, "BankID"),",")        
         'scad = asBankID(0)
         sExtraConditions = sExtraConditions & "  And ((EHL.BankID = "&asBankID(0)&") "
         For iIndex = 1 To UBound(asBankID)+1     
            'scad = scad&","&asBankID(iIndex)
            sExtraConditions = sExtraConditions & " Or (EHL.BankID = "&asBankID(iIndex)&")"      
            Call GetNameFromTable(oADODBConnection, "Banks", asBankID(iIndex-1), "", "", sBankName, "")
            sBanks = sBanks & "," & sBankName
         Next
         sExtraConditions = sExtraConditions & ")" 
    End If
    If len(oRequest("CheckConceptID").Item)>0 Then
        If StrComp(oRequest("CheckConceptID").Item, "-1", vbBinaryCompare) = 0 Then
            sExtraConditions = sExtraConditions&" And (EHL.AccountNumber='.')"
            asCheckConcept = "Cheque"
        Else 
            asCheckConcept = "Depósito bancario"
            sExtraConditions = sExtraConditions &" And (EHL.AccountNumber Not Like '.')" 
        End If
    Else
         asCheckConcept = "Depósito y cheque"
    End If
    If Len(oRequest("PayrollTypeID").Item) > 0 Then
       sExtraConditions = sExtraConditions&" And (CLC.PayrollTypeID = "& lPayrollTypeID &") "       
    End If

	sErrorDescription = "No se pudieron obtener las CLCs generadas para la nómina especificada."
	If aReportsComponent(N_STEP_REPORTS) = 3 Then
        sQuery = "Select PayrollCLC, PayrollCode, Count(*)  Regs From EmployeesHistoryListForPayroll EHL, PayrollsCLCs CLC, Zones Z, Areas A Where (EHL.PayrollID = "& lPayrollID &") And (PayrollCode = '" & lPeriod & "') And (EHL.PayrollID = CLC.PayrollID) And (EHL.EmployeeID = CLC.EmployeeID) And (EHL.AreaID = A.AreaID) And (EHL.ZoneID = Z.ZoneID) "& sExtraConditions &" Group by PayrollCLC, PayrollCode"
	Else
        lErrorNumber = UpdateCLC_RPT(oRequest, oADODBConnection, lPayrollCLC, lPeriod,lPayrollID,lPayrollTypeID,lPayrollDescription,lMemorandum,lPayrollFile, lCancelDate, lPayrollYear, lPayrollMonth, lFortNightly, sExtraConditions, sErrorDescription)
        sQuery = "Select PayrollCLC, PayrollCode, Count(PayrollCLC) As Regs From PayrollsCLCs Where (PayrollCLC='" & lPayrollCLC & "') And PayrollId = "& lPayrollID &" Group By PayrollCLC, PayrollCode Order By PayrollCLC"
	End If
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: " & sQuery & " -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			asCLCs = oRecordset.GetRows()
			Response.Write "<TABLE WIDTH=""1100"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			asColumnsTitles = Split("CLC,Unidad,Qna,Banco,F-Pago,CP,Registros,Percepciones,Deducciones,Líquido,", ",", -1, vbBinaryCompare)
			asCellWidths = Split("100,100,100,100,100,100,100,100,100,100,50,", ",", -1, vbBinaryCompare)
			If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
				lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			Else
				lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			End If
			asCellAlignments = Split(",,CENTER,,RIGHT,CENTER,RIGHT,RIGHT,RIGHT,RIGHT,CENTER", ",", -1, vbBinaryCompare)
				For iIndex = 0 To UBound(asCLCs,2)                   
                    If lErrorNumber = 0 Then 
                        If aReportsComponent(N_STEP_REPORTS) = 3 Then
                            'sQuery = "Select ConceptID, Sum(ConceptAmount) Importe From Payroll_" & Mid(lPayrollID,1,4) & " Pr Where RecordDate =" & lPayrollID & " And EmployeeID In (Select EHL.EmployeeID  From EmployeesHistoryListForPayroll EHL, PayrollsCLCs CLC , Zones Z, Areas A Where (EHL.PayrollID = " & lPayrollID & ") And (PayrollCode =  '"& asCLCs(1,iIndex) &"') And (PayrollCLC = '"& asCLCs(0,iIndex) &"') And (EHL.PayrollID = CLC.PayrollID) And (EHL.EmployeeID = CLC.EmployeeID) And (EHL.AreaID = A.AreaID) And (EHL.ZoneID = Z.ZoneID) "&sExtraConditions&") And Pr.ConceptID In (-2,-1,0) Group By ConceptID Order by ConceptID"
							sQuery = "Select ConceptID, Sum(ConceptAmount) Importe From suma_cptos_payroll_" & Mid(lPayrollID,1,4) & " Pr Where RecordDate =" & lPayrollID & " And EmployeeID In (Select EHL.EmployeeID  From EmployeesHistoryListForPayroll EHL, PayrollsCLCs CLC , Zones Z, Areas A Where (EHL.PayrollID = " & lPayrollID & ") And (PayrollCode =  '"& asCLCs(1,iIndex) &"') And (PayrollCLC = '"& asCLCs(0,iIndex) &"') And (EHL.PayrollID = CLC.PayrollID) And (EHL.EmployeeID = CLC.EmployeeID) And (EHL.AreaID = A.AreaID) And (EHL.ZoneID = Z.ZoneID) "&sExtraConditions&") And Pr.ConceptID In (-2,-1,0) Group By ConceptID Order by ConceptID"
		                    'sQuery = "Select ConceptID, Sum(ConceptAmount) Importe From Payroll_" & lPayrollID & "  Pr, EmployeesHistoryListForPayroll EHL, PayrollsCLCs CLC "&sExtraTables&" Where Pr.EmployeeID = EHL.EmployeeID And (EHL.EmployeeID = CLC.EmployeeID) And (EHL.PayrollID = " & lPayrollID & " ) And (EHL.PayrollID = CLC.PayrollID) And (PayrollCode = '" & lPeriod & "') And Pr.ConceptID In (-2,-1,0) "&sExtraConditions& " Group By ConceptID Order by ConceptID"
                            'sQuery = "Select ConceptID, Sum(ConceptAmount) Importe From Payroll_" & lPayrollID & "  Pr Where RecordDate =" & lPayrollID & " And EmployeeID In ( Select EHL.EmployeeID  From EmployeesHistoryListForPayroll EHL, PayrollsCLCs CLC, Zones Z, Areas A Where (EHL.PayrollID = " & lPayrollID & ") And (PayrollCode =  '"& lPeriod &"') And (EHL.PayrollID = CLC.PayrollID) And (EHL.EmployeeID = CLC.EmployeeID) And (EHL.AreaID = A.AreaID) And (EHL.ZoneID = Z.ZoneID) "&sExtraConditions& ") And Pr.ConceptID In (-2,-1,0) Group By ConceptID Order by ConceptID"
                        Else
							sQuery = "Select ConceptID, Sum(ConceptAmount) Importe From suma_cptos_payroll_" & Mid(lPayrollID,1,4) & " Pr, PayrollsCLCs CLC Where CLC.PayrollCLC = '"& lPayrollCLC &"' And Pr.EmployeeID = CLC.EmployeeID And Pr.RecordDate = CLC.PayrollID And CLC.payrollID ='"& lPayrollID &"' And CLC.PayrollCode = '"& lPeriod &"' And Pr.ConceptID In (-2,-1,0) Group By ConceptID Order by ConceptID"
                        End If

					    lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					    alConcepts = oRecordset.GetRows()

					    sRowContents = asCLCs(0,iIndex)
					    sRowContents = sRowContents & TABLE_SEPARATOR & sCompanies
					    sRowContents = sRowContents & TABLE_SEPARATOR & asCLCs(1,iIndex)
					    sRowContents = sRowContents & TABLE_SEPARATOR & sBanks
                        sRowContents = sRowContents & TABLE_SEPARATOR & asCheckConcept

					   
					    If InStr(1,asZoneID(0),"9",vbBinaryCompare) = 0 Then
						    sRowContents = sRowContents & TABLE_SEPARATOR & "Foráneo"
					    Else
						    sRowContents = sRowContents & TABLE_SEPARATOR & "Local"
					    End If
					    sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(asCLCs(2,iIndex)), 0, True, True, True)
					    sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(alConcepts(1,1)), 2, True, True, True)
					    sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(alConcepts(1,0)), 2, True, True, True)
					    sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(alConcepts(1,2)), 2, True, True, True)
                    
					    If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_DELETE_PERMISSIONS) = N_DELETE_PERMISSIONS Then
						    sRowContents = sRowContents & TABLE_SEPARATOR  & "<A HREF=""Catalogs.asp?Action=PayrollsClcs&PayrollCLC=" & CStr(asCLCs(0,iIndex)) & "&PayrollID=" & lPayrollID & "&Remove=1"">"
						    sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
					    End If
					    asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					    lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
                    End If
				Next
			Response.Write "</TABLE>" & vbNewLine
		End If
	End If
   End Function


Function BuildReport1401(oRequest, oADODBConnection, sConceptIDs, sErrorDescription)
'************************************************************
'Purpose: To display the payroll for the given concepts
'         as a pipe-separated text file
'Inputs:  oRequest, oADODBConnection, sConceptIDs
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1401"
	Dim sQueryBegin
    Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim bPayrollIsClosed
	Dim lPayrollNumber
	Dim iConceptCounter
	Dim iIndex
	Dim sDate
	Dim sFilePath
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim sTemp
	Dim sCurrentID
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

	Dim sEmployeeNumber
	Dim sEmployeeName
	Dim sEmployeeLastName 
	Dim sEmployeeLastName2
	Dim sRFC
	Dim sGroupGradeLevelShortName
	Dim sIntegrationID
	Dim sCheckNumber
	Dim sAccountNumber
	Dim sPositionShortName
	Dim sLevelShortName
	Dim sJourneyShortName
	Dim	sShiftShortName
	Dim sPayrollCLC

	dim sTipo
	dim sCompanyID
	dim sBANKSHORTNAME
	dim sEmployeeTypeshortname
	dim sZonePath
	dim sPayrollID
	dim sYearA
	dim sMonthA 
	dim sQuincena
	dim sTabulador

	sQueryBegin = ""
    Call GetConditionFromURL(oRequest, sCondition, lPayrollID, lForPayrollID)
    If InStr(1, sCondition, " And (EmployeesHistoryList.JobID=Jobs.JobID)", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Jobs"
	sCondition = Replace(Replace(Replace(sCondition, "Banks.", "BankAccounts."), "Companies.", "EmployeesHistoryList."), "EmployeeTypes.", "EmployeesHistoryList.")
	lPayrollNumber = (CInt(Left(lForPayrollID, Len("0000"))) * 100) + CInt(GetPayrollNumber(lForPayrollID))
	oStartDate = Now()
	sDate = GetSerialNumberForDate("")
	sFilePath = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".txt"
	Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(Replace(sFilePath, ".txt", ".zip")) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
	sFilePath = Server.MapPath(sFilePath)
	Response.Flush()

	Call IsPayrollClosed(oADODBConnection, lPayrollID, sCondition, bPayrollIsClosed, sErrorDescription)

	If Len(sConceptIDs) > 0 Then sCondition = sCondition & " And (Concepts.ConceptID In (" & sConceptIDs & "))"
	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
'		sCondition = sCondition & " And ((EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (EmployeesHistoryList.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
		sCondition = sCondition & " And (EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))"
	End If
	sErrorDescription = "No se pudieron obtener los montos pagados."
	If Len(oRequest("TotalByArea").Item) > 0 Then
		If StrComp(oRequest("CheckConceptID").Item, "69", vbBinaryCompare) = 0 Then
		ElseIf StrComp(oRequest("CheckConceptID").Item, "155", vbBinaryCompare) = 0 Then
		Else
			If bPayrollIsClosed Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Areas.AreaID As EmployeeID, '' As EmployeeNumber, '' As EmployeeName, '' As EmployeeLastName, '' As EmployeeLastName2, '' As RFC, '1' As EmployeeTypeID, CompanyName, '' As GroupGradeLevelShortName, '' As IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, '' As CheckNumber, '' As AccountNumber, PayrollsCLCs.PayrollCLC, '' As PositionShortName, '' As LevelShortName, '' As JourneyShortName, '' As ShiftShortName, Min(RecordDate) As MinRecordDate, Max(RecordDate) As MaxRecordDate, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, Payrolls, PayrollsCLCs, Employees, EmployeesHistoryListForPayroll, Companies, Areas, Positions, Levels, Journeys, Shifts, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones" & sQueryBegin & " Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.JourneyID=Journeys.JourneyID) And (EmployeesHistoryListForPayroll.ShiftID=Shifts.ShiftID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (Payrolls.PayrollID=" & lPayrollID & ") And (PayrollsCLCs.PayrollID=" & lPayrollID & ") And (PayrollsCLCs.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Journeys.StartDate<=" & lForPayrollID & ") And (Journeys.EndDate>=" & lForPayrollID & ") And (Shifts.StartDate<=" & lForPayrollID & ") And (Shifts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID>0) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By Areas.AreaID, CompanyName, Areas.AreaShortName, ConceptShortName, IsDeduction, PayrollsCLCs.PayrollCLC Order By Areas.AreaShortName, Concepts.IsDeduction, ConceptShortName", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select Areas.AreaID As EmployeeID, '' As EmployeeNumber, '' As EmployeeName, '' As EmployeeLastName, '' As EmployeeLastName2, '' As RFC, '1' As EmployeeTypeID, CompanyName, '' As GroupGradeLevelShortName, '' As IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, '' As CheckNumber, '' As AccountNumber, PayrollsCLCs.PayrollCLC, '' As PositionShortName, '' As LevelShortName, '' As JourneyShortName, '' As ShiftShortName, Min(RecordDate) As MinRecordDate, Max(RecordDate) As MaxRecordDate, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, Payrolls, PayrollsCLCs, Employees, EmployeesHistoryListForPayroll, Companies, Areas, Positions, Levels, Journeys, Shifts, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones" & sQueryBegin & " Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.JourneyID=Journeys.JourneyID) And (EmployeesHistoryListForPayroll.ShiftID=Shifts.ShiftID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (Payrolls.PayrollID=" & lPayrollID & ") And (PayrollsCLCs.PayrollID=" & lPayrollID & ") And (PayrollsCLCs.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Journeys.StartDate<=" & lForPayrollID & ") And (Journeys.EndDate>=" & lForPayrollID & ") And (Shifts.StartDate<=" & lForPayrollID & ") And (Shifts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID>0) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By Areas.AreaID, CompanyName, Areas.AreaShortName, ConceptShortName, IsDeduction, PayrollsCLCs.PayrollCLC Order By Areas.AreaShortName, Concepts.IsDeduction, ConceptShortName -->" & vbNewLine
			Else
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Areas.AreaID As EmployeeID, '' As EmployeeNumber, '' As EmployeeName, '' As EmployeeLastName, '' As EmployeeLastName2, '' As RFC, '1' As EmployeeTypeID, CompanyName, '' As GroupGradeLevelShortName, '' As IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, '' As CheckNumber, '' As AccountNumber, PayrollsCLCs.PayrollCLC, '' As PositionShortName, '' As LevelShortName, '' As JourneyShortName, '' As ShiftShortName, Min(RecordDate) As MinRecordDate, Max(RecordDate) As MaxRecordDate, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, Payrolls, PayrollsCLCs, BankAccounts, Employees, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas, Positions, Levels, Journeys, Shifts, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones" & sQueryBegin & " Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (Payrolls.PayrollID=" & lPayrollID & ") And (PayrollsCLCs.PayrollID=" & lPayrollID & ") And (PayrollsCLCs.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Journeys.StartDate<=" & lForPayrollID & ") And (Journeys.EndDate>=" & lForPayrollID & ") And (Shifts.StartDate<=" & lForPayrollID & ") And (Shifts.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (Concepts.ConceptID>0) " & sCondition & " Group By Areas.AreaID, CompanyName, Areas.AreaShortName, ConceptShortName, IsDeduction, PayrollsCLCs.PayrollCLC Order By Areas.AreaShortName, Concepts.IsDeduction, ConceptShortName", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select Areas.AreaID As EmployeeID, '' As EmployeeNumber, '' As EmployeeName, '' As EmployeeLastName, '' As EmployeeLastName2, '' As RFC, '1' As EmployeeTypeID, CompanyName, '' As GroupGradeLevelShortName, '' As IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, '' As CheckNumber, '' As AccountNumber, PayrollsCLCs.PayrollCLC, '' As PositionShortName, '' As LevelShortName, '' As JourneyShortName, '' As ShiftShortName, Min(RecordDate) As MinRecordDate, Max(RecordDate) As MaxRecordDate, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, Payrolls, PayrollsCLCs, BankAccounts, Employees, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas, Positions, Levels, Journeys, Shifts, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones" & sQueryBegin & " Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (Payrolls.PayrollID=" & lPayrollID & ") And (PayrollsCLCs.PayrollID=" & lPayrollID & ") And (PayrollsCLCs.EmployeeID=Payroll_" & lPayrollID & ".EmployeeID) And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Journeys.StartDate<=" & lForPayrollID & ") And (Journeys.EndDate>=" & lForPayrollID & ") And (Shifts.StartDate<=" & lForPayrollID & ") And (Shifts.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (Concepts.ConceptID>0) " & sCondition & " Group By Areas.AreaID, CompanyName, Areas.AreaShortName, ConceptShortName, IsDeduction, PayrollsCLCs.PayrollCLC Order By Areas.AreaShortName, Concepts.IsDeduction, ConceptShortName -->" & vbNewLine
			End If
			If oRecordset.EOF Then
				oRecordset.Close
				If bPayrollIsClosed Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Areas.AreaID As EmployeeID, '' As EmployeeNumber, '' As EmployeeName, '' As EmployeeLastName, '' As EmployeeLastName2, '' As RFC, '1' As EmployeeTypeID, CompanyName, '' As GroupGradeLevelShortName, '' As IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, '' As CheckNumber, '' As AccountNumber, '' As PayrollCLC, '' As PositionShortName, '' As LevelShortName, '' As JourneyShortName, '' As ShiftShortName, Min(RecordDate) As MinRecordDate, Max(RecordDate) As MaxRecordDate, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, Payrolls, Employees, EmployeesHistoryListForPayroll, Companies, Areas, Positions, Levels, Journeys, Shifts, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones" & sQueryBegin & " Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.JourneyID=Journeys.JourneyID) And (EmployeesHistoryListForPayroll.ShiftID=Shifts.ShiftID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (Payrolls.PayrollID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Journeys.StartDate<=" & lForPayrollID & ") And (Journeys.EndDate>=" & lForPayrollID & ") And (Shifts.StartDate<=" & lForPayrollID & ") And (Shifts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID>0) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By Areas.AreaID, CompanyName, Areas.AreaShortName, ConceptShortName, IsDeduction Order By Areas.AreaShortName, Concepts.IsDeduction, ConceptShortName", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Areas.AreaID As EmployeeID, '' As EmployeeNumber, '' As EmployeeName, '' As EmployeeLastName, '' As EmployeeLastName2, '' As RFC, '1' As EmployeeTypeID, CompanyName, '' As GroupGradeLevelShortName, '' As IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, '' As CheckNumber, '' As AccountNumber, '' As PayrollCLC, '' As PositionShortName, '' As LevelShortName, '' As JourneyShortName, '' As ShiftShortName, Min(RecordDate) As MinRecordDate, Max(RecordDate) As MaxRecordDate, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, Payrolls, Employees, EmployeesHistoryListForPayroll, Companies, Areas, Positions, Levels, Journeys, Shifts, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones" & sQueryBegin & " Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.JourneyID=Journeys.JourneyID) And (EmployeesHistoryListForPayroll.ShiftID=Shifts.ShiftID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (Payrolls.PayrollID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Journeys.StartDate<=" & lForPayrollID & ") And (Journeys.EndDate>=" & lForPayrollID & ") And (Shifts.StartDate<=" & lForPayrollID & ") And (Shifts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID>0) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " Group By Areas.AreaID, CompanyName, Areas.AreaShortName, ConceptShortName, IsDeduction Order By Areas.AreaShortName, Concepts.IsDeduction, ConceptShortName -->" & vbNewLine
				Else
					'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Areas.AreaID As EmployeeID, '' As EmployeeNumber, '' As EmployeeName, '' As EmployeeLastName, '' As EmployeeLastName2, '' As RFC, '1' As EmployeeTypeID, CompanyName, '' As GroupGradeLevelShortName, '' As IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, '' As CheckNumber, '' As AccountNumber, '' As PayrollCLC, '' As PositionShortName, '' As LevelShortName, '' As JourneyShortName, '' As ShiftShortName, Min(RecordDate) As MinRecordDate, Max(RecordDate) As MaxRecordDate, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, Payrolls, BankAccounts, Employees, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas, Positions, Levels, Journeys, Shifts, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones" & sQueryBegin & " Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (Payrolls.PayrollID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Journeys.StartDate<=" & lForPayrollID & ") And (Journeys.EndDate>=" & lForPayrollID & ") And (Shifts.StartDate<=" & lForPayrollID & ") And (Shifts.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (Concepts.ConceptID>0) " & sCondition & " Group By Areas.AreaID, CompanyName, Areas.AreaShortName, ConceptShortName, IsDeduction Order By Areas.AreaShortName, Concepts.IsDeduction, ConceptShortName", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
                    lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Areas.AreaID As EmployeeID, '' As EmployeeNumber, '' As EmployeeName, '' As EmployeeLastName, '' As EmployeeLastName2, '' As RFC, '1' As EmployeeTypeID, CompanyName, '' As GroupGradeLevelShortName, '' As IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, '' As CheckNumber, '' As AccountNumber, '' As PayrollCLC, '' As PositionShortName, '' As LevelShortName, '' As JourneyShortName, '' As ShiftShortName, Min(RecordDate) As MinRecordDate, Max(RecordDate) As MaxRecordDate, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, Payrolls, BankAccounts, Employees, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas, Positions, Levels, Journeys, Shifts, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones" & sQueryBegin & " Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (Payrolls.PayrollID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Journeys.StartDate<=" & lForPayrollID & ") And (Journeys.EndDate>=" & lForPayrollID & ") And (Shifts.StartDate<=" & lForPayrollID & ") And (Shifts.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (EmployeesChangesLKP.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesChangesLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (Concepts.ConceptID>0) " & sCondition & " Group By Areas.AreaID, CompanyName, Areas.AreaShortName, ConceptShortName, IsDeduction Order By Areas.AreaShortName, Concepts.IsDeduction, ConceptShortName", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Response.Write vbNewLine & "<!-- Query: Select Areas.AreaID As EmployeeID, '' As EmployeeNumber, '' As EmployeeName, '' As EmployeeLastName, '' As EmployeeLastName2, '' As RFC, '1' As EmployeeTypeID, CompanyName, '' As GroupGradeLevelShortName, '' As IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, '' As CheckNumber, '' As AccountNumber, '' As PayrollCLC, '' As PositionShortName, '' As LevelShortName, '' As JourneyShortName, '' As ShiftShortName, Min(RecordDate) As MinRecordDate, Max(RecordDate) As MaxRecordDate, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, Payrolls, BankAccounts, Employees, EmployeesChangesLKP, EmployeesHistoryList, Companies, Areas, Positions, Levels, Journeys, Shifts, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones" & sQueryBegin & " Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (Payrolls.PayrollID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Journeys.StartDate<=" & lForPayrollID & ") And (Journeys.EndDate>=" & lForPayrollID & ") And (Shifts.StartDate<=" & lForPayrollID & ") And (Shifts.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (EmployeesChangesLKP.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesChangesLKP.PayrollDate=Payroll_" & lPayrollID & ".RecordDate) And (Concepts.ConceptID>0) " & sCondition & " Group By Areas.AreaID, CompanyName, Areas.AreaShortName, ConceptShortName, IsDeduction Order By Areas.AreaShortName, Concepts.IsDeduction, ConceptShortName -->" & vbNewLine
				End If
			End If
		End If
	ElseIf bPayrollIsClosed Then
		If StrComp(oRequest("CheckConceptID").Item, "69", vbBinaryCompare) = 0 Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryListForPayroll.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, '' As CheckNumber, AccountNumber, '' As PayrollCLC, PositionShortName, LevelShortName, JourneyShortName, ShiftShortName, Min(RecordDate) As MinRecordDate, Max(RecordDate) As MaxRecordDate, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", Concepts, Payrolls, BankAccounts, Employees, EmployeesHistoryListForPayroll, GroupGradeLevels, Companies, Areas, Positions, Levels, Journeys, Shifts, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (EmployeesBeneficiariesLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.JourneyID=Journeys.JourneyID) And (EmployeesHistoryListForPayroll.ShiftID=Shifts.ShiftID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (Payrolls.PayrollID=" & lPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Journeys.StartDate<=" & lForPayrollID & ") And (Journeys.EndDate>=" & lForPayrollID & ") And (Shifts.StartDate<=" & lForPayrollID & ") And (Shifts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID>0) " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By EmployeesBeneficiariesLKP.BeneficiaryNumber, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryListForPayroll.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, AccountNumber, PositionShortName, LevelShortName, JourneyShortName, ShiftShortName Order By EmployeesHistoryListForPayroll.EmployeeNumber, Concepts.IsDeduction, ConceptShortName", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryListForPayroll.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, '' As CheckNumber, AccountNumber, '' As PayrollCLC, PositionShortName, LevelShortName, JourneyShortName, ShiftShortName, Min(RecordDate) As MinRecordDate, Max(RecordDate) As MaxRecordDate, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", Concepts, Payrolls, BankAccounts, Employees, EmployeesHistoryListForPayroll, GroupGradeLevels, Companies, Areas, Positions, Levels, Journeys, Shifts, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (EmployeesBeneficiariesLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.JourneyID=Journeys.JourneyID) And (EmployeesHistoryListForPayroll.ShiftID=Shifts.ShiftID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (Payrolls.PayrollID=" & lPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Journeys.StartDate<=" & lForPayrollID & ") And (Journeys.EndDate>=" & lForPayrollID & ") And (Shifts.StartDate<=" & lForPayrollID & ") And (Shifts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID>0) " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By EmployeesBeneficiariesLKP.BeneficiaryNumber, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryListForPayroll.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, AccountNumber, PositionShortName, LevelShortName, JourneyShortName, ShiftShortName Order By EmployeesHistoryListForPayroll.EmployeeNumber, Concepts.IsDeduction, ConceptShortName -->" & vbNewLine
		ElseIf StrComp(oRequest("CheckConceptID").Item, "155", vbBinaryCompare) = 0 Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesCreditorsLKP.CreditorNumber As EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryListForPayroll.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, '' As CheckNumber, AccountNumber, '' As PayrollCLC, PositionShortName, LevelShortName, JourneyShortName, ShiftShortName, Min(RecordDate) As MinRecordDate, Max(RecordDate) As MaxRecordDate, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", Concepts, Payrolls, BankAccounts, Employees, EmployeesHistoryListForPayroll, GroupGradeLevels, Companies, Areas, Positions, Levels, Journeys, Shifts, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (EmployeesCreditorsLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.JourneyID=Journeys.JourneyID) And (EmployeesHistoryListForPayroll.ShiftID=Shifts.ShiftID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (Payrolls.PayrollID=" & lPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Journeys.StartDate<=" & lForPayrollID & ") And (Journeys.EndDate>=" & lForPayrollID & ") And (Shifts.StartDate<=" & lForPayrollID & ") And (Shifts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID>0) " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By EmployeesCreditorsLKP.CreditorNumber, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryListForPayroll.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, AccountNumber, PositionShortName, LevelShortName, JourneyShortName, ShiftShortName Order By EmployeesHistoryListForPayroll.EmployeeNumber, Concepts.IsDeduction, ConceptShortName", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select EmployeesCreditorsLKP.CreditorNumber As EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryListForPayroll.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, '' As CheckNumber, AccountNumber, '' As PayrollCLC, PositionShortName, LevelShortName, JourneyShortName, ShiftShortName, Min(RecordDate) As MinRecordDate, Max(RecordDate) As MaxRecordDate, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", Concepts, Payrolls, BankAccounts, Employees, EmployeesHistoryListForPayroll, GroupGradeLevels, Companies, Areas, Positions, Levels, Journeys, Shifts, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (EmployeesCreditorsLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.JourneyID=Journeys.JourneyID) And (EmployeesHistoryListForPayroll.ShiftID=Shifts.ShiftID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (Payrolls.PayrollID=" & lPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Journeys.StartDate<=" & lForPayrollID & ") And (Journeys.EndDate>=" & lForPayrollID & ") And (Shifts.StartDate<=" & lForPayrollID & ") And (Shifts.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID>0) " & Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll.") & " Group By EmployeesCreditorsLKP.CreditorNumber, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryListForPayroll.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, AccountNumber, PositionShortName, LevelShortName, JourneyShortName, ShiftShortName Order By EmployeesHistoryListForPayroll.EmployeeNumber, Concepts.IsDeduction, ConceptShortName -->" & vbNewLine
		Else 'ARCHIVOS SPEP GAGR'
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryListForPayroll.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, '' As CheckNumber, AccountNumber, '' As PayrollCLC, PositionShortName, LevelShortName, JourneyShortName, '' As ShiftShortName, Companies.CompanyID, BANKS.BANKSHORTNAME, EmployeeTypes.EmployeeTypeshortname, Zones.zonepath,  Payrolls.PayrollID, Payrolls.PayrollTypeID, Min(RecordDate) As MinRecordDate, Max(RecordDate) As MaxRecordDate, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ServiceShortName From Payroll_" & lPayrollID & ", Concepts, Payrolls, Employees, EmployeesHistoryListForPayroll, GroupGradeLevels, Companies, Areas, Positions, Levels, Journeys, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, banks, employeetypes, services  " & sQueryBegin & " Where (EmployeesHistoryListForPayroll.serviceid = services.serviceid) AND (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.JourneyID=Journeys.JourneyID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (Payrolls.PayrollID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Journeys.StartDate<=" & lForPayrollID & ") And (Journeys.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID>0) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " AND (banks.bankid = EmployeesHistoryListForPayroll.bankid) and (EmployeeTypes.EmployeeTypeid = EmployeesHistoryListForPayroll.EmployeeTypeID) Group By EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryListForPayroll.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, AccountNumber, PositionShortName, LevelShortName, JourneyShortName, Companies.CompanyID,  BANKS.BANKSHORTNAME,EmployeeTypes.EmployeeTypeshortname,Zones.zonepath, Payrolls.PayrollID, Payrolls.PayrollTypeID, serviceshortname Order By EmployeesHistoryListForPayroll.EmployeeNumber, Concepts.IsDeduction, ConceptShortName", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryListForPayroll.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, '' As CheckNumber, AccountNumber, '' As PayrollCLC, PositionShortName, LevelShortName, JourneyShortName, '' As ShiftShortName, Companies.CompanyID, BANKS.BANKSHORTNAME, EmployeeTypes.EmployeeTypeshortname, Zones.zonepath,  Payrolls.PayrollID, Payrolls.PayrollTypeID, Min(RecordDate) As MinRecordDate, Max(RecordDate) As MaxRecordDate, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount, ServiceShortName From Payroll_" & lPayrollID & ", Concepts, Payrolls, Employees, EmployeesHistoryListForPayroll, GroupGradeLevels, Companies, Areas, Positions, Levels, Journeys, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones, banks, employeetypes, services  " & sQueryBegin & " Where (EmployeesHistoryListForPayroll.serviceid = services.serviceid) AND (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryListForPayroll.JourneyID=Journeys.JourneyID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (Payrolls.PayrollID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Journeys.StartDate<=" & lForPayrollID & ") And (Journeys.EndDate>=" & lForPayrollID & ") And (Concepts.ConceptID>0) " & Replace(Replace(sCondition, "EmployeesHistoryList.", "EmployeesHistoryListForPayroll."), "BankAccounts.", "EmployeesHistoryListForPayroll.") & " AND (banks.bankid = EmployeesHistoryListForPayroll.bankid) and (EmployeeTypes.EmployeeTypeid = EmployeesHistoryListForPayroll.EmployeeTypeID) Group By EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryListForPayroll.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryListForPayroll.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, AccountNumber, PositionShortName, LevelShortName, JourneyShortName, Companies.CompanyID,  BANKS.BANKSHORTNAME,EmployeeTypes.EmployeeTypeshortname,Zones.zonepath, Payrolls.PayrollID, Payrolls.PayrollTypeID Order By EmployeesHistoryListForPayroll.EmployeeNumber, Concepts.IsDeduction, ConceptShortName -->" & vbNewLine
		End If
	Else
		If StrComp(oRequest("CheckConceptID").Item, "69", vbBinaryCompare) = 0 Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryList.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryList.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, '' As CheckNumber, AccountNumber, '' As PayrollCLC, PositionShortName, LevelShortName, JourneyShortName, ShiftShortName, Min(RecordDate) As MinRecordDate, Max(RecordDate) As MaxRecordDate, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", Concepts, Payrolls, BankAccounts, Employees, EmployeesChangesLKP, EmployeesHistoryList, GroupGradeLevels, Companies, Areas, Positions, Levels, Journeys, Shifts, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (EmployeesBeneficiariesLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (Payrolls.PayrollID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Journeys.StartDate<=" & lForPayrollID & ") And (Journeys.EndDate>=" & lForPayrollID & ") And (Shifts.StartDate<=" & lForPayrollID & ") And (Shifts.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Concepts.ConceptID>0) " & sCondition & " Group By EmployeesBeneficiariesLKP.BeneficiaryNumber, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryList.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryList.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, AccountNumber, PositionShortName, LevelShortName, JourneyShortName, ShiftShortName Order By EmployeesHistoryList.EmployeeNumber, Concepts.IsDeduction, ConceptShortName", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryList.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryList.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, '' As CheckNumber, AccountNumber, '' As PayrollCLC, PositionShortName, LevelShortName, JourneyShortName, ShiftShortName, Min(RecordDate) As MinRecordDate, Max(RecordDate) As MaxRecordDate, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & lPayrollID & ", Concepts, Payrolls, BankAccounts, Employees, EmployeesChangesLKP, EmployeesHistoryList, GroupGradeLevels, Companies, Areas, Positions, Levels, Journeys, Shifts, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones Where (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (EmployeesBeneficiariesLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (Payrolls.PayrollID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.StartDate<=" & lForPayrollID & ") And (EmployeesBeneficiariesLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Journeys.StartDate<=" & lForPayrollID & ") And (Journeys.EndDate>=" & lForPayrollID & ") And (Shifts.StartDate<=" & lForPayrollID & ") And (Shifts.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Concepts.ConceptID>0) " & sCondition & " Group By EmployeesBeneficiariesLKP.BeneficiaryNumber, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryList.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryList.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, AccountNumber, PositionShortName, LevelShortName, JourneyShortName, ShiftShortName Order By EmployeesHistoryList.EmployeeNumber, Concepts.IsDeduction, ConceptShortName -->" & vbNewLine
		ElseIf StrComp(oRequest("CheckConceptID").Item, "155", vbBinaryCompare) = 0 Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesCreditorsLKP.CreditorNumber As EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryList.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryList.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, '' As CheckNumber, AccountNumber, '' As PayrollCLC, PositionShortName, LevelShortName, JourneyShortName, ShiftShortName, Min(RecordDate) As MinRecordDate, Max(RecordDate) As MaxRecordDate, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", Concepts, Payrolls, BankAccounts, Employees, EmployeesChangesLKP, EmployeesHistoryList, GroupGradeLevels, Companies, Areas, Positions, Levels, Journeys, Shifts, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (EmployeesCreditorsLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (Payrolls.PayrollID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Journeys.StartDate<=" & lForPayrollID & ") And (Journeys.EndDate>=" & lForPayrollID & ") And (Shifts.StartDate<=" & lForPayrollID & ") And (Shifts.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Concepts.ConceptID>0) " & sCondition & " Group By EmployeesCreditorsLKP.CreditorNumber, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryList.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryList.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, AccountNumber, PositionShortName, LevelShortName, JourneyShortName, ShiftShortName Order By EmployeesHistoryList.EmployeeNumber, Concepts.IsDeduction, ConceptShortName", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select EmployeesCreditorsLKP.CreditorNumber As EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryList.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryList.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, '' As CheckNumber, AccountNumber, '' As PayrollCLC, PositionShortName, LevelShortName, JourneyShortName, ShiftShortName, Min(RecordDate) As MinRecordDate, Max(RecordDate) As MaxRecordDate, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & lPayrollID & ", Concepts, Payrolls, BankAccounts, Employees, EmployeesChangesLKP, EmployeesHistoryList, GroupGradeLevels, Companies, Areas, Positions, Levels, Journeys, Shifts, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones Where (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (EmployeesCreditorsLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (Payrolls.PayrollID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (EmployeesCreditorsLKP.StartDate<=" & lForPayrollID & ") And (EmployeesCreditorsLKP.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Journeys.StartDate<=" & lForPayrollID & ") And (Journeys.EndDate>=" & lForPayrollID & ") And (Shifts.StartDate<=" & lForPayrollID & ") And (Shifts.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Concepts.ConceptID>0) " & sCondition & " Group By EmployeesCreditorsLKP.CreditorNumber, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryList.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryList.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, AccountNumber, PositionShortName, LevelShortName, JourneyShortName, ShiftShortName Order By EmployeesHistoryList.EmployeeNumber, Concepts.IsDeduction, ConceptShortName -->" & vbNewLine
		Else
			'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryList.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryList.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, '' As CheckNumber, AccountNumber, '' As PayrollCLC, PositionShortName, LevelShortName, JourneyShortName, ShiftShortName, Min(RecordDate) As MinRecordDate, Max(RecordDate) As MaxRecordDate, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, Payrolls, BankAccounts, Employees, EmployeesChangesLKP, EmployeesHistoryList, GroupGradeLevels, Companies, Areas, Positions, Levels, Journeys, Shifts, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones" & sQueryBegin & " Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (Payrolls.PayrollID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Journeys.StartDate<=" & lForPayrollID & ") And (Journeys.EndDate>=" & lForPayrollID & ") And (Shifts.StartDate<=" & lForPayrollID & ") And (Shifts.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=" & lForPayrollID & ") And (BankAccounts.EndDate>=" & lForPayrollID & ") And (BankAccounts.Active=1) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (Concepts.ConceptID>0) " & sCondition & " Group By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryList.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryList.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, AccountNumber, PositionShortName, LevelShortName, JourneyShortName, ShiftShortName Order By EmployeesHistoryList.EmployeeNumber, Concepts.IsDeduction, ConceptShortName", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
            lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryList.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryList.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, '' As CheckNumber, AccountNumber, '' As PayrollCLC, PositionShortName, LevelShortName, JourneyShortName, ShiftShortName, Min(RecordDate) As MinRecordDate, Max(RecordDate) As MaxRecordDate, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, Payrolls, BankAccounts, Employees, EmployeesChangesLKP, EmployeesHistoryList, GroupGradeLevels, Companies, Areas, Positions, Levels, Journeys, Shifts, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones" & sQueryBegin & " Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (Payrolls.PayrollID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Journeys.StartDate<=" & lForPayrollID & ") And (Journeys.EndDate>=" & lForPayrollID & ") And (Shifts.StartDate<=" & lForPayrollID & ") And (Shifts.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.EndDate>=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.Active=1) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (Concepts.ConceptID>0) " & sCondition & " Group By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryList.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryList.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, AccountNumber, PositionShortName, LevelShortName, JourneyShortName, ShiftShortName Order By EmployeesHistoryList.EmployeeNumber, Concepts.IsDeduction, ConceptShortName", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryList.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryList.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, '' As CheckNumber, AccountNumber, '' As PayrollCLC, PositionShortName, LevelShortName, JourneyShortName, ShiftShortName, Min(RecordDate) As MinRecordDate, Max(RecordDate) As MaxRecordDate, Sum(Payroll_" & lPayrollID & ".ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, Payrolls, BankAccounts, Employees, EmployeesChangesLKP, EmployeesHistoryList, GroupGradeLevels, Companies, Areas, Positions, Levels, Journeys, Shifts, Areas As PaymentCenters, Zones, Zones As Zones2, Zones As ParentZones" & sQueryBegin & " Where (Payroll_" & lPayrollID & ".EmployeeID=BankAccounts.EmployeeID) And (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (PaymentCenters.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=ParentZones.ZoneID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (Payrolls.PayrollID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Journeys.StartDate<=" & lForPayrollID & ") And (Journeys.EndDate>=" & lForPayrollID & ") And (Shifts.StartDate<=" & lForPayrollID & ") And (Shifts.EndDate>=" & lForPayrollID & ") And (BankAccounts.StartDate<=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.EndDate>=Payroll_" & lPayrollID & ".RecordDate) And (BankAccounts.Active=1) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & CLng(Left(lPayrollID, (Len("00000000")))) & ") And (Concepts.ConceptID>0) " & sCondition & " Group By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, EmployeesHistoryList.EmployeeTypeID, CompanyName, GroupGradeLevels.GroupGradeLevelShortName, EmployeesHistoryList.IntegrationID, Areas.AreaShortName, ConceptShortName, IsDeduction, AccountNumber, PositionShortName, LevelShortName, JourneyShortName, ShiftShortName Order By EmployeesHistoryList.EmployeeNumber, Concepts.IsDeduction, ConceptShortName -->" & vbNewLine
		End If
	End If
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sCurrentID = ""
			adTotal = Split("0,0", ",")
			adTotal(0) = 0
			adTotal(1) = 0
			Do While Not oRecordset.EOF
				If Not IsNull(oRecordset.Fields("EmployeeNumber").Value) Then
					sEmployeeNumber = CStr(oRecordset.Fields("EmployeeNumber").Value)
				Else
					sEmployeeNumber = ""
				End If
				If Not IsNull(oRecordset.Fields("EmployeeName").Value) Then
					sEmployeeName = CStr(oRecordset.Fields("EmployeeName").Value)
				Else
					sEmployeeName = ""
				End If
				If Not IsNull(oRecordset.Fields("EmployeeLastName").Value) Then
					sEmployeeLastName = CStr(oRecordset.Fields("EmployeeLastName").Value)
				Else
					sEmployeeLastName = ""
				End If
				If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
					sEmployeeLastName2 = CStr(oRecordset.Fields("EmployeeLastName2").Value)
				Else
					sEmployeeLastName2 = ""
				End If
				If Not IsNull(oRecordset.Fields("RFC").Value) Then
					sRFC = CStr(oRecordset.Fields("RFC").Value)
				Else
					sRFC = ""
				End If
				If Not IsNull(oRecordset.Fields("GroupGradeLevelShortName").Value) Then
					sGroupGradeLevelShortName = CStr(oRecordset.Fields("GroupGradeLevelShortName").Value)
				Else
					sGroupGradeLevelShortName = ""
				End If
				If Not IsNull(oRecordset.Fields("IntegrationID").Value) Then
					sIntegrationID = CStr(oRecordset.Fields("IntegrationID").Value)
				Else
					sIntegrationID = ""
				End If

				If Not IsNull(oRecordset.Fields("CheckNumber").Value) Then
					sCheckNumber = CStr(oRecordset.Fields("CheckNumber").Value)
				Else
					sCheckNumber = ""
				End If
				If Not IsNull(oRecordset.Fields("AccountNumber").Value) Then
					sAccountNumber = CStr(oRecordset.Fields("AccountNumber").Value)
					if sAccountNumber = "." then 
						sTabulador = "CHEQUE"
					else 
						sTabulador  = "TD"
					end if
				Else
					sAccountNumber = ""
				End If
				If Not IsNull(oRecordset.Fields("PositionShortName").Value) Then
					sPositionShortName = CStr(oRecordset.Fields("PositionShortName").Value)
				Else
					sPositionShortName = ""
				End If
				If Not IsNull(oRecordset.Fields("LevelShortName").Value) Then
					sLevelShortName = CStr(oRecordset.Fields("LevelShortName").Value)
				Else
					sLevelShortName = ""
				End If
				If Not IsNull(oRecordset.Fields("JourneyShortName").Value) Then
					sJourneyShortName = CStr(oRecordset.Fields("JourneyShortName").Value)
				Else
					sJourneyShortName = ""
				End If
				If Not IsNull(oRecordset.Fields("ShiftShortName").Value) Then
					sShiftShortName = CStr(oRecordset.Fields("ShiftShortName").Value)
				Else
					sShiftShortName = ""
				End If
				If Not IsNull(oRecordset.Fields("PayrollCLC").Value) Then
					sPayrollCLC = CStr(oRecordset.Fields("PayrollCLC").Value)
				Else
					sPayrollCLC = ""
				End If
				If Not IsNull(oRecordset.Fields("CompanyID").Value) Then
					sCompanyID = CStr(oRecordset.Fields("CompanyID").Value)
					if len (sCompanyID) = 1 then sCompanyID = "0" & sCompanyID
				Else
					sCompanyID = ""
				End If	
				If Not IsNull(oRecordset.Fields("BANKSHORTNAME").Value) Then
					sBANKSHORTNAME = CStr(oRecordset.Fields("BANKSHORTNAME").Value)
				Else
					sBANKSHORTNAME = ""
				End If	
				If Not IsNull(oRecordset.Fields("EmployeeTypeshortname").Value) Then
					sEmployeeTypeshortname = CStr(oRecordset.Fields("EmployeeTypeshortname").Value)
				Else
					sEmployeeTypeshortname = ""
				End If	
				If Not IsNull(oRecordset.Fields("zonepath").Value) Then
					szonepath = mid(oRecordset.Fields("zonepath").Value,5,InStr(5, oRecordset.Fields("zonepath").Value, ",")-5)	
					if len (szonepath) = 1 then szonepath = "0" & szonepath
				Else
					szonepath = ""
				End If	
				If Not IsNull(oRecordset.Fields("payrollID").Value) Then
					spayrollID = DisplayNumericDateFromSerialNumber(oRecordset.Fields("payrollID").Value)
					sYearA= mid(oRecordset.Fields("payrollID").Value,1,4)
					sMonthA= mid(oRecordset.Fields("payrollID").Value,5,2)
					if mid(oRecordset.Fields("payrollID").Value,7,2) < 16 then
						sQuincena =  int(mid(oRecordset.Fields("payrollID").Value,5,2)) * 2-1 
					else
						sQuincena =  int(mid(oRecordset.Fields("payrollID").Value,5,2)) * 2
					end if
				Else
					spayrollID = ""
					sYearA = ""
					sMonthA = ""
					sQuincena
				End If
				If Not IsNull(oRecordset.Fields("PayrollTypeID").Value) Then
					select case cint(oRecordset.Fields("PayrollTypeID").Value)
						case 0 
							sTipo = "4"
						case 1
							sTipo = "0"
						case 2,4
							stipo = "1"
						case Else
							stipo = "X"
					end select 
				Else
					sTipo = ""
				End If	

				'If StrComp(sCurrentID, CStr(oRecordset.Fields("EmployeeID").Value) & "," & CStr(oRecordset.Fields("RecordDate").Value), vbBinaryCompare) <> 0 Then
				If StrComp(sCurrentID, CStr(oRecordset.Fields("EmployeeID").Value), vbBinaryCompare) <> 0 Then
					If Len(sCurrentID) > 0 Then
						sRowContents = Replace(sRowContents, "<PERCEPCIONES />", adTotal(0))
						sRowContents = Replace(sRowContents, "<DEDUCCIONES />", adTotal(1))
						sRowContents = Replace(sRowContents, "<TOTAL />", adTotal(0) - adTotal(1))
						For iIndex = iConceptCounter + 1 To 40
							sRowContents = Replace(sRowContents, "<CONCEPTO />", "|||0.0<CONCEPTO />")
						Next
						sRowContents = Replace(sRowContents, "<CONCEPTO />", "")
						lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
						adTotal(0) = 0
						adTotal(1) = 0
					End If
					sRowContents = "N|" & sTipo
					sRowContents = sRowContents & "|" & sEmployeeNumber
					sRowContents = sRowContents & "|" & sRFC
					sRowContents = sRowContents & "|" & sEmployeeLastName & " " & sEmployeeLastName2 & " " & sEmployeeName
					sRowContents = sRowContents & "|" & (CInt(Left(CStr(oRecordset.Fields("MinRecordDate").Value), Len("0000"))) * 100) + CInt(GetPayrollNumber(CLng(oRecordset.Fields("MinRecordDate").Value)))
					sRowContents = sRowContents & "|" & (CInt(Left(CStr(oRecordset.Fields("MaxRecordDate").Value), Len("0000"))) * 100) + CInt(GetPayrollNumber(CLng(oRecordset.Fields("MaxRecordDate").Value)))
					sRowContents = sRowContents & "|||0|0|0.0|||0|0|0.0|||0|0|0.0|||0|0|0.0|||0|0|0.0|||0|0|0.0|||0|0|0.0|||0|0|0.0|||0|0|0.0|||0|0|0.0|||0|0|0.0|||0|0|0.0"
					sRowContents = sRowContents & "|" & CStr(oRecordset.Fields("AreaShortName").Value)
					sRowContents = sRowContents & "|" & "<PERCEPCIONES />"
					sRowContents = sRowContents & "|" & "<DEDUCCIONES />"
					sRowContents = sRowContents & "|" & "<TOTAL />"
					sRowContents = sRowContents & "<CONCEPTO />"
					sRowContents = sRowContents & "|" &  sCheckNumber
					sRowContents = sRowContents & "|"
						If StrComp(sAccountNumber, ".", vbBinaryCompare) <> 0 Then
							If Len(sAccountNumber) <> 0 Then
								asTemp = Split(sAccountNumber, LIST_SEPARATOR)
								sRowContents = sRowContents & asTemp(0)
							Else
								sRowContents = sRowContents & ""
							End If
						End If
					sRowContents = sRowContents & "|" & sPayrollCLC
					sRowContents = sRowContents & "|" & sPositionShortName
					If CStr(oRecordset.Fields("EmployeeTypeID").Value) = 1 Then
						sRowContents = sRowContents & "|" & sGroupGradeLevelShortName
						sRowContents = sRowContents & "|" & sIntegrationID
					Else
						sRowContents = sRowContents & "|" & Left(sLevelShortName, Len("0"))
						sRowContents = sRowContents & "|" & Right(CStr(oRecordset.Fields("LevelShortName").Value), Len("00"))
					End If
					sRowContents = sRowContents & "|" & Left(sJourneyShortName, Len("0"))
					sRowContents = sRowContents & "|" & sJourneyShortName

					sRowContents = sRowContents & "|" & sCompanyID & "|" & sBANKSHORTNAME & "|" & sEmployeeTypeshortname & "|" & sTabulador 
					sRowContents = sRowContents & "|" & sZonePath  & "||" & sPayrollID  & "|||||" & sQuincena & "|" & sMonthA 
					sRowContents = sRowContents & "|" & sYearA  & "|||||"
					sRowContents = sRowContents & CStr(oRecordset.Fields("AreaShortName").Value) & "|" 

					iConceptCounter = 0
					sCurrentID = CStr(oRecordset.Fields("EmployeeID").Value)' & "," & CStr(oRecordset.Fields("RecordDate").Value)
				End If
				If iConceptCounter < 40 Then
					If CInt(oRecordset.Fields("IsDeduction").Value) = 1 Then
						sRowContents = Replace(sRowContents, "<CONCEPTO />", "|D|" & CStr(oRecordset.Fields("ConceptShortName").Value) & "|" & CStr(oRecordset.Fields("TotalAmount").Value) & "<CONCEPTO />")
						adTotal(1) = adTotal(1) + CDbl(oRecordset.Fields("TotalAmount").Value)
					Else
						sRowContents = Replace(sRowContents, "<CONCEPTO />", "|P|" & CStr(oRecordset.Fields("ConceptShortName").Value) & "|" & CStr(oRecordset.Fields("TotalAmount").Value) & "<CONCEPTO />")
						adTotal(0) = adTotal(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
					End If
					iConceptCounter = iConceptCounter + 1
				End If
				'LAYOUT SPEP GAGR'
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
			sRowContents = Replace(sRowContents, "<PERCEPCIONES />", adTotal(0))
			sRowContents = Replace(sRowContents, "<DEDUCCIONES />", adTotal(1))
			sRowContents = Replace(sRowContents, "<TOTAL />", adTotal(0) - adTotal(1))
			For iIndex = iConceptCounter + 1 To 40
				sRowContents = Replace(sRowContents, "<CONCEPTO />", "|||0.0<CONCEPTO />")
			Next
			sRowContents = Replace(sRowContents, "<CONCEPTO />", "")
			lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
		End If
		If FileExists(sFilePath, sErrorDescription) Then
			lErrorNumber = ZipFile(sFilePath, Replace(sFilePath, ".txt", ".zip"), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudieron guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', ' ')", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
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
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "window.CheckFileIFrame.location.href = 'CheckFile.asp?bNoReport=1';" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
		End If
	Else
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "window.CheckFileIFrame.location.href = 'CheckFile.asp?bNoReport=1';" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	Set oRecordset = Nothing
	BuildReport1401 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1403(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the monthly payroll resume based on CLCs
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1403"
	
	Dim oRecordset
	Dim sRowContents
	Dim lErrorNumber
	Dim sQuery
	Dim lPeriod
	Dim lPayrollID
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim sMonth
    Dim sYear
    Dim sZonesID
    Dim sExtraConditions
    Dim sExtraTables
    Dim lCount
    Dim sPayrollType
    Dim lTotReg
    Dim lTotPer
    Dim lTotDed
    Dim lTotLiq
    Dim lReg
    Dim lPer
    Dim lDed
    Dim lLiq
    Dim lTypeMax
    Dim i
    Dim sPayrollTypeCondition
   
    lPayrollID = oRequest("PayrollID").Item


	asPeriods = Split("1,1,2,2,3,3,4,4,5,5,6,6",",")
	sYear = oRequest("YearID").Item
	sErrorDescription = "No se pudieron obtener las CLCs generadas para la nómina especificada."
    sExtraConditions = ""
    sExtraTables = ""
    lTotReg = 0
    lTotPer = 0
    lTotDed = 0
    lTotLiq = 0
    i=1
    sRowContents = ""
    sDate=""

    sExtraConditions = "And (rpt.ZoneID = Z.ZoneID)"
    sExtraTables = "rpt, Zones Z "
    sPayrollTypeCondition = ""

     If Len(oRequest("ZoneID").Item) >0 Then
        sZonesID = split(GetParameterFromURLString(oRequest, "ZoneID"),",")
        sExtraConditions = "And (Z.ZonePath like'%,"&trim(sZonesID(0))&",%' "
        For iIndex = 1 To UBound(sZonesID)     
               sExtraConditions = sExtraConditions & " OR Z.ZonePath like'%,"&trim(sZonesID(iIndex))&",%' "
        Next    
        sExtraConditions =sExtraConditions & ")"        
    End If

    If Len(oRequest("PayrollTypeID").Item)=0 Then
        lTypeMax = 2
    Else
        lTypeMax = 1
    End If

    

    Do While i <= lTypeMax

        If Len(oRequest("PayrollTypeID").Item)=0 Then
            If (i=1) Then
                sPayrollType="ORDINARIA"
            Else
                sPayrollType="EXTRAORDINARIA" 
            End If
            sPayrollTypeCondition = sPayrollTypeCondition & " And rpt.PayrollTypeID = "& i
        Else 
            sQuery = "select PayrollTypeName from PayrollTypes Where PayrollTypeID = " & oRequest("PayrollTypeID").Item
            lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
            sPayrollType = CStr(oRecordset.Fields("PayrollTypeName").Value)
            sPayrollTypeCondition = sPayrollTypeCondition & " And rpt.PayrollTypeID = "& oRequest("PayrollTypeID").Item 
        End If
        If (oRequest("PayrollID").Item <> 0) Then
            lPeriod = Mid(lPayrollID,1,4) & "0" & asPeriods(CInt(Mid(lPayrollID,5,2))-1)
            sQuery = "Select  PayrollCLC, CompanyName, PayrollCode, Fortnightly,BankName, Case When AccountNumber Like '.' Then 'CHEQUE' Else 'DEBITO' End As FPago, Case When Z.ZonePath Like'%,9,%' Then 'LOCAL' Else 'FORANEO' End as Area, PayrollMemorandum, PayrollFile,Count(PayrollCLC) as Registros, Sum(Percepciones) as Percepciones, Sum (Deducciones) as Deducciones, Sum(Liquido) as Liquido From rpt_resumen_mensual RPT, Zones Z Where RPT.ZoneID = Z.ZoneID And RPT.PayrollID  = " & lPayrollID & " And PayrollCode = '" & lPeriod & "' "& sExtraConditions & sPayrollTypeCondition &" Group By PayrollCLC,CompanyName,PayrollCode, Fortnightly,BankName,Case When AccountNumber Like '.' Then 'CHEQUE' Else 'DEBITO' End, Case When Z.ZonePath Like'%,9,%' Then 'LOCAL' Else 'FORANEO' End , PayrollMemorandum, PayrollFile Order by PayrollCLC"
        ElseIf (oRequest("MonthID").Item<> 0) Then 
            sMonth = (oRequest("MonthID").Item)
            If Len(sMonth) < 2 Then
                sMonth = "0"&sMonth
            End If
            lPayrollID = sYear & sMonth
            lPeriod = sYear  & "0" & asPeriods(CInt(sMonth)-1)     
            sQuery = "Select PayrollCLC, CompanyName, PayrollCode, Fortnightly, BankName, CASE WHEN AccountNumber LIKE '.' THEN 'CHEQUE' ELSE 'DEBITO' END AS FPago, CASE WHEN Z.ZonePath LIKE '%,9,%' THEN 'LOCAL' ELSE 'FORANEO' END AS Area, PayrollMemorandum, PayrollFile, COUNT (PayrollCLC) AS Registros, SUM (Percepciones) AS Percepciones, SUM (Deducciones) AS Deducciones, SUM (Liquido) AS Liquido From rpt_resumen_mensual "& sExtraTables &" Where payrollID between "& lPayrollID &"01 and "& lPayrollID &"31 And PayrollCode = '"& lPeriod &"' "& sExtraConditions & sPayrollTypeCondition &" Group by PayrollCLC, CompanyName, PayrollCode, Fortnightly, BankName, CASE WHEN AccountNumber LIKE '.' THEN 'CHEQUE' ELSE 'DEBITO' END, CASE WHEN Z.ZonePath LIKE '%,9,%' THEN 'LOCAL' ELSE 'FORANEO' END, PayrollMemorandum, PayrollFile Order by PayrollCLC"
        ElseIf (oRequest("QuarterID").Item<>0) Then
            lPeriod = sYear & "0" & oRequest("QuarterID").Item
            sQuery = "Select PayrollCLC, CompanyName, PayrollCode, Fortnightly, BankName, CASE WHEN AccountNumber LIKE '.' THEN 'CHEQUE' ELSE 'DEBITO' END AS FPago, CASE WHEN Z.ZonePath LIKE '%,9,%' THEN 'LOCAL' ELSE 'FORANEO' END AS Area, PayrollMemorandum, PayrollFile, COUNT (PayrollCLC) AS Registros, SUM (Percepciones) AS Percepciones, SUM (Deducciones) AS Deducciones, SUM (Liquido) AS Liquido From rpt_resumen_mensual "& sExtraTables &" Where PayrollCode = '"& lPeriod &"' "& sExtraConditions & sPayrollTypeCondition &" Group by PayrollCLC, CompanyName, PayrollCode, Fortnightly, BankName, CASE WHEN AccountNumber LIKE '.' THEN 'CHEQUE' ELSE 'DEBITO' END, CASE WHEN Z.ZonePath LIKE '%,9,%' THEN 'LOCAL' ELSE 'FORANEO' END, PayrollMemorandum, PayrollFile Order by PayrollCLC"
        Else 
            sQuery = "Select PayrollCLC, CompanyName, PayrollCode, Fortnightly, BankName, CASE WHEN AccountNumber LIKE '.' THEN 'CHEQUE' ELSE 'DEBITO' END AS FPago, CASE WHEN Z.ZonePath LIKE '%,9,%' THEN 'LOCAL' ELSE 'FORANEO' END AS Area, PayrollMemorandum, PayrollFile, COUNT (PayrollCLC) AS Registros, SUM (Percepciones) AS Percepciones, SUM (Deducciones) AS Deducciones, SUM (Liquido) AS Liquido From rpt_resumen_mensual "& sExtraTables &" Where PayrollID Between "& sYear &"0101 And "& sYear &"1231 And PayrollCode In( '"& sYear &"01','"& sYear &"02', '"& sYear &"03', '"& sYear &"04', '"& sYear &"05', '"& sYear &"06') "& sExtraConditions & sPayrollTypeCondition &" Group by PayrollCLC, CompanyName, PayrollCode, Fortnightly, BankName, CASE WHEN AccountNumber LIKE '.' THEN 'CHEQUE' ELSE 'DEBITO' END, CASE WHEN Z.ZonePath LIKE '%,9,%' THEN 'LOCAL' ELSE 'FORANEO' END, PayrollMemorandum, PayrollFile Order by  PayrollCLC, PayrollCode"
        End If

        lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	    Response.Write vbNewLine & "<!-- Query: " & sQuery & " -->" & vbNewLine
	    If lErrorNumber = 0 Then
            If Not oRecordset.EOF Then
                If(i=1)Then
                    sDate = GetSerialNumberForDate("")
    			    sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
    			    lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
    			    sFilePath = sFilePath & "\"
    			    sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
    			    Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
    			    Response.Flush()
    			    sDocumentName = sFilePath & "Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"
                    sRowContents = sRowContents & "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
                        sRowContents = sRowContents & "<TR>"
                            sRowContents = sRowContents & "<TD COLSPAN=3 ROWSPAN=7><IMG SRC=""http://"& Request.ServerVariables("SERVER_NAME")&"/SIAP/Images/LogoISSSTE.gif"" WIDTH=""260"" HEIGHT=""120"" ALT=""LogoISSSTE"" BORDER=""0"" /></TD>"    
                            sRowContents = sRowContents & "<TD COLSPAN=8 ROWSPAN=7> </TD>"    
                            sRowContents = sRowContents & "<TD COLSPAN=2 ROWSPAN=7><IMG SRC=""http://"& Request.ServerVariables("SERVER_NAME")&"/SIAP/Images/Logo_ISSSTE.gif"" WIDTH=""120"" HEIGHT=""130""  BORDER=""0"" /></TD></TR>"    
                    sRowContents = sRowContents &"</TABLE>"
                    sRowContents = sRowContents & " <br/>"       
                    sRowContents = sRowContents & "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
                        sRowContents = sRowContents & "<TR>"
                            sRowContents = sRowContents & "<TD COLSPAN=3 FACE=""verdana"">DIRECCION DE ADMINISTRACIÓN</TD>"   
                            sRowContents = sRowContents & "<TD COLSPAN=6></TD>"             
                            sRowContents = sRowContents & "<TD COLSPAN=4 style=""border-width: 1px;border: solid;""><b>Anexo del memorándum número SP/     /"& oRequest("YearID").Item &"</b></TD></TR>"    
                        sRowContents = sRowContents & "<TR>"
                            sRowContents = sRowContents & "<TD COLSPAN=3 FACE=""verdana"">Subdirección de Personal</TD></TR>"   
                        sRowContents = sRowContents & "<TR>"
                            sRowContents = sRowContents & "<TD COLSPAN=3 FACE=""verdana"">Jefatura de Servicios de Informática</TD></TR>"   
                        sRowContents = sRowContents & "<TR>"
                            sRowContents = sRowContents & "<TD></TD>"
                            If Len(oRequest("PayrollID").Item)>1 Then
                                sDate = "(" & Mid(oRequest("PayrollID").Item,1,4) & "/" & Mid(oRequest("PayrollID").Item,5,2) & "/" & Mid(oRequest("PayrollID").Item,7,2) & ")"
                            ElseIf (oRequest("MonthID").Item<> 0) Then 
                                sDate = "(" & oRequest("YearID").Item & "/" & oRequest("MonthID").Item & ")"
                            Else 
                                sDate = "("& oRequest("YearID").Item &")"
                            End If 
                            sRowContents = sRowContents & "<TD align=""center"" COLSPAN=12 style=""border-width: 1px;border: solid;""><b><i>RESUMEN MENSUAL DE NÓMINAS "& sDate &"</i></b></TD>"   
                            sRowContents = sRowContents & "<TD></TD></TR>"   
                    sRowContents = sRowContents &"</TABLE>" 
                End If
                sRowContents =  sRowContents & "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
			        sRowContents = sRowContents & "<TR>"
                         
                        sRowContents = sRowContents & "<TD COLSPAN =""2"" style=""border-width: 1px;border: solid;""><B>"& sPayrollType &"</B></TD>"
                        sRowContents = sRowContents & "<TD COLSPAN =""11""><B></B></TD>"
                    sRowContents = sRowContents & "</TR>"
                    sRowContents = sRowContents & "<TR>"
                        sRowContents = sRowContents & "<TD><B>No.</B></TD>"
		                sRowContents = sRowContents & "<TD><B>C.L.C/LOCAL</B></TD>"
				        sRowContents = sRowContents & "<TD><B>UNIDAD</B></TD>"
				        sRowContents = sRowContents & "<TD><B>QNA.</B></TD>"
                        sRowContents = sRowContents & "<TD><B>BANCO</B></TD>"
                        sRowContents = sRowContents & "<TD><B>F/PAGO</B></TD>"
                        sRowContents = sRowContents & "<TD><B>AREA</B></TD>"
				        sRowContents = sRowContents & "<TD><B>No. MEMORANDA</B></TD>"
                        sRowContents = sRowContents & "<TD><B>ARCHIVO</B></TD>"				
				        sRowContents = sRowContents & "<TD><B>REG.</B></TD>"
				        sRowContents = sRowContents & "<TD><B>PERCEPCIONES</B></TD>"
				        sRowContents = sRowContents & "<TD><B>DEDUCCIONES</B></TD>"
				        sRowContents = sRowContents & "<TD><B>LIQUIDO</B></TD>"
			        sRowContents = sRowContents & "</TR>"
                    lCount = 1
                    Do While Not oRecordset.EOF
                       sRowContents = sRowContents & "<TR>"
                           sRowContents = sRowContents & "<TD>" & lCount & "</TD>"
				           sRowContents = sRowContents & "<TD>'" & CStr(oRecordset.Fields("PayrollCLC").Value) & "</TD>"
					       sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("CompanyName").Value) & "</TD>"
                           If Len(oRecordset.Fields("Fortnightly").Value)>0 Then 
                               sRowContents = sRowContents & "<TD>'" & CStr(oRecordset.Fields("Fortnightly").Value)  & "/"& oRequest("YearID").Item & "</TD>"
                           Else
                               sRowContents = sRowContents & "<TD>  </TD>"
                           End If
                           sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("BankName").Value) & "</TD>"
                           sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("FPago").Value) & "</TD>"
                           sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("Area").Value) & "</TD>"
                           If Len(oRecordset.Fields("PayrollMemorandum").Value)>0 Then 
                               sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("PayrollMemorandum").Value) & "</TD>"
                           Else
				           sRowContents = sRowContents & "<TD>  </TD>"
                           End IF
                           If Len(oRecordset.Fields("PayrollFile").Value)>0 Then 
                               sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("PayrollFile").Value) & "</TD>"
                           Else
					           sRowContents = sRowContents & "<TD>  </TD>"
                           End IF
                           sRowContents = sRowContents & "<TD>" & CInt(oRecordset.Fields("Registros").Value) & "</TD>"
                           lReg= lReg + CInt(oRecordset.Fields("Registros").Value)
					       sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(oRecordset.Fields("Percepciones").Value),2,True,True,True) & "</TD>"
					       lPer = lPer+ CDbl(oRecordset.Fields("Percepciones").Value)
                           sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(oRecordset.Fields("Deducciones").Value),2,True,True,True) & "</TD>"
                           lDed = lDed + CDbl(oRecordset.Fields("Deducciones").Value)                            
                           sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(oRecordset.Fields("Liquido").Value),2,True,True,True) & "</TD>"
					       lLiq = lLiq + CDbl(oRecordset.Fields("Liquido").Value)
                       sRowContents = sRowContents & "</TR>"
				       lCount = lCount + 1
                       oRecordset.MoveNext
                    Loop

                    lTotReg = lTotReg+lReg
                    lTotPer = lTotPer+lPer
                    lTotDed = lTotDed+lDed
                    lTotLiq = lTotLiq+lLiq
                sRowContents = sRowContents & "</TABLE>"
                sRowContents = sRowContents & " <br/> "
                sRowContents = sRowContents & "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
                sRowContents = sRowContents & "<TR><TD COLSPAN = 8> </TD>"
                sRowContents = sRowContents & "<TD style=""border-width: 1px;border: solid;"">TOTAL QNA.</TD>"
                sRowContents = sRowContents & "<TD style=""border-width: 1px;border: solid;"">"& lReg &"</TD>"
                sRowContents = sRowContents & "<TD style=""border-width: 1px;border: solid;"">"& FormatNumber(lPer,2,True,True,True) &"</TD>"
                sRowContents = sRowContents & "<TD style=""border-width: 1px;border: solid;"">"& FormatNumber(lDed,2,True,True,True) &"</TD>"
                sRowContents = sRowContents & "<TD style=""border-width: 1px;border: solid;"">"& FormatNumber(lLiq,2,True,True,True) &"</TD></TR>"
                sRowContents = sRowContents & "</TABLE>"
                sRowContents = sRowContents & " <br/> <br/>"
            End If
        Else
            lErrorNumber = -1
			sErrorDescription = "No se encontraron documentos capturados para la quincena elegida."
            'Exit While
        End If
        i=i+1
        lPer = 0
        lDed = 0
        lLiq = 0
        lReg = 0
        sPayrollTypeCondition = ""       
    Loop   
    lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
    If lErrorNumber = 0 Then
        sRowContents =  "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
        sRowContents = sRowContents & "<TR><TD COLSPAN = 8> </TD>"
        sRowContents = sRowContents & "<TD style=""border-width: 1px;border: solid;"" >TOTAL GENERAL</TD>"
        sRowContents = sRowContents & "<TD style=""border-width: 1px;border: solid;"">"& lTotReg &"</TD>"
        sRowContents = sRowContents & "<TD style=""border-width: 1px;border: solid;"">"& FormatNumber(lTotPer,2,True,True,True) &"</TD>"
        sRowContents = sRowContents & "<TD style=""border-width: 1px;border: solid;"">"& FormatNumber(lTotDed,2,True,True,True) &"</TD>"
        sRowContents = sRowContents & "<TD style=""border-width: 1px;border: solid;"">"& FormatNumber(lTotLiq,2,True,True,True) &"</TD></TR>"
        sRowContents = sRowContents & "</TABLE>"
        lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
        If lErrorNumber = 0 Then
            lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
            lErrorNumber = DeleteFolder(sFilePath, sErrorDescription)
            oEndDate = Now()
            If (lErrorNumber = 0) And B_USE_SMTP Then
		        If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
		    End If       
        End If
    Else
        lErrorNumber = -1
		sErrorDescription = "No se encontraron documentos capturados para la quincena elegida."
    End If

	Set oRecordset = Nothing
	BuildReport1403 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1403b(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the monthly payroll resume based on CLCs
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1403"
	Dim alConcepts
	Dim asCLCs
	Dim asCondition
	Dim asPeriods
	Dim iIndex
	Dim jIndex
	Dim sCondition
	Dim oRecordset
	Dim oRecordsetConcepts
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim sQuery
	Dim sQueryC
	Dim lPeriod
	Dim lPayrollID
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim alConcepts2
	Dim alImports
    Dim sMonth
    Dim sYear
    Dim sZonesID
    Dim sExtraConditions
    Dim sExtraTables
    Dim lCount
    Dim sPayrollType
    Dim lTotReg
    Dim lTotPer
    Dim lTotDed
    Dim lTotLiq
    Dim lReg
    Dim lPer
    Dim lDed
    Dim lLiq
    Dim lType()
    dim i
    Dim sPayrollTypeCondition
   

	asPeriods = Split("1,1,2,2,3,3,4,4,5,5,6,6",",")
	alConcepts2 = Split("1,4,5,6,7,8,13,47,89,146",",")
	alImports = Split("0,0,0,0,0,0,0,0,0,0",",")
	sYear = oRequest("YearID").Item
	sErrorDescription = "No se pudieron obtener las CLCs generadas para la nómina especificada."
    sExtraConditions = ""
    sExtraTables = ""
    lTotReg = 0
    lTotPer = 0
    lTotDed = 0
    lTotLiq = 0
    i=0
    sDate=""

    sExtraConditions = "And (rpt.ZoneID = Z.ZoneID)"
    sExtraTables = "rpt, Zones Z "
    sPayrollTypeCondition = ""

     If Len(oRequest("ZoneID").Item) >0 Then
        sZonesID = split(GetParameterFromURLString(oRequest, "ZoneID"),",")
        sExtraConditions = "And (Z.ZonePath like'%,"&trim(sZonesID(0))&",%' "
        For iIndex = 1 To UBound(sZonesID)     
               sExtraConditions = sExtraConditions & " OR Z.ZonePath like'%,"&trim(sZonesID(iIndex))&",%' "
        Next    
        sExtraConditions =sExtraConditions & ")"        
    End If

    If Len(oRequest("PayrollTypeID").Item) > 0 Then
        Redim lType(0)
        lType(0) = CInt(oRequest("PayrollTypeID").Item)
    Else
        Redim lType(1)
        lType(0) = 1
        lType(1) = 2
    End If

    'If Len(oRequest("PayrollTypeID").Item) > 0 Then
    '    sExtraConditions = sExtraConditions & " And PayrollTypeID = "& oRequest("PayrollTypeID").Item 
    '    sQuery = "select PayrollTypeName from PayrollTypes Where PayrollTypeID = " & oRequest("PayrollTypeID").Item
    '    lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
    '    sPayrollType = CStr(oRecordset.Fields("PayrollTypeName").Value)
    'End If

    sRowContents = sRowContents & "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
        sRowContents = sRowContents & "<TR>"
            sRowContents = sRowContents & "<TD COLSPAN=3 ROWSPAN=7><IMG SRC=""http://"& Request.ServerVariables("SERVER_NAME")&"/SIAP/Images/LogoISSSTE.gif"" WIDTH=""260"" HEIGHT=""120"" ALT=""LogoISSSTE"" BORDER=""0"" /></TD>"    
            sRowContents = sRowContents & "<TD COLSPAN=8 ROWSPAN=7> </TD>"    
            sRowContents = sRowContents & "<TD COLSPAN=2 ROWSPAN=7><IMG SRC=""http://"& Request.ServerVariables("SERVER_NAME")&"/SIAP/Images/Logo_ISSSTE.gif"" WIDTH=""120"" HEIGHT=""130""  BORDER=""0"" /></TD>"    
        sRowContents = sRowContens & "</TR>"
     sRowContents = sRowContents &"</TABLE>"

    sRowContents = sRowContents & " <br/>"       
    sRowContents = sRowContents & "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
        'sRowContents = sRowContents & "<TR>"
        '    sRowContents = sRowContents & "<TD COLSPAN=3 ROWSPAN=8><IMG SRC=""http://"& Request.ServerVariables("SERVER_NAME")&"/SIAP/Images/LogoISSSTE.gif"" WIDTH=""260"" HEIGHT=""120"" ALT=""LogoISSSTE"" BORDER=""0"" /></TD>"    
        '    sRowContents = sRowContents & "<TD COLSPAN=8 ROWSPAN=8> </TD>"    
        '    sRowContents = sRowContents & "<TD COLSPAN=2 ROWSPAN=8><IMG SRC=""http://"& Request.ServerVariables("SERVER_NAME")&"/SIAP/Images/Logo_ISSSTE.gif"" WIDTH=""120"" HEIGHT=""130""  BORDER=""0"" /></TD>"    
        'sRowContents = sRowContens & "</TR>"
        sRowContents = sRowContents & "<TR>"
            sRowContents = sRowContents & "<TD COLSPAN=3 FACE=""verdana"">DIRECCION DE ADMINISTRACIÓN</TD>"   
            sRowContents = sRowContents & "<TD COLSPAN=7></TD>"             
            sRowContents = sRowContents & "<TD COLSPAN=3 style=""border-width: 1px;border: solid;""><b>Anexo del memorándum número SP/   /"& oRequest("YearID").Item &"</b></TD>"    
         sRowContents = sRowContens & "</TR>"
         sRowContents = sRowContents & "<TR>"
            sRowContents = sRowContents & "<TD COLSPAN=3 FACE=""verdana"">Subdirección de Personal</TD>"   
         sRowContents = sRowContens & "</TR>"
         sRowContents = sRowContents & "<TR>"
            sRowContents = sRowContents & "<TD COLSPAN=3 FACE=""verdana"">Jefatura de Servicios de Informática</TD>"   
         sRowContents = sRowContens & "</TR>"
         sRowContents = sRowContents & "<TR>"
            sRowContents = sRowContents & "<TD></TD>"
            If Len(oRequest("PayrollID").Item)>1 Then
                  sDate = "(" & Mid(oRequest("PayrollID").Item,1,4) & "/" & Mid(oRequest("PayrollID").Item,5,2) & "/" & Mid(oRequest("PayrollID").Item,7,2) & ")"
            Else 
                sDate = "("& oRequest("YearID").Item &")"
            End If 
            sRowContents = sRowContents & "<TD align=""center"" COLSPAN=12 style=""border-width: 1px;border: solid;""><b><i>RESUMEN MENSUAL DE NÓMINAS "& sDate &"</i></b></TD>"   
            sRowContents = sRowContents & "<TD></TD>"      
         sRowContents = sRowContens & "</TR>"
    sRowContents = sRowContents &"</TABLE>"
    
    sRowContents = sRowContents & " <br/> <br/>"           
    
    Do While i<=UBound(lType)
         sPayrollTypeCondition = ""
         lReg = 0
         lPer = 0
         lDed = 0
         lLiq = 0
    
          
        sPayrollTypeCondition =" And PayrollTypeID = "& lType(i) 
        sQuery = "select PayrollTypeName from PayrollTypes Where PayrollTypeID = " & lType(i)
        lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
        sPayrollType = CStr(oRecordset.Fields("PayrollTypeName").Value)

        

        If (oRequest("PayrollID").Item <> 0) Then
            lPayrollID = oRequest("PayrollID").Item
	        lPeriod = Mid(lPayrollID,1,4) & "0" & asPeriods(CInt(Mid(lPayrollID,5,2))-1)
            sQuery = "Select  PayrollCLC, CompanyName, PayrollCode, Fortnightly,BankName, Case When AccountNumber Like '.' Then 'CHEQUE' Else 'DEBITO' End As FPago, Case When Z.ZonePath Like'%,9,%' Then 'LOCAL' Else 'FORANEO' End as Area, PayrollMemorandum, PayrollFile,Count(PayrollCLC) as Registros, Sum(Percepciones) as Percepciones, Sum (Deducciones) as Deducciones, Sum(Liquido) as Liquido From rpt_resumen_mensual RPT, Zones Z Where RPT.ZoneID = Z.ZoneID And RPT.PayrollID  = " & lPayrollID & " And PayrollCode = '" & lPeriod & "' "& sExtraConditions & sPayrollTypeCondition &" Group By PayrollCLC,CompanyName,PayrollCode, Fortnightly,BankName,Case When AccountNumber Like '.' Then 'CHEQUE' Else 'DEBITO' End, Case When Z.ZonePath Like'%,9,%' Then 'LOCAL' Else 'FORANEO' End , PayrollMemorandum, PayrollFile Order by PayrollCLC"
            'SQuery = "Select PayrollCLC,PayrollCode,CompanyID, CompanyName, PayrollID, BankID,BankName, Count(PayrollCLC) as Registros, Sum(Percepciones) as Percepciones, Sum (Deducciones) as Deducciones, Sum(Liquido) as Liquido From rpt_resumen_mensual "& sExtraTables &" Where payrollID = " & lPayrollID & " And PayrollCode = '" & lPeriod & "' "& sExtraConditions &" Group by PayrollCLC,PayrollCode,CompanyID, CompanyName, payrollID, BankID,BankName Order by PayrollCLC;"
            'sQuery = "Select rpt.PayrollCLC,rpt.PayrollCode,rpt.CompanyID, rpt.CompanyName, rpt.PayrollID, rpt.BankID,rpt.BankName,  case when ACCOUNTNUMBER like '.' then 'CHEQUE' else 'DEBITO' end as fpago,case when z.zonePath like '%,9,%' then 'local' else 'foraneo' end as area, Count(PayrollCLC) as Registros, Sum(Percepciones) as Percepciones, Sum (Deducciones) as Deducciones, Sum(Liquido) as Liquido From rpt_resumen_mensual rpt, zones z  "& sExtraTables &" Where payrollID = " & lPayrollID & " And PayrollCode ='" & lPeriod & "' "& sExtraConditions &" And  rpt.zoneID = z.zoneID Group by rpt.PayrollCLC,rpt.PayrollCode,rpt.CompanyID, rpt.CompanyName, rpt.PayrollID, rpt.BankID,rpt.BankName,  case when ACCOUNTNUMBER like '.' then 'CHEQUE' else 'DEBITO' end ,case when z.zonePath like '%,9,%' then 'local' else 'foraneo' end Order by PayrollCLC"
        ElseIf (oRequest("MonthID").Item<> 0) Then 
            sMonth = (oRequest("MonthID").Item)
            If Len(sMonth) < 2 Then
                sMonth = "0"&sMonth
            End If
            lPayrollID = sYear & sMonth
            lPeriod = sYear  & "0" & asPeriods(CInt(sMonth)-1)     
            sQuery = "Select PayrollCLC, CompanyName, PayrollCode, Fortnightly, BankName, CASE WHEN AccountNumber LIKE '.' THEN 'CHEQUE' ELSE 'DEBITO' END AS FPago, CASE WHEN Z.ZonePath LIKE '%,9,%' THEN 'LOCAL' ELSE 'FORANEO' END AS Area, PayrollMemorandum, PayrollFile, COUNT (PayrollCLC) AS Registros, SUM (Percepciones) AS Percepciones, SUM (Deducciones) AS Deducciones, SUM (Liquido) AS Liquido From rpt_resumen_mensual "& sExtraTables &" Where payrollID between "& lPayrollID &"01 and "& lPayrollID &"31 And PayrollCode = '"& lPeriod &"' "& sExtraConditions & sPayrollTypeCondition &" Group by PayrollCLC, CompanyName, PayrollCode, Fortnightly, BankName, CASE WHEN AccountNumber LIKE '.' THEN 'CHEQUE' ELSE 'DEBITO' END, CASE WHEN Z.ZonePath LIKE '%,9,%' THEN 'LOCAL' ELSE 'FORANEO' END, PayrollMemorandum, PayrollFile Order by PayrollCLC"
            'SQuery = "Select PayrollCLC,PayrollCode,CompanyID, CompanyName, PayrollID, BankID,BankName, Count(PayrollCLC) as Registros, Sum(Percepciones) as Percepciones, Sum (Deducciones) as Deducciones, Sum(Liquido) as Liquido From rpt_resumen_mensual "& sExtraTables &" Where payrollID between "& lPayrollID &"01 and "& lPayrollID &"31 And PayrollCode = '"& lPeriod &"' "& sExtraConditions &" Group by PayrollCLC,PayrollCode,CompanyID, CompanyName, payrollID, BankID,BankName Order by PayrollCLC"
        ElseIf (oRequest("QuarterID").Item<>0) Then
            lPeriod = sYear & "0" & oRequest("QuarterID").Item
            sQuery = "Select PayrollCLC, CompanyName, PayrollCode, Fortnightly, BankName, CASE WHEN AccountNumber LIKE '.' THEN 'CHEQUE' ELSE 'DEBITO' END AS FPago, CASE WHEN Z.ZonePath LIKE '%,9,%' THEN 'LOCAL' ELSE 'FORANEO' END AS Area, PayrollMemorandum, PayrollFile, COUNT (PayrollCLC) AS Registros, SUM (Percepciones) AS Percepciones, SUM (Deducciones) AS Deducciones, SUM (Liquido) AS Liquido From rpt_resumen_mensual "& sExtraTables &" Where PayrollCode = '"& lPeriod &"' "& sExtraConditions & sPayrollTypeCondition &" Group by PayrollCLC, CompanyName, PayrollCode, Fortnightly, BankName, CASE WHEN AccountNumber LIKE '.' THEN 'CHEQUE' ELSE 'DEBITO' END, CASE WHEN Z.ZonePath LIKE '%,9,%' THEN 'LOCAL' ELSE 'FORANEO' END, PayrollMemorandum, PayrollFile Order by PayrollCLC"
        Else 
            sQuery = "Select PayrollCLC, CompanyName, PayrollCode, Fortnightly, BankName, CASE WHEN AccountNumber LIKE '.' THEN 'CHEQUE' ELSE 'DEBITO' END AS FPago, CASE WHEN Z.ZonePath LIKE '%,9,%' THEN 'LOCAL' ELSE 'FORANEO' END AS Area, PayrollMemorandum, PayrollFile, COUNT (PayrollCLC) AS Registros, SUM (Percepciones) AS Percepciones, SUM (Deducciones) AS Deducciones, SUM (Liquido) AS Liquido From rpt_resumen_mensual "& sExtraTables &" Where PayrollID Between "& sYear &"0101 And "& sYear &"1231 And PayrollCode In( '"& sYear &"01','"& sYear &"02', '"& sYear &"03', '"& sYear &"04', '"& sYear &"05', '"& sYear &"06') "& sExtraConditions & sPayrollTypeCondition &" Group by PayrollCLC, CompanyName, PayrollCode, Fortnightly, BankName, CASE WHEN AccountNumber LIKE '.' THEN 'CHEQUE' ELSE 'DEBITO' END, CASE WHEN Z.ZonePath LIKE '%,9,%' THEN 'LOCAL' ELSE 'FORANEO' END, PayrollMemorandum, PayrollFile Order by  PayrollCode"
        End If

        'sQuery = "Select PayrollCLC, PayrollID, CompanyID, CompanyName, PayrollCode, BankID, BankName, Count(PayrollCLC) As Registros, SUM(Percepciones) As Percepciones, SUM(Deducciones) As Deducciones, SUM(Liquido) As Liquido From rpt_resumen_mensual Where (PayrollCode='" & lPeriod & "') And (PayrollID = " & lPayrollID & ") Group By PayrollID, PayrollCLC, CompanyId, CompanyName, PayrollCode, BankID, BankName Order By PayrollCLC"
	     lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	    Response.Write vbNewLine & "<!-- Query: " & sQuery & " -->" & vbNewLine
	    If lErrorNumber = 0 Then
    	    If Not oRecordset.EOF Then
                sDate = GetSerialNumberForDate("")
			    sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			    lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
			    sFilePath = sFilePath & "\"
			    sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
			    Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			    Response.Flush()
			    sDocumentName = sFilePath & "Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"
			    
                sRowContents = sRowContents & "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">"
			    'lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
                sRowContents = sRowContents & "<TR><TD COLSPAN =""2"" style=""border-width: 1px;border: solid;""><B>"& sPayrollType &"</B></TD>"
                sRowContents = sRowContents & "<TD COLSPAN =""11""><B></B></TD></TR>"
                sRowContents = sRowContents & "<TR>"
                    sRowContents = sRowContents & "<TD><B>No.</B></TD>"
			        sRowContents = sRowContents & "<TD><B>C.L.C/LOCAL</B></TD>"
				    sRowContents = sRowContents & "<TD><B>UNIDAD</B></TD>"
				    sRowContents = sRowContents & "<TD><B>QNA.</B></TD>"
                    sRowContents = sRowContents & "<TD><B>BANCO</B></TD>"
                    sRowContents = sRowContents & "<TD><B>F/PAGO</B></TD>"
                    sRowContents = sRowContents & "<TD><B>AREA</B></TD>"
				    sRowContents = sRowContents & "<TD><B>No. MEMORANDA</B></TD>"
                    sRowContents = sRowContents & "<TD><B>ARCHIVO</B></TD>"				
				    sRowContents = sRowContents & "<TD><B>REG.</B></TD>"
				    sRowContents = sRowContents & "<TD><B>PERCEPCIONES</B></TD>"
				    sRowContents = sRowContents & "<TD><B>DEDUCCIONES</B></TD>"
				    sRowContents = sRowContents & "<TD><B>LIQUIDO</B></TD>"
			    sRowContents = sRowContents & "</TR>"
                lCount = 1
                Do While Not oRecordset.EOF
                    sRowContents = sRowContents& "<TR>"
                        sRowContents = sRowContents & "<TD>" & lCount & "</TD>"
					    sRowContents = sRowContents & "<TD>'" & CStr(oRecordset.Fields("PayrollCLC").Value) & "</TD>"
					    sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("CompanyName").Value) & "</TD>"
                        If Len(oRecordset.Fields("Fortnightly").Value)>0 Then 
                            sRowContents = sRowContents & "<TD>'" & CStr(oRecordset.Fields("Fortnightly").Value)  & "/"& oRequest("YearID").Item & "</TD>"
                        Else
                            sRowContents = sRowContents & "<TD>  </TD>"
                        End If
                    
                        sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("BankName").Value) & "</TD>"
                        sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("FPago").Value) & "</TD>"
                        sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("Area").Value) & "</TD>"
                        If Len(oRecordset.Fields("PayrollMemorandum").Value)>0 Then 
                            sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("PayrollMemorandum").Value) & "</TD>"
                        Else
					        sRowContents = sRowContents & "<TD>  </TD>"
                        End IF
                        If Len(oRecordset.Fields("PayrollFile").Value)>0 Then 
                            sRowContents = sRowContents & "<TD>" & CStr(oRecordset.Fields("PayrollFile").Value) & "</TD>"
                        Else
					        sRowContents = sRowContents & "<TD>  </TD>"
                        End IF
                        sRowContents = sRowContents & "<TD>" & CInt(oRecordset.Fields("Registros").Value) & "</TD>"
                        lReg= lReg + CInt(oRecordset.Fields("Registros").Value)
					    sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(oRecordset.Fields("Percepciones").Value),2,True,True,True) & "</TD>"
					    lPer = lPer+ CDbl(oRecordset.Fields("Percepciones").Value)
                        sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(oRecordset.Fields("Deducciones").Value),2,True,True,True) & "</TD>"
					    lDed = lDed + CDbl(oRecordset.Fields("Deducciones").Value)
                        sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(oRecordset.Fields("Liquido").Value),2,True,True,True) & "</TD>"
					    lLiq = lLiq + CDbl(oRecordset.Fields("Liquido").Value)

				    sRowContents = sRowContents & "</TR>"
				    lCount = lCount + 1
                    oRecordset.MoveNext
                Loop
                sRowContents = sRowContents &"</TABLE>"
            
			    sRowContents = sRowContents & " <br/> "
                sRowContents = sRowContents & "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
                sRowContents = sRowContents & "<TR><TD COLSPAN = 8> </TD>"
                sRowContents = sRowContents & "<TD style=""border-width: 1px;border: solid;"">TOTAL QNA.</TD>"
                sRowContents = sRowContents & "<TD style=""border-width: 1px;border: solid;"">"& lReg &"</TD>"
                sRowContents = sRowContents & "<TD style=""border-width: 1px;border: solid;"">"& FormatNumber(lPer,2,True,True,True) &"</TD>"
                sRowContents = sRowContents & "<TD style=""border-width: 1px;border: solid;"">"& FormatNumber(lDed,2,True,True,True) &"</TD>"
                sRowContents = sRowContents & "<TD style=""border-width: 1px;border: solid;"">"& FormatNumber(lLiq,2,True,True,True) &"</TD></TR>"
                sRowContents = sRowContents & "</TABLE>"
                sRowContents = sRowContents & " <br/> <br/>"
               lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
               lTotReg = lTotReg+lReg
               lTotPer = lTotPer+lPer
               lTotDed = lTotDed+lDed
               lTotLiq = lTotLiq+lLiq
			    
            End If
	    Else
			lErrorNumber = -1
			sErrorDescription = "No se encontraron documentos capturados para la quincena elegida."
	        '	End If
	    End If
        i= i+1
    Loop

    sRowContents =  "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
        sRowContents = sRowContents & "<TR><TD COLSPAN = 8> </TD>"
        sRowContents = sRowContents & "<TD style=""border-width: 1px;border: solid;"" >TOTAL GENERAL</TD>"
        sRowContents = sRowContents & "<TD style=""border-width: 1px;border: solid;"">"& lTotReg &"</TD>"
        sRowContents = sRowContents & "<TD style=""border-width: 1px;border: solid;"">"& FormatNumber(lTotPer,2,True,True,True) &"</TD>"
        sRowContents = sRowContents & "<TD style=""border-width: 1px;border: solid;"">"& FormatNumber(lTotDed,2,True,True,True) &"</TD>"
        sRowContents = sRowContents & "<TD style=""border-width: 1px;border: solid;"">"& FormatNumber(lTotLiq,2,True,True,True) &"</TD></TR>"
    sRowContents = sRowContents & "</TABLE>"
    lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)

    If lErrorNumber = 0 Then
        lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
        lErrorNumber = DeleteFolder(sFilePath, sErrorDescription)
        oEndDate = Now()
        If (lErrorNumber = 0) And B_USE_SMTP Then
		    If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
		End If
    Else
        lErrorNumber = -1
		sErrorDescription = "No se encontraron documentos capturados para la quincena elegida."
    End If

	Set oRecordset = Nothing
	BuildReport1403 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1403a(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the monthly payroll resume based on CLCs
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1403"
	Dim alConcepts
	Dim asCLCs
	Dim asCondition
	Dim asPeriods
	Dim iIndex
	Dim jIndex
	Dim sCondition
	Dim oRecordset
	Dim oRecordsetConcepts
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim sQuery
	Dim sQueryC
	Dim lPeriod
	Dim lPayrollID
	Dim sDate
	Dim sFilePath
	Dim sFileName
	Dim sDocumentName
	Dim alConcepts2
	Dim alImports
    Dim sYear

	asPeriods = Split("1,1,2,2,3,3,4,4,5,5,6,6",",")
	alConcepts2 = Split("1,4,5,6,7,8,13,47,89,146",",")
	alImports = Split("0,0,0,0,0,0,0,0,0,0",",")
	lPayrollID = oRequest("PayrollID").Item
	lPeriod = Mid(lPayrollID,1,4) & "0" & asPeriods(CInt(Mid(lPayrollID,5,2))-1)
	asCLCs = ""
	sErrorDescription = "No se pudieron obtener las CLCs generadas para la nómina especificada."
	sQuery = "Select PayrollCLC, CLC.PayrollID, EHL.CompanyID, CompanyName, PayrollCode, EHL.BankID, BankName, Count(PayrollCLC) As Regs From PayrollsCLCs As CLC, EmployeesHistoryListForPayroll As EHL, Companies As C, Banks As B Where (PayrollCode='" & lPeriod & "') And (CLC.EmployeeID = EHL.EmployeeID) And (EHL.BankID = B.BankID) And (EHL.CompanyID = C.CompanyID) And (EHL.PayrollID = " & lPayrollID & ") Group By CLC.PayrollID, PayrollCLC, EHL.CompanyId, CompanyName, PayrollCode, EHL.BankID, BankName Order By PayrollCLC"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: " & sQuery & " -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			asCLCs = oRecordset.GetRows()
			sDate = GetSerialNumberForDate("")
			sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate)
			lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
			sFilePath = sFilePath & "\"
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".zip"
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			Response.Flush()
			sDocumentName = sFilePath & "Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".xls"
			sRowContents = "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
			lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
			sRowContents = "<TR>"
				sRowContents = sRowContents & "<TD>CLC</TD>"
				sRowContents = sRowContents & "<TD>Unidad</TD>"
				sRowContents = sRowContents & "<TD>Quincena</TD>"
				sRowContents = sRowContents & "<TD>Bimestre</TD>"
				sRowContents = sRowContents & "<TD>Banco</TD>"
				sRowContents = sRowContents & "<TD>Registros</TD>"
				sRowContents = sRowContents & "<TD>Percepciones</TD>"
				sRowContents = sRowContents & "<TD>Deducciones</TD>"
				sRowContents = sRowContents & "<TD>Líquido</TD>"
				sRowContents = sRowContents & "<TD>01</TD>"
				sRowContents = sRowContents & "<TD>04</TD>"
				sRowContents = sRowContents & "<TD>05</TD>"
				sRowContents = sRowContents & "<TD>06</TD>"
				sRowContents = sRowContents & "<TD>07</TD>"
				sRowContents = sRowContents & "<TD>08</TD>"
				sRowContents = sRowContents & "<TD>11</TD>"
				sRowContents = sRowContents & "<TD>44</TD>"
				sRowContents = sRowContents & "<TD>B2</TD>"
				sRowContents = sRowContents & "<TD>7S</TD>"
			sRowContents = sRowContents & "</TR>"
			lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
            sYear = Mid(lPayrollID,1,4)
			For iIndex = 0 To UBound(asCLCs,2)
				sQuery = "Select ConceptName, Sum(ConceptAmount) Importe From Payroll_" & sYear & " Pr, Concepts C, EmployeesHistoryListForPayroll EHL Where Pr.EmployeeID In (Select EmployeeID From PayrollsCLCs Where PayrollCLC = '" & asCLCs(0,iIndex) & "') And Pr.ConceptID In (-2,-1,0) And (Pr.ConceptID = C.ConceptID) And (Pr.EmployeeID = EHL.EmployeeID) And (EHL.PayrollID = " & lPayrollID & ") And (EHL.BankID = " & asCLCs(5,iIndex) & ") And (EHL.CompanyID = " & asCLCs(2,iIndex) & ") Group By ConceptName Order By ConceptName Desc"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				alConcepts = oRecordset.GetRows()
				For jIndex = 0 To UBound(alConcepts2)
					sQueryC = "Select ConceptShortName, Sum(ConceptAmount) Importe From Payroll_" & sYear & " Pr, Concepts C, EmployeesHistoryListForPayroll EHL Where Pr.EmployeeID In (Select EmployeeID From PayrollsCLCs Where PayrollCLC = '" & asCLCs(0,iIndex) & "') And (Pr.ConceptID = " & alConcepts2(jIndex) & ") And (Pr.ConceptID = C.ConceptID) And (Pr.EmployeeID = EHL.EmployeeID) And (EHL.PayrollID = " & lPayrollID & ") And (EHL.BankID = " & asCLCs(5,iIndex) & ") And (EHL.CompanyID = " & asCLCs(2,iIndex) & ") Group By ConceptShortName"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQueryC, "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordsetConcepts)
					If oRecordsetConcepts.EOF Then
						alImports(jIndex) = 0
					Else
						alImports(jIndex) = oRecordsetConcepts.Fields("Importe").Value
					End If
				Next
				sRowContents = "<TR>"
					sRowContents = sRowContents & "<TD>'" & asCLCs(0,iIndex) & "'</TD>"
					sRowContents = sRowContents & "<TD>" & asCLCs(3,iIndex) & "</TD>"
					sRowContents = sRowContents & "<TD>" & asCLCs(1,iIndex) & "</TD>"
					sRowContents = sRowContents & "<TD>" & asCLCs(4,iIndex) & "</TD>"
					sRowContents = sRowContents & "<TD>" & asCLCs(6,iIndex) & "</TD>"
					sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(asCLCs(7,iIndex)),2,True,True,True) & "</TD>"
					sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(alConcepts(1,0)),2,True,True,True) & "</TD>"
					sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(alConcepts(1,2)),2,True,True,True) & "</TD>"
					sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(alConcepts(1,1)),2,True,True,True) & "</TD>"
					sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(alImports(0)),2,True,True,True) & "</TD>"
					sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(alImports(1)),2,True,True,True) & "</TD>"
					sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(alImports(2)),2,True,True,True) & "</TD>"
					sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(alImports(3)),2,True,True,True) & "</TD>"
					sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(alImports(4)),2,True,True,True) & "</TD>"
					sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(alImports(5)),2,True,True,True) & "</TD>"
					sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(alImports(6)),2,True,True,True) & "</TD>"
					sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(alImports(7)),2,True,True,True) & "</TD>"
					sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(alImports(8)),2,True,True,True) & "</TD>"
					sRowContents = sRowContents & "<TD>" & FormatNumber(CDbl(alImports(9)),2,True,True,True) & "</TD>"
				sRowContents = sRowContents & "</TR>"
				lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
			Next
			sRowContents = "</TABLE>"
			lErrorNumber = AppendTextToFile(sDocumentName, sRowContents, sErrorDescription)
			If lErrorNumber = 0 Then
				lErrorNumber = ZipFolder(sFilePath, Server.MapPath(sFileName), sErrorDescription)
			End If
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudieron guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', ' ')", "ReportsQueries1100Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				lErrorNumber = DeleteFolder(sFilePath, sErrorDescription)
			End If
			oEndDate = Now()
			If (lErrorNumber = 0) And B_USE_SMTP Then
				If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
			End If
		Else
			lErrorNumber = -1
			sErrorDescription = "No se encontraron documentos capturados para la quincena elegida."
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1403 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1411(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the payroll for concepts 56, 76, and 77
'         group by companies
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1411"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim adTotal
	Dim adGranTotal
	Dim lCurrentID
	Dim sCurrentName
	Dim sCurrentTypeName
	Dim dCurrentAmount
	Dim sContents
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, "", lPayrollID, lForPayrollID)
	sContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1411.htm"), sErrorDescription)
	sContents = Replace(sContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
	sContents = Replace(sContents, "<PAYROLL_NUMBER />", GetPayrollNumber(lForPayrollID))
	sContents = Replace(sContents, "<PAYROLL_YEAR />", Left(lForPayrollID, Len("0000")))
	Response.Write sContents
	Response.Write "<TABLE BORDER="""
		If Not bForExport Then
			Response.Write "0"
		Else
			Response.Write "1"
		End If
	Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
		asColumnsTitles = Split("Nómina,Concepto 77,Concepto 54,Concepto 76", ",", -1, vbBinaryCompare)
		If bForExport Then
			lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
		Else
			If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
				lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			Else
				lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			End If
		End If
		asCellAlignments = Split(",RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)

		adTotal = Split("0,0,0", ",")
		adTotal(0) = 0
		adTotal(1) = 0
		adTotal(2) = 0
		adGranTotal = Split("0,0,0", ",")
		adGranTotal(0) = 0
		adGranTotal(1) = 0
		adGranTotal(2) = 0

		sErrorDescription = "No se pudieron obtener los montos pagados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(Concept77Amount) As Total77Amount, Sum(Concept54Amount) As Total54Amount From EmployeesFONAC Where (PayrollID=" & lPayrollID & ")", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select Sum(Concept77Amount) As Total77Amount, Sum(Concept54Amount) As Total54Amount From EmployeesFONAC Where (PayrollID=" & lPayrollID & ") -->" & vbNewLine
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sRowContents = "FOVISSSTE"
				If IsNull(oRecordset.Fields("Total77Amount").Value) Then
					sRowContents = sRowContents & TABLE_SEPARATOR & "0.00"
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Total77Amount").Value), 2, True, False, True)
				End If
				If IsNull(oRecordset.Fields("Total54Amount").Value) Then
					sRowContents = sRowContents & TABLE_SEPARATOR & "0.00"
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("Total54Amount").Value), 2, True, False, True)
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & "0.00"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				adGranTotal(0) = adGranTotal(0) + CDbl(oRecordset.Fields("Total77Amount").Value)
				adGranTotal(1) = adGranTotal(1) + CDbl(oRecordset.Fields("Total54Amount").Value)
				adGranTotal(2) = adGranTotal(2) + 0
				oRecordset.Close
			End If
		End If

		'sCondition = " And (Payroll_" & lPayrollID & ".ConceptID In (76,77))"
		sErrorDescription = "No se pudieron obtener los montos pagados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Companies.CompanyID, CompanyName, ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryListForPayroll, Companies Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryListForPayroll.EmployeeDate) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID In (76,77)) Group By Companies.CompanyID, CompanyName, ConceptID Union All Select Companies.CompanyID, CompanyName, ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryListForPayroll, Companies Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryListForPayroll.EmployeeDate) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID = 56) And (Payroll_" & lPayrollID & ".EmployeeID IN (Select Distinct EmployeeID From Payroll_" & lPayrollID & " Where ConceptID = 77)) Group By Companies.CompanyID, CompanyName, ConceptID  Order By CompanyName, ConceptID", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select Companies.CompanyID, CompanyName, ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryListForPayroll, Companies Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryListForPayroll.EmployeeDate) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID In (76,77)) Group By Companies.CompanyID, CompanyName, ConceptID Union All Select Companies.CompanyID, CompanyName, ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryListForPayroll, Companies Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryListForPayroll.EmployeeDate) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (Payroll_" & lPayrollID & ".ConceptID = 56) And (Payroll_" & lPayrollID & ".EmployeeID IN (Select Distinct EmployeeID From Payroll_" & lPayrollID & " Where ConceptID = 77)) Group By Companies.CompanyID, CompanyName, ConceptID  Order By CompanyName, ConceptID -->" & vbNewLine
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				lCurrentID = CLng(oRecordset.Fields("CompanyID").Value)
				sCurrentName = CStr(oRecordset.Fields("CompanyName").Value)
				Do While Not oRecordset.EOF
					If lCurrentID <> CLng(oRecordset.Fields("CompanyID").Value) Then
						sRowContents = CleanStringForHTML(sCurrentName)
						sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotal(0), 2, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotal(1), 2, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotal(2), 2, True, False, True)
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						adGranTotal(0) = adGranTotal(0) + adTotal(0)
						adGranTotal(1) = adGranTotal(1) + adTotal(1)
						adGranTotal(2) = adGranTotal(2) + adTotal(2)
						adTotal(0) = 0
						adTotal(1) = 0
						adTotal(2) = 0
						lCurrentID = CLng(oRecordset.Fields("CompanyID").Value)
						sCurrentName = CStr(oRecordset.Fields("CompanyName").Value)
					End If
					Select Case CLng(oRecordset.Fields("ConceptID").Value)
						Case 56
							adTotal(1) = CDbl(oRecordset.Fields("TotalAmount").Value)
						Case 76
							adTotal(2) = CDbl(oRecordset.Fields("TotalAmount").Value)
						Case 77
							adTotal(0) = CDbl(oRecordset.Fields("TotalAmount").Value)
					End Select

					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close

				sRowContents = CleanStringForHTML(sCurrentName)
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotal(0), 2, True, False, True)
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotal(1), 2, True, False, True)
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotal(2), 2, True, False, True)
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				adGranTotal(0) = adGranTotal(0) + adTotal(0)
				adGranTotal(1) = adGranTotal(1) + adTotal(1)
				adGranTotal(2) = adGranTotal(2) + adTotal(2)

				sRowContents = "<B>TOTAL</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adGranTotal(0), 2, True, False, True) & "</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adGranTotal(1), 2, True, False, True) & "</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adGranTotal(2), 2, True, False, True) & "</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			End If
		End If
	Response.Write "</TABLE><BR />"

	Response.Write "<TABLE BORDER="""
		If Not bForExport Then
			Response.Write "0"
		Else
			Response.Write "1"
		End If
	Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
		asColumnsTitles = Split("Nómina,Concepto 77,Registros,Total", ",", -1, vbBinaryCompare)
		If bForExport Then
			lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
		Else
			If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
				lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			Else
				lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			End If
		End If
		asCellAlignments = Split(",RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)

		adTotal = Split("0,0,0", ",")
		adTotal(0) = 0
		adTotal(1) = 0
		adTotal(2) = 0
		adGranTotal = Split("0,0,0", ",")
		adGranTotal(0) = 0
		adGranTotal(1) = 0
		adGranTotal(2) = 0

		sErrorDescription = "No se pudieron obtener los montos pagados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, Sum(Concept77Amount) As TotalAmount From EmployeesFONAC Where (PayrollID=" & lPayrollID & ") Group By EmployeeID Order By Sum(Concept77Amount)", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select EmployeeID, Sum(Concept77Amount) As TotalAmount From EmployeesFONAC Where (PayrollID=" & lPayrollID & ") Group By EmployeeID Order By Sum(Concept77Amount) -->" & vbNewLine
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				lCurrentID = -2
				dCurrentAmount = CDbl(oRecordset.Fields("TotalAmount").Value)
				asRowContents = Split("<SPAN COLS=""4"" />FOVISSSTE", TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				Do While Not oRecordset.EOF
					If dCurrentAmount <> CDbl(oRecordset.Fields("TotalAmount").Value) Then
						sRowContents = "&nbsp;" & TABLE_SEPARATOR & FormatNumber(dCurrentAmount, 2, True, False, True) & TABLE_SEPARATOR & FormatNumber(adTotal(1), 0, True, False, True) & TABLE_SEPARATOR & FormatNumber((dCurrentAmount * adTotal(1)), 2, True, False, True)
						adGranTotal(2) = adGranTotal(2) + (dCurrentAmount * adTotal(1))
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						dCurrentAmount = CDbl(oRecordset.Fields("TotalAmount").Value)
						adTotal(1) = 0
					End If
					adTotal(1) = adTotal(1) + 1
					adGranTotal(1) = adGranTotal(1) + 1

					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close

				sRowContents = "&nbsp;" & TABLE_SEPARATOR & FormatNumber(dCurrentAmount, 2, True, False, True) & TABLE_SEPARATOR & FormatNumber(adTotal(1), 0, True, False, True) & TABLE_SEPARATOR & FormatNumber((dCurrentAmount * adTotal(1)), 2, True, False, True)
				adGranTotal(2) = adGranTotal(2) + (dCurrentAmount * adTotal(1))
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			End If
		End If

		sCondition = " And (Payroll_" & lPayrollID & ".ConceptID = 77)"
		sErrorDescription = "No se pudieron obtener los montos pagados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Companies.CompanyID, CompanyName, Payroll_" & lPayrollID & ".EmployeeID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryListForPayroll, Companies Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryListForPayroll.EmployeeDate) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Companies.CompanyID, CompanyName, Payroll_" & lPayrollID & ".EmployeeID Order By CompanyName, Sum(ConceptAmount)", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select Companies.CompanyID, CompanyName, Payroll_" & lPayrollID & ".EmployeeID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryListForPayroll, Companies Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryListForPayroll.EmployeeDate) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Companies.CompanyID, CompanyName, Payroll_" & lPayrollID & ".EmployeeID Order By CompanyName, Sum(ConceptAmount) -->" & vbNewLine
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				lCurrentID = -2
				dCurrentAmount = CDbl(oRecordset.Fields("TotalAmount").Value)
				Do While Not oRecordset.EOF
					If lCurrentID <> CLng(oRecordset.Fields("CompanyID").Value) Then
						If lCurrentID <> -2 Then
							sRowContents = "&nbsp;" & TABLE_SEPARATOR & FormatNumber(dCurrentAmount, 2, True, False, True) & TABLE_SEPARATOR & FormatNumber(adTotal(1), 0, True, False, True) & TABLE_SEPARATOR & FormatNumber((dCurrentAmount * adTotal(1)), 2, True, False, True)
							adGranTotal(2) = adGranTotal(2) + (dCurrentAmount * adTotal(1))
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
							dCurrentAmount = CDbl(oRecordset.Fields("TotalAmount").Value)
						End If

						asRowContents = Split("<SPAN COLS=""4"" />" & CleanStringForHTML(CStr(oRecordset.Fields("CompanyName").Value)), TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						lCurrentID = CLng(oRecordset.Fields("CompanyID").Value)
						adTotal(1) = 0
					ElseIf dCurrentAmount <> CDbl(oRecordset.Fields("TotalAmount").Value) Then
						sRowContents = "&nbsp;" & TABLE_SEPARATOR & FormatNumber(dCurrentAmount, 2, True, False, True) & TABLE_SEPARATOR & FormatNumber(adTotal(1), 0, True, False, True) & TABLE_SEPARATOR & FormatNumber((dCurrentAmount * adTotal(1)), 2, True, False, True)
						adGranTotal(2) = adGranTotal(2) + (dCurrentAmount * adTotal(1))
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						dCurrentAmount = CDbl(oRecordset.Fields("TotalAmount").Value)
						adTotal(1) = 0
					End If
					adTotal(1) = adTotal(1) + 1
					adGranTotal(1) = adGranTotal(1) + 1

					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
				sRowContents = "&nbsp;" & TABLE_SEPARATOR & FormatNumber(dCurrentAmount, 2, True, False, True) & TABLE_SEPARATOR & FormatNumber(adTotal(1), 0, True, False, True) & TABLE_SEPARATOR & FormatNumber((dCurrentAmount * adTotal(1)), 2, True, False, True)
				adGranTotal(2) = adGranTotal(2) + (dCurrentAmount * adTotal(1))
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If

				sRowContents = "<B>TOTAL</B>" & TABLE_SEPARATOR & "<CENTER><B>---</B></CENTER>" & TABLE_SEPARATOR & "<B>" & FormatNumber(adGranTotal(1), 0, True, False, True) & "</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(adGranTotal(2), 2, True, False, True) & "</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			End If
		End If
	Response.Write "</TABLE><BR />"

	Response.Write "<TABLE BORDER="""
		If Not bForExport Then
			Response.Write "0"
		Else
			Response.Write "1"
		End If
	Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
		asColumnsTitles = Split("Nómina,Concepto 77,Concepto 54,Registros", ",", -1, vbBinaryCompare)
		If bForExport Then
			lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
		Else
			If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
				lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			Else
				lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			End If
		End If
		asCellAlignments = Split(",RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)

		adTotal = Split("0,0,0", ",")
		adTotal(0) = 0
		adTotal(1) = 0
		adTotal(2) = 0
		adGranTotal = Split(",", ",")
		adGranTotal(0) = Split("0,0,0", ",")
		adGranTotal(0)(0) = 0
		adGranTotal(0)(1) = 0
		adGranTotal(0)(2) = 0
		adGranTotal(1) = Split("0,0,0", ",")
		adGranTotal(1)(0) = 0
		adGranTotal(1)(1) = 0
		adGranTotal(1)(2) = 0

		sErrorDescription = "No se pudieron obtener los montos pagados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PositionTypes.PositionTypeID, PositionTypeName, EmployeeID, Sum(Concept77Amount) As Total77Amount, Sum(Concept54Amount) As Total54Amount From EmployeesFONAC, PositionTypes Where (EmployeesFONAC.PositionTypeID=PositionTypes.PositionTypeID) And (PayrollID=" & lPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") Group By PositionTypes.PositionTypeID, PositionTypeName, EmployeeID Order By PositionTypeName, EmployeeID", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select PositionTypes.PositionTypeID, PositionTypeName, EmployeeID, Sum(Concept77Amount) As Total77Amount, Sum(Concept54Amount) As Total54Amount From EmployeesFONAC, PositionTypes Where (EmployeesFONAC.PositionTypeID=PositionTypes.PositionTypeID) And (PayrollID=" & lPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") Group By PositionTypes.PositionTypeID, PositionTypeName, EmployeeID Order By PositionTypeName, EmployeeID -->" & vbNewLine
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				lCurrentID = CLng(oRecordset.Fields("PositionTypeID").Value)
				sCurrentTypeName = CStr(oRecordset.Fields("PositionTypeName").Value)
				sRowContents = "<SPAN COLS=""4"" /><B>FOVISSSTE</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				Do While Not oRecordset.EOF
					If lCurrentID <> CLng(oRecordset.Fields("PositionTypeID").Value) Then
						sRowContents = "&nbsp;&nbsp;&nbsp;" & CleanStringForHTML(sCurrentTypeName)
						sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotal(0), 2, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotal(1), 2, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotal(2), 0, True, False, True)
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						adGranTotal(0)(0) = adGranTotal(0)(0) + adTotal(0)
						adGranTotal(0)(1) = adGranTotal(0)(1) + adTotal(1)
						adGranTotal(0)(2) = adGranTotal(0)(2) + adTotal(2)
						adTotal(0) = 0
						adTotal(1) = 0
						adTotal(2) = 0

						lCurrentID = CLng(oRecordset.Fields("PositionTypeID").Value)
						sCurrentTypeName = CStr(oRecordset.Fields("PositionTypeName").Value)
					End If
					adTotal(1) = adTotal(1) + CDbl(oRecordset.Fields("Total54Amount").Value)
					adTotal(0) = adTotal(0) + CDbl(oRecordset.Fields("Total77Amount").Value)
					adTotal(2) = adTotal(2) + 1

					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close

				sRowContents = "&nbsp;&nbsp;&nbsp;" & CleanStringForHTML(sCurrentTypeName)
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotal(0), 2, True, False, True)
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotal(1), 2, True, False, True)
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotal(2), 0, True, False, True)
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				adGranTotal(0)(0) = adGranTotal(0)(0) + adTotal(0)
				adGranTotal(0)(1) = adGranTotal(0)(1) + adTotal(1)
				adGranTotal(0)(2) = adGranTotal(0)(2) + adTotal(2)
				adTotal(0) = 0
				adTotal(1) = 0
				adTotal(2) = 0

				sRowContents = "<B>TOTAL</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adGranTotal(0)(0), 2, True, False, True)
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adGranTotal(0)(1), 2, True, False, True)
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adGranTotal(0)(2), 0, True, False, True)
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				adGranTotal(1)(0) = adGranTotal(1)(0) + adGranTotal(0)(0)
				adGranTotal(1)(1) = adGranTotal(1)(1) + adGranTotal(0)(1)
				adGranTotal(1)(2) = adGranTotal(1)(2) + adGranTotal(0)(2)
				adGranTotal(0)(0) = 0
				adGranTotal(0)(1) = 0
				adGranTotal(0)(2) = 0
			End If
		End If

		sCondition = " And (Payroll_" & lPayrollID & ".ConceptID In (56,77)) And (Payroll_" & lPayrollID & ".EmployeeID In (Select Distinct EmployeeID From Payroll_" & lPayrollID & " Where ConceptID = 77))"
		sErrorDescription = "No se pudieron obtener los montos pagados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Companies.CompanyID, CompanyName, PositionTypes.PositionTypeID, PositionTypeName, Payroll_" & lPayrollID & ".EmployeeID, ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryListForPayroll, Companies, PositionTypes Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryListForPayroll.EmployeeDate) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Companies.CompanyID, CompanyName, PositionTypes.PositionTypeID, PositionTypeName, Payroll_" & lPayrollID & ".EmployeeID, ConceptID Order By CompanyName, PositionTypeName, EmployeeID, ConceptID", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select Companies.CompanyID, CompanyName, PositionTypes.PositionTypeID, PositionTypeName, Payroll_" & lPayrollID & ".EmployeeID, ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryListForPayroll, Companies, PositionTypes Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryListForPayroll.EmployeeDate) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") And (PositionTypes.StartDate<=" & lForPayrollID & ") And (PositionTypes.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Companies.CompanyID, CompanyName, PositionTypes.PositionTypeID, PositionTypeName, Payroll_" & lPayrollID & ".EmployeeID, ConceptID Order By CompanyName, PositionTypeName, EmployeeID, ConceptID -->" & vbNewLine
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				lCurrentID = (CLng(oRecordset.Fields("CompanyID").Value) * 10) + CLng(oRecordset.Fields("PositionTypeID").Value)
				sCurrentName = CStr(oRecordset.Fields("CompanyName").Value)
				sCurrentTypeName = CStr(oRecordset.Fields("PositionTypeName").Value)
				sRowContents = "<SPAN COLS=""4"" /><B>" & CleanStringForHTML(sCurrentName) & "</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				Do While Not oRecordset.EOF
					If lCurrentID <> (CLng(oRecordset.Fields("CompanyID").Value) * 10) + CLng(oRecordset.Fields("PositionTypeID").Value) Then
						sRowContents = "&nbsp;&nbsp;&nbsp;" & CleanStringForHTML(sCurrentTypeName)
						sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotal(0), 2, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotal(1), 2, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotal(2), 0, True, False, True)
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						adGranTotal(0)(0) = adGranTotal(0)(0) + adTotal(0)
						adGranTotal(0)(1) = adGranTotal(0)(1) + adTotal(1)
						adGranTotal(0)(2) = adGranTotal(0)(2) + adTotal(2)
						adTotal(0) = 0
						adTotal(1) = 0
						adTotal(2) = 0

						If StrComp(sCurrentName, CStr(oRecordset.Fields("CompanyName").Value), vbBinaryCompare) <> 0 Then
							sRowContents = "<B>TOTAL</B>"
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adGranTotal(0)(0), 2, True, False, True)
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adGranTotal(0)(1), 2, True, False, True)
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adGranTotal(0)(2), 0, True, False, True)
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
							adGranTotal(1)(0) = adGranTotal(1)(0) + adGranTotal(0)(0)
							adGranTotal(1)(1) = adGranTotal(1)(1) + adGranTotal(0)(1)
							adGranTotal(1)(2) = adGranTotal(1)(2) + adGranTotal(0)(2)
							adGranTotal(0)(0) = 0
							adGranTotal(0)(1) = 0
							adGranTotal(0)(2) = 0

							sCurrentName = CStr(oRecordset.Fields("CompanyName").Value)
							sRowContents = "<SPAN COLS=""4"" /><B>" & CleanStringForHTML(sCurrentName) & "</B>"
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
						End If

						lCurrentID = (CLng(oRecordset.Fields("CompanyID").Value) * 10) + CLng(oRecordset.Fields("PositionTypeID").Value)
						sCurrentTypeName = CStr(oRecordset.Fields("PositionTypeName").Value)
					End If
					Select Case CLng(oRecordset.Fields("ConceptID").Value)
						Case 56
							adTotal(1) = adTotal(1) + CDbl(oRecordset.Fields("TotalAmount").Value)
						Case 76
							adTotal(0) = adTotal(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
						Case 77
							adTotal(0) = adTotal(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
							adTotal(2) = adTotal(2) + 1
					End Select

					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close

				sRowContents = "&nbsp;&nbsp;&nbsp;" & CleanStringForHTML(sCurrentTypeName)
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotal(0), 2, True, False, True)
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotal(1), 2, True, False, True)
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotal(2), 0, True, False, True)
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				adGranTotal(0)(0) = adGranTotal(0)(0) + adTotal(0)
				adGranTotal(0)(1) = adGranTotal(0)(1) + adTotal(1)
				adGranTotal(0)(2) = adGranTotal(0)(2) + adTotal(2)

				sRowContents = "<B>TOTAL</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adGranTotal(0)(0), 2, True, False, True)
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adGranTotal(0)(1), 2, True, False, True)
				sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adGranTotal(0)(2), 0, True, False, True)
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				adGranTotal(1)(0) = adGranTotal(1)(0) + adGranTotal(0)(0)
				adGranTotal(1)(1) = adGranTotal(1)(1) + adGranTotal(0)(1)
				adGranTotal(1)(2) = adGranTotal(1)(2) + adGranTotal(0)(2)

				sRowContents = "<B>TOTAL</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adGranTotal(1)(0), 2, True, False, True) & "</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adGranTotal(1)(1), 2, True, False, True) & "</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adGranTotal(1)(2), 0, True, False, True) & "</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			End If
		End If
	Response.Write "</TABLE><BR />"

	Set oRecordset = Nothing
	BuildReport1411 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1413(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the payroll for concepts 56, 76, and 77
'         as a coma-separated text file, group by area
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1413"
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
	Dim lCurrentID
	Dim dTotal
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, "", lPayrollID, lForPayrollID)
	lPayrollNumber = (CInt(Left(lForPayrollID, Len("0000"))) * 100) + CInt(GetPayrollNumber(lForPayrollID))
	oStartDate = Now()
	sDate = GetSerialNumberForDate("")
	sFilePath = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".txt"
	Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(Replace(sFilePath, ".txt", ".zip")) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
	sFilePath = Server.MapPath(sFilePath)
	Response.Flush()

	sErrorDescription = "No se pudieron obtener los montos pagados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Areas.AreaID, AreaShortName, AreaName, Sum(Concept77Amount) As Total77Amount, Sum(Concept54Amount) As Total54Amount From EmployeesFONAC, Areas Where (EmployeesFONAC.AreaID=Areas.AreaID) And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PayrollID=" & lPayrollID & ") Group By Areas.AreaID, AreaShortName, AreaName Order By AreaShortName, AreaName", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select Areas.AreaID, AreaShortName, AreaName, Sum(Concept77Amount) As Total77Amount, Sum(Concept54Amount) As Total54Amount From EmployeesFONAC, Areas Where (EmployeesFONAC.AreaID=Areas.AreaID) And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (PayrollID=" & lPayrollID & ") Group By Areas.AreaID, AreaShortName, AreaName Order By AreaShortName, AreaName -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			lCurrentID = -2
			dTotal = 0
			Do While Not oRecordset.EOF
				sRowContents = Right(("0000000000" & CStr(oRecordset.Fields("AreaShortName").Value)), Len("0000000000"))
				sRowContents = sRowContents & SizeText(CStr(oRecordset.Fields("AreaName").Value), " ", 30, 1)
				sRowContents = sRowContents & Right(("00000000" & FormatNumber((CDbl(oRecordset.Fields("Total77Amount").Value) * 0.15), 2, True, False, False)), Len("00000000.00"))
				sRowContents = sRowContents & Right(("00000000" & FormatNumber(CDbl(oRecordset.Fields("Total77Amount").Value), 2, True, False, False)), Len("00000000.00"))
				sRowContents = sRowContents & Right(("00000000" & FormatNumber(CDbl(oRecordset.Fields("Total54Amount").Value), 2, True, False, False)), Len("00000000.00"))
				sRowContents = sRowContents & Right(("00000000" & FormatNumber((CDbl(oRecordset.Fields("Total54Amount").Value) * 0.25), 2, True, False, False)), Len("00000000.00"))
				sRowContents = sRowContents & lForPayrollID
				sRowContents = sRowContents & Right(lForPayrollID, Len("00"))
				lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)

				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
		End If
	End If

'	sCondition = " And (Concepts.ConceptID In (56,77))"
	sCondition = " And (Payroll_" & lPayrollID & ".ConceptID In (56,77)) And (Payroll_" & lPayrollID & ".EmployeeID In (Select Distinct EmployeeID From Payroll_" & lPayrollID & " Where ConceptID = 77))"
	sErrorDescription = "No se pudieron obtener los montos pagados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Areas.AreaID, AreaShortName, AreaName, Concepts.ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesChangesLKP, EmployeesHistoryListForPayroll, Areas Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryListForPayroll.EmployeeDate) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") " & sCondition & " Group By Areas.AreaID, AreaShortName, AreaName, Concepts.ConceptID Order By AreaShortName, AreaName", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select Areas.AreaID, AreaShortName, AreaName, Concepts.ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesChangesLKP, EmployeesHistoryListForPayroll, Areas Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryListForPayroll.EmployeeDate) And (EmployeesHistoryListForPayroll.AreaID=Areas.AreaID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") " & sCondition & " Group By Areas.AreaID, AreaShortName, AreaName, Concepts.ConceptID Order By AreaShortName, AreaName -->" & vbNewLine
'	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Areas.AreaID, AreaShortName, AreaName, Concepts.ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesChangesLKP, EmployeesHistoryList, Areas Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") " & sCondition & " Group By Areas.AreaID, AreaShortName, AreaName, Concepts.ConceptID Order By AreaShortName, AreaName", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
'	Response.Write vbNewLine & "<!-- Query: Select Areas.AreaID, AreaShortName, AreaName, Concepts.ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesChangesLKP, EmployeesHistoryList, Areas Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (Areas.StartDate<=" & lForPayrollID & ") And (Areas.EndDate>=" & lForPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") " & sCondition & " Group By Areas.AreaID, AreaShortName, AreaName, Concepts.ConceptID Order By AreaShortName, AreaName -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			lCurrentID = -2
			dTotal = 0
			Do While Not oRecordset.EOF
				If lCurrentID <> CLng(oRecordset.Fields("AreaID").Value) Then
					If lCurrentID <> -2 Then
						sRowContents = Replace(sRowContents, "<CONCEPT_77_15 />", Right(("00000000" & FormatNumber((dTotal * 0.15), 2, True, False, False)), Len("00000000.00")))
						sRowContents = Replace(sRowContents, "<CONCEPT_77 />", Right(("00000000" & FormatNumber(dTotal, 2, True, False, False)), Len("00000000.00")))
						sRowContents = Replace(sRowContents, "<CONCEPT_56 />", "00000000.00")
						sRowContents = Replace(sRowContents, "<CONCEPT_56_25 />", "00000000.00")
						lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
					End If
					sRowContents = Right(("0000000000" & CStr(oRecordset.Fields("AreaShortName").Value)), Len("0000000000"))
					sRowContents = sRowContents & SizeText(CStr(oRecordset.Fields("AreaName").Value), " ", 30, 1)
					sRowContents = sRowContents & "<CONCEPT_77_15 />"
					sRowContents = sRowContents & "<CONCEPT_77 />"
					sRowContents = sRowContents & "<CONCEPT_56 />"
					sRowContents = sRowContents & "<CONCEPT_56_25 />"
					sRowContents = sRowContents & lForPayrollID
					sRowContents = sRowContents & Right(lForPayrollID, Len("00"))

					lCurrentID = CLng(oRecordset.Fields("AreaID").Value)
					dTotal = 0
				End If
				If InStr(1, ",77,76,", "," & CStr(oRecordset.Fields("ConceptID").Value) & ",", vbBinaryCompare) > 0 Then
					dTotal = dTotal + CDbl(oRecordset.Fields("TotalAmount").Value)
				Else
					sRowContents = Replace(sRowContents, "<CONCEPT_56 />", Right(("00000000" & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, False)), Len("00000000.00")))
					sRowContents = Replace(sRowContents, "<CONCEPT_56_25 />", Right(("00000000" & FormatNumber((CDbl(oRecordset.Fields("TotalAmount").Value) * 0.25), 2, True, False, False)), Len("00000000.00")))
				End If

				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			sRowContents = Replace(sRowContents, "<CONCEPT_77_15 />", Right(("00000000" & FormatNumber((dTotal * 0.15), 2, True, False, False)), Len("00000000.00")))
			sRowContents = Replace(sRowContents, "<CONCEPT_77 />", Right(("00000000" & FormatNumber(dTotal, 2, True, False, False)), Len("00000000.00")))
			sRowContents = Replace(sRowContents, "<CONCEPT_56 />", "00000000.00")
			sRowContents = Replace(sRowContents, "<CONCEPT_56_25 />", "00000000.00")
			lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
			oRecordset.Close
		End If
	End If
	If FileExists(sFilePath, sErrorDescription) Then
		lErrorNumber = ZipFile(sFilePath, Replace(sFilePath, ".txt", ".zip"), sErrorDescription)
		If lErrorNumber = 0 Then
			Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
			sErrorDescription = "No se pudieron guardar la información del reporte."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			Response.Write vbNewLine & "<!-- Query: Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '') -->" & vbNewLine
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
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "window.CheckFileIFrame.location.href = 'CheckFile.asp?bNoReport=1';" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	Set oRecordset = Nothing
	BuildReport1413 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1414(oRequest, oADODBConnection, bFull, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the payroll for concepts 56, 76, and 77
'         as a coma-separated text file
'Inputs:  oRequest, oADODBConnection, bFull, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1414"
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
	Dim lCurrentID
	Dim dTotal
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Const FONAC_FACTOR = 0.25

	Call GetConditionFromURL(oRequest, "", lPayrollID, lForPayrollID)
	lPayrollNumber = (CInt(Left(lForPayrollID, Len("0000"))) * 100) + CInt(GetPayrollNumber(lForPayrollID))
	oStartDate = Now()
	sDate = GetSerialNumberForDate("")
	sFilePath = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".txt"
	Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(Replace(sFilePath, ".txt", ".zip")) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
	sFilePath = Server.MapPath(sFilePath)
	Response.Flush()

'	sCondition = " And (Concepts.ConceptID In (56,77))"
	sCondition = " And (Payroll_" & lPayrollID & ".ConceptID In (56,77)) And (Payroll_" & lPayrollID & ".EmployeeID In (Select Distinct EmployeeID From Payroll_" & lPayrollID & " Where ConceptID = 77))"
	sErrorDescription = "No se pudieron obtener los montos pagados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case EmployeeLastName2 When Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, Case RFC When Null Then ' ' Else RFC End RFC, Concepts.ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesChangesLKP, EmployeesHistoryListForPayroll, Employees Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryListForPayroll.EmployeeDate) And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") " & sCondition & " Group By EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Concepts.ConceptID Order By EmployeeNumber", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, Case EmployeeLastName2 When Null Then ' ' Else EmployeeLastName2 End EmployeeLastName2, Case RFC When Null Then ' ' Else RFC End RFC, Concepts.ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesChangesLKP, EmployeesHistoryListForPayroll, Employees Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryListForPayroll.EmployeeDate) And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") " & sCondition & " Group By EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, Concepts.ConceptID Order By EmployeeNumber -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			lCurrentID = -2
			dTotal = 0
			Do While Not oRecordset.EOF
				If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
					If lCurrentID <> -2 Then
						If bFull Then
							sRowContents = Replace(sRowContents, "<CONCEPT_77_76 />", FormatNumber(dTotal, 2, True, False, False))
							sRowContents = Replace(sRowContents, "<CONCEPT_77_125 />", "0.00")
							sRowContents = Replace(sRowContents, "<CONCEPT_77_25 />", "0.00")
							sRowContents = Replace(sRowContents, "<CONCEPT_54_25 />", "0.00")
						End If
						lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
					End If
					sRowContents = "I"
					sRowContents = sRowContents & "," & CStr(oRecordset.Fields("EmployeeNumber").Value)
					If bFull Then
						sRowContents = sRowContents & "," & lPayrollNumber
						sRowContents = sRowContents & "," & "<CONCEPT_77_76 />"
						sRowContents = sRowContents & "," & "<CONCEPT_77_125 />"
						sRowContents = sRowContents & "," & "<CONCEPT_77_25 />"
						sRowContents = sRowContents & "," & "<CONCEPT_54_25 />"
					End If
					sTemp = CStr(oRecordset.Fields("EmployeeName").Value)
					sTemp = sTemp & " " & CStr(oRecordset.Fields("EmployeeLastName").Value)
					If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then sTemp = sTemp & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value)
					Err.Clear
					sRowContents = sRowContents & "," & sTemp
					sRowContents = sRowContents & "," & CStr(oRecordset.Fields("RFC").Value)
					sRowContents = sRowContents & ","

					lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
					dTotal = 0
				End If
				If bFull Then
					Select Case CLng(oRecordset.Fields("ConceptID").Value)
						Case 56
							sRowContents = Replace(sRowContents, "<CONCEPT_54_25 />", FormatNumber((CDbl(oRecordset.Fields("TotalAmount").Value) * FONAC_FACTOR), 2, True, False, False))
						Case 76
							dTotal = dTotal + CDbl(oRecordset.Fields("TotalAmount").Value)
						Case 77
							sRowContents = Replace(sRowContents, "<CONCEPT_77_125 />", FormatNumber((CDbl(oRecordset.Fields("TotalAmount").Value) * (1 + FONAC_FACTOR)), 2, True, False, False))
							sRowContents = Replace(sRowContents, "<CONCEPT_77_25 />", FormatNumber((CDbl(oRecordset.Fields("TotalAmount").Value) * FONAC_FACTOR), 2, True, False, False))
							dTotal = dTotal + CDbl(oRecordset.Fields("TotalAmount").Value)
					End Select
				End If

				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			If bFull Then
				sRowContents = Replace(sRowContents, "<CONCEPT_77_76 />", FormatNumber(dTotal, 2, True, False, False))
				sRowContents = Replace(sRowContents, "<CONCEPT_77_125 />", "0.00")
				sRowContents = Replace(sRowContents, "<CONCEPT_77_25 />", "0.00")
				sRowContents = Replace(sRowContents, "<CONCEPT_54_25 />", "0.00")
			End If
			lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
			oRecordset.Close
		End If
	End If

	sErrorDescription = "No se pudieron obtener los montos pagados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesFONAC Where (PayrollID=" & lPayrollID & ") Order By EmployeeNumber", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select * From EmployeesFONAC Where (PayrollID=" & lPayrollID & ") Order By EmployeeNumber -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Do While Not oRecordset.EOF
				sRowContents = "I"
				sRowContents = sRowContents & "," & CStr(oRecordset.Fields("EmployeeNumber").Value)
				If bFull Then
					sRowContents = sRowContents & "," & lPayrollNumber
					sRowContents = sRowContents & "," & CStr(oRecordset.Fields("Concept77Amount").Value)
					sRowContents = sRowContents & "," & CDbl(oRecordset.Fields("Concept77Amount").Value) + CDbl(oRecordset.Fields("Concept54Amount").Value)
					sRowContents = sRowContents & "," & CStr(oRecordset.Fields("Concept54Amount").Value)
					sRowContents = sRowContents & "," & 0.00
				End If
				sRowContents = sRowContents & "," & CStr(oRecordset.Fields("EmployeeName").Value)
				sRowContents = sRowContents & "," & CStr(oRecordset.Fields("EmployeeRFC").Value)
				sRowContents = sRowContents & ","
				lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
		End If
	End If

	If FileExists(sFilePath, sErrorDescription) Then
		lErrorNumber = ZipFile(sFilePath, Replace(sFilePath, ".txt", ".zip"), sErrorDescription)
		If lErrorNumber = 0 Then
			Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
			sErrorDescription = "No se pudieron guardar la información del reporte."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
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
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "window.CheckFileIFrame.location.href = 'CheckFile.asp?bNoReport=1';" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	Set oRecordset = Nothing
	BuildReport1414 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1415(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the payroll for concepts 56, 76, and 77
'         as a coma-separated text file
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1415"
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
	Dim lCurrentID
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, "", lPayrollID, lForPayrollID)
	lPayrollNumber = Right(("00" & GetPayrollNumber(lForPayrollID)), Len("00")) & Left(lForPayrollID, Len("0000"))
	sCondition = " And (Concepts.ConceptID In (56,77))"
	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
		sCondition = sCondition & " And ((EmployeesHistoryList.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (EmployeesHistoryList.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
	End If
	oStartDate = Now()
	sErrorDescription = "No se pudieron obtener los montos pagados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, AreaCode, Concepts.ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesChangesLKP, EmployeesHistoryList, Employees, Areas Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, AreaCode, Concepts.ConceptID Order By EmployeesHistoryList.EmployeeNumber", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, AreaCode, Concepts.ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", Concepts, EmployeesChangesLKP, EmployeesHistoryList, Employees, Areas Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Concepts.StartDate<=" & lForPayrollID & ") And (Concepts.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, AreaCode, Concepts.ConceptID Order By EmployeesHistoryList.EmployeeNumber -->" & vbNewLine

	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFilePath = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate & ".txt"
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(Replace(sFilePath, ".txt", ".zip")) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			sFilePath = Server.MapPath(sFilePath)
			Response.Flush()

			lCurrentID = -2
			dTotal = 0
			Do While Not oRecordset.EOF
				If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
					If lCurrentID <> -2 Then
						sRowContents = Replace(sRowContents, "<CONCEPT_77 />", "000000")
						sRowContents = Replace(sRowContents, "<CONCEPT_56 />", "C000000")
						lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
					End If
					sRowContents = SizeText(CStr(oRecordset.Fields("EmployeeNumber").Value), " ", 6, 1)
					sRowContents = sRowContents & SizeText(CStr(oRecordset.Fields("RFC").Value), " ", 10, 1)
					sTemp = CStr(oRecordset.Fields("EmployeeName").Value)
					sTemp = sTemp & " " & CStr(oRecordset.Fields("EmployeeLastName").Value)
					If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then sTemp = sTemp & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value)
					Err.Clear
					sRowContents = sRowContents & SizeText(sTemp, " ", 50, 1) 'Nombre
					sRowContents = sRowContents & Right(("000000" & CStr(oRecordset.Fields("AreaCode").Value)), Len("000000"))
					sRowContents = sRowContents & "A00000000"
					sRowContents = sRowContents & lPayrollNumber
					sRowContents = sRowContents & "0000000000"
					sRowContents = sRowContents & "<CONCEPT_77 />"
					sRowContents = sRowContents & "<CONCEPT_56 />"

					lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
				End If
				Select Case CLng(oRecordset.Fields("ConceptID").Value)
					Case 56
						If CDbl(oRecordset.Fields("TotalAmount").Value) > 0 Then
							sRowContents = Replace(sRowContents, "<CONCEPT_56 />", "S" & Right(("000000" & FormatNumber((CDbl(oRecordset.Fields("TotalAmount").Value) * 100), 0, True, False, False)), Len("000000")))
						Else
							sRowContents = Replace(sRowContents, "<CONCEPT_56 />", "C000000")
						End If
					Case 77
						sRowContents = Replace(sRowContents, "<CONCEPT_77 />", Right(("000000" & FormatNumber((CDbl(oRecordset.Fields("TotalAmount").Value) * 100), 0, True, False, False)), Len("000000")))
				End Select
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			sRowContents = Replace(sRowContents, "<CONCEPT_77 />", "000000")
			sRowContents = Replace(sRowContents, "<CONCEPT_56 />", "C000000")
			lErrorNumber = AppendTextToFile(sFilePath, sRowContents, sErrorDescription)
			oRecordset.Close

			lErrorNumber = ZipFile(sFilePath, Replace(sFilePath, ".txt", ".zip"), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudieron guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				lErrorNumber = DeleteFile(sFilePath, sErrorDescription)
			End If
			oEndDate = Now()
			If (lErrorNumber = 0) And B_USE_SMTP Then
				If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(Replace(sFilePath, ".txt", ".zip"), CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
			End If
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1415 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1417(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the payroll for concepts 56, 76, and 77
'         group by companies, showing only grand totals
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1417"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim lCurrentID
	Dim lEmployeeID
	Dim sCurrentName
	Dim adAmount
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Const FONAC_FACTOR = 0.25

	Call GetConditionFromURL(oRequest, "", lPayrollID, lForPayrollID)
	Response.Write "<TABLE BORDER="""
		If Not bForExport Then
			Response.Write "0"
		Else
			Response.Write "1"
		End If
	Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
		asColumnsTitles = Split("Nómina,Quincena,Concepto 77,Aportación del Instituto,Aportación de la dependencia,Concepto 54,Concepto x54,Aportación", ",", -1, vbBinaryCompare)
		If bForExport Then
			lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
		Else
			If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
				lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			Else
				lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
			End If
		End If

		asCellAlignments = Split(",RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
		adAmount = Split("0,0,0,0,0,0", ",")
		adAmount(0) = 0
		adAmount(1) = 0
		adAmount(2) = 0
		adAmount(3) = 0
		adAmount(4) = 0
		adAmount(5) = 0
		sErrorDescription = "No se pudieron obtener los montos pagados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, Sum(Concept77Amount) As Total77Amount, Sum(Concept54Amount) As Total54Amount From EmployeesFONAC Where (PayrollID=" & lPayrollID & ") " & sCondition & " Group By EmployeeID Order By EmployeeID", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select EmployeeID, Sum(Concept77Amount) As Total77Amount, Sum(Concept54Amount) As Total54Amount From EmployeesFONAC Where (PayrollID=" & lPayrollID & ") " & sCondition & " Group By EmployeeID Order By EmployeeID -->" & vbNewLine
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				lCurrentID = -2
				lEmployeeID = CLng(oRecordset.Fields("EmployeeID").Value)
				sRowContents = "FOVISSSTE"
				sRowContents = sRowContents & TABLE_SEPARATOR & Left(lForPayrollID, Len("0000")) & Right(("00" & GetPayrollNumber(lForPayrollID)), Len("00"))
				sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_77 />"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_INSTITUTO />"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_DEPENDENCIA />"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_54 />"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_x54 />"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_77 />"
				adAmount(0) = 0
				adAmount(1) = 0
				adAmount(2) = 0
				adAmount(3) = 0
				adAmount(4) = 0
				adAmount(5) = 0
				Do While Not oRecordset.EOF
					adAmount(0) = adAmount(0) + CDbl(oRecordset.Fields("Total77Amount").Value)
					adAmount(1) = adAmount(1) + CDbl(oRecordset.Fields("Total54Amount").Value)

'					adAmount(2) = adAmount(2) + (adAmount(4) * CDbl(GetAdminOption(aAdminOptionsComponent, FONAC_02_OPTION)) / CDbl(GetAdminOption(aAdminOptionsComponent, FONAC_01_OPTION)))
'					adAmount(3) = adAmount(3) + (adAmount(4) * CDbl(GetAdminOption(aAdminOptionsComponent, FONAC_03_OPTION)) / CDbl(GetAdminOption(aAdminOptionsComponent, FONAC_01_OPTION)))
					adAmount(2) = adAmount(2) + CDbl(adAmount(4) * (1 + FONAC_FACTOR))
					adAmount(3) = adAmount(3) + CDbl(adAmount(4) * FONAC_FACTOR)
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close

				sRowContents = Replace(sRowContents, "<CONCEPT_77 />", FormatNumber(adAmount(0), 2, True, False, True))
				sRowContents = Replace(sRowContents, "<CONCEPT_INSTITUTO />", FormatNumber(adAmount(2), 2, True, False, True))
				sRowContents = Replace(sRowContents, "<CONCEPT_DEPENDENCIA />", FormatNumber(adAmount(3), 2, True, False, True))
				sRowContents = Replace(sRowContents, "<CONCEPT_54 />", FormatNumber(adAmount(1), 2, True, False, True))
				sRowContents = Replace(sRowContents, "<CONCEPT_x54 />", FormatNumber((adAmount(1) * CDbl(GetAdminOption(aAdminOptionsComponent, FONAC_04_OPTION))), 2, True, False, True))
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			End If
		End If

		sCondition = " And (Payroll_" & lPayrollID & ".ConceptID In (56,77)) And (Payroll_" & lPayrollID & ".EmployeeID In (Select Distinct EmployeeID From Payroll_" & lPayrollID & " Where ConceptID = 77))"
		sErrorDescription = "No se pudieron obtener los montos pagados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Companies.CompanyID, CompanyName, EmployeesHistoryListForPayroll.EmployeeID, ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryListForPayroll, Companies Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryListForPayroll.EmployeeDate) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Companies.CompanyID, CompanyName, EmployeesHistoryListForPayroll.EmployeeID, ConceptID Order By CompanyName, EmployeesHistoryListForPayroll.EmployeeID, ConceptID", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select Companies.CompanyID, CompanyName, EmployeesHistoryListForPayroll.EmployeeID, ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryListForPayroll, Companies Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryListForPayroll.EmployeeDate) And (EmployeesHistoryListForPayroll.CompanyID=Companies.CompanyID) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryListForPayroll.PayrollID=" & lPayrollID & ") And (Companies.StartDate<=" & lForPayrollID & ") And (Companies.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Companies.CompanyID, CompanyName, EmployeesHistoryListForPayroll.EmployeeID, ConceptID Order By CompanyName, EmployeesHistoryListForPayroll.EmployeeID, ConceptID -->" & vbNewLine
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				lCurrentID = -2
				lEmployeeID = CLng(oRecordset.Fields("EmployeeID").Value)
				Do While Not oRecordset.EOF
					If lEmployeeID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
'						adAmount(2) = adAmount(2) + (adAmount(4) * CDbl(GetAdminOption(aAdminOptionsComponent, FONAC_02_OPTION)) / CDbl(GetAdminOption(aAdminOptionsComponent, FONAC_01_OPTION)))
'						adAmount(3) = adAmount(3) + (adAmount(4) * CDbl(GetAdminOption(aAdminOptionsComponent, FONAC_03_OPTION)) / CDbl(GetAdminOption(aAdminOptionsComponent, FONAC_01_OPTION)))						
						adAmount(2) = adAmount(2) + CDbl(adAmount(4) * (1 + FONAC_FACTOR))
						adAmount(3) = adAmount(3) + CDbl(adAmount(4) * FONAC_FACTOR)
						lEmployeeID = CLng(oRecordset.Fields("EmployeeID").Value)
						adAmount(4) = 0
						adAmount(5) = 0
					End If
					If lCurrentID <> CLng(oRecordset.Fields("CompanyID").Value) Then
						If lCurrentID <> -2 Then
							sRowContents = Replace(sRowContents, "<CONCEPT_77 />", FormatNumber(adAmount(0), 2, True, False, True))
							sRowContents = Replace(sRowContents, "<CONCEPT_INSTITUTO />", FormatNumber(adAmount(2), 2, True, False, True))
							sRowContents = Replace(sRowContents, "<CONCEPT_DEPENDENCIA />", FormatNumber(adAmount(3), 2, True, False, True))
							sRowContents = Replace(sRowContents, "<CONCEPT_54 />", FormatNumber(adAmount(1), 2, True, False, True))
							sRowContents = Replace(sRowContents, "<CONCEPT_x54 />", FormatNumber((adAmount(1) * CDbl(GetAdminOption(aAdminOptionsComponent, FONAC_04_OPTION))), 2, True, False, True))
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
						End If
						sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("CompanyName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & Left(lForPayrollID, Len("0000")) & Right(("00" & GetPayrollNumber(lForPayrollID)), Len("00"))
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_77 />"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_INSTITUTO />"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_DEPENDENCIA />"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_54 />"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_x54 />"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_77 />"
						lCurrentID = CLng(oRecordset.Fields("CompanyID").Value)
						adAmount(0) = 0
						adAmount(1) = 0
						adAmount(2) = 0
						adAmount(3) = 0
						adAmount(4) = 0
						adAmount(5) = 0
					End If
					Select Case CLng(oRecordset.Fields("ConceptID").Value)
						Case 56
							adAmount(1) = adAmount(1) + CDbl(oRecordset.Fields("TotalAmount").Value)
							adAmount(5) = adAmount(5) + CDbl(oRecordset.Fields("TotalAmount").Value)
						Case 76
							adAmount(0) = adAmount(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
							adAmount(4) = adAmount(4) + CDbl(oRecordset.Fields("TotalAmount").Value)
						Case 77
							adAmount(0) = adAmount(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
							adAmount(4) = adAmount(4) + CDbl(oRecordset.Fields("TotalAmount").Value)
					End Select
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close

'				adAmount(2) = adAmount(2) + (adAmount(4) * CDbl(GetAdminOption(aAdminOptionsComponent, FONAC_02_OPTION)) / CDbl(GetAdminOption(aAdminOptionsComponent, FONAC_01_OPTION)))
'				adAmount(3) = adAmount(3) + (adAmount(4) * CDbl(GetAdminOption(aAdminOptionsComponent, FONAC_03_OPTION)) / CDbl(GetAdminOption(aAdminOptionsComponent, FONAC_01_OPTION)))
				adAmount(2) = adAmount(2) + CDbl(adAmount(4) * (1 + FONAC_FACTOR))
				adAmount(3) = adAmount(3) + CDbl(adAmount(4) * FONAC_FACTOR)
				sRowContents = Replace(sRowContents, "<CONCEPT_77 />", FormatNumber(adAmount(0), 2, True, False, True))
				sRowContents = Replace(sRowContents, "<CONCEPT_INSTITUTO />", FormatNumber(adAmount(2), 2, True, False, True))
				sRowContents = Replace(sRowContents, "<CONCEPT_DEPENDENCIA />", FormatNumber(adAmount(3), 2, True, False, True))
				sRowContents = Replace(sRowContents, "<CONCEPT_54 />", FormatNumber(adAmount(1), 2, True, False, True))
				sRowContents = Replace(sRowContents, "<CONCEPT_x54 />", FormatNumber((adAmount(1) * CDbl(GetAdminOption(aAdminOptionsComponent, FONAC_04_OPTION))), 2, True, False, True))
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			End If
		End If
	Response.Write "</TABLE><BR />"

	Set oRecordset = Nothing
	BuildReport1417 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1420(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the payroll for the concept 15, 31, C5
'         group by areas
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1420"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim lPayrollNumber
	Dim sCompanyName
	Dim sConceptShortName
	Dim sConceptName
	Dim adTotal
	Dim adGranTotal
	Dim lCurrentArea1ID
	Dim lCurrentArea2ID
	Dim sContents
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, "", lPayrollID, lForPayrollID)
	lPayrollNumber = CInt(GetPayrollNumber(lForPayrollID))
	If Len(oRequest("SubAreaID").Item) > 0 Then
		sCondition = " And (Areas2.AreaID In (" & oRequest("SubAreaID").Item & "))"
	ElseIf Len(oRequest("AreaID").Item) > 0 Then
		sCondition = " And (Areas1.AreaID=" & oRequest("AreaID").Item & ")"
	End If
	If Len(oRequest("CompanyID").Item) > 0 Then
		sCondition = " And (EmployeesHistoryList.CompanyID=" & oRequest("CompanyID").Item & ")"
		Call GetNameFromTable(oADODBConnection, "Companies", oRequest("CompanyID").Item, "", "", sCompanyName, "")
		sCompanyName = Split(sCompanyName, " ", 2)
	End If
	If Len(oRequest("ConceptID").Item) > 0 Then
		sCondition = " And (Payroll_" & lPayrollID & ".ConceptID In (" & oRequest("ConceptID").Item & "))"
	End If
	sErrorDescription = "No se pudieron obtener los montos registrados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Areas1.AreaID As AreaID1, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaID As AreaID2, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, Payroll_" & lPayrollID & ".ConceptID, Count(Payroll_" & lPayrollID & ".EmployeeID) As TotalCount, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2 Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Areas1.AreaID, Areas1.AreaCode, Areas1.AreaName, Areas2.AreaID, Areas2.AreaCode, Areas2.AreaName, Payroll_" & lPayrollID & ".ConceptID Order By Areas1.AreaCode, Areas2.AreaCode, Payroll_" & lPayrollID & ".ConceptID", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select Areas1.AreaID As AreaID1, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaID As AreaID2, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, Payroll_" & lPayrollID & ".ConceptID, Count(Payroll_" & lPayrollID & ".EmployeeID) As TotalCount, Sum(ConceptAmount) As TotalAmount From Payroll_" & lPayrollID & ", EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2 Where (Payroll_" & lPayrollID & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Areas1.AreaID, Areas1.AreaCode, Areas1.AreaName, Areas2.AreaID, Areas2.AreaCode, Areas2.AreaName, Payroll_" & lPayrollID & ".ConceptID Order By Areas1.AreaCode, Areas2.AreaCode, Payroll_" & lPayrollID & ".ConceptID -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1420.htm"), sErrorDescription)
			sContents = Replace(sContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
			sContents = Replace(sContents, "<CURRENT_DATE />", DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, -1))
			sContents = Replace(sContents, "<CURRENT_HOUR />", DisplayTimeFromSerialNumber(Right(GetSerialNumberForDate(""), Len("000000"))))
			sContents = Replace(sContents, "<COMPANY_SHORT_NAME />", sCompanyName(0))
			sContents = Replace(sContents, "<COMPANY_NAME />", sCompanyName(1))
			If StrComp(oRequest("ConceptID").Item, "96", vbBinaryCompare) = 0 Then
				sContents = Replace(sContents, "<CONCEPT_NAME />", "Guardias de PROVAC")
				sContents = Replace(sContents, "<CONCEPTS_NAMES />", "<TD COLSPAN=""2"" ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2""><B>Guardias de Provac</B></FONT></TD>")
			Else
				sContents = Replace(sContents, "<CONCEPT_NAME />", "Guardias y Suplencias")
				sContents = Replace(sContents, "<CONCEPTS_NAMES />", "<TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2""><B>Remuneración por Guardias</B></FONT></TD><TD ALIGN=""RIGHT""><FONT FACE=""Arial"" SIZE=""2""><B>Remuneración por Suplencias</B></FONT></TD>")
			End If
			sContents = Replace(sContents, "<PAYROLL_NUMBER />", lPayrollNumber)
			sContents = Replace(sContents, "<PAYROLL_YEAR />", Left(lForPayrollID, Len("0000")))
			sContents = Replace(sContents, "<PAYROLL_DATE />", DisplayDateFromSerialNumber(lForPayrollID, -1, -1, -1))
			Response.Write sContents
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asCellAlignments = Split(",,,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
				adTotal = Split("0,0,0", ",")
				adTotal(0) = 0
				adTotal(1) = 0
				adTotal(2) = 0
				adGranTotal = Split("0,0,0", ",")
				adGranTotal(0) = 0
				adGranTotal(1) = 0
				adGranTotal(2) = 0
				lCurrentArea1ID = -2
				lCurrentArea2ID = -2

				sRowContents = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode1").Value)) & "</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""4"" /><B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaName1").Value)) & "</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				lCurrentArea1ID = CLng(oRecordset.Fields("AreaID1").Value)
				Do While Not oRecordset.EOF
					If lCurrentArea2ID <> CLng(oRecordset.Fields("AreaID2").Value) Then
						If lCurrentArea2ID <> -2 Then
							sRowContents = Replace(sRowContents, "<CONCEPT_18 />", "0.00")
							sRowContents = Replace(sRowContents, "<CONCEPT_34 />", "0.00")
							sRowContents = Replace(sRowContents, "<CONCEPT_96 />", "0.00")
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
						End If

						sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("AreaCode2").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""2"" />" & CleanStringForHTML(CStr(oRecordset.Fields("AreaName2").Value))
						If StrComp(oRequest("ConceptID").Item, "96", vbBinaryCompare) = 0 Then
							sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""2"" /><CONCEPT_96 />"
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_18 />"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_34 />"
						End If
						lCurrentArea2ID = CLng(oRecordset.Fields("AreaID2").Value)
					End If
					If lCurrentArea1ID <> CLng(oRecordset.Fields("AreaID1").Value) Then
						sRowContents = "<SPAN COLS=""2"" /><B>TOTAL POR FECHA DE PAGO</B>" & TABLE_SEPARATOR & "QUINCENA " & lPayrollNumber & " DE " & Left(lForPayrollID, Len("0000"))
						If StrComp(oRequest("ConceptID").Item, "96", vbBinaryCompare) = 0 Then
							sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""2"" />" & FormatNumber(adTotal(2), 2, True, False, True)
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotal(0), 2, True, False, True)
							sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotal(1), 2, True, False, True)
						End If
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						adTotal(0) = 0
						adTotal(1) = 0
						adTotal(2) = 0

						sRowContents = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode1").Value)) & "</B>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""4"" /><B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaName1").Value)) & "</B>"
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						lCurrentArea1ID = CLng(oRecordset.Fields("AreaID1").Value)

						sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("AreaCode2").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""2"" />" & CleanStringForHTML(CStr(oRecordset.Fields("AreaName2").Value))
						If StrComp(oRequest("ConceptID").Item, "96", vbBinaryCompare) = 0 Then
							sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""2"" /><CONCEPT_96 />"
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_18 />"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<CONCEPT_34 />"
						End If
					End If
					Select Case CLng(oRecordset.Fields("ConceptID").Value)
						Case 18
							sRowContents = Replace(sRowContents, "<CONCEPT_18 />", FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True))
							adTotal(0) = adTotal(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
							adGranTotal(0) = adGranTotal(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
						Case 34
							sRowContents = Replace(sRowContents, "<CONCEPT_34 />", FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True))
							adTotal(1) = adTotal(1) + CDbl(oRecordset.Fields("TotalAmount").Value)
							adGranTotal(1) = adGranTotal(1) + CDbl(oRecordset.Fields("TotalAmount").Value)
						Case 96
							sRowContents = Replace(sRowContents, "<CONCEPT_96 />", FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True))
							adTotal(2) = adTotal(2) + CDbl(oRecordset.Fields("TotalAmount").Value)
							adGranTotal(2) = adGranTotal(2) + CDbl(oRecordset.Fields("TotalAmount").Value)
					End Select

					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close

				sRowContents = Replace(sRowContents, "<CONCEPT_18 />", "0.00")
				sRowContents = Replace(sRowContents, "<CONCEPT_34 />", "0.00")
				sRowContents = Replace(sRowContents, "<CONCEPT_96 />", "0.00")
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If

				sRowContents = "<SPAN COLS=""2"" /><B>TOTAL POR FECHA DE PAGO</B>" & TABLE_SEPARATOR & "QUINCENA " & lPayrollNumber & " DE " & Left(lForPayrollID, Len("0000"))
				If StrComp(oRequest("ConceptID").Item, "96", vbBinaryCompare) = 0 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""2"" />" & FormatNumber(adTotal(2), 2, True, False, True)
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotal(0), 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotal(1), 2, True, False, True)
				End If
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If

				sRowContents = "<SPAN COLS=""3"" /><B>TOTAL GENERAL</B>"
				If StrComp(oRequest("ConceptID").Item, "96", vbBinaryCompare) = 0 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""2"" />" & FormatNumber(adGranTotal(2), 2, True, False, True)
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adGranTotal(0), 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adGranTotal(1), 2, True, False, True)
				End If
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			Response.Write "</TABLE>"
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1420 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1421(oRequest, oADODBConnection, bExternal, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the payroll for the concept 15, 31, C5
'         group by areas, employees, and dates
'Inputs:  oRequest, oADODBConnection, bExternal, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1421"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim lPayrollNumber
	Dim sConceptName
	Dim adTotal
	Dim lCurrentEmployeeID
	Dim lCurrentArea1ID
	Dim lCurrentArea2ID
	Dim sContents
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, "", lPayrollID, lForPayrollID)
	lPayrollNumber = "" & GetPayrollNumber(lForPayrollID) & Right(Left(lForPayrollID, Len("0000")), Len("00"))
	If Len(oRequest("SubAreaID").Item) > 0 Then
		sCondition = " And (Areas2.AreaID In (" & oRequest("SubAreaID").Item & "))"
	ElseIf Len(oRequest("AreaID").Item) > 0 Then
		sCondition = " And (Areas1.AreaID=" & oRequest("AreaID").Item & ")"
	End If
	If Len(oRequest("ConceptID").Item) > 0 Then
		Select Case CLng(oRequest("ConceptID").Item)
			Case 18
				sCondition = sCondition & " And (EmployeesSpecialJourneys.SpecialJourneyID=423)"
			Case 34
				sCondition = sCondition & " And (EmployeesSpecialJourneys.SpecialJourneyID=424)"
			Case 96
				sCondition = sCondition & " And (EmployeesSpecialJourneys.SpecialJourneyID=425)"
		End Select
	End If
	sErrorDescription = "No se pudieron obtener los montos registrados."
	If bExternal Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Areas1.AreaID As AreaID1, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaID As AreaID2, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, EmployeesSpecialJourneys.RFC As EmpID, EmployeesSpecialJourneys.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, PositionShortName, PositionName, LevelShortName, AppliedDate, EmployeesSpecialJourneys.StartDate, EmployeesSpecialJourneys.EndDate, ReasonShortName, WorkedHours, ConceptAmount From EmployeesSpecialJourneys, Areas As Areas1, Areas As Areas2, Positions, Levels, SpecialJourneysReasons Where (EmployeesSpecialJourneys.PositionID=Positions.PositionID) And (EmployeesSpecialJourneys.LevelID=Levels.LevelID) And (EmployeesSpecialJourneys.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesSpecialJourneys.ReasonID=SpecialJourneysReasons.ReasonID) And (EmployeesSpecialJourneys.AppliedDate=" & lPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (EmployeesSpecialJourneys.EmployeeID>=800000) " & sCondition & " Order By Areas1.AreaCode, Areas2.AreaCode, EmployeesSpecialJourneys.RFC, EmployeesSpecialJourneys.StartDate", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select Areas1.AreaID As AreaID1, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaID As AreaID2, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, EmployeesSpecialJourneys.RFC As EmpID, EmployeesSpecialJourneys.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, PositionShortName, PositionName, LevelShortName, AppliedDate, EmployeesSpecialJourneys.StartDate, EmployeesSpecialJourneys.EndDate, ReasonShortName, WorkedHours, ConceptAmount From EmployeesSpecialJourneys, Areas As Areas1, Areas As Areas2, Positions, Levels, SpecialJourneysReasons Where (EmployeesSpecialJourneys.PositionID=Positions.PositionID) And (EmployeesSpecialJourneys.LevelID=Levels.LevelID) And (EmployeesSpecialJourneys.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesSpecialJourneys.ReasonID=SpecialJourneysReasons.ReasonID) And (EmployeesSpecialJourneys.AppliedDate=" & lPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (EmployeesSpecialJourneys.EmployeeID>=800000) " & sCondition & " Order By Areas1.AreaCode, Areas2.AreaCode, EmployeesSpecialJourneys.RFC, EmployeesSpecialJourneys.StartDate -->" & vbNewLine
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Areas1.AreaID As AreaID1, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaID As AreaID2, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, EmployeesHistoryList.EmployeeID As EmpID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, PositionShortName, PositionName, LevelShortName, AppliedDate, EmployeesSpecialJourneys.StartDate, EmployeesSpecialJourneys.EndDate, ReasonShortName, WorkedHours, ConceptAmount From EmployeesSpecialJourneys, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Positions, Levels, SpecialJourneysReasons Where (EmployeesSpecialJourneys.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesSpecialJourneys.ReasonID=SpecialJourneysReasons.ReasonID) And (EmployeesSpecialJourneys.AppliedDate=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") " & sCondition & " Order By Areas1.AreaCode, Areas2.AreaCode, EmployeesHistoryList.EmployeeNumber, EmployeesSpecialJourneys.StartDate", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select Areas1.AreaID As AreaID1, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaID As AreaID2, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, EmployeesHistoryList.EmployeeID As EmpID, EmployeesHistoryList.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, PositionShortName, PositionName, LevelShortName, AppliedDate, EmployeesSpecialJourneys.StartDate, EmployeesSpecialJourneys.EndDate, ReasonShortName, WorkedHours, ConceptAmount From EmployeesSpecialJourneys, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Positions, Levels, SpecialJourneysReasons Where (EmployeesSpecialJourneys.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesSpecialJourneys.ReasonID=SpecialJourneysReasons.ReasonID) And (EmployeesSpecialJourneys.AppliedDate=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") " & sCondition & " Order By Areas1.AreaCode, Areas2.AreaCode, EmployeesHistoryList.EmployeeNumber, EmployeesSpecialJourneys.StartDate -->" & vbNewLine
	End If
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1421.htm"), sErrorDescription)
			sContents = Replace(sContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
			Select Case CLng(oRequest("ConceptID").Item)
				Case 18
					sContents = Replace(sContents, "<CONCEPT_NAME />", "15 (Guardias)")
				Case 34
					sContents = Replace(sContents, "<CONCEPT_NAME />", "31 (Suplencias)")
				Case 96
					sContents = Replace(sContents, "<CONCEPT_NAME />", "C5 (Guardias de PROVAC)")
			End Select
			Response.Write sContents
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If CLng(oRequest("ConceptID").Item) = 34 Then
					asColumnsTitles = Split("No. del empleado,Ordinal,Fecha de pago,Fecha de inicio,Fecha de fin,Clave de movimiento,Días suplidos,Quincena", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("No. del empleado,Ordinal,Fecha de pago,Fecha de inicio,Fecha de fin,Clave de movimiento,Horas de guardia,Quincena", ",", -1, vbBinaryCompare)
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
				asCellAlignments = Split(",,,,,,RIGHT,", ",", -1, vbBinaryCompare)
				adTotal = Split(",,", ",")
				adTotal(0) = Split("0,0,0", ",")
				adTotal(0)(0) = 0
				adTotal(0)(1) = 0
				adTotal(0)(2) = 0
				adTotal(1) = Split("0,0,0", ",")
				adTotal(1)(0) = 0
				adTotal(1)(1) = 0
				adTotal(1)(2) = 0
				adTotal(2) = Split("0,0,0", ",")
				adTotal(2)(0) = 0
				adTotal(2)(1) = 0
				adTotal(2)(2) = 0
				lCurrentEmployeeID = "-2"
				lCurrentArea1ID = -2
				lCurrentArea2ID = -2

				sRowContents = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode1").Value)) & "</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""7"" /><B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaName1").Value)) & "</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				lCurrentArea1ID = CLng(oRecordset.Fields("AreaID1").Value)
				Do While Not oRecordset.EOF
					If lCurrentArea1ID <> CLng(oRecordset.Fields("AreaID1").Value) Then
						If StrComp(lCurrentEmployeeID, "-2", vbBinaryCompare) <> 0 Then
							sRowContents = TABLE_SEPARATOR & "<SPAN COLS=""2"" />Fecha de pago: " & DisplayNumericDateFromSerialNumber(lForPayrollID)
							sRowContents = sRowContents & TABLE_SEPARATOR & "Importe total: " & TABLE_SEPARATOR & FormatNumber(adTotal(0)(2), 2, True, False, True)
							If CLng(oRequest("ConceptID").Item) = 34 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "Total horas: "
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & "Total días: "
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & adTotal(0)(1) & TABLE_SEPARATOR
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
							lErrorNumber = DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
							adTotal(1)(0) = adTotal(1)(0) + adTotal(0)(0)
							adTotal(1)(1) = adTotal(1)(1) + adTotal(0)(1)
							adTotal(1)(2) = adTotal(1)(2) + adTotal(0)(2)
							adTotal(0)(0) = 0
							adTotal(0)(1) = 0
							adTotal(0)(2) = 0
						End If

						sRowContents = "<B>TOTALES</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotal(1)(0), 0, True, False, True) & "</B>" & TABLE_SEPARATOR & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotal(1)(2), 2, True, False, True) & "</B>" & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						lErrorNumber = DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
						adTotal(2)(0) = adTotal(2)(0) + adTotal(1)(0)
						adTotal(2)(1) = adTotal(2)(1) + adTotal(1)(1)
						adTotal(2)(2) = adTotal(2)(2) + adTotal(1)(2)
						adTotal(1)(0) = 0
						adTotal(1)(1) = 0
						adTotal(1)(2) = 0

						sRowContents = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode1").Value)) & "</B>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""7"" /><B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaName1").Value)) & "</B>"
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						lCurrentArea1ID = CLng(oRecordset.Fields("AreaID1").Value)
						lCurrentArea2ID = -2
						lCurrentEmployeeID = "-2"
					End If
					If lCurrentArea2ID <> CLng(oRecordset.Fields("AreaID2").Value) Then
						If StrComp(lCurrentEmployeeID, "-2", vbBinaryCompare) <> 0 Then
							sRowContents = TABLE_SEPARATOR & "<SPAN COLS=""2"" />Fecha de pago: " & DisplayNumericDateFromSerialNumber(lForPayrollID)
							sRowContents = sRowContents & TABLE_SEPARATOR & "Importe total: " & TABLE_SEPARATOR & FormatNumber(adTotal(0)(2), 2, True, False, True)
							If CLng(oRequest("ConceptID").Item) = 34 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "Total horas: "
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & "Total días: "
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & adTotal(0)(1) & TABLE_SEPARATOR
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
							lErrorNumber = DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
							adTotal(1)(0) = adTotal(1)(0) + adTotal(0)(0)
							adTotal(1)(1) = adTotal(1)(1) + adTotal(0)(1)
							adTotal(1)(2) = adTotal(1)(2) + adTotal(0)(2)
							adTotal(0)(0) = 0
							adTotal(0)(1) = 0
							adTotal(0)(2) = 0
						End If

						sRowContents = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode2").Value)) & "</B>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""7"" /><B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaName2").Value)) & "</B>"
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						lCurrentArea2ID = CLng(oRecordset.Fields("AreaID2").Value)
						lCurrentEmployeeID = "-2"
					End If
					If StrComp(lCurrentEmployeeID, CStr(oRecordset.Fields("EmpID").Value), vbBinaryCompare) <> 0 Then
						If StrComp(lCurrentEmployeeID, "-2", vbBinaryCompare) <> 0 Then
							sRowContents = TABLE_SEPARATOR & "<SPAN COLS=""2"" />Fecha de pago: " & DisplayNumericDateFromSerialNumber(lForPayrollID)
							sRowContents = sRowContents & TABLE_SEPARATOR & "Importe total: " & TABLE_SEPARATOR & FormatNumber(adTotal(0)(2), 2, True, False, True)
							If CLng(oRequest("ConceptID").Item) = 34 Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "Total horas: "
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & "Total días: "
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR & adTotal(0)(1) & TABLE_SEPARATOR
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
							lErrorNumber = DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
							adTotal(1)(0) = adTotal(1)(0) + adTotal(0)(0)
							adTotal(1)(1) = adTotal(1)(1) + adTotal(0)(1)
							adTotal(1)(2) = adTotal(1)(2) + adTotal(0)(2)
							adTotal(0)(0) = 0
							adTotal(0)(1) = 0
							adTotal(0)(2) = 0
						End If
						sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
						If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
							sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""2"" />" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & ", " & CStr(oRecordset.Fields("EmployeeName").Value))
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""2"" />" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value))
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""2"" />" & CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(Left(("000" & CStr(oRecordset.Fields("LevelShortName").Value)), Len("00")))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(Right(CStr(oRecordset.Fields("LevelShortName").Value), Len("0")))
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						lCurrentEmployeeID = CStr(oRecordset.Fields("EmpID").Value)
					End If

					adTotal(0)(0) = adTotal(0)(0) + 1
					sRowContents = "&nbsp;" & TABLE_SEPARATOR & adTotal(0)(0)
					sRowContents = sRowContents & TABLE_SEPARATOR & "Fecha de pago: " & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("AppliedDate").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & "Fecha de inicio: " & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & "Fecha de fin: " & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ReasonShortName").Value))
					If CLng(oRequest("ConceptID").Item) = 34 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & (Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)), GetDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)))) + 1)
						adTotal(0)(1) = adTotal(0)(1) + Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)), GetDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)))) + 1
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & CDbl(oRecordset.Fields("WorkedHours").Value) * (Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)), GetDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)))) + 1)
						adTotal(0)(1) = adTotal(0)(1) + CDbl(oRecordset.Fields("WorkedHours").Value) * (Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)), GetDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)))) + 1)
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & "QNA" & lPayrollNumber
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					adTotal(0)(2) = adTotal(0)(2) + CDbl(oRecordset.Fields("ConceptAmount").Value)

					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close

				sRowContents = TABLE_SEPARATOR & "<SPAN COLS=""2"" />Fecha de pago: " & DisplayNumericDateFromSerialNumber(lForPayrollID)
				sRowContents = sRowContents & TABLE_SEPARATOR & "Importe total: " & TABLE_SEPARATOR & FormatNumber(adTotal(0)(2), 2, True, False, True)
				If CLng(oRequest("ConceptID").Item) = 34 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & "Total horas: "
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & "Total días: "
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & adTotal(0)(1) & TABLE_SEPARATOR
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				adTotal(1)(0) = adTotal(1)(0) + adTotal(0)(0)
				adTotal(1)(1) = adTotal(1)(1) + adTotal(0)(1)
				adTotal(1)(2) = adTotal(1)(2) + adTotal(0)(2)
				adTotal(0)(0) = 0
				adTotal(0)(1) = 0
				adTotal(0)(2) = 0

				sRowContents = "<B>TOTALES</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotal(1)(0), 0, True, False, True) & "</B>" & TABLE_SEPARATOR & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotal(1)(2), 2, True, False, True) & "</B>" & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				lErrorNumber = DisplayLine(asColumnsTitles, "", bForExport, sErrorDescription)
				adTotal(2)(0) = adTotal(2)(0) + adTotal(1)(0)
				adTotal(2)(1) = adTotal(2)(1) + adTotal(1)(1)
				adTotal(2)(2) = adTotal(2)(2) + adTotal(1)(2)
				adTotal(1)(0) = 0
				adTotal(1)(1) = 0
				adTotal(1)(2) = 0

				sRowContents = "<B>REGISTROS LEIDOS</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotal(2)(0), 0, True, False, True) & "</B>" & TABLE_SEPARATOR & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotal(2)(2), 2, True, False, True) & "</B>" & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			Response.Write "</TABLE>"
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1421 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1422(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the employee's records for concepts
'         15, 31, C5 sort by RecordID
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1422"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim iCounter
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, "", lPayrollID, lForPayrollID)
	If Len(oRequest("SubAreaID").Item) > 0 Then
		sCondition = " And (Areas2.AreaID In (" & oRequest("SubAreaID").Item & "))"
	ElseIf Len(oRequest("AreaID").Item) > 0 Then
		sCondition = " And (Areas1.AreaID=" & oRequest("AreaID").Item & ")"
	End If
	If Len(oRequest("ConceptID").Item) > 0 Then
		Select Case CLng(oRequest("ConceptID").Item)
			Case 18
				sCondition = sCondition & " And (EmployeesSpecialJourneys.SpecialJourneyID=423)"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>REPORTE EN ORDEN DE CAPTURA DE MOVIMIENTOS DE GUARDIAS</B></FONT><BR /><BR />"
			Case 34
				sCondition = sCondition & " And (EmployeesSpecialJourneys.SpecialJourneyID=424)"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>REPORTE EN ORDEN DE CAPTURA DE MOVIMIENTOS DE SUPLENCIAS</B></FONT><BR /><BR />"
			Case 96
				sCondition = sCondition & " And (EmployeesSpecialJourneys.SpecialJourneyID=425)"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>REPORTE EN ORDEN DE CAPTURA DE MOVIMIENTOS DE REZAGO QUIRÚRGICO</B></FONT><BR /><BR />"
		End Select
	End If
	sErrorDescription = "No se pudieron obtener los montos registrados."
	If True Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeNumber, RFC, EmployeeName, EmployeeLastName, EmployeeLastName2, EmployeesSpecialJourneys.WorkingHours, Areas2.AreaShortName, ServiceShortName, PositionShortName, LevelShortName, OriginalEmployeeID, RiskLevelID, DocumentNumber, EmployeesSpecialJourneys.StartDate, EmployeesSpecialJourneys.EndDate, WorkedHours, JourneyShortName, MovementShortName, ReasonShortName, LevelShortName As OriginalLevelShortName From EmployeesSpecialJourneys, Areas As Areas1, Areas As Areas2, Services, Positions, Levels, SpecialJourneys, SpecialJourneysMovements, SpecialJourneysReasons Where (EmployeesSpecialJourneys.ServiceID=Services.ServiceID) And (EmployeesSpecialJourneys.PositionID=Positions.PositionID) And (EmployeesSpecialJourneys.LevelID=Levels.LevelID) And (EmployeesSpecialJourneys.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesSpecialJourneys.JourneyID=SpecialJourneys.JourneyID) And (EmployeesSpecialJourneys.MovementID=SpecialJourneysMovements.MovementID) And (EmployeesSpecialJourneys.ReasonID=SpecialJourneysReasons.ReasonID) And (EmployeesSpecialJourneys.AppliedDate=" & lPayrollID & ") And (Services.StartDate<=" & lForPayrollID & ") And (Services.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (EmployeeID>=800000) " & sCondition & " Order By RecordID", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select EmployeeNumber, RFC, EmployeeName, EmployeeLastName, EmployeeLastName2, EmployeesSpecialJourneys.WorkingHours, Areas2.AreaShortName, ServiceShortName, PositionShortName, LevelShortName, OriginalEmployeeID, RiskLevelID, DocumentNumber, EmployeesSpecialJourneys.StartDate, EmployeesSpecialJourneys.EndDate, WorkedHours, JourneyShortName, MovementShortName, ReasonShortName, LevelShortName As OriginalLevelShortName From EmployeesSpecialJourneys, Areas As Areas1, Areas As Areas2, Services, Positions, Levels, SpecialJourneys, SpecialJourneysMovements, SpecialJourneysReasons Where (EmployeesSpecialJourneys.ServiceID=Services.ServiceID) And (EmployeesSpecialJourneys.PositionID=Positions.PositionID) And (EmployeesSpecialJourneys.LevelID=Levels.LevelID) And (EmployeesSpecialJourneys.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesSpecialJourneys.JourneyID=SpecialJourneys.JourneyID) And (EmployeesSpecialJourneys.MovementID=SpecialJourneysMovements.MovementID) And (EmployeesSpecialJourneys.ReasonID=SpecialJourneysReasons.ReasonID) And (EmployeesSpecialJourneys.AppliedDate=" & lPayrollID & ") And (Services.StartDate<=" & lForPayrollID & ") And (Services.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (EmployeeID>=800000) " & sCondition & " Order By RecordID -->" & vbNewLine
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeNumber, RFC, EmployeeName, EmployeeLastName, EmployeeLastName2, EmployeesHistoryList.WorkingHours, Areas2.AreaShortName, ServiceShortName, PositionShortName, Levels.LevelShortName, EmployeesSpecialJourneys.OriginalEmployeeID, EmployeesSpecialJourneys.RiskLevelID, EmployeesSpecialJourneys.DocumentNumber, EmployeesSpecialJourneys.StartDate, EmployeesSpecialJourneys.EndDate, EmployeesSpecialJourneys.WorkedHours, JourneyShortName, MovementShortName, ReasonShortName, OriginalLevels.LevelShortName As OriginalLevelShortName From EmployeesSpecialJourneys, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Services, Positions, Levels, Levels As OriginalLevels, SpecialJourneys, SpecialJourneysMovements, SpecialJourneysReasons Where (EmployeesSpecialJourneys.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.ServiceID=Services.ServiceID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesSpecialJourneys.JourneyID=SpecialJourneys.JourneyID) And (EmployeesSpecialJourneys.MovementID=SpecialJourneysMovements.MovementID) And (EmployeesSpecialJourneys.ReasonID=SpecialJourneysReasons.ReasonID) And (EmployeesSpecialJourneys.LevelID=OriginalLevels.LevelID) And (EmployeesSpecialJourneys.AppliedDate=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Services.StartDate<=" & lForPayrollID & ") And (Services.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (OriginalLevels.StartDate<=" & lForPayrollID & ") And (OriginalLevels.EndDate>=" & lForPayrollID & ") " & sCondition & " Order By RecordID", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryList.EmployeeNumber, RFC, EmployeeName, EmployeeLastName, EmployeeLastName2, EmployeesHistoryList.WorkingHours, Areas2.AreaShortName, ServiceShortName, PositionShortName, Levels.LevelShortName, EmployeesSpecialJourneys.OriginalEmployeeID, EmployeesSpecialJourneys.RiskLevelID, EmployeesSpecialJourneys.DocumentNumber, EmployeesSpecialJourneys.StartDate, EmployeesSpecialJourneys.EndDate, EmployeesSpecialJourneys.WorkedHours, JourneyShortName, MovementShortName, ReasonShortName, OriginalLevels.LevelShortName As OriginalLevelShortName From EmployeesSpecialJourneys, EmployeesChangesLKP, EmployeesHistoryList, Areas As Areas1, Areas As Areas2, Services, Positions, Levels, Levels As OriginalLevels, SpecialJourneys, SpecialJourneysMovements, SpecialJourneysReasons Where (EmployeesSpecialJourneys.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.ServiceID=Services.ServiceID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesSpecialJourneys.JourneyID=SpecialJourneys.JourneyID) And (EmployeesSpecialJourneys.MovementID=SpecialJourneysMovements.MovementID) And (EmployeesSpecialJourneys.ReasonID=SpecialJourneysReasons.ReasonID) And (EmployeesSpecialJourneys.LevelID=OriginalLevels.LevelID) And (EmployeesSpecialJourneys.AppliedDate=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollDate=" & lPayrollID & ") And (Services.StartDate<=" & lForPayrollID & ") And (Services.EndDate>=" & lForPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (OriginalLevels.StartDate<=" & lForPayrollID & ") And (OriginalLevels.EndDate>=" & lForPayrollID & ") " & sCondition & " Order By RecordID -->" & vbNewLine
	End If
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("Registro,No. del empleado,RFC,Turno,Nombre,Adscripción,Servicio,Puesto,RP,Nivel,Jornada,Movimiento,Fecha inicio,Fecha final,Horas/días,Empleado suplido,Motivo,Folio,Nivel suplido", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				asCellAlignments = Split(",,,,,,,,,,,,,,RIGHT,,,,,", ",", -1, vbBinaryCompare)
				iCounter = 1
				Do While Not oRecordset.EOF
					sRowContents = iCounter
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("JourneyShortName").Value))
					If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & ", " & CStr(oRecordset.Fields("EmployeeName").Value))
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value))
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaShortName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ServiceShortName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
					If CLng(oRecordset.Fields("RiskLevelID").Value) > 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & CStr(oRecordset.Fields("RiskLevelID").Value)
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("WorkingHours").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("MovementShortName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("StartDate").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EndDate").Value))
					If CLng(oRequest("ConceptID").Item) = 34 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & (Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)), GetDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)))) + 1)
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(Right(("000000" & CStr(oRecordset.Fields("OriginalEmployeeID").Value)), Len("000000")))
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & CDbl(oRecordset.Fields("WorkedHours").Value) * (Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)), GetDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)))) + 1)
						sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ReasonShortName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("DocumentNumber").Value))
					If CLng(oRequest("ConceptID").Item) = 34 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("OriginalLevelShortName").Value))
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & "QNA" & lPayrollNumber
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					iCounter = iCounter + 1
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
			Response.Write "</TABLE>"
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1422 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1424(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the payroll for the concept RQ group by areas
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1424"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim lCounter
	Dim lGranCounter
	Dim dTotal
	Dim dGranTotal
	Dim lCurrentID
	Dim sContents
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, "", lPayrollID, lForPayrollID)
	If Len(oRequest("SubAreaID").Item) > 0 Then
		sCondition = " And (Areas2.AreaID In (" & oRequest("SubAreaID").Item & "))"
	ElseIf Len(oRequest("AreaID").Item) > 0 Then
		sCondition = " And (Areas1.AreaID=" & oRequest("AreaID").Item & ")"
	End If
	If Len(oRequest("External").Item) = 0 Then
		sCondition = sCondition & " And (EmployeesSpecialJourneys.EmployeeID<800000)"
		sContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1424a.htm"), sErrorDescription)
	Else
		sCondition = sCondition & " And (EmployeesSpecialJourneys.EmployeeID>=800000)"
		sContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1424b.htm"), sErrorDescription)
	End If
	sCondition = sCondition & " And (EmployeesSpecialJourneys.SpecialJourneyID=425)"
	sErrorDescription = "No se pudieron obtener los montos registrados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Areas1.AreaID As AreaID1, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaID As AreaID2, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, Count(EmployeeID) As TotalCount, Sum(ConceptAmount) As TotalAmount From EmployeesSpecialJourneys, Areas As Areas1, Areas As Areas2 Where (EmployeesSpecialJourneys.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesSpecialJourneys.AppliedDate=" & lPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Areas1.AreaID, Areas1.AreaCode, Areas1.AreaName, Areas2.AreaID, Areas2.AreaCode, Areas2.AreaName Order By Areas1.AreaCode, Areas2.AreaCode", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select Areas1.AreaID As AreaID1, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaID As AreaID2, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, Count(EmployeeID) As TotalCount, Sum(ConceptAmount) As TotalAmount From EmployeesSpecialJourneys, Areas As Areas1, Areas As Areas2 Where (EmployeesSpecialJourneys.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesSpecialJourneys.AppliedDate=" & lPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By Areas1.AreaID, Areas1.AreaCode, Areas1.AreaName, Areas2.AreaID, Areas2.AreaCode, Areas2.AreaName Order By Areas1.AreaCode, Areas2.AreaCode -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sContents = Replace(sContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
			sContents = Replace(sContents, "<PAYROLL_DATE />", DisplayDateFromSerialNumber(lForPayrollID, -1, -1, -1))
			Response.Write sContents
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("No. del centro,Centro de trabajo,Registros,Importe", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				asCellAlignments = Split(",,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
				lCounter = 0
				lGranCounter = 0
				dTotal = 0
				dGranTotal = 0
				lCurrentID = -2
				sRowContents = "<SPAN COLS=""4"" /><B>UNIDAD ADMINISTRATIVA: " & CleanStringForHTML(CStr(oRecordset.Fields("AreaName1").Value)) & "</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				Do While Not oRecordset.EOF
					If lCurrentID <> CLng(oRecordset.Fields("AreaID1").Value) Then
						If lCurrentID <> -2 Then
							sRowContents = "&nbsp;" & TABLE_SEPARATOR & "<B>Total por Estado:</B>" & TABLE_SEPARATOR & FormatNumber(lCounter, 0, True, False, True) & "</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(dTotal, 2, True, False, True) & "</B>"
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
							lGranCounter = lGranCounter + lCounter
							dGranTotal = dGranTotal + dTotal
							lCounter = 0
							dTotal = 0
							asRowContents = Split("<SPAN COLS=""4"" />&nbsp;", TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
						End If

						sRowContents = "<SPAN COLS=""4"" /><B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode1").Value) & ". " & CStr(oRecordset.Fields("AreaName1").Value)) & "</B>"
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						lCurrentID = CLng(oRecordset.Fields("AreaID1").Value)
					End If
					sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("AreaCode2").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaName2").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CLng(oRecordset.Fields("TotalCount").Value), 0, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True)
					lCounter = lCounter + CLng(oRecordset.Fields("TotalCount").Value)
					dTotal = dTotal + CDbl(oRecordset.Fields("TotalAmount").Value)

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If

					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close

				sRowContents = "&nbsp;" & TABLE_SEPARATOR & "<B>Total por Estado:</B>" & TABLE_SEPARATOR & FormatNumber(lCounter, 0, True, False, True) & "</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(dTotal, 2, True, False, True) & "</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				lGranCounter = lGranCounter + lCounter
				dGranTotal = dGranTotal + dTotal
				asRowContents = Split("<SPAN COLS=""4"" />&nbsp;", TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If

				sRowContents = "&nbsp;" & TABLE_SEPARATOR & "<B>Total por Estado:</B>" & TABLE_SEPARATOR & FormatNumber(lGranCounter, 0, True, False, True) & "</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(dGranTotal, 2, True, False, True) & "</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			Response.Write "</TABLE>"
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1424 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1425(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the payroll for the concept RQ group by areas
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1425"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim dTotal
	Dim dGranTotal
	Dim lCounter
	Dim lGranCounter
	Dim lCurrentID
	Dim lCurrentAreaID
	Dim sContents
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, "", lPayrollID, lForPayrollID)
	If Len(oRequest("SubAreaID").Item) > 0 Then
		sCondition = " And (Areas2.AreaID In (" & oRequest("SubAreaID").Item & "))"
	ElseIf Len(oRequest("AreaID").Item) > 0 Then
		sCondition = " And (Areas1.AreaID=" & oRequest("AreaID").Item & ")"
	End If
	If Len(oRequest("External").Item) = 0 Then
		sCondition = sCondition & " And (EmployeesSpecialJourneys.EmployeeID<800000)"
		sContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1425a.htm"), sErrorDescription)
	Else
		sCondition = sCondition & " And (EmployeesSpecialJourneys.EmployeeID>=800000)"
		sContents = GetFileContents(Server.MapPath("Templates\HeaderForReport_1425b.htm"), sErrorDescription)
	End If
	sCondition = sCondition & " And (EmployeesSpecialJourneys.SpecialJourneyID=425)"
	sErrorDescription = "No se pudieron obtener los montos registrados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, Areas1.AreaID As AreaID1, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaID As AreaID2, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, Sum(ConceptAmount) As TotalAmount From EmployeesSpecialJourneys, Areas As Areas1, Areas As Areas2 Where (EmployeesSpecialJourneys.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesSpecialJourneys.AppliedDate=" & lPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, Areas1.AreaID, Areas1.AreaCode, Areas1.AreaName, Areas2.AreaID, Areas2.AreaCode, Areas2.AreaName Order By Areas1.AreaCode, Areas2.AreaCode, EmployeeNumber", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, Areas1.AreaID As AreaID1, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaID As AreaID2, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, Sum(ConceptAmount) As TotalAmount From EmployeesSpecialJourneys, Areas As Areas1, Areas As Areas2 Where (EmployeesSpecialJourneys.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesSpecialJourneys.AppliedDate=" & lPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") " & sCondition & " Group By EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, Areas1.AreaID, Areas1.AreaCode, Areas1.AreaName, Areas2.AreaID, Areas2.AreaCode, Areas2.AreaName Order By Areas1.AreaCode, Areas2.AreaCode, EmployeeNumber -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sContents = Replace(sContents, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
			sContents = Replace(sContents, "<PAYROLL_DATE />", DisplayDateFromSerialNumber(lForPayrollID, -1, -1, -1))
			Response.Write sContents
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("No. del empleado,Nombre,Fecha de pago,Importe", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				asCellAlignments = Split(",,,RIGHT", ",", -1, vbBinaryCompare)
				dTotal = 0
				dGranTotal = 0
				lCounter = 0
				lGranCounter = 0
				lCurrentID = -2
				lCurrentAreaID = CLng(oRecordset.Fields("AreaID1").Value)
				sRowContents = "<SPAN COLS=""4"" /><B>UNIDAD ADMINISTRATIVA: " & CleanStringForHTML(CStr(oRecordset.Fields("AreaName1").Value)) & "</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				Do While Not oRecordset.EOF
					If lCurrentID <> CLng(oRecordset.Fields("AreaID2").Value) Then
						If lCurrentID <> -2 Then
							sRowContents = "&nbsp;" & TABLE_SEPARATOR & "<B>Registros: " & FormatNumber(lCounter, 0, True, False, True) & "</B>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B>Total por centro de trabajo:</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(dTotal, 2, True, False, True) & "</B>"
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
							dGranTotal = dGranTotal + dTotal
							lGranCounter = lGranCounter + lCounter
							dTotal = 0
							lCounter = 0
							asRowContents = Split("<SPAN COLS=""4"" />&nbsp;", TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
						End If

						If lCurrentAreaID <> CLng(oRecordset.Fields("AreaID1").Value) Then
							sRowContents = "&nbsp;" & TABLE_SEPARATOR & "<B>Registros: " & FormatNumber(lGranCounter, 0, True, False, True) & "</B>"
							sRowContents = sRowContents & TABLE_SEPARATOR & "<B>Total por delegación:</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(dGranTotal, 2, True, False, True) & "</B>"
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
							dGranTotal = 0
							lGranCounter = 0
							asRowContents = Split("<SPAN COLS=""4"" />&nbsp;", TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If

							lCurrentAreaID = CLng(oRecordset.Fields("AreaID1").Value)
							sRowContents = "<SPAN COLS=""4"" /><B>UNIDAD ADMINISTRATIVA: " & CleanStringForHTML(CStr(oRecordset.Fields("AreaName1").Value)) & "</B>"
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
						End If

						sRowContents = "<SPAN COLS=""4"" /><B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode2").Value) & ". " & CStr(oRecordset.Fields("AreaName2").Value)) & "</B>"
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						lCurrentID = CLng(oRecordset.Fields("AreaID2").Value)
					End If
					sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
					If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value))
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value))
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(lForPayrollID, -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True)
					dTotal = dTotal + CDbl(oRecordset.Fields("TotalAmount").Value)
					lCounter = lCounter + 1

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If

					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close

				sRowContents = "&nbsp;" & TABLE_SEPARATOR & "<B>Registros: " & FormatNumber(lCounter, 0, True, False, True) & "</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>Total por centro de trabajo:</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(dTotal, 2, True, False, True) & "</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				dGranTotal = dGranTotal + dTotal
				lGranCounter = lGranCounter + lCounter

				asRowContents = Split("<SPAN COLS=""4"" />&nbsp;", TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If

				sRowContents = "&nbsp;" & TABLE_SEPARATOR & "<B>Registros: " & FormatNumber(lGranCounter, 0, True, False, True) & "</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>Total por delegación:</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(dGranTotal, 2, True, False, True) & "</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			Response.Write "</TABLE>"
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1425 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1426(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the employee's records for concepts
'         15, 31, C5 sort by RecordID
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1426"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim lStartPayrollDate
	Dim sPeriod
	Dim iCounter
	Dim sCurrentRFC
	Dim dTotalHours
	Dim dTotalAmount
	Dim oRecordset
	Dim sRowContents
	Dim sDate
	Dim sFileName
	Dim lReportID
	Dim oStartDate
	Dim oEndDate
	Dim lErrorNumber

	oStartDate = Now()
	Call GetConditionFromURL(oRequest, "", lPayrollID, lForPayrollID)
	If Len(oRequest("SubAreaID").Item) > 0 Then
		sCondition = " And (Areas2.AreaID In (" & oRequest("SubAreaID").Item & "))"
	ElseIf Len(oRequest("AreaID").Item) > 0 Then
		sCondition = " And (Areas1.AreaID=" & oRequest("AreaID").Item & ")"
	End If
	If Len(oRequest("ConceptID").Item) > 0 Then
		Select Case CLng(oRequest("ConceptID").Item)
			Case 18
				sCondition = sCondition & " And (EmployeesSpecialJourneys.SpecialJourneyID=423)"
			Case 34
				sCondition = sCondition & " And (EmployeesSpecialJourneys.SpecialJourneyID=424)"
			Case 96
				sCondition = sCondition & " And (EmployeesSpecialJourneys.SpecialJourneyID=425)"
		End Select
	End If
	sErrorDescription = "No se pudieron obtener los montos registrados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeNumber, RFC, EmployeeName, EmployeeLastName, EmployeeLastName2, EmployeesSpecialJourneys.WorkingHours, Areas2.AreaShortName, PositionShortName, PositionName, LevelShortName, EmployeesSpecialJourneys.AppliedDate, EmployeesSpecialJourneys.StartDate, EmployeesSpecialJourneys.EndDate, WorkedHours, ConceptAmount From EmployeesSpecialJourneys, Areas As Areas1, Areas As Areas2, Positions, Levels Where (EmployeesSpecialJourneys.PositionID=Positions.PositionID) And (EmployeesSpecialJourneys.LevelID=Levels.LevelID) And (EmployeesSpecialJourneys.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesSpecialJourneys.AppliedDate=" & lPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (EmployeeID>=800000) " & sCondition & " Order By EmployeeLastName, EmployeeLastName2, EmployeeName", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select EmployeeNumber, RFC, EmployeeName, EmployeeLastName, EmployeeLastName2, EmployeesSpecialJourneys.WorkingHours, Areas2.AreaShortName, PositionShortName, PositionName, LevelShortName, EmployeesSpecialJourneys.AppliedDate, EmployeesSpecialJourneys.StartDate, EmployeesSpecialJourneys.EndDate, WorkedHours, ConceptAmount From EmployeesSpecialJourneys, Areas As Areas1, Areas As Areas2, Positions, Levels Where (EmployeesSpecialJourneys.PositionID=Positions.PositionID) And (EmployeesSpecialJourneys.LevelID=Levels.LevelID) And (EmployeesSpecialJourneys.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesSpecialJourneys.AppliedDate=" & lPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (EmployeeID>=800000) " & sCondition & " Order By EmployeeLastName, EmployeeLastName2, EmployeeName -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sDate = GetSerialNumberForDate("")
			sFileName = REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_" & aReportsComponent(N_ID_REPORTS) & "_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & sDate
			Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName & ".zip") & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
			Response.Flush()

			lStartPayrollDate = GetPayrollStartDate(lForPayrollID)
			sPeriod = Left(CStr(lStartPayrollDate), Len("0000")) & "-" & Mid(CStr(lStartPayrollDate), Len("00000"), Len("00")) & "-" & Right(CStr(lStartPayrollDate), Len("00")) & " " & Left(CStr(oRecordset.Fields("AppliedDate").Value), Len("0000")) & "-" & Mid(CStr(oRecordset.Fields("AppliedDate").Value), Len("00000"), Len("00")) & "-" & Right(CStr(oRecordset.Fields("AppliedDate").Value), Len("00"))
			iCounter = 1
			sCurrentRFC = ""
			dTotalHours = 0
			dTotalAmount = 0
			Do While Not oRecordset.EOF
				If StrComp(sCurrentRFC, CStr(oRecordset.Fields("RFC").Value), vbBinaryCompare) <> 0 Then
					If Len(sCurrentRFC) > 0 Then
						sRowContents = Replace(sRowContents, "<TOTAL_HOURS />", dTotalHours)
						sRowContents = Replace(sRowContents, "<TOTAL_AMOUNT />", Right(("         " & FormatNumber(dTotalAmount, 2, True, False, True)), Len("         ")))
						lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".txt"), sRowContents, sErrorDescription)
						dTotalHours = 0
						dTotalAmount = 0
					End If
				
					sRowContents = "                             "
					sRowContents = sRowContents & CStr(oRecordset.Fields("PositionShortName").Value) & "   "
					sRowContents = sRowContents & CStr(oRecordset.Fields("LevelShortName").Value) & " "
					sRowContents = sRowContents & sPeriod & "   " & "<TOTAL_HOURS />"
					sRowContents = sRowContents & "        "
					sRowContents = sRowContents & "$<TOTAL_AMOUNT />"
					sRowContents = sRowContents & vbNewLine & vbNewLine & vbNewLine & vbNewLine
					If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
						sRowContents = sRowContents & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & ", " & CStr(oRecordset.Fields("EmployeeName").Value)
					Else
						sRowContents = sRowContents & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value)
					End If
					sRowContents = sRowContents & vbNewLine & vbNewLine & vbNewLine & vbNewLine

					sRowContents = sRowContents & "  " & CStr(oRecordset.Fields("EmployeeNumber").Value)
					sRowContents = sRowContents & "           "
					sRowContents = sRowContents & CStr(oRecordset.Fields("RFC").Value)
					sRowContents = sRowContents & vbNewLine & vbNewLine & vbNewLine & vbNewLine

					sRowContents = sRowContents & "  " & CStr(oRecordset.Fields("AreaShortName").Value)
					sRowContents = sRowContents & "       "
					sRowContents = sRowContents & CStr(oRecordset.Fields("PositionShortName").Value)
					sRowContents = sRowContents & "   "
					sRowContents = sRowContents & vbNewLine & vbNewLine & vbNewLine & vbNewLine

					sRowContents = sRowContents & "  " & Left((CStr(oRecordset.Fields("PositionName").Value) & "                                                                           "), Len ("                                                                           "))
					sRowContents = sRowContents & "$<TOTAL_AMOUNT />"
					sRowContents = sRowContents & vbNewLine

					sRowContents = sRowContents & "                                                                             "
					sRowContents = sRowContents & "$     0.00"
					sRowContents = sRowContents & vbNewLine

					sRowContents = sRowContents & "  " & Left(CStr(oRecordset.Fields("AppliedDate").Value), Len("0000")) & "-" & Mid(CStr(oRecordset.Fields("AppliedDate").Value), Len("00000"), Len("00")) & "-" & Right(CStr(oRecordset.Fields("AppliedDate").Value), Len("00"))
					sRowContents = sRowContents & Right(("                " & iCounter), Len("                "))
					sRowContents = sRowContents & "                                                 "
					sRowContents = sRowContents & "$<TOTAL_AMOUNT />"
					sRowContents = sRowContents & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine
					sCurrentRFC = CStr(oRecordset.Fields("RFC").Value)
				End If
				If CLng(oRequest("ConceptID").Item) = 34 Then
					dTotalHours = dTotalHours + (Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)), GetDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)))) + 1)
				Else
					dTotalHours = dTotalHours + CDbl(oRecordset.Fields("WorkedHours").Value) * (Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)), GetDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)))) + 1)
				End If
				dTotalAmount = dTotalAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
				iCounter = iCounter + 1
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
			sRowContents = Replace(sRowContents, "<TOTAL_HOURS />", dTotalHours)
			sRowContents = Replace(sRowContents, "<TOTAL_AMOUNT />", Right(("         " & FormatNumber(dTotalAmount, 2, True, False, True)), Len("         ")))
			lErrorNumber = AppendTextToFile(Server.MapPath(sFileName & ".txt"), sRowContents, sErrorDescription)

			lErrorNumber = ZipFile(Server.MapPath(sFileName & ".txt"), Server.MapPath(sFileName & ".zip"), sErrorDescription)
			If lErrorNumber = 0 Then
				Call GetNewIDFromTable(oADODBConnection, "Reports", "ReportID", "", 1, lReportID, sErrorDescription)
				sErrorDescription = "No se pudieron guardar la información del reporte."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Reports (ReportID, UserID, ConstantID, ReportName, ReportDescription, ReportParameters1, ReportParameters2, ReportParameters3) Values (" & lReportID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aReportsComponent(N_ID_REPORTS) & ", " & sDate & ", '" & CATALOG_SEPARATOR & "', '" & oRequest & "', '" & sFlags & "', '')", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				lErrorNumber = DeleteFile(Server.MapPath(sFileName & ".txt"), sErrorDescription)
			End If
			oEndDate = Now()
			If (lErrorNumber = 0) And B_USE_SMTP Then
				If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
			End If
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1426 = lErrorNumber
	Err.Clear
End Function

Function BuildReport1427(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the employee's amounts for concepts
'         15, 31, C5 sort by RecordID
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildReport1427"
	Dim sCondition
	Dim lPayrollID
	Dim lForPayrollID
	Dim iCounter
	Dim dTotal
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	Call GetConditionFromURL(oRequest, "", lPayrollID, lForPayrollID)
	If Len(oRequest("SubAreaID").Item) > 0 Then
		sCondition = " And (Areas2.AreaID In (" & oRequest("SubAreaID").Item & "))"
	ElseIf Len(oRequest("AreaID").Item) > 0 Then
		sCondition = " And (Areas1.AreaID=" & oRequest("AreaID").Item & ")"
	End If
	If Len(oRequest("ConceptID").Item) > 0 Then
		Select Case CLng(oRequest("ConceptID").Item)
			Case 18
				sCondition = sCondition & " And (EmployeesSpecialJourneys.SpecialJourneyID=423)"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>LISTADO DE FIRMAS DE GUARDIAS CORRESPONDIENTE A LA QUINCENA</B></FONT><BR /><BR />"
			Case 34
				sCondition = sCondition & " And (EmployeesSpecialJourneys.SpecialJourneyID=424)"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>LISTADO DE FIRMAS DE SUPLENCIAS CORRESPONDIENTE A LA QUINCENA</B></FONT><BR /><BR />"
			Case 96
				sCondition = sCondition & " And (EmployeesSpecialJourneys.SpecialJourneyID=425)"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>LISTADO DE FIRMAS DE REZAGO QUIRÚRGICO CORRESPONDIENTE A LA QUINCENA</B></FONT><BR /><BR />"
		End Select
	End If
	sErrorDescription = "No se pudieron obtener los montos registrados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeNumber, RFC, EmployeeName, EmployeeLastName, EmployeeLastName2, EmployeesSpecialJourneys.WorkingHours, Areas2.AreaShortName, PositionShortName, LevelShortName, RiskLevelID, Sum(ConceptAmount) As TotalAmount From EmployeesSpecialJourneys, Areas As Areas1, Areas As Areas2, Positions, Levels, SpecialJourneys, SpecialJourneysMovements, SpecialJourneysReasons Where (EmployeesSpecialJourneys.PositionID=Positions.PositionID) And (EmployeesSpecialJourneys.LevelID=Levels.LevelID) And (EmployeesSpecialJourneys.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesSpecialJourneys.AppliedDate=" & lPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (EmployeeID>=800000) " & sCondition & " Group by EmployeeNumber, RFC, EmployeeName, EmployeeLastName, EmployeeLastName2, EmployeesSpecialJourneys.WorkingHours, Areas2.AreaShortName, PositionShortName, LevelShortName, RiskLevelID Order By EmployeeNumber", "ReportsQueries1400Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select EmployeeNumber, RFC, EmployeeName, EmployeeLastName, EmployeeLastName2, EmployeesSpecialJourneys.WorkingHours, Areas2.AreaShortName, PositionShortName, LevelShortName, RiskLevelID, Sum(ConceptAmount) As TotalAmount From EmployeesSpecialJourneys, Areas As Areas1, Areas As Areas2, Positions, Levels, SpecialJourneys, SpecialJourneysMovements, SpecialJourneysReasons Where (EmployeesSpecialJourneys.PositionID=Positions.PositionID) And (EmployeesSpecialJourneys.LevelID=Levels.LevelID) And (EmployeesSpecialJourneys.AreaID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesSpecialJourneys.AppliedDate=" & lPayrollID & ") And (Positions.StartDate<=" & lForPayrollID & ") And (Positions.EndDate>=" & lForPayrollID & ") And (Levels.StartDate<=" & lForPayrollID & ") And (Levels.EndDate>=" & lForPayrollID & ") And (Areas1.StartDate<=" & lForPayrollID & ") And (Areas1.EndDate>=" & lForPayrollID & ") And (Areas2.StartDate<=" & lForPayrollID & ") And (Areas2.EndDate>=" & lForPayrollID & ") And (EmployeeID>=800000) " & sCondition & " Group by EmployeeNumber, RFC, EmployeeName, EmployeeLastName, EmployeeLastName2, EmployeesSpecialJourneys.WorkingHours, Areas2.AreaShortName, PositionShortName, LevelShortName, RiskLevelID Order By EmployeeNumber -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("Registro,No. del empleado,RFC,Nombre,Puesto,RP,Nivel,Jornada,Percepción,Deducción,Total,Firma", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				asCellAlignments = Split(",,,,,,,,RIGHT,RIGHT,RIGHT,", ",", -1, vbBinaryCompare)
				iCounter = 1
				dTotal = 0
				Do While Not oRecordset.EOF
					sRowContents = iCounter
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value))
					If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & ", " & CStr(oRecordset.Fields("EmployeeName").Value))
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value))
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
					If CLng(oRecordset.Fields("RiskLevelID").Value) > 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & CStr(oRecordset.Fields("RiskLevelID").Value)
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("LevelShortName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("WorkingHours").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & "$ " & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & "$ 0.00"
					sRowContents = sRowContents & TABLE_SEPARATOR & "$ " & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True)
					dTotal = dTotal + CDbl(oRecordset.Fields("TotalAmount").Value)
					sRowContents = sRowContents & TABLE_SEPARATOR & "________________________________"

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					iCounter = iCounter + 1
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop

				sRowContents = "<SPAN COLS=""10"" />&nbsp;"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				sRowContents = "<SPAN COLS=""7"" />&nbsp;"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>TOTALES</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(dTotal, 2, True, False, True) & "</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>0.00</B>"
				sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(dTotal, 2, True, False, True) & "</B>"
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If

				oRecordset.Close
			Response.Write "</TABLE>"
		End If
	End If

	Set oRecordset = Nothing
	BuildReport1427 = lErrorNumber
	Err.Clear
End Function

Function UpdateCLC(oRequest, oADODBConnection, lPayrollCLC, lPayrollCode,lPayrollID,lConditions, sErrorDescription)

	On Error Resume Next
	Const S_FUNCTION_NAME = "UpdateCLC"	

	Set oADODBCommand = Server.CreateObject("ADODB.Command")
	Set oADODBCommand.ActiveConnection = oADODBConnection
	oADODBCommand.commandtype=4
	oADODBCommand.commandtext = "SIAP.UpdateCLC"
	Set param = oADODBCommand.Parameters
	param.append oADODBCommand.createparameter("lPayrollCLC",8,1)
	param.append oADODBCommand.createparameter("lPayrollCode",8,1)
    param.append oADODBCommand.createparameter("lPayrollID",3,1)
    param.append oADODBCommand.createparameter("lConditions",8,1)
  
 	oADODBCommand("lPayrollCLC") = lPayrollCLC
	oADODBCommand("lPayrollCode") = lPayrollCode
    oADODBCommand("lPayrollID") = lPayrollID
    oADODBCommand("lConditions") = lConditions
	
	oADODBCommand.Execute

	Set oADODBCommand = Nothing
	Set param = Nothing
	UpdateCLC = lErrorNumber
    Err.Clear
End Function

Function UpdateCLC_RPT(oRequest, oADODBConnection, lPayrollCLC, lPayrollCode,lPayrollID,lPayrollTypeID, lPayrollDescription, lMemorandum, lPayrollFile, lCancelDate, lPayrollYear, lPayrollMonth, lFortNightly,lConditions, sErrorDescription)

	On Error Resume Next
	Const S_FUNCTION_NAME = "UpdateCLC_RPT"	
    Dim iCount
    
	Set oADODBCommand = Server.CreateObject("ADODB.Command")
	Set oADODBCommand.ActiveConnection = oADODBConnection
	oADODBCommand.commandtype=4
	oADODBCommand.commandtext = "SIAP.UpdateCLC_RPT"
	Set param = oADODBCommand.Parameters
	param.append oADODBCommand.createparameter("lPayrollCLC",8,1)
	param.append oADODBCommand.createparameter("lPayrollCode",8,1)
    param.append oADODBCommand.createparameter("lPayrollID",3,1)
    param.append oADODBCommand.createparameter("lPayrollTypeID",3,1)
    param.append oADODBCommand.createparameter("lPayrollDescription",8,1)
    param.append oADODBCommand.createparameter("lMemorandum",8,1)
    param.append oADODBCommand.createparameter("lPayrollFile",8,1)
    param.append oADODBCommand.createparameter("lCancelDate",3,1)
    param.append oADODBCommand.createparameter("lPayrollYear", 3, 1)
    param.append oADODBCommand.createparameter("lPayrollMonth", 3, 1)
    param.append oADODBCommand.createparameter("lFortNightly", 8, 1)
    param.append oADODBCommand.createparameter("lConditions",8,1)  
    param.append oADODBCommand.createparameter("lcount", 3, 2)

	oADODBCommand("lPayrollCLC") = lPayrollCLC
	oADODBCommand("lPayrollCode") = lPayrollCode
    oADODBCommand("lPayrollID") = lPayrollID
    oADODBCommand("lPayrollTypeID") = lPayrollTypeID
    oADODBCommand("lPayrollDescription") = lPayrollDescription
    oADODBCommand("lMemorandum") = lMemorandum
    oADODBCommand("lPayrollFile") = lPayrollFile
    oADODBCommand("lCancelDate") = lCancelDate
    oADODBCommand("lPayrollYear") = lPayrollYear
    oADODBCommand("lPayrollMonth") = lPayrollMonth
    oADODBCommand("lFortNightly") = lFortNightly
    oADODBCommand("lConditions") = lConditions

   If (StrComp(lCancelDate,"",vbBinaryCompare) = 0 ) Then
        oADODBCommand("lCancelDate").Value= Mid(lPayrollID,1,4)
   End If
   If (StrComp(lPayrollTypeID,"",vbBinaryCompare) = 0 ) Then
        oADODBCommand("lPayrollTypeID") = 0
   End If
	
	oADODBCommand.Execute
    iCount = oADODBCommand("lcount")

	Set oADODBCommand = Nothing
	Set param = Nothing
	UpdateCLC_RPT = lErrorNumber
    Err.Clear
End Function

Function prc_actualiza_clc(oRequest, oADODBConnection, lPayrollCLC, lPayrollCode,lPayrollID, sErrorDescription)

	On Error Resume Next
	Const S_FUNCTION_NAME = "prc_actualiza_clc"	

	Set oADODBCommand = Server.CreateObject("ADODB.Command")
	Set oADODBCommand.ActiveConnection = oADODBConnection
	oADODBCommand.commandtype=4
	oADODBCommand.commandtext = "SIAP.prc_actualiza_clc"
	'oADODBCommand.commandtext = "SIAP."&lProcedure
    Set param = oADODBCommand.Parameters
	param.append oADODBCommand.createparameter("lPayrollCLC",8,1)
	param.append oADODBCommand.createparameter("lPayrollCode",8,1)
    param.append oADODBCommand.createparameter("lPayrollID",3,1)

	oADODBCommand("lPayrollCLC") = lPayrollCLC
	oADODBCommand("lPayrollCode") = lPayrollCode
    oADODBCommand("lPayrollID") = lPayrollID
	
	oADODBCommand.Execute

	Set oADODBCommand = Nothing
	Set param = Nothing
	prc_actualiza_clc = lErrorNumber
    Err.Clear
End Function

%>