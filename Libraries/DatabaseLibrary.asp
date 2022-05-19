<!-- #include file="AuditComponent.asp" -->
<%
Const SQL_SERVER = 1
Const ACCESS = 2
Const ACCESS_DSN = 3
Const ORACLE = 4
Const MYSQL = 5
Const SQL_SERVER_64_OLE = 6
Const SQL_SERVER_64_DSNLess = 7
Const SQL_SERVER_2008 = 8
Dim S_WILD_CHAR

Function CreateADODBConnection(sDatabasePath, sUserName, sPassword, iConnectionType, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To create the connection with the database
'Inputs:  sDatabasePath, sUserName, sPassword, iConnectionType
'Outputs: oADODBConnection, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CreateADODBConnection"
	Dim sDBPath
	Dim lErrorNumber

	If InStr(1, sDatabasePath, ":", vbBinaryCompare) <> 2 Then
		sDBPath = Server.MapPath(sDatabasePath)
		If Err.number <> 0 Then
			Err.Clear
			sDBPath = sDatabasePath
		End If
	Else
		sDBPath = sDatabasePath
	End If

	Set oADODBConnection = Server.CreateObject("ADODB.Connection")
	lErrorNumber = Err.number
	If lErrorNumber <> 0 Then
		sErrorDescription = "No se pudo crear una instancia del objeto 'ADODB.Connection' porque el archivo 'msado10.dll' no se encuentra correctamente registrado en el servidor de Web. Favor de contactar al Administrador."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "DatabaseLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	Else
		oADODBConnection.CommandTimeout = 3600
		oADODBConnection.ConnectionTimeout = 150
		Select Case iConnectionType
			Case SQL_SERVER, MYSQL, ACCESS_DSN
				sErrorDescription = "No se pudo abrir la base de datos '" & sDatabasePath & "'"
				Call oADODBConnection.Open(sDatabasePath, sUserName, sPassword)
				lErrorNumber = Err.number
				If lErrorNumber <> 0 Then
					Call LogErrorInXMLFile(lErrorNumber, sErrorDescription & ".<BR /><B>Error del servidor Web: </B>" & Err.Description, 000, "DatabaseLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
					Err.Clear
					oADODBConnection.DefaultDatabase = SIAP_DATABASE_NAME
					Call oADODBConnection.Open(sDatabasePath, sUserName, sPassword)
					lErrorNumber = Err.number
					If lErrorNumber <> 0 Then
						Call LogErrorInXMLFile(lErrorNumber, sErrorDescription & ".<BR /><B>Error del servidor Web: </B>" & Err.Description, 000, "DatabaseLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
					End If
				End If
				If iConnectionType = ACCESS_DSN Then
					S_WILD_CHAR = "%"
				Else
					S_WILD_CHAR = "%"
				End If
			Case ACCESS
				Call oADODBConnection.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDBPath & ";User ID=" & sUserName & ";Password=" & sPassword & ";")
				sErrorDescription = "No se pudo abrir la base de datos '" & sDBPath & "'"
				S_WILD_CHAR = "%"
			Case ORACLE
				'Call oADODBConnection.Open("Provider=OraOLEDB.Oracle;Data Source=" & sDatabasePath & ";User ID=" & sUserName & ";Password=" & sPassword & ";FetchSize=200;CacheType=Memory;OSAuthent=0;PLSQLRSet=1;")
				'Call oADODBConnection.Open("Provider=OraOLEDB.Oracle;Data Source=" & sDatabasePath & ";User ID=" & sUserName & ";Password=" & sPassword & ";")
				'Call oADODBConnection.Open("DSN=" & sDatabasePath & ";Uid=" & sUserName & ";Pwd=" & sPassword & ";")
				Call oADODBConnection.Open("DSN=" & sDatabasePath & ";Uid=" & sUserName & ";Pwd=" & sPassword & ";")
				sErrorDescription = "No se pudo abrir la base de datos '" & sDatabasePath & "'"
				S_WILD_CHAR = "%"
			Case SQL_SERVER_64_OLE, SQL_SERVER_64_DSNLess
				sErrorDescription = "No se pudo abrir la base de datos '" & sDatabasePath & "'."
				If iConnectionType = SQL_SERVER_64_OLE Then
					Call oADODBConnection.Open("Provider=MSDASQL;Driver={SQL Server};Server=" & SERVER_NAME_FOR_LICENSE & ";UID=" & sUserName & ";PWD=" & sPassword & ";")
				Else
					Call oADODBConnection.Open("Provider=SQLOLEDB;Data Source=" & sDatabasePath & ";User ID=" & sUserName & ";Password=" & sPassword & ";")
				End If
				lErrorNumber = Err.number
				If lErrorNumber <> 0 Then
					Call LogErrorInXMLFile(lErrorNumber, sErrorDescription & ".<BR /><B>Error del servidor Web:&nbsp;</B>" & Err.Description, 000, "DatabaseLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
					Err.Clear
					oADODBConnection.DefaultDatabase = SIAP_DATABASE_NAME
					Call oADODBConnection.Open(sDatabasePath, sUserName, sPassword)
					lErrorNumber = Err.number
					If lErrorNumber <> 0 Then
						Call LogErrorInXMLFile(lErrorNumber, sErrorDescription & ".<BR /><B>Error del servidor Web:&nbsp;</B>" & Err.Description, 000, "DatabaseLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
					End If
				End If
				S_WILD_CHAR = "%"
			Case SQL_SERVER_2008
				sErrorDescription = "No se pudo abrir la base de datos '" & sDatabasePath & "'."
				'Call oADODBConnection.Open("Provider=SQLOLEDB.1;Server=" & SERVER_IP_FOR_LICENSE & ";Data Source=" & sDatabasePath & ";Database=" & SIAP_DATABASE_NAME & ";UID=" & sUserName & ";PWD=" & sPassword & ";Integrated Security=SSPI;Persist Security Info=True;Trusted_Connection=Yes;")
				Call oADODBConnection.Open("Provider=SQLNCLI10;Driver={SQL Server};Server=" & SERVER_IP_FOR_LICENSE & ";Data Source=" & sDatabasePath & ";Database=" & SIAP_DATABASE_NAME & ";UID=" & sUserName & ";PWD=" & sPassword & ";Integrated Security=SSPI;Persist Security Info=True;Trusted_Connection=Yes;")
				lErrorNumber = Err.number
				If lErrorNumber <> 0 Then
					Call LogErrorInXMLFile(lErrorNumber, sErrorDescription & ".<BR /><B>Error del servidor Web:&nbsp;</B>" & Err.Description, 000, "DatabaseLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
					Err.Clear
					oADODBConnection.DefaultDatabase = SIAP_DATABASE_NAME
					Call oADODBConnection.Open(sDatabasePath, sUserName, sPassword)
					lErrorNumber = Err.number
					If lErrorNumber <> 0 Then
						Call LogErrorInXMLFile(lErrorNumber, sErrorDescription & ".<BR /><B>Error del servidor Web:&nbsp;</B>" & Err.Description, 000, "DatabaseLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
					End If
				End If
				S_WILD_CHAR = "%"
		End Select
		lErrorNumber = Err.number
		If lErrorNumber <> 0 Then
			If Len(sUserName) > 0 Then sErrorDescription = sErrorDescription & " utilizando el usuario '" & sUserName & "'"
			sErrorDescription = sErrorDescription & ". Favor de contactar al Administrador."
			If Len(Err.description) > 0 Then
				sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
			End If
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "DatabaseLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
		Else
			sErrorDescription = ""
		End If
	End If

	CreateADODBConnection = lErrorNumber
	Err.Clear
End Function

Function ShowADODBConnectionProperties(oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To show the database connection properties
'Inputs:  oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ShowADODBConnectionProperties"
	Dim lErrorNumber

	If Not IsObject(oADODBConnection) Then
		lErrorNumber = L_ERR_NO_DB_CONNECTION
		sErrorDescription = "No existe una conexión con la base de datos.<BR />"
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "DatabaseLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	Else
		Response.Write "<B>PROPIEDADES DE LA BASE DE DATOS</B><BR /><BR />"
		Response.Write "<B>Properties.Count</B>: " & oADODBConnection.Properties.Count & "<BR />"
		Response.Write "<B>ConnectionString</B>: " & oADODBConnection.ConnectionString & "<BR />"
		Response.Write "<B>CommandTimeout</B>: " & oADODBConnection.CommandTimeout & "<BR />"
		Response.Write "<B>ConnectionTimeout</B>: " & oADODBConnection.ConnectionTimeout & "<BR />"
		Response.Write "<B>Version</B>: " & oADODBConnection.Version & "<BR />"
		Response.Write "<B>Errors.Count</B>: " & oADODBConnection.Errors.Count & "<BR />"
		Response.Write "<B>DefaultDatabase</B>: " & oADODBConnection.DefaultDatabase & "<BR />"
		Response.Write "<B>IsolationLevel</B>: " & oADODBConnection.IsolationLevel & "<BR />"
		Response.Write "<B>Attributes</B>: " & oADODBConnection.Attributes & "<BR />"
		Response.Write "<B>CursorLocation</B>: " & oADODBConnection.CursorLocation & "<BR />"
		Response.Write "<B>Mode</B>: " & oADODBConnection.Mode & "<BR />"
		Response.Write "<B>Provider</B>: " & oADODBConnection.Provider & "<BR />"
		Response.Write "<B>State</B>: " & oADODBConnection.State & "<BR />"
	End If
	ShowADODBConnectionProperties = lErrorNumber
	Err.Clear
End Function

Function ExecuteInsertQuerySp(oADODBConnection, sQuery, sLibraryFile, sFuntionName, iDescriptorID, sErrorDescription)
'************************************************************
'Purpose: To execute an Insert query using the DB connection
'         ignoring the error -2147217900 (duplicated entries)
'Inputs:  oADODBConnection, sQuery, sLibraryFile, sFuntionName, iDescriptorID, sErrorDescription
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ExecuteInsertQuerySp"
	Dim sTable
	Dim iTableNamePos
	Dim lErrorNumber

	If Not IsObject(oADODBConnection) Then
		lErrorNumber = L_ERR_NO_DB_CONNECTION
		sErrorDescription = "No existe una conexión con la base de datos.<BR />"
		sErrorDescription = sErrorDescription & "Librería: " & sLibraryFile & "<BR />"
		sErrorDescription = sErrorDescription & "Función: " & sFuntionName & "<BR />"
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "DatabaseLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	ElseIf (StrComp(sLibraryFile, "PayrollComponent.asp", vbBinaryCompare) = 0) And (InStr(1, "CalculatePayroll,DoCalculations", sFuntionName, vbBinaryCompare) > 0) And (FileExists(Server.MapPath("Database\Stop.txt"), "")) Then
		lErrorNumber = -69
		sErrorDescription = "Interrupción del proceso de prenómina."
	Else
		If iConnectionType <> ORACLE Then
			Call oADODBConnection.Execute(sQuery, null, 128)
		Else
			Call oADODBConnection.Execute(Replace(Replace(sQuery, " As ", " ", vbTextCompare), " as ", " ", vbTextCompare), null, 128)
		End If
		lErrorNumber = Err.number
		If (lErrorNumber <> 0) And (lErrorNumber <> -2147217900) Then
			If Len(sErrorDescription) > 0 Then sErrorDescription = sErrorDescription & "<BR />"
			sErrorDescription = sErrorDescription & "Query: " & sQuery & "<BR />"
			If Len(Err.description) > 0 Then
				sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor de Web: </B>" & Err.description & "<BR />"
			End If
			sErrorDescription = sErrorDescription & "<B>Librería: </B>" & sLibraryFile & "<BR />"
			sErrorDescription = sErrorDescription & "<B>Función: </B>" & sFuntionName & "<BR />"
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, iDescriptorID, "DatabaseLibrary.asp", S_FUNCTION_NAME, N_SQL_ERROR_LEVEL)
			bTimeout = (lErrorNumber = -2147217871)
'		ElseIf (lErrorNumber = -2147217900) Then
'			If Len(sErrorDescription) > 0 Then sErrorDescription = sErrorDescription & "<BR />"
'			sErrorDescription = sErrorDescription & "Query: " & sQuery & "<BR />"
'			If Len(Err.description) > 0 Then
'				sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor de Web: </B>" & Err.description & "<BR />"
'			End If
'			sErrorDescription = sErrorDescription & "<B>Librería: </B>" & sLibraryFile & "<BR />"
'			sErrorDescription = sErrorDescription & "<B>Función: </B>" & sFuntionName & "<BR />"
'			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, iDescriptorID, "DatabaseLibrary.asp", S_FUNCTION_NAME, N_SQL_ERROR_LEVEL)
'			lErrorNumber = 0
'			sErrorDescription = ""
		Else
'			If (Len(GetAdminOption(aAdminOptionsComponent, 0)) > 0) And (InStr(1, aLoginComponent(S_ACCESS_KEY_LOGIN), "vac", vbBinaryCompare) = 1) Then
'				Response.Write vbNewLine & vbNewLine & "<!-- Start Tracing Query: " & vbNewLine & sQuery & vbNewLine & "End Tracing Query -->" & vbNewLine
'			End If
			lErrorNumber = 0
			sErrorDescription = ""
		End If
	End If

	ExecuteInsertQuerySp = lErrorNumber
	Err.Clear
End Function

Function ExecuteSQLQuery(oADODBConnection, sQuery, sLibraryFile, sFuntionName, iDescriptorID, sErrorDescription, oRecordset)
'************************************************************
'Purpose: To execute a SQL query using the DB connection
'Inputs:  oADODBConnection, sQuery, sLibraryFile, sFuntionName, iDescriptorID, sErrorDescription
'Outputs: sErrorDescription, oRecordset
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ExecuteSQLQuery"
	Dim sTable
	Dim iTableNamePos
	Dim lRecordsAffected
	Dim lErrorNumber
	Dim sTempTop
	Dim lRows
	Dim lStartPos
	Dim lEndPos

	If Not IsObject(oADODBConnection) Then
		lErrorNumber = L_ERR_NO_DB_CONNECTION
		sErrorDescription = "No existe una conexión con la base de datos.<BR />"
		sErrorDescription = sErrorDescription & "Librería: " & sLibraryFile & "<BR />"
		sErrorDescription = sErrorDescription & "Función: " & sFuntionName & "<BR />"
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "DatabaseLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	ElseIf (StrComp(sLibraryFile, "PayrollComponent.asp", vbBinaryCompare) = 0) And (InStr(1, "CalculatePayroll,DoCalculations", sFuntionName, vbBinaryCompare) > 0) And (FileExists(Server.MapPath("Database\Stop.txt"), "")) Then
		lErrorNumber = -69
		sErrorDescription = "Interrupción del proceso de prenómina."
	Else
		If (InStr(1, sQuery, "Select", vbTextCompare) > 0) Or (InStr(1, sQuery, " ", vbBinaryCompare) = 0) Then
			If iConnectionType <> ORACLE Then
				Set oRecordset = oADODBConnection.Execute(sQuery, lRecordsAffected, -1)
			Else
				If InStr(1, sQuery, " Top ", vbBinaryCompare) > 0 Then
					lStartPos = InStr(1, sQuery, " Top ", vbBinaryCompare)
					lEndPos = InStr(lStartPos + Len(" Top "), sQuery, " ", vbBinaryCompare)
					sTempTop = Mid(sQuery, lStartPos, (lEndPos - lStartPos))
					lRows = Mid(sQuery, (lStartPos + Len(" Top ")), (lEndPos - lStartPos - Len(" Top ")))
					sQuery = Replace(sQuery, sTempTop, " ", vbTextCompare)
					sQuery = "Select * From (" & sQuery & ") Where ROWNUM <= " & lRows
				End If
				Set oRecordset = oADODBConnection.Execute(Replace(Replace(Replace(Replace(Replace(sQuery, "''", "' '", vbTextCompare), " As ", " ", vbTextCompare), " as ", " ", vbTextCompare)," SUBSTRING "," SUBSTR "), "+ ' ' +", "|| ' ' ||", vbTextCompare), lRecordsAffected, -1)
			End If
		Else
			If ((iConnectionType = SQL_SERVER) Or (iConnectionType = ORACLE)) And (InStr(1, sQuery, "Delete", vbTextCompare) > 0) Then
				If (InStr(1, sQuery, "Where", vbTextCompare) = 0) Then
					sQuery = Replace(sQuery,"Delete From","Truncate Table",vbTextCompare)
				End If
			End If
			If iConnectionType <> ORACLE Then
				Call oADODBConnection.Execute(sQuery, lRecordsAffected, 128)
			Else
				Call oADODBConnection.Execute(Replace(Replace(Replace(Replace(sQuery, "''", "' '", vbTextCompare), " As ", " ", vbTextCompare), " as ", " ", vbTextCompare), "+ ' ' +", "|| ' ' ||", vbTextCompare), lRecordsAffected, 128)
			End If
		End If
		lErrorNumber = Err.number
		If lErrorNumber <> 0 Then
			If Len(sErrorDescription) > 0 Then sErrorDescription = sErrorDescription & "<BR />"
			sErrorDescription = sErrorDescription & "Query: " & sQuery & "<BR />"
			If Len(Err.description) > 0 Then
				sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor de Web: </B>" & Err.description & "<BR />"
			End If
			sErrorDescription = sErrorDescription & "<B>Librería: </B>" & sLibraryFile & "<BR />"
			sErrorDescription = sErrorDescription & "<B>Función: </B>" & sFuntionName & "<BR />"
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, iDescriptorID, "DatabaseLibrary.asp", S_FUNCTION_NAME, N_SQL_ERROR_LEVEL)
			bTimeout = (lErrorNumber = -2147217871)
		Else
			sTable = Trim(sQuery)
			'False|UPDATE_OPTION
			If (InStr(1, sTable, "Update", vbBinaryCompare) = 1) And True Then 'UPDATE_OPTION
				iTableNamePos = InStr(1, sQuery, "Update ", vbBinaryCompare) + Len("Update ")
				sTable = "," & Mid(sQuery, iTableNamePos, (InStr(iTableNamePos, sQuery, " ") - iTableNamePos)) & ","
				If InStr(1, S_EXCLUDED_TABLES, sTable, vbBinaryCompare) = 0 Then
					Call LogErrorInXMLFile(0, ("Query: " & sQuery & "<BR />" & "<B>Librería: </B>" & sLibraryFile & "<BR />" & "<B>Función: </B>" & sFuntionName & "<BR />"), iDescriptorID, "DatabaseLibrary.asp", S_FUNCTION_NAME, N_SQL_QUERY_LEVEL)
				End If
			'False|DELETE_OPTION
			ElseIf (InStr(1, sTable, "Delete", vbBinaryCompare) = 1) And True Then 'DELETE_OPTION
				iTableNamePos = InStr(1, sQuery, "Delete ", vbBinaryCompare) + Len("Delete ")
				sTable = "," & Mid(sQuery, iTableNamePos, (InStr(iTableNamePos, sQuery, " ") - iTableNamePos)) & ","
				If InStr(1, S_EXCLUDED_TABLES, sTable, vbBinaryCompare) = 0 Then
					Call LogErrorInXMLFile(0, ("Query: " & sQuery & "<BR />" & "<B>Librería: </B>" & sLibraryFile & "<BR />" & "<B>Función: </B>" & sFuntionName & "<BR />"), iDescriptorID, "DatabaseLibrary.asp", S_FUNCTION_NAME, N_SQL_QUERY_LEVEL)
				End If
			'True|INSERT_OPTION
			ElseIf (InStr(1, sTable, "Insert", vbBinaryCompare) = 1) And True Then 'INSERT_OPTION
				iTableNamePos = InStr(1, sQuery, "Insert ", vbBinaryCompare) + Len("Insert ")
				sTable = "," & Mid(sQuery, iTableNamePos, (InStr(iTableNamePos, sQuery, " ") - iTableNamePos)) & ","
				If InStr(1, S_EXCLUDED_TABLES, sTable, vbBinaryCompare) = 0 Then
					Call LogErrorInXMLFile(0, ("Query: " & sQuery & "<BR />" & "<B>Librería: </B>" & sLibraryFile & "<BR />" & "<B>Función: </B>" & sFuntionName & "<BR />"), iDescriptorID, "DatabaseLibrary.asp", S_FUNCTION_NAME, N_SQL_QUERY_LEVEL)
				End If
			ElseIf StrComp(GetASPFileName(""), "QueryConsole.asp", vbTextCompare) = 0 Then
				Call LogErrorInXMLFile(0, ("Query: " & sQuery & "<BR />" & "<B>Librería: </B>" & sLibraryFile & "<BR />" & "<B>Función: </B>" & sFuntionName & "<BR />"), iDescriptorID, "DatabaseLibrary.asp", S_FUNCTION_NAME, N_SQL_QUERY_LEVEL)
				Response.Write vbNewLine & "<!-- Query: " & sQuery & " -->" & vbNewLine
			End If
			If Instr(1, S_TABLES_FOR_AUDIT, "," & sTable & ",", vbBinaryCompare) = 0 Then
				Select Case sTable
					
				End Select
			End If
'			If (Len(GetAdminOption(aAdminOptionsComponent, 0)) > 0) And (InStr(1, aLoginComponent(S_ACCESS_KEY_LOGIN), "vac", vbBinaryCompare) = 1) Then
'				Response.Write vbNewLine & vbNewLine & "<!-- Start Tracing Query: " & vbNewLine & sQuery & vbNewLine & "End Tracing Query -->" & vbNewLine
'			End If
			sErrorDescription = ""
		End If
	End If

	ExecuteSQLQuery = lErrorNumber
	Err.Clear
End Function

Function ExecuteUpdateQuerySp(oADODBConnection, sQuery, sLibraryFile, sFuntionName, iDescriptorID, sErrorDescription)
'************************************************************
'Purpose: To execute an Update query using the DB connection
'         not logging the query in the Error Log
'Inputs:  oADODBConnection, sQuery, sLibraryFile, sFuntionName, iDescriptorID, sErrorDescription
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ExecuteUpdateQuerySp"
	Dim sTable
	Dim iTableNamePos
	Dim lRecordsAffected
	Dim lErrorNumber

	If Not IsObject(oADODBConnection) Then
		lErrorNumber = L_ERR_NO_DB_CONNECTION
		sErrorDescription = "No existe una conexión con la base de datos.<BR />"
		sErrorDescription = sErrorDescription & "Librería: " & sLibraryFile & "<BR />"
		sErrorDescription = sErrorDescription & "Función: " & sFuntionName & "<BR />"
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "DatabaseLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	ElseIf (StrComp(sLibraryFile, "PayrollComponent.asp", vbBinaryCompare) = 0) And (InStr(1, "CalculatePayroll,DoCalculations", sFuntionName, vbBinaryCompare) > 0) And (FileExists(Server.MapPath("Database\Stop.txt"), "")) Then
		lErrorNumber = -69
		sErrorDescription = "Interrupción del proceso de prenómina."
	Else
		If iConnectionType <> ORACLE Then
			Call oADODBConnection.Execute(sQuery, lRecordsAffected, 128)
		Else
			Call oADODBConnection.Execute(Replace(Replace(Replace(Replace(sQuery, "''", "' '", vbTextCompare), " As ", " ", vbTextCompare), " as ", " ", vbTextCompare), "+ ' ' +", "|| ' ' ||", vbTextCompare), lRecordsAffected, 128)
		End If
		lErrorNumber = Err.number
		If lErrorNumber <> 0 Then
			If Len(sErrorDescription) > 0 Then sErrorDescription = sErrorDescription & "<BR />"
			sErrorDescription = sErrorDescription & "Query: " & sQuery & "<BR />"
			If Len(Err.description) > 0 Then
				sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor de Web: </B>" & Err.description & "<BR />"
			End If
			sErrorDescription = sErrorDescription & "<B>Librería: </B>" & sLibraryFile & "<BR />"
			sErrorDescription = sErrorDescription & "<B>Función: </B>" & sFuntionName & "<BR />"
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, iDescriptorID, "DatabaseLibrary.asp", S_FUNCTION_NAME, N_SQL_ERROR_LEVEL)
			bTimeout = (lErrorNumber = -2147217871)
		Else
'			If (Len(GetAdminOption(aAdminOptionsComponent, 0)) > 0) And (InStr(1, aLoginComponent(S_ACCESS_KEY_LOGIN), "vac", vbBinaryCompare) = 1) Then
'				Response.Write vbNewLine & vbNewLine & "<!-- Start Tracing Query: " & vbNewLine & sQuery & vbNewLine & "End Tracing Query -->" & vbNewLine
'			End If
			lErrorNumber = 0
			sErrorDescription = ""
		End If
	End If

	ExecuteUpdateQuerySp = lErrorNumber
	Err.Clear
End Function

Function GenerateCheckboxesFromQuery(oADODBConnection, sTableName, sIDField, sValueField, sCondition, sOrderBy, sListOfCheckedItems, sCollectionName, sErrorDescription)
'************************************************************
'Purpose: To generate checkbox items for every entry in the
'		  recordset that will result by executing the query.
'Inputs:  oADODBConnection, sTableName, sIDField, sValueField, sCondition, sOrderBy, sListOfCheckedItems, sCollectionName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GenerateCheckboxesFromQuery"
	Dim sQuery
	Dim oRecordset
	Dim iIndex
	Dim asID
	Dim asFieldName
	Dim asValue
	Dim sValue
	Dim sSeparator
	Dim lErrorNumber

	sQuery = "Select " & sIDField & ", " & sValueField & " From " & sTableName
	sCondition = Trim(sCondition)
	If Len(sCondition) > 0 Then
		If InStr(1, sCondition, "Where ", vbTextCompare) = 1 Then sCondition = Replace(sCondition, "Where ", "", 1, 1, vbTextCompare)
		If InStr(1, sCondition, "And ", vbTextCompare) = 1 Then sCondition = Replace(sCondition, "And ", "", 1, 1, vbTextCompare)
		sQuery = sQuery & " Where " & sCondition
	End If
	sQuery = sQuery & " Order By "
	If Len(sOrderBy) > 0 Then
		sQuery = sQuery & sOrderBy
	Else
		sQuery = sQuery & sValueField
	End If

	asID = Split(sIDField, ",", -1, vbBinaryCompare)
	asValue = Split(sValueField, ",", -1, vbBinaryCompare)
	If InStr(1, sListOfCheckedItems, LIST_SEPARATOR, vbBinaryCompare) = 0 Then
		sSeparator = ","
	Else
		sSeparator = LIST_SEPARATOR
	End If
	sErrorDescription = "No se pudo ejecutar la petición sobre la base de datos."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "DatabaseLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Do While Not oRecordset.EOF
				If Not IsNull(oRecordset.Fields(0).Value) Then
					Response.Write "<NOBR><INPUT TYPE=""CHECKBOX"" NAME=""" & sCollectionName & """ ID=""" & sCollectionName & "Chk"" VALUE="""
						For iIndex = 0 To UBound(asID)
							Response.Write CleanStringForHTML(CStr(oRecordset.Fields(iIndex).Value))
							If iIndex < UBound(asID) Then Response.Write ","
						Next
					Response.Write """"
						If (InStr(1, sSeparator & sListOfCheckedItems & sSeparator, sSeparator & CStr(oRecordset.Fields(0).Value) & sSeparator, vbTextCompare) > 0) Or (StrComp(sListOfCheckedItems, "True", vbBinaryCompare) = 0) Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " />"
					For iIndex = 0 To UBound(asValue)
						sValue = ""
						sValue = CStr(oRecordset.Fields(UBound(asID) + 1 + iIndex).Value)
						Err.Clear
						If InStr(1, oRecordset.Fields(UBound(asID) + 1 + iIndex).Name, "FormatNumber") = 0 Then
							If B_UPPERCASE Then
								Response.Write CleanStringForHTML(UCase(CStr(sValue)))
							Else
								Response.Write CleanStringForHTML(CStr(sValue))
							End If
						Else
							asFieldName = Split(oRecordset.Fields(UBound(asID) + 1 + iIndex).Name, "_", -1, vbBinaryCompare)
							Response.Write FormatNumber(sValue, asFieldName(1), True, False, True)
						End If
						If iIndex < UBound(asValue) Then Response.Write " "
					Next
					Response.Write "</NOBR><BR />"
					Response.Flush()
				End If
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	GenerateCheckboxesFromQuery = lErrorNumber
	Err.Clear
End Function

Function GenerateJavaScriptArrayFromQuery(oADODBConnection, sTableName, sIDField, sValueField, sCondition, sOrderBy, sErrorDescription)
'************************************************************
'Purpose: To generate JavaScript code with the array elements
'		  for every entry in the recordset that will result
'		  by executing the query.
'Inputs:  oADODBConnection, sTableName, sIDField, sValueField, sCondition, sOrderBy
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GenerateJavaScriptArrayFromQuery"
	Dim sQuery
	Dim iIndex
	Dim oRecordset
	Dim asFieldName
	Dim asEmptyOption
	Dim sValue
	Dim lErrorNumber

	sQuery = "Select " & sIDField & ", " & sValueField & " From " & sTableName
	sCondition = Trim(sCondition)
	If Len(sCondition) > 0 Then
		If InStr(1, sCondition, "Where ", vbTextCompare) = 1 Then sCondition = Replace(sCondition, "Where ", "", 1, 1, vbTextCompare)
		If InStr(1, sCondition, "And ", vbTextCompare) = 1 Then sCondition = Replace(sCondition, "And ", "", 1, 1, vbTextCompare)
		sQuery = sQuery & " Where " & sCondition
	End If
	sQuery = sQuery & " Order By "
	If Len(sOrderBy) > 0 Then
		sQuery = sQuery & sOrderBy
	Else
		sQuery = sQuery & sValueField
	End If

	sErrorDescription = "No se pudo ejecutar la petición sobre la base de datos."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "DatabaseLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Do While Not oRecordset.EOF
				If Not IsNull(oRecordset.Fields(0).Value) Then
					Response.Write "['"
						For iIndex = 0 To oRecordset.Fields.Count - 2
							sValue = ""
							sValue = CStr(oRecordset.Fields(iIndex).Value)
							Err.Clear
							If InStr(1, oRecordset.Fields(iIndex).Name, "FormatNumber") = 0 Then
								Response.Write CleanStringForJavaScript(CStr(sValue)) & "', '"
							Else
								asFieldName = Split(oRecordset.Fields(iIndex).Name, "_", -1, vbBinaryCompare)
								Response.Write FormatNumber(CStr(sValue), asFieldName(1), True, False, True) & "', '"
							End If
						Next
						sValue = ""
						sValue = CStr(oRecordset.Fields(iIndex).Value)
						Err.Clear
						Response.Write CleanStringForJavaScript(CStr(sValue))
					Response.Write "'], "
					Response.Flush()
				End If
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	GenerateJavaScriptArrayFromQuery = lErrorNumber
	Err.Clear
End Function

Function GenerateListOptionsFromQuery(oADODBConnection, sTableName, sIDField, sValueField, sCondition, sOrderBy, sListOfSelectedItems, sEmptyOption, sErrorDescription)
'************************************************************
'Purpose: To generate HTML code with the <OPTION> tags for
'		  every entry in the recordset that will result by
'		  executing the query.
'Inputs:  oADODBConnection, sTableName, sIDField, sValueField, sCondition, sOrderBy, sListOfSelectedItems, sEmptyOption
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GenerateListOptionsFromQuery"
	Dim sQuery
	Dim oRecordset
	Dim asEmptyOption
	Dim iIndex
	Dim asID
	Dim asFieldName
	Dim asValue
	Dim sValue
	Dim sSeparator
	Dim lErrorNumber

	sQuery = "Select " & sIDField & ", " & sValueField & " From " & sTableName
	sCondition = Trim(sCondition)
	If Len(sCondition) > 0 Then
		If InStr(1, sCondition, "Where ", vbTextCompare) = 1 Then sCondition = Replace(sCondition, "Where ", "", 1, 1, vbTextCompare)
		If InStr(1, sCondition, "And ", vbTextCompare) = 1 Then sCondition = Replace(sCondition, "And ", "", 1, 1, vbTextCompare)
		sQuery = sQuery & " Where " & sCondition
	End If
	sQuery = sQuery & " Order By "
	If Len(sOrderBy) > 0 Then
		sQuery = sQuery & sOrderBy
	Else
		sQuery = sQuery & sValueField
	End If

	asID = Split(Replace(sIDField, "',' As", "'Þ' As"), ",", -1, vbBinaryCompare)
	asValue = Split(Replace(sValueField, "',' As", "'Þ' As"), ",", -1, vbBinaryCompare)
	If InStr(1, sListOfSelectedItems, LIST_SEPARATOR, vbBinaryCompare) = 0 Then
		sSeparator = ","
	Else
		sSeparator = LIST_SEPARATOR
	End If
	sErrorDescription = "No se pudo ejecutar la petición sobre la base de datos."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "DatabaseLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			If Len(sEmptyOption) > 0 Then
				asEmptyOption = Split(sEmptyOption, LIST_SEPARATOR, 2, vbBinaryCompare)
				GenerateListOptionsFromQuery = "<OPTION VALUE=""" & CleanStringForHTML(asEmptyOption(1)) & """>" & CleanStringForHTML(asEmptyOption(0)) & "</OPTION>"
			End If
		Else
			GenerateListOptionsFromQuery = ""
			Do While Not oRecordset.EOF
				If Not IsNull(oRecordset.Fields(0).Value) Then
					GenerateListOptionsFromQuery = GenerateListOptionsFromQuery & "<OPTION VALUE="""
						For iIndex = 0 To UBound(asID)
							If Not IsNull(oRecordset.Fields(iIndex)) Then
								GenerateListOptionsFromQuery = GenerateListOptionsFromQuery & CleanStringForHTML(CStr(oRecordset.Fields(iIndex).Value))
							End If
							If iIndex < UBound(asID) Then GenerateListOptionsFromQuery = GenerateListOptionsFromQuery & ","
						Next
					GenerateListOptionsFromQuery = GenerateListOptionsFromQuery & """"
						If (InStr(1, sSeparator & sListOfSelectedItems & sSeparator, sSeparator & CStr(oRecordset.Fields(0).Value) & sSeparator, vbTextCompare) > 0) Or (StrComp(sListOfSelectedItems, "True", vbBinaryCompare) = 0) Then
							GenerateListOptionsFromQuery = GenerateListOptionsFromQuery & " SELECTED=""1"""
						End If
					GenerateListOptionsFromQuery = GenerateListOptionsFromQuery & ">"
					For iIndex = 0 To UBound(asValue)
						sValue = ""
						sValue = CStr(oRecordset.Fields(UBound(asID) + 1 + iIndex).Value)
						Err.Clear
						If InStr(1, oRecordset.Fields(UBound(asID) + 1 + iIndex).Name, "FormatNumber") = 0 Then
							If B_UPPERCASE Then
								GenerateListOptionsFromQuery = GenerateListOptionsFromQuery & CleanStringForHTML(UCase(CStr(sValue)))
							Else
								GenerateListOptionsFromQuery = GenerateListOptionsFromQuery & CleanStringForHTML(CStr(sValue))
							End If
						Else
							asFieldName = Split(oRecordset.Fields(UBound(asID) + 1 + iIndex).Name, "_", -1, vbBinaryCompare)
							GenerateListOptionsFromQuery = GenerateListOptionsFromQuery & FormatNumber(CStr(sValue), asFieldName(1), True, False, True)
						End If
						If iIndex < UBound(asValue) Then GenerateListOptionsFromQuery = GenerateListOptionsFromQuery & " "
					Next
					GenerateListOptionsFromQuery = GenerateListOptionsFromQuery & "</OPTION>"
					Response.Flush()
				End If
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function GenerateHierarchyListOptionsFromQuery(oADODBConnection, sTableName, sIDField, sParentIDField, sValueField, sCondition, lParentID, sOrderBy, sListOfSelectedItems, sTab, sEmptyOption, sOutput, sErrorDescription)
'************************************************************
'Purpose: To generate HTML code with the <OPTION> tags for
'		  every entry in the recordset that will result by
'		  executing the query following their hierarchy.
'Inputs:  oADODBConnection, sTableName, sIDField, sParentIDField, sValueField, sCondition, lParentID, sOrderBy, sListOfSelectedItems, sTab, sEmptyOption
'Outputs: sOutput, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GenerateHierarchyListOptionsFromQuery"
	Dim sQuery
	Dim oRecordset
	Dim asEmptyOption
	Dim iIndex
	Dim asID
	Dim asFieldName
	Dim asValue
	Dim sSeparator
	Dim lErrorNumber

	sQuery = "Select " & sIDField & ", " & sParentIDField & ", " & sValueField & " From " & sTableName
	sQuery = sQuery & " Where (" & sParentIDField & "=" & lParentID & ") "
	sCondition = Trim(sCondition)
	If Len(sCondition) > 0 Then
		If InStr(1, sCondition, "Where ", vbTextCompare) = 1 Then sCondition = Replace(sCondition, "Where ", "", 1, 1, vbTextCompare)
		If InStr(1, sCondition, "And ", vbTextCompare) <> 1 Then sCondition = "And " & sCondition
		sQuery = sQuery & sCondition
	End If
	sQuery = sQuery & " Order By "
	If Len(sOrderBy) > 0 Then
		sQuery = sQuery & sOrderBy
	Else
		sQuery = sQuery & sValueField
	End If

	asID = Split(sIDField, ",", -1, vbBinaryCompare)
	asValue = Split(sValueField, ",", -1, vbBinaryCompare)
	If InStr(1, sListOfSelectedItems, LIST_SEPARATOR, vbBinaryCompare) = 0 Then
		sSeparator = ","
	Else
		sSeparator = LIST_SEPARATOR
	End If
	sErrorDescription = "No se pudo ejecutar la petición sobre la base de datos."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "DatabaseLibrary.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			If Len(sEmptyOption) > 0 Then
				asEmptyOption = Split(sEmptyOption, LIST_SEPARATOR, 2, vbBinaryCompare)
				sOutput = sOutput & "<OPTION VALUE=""" & CleanStringForHTML(asEmptyOption(1)) & """>" & CleanStringForHTML(asEmptyOption(0)) & "</OPTION>"
			End If
		Else
			Do While Not oRecordset.EOF
				If Not IsNull(oRecordset.Fields(0).Value) Then
				If Not IsNull(oRecordset.Fields(2).Value) Then
					sOutput = sOutput & "<OPTION VALUE="""
						For iIndex = 0 To UBound(asID)
							If Not IsNull(oRecordset.Fields(iIndex)) Then
								sOutput = sOutput & CleanStringForHTML(CStr(oRecordset.Fields(iIndex).Value))
							End If
							If iIndex < UBound(asID) Then sOutput = sOutput & ","
						Next
					sOutput = sOutput & """"
						If (InStr(1, sSeparator & sListOfSelectedItems & sSeparator, sSeparator & CStr(oRecordset.Fields(0).Value) & sSeparator, vbTextCompare) > 0) Or (StrComp(sListOfSelectedItems, "True", vbBinaryCompare) = 0) Then
							sOutput = sOutput & " SELECTED=""1"""
						End If
					sOutput = sOutput & ">" & sTab
					For iIndex = 0 To UBound(asValue)
						If Not IsNull(oRecordset.Fields(UBound(asID) + 2 + iIndex)) Then
							If InStr(1, oRecordset.Fields(UBound(asID) + 2 + iIndex).Name, "FormatNumber") = 0 Then
								If B_UPPERCASE Then
									sOutput = sOutput & CleanStringForHTML(UCase(CStr(oRecordset.Fields(UBound(asID) + 2 + iIndex).Value)))
								Else
									sOutput = sOutput & CleanStringForHTML(CStr(oRecordset.Fields(UBound(asID) + 2 + iIndex).Value))
								End If
							Else
								asFieldName = Split(oRecordset.Fields(UBound(asID) + 2 + iIndex).Name, "_", -1, vbBinaryCompare)
								sOutput = sOutput & FormatNumber(CStr(oRecordset.Fields(UBound(asID) + 2 + iIndex).Value), asFieldName(1), True, False, True)
							End If
						End If
						If iIndex < UBound(asValue) Then sOutput = sOutput & " "
					Next
					sOutput = sOutput & "</OPTION>"
					lErrorNumber = GenerateHierarchyListOptionsFromQuery(oADODBConnection, sTableName, sIDField, sParentIDField, sValueField, sCondition, CLng(oRecordset.Fields(0).Value), sOrderBy, sListOfSelectedItems, sTab & "&nbsp;&nbsp;&nbsp;", sEmptyOption, sOutput, sErrorDescription)
					Response.Flush()
				End If
				End If
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	GenerateHierarchyListOptionsFromQuery = lErrorNumber
	Err.Clear
End Function

Function GetNumberOfEntriesFromRecordset(oRecordset)
'************************************************************
'Purpose: To count the number of entries in a recordset
'Inputs:  oRecordset
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetNumberOfEntriesFromRecordset"
	Dim lEntryCounter

	lEntryCounter = 0
	oRecordset.MoveFirst
	Do While Not oRecordset.EOF
		oRecordset.MoveNext
		lEntryCounter = lEntryCounter + 1
		If Err.number <> 0 Then
			lEntryCounter = 0
			Exit Do
		End If
	Loop
	oRecordset.MoveFirst

	GetNumberOfEntriesFromRecordset = lEntryCounter
	Err.Clear
End Function

Function MoveRecordsetToLastItem(oRecordset)
'************************************************************
'Purpose: To move a recordset to its last item
'Inputs:  oRecordset
'Outputs: oRecordset
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "MoveRecordsetToLastItem"

	If Not oRecordset.EOF Then
		Do While Not oRecordset.EOF
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
	End If
	oRecordset.MovePrevious

	MoveRecordsetToLastItem = Err.number
	Err.Clear
End Function
%>