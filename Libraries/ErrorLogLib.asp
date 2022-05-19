<%
Private Const S_DATE_ERROR_LOG = 0
Private Const N_SHOW_LEVEL_ERROR_LOG = 1
Private Const S_SHOW_ACCESS_KEY_ERROR_LOG = 2
Private Const S_SORT_ERROR_LOG = 3
Private Const S_FOLDER_LOG = 4
Private Const S_ERROR_FILE_LOG = 5
Private Const B_COMPONENT_INITIALIZED_ERROR_LOG = 6

Private Const N_ERROR_LOG_COMPONENT_SIZE = 6

Const S_ERROR_LOG_FILE_NAME = "Logs\LogXXX.txt"

Dim aErrorLogComponent()
ReDim aErrorLogComponent(N_ERROR_LOG_COMPONENT_SIZE)

Dim aMessagesTypes
aMessagesTypes = Split("Mensajes,Advertencias,Errores,Errores de SQL,Operaciones en la base de datos,Mensajes enviados", ",", -1, vbBinaryCompare)

Call InitializeErrorLogComponent(oRequest, aErrorLogComponent)

Function InitializeErrorLogComponent(oRequest, aErrorLogComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Error Log
'         Component using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aErrorLogComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeErrorLogComponent"
	Dim iItem
	Redim Preserve aErrorLogComponent(N_ERROR_LOG_COMPONENT_SIZE)

	If IsEmpty(aErrorLogComponent(S_DATE_ERROR_LOG)) Then
		If Len(oRequest("ErrorLogDate").Item) > 0 Then
			aErrorLogComponent(S_DATE_ERROR_LOG) = oRequest("ErrorLogDate").Item
		ElseIf (Len(oRequest("Year").Item) > 0) And (Len(oRequest("Month").Item) > 0) And (Len(oRequest("Day").Item) > 0) Then
			aErrorLogComponent(S_DATE_ERROR_LOG) = oRequest("Year").Item & Right(("0" & oRequest("Month").Item), Len("00")) &  Right(("0" & oRequest("Day").Item), Len("00"))
		Else
			aErrorLogComponent(S_DATE_ERROR_LOG) = Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00"))
		End If
	End If

	If IsEmpty(aErrorLogComponent(N_SHOW_LEVEL_ERROR_LOG)) Then
		If Len(oRequest("ErrorLogLevel").Item) > 0 Then
			If InStr(1, oRequest("ErrorLogLevel").Item, ",", vbBinaryCompare) > 1 Then
				aErrorLogComponent(N_SHOW_LEVEL_ERROR_LOG) = 0
				For Each iItem In oRequest("ErrorLogLevel")
					aErrorLogComponent(N_SHOW_LEVEL_ERROR_LOG) = aErrorLogComponent(N_SHOW_LEVEL_ERROR_LOG) + CLng(iItem)
				Next
			Else
				aErrorLogComponent(N_SHOW_LEVEL_ERROR_LOG) = CLng(oRequest("ErrorLogLevel").Item)
			End If
		Else
			aErrorLogComponent(N_SHOW_LEVEL_ERROR_LOG) = N_MESSAGE_LEVEL Or N_WARNING_LEVEL Or N_ERROR_LEVEL Or N_SQL_ERROR_LEVEL Or N_EMAIL_LEVEL
		End If
	End If

	If IsEmpty(aErrorLogComponent(S_SHOW_ACCESS_KEY_ERROR_LOG)) Then
		If Len(oRequest("ErrorLogAccessKey").Item) > 0 Then
			aErrorLogComponent(S_SHOW_ACCESS_KEY_ERROR_LOG) = oRequest("ErrorLogAccessKey").Item
		Else
			aErrorLogComponent(S_SHOW_ACCESS_KEY_ERROR_LOG) = ""
		End If
	End If

	If IsEmpty(aErrorLogComponent(S_SORT_ERROR_LOG)) Then
		If Len(oRequest("ErrorLogSort").Item) > 0 Then
			aErrorLogComponent(S_SORT_ERROR_LOG) = oRequest("ErrorLogSort").Item
		Else
			aErrorLogComponent(S_SORT_ERROR_LOG) = ""
		End If
	End If

	If IsEmpty(aErrorLogComponent(S_FOLDER_LOG)) Then
		If Len(Request.Cookies("SIAP_LogFolder")) > 0 Then
			aErrorLogComponent(S_FOLDER_LOG) = Request.Cookies("SIAP_LogFolder").Item
		ElseIf Len(oRequest("LogFolder").Item) > 0 Then
			aErrorLogComponent(S_FOLDER_LOG) = oRequest("LogFolder").Item
		End If
	End If
	Response.Cookies("SIAP_LogFolder") = aErrorLogComponent(S_FOLDER_LOG)
	aErrorLogComponent(S_ERROR_FILE_LOG) = Replace(Replace(S_ERROR_LOG_FILE_NAME, "\", ("\" & aErrorLogComponent(S_FOLDER_LOG) & "\"), 1, 1, vbBinaryCompare), "\\", "\", 1, -1, vbBinaryCompare)

	aErrorLogComponent(B_COMPONENT_INITIALIZED_ERROR_LOG) = True
	InitializeErrorLogComponent = Err.number
	Err.Clear
End Function

Function LogErrorInXMLFile(lErrorNumber, sErrorDescription, iDescriptorID, sLibraryFile, sFunctionName, iLevel)
'************************************************************
'Purpose: To append a new node into the Error Log XML File
'Inputs:  lErrorNumber, sErrorDescription, iDescriptorID, sLibraryFile, sFunctionName, iLevel
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "LogErrorInXMLFile"
	Dim sAccessKey
	Dim sXMLNode
	Dim sDate
	Dim sURL
	Dim aParameters
	Dim iIndex
	Dim oFileSystemObject
	Dim oTextFile
	
	sAccessKey = ""
	sAccessKey = aLoginComponent(S_ACCESS_KEY_LOGIN)
	If Len(sAccessKey) = 0 Then
		sAccessKey = Request.Cookies("SIAP_CurrentAccessKey").Item
	End If
	Err.Clear
	sDate = GetSerialNumberForDate("")
	sURL = CStr(oRequest)
	aParameters = Split("UserPassword,ReportContents", ",", -1, vbBinaryCompare)
	For iIndex = 0 To UBound(aParameters)
		sURL = RemoveParameterFromURLString(sURL, aParameters(iIndex))
	Next
	sURL = Server.HTMLEncode(sURL)
	If (lErrorNumber = 0) And ((iLevel = N_SQL_QUERY_LEVEL) Or (iLevel = N_EMAIL_LEVEL)) Then
		sURL = ""
	End If
	sXMLNode = "<ERROR IPAddress=""" & Request.ServerVariables("REMOTE_ADDR") & """ UserAccessKey=""" & sAccessKey & """ ErrorNumber=""" & lErrorNumber & """ ErrorDescription=""" & Server.HTMLEncode(sErrorDescription) & """ DescriptorID=""" & iDescriptorID & """ ASPFile=""" & Request.ServerVariables("PATH_INFO") & """ LibraryFile=""" & sLibraryFile & """ FunctionName=""" & sFunctionName & """ URL=""" & sURL & """ Date=""" & sDate & """ Level=""" & iLevel & """ />"

	Set oFileSystemObject = Server.CreateObject("Scripting.FileSystemObject")
	If Err.number <> 0 Then
		lErrorNumber = Err.number
		sErrorDescription = "El archivo 'scrrun.dll' no se encuentra registrado correctamente en el servidor Web. No se pudo escribir en la bitácora. Favor de contactar al Administrador."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
		End If
	Else
		Set oTextFile = oFileSystemObject.OpenTextFile(SYSTEM_PHYSICAL_PATH & (Replace(aErrorLogComponent(S_ERROR_FILE_LOG), "XXX", Left(sDate, Len("YYYYMMDD")), 1, 1, vbBinaryCompare)), 8, True)
		If Err.Number = 0 Then
			oTextFile.WriteLine sXMLNode
			oTextFile.Close
			Set oTextFile = Nothing
		End If
		Set oFileSystemObject = Nothing
	End If

	LogErrorInXMLFile = Err.number
	Err.Clear
End Function

Function TransformDateFromErrorLog(sCodedDate, bDisplayDate, bDisplayTime)
'************************************************************
'Purpose: To transform the coded date into a readable date
'Inputs:  sCodedDate, bDisplayDate, bDisplayTime
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "TransformDateFromErrorLog"

	If bDisplayDate And bDisplayTime Then
		TransformDateFromErrorLog = DisplayDate(CInt(Mid(sCodedDate, 1, 4)), CInt(Mid(sCodedDate, 5, 2)), CInt(Mid(sCodedDate, 7, 2)), CInt(Mid(sCodedDate, 9, 2)), CInt(Mid(sCodedDate, 11, 2)), CInt(Mid(sCodedDate, 13, 2)))
	ElseIf bDisplayDate Then
		TransformDateFromErrorLog = DisplayDate(CInt(Mid(sCodedDate, 1, 4)), CInt(Mid(sCodedDate, 5, 2)), CInt(Mid(sCodedDate, 7, 2)), -1, -1, -1)
	ElseIf bDisplayTime Then
		TransformDateFromErrorLog = Mid(sCodedDate, 9, 2) & ":" & Mid(sCodedDate, 11, 2) & ":" & Mid(sCodedDate, 13, 2)
	End If

	Err.Clear
End Function

Function GetLogFileContents(aErrorLogComponent, sContents, sErrorDescription)
'************************************************************
'Purpose: To display a table with the errors logged in the
'         given date
'Inputs:  aErrorLogComponent
'Outputs: aErrorLogComponent, sContents, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetLogFileContents"
    Dim oFileSystemObject
    Dim oFile
    Dim oTextFile
	Dim sFileName
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aErrorLogComponent(B_COMPONENT_INITIALIZED_ERROR_LOG)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeErrorLogComponent(oRequest, aErrorLogComponent)
	End If

	sContents = ""
	Set oFileSystemObject = Server.CreateObject("Scripting.FileSystemObject")
	If Err.number <> 0 Then
		lErrorNumber = Err.number
		sErrorDescription = "El archivo 'scrrun.dll' no se encuentra registrado correctamente en el servidor Web. No se pudo obtener el contenido de la bitácora. Favor de contactar al Administrador."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ErrorLogLib.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	Else
		sFileName = Server.MapPath(Replace(aErrorLogComponent(S_ERROR_FILE_LOG), "XXX", aErrorLogComponent(S_DATE_ERROR_LOG), 1, 1, vbBinaryCompare))
	    Set oFile = oFileSystemObject.GetFile(sFileName)
	    If Err.Number = 0 Then Set oTextFile = oFile.OpenAsTextStream(1)
		If Err.Number = 53 Then
			lErrorNumber = Err.number
			Err.clear
			sErrorDescription = "No existe bitácora para el día " & DisplayDate(CInt(Mid(aErrorLogComponent(S_DATE_ERROR_LOG), 1, 4)), CInt(Mid(aErrorLogComponent(S_DATE_ERROR_LOG), 5, 2)), CInt(Mid(aErrorLogComponent(S_DATE_ERROR_LOG), 7, 2)), -1, -1, -1)
			If Err.number <> 0 Then
				lErrorNumber = Err.number
				sErrorDescription = "El archivo '" & sFileName & "' no existe."
				If Len(Err.description) > 0 Then
					sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
				End If
			End If
		ElseIf Err.Number <> 0 Then
			lErrorNumber = Err.number
			sErrorDescription = "Hubo un problema al abrir el archivo '" & sFileName & "'."
			If Len(Err.description) > 0 Then
				sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
			End If
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ErrorLogLib.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
		Else
			sContents = "<ERRORS>" & oTextFile.ReadAll() & "</ERRORS>"
			If Err.Number <> 0 Then
				lErrorNumber = Err.number
				sErrorDescription = "Hubo un problema al leer el contenido del archivo '" & sFileName & "'."
				If Len(Err.description) > 0 Then
					sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
				End If
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ErrorLogLib.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
			Else
				oTextFile.Close
				Set oTextFile = Nothing
				Set oFile = Nothing
			End If
		End If
		Set oFileSystemObject = Nothing
	End If

	GetLogFileContents = lErrorNumber
	Err.Clear
End Function

Function GetLogFilesNames(aErrorLogComponent, sFilter, sFilesNames, sErrorDescription)
'************************************************************
'Purpose: To get the name of the files using the specified
'         filter
'Inputs:  aErrorLogComponent, sFilter
'Outputs: sFilesNames, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetLogFilesNames"
	Dim oFileSystemObject
	Dim oFolderObject
	Dim oItemInFolder
	Dim sTempItemName
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aErrorLogComponent(B_COMPONENT_INITIALIZED_ERROR_LOG)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeErrorLogComponent(oRequest, aErrorLogComponent)
	End If
	
	Set oFileSystemObject = Server.CreateObject("Scripting.FileSystemObject")
	lErrorNumber = Err.number
	If lErrorNumber <> 0 Then
		sErrorDescription = "El archivo 'scrrun.dll' no se encuentra registrado correctamente en el servidor Web. No se pudieron obtener los archivos de la bitácora. Favor de contactar al Administrador."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ErrorLogLib.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	Else
		Set oFolderObject = oFileSystemObject.GetFolder(Server.MapPath(Replace(aErrorLogComponent(S_ERROR_FILE_LOG), "\LogXXX.txt", "", 1, 1, vbBinaryCompare)))
		lErrorNumber = Err.number
		If lErrorNumber = 53 Then
			sErrorDescription = "El directorio '" & Server.MapPath(Replace(aErrorLogComponent(S_ERROR_FILE_LOG), "\LogXXX.txt", "", 1, 1, vbBinaryCompare)) & "' no existe."
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ErrorLogLib.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
		ElseIf lErrorNumber <> 0 Then
			sErrorDescription = "El directorio '" & Server.MapPath(Replace(aErrorLogComponent(S_ERROR_FILE_LOG), "\LogXXX.txt", "", 1, 1, vbBinaryCompare)) & "' no pudo ser abierto."
			If Len(Err.description) > 0 Then
				sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
			End If
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ErrorLogLib.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
		Else
			sFilesNames = ""
			For Each oItemInFolder In oFolderObject.Files
				sTempItemName = oItemInFolder.Name
				If Len(sTempItemName) > 0 Then
					If InStr(1, sTempItemName, sFilter, vbTextCompare) > 0 Then
						sFilesNames = sFilesNames & sTempItemName & ","
					End If
				End If
			Next
			sFilesNames = Left(sFilesNames, (Len(sFilesNames) - Len(",")))
		End If
	End If

	GetLogFilesNames = lErrorNumber
	Err.Clear
End Function

Function DisplayLogFile(aErrorLogComponent, sErrorDescription)
'************************************************************
'Purpose: To display a table with the errors logged in the
'         given date
'Inputs:  aErrorLogComponent
'Outputs: aErrorLogComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayLogFile"
	Dim oXML
	Dim oXMLNode
	Dim sContents
	Dim bEmpty
	Dim sLevelFilter
	Dim iIndex
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	
	bEmpty = True
	lErrorNumber = GetLogFileContents(aErrorLogComponent, sContents, sErrorDescription)
	If lErrorNumber = 0 Then
		lErrorNumber = CreateXMLDOMObject(oXML, sErrorDescription)
		If lErrorNumber = 0 Then
			lErrorNumber = LoadXMLToObject(sContents, oXML, sErrorDescription)
			If lErrorNumber = 0 Then
				Response.Write "<TABLE WIDTH=""700"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
					asColumnsTitles = Split("&nbsp;,Hora,Usuario,Mensaje o Error,Origen", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,60,150,270,200", ",", -1, vbBinaryCompare)
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If

					asCellAlignments = Split(",,,,", ",", -1, vbBinaryCompare)
					sLevelFilter = ""
					If aErrorLogComponent(N_SHOW_LEVEL_ERROR_LOG) <> - 1 Then
						For iIndex = 0 To 10
							If (aErrorLogComponent(N_SHOW_LEVEL_ERROR_LOG) And (2 ^ iIndex)) <> 0 Then
								If Len(sLevelFilter) > 0 Then sLevelFilter = sLevelFilter & " $or$ "
								sLevelFilter = sLevelFilter & "(@Level = '" & 2 ^ iIndex & "')"
							End If
						Next
					End If
					If Len(sLevelFilter) > 0 Then sLevelFilter = "[" & sLevelFilter & "]"
					For Each oXMLNode In oXML.documentElement.selectNodes("/ERRORS/ERROR" & sLevelFilter)
						bEmpty = False
						sRowContents = ""
						sRowContents = sRowContents  & "<IMG SRC=""Images/IcnErrorLevel" & oXMLNode.getAttribute("Level") & ".gif"" WIDTH=""16"" HEIGHT=""16"" />"
						sRowContents = sRowContents & TABLE_SEPARATOR & TransformDateFromErrorLog(oXMLNode.getAttribute("Date"), False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR & "<B>Clave&nbsp;del&nbsp;usuario:&nbsp;</B>" & oXMLNode.getAttribute("UserAccessKey") & "<BR />"
						sRowContents = sRowContents & "<B>Dirección&nbsp;IP:&nbsp;</B>" & oXMLNode.getAttribute("IPAddress")
						sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & oXMLNode.getAttribute("ErrorNumber") & ":&nbsp;</B>"
						sRowContents = sRowContents & oXMLNode.getAttribute("ErrorDescription") & "<BR /><B>URL:&nbsp;</B>" & Replace(oXMLNode.getAttribute("URL"), "&", "&<BR />", 1, -1, vbBinaryCompare)
						sRowContents = sRowContents & TABLE_SEPARATOR & "<B>ASP:&nbsp;</B>" & oXMLNode.getAttribute("ASPFile") & "<BR />"
						sRowContents = sRowContents & "<B>Librería:&nbsp;</B>" & oXMLNode.getAttribute("LibraryFile") & "<BR />"
						sRowContents = sRowContents & "<B>Función:&nbsp;</B>" & oXMLNode.getAttribute("FunctionName")
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						If Err.number <> 0 Then Exit For
					Next
				Response.Write "</TABLE>" & vbNewLine
				If bEmpty Then
					lErrorNumber = L_ERR_NO_RECORDS
					sErrorDescription = "No se han registrado eventos del(os) tipo(s) de mensaje seleccionado(s) en la bitácora para el día " & Mid(aErrorLogComponent(S_DATE_ERROR_LOG), Len("YYYYMMD"), Len("DD")) & " de " & asMonthNames_es(Mid(aErrorLogComponent(S_DATE_ERROR_LOG), Len("YYYYM"), Len("MM"))) & " de " & Left(aErrorLogComponent(S_DATE_ERROR_LOG), Len("YYYY")) & "."
				End If
			End If
		End If
	End If

	DisplayLogFile = lErrorNumber
	Err.Clear
End Function

Function SendLogFileViaEmail()
'************************************************************
'Purpose: To send today's log file via e-mail
'************************************************************
    On Error Resume Next
	Const S_FUNCTION_NAME = "SendLogFileViaEmail"
	Dim oDate
	Dim sDate
	Dim sFileName
	Dim sContents
	Dim oFileSystemObject
    Dim oTextFile
    Dim oFile
	Dim lErrorNumber

	If B_USE_SMTP And B_SEND_LOG_FILES Then
		If (CLng(Application.Contents("SIAP_DaemonStatus")) And N_SEND_LOGS_DAEMON) Then 
		Else
			oDate = DateAdd("d", -1, Now())
			sDate = Year(oDate) & Right(("0" & Month(oDate)), Len("00")) & Right(("0" & Day(oDate)), Len("00")) & Right(("0" & Hour(oDate)), Len("00")) & Right(("0" & Minute(oDate)), Len("00")) & Right(("0" & Second(oDate)), Len("00"))
			sFileName = Server.MapPath(Replace(S_ERROR_LOG_FILE_NAME, "XXX", Left(sDate, Len("YYYYMMDD")), 1, 1, vbBinaryCompare))
			sContents = ""
			Set oFileSystemObject = Server.CreateObject("Scripting.FileSystemObject")
			lErrorNumber = Err.number
			If lErrorNumber <> 0 Then
				sErrorDescription = "El archivo 'scrrun.dll' no se encuentra registrado correctamente en el servidor Web. No se pudo enviar el archivo de la bitácora por correo electrónico. Favor de contactar al Administrador."
				If Len(Err.description) > 0 Then
					sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
				End If
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ErrorLogLib.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
			Else
			    Set oFile = oFileSystemObject.GetFile(sFileName)
			    lErrorNumber = Err.Number
			    If lErrorNumber = 0 Then Set oTextFile = oFile.OpenAsTextStream(1)
			    lErrorNumber = Err.Number
				ReDim aEmailComponent(N_EMAIL_COMPONENT_SIZE)
				aEmailComponent(S_TO_EMAIL) = "victor@jibda.com"
				aEmailComponent(S_FROM_EMAIL) = S_ADMIN_EMAIL_ACCOUNT
				aEmailComponent(S_SUBJECT_EMAIL) = "Bitácora de errores proveniente de " & SERVER_NAME_FOR_LICENSE & " (" & SERVER_IP_FOR_LICENSE & ")"
				If lErrorNumber = 53 Then
					aEmailComponent(S_BODY_EMAIL) = "<FONT FACE=""Arial"" SIZE=""2"">Este mensaje no contiene la bitácora de errores de " & SYSTEM_PATH & " pues no se generó ningún registro.<BR /><BR />" & _
													"El mensaje proviene del servidor <B>" & Request.ServerVariables("SERVER_NAME") & "</B> con dirección IP <B>" & Request.ServerVariables("LOCAL_ADDR") & "</B>.<BR /><BR /></FONT>"
					lErrorNumber = SendEmail(oRequest, aEmailComponent, sErrorDescription)
					If Err.Number = 0 Then
						Application.Contents("SIAP_DaemonStatus") = Application.Contents("SIAP_DaemonStatus") Or N_SEND_LOGS_DAEMON
					End If
				ElseIf lErrorNumber <> 0 Then
					sErrorDescription = "Hubo un problema al abrir el archivo '" & sFileName & "'."
					If Len(Err.description) > 0 Then
						sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
					End If
					Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ErrorLogLib.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
				Else
					sContents = oTextFile.ReadAll()
					oTextFile.Close
					Set oTextFile = Nothing
					If InStr(1, sContents, "<LOG_STATUS", vbBinaryCompare) = 0 Then
						aEmailComponent(S_BODY_EMAIL) = "<FONT FACE=""Arial"" SIZE=""2"">Este mensaje contiene la bitácora de errores de " & SYSTEM_PATH & ".<BR /><BR />" & _
														"El archivo adjunto proviene del servidor <B>" & Request.ServerVariables("SERVER_NAME") & "</B> con dirección IP <B>" & Request.ServerVariables("LOCAL_ADDR") & "</B>.<BR /><BR /></FONT>"
						aEmailComponent(S_ATTACHMENTS_EMAIL) = sFileName
						lErrorNumber = SendEmail(oRequest, aEmailComponent, sErrorDescription)
						If lErrorNumber = 0 Then
							Set oTextFile = oFileSystemObject.OpenTextFile(sFileName, 8, True)
							If Err.Number = 0 Then
								oTextFile.WriteLine "<LOG_STATUS SendDate=""" & sDate & """ />"
								If Err.Number = 0 Then
									Application.Contents("SIAP_DaemonStatus") = Application.Contents("SIAP_DaemonStatus") Or N_SEND_LOGS_DAEMON
								End If
							End If
							oTextFile.Close
							Set oTextFile = Nothing
						End If
					End If
					Set oFile = Nothing
				End If
				Set oFileSystemObject = Nothing
			End If
		End If
	End If

	SendLogFileViaEmail = lErrorNumber
	Err.Clear
End Function
%>