<%
Function ZipFile(sFilePath, sZIPFile, sErrorDescription)
'************************************************************
'Purpose: To create a zip file withg the given file contents
'Inputs:  sFilePath, sZIPFile
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ZipFile"
	Dim oZIPFile
	Dim lErrorNumber

	Set oZIPFile = Server.CreateObject("XStandard.Zip")
	lErrorNumber = Err.Number
	If lErrorNumber <> 0 Then
		sErrorDescription = "El objeto XStandard.Zip no pudo ser creado. El archivo XZip.dll puede no estar registrado en el servidor Web. Favor de contactar al Administrador."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web:&nbsp;</B>" & Err.description
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ZIPLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	Else
		Call oZIPFile.Pack(sFilePath, sZIPFile)
		lErrorNumber = Err.Number
		If lErrorNumber <> 0 Then
			sErrorDescription = "No se pudo comprimir el archivo '" & sFilePath & "'."
			If Len(Err.description) > 0 Then
				sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web:&nbsp;</B>" & Err.description
			End If
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ZIPLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
		End If
	End If

	Set oZIPFile = Nothing
	ZipFile = lErrorNumber
	Err.Clear
End Function

Function ZipFolder(sFolderPath, sZIPFile, sErrorDescription)
'************************************************************
'Purpose: To create a zip file with the given folder contents
'Inputs:  sFolderPath, sZIPFile
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ZipFolder"
	Dim oZIPFile
	Dim lErrorNumber

	Set oZIPFile = Server.CreateObject("XStandard.Zip")
	lErrorNumber = Err.Number
	If lErrorNumber <> 0 Then
		sErrorDescription = "El objeto XStandard.Zip no pudo ser creado. El archivo XZip.dll puede no estar registrado en el servidor Web. Favor de contactar al Administrador."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web:&nbsp;</B>" & Err.description
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ZIPLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	Else
		Call oZIPFile.Pack(sFolderPath, sZIPFile)
		lErrorNumber = Err.Number
		If lErrorNumber <> 0 Then
			sErrorDescription = "No se pudo comprimir el directorio '" & sFolderPath & "'."
			If Len(Err.description) > 0 Then
				sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web:&nbsp;</B>" & Err.description
			End If
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ZIPLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
		End If
	End If

	Set oZIPFile = Nothing
	ZipFolder = lErrorNumber
	Err.Clear
End Function
%>