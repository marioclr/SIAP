<%
Function AppendTextToFile(sFilePath, sContents, sErrorDescription)
'************************************************************
'Purpose: To save the given text into the specified file
'Inputs:  sFilePath, sContents
'Outputs: sErrorDescription
'************************************************************
    On Error Resume Next
	Const S_FUNCTION_NAME = "AppendTextToFile"
    Dim oFileSystem
    Dim oTextFile
    Dim sFolderPath
    Dim lErrorNumber

    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
	lErrorNumber = Err.number
	If lErrorNumber <> 0 Then
		sErrorDescription = "No se pudo crear una instancia del objeto 'Scripting.FileSystemObject' porque el archivo 'scrrun.dll' no se encuentra correctamente registrado en el servidor de Web. Favor de contactar al Administrador."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	Else
		sFolderPath = Left(sFilePath, (InStrRev(sFilePath, "\") - Len("\")))
		If Not FolderExists(sFolderPath, sErrorDescription) Then Call CreateFolder(sFolderPath, "")
		Set oTextFile = oFileSystem.OpenTextFile(sFilePath, 8, True)
		lErrorNumber = Err.number
		If lErrorNumber <> 0 Then
			sErrorDescription = "No se pudo abrir el archivo '" & sFilePath & "'. Favor de contactar al Administrador."
			If Len(Err.description) > 0 Then
				sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
			End If
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
		Else
			oTextFile.WriteLine sContents
			lErrorNumber = Err.number
			If lErrorNumber <> 0 Then
				sErrorDescription = "No se pudieron escribir los contenidos especificados al archivo '" & sFilePath & "'. Favor de contactar al Administrador."
				If Len(Err.description) > 0 Then
					sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
				End If
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
			Else
			    oTextFile.Close
			End If
		End If
	End If

    Set oTextFile = Nothing
    Set oFileSystem = Nothing
    AppendTextToFile = lErrorNumber
	Err.Clear
End Function 'End of AppendTextToFile

Function CopyFile(sFilePath, sTargetFilePath, sErrorDescription)
'************************************************************
'Purpose: To delete the specified file
'Inputs:  sFilePath, sTargetFilePath
'Outputs: sErrorDescription
'************************************************************
    On Error Resume Next
	Const S_FUNCTION_NAME = "CopyFile"
    Dim oFileSystem
    Dim oFile
    Dim sTargetFolderPath
    Dim lErrorNumber

    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
    lErrorNumber = Err.number
    If lErrorNumber <> 0 Then
		sErrorDescription = "No se pudo crear una instancia del objeto 'Scripting.FileSystemObject'. El archivo 'scrrun.dll' no está correctamente registrado en el servidor de Web. Favor de contactar al Administrador."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web:&nbsp;</B>" & Err.description
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
    Else
		Set oFile = oFileSystem.GetFile(sFilePath)
		lErrorNumber = Err.number
		If lErrorNumber <> 0 Then
			sErrorDescription = "No se pudo encontrar el archivo '" & sFilePath & "'. Favor de contactar al Administrador."
			If Len(Err.description) > 0 Then
				sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web:&nbsp;</B>" & Err.description
			End If
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
		Else
			sTargetFolderPath = Left(sTargetFilePath, (InStrRev(sTargetFilePath, "\") - Len("\")))
			If Not FolderExists(sTargetFolderPath, sErrorDescription) Then Call CreateFolder(sTargetFolderPath, "")
			Call oFile.Copy(sTargetFilePath, True)
			lErrorNumber = Err.number
			If lErrorNumber <> 0 Then
				sErrorDescription = "No se pudo copiar el archivo '" & sSourceFilePath & "' a '" & sTargetFilePath & "'. Favor de contactar al Administrador."
				If Len(Err.description) > 0 Then
					sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web:&nbsp;</B>" & Err.description
				End If
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
			End If
		End If
    End If

	Set oFileSystem = Nothing
    CopyFile = lErrorNumber
    Err.Clear
End Function 'End of CopyFile

Function CopyFiles(sFilePath, sTargetFilePath, sErrorDescription)
'************************************************************
'Purpose: To delete the specified file
'Inputs:  sFilePath, sTargetFilePath
'Outputs: sErrorDescription
'************************************************************
    On Error Resume Next
	Const S_FUNCTION_NAME = "CopyFiles"
    Dim oFileSystem
    Dim oFile
    Dim sTargetFolderPath
    Dim lErrorNumber

    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
    lErrorNumber = Err.number
    If lErrorNumber <> 0 Then
		sErrorDescription = "No se pudo crear una instancia del objeto 'Scripting.FileSystemObject'. El archivo 'scrrun.dll' no está correctamente registrado en el servidor de Web. Favor de contactar al Administrador."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web:&nbsp;</B>" & Err.description
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
    Else
        'Response.Write "Copiar 1" & Err.description & " " & Err.number & "<BR />"
		Call oFileSystem.CopyFile(sFilePath, sTargetFilePath)
        'Response.Write "Copiar 2" & Err.description & " " & Err.number & "<BR />"
		lErrorNumber = Err.number
		If lErrorNumber <> 0 Then
			sErrorDescription = "No se pudo copiar el archivo '" & sFilePath & "'. Favor de contactar al Administrador."
			If Len(Err.description) > 0 Then
				sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web:&nbsp;</B>" & Err.description
			End If
			'Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
		End If
    End If

	Set oFileSystem = Nothing
    CopyFiles = lErrorNumber
    Err.Clear
End Function 'End of CopyFiles

Function CopyFolder(sSourceFolderPath, sTargetFolderPath, sErrorDescription)
'************************************************************
'Purpose: To copy the specified folder to the target path
'Inputs:  sSourceFolderPath, sTargetFolderPath
'Outputs: sErrorDescription
'************************************************************
    On Error Resume Next
	Const S_FUNCTION_NAME = "CopyFolder"
    Dim oFileSystem
    Dim oFolder
    Dim lErrorNumber

    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
	lErrorNumber = Err.number
	If lErrorNumber <> 0 Then
		sErrorDescription = "No se pudo crear una instancia del objeto 'Scripting.FileSystemObject' porque el archivo 'scrrun.dll' no se encuentra correctamente registrado en el servidor de Web. Favor de contactar al Administrador."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	Else
		Set oFolder = oFileSystem.GetFolder(sSourceFolderPath)
		lErrorNumber = Err.number
		If lErrorNumber <> 0 Then
			sErrorDescription = "El directorio '" & sSourceFolderPath & "' no se pudo abrir. Favor de contactar al Administrador."
			If Len(Err.description) > 0 Then
				sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
			End If
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
		Else
			If Not FolderExists(sTargetFolderPath, sErrorDescription) Then Call CreateFolder(sTargetFolderPath, "")
			Call oFolder.Copy(sTargetFolderPath, True)
			lErrorNumber = Err.number
			If lErrorNumber <> 0 Then
				sErrorDescription = "El directorio '" & sSourceFolderPath & "' no pudo copiarse a '" & sTargetFolderPath & "'. Favor de contactar al Administrador."
				If Len(Err.description) > 0 Then
					sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
				End If
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
			End If
		End If
	End If

	Set oFolder = Nothing
	Set oFileSystem = Nothing
	CopyFolder = lErrorNumber
    Err.Clear
End Function 'End of CopyFolder

Function CreateFolder(sFolderPath, sErrorDescription)
'************************************************************
'Purpose: To create a new folder
'Inputs:  sFolderPath
'Outputs: sErrorDescription
'************************************************************
    On Error Resume Next
	Const S_FUNCTION_NAME = "CreateFolder"
    Dim oFileSystem
    Dim sTempFolderPath
    Dim lErrorNumber

    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
    lErrorNumber = Err.number
    If lErrorNumber <> 0 Then
		sErrorDescription = "No se pudo crear una instancia del objeto 'Scripting.FileSystemObject' porque el archivo 'scrrun.dll' no se encuentra correctamente registrado en el servidor de Web. Favor de contactar al Administrador."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
    Else
	    sTempFolderPath = Left(sFolderPath, (InStrRev(sFolderPath, "\", -1, vbBinaryCompare) - Len("\")))
	    If Not FolderExists(sTempFolderPath, sErrorDescription) Then
			'lErrorNumber = CreateFolder(sTempFolderPath, sErrorDescription)
			If lErrorNumber = 0 Then
				oFileSystem.CreateFolder(sTempFolderPath)
				lErrorNumber = Err.number
				If lErrorNumber <> 0 Then
					sErrorDescription = "No se pudo crear el directorio '" & sFolderPath & "'. Favor de contactar al Administrador."
					If Len(Err.description) > 0 Then
						sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
					End If
					Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
				End If
			End If
	    End If
    End If

	Set oFileSystem = Nothing
    CreateFolder = lErrorNumber
    Err.Clear
End Function 'End of CreateFolder

Function DeleteFile(sFilePath, sErrorDescription)
'************************************************************
'Purpose: To delete the specified file
'Inputs:  sFilePath
'Outputs: sErrorDescription
'************************************************************
    On Error Resume Next
	Const S_FUNCTION_NAME = "DeleteFile"
    Dim oFileSystem
    Dim lErrorNumber

    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
    lErrorNumber = Err.number
    If lErrorNumber <> 0 Then
		sErrorDescription = "No se pudo crear una instancia del objeto 'Scripting.FileSystemObject' porque el archivo 'scrrun.dll' no se encuentra correctamente registrado en el servidor de Web. Favor de contactar al Administrador."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
    Else
	    oFileSystem.DeleteFile(sFilePath)
		lErrorNumber = Err.number
		If lErrorNumber <> 0 Then
			sErrorDescription = "No se pudo eliminar el archivo '" & sFilePath & "'. Favor de contactar al Administrador."
			If Len(Err.description) > 0 Then
				sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
			End If
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
		End If
    End If

	Set oFileSystem = Nothing
    DeleteFile = lErrorNumber
    Err.Clear
End Function 'End of DeleteFile

Function DeleteFolder(sFolderPath, sErrorDescription)
'************************************************************
'Purpose: To delete the specified folder
'Inputs:  sFolderPath
'Outputs: sErrorDescription
'************************************************************
    On Error Resume Next
	Const S_FUNCTION_NAME = "DeleteFolder"
    Dim oFileSystem
    Dim oFolder
    Dim lErrorNumber

    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
	lErrorNumber = Err.number
	If lErrorNumber <> 0 Then
		sErrorDescription = "No se pudo crear una instancia del objeto 'Scripting.FileSystemObject' porque el archivo 'scrrun.dll' no se encuentra correctamente registrado en el servidor de Web. Favor de contactar al Administrador."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	Else
		Set oFolder = oFileSystem.GetFolder(sFolderPath)
		lErrorNumber = Err.number
		If lErrorNumber <> 0 Then
			sErrorDescription = "El directorio '" & sFolderPath & "' no se pudo abrir. Favor de contactar al Administrador."
			If Len(Err.description) > 0 Then
				sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
			End If
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
		Else
			Call oFolder.Delete(True)
			lErrorNumber = Err.number
			If lErrorNumber <> 0 Then
				sErrorDescription = "El directorio '" & sFolderPath & "' no pudo ser borrado. Favor de contactar al Administrador."
				If Len(Err.description) > 0 Then
					sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
				End If
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
			End If
		End If
	End If

	Set oFolder = Nothing
	Set oFileSystem = Nothing
	DeleteFolder = lErrorNumber
    Err.Clear
End Function 'End of DeleteFolder

Function FileExists(sFilePath, sErrorDescription)
'************************************************************
'Purpose: To check if the specified file exists
'Inputs:  sFilePath
'Outputs: sErrorDescription
'************************************************************
    On Error Resume Next
	Const S_FUNCTION_NAME = "FileExists"
    Dim oFileSystem
    Dim oTextFile
    Dim lErrorNumber

    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
	lErrorNumber = Err.number
	If lErrorNumber <> 0 Then
		sErrorDescription = "No se pudo crear una instancia del objeto 'Scripting.FileSystemObject' porque el archivo 'scrrun.dll' no se encuentra correctamente registrado en el servidor de Web. Favor de contactar al Administrador."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
		FileExists = False
	Else
		Set oTextFile = oFileSystem.OpenTextFile(sFilePath)
		FileExists = (Err.Number = 0)
	End If

    Err.Clear
End Function 'End of FileExists

Function FolderExists(sFolderPath, sErrorDescription)
'************************************************************
'Purpose: To check if the specified folder exists
'Inputs:  sFolderPath
'Outputs: sErrorDescription
'************************************************************
    On Error Resume Next
	Const S_FUNCTION_NAME = "FolderExists"
    Dim oFileSystem
    Dim lErrorNumber

    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
	lErrorNumber = Err.number
	If lErrorNumber <> 0 Then
		sErrorDescription = "No se pudo crear una instancia del objeto 'Scripting.FileSystemObject' porque el archivo 'scrrun.dll' no se encuentra correctamente registrado en el servidor de Web. Favor de contactar al Administrador."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
		FolderExists = False
	Else
		FolderExists = oFileSystem.FolderExists(sFolderPath)
	End If

    Err.Clear
End Function 'End of FolderExists

Function GetFileContents(sFilePath, sErrorDescription)
'************************************************************
'Purpose: To get the contents of the specified file
'Inputs:  sFilePath
'Outputs: sErrorDescription. A String with the file contents
'************************************************************
    On Error Resume Next
	Const S_FUNCTION_NAME = "GetFileContents"
    Dim oFileSystem
    Dim oFile
    Dim oTextFile
    Dim lErrorNumber

	GetFileContents = ""
    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
	lErrorNumber = Err.number
	If lErrorNumber <> 0 Then
		sErrorDescription = "No se pudo crear una instancia del objeto 'Scripting.FileSystemObject' porque el archivo 'scrrun.dll' no se encuentra correctamente registrado en el servidor de Web. Favor de contactar al Administrador."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	Else
		Set oFile = oFileSystem.GetFile(sFilePath)
		lErrorNumber = Err.number
		If lErrorNumber <> 0 Then
			sErrorDescription = "No se pudo encontrar el archivo '" & sFilePath & "'. Favor de contactar al Administrador."
			If Len(Err.description) > 0 Then
				sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
			End If
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
		Else
			Set oTextFile = oFile.OpenAsTextStream(1)
			lErrorNumber = Err.number
			If lErrorNumber <> 0 Then
				sErrorDescription = "No se pudo abrir el archivo '" & sFilePath & "'. Favor de contactar al Administrador."
				If Len(Err.description) > 0 Then
					sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
				End If
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
			Else
				GetFileContents = oTextFile.ReadAll()
				lErrorNumber = Err.number
				If lErrorNumber <> 0 Then
					sErrorDescription = "No se pudo leer el archivo '" & sFilePath & "'. Favor de contactar al Administrador."
					If Len(Err.description) > 0 Then
						sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
					End If
					Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
				End If
				oTextFile.Close
			End If
		End If
	End If

	Set oTextFile = Nothing
	Set oFile = Nothing
	Set oFileSystem = Nothing
	Err.Clear
End Function 'End of GetFileContents

Function GetFolderContents(sFolderPath, bIncludeFolders, sFolderContents, sErrorDescription)
'************************************************************
'Purpose: To get the contents of the specified folder
'Inputs:  sFolderPath, bIncludeFolders
'Outputs: sFolderContents, sErrorDescription
'************************************************************
    On Error Resume Next
	Const S_FUNCTION_NAME = "GetFolderContents"
    Dim oFileSystem
    Dim oFolder
    Dim oFile
    Dim lErrorNumber

    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
	lErrorNumber = Err.number
	If lErrorNumber <> 0 Then
		sErrorDescription = "No se pudo crear una instancia del objeto 'Scripting.FileSystemObject' porque el archivo 'scrrun.dll' no se encuentra correctamente registrado en el servidor de Web. Favor de contactar al Administrador."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	Else
		Set oFolder = oFileSystem.GetFolder(sFolderPath)
		lErrorNumber = Err.number
		If lErrorNumber <> 0 Then
			sErrorDescription = "No se pudo abrir el directorio '" & sFolderPath & "'."
			If Len(Err.description) > 0 Then
				sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
			End If
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
		Else
			If bIncludeFolders Then
				For Each oFile In oFolder.SubFolders
					sFolderContents = sFolderContents & oFile.Name & LIST_SEPARATOR
				Next
			End If
			For Each oFile In oFolder.Files
				sFolderContents = sFolderContents & oFile.Name & LIST_SEPARATOR
			Next
			sFolderContents = Left(sFolderContents, (Len(sFolderContents) - Len(LIST_SEPARATOR)))
		End If
	End If

	Set oFileSystem = Nothing
	GetFolderContents = lErrorNumber
	Err.Clear
End Function 'End of GetFolderContents

Function IsFolderEmpty(sFolderPath, sErrorDescription)
'************************************************************
'Purpose: To check if the specified folder is empty
'Inputs:  sFolderPath
'Outputs: sErrorDescription
'************************************************************
    On Error Resume Next
	Const S_FUNCTION_NAME = "IsFolderEmpty"
    Dim oFileSystem
    Dim oFolder
    Dim lErrorNumber

    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
	lErrorNumber = Err.number
	If lErrorNumber <> 0 Then
		sErrorDescription = "No se pudo crear una instancia del objeto 'Scripting.FileSystemObject' porque el archivo 'scrrun.dll' no se encuentra correctamente registrado en el servidor de Web. Favor de contactar al Administrador."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	Else
		Set oFolder = oFileSystem.GetFolder(sFolderPath)
		If Err.number <> 0 Then
			IsFolderEmpty = True
		Else
			IsFolderEmpty = (oFolder.Files.Count = 0)
		End If
	End If

	Set oFolder = Nothing
	Set oFileSystem = Nothing
    Err.Clear
End Function 'End of IsFolderEmpty

Function MoveAllFiles(sSourceFolderPath, sTargetFolderPath, sErrorDescription)
'************************************************************
'Purpose: To move the files in the specified folder to the
'         target path
'Inputs:  sSourceFolderPath, sTargetFolderPath
'Outputs: sErrorDescription
'************************************************************
    On Error Resume Next
	Const S_FUNCTION_NAME = "MoveAllFiles"
    Dim oFileSystem
    Dim oFolder
    Dim oFile
    Dim lErrorNumber

    If StrComp(sSourceFolderPath, sTargetFolderPath, vbTextCompare) <> 0 Then
		Set oFileSystem = CreateObject("Scripting.FileSystemObject")
		lErrorNumber = Err.number
		If lErrorNumber <> 0 Then
			sErrorDescription = "No se pudo crear una instancia del objeto 'Scripting.FileSystemObject'. El archivo 'scrrun.dll' no está correctamente registrado en el servidor de Web. Favor de contactar al Administrador."
			If Len(Err.description) > 0 Then
				sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web:&nbsp;</B>" & Err.description
			End If
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
		Else
			Set oFolder = oFileSystem.GetFolder(sSourceFolderPath)
			lErrorNumber = Err.number
			If lErrorNumber <> 0 Then
				sErrorDescription = "No se pudo abrir el directorio '" & sFolderPath & "'. Favor de contactar al Administrador."
				If Len(Err.description) > 0 Then
					sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web:&nbsp;</B>" & Err.description
				End If
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
			Else
				For Each oFile In oFolder.Files
					If InStr(1, oFile.Name, ".xml", vbTextCompare) > 0 Then
						lErrorNumber = CopyFile(sSourceFolderPath & "\" & oFile.Name, sTargetFolderPath & "\" & oFile.Name, sErrorDescription)
						If lErrorNumber = 0 Then lErrorNumber = DeleteFile(sSourceFolderPath & "\" & oFile.Name, sErrorDescription)
						If lErrorNumber <> 0 Then Exit For
					End If
				Next
			End If
		End If
	End If

	Set oFileSystem = Nothing
	GetFolderContents = lErrorNumber
	Err.Clear
End Function 'End of MoveAllFiles

Function RenameFileExtension(sFilePath, sOrgExt, sDesExt, sErrorDescription)
'************************************************************
'Purpose: To change the specified file extension
'Inputs:  sFilePath
'Outputs: sErrorDescription
'************************************************************
    On Error Resume Next
	Const S_FUNCTION_NAME = "RenameFileExtension"
    Dim oFileSystem
    Dim lErrorNumber

    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
    lErrorNumber = Err.number
    If lErrorNumber <> 0 Then
		sErrorDescription = "No se pudo crear una instancia del objeto 'Scripting.FileSystemObject' porque el archivo 'scrrun.dll' no se encuentra correctamente registrado en el servidor de Web. Favor de contactar al Administrador."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
    Else
	    oFileSystem.MoveFile sFilePath, Replace(sFilePath, sOrgExt, sDesExt)
		lErrorNumber = Err.number
		If lErrorNumber <> 0 Then
			sErrorDescription = "No se pudo eliminar el archivo '" & sFilePath & "'. Favor de contactar al Administrador."
			If Len(Err.description) > 0 Then
				sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
			End If
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
		End If
    End If

	Set oFileSystem = Nothing
    RenameFileExtension = lErrorNumber
    Err.Clear
End Function 'End of RenameFileExtension

Function SaveTextToFile(sFilePath, sContents, sErrorDescription)
'************************************************************
'Purpose: To save the given text into the specified file
'Inputs:  sFilePath, sContents
'Outputs: sErrorDescription
'************************************************************
    On Error Resume Next
	Const S_FUNCTION_NAME = "SaveTextToFile"
    Dim oFileSystem 
    Dim oTextFile 
    Dim lErrorNumber

    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
	lErrorNumber = Err.number
	If lErrorNumber <> 0 Then
		sErrorDescription = "No se pudo crear una instancia del objeto 'Scripting.FileSystemObject' porque el archivo 'scrrun.dll' no se encuentra correctamente registrado en el servidor de Web. Favor de contactar al Administrador."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	Else
		Set oTextFile = oFileSystem.OpenTextFile(sFilePath, 2, True)
		lErrorNumber = Err.number
		If lErrorNumber <> 0 Then
			sErrorDescription = "No se pudo abrir el archivo '" & sFilePath & "'. Favor de contactar al Administrador."
			If Len(Err.description) > 0 Then
				sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
			End If
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
		Else
			oTextFile.Write sContents
			lErrorNumber = Err.number
			If lErrorNumber <> 0 Then
				sErrorDescription = "No se pudieron escribir los contenidos especificados al archivo '" & sFilePath & "'. Favor de contactar al Administrador."
				If Len(Err.description) > 0 Then
					sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
				End If
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FileLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
			Else
			    oTextFile.Close
			End If
		End If
	End If

    Set oTextFile = Nothing
    Set oFileSystem = Nothing
    SaveTextToFile = lErrorNumber
	Err.Clear
End Function 'End of SaveTextToFile

Function TestFileSystemObject(sFolderPath, sErrorDescription)
'************************************************************
'Purpose: To check if the FileSystemObject is correctly installed
'Inputs:  sFolderPath
'Outputs: sErrorDescription
'************************************************************
    On Error Resume Next
	Const S_FUNCTION_NAME = "TestFileSystemObject"
    Dim oFileSystem
    Dim oFolder
    Dim oFile
    Dim sFolderContents

    Response.Write "<B>PRUEBA DEL OBJECTO FileSystemObject</B><BR /><B>Path: </B>" & sFolderPath & "<BR /><BR />"
    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
    Response.Write "<B>CreateObject(""Scripting.FileSystemObject""):</B> " & Err.number & " - " & Err.Description & "<BR />"
    Response.Write "<B>IsObject(oFileSystem):</B> " & IsObject(oFileSystem) & "<BR />"
    Response.Write "<B>IsNull(oFileSystem):</B> " & IsNull(oFileSystem) & "<BR />"
    Response.Write "<B>IsEmpty(oFileSystem):</B> " & IsEmpty(oFileSystem) & "<BR />"
    Response.Write "<B>oFileSystem.FolderExists(sFolderPath):</B> " & oFileSystem.FolderExists(sFolderPath) & "<BR />"
    If Err.number = 0 Then
		Set oFolder = oFileSystem.GetFolder(sFolderPath)
	    Response.Write "<B>oFolder = oFileSystem.GetFolder(sFolderPath):</B> " & Err.number & " - " & Err.Description & "<BR />"
	    Response.Write "<B>IsObject(oFolder.Files):</B> " & IsObject(oFolder.Files) & "<BR />"
	    Response.Write "<B>IsNull(oFolder.Files):</B> " & IsNull(oFolder.Files) & "<BR />"
	    Response.Write "<B>IsEmpty(oFolder.Files):</B> " & IsEmpty(oFolder.Files) & "<BR />"
		For Each oFile In oFolder.Files
			sFolderContents = sFolderContents & oFile.Name & LIST_SEPARATOR
			If err.number <> 0 Then
				Response.Write "<B>Err:</B> " & Err.number & " - " & Err.Description & "<BR />"
				Exit For
			End If
		Next
		sFolderContents = Left(sFolderContents, (Len(sFolderContents) - Len(LIST_SEPARATOR)))
		Response.Write "<B>Contenidos del directorio:</B> " & sFolderContents & "<BR />"
	End If

	Set oFile = Nothing
	Set oFolder = Nothing
	Set oFileSystem = Nothing
    TestFileSystemObject = Err.number
    Err.Clear
End Function 'End of TestFileSystemObject
%>