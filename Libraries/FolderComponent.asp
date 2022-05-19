<%
Const S_NAME_FOLDER = 0
Const S_PATH_FOLDER = 1
Const S_PARENT_NAME_FOLDER = 2
Const N_SIZE_FOLDER = 3
Const S_TYPE_FOLDER = 4
Const S_DATE_CREATED_FOLDER = 5
Const S_DATE_MODIFIED_FOLDER = 6
Const S_DATE_ACCESSED_FOLDER = 7
Const O_FILE_SYSTEM_OBJECT_FOLDER = 8
Const O_FOLDER_OBJECT_FOLDER = 9
Const O_FILES_FOLDER = 10
Const O_SUBFOLDERS_FOLDER = 11
Const S_TARGET_PAGE_FOLDER = 12
Const S_FILE_TARGET_PAGE_FOLDER = 13
Const S_EXTRA_URL_FOLDER = 14
Const S_JAVASCRIPT_FOLDER = 15
Const S_FILE_JAVASCRIPT_FOLDER = 16
Const N_START_LEVEL_FOLDER = 17
Const N_DISPLAY_LEVEL_FOLDER = 18
Const B_DISPLAY_SUBFOLDERS_FOLDER = 19
Const B_DISPLAY_FILES_FOLDER = 20
Const S_FILTER_FOR_FILES_FOLDER = 21
Const B_IS_EMPTY_FOLDER = 22
Const B_COMPONENT_INITIALIZED_FOLDER = 23

Const N_FOLDER_COMPONENT_SIZE = 23

Dim aFolderComponent()
Redim aFolderComponent(N_FOLDER_COMPONENT_SIZE)

Function InitializeFolderComponent(oRequest, aFolderComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Folder Component
'         using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aFolderComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeFolderComponent"
	Redim Preserve aFolderComponent(N_FOLDER_COMPONENT_SIZE)

	If IsEmpty(aFolderComponent(S_NAME_FOLDER)) Then
		If Len(oRequest("FolderName").Item) > 0 Then
			aFolderComponent(S_NAME_FOLDER) = oRequest("FolderName").Item
		Else
			aFolderComponent(S_NAME_FOLDER) = ""
		End If
	End If
	If IsEmpty(aFolderComponent(S_PATH_FOLDER)) Then
		If Len(oRequest("FolderPath").Item) > 0 Then
			aFolderComponent(S_PATH_FOLDER) = oRequest("FolderPath").Item
		ElseIf InStr(1, aFolderComponent(S_NAME_FOLDER), "\") > 0 Then
			aFolderComponent(S_PATH_FOLDER) = Left(aFolderComponent(S_NAME_FOLDER), InStrRev(aFolderComponent(S_NAME_FOLDER), "\"))
			aFolderComponent(S_NAME_FOLDER) = Right(aFolderComponent(S_NAME_FOLDER), Len(aFolderComponent(S_NAME_FOLDER)) - Len(aFolderComponent(S_PATH_FOLDER)))
		Else
			aFolderComponent(S_PATH_FOLDER) = ""
		End If
	End If
	aFolderComponent(S_NAME_FOLDER) = Replace(aFolderComponent(S_NAME_FOLDER), "/", "", 1, -1, vbBinaryCompare)
	aFolderComponent(S_PATH_FOLDER) = Replace(aFolderComponent(S_PATH_FOLDER), "/", "\", 1, -1, vbBinaryCompare)
	If Len(aFolderComponent(S_PATH_FOLDER)) > 0 Then
		If InStr(1, ".\", Left(aFolderComponent(S_PATH_FOLDER), Len(".")), vbBinaryCompare) > 0 Then
			aFolderComponent(S_PATH_FOLDER) = Server.MapPath(aFolderComponent(S_PATH_FOLDER))
		End If
		If StrComp(Right(aFolderComponent(S_PATH_FOLDER), Len("\")), "\", 1) <> 0 Then
			aFolderComponent(S_PATH_FOLDER) = aFolderComponent(S_PATH_FOLDER) & "\"
		End If
	End If

	If IsEmpty(aFolderComponent(S_PARENT_NAME_FOLDER)) Then
		aFolderComponent(S_PARENT_NAME_FOLDER) = ""
	End If
	If IsEmpty(aFolderComponent(N_SIZE_FOLDER)) Then
		aFolderComponent(N_SIZE_FOLDER) = ""
	End If
	If IsEmpty(aFolderComponent(S_TYPE_FOLDER)) Then
		aFolderComponent(S_TYPE_FOLDER) = ""
	End If
	If IsEmpty(aFolderComponent(S_DATE_CREATED_FOLDER)) Then
		aFolderComponent(S_DATE_CREATED_FOLDER) = ""
	End If
	If IsEmpty(aFolderComponent(S_DATE_MODIFIED_FOLDER)) Then
		aFolderComponent(S_DATE_MODIFIED_FOLDER) = ""
	End If
	If IsEmpty(aFolderComponent(S_DATE_ACCESSED_FOLDER)) Then
		aFolderComponent(S_DATE_ACCESSED_FOLDER) = ""
	End If

	If IsEmpty(aFolderComponent(S_TARGET_PAGE_FOLDER)) Then
		aFolderComponent(S_TARGET_PAGE_FOLDER) = GetASPFileName("")
	End If
	If IsEmpty(aFolderComponent(S_FILE_TARGET_PAGE_FOLDER)) Then
		aFolderComponent(S_FILE_TARGET_PAGE_FOLDER) = GetASPFileName("")
	End If
	If IsEmpty(aFolderComponent(S_EXTRA_URL_FOLDER)) Then
		aFolderComponent(S_EXTRA_URL_FOLDER) = ""
	End If
	If IsEmpty(aFolderComponent(S_JAVASCRIPT_FOLDER)) Then
		aFolderComponent(S_JAVASCRIPT_FOLDER) = ""
	End If
	If IsEmpty(aFolderComponent(S_FILE_JAVASCRIPT_FOLDER)) Then
		aFolderComponent(S_FILE_JAVASCRIPT_FOLDER) = ""
	End If

	If IsEmpty(aFolderComponent(N_START_LEVEL_FOLDER)) Then
		If Len(oRequest("FolderStartLevel").Item) > 0 Then
			aFolderComponent(N_START_LEVEL_FOLDER) = CInt(oRequest("FolderStartLevel").Item)
		Else
			aFolderComponent(N_START_LEVEL_FOLDER) = -1
		End If
	End If

	If IsEmpty(aFolderComponent(N_DISPLAY_LEVEL_FOLDER)) Then
		If Len(oRequest("FolderDisplayLevel").Item) > 0 Then
			aFolderComponent(N_DISPLAY_LEVEL_FOLDER) = CInt(oRequest("FolderDisplayLevel").Item)
		Else
			aFolderComponent(N_DISPLAY_LEVEL_FOLDER) = -1
		End If
	End If

	If IsEmpty(aFolderComponent(B_DISPLAY_SUBFOLDERS_FOLDER)) Then
		aFolderComponent(B_DISPLAY_SUBFOLDERS_FOLDER) = (Len(oRequest("HideSubfolders").Item) = 0)
	End If
	If IsEmpty(aFolderComponent(B_DISPLAY_FILES_FOLDER)) Then
		aFolderComponent(B_DISPLAY_FILES_FOLDER) = (Len(oRequest("HideFiles").Item) = 0)
	End If

	If IsEmpty(aFolderComponent(S_FILTER_FOR_FILES_FOLDER)) Then
		If Len(oRequest("FolderFilter").Item) > 0 Then
			aFolderComponent(S_FILTER_FOR_FILES_FOLDER) = oRequest("FolderFilter").Item
		Else
			aFolderComponent(S_FILTER_FOR_FILES_FOLDER) = ""
		End If
	End If

	aFolderComponent(B_COMPONENT_INITIALIZED_FOLDER) = True
	InitializeFolderComponent = Err.number
	Err.Clear
End Function

Function CreateFolderObject(oRequest, aFolderComponent, sErrorDescription)
'************************************************************
'Purpose: To create a File System Object
'Inputs:  oRequest, aFolderComponent
'Outputs: aFolderComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CreateFolderObject"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aFolderComponent(B_COMPONENT_INITIALIZED_FOLDER)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeFolderComponent(oRequest, aFolderComponent)
	End If

	Set aFolderComponent(O_FILE_SYSTEM_OBJECT_FOLDER) = Server.CreateObject("Scripting.FileSystemObject")
	lErrorNumber = Err.number
	If lErrorNumber <> 0 Then
		sErrorDescription = "El archivo 'scrrun.dll' no se encuentra registrado correctamente en el servidor Web. Favor de contactar al Administrador."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FolderComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CreateFolderObject = lErrorNumber
	Err.Clear
End Function

Function OpenFolder(oRequest, aFolderComponent, sErrorDescription)
'************************************************************
'Purpose: To open a folder given the path and the name of it
'Inputs:  oRequest, aFolderComponent
'Outputs: aFolderComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "OpenFolder"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aFolderComponent(B_COMPONENT_INITIALIZED_FOLDER)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeFolderComponent(oRequest, aFolderComponent)
	End If

	If Len(aFolderComponent(S_NAME_FOLDER)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el nombre del directorio."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FolderComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If (Not IsObject(aFolderComponent(O_FILE_SYSTEM_OBJECT_FOLDER))) Or (aFolderComponent(O_FILE_SYSTEM_OBJECT_FOLDER) Is Nothing) Then
			lErrorNumber = CreateFolderObject(oRequest, aFolderComponent, sErrorDescription)
		End If

		If lErrorNumber = 0 Then
			Set aFolderComponent(O_FOLDER_OBJECT_FOLDER) = aFolderComponent(O_FILE_SYSTEM_OBJECT_FOLDER).GetFolder(aFolderComponent(S_PATH_FOLDER) & aFolderComponent(S_NAME_FOLDER))
			lErrorNumber = Err.number
			If lErrorNumber = 53 Then
				sErrorDescription = "El directorio '" & aFolderComponent(S_PATH_FOLDER) & aFolderComponent(S_NAME_FOLDER) & "' no existe."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FolderComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			ElseIf lErrorNumber <> 0 Then
				sErrorDescription = "El directorio '" & aFolderComponent(S_PATH_FOLDER) & aFolderComponent(S_NAME_FOLDER) & "' no pudo ser abierto."
				If Len(Err.description) > 0 Then
					sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
				End If
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FolderComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			End If
		End If
	End If

	OpenFolder = lErrorNumber
	Err.Clear
End Function

Function GetFolderInformation(oRequest, aFolderComponent, sErrorDescription)
'************************************************************
'Purpose: To retrieve the information about the given folder
'Inputs:  oRequest, aFolderComponent
'Outputs: aFolderComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetFolderInformation"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aFolderComponent(B_COMPONENT_INITIALIZED_FOLDER)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeFolderComponent(oRequest, aFolderComponent)
	End If

	If (Not IsObject(aFolderComponent(O_FOLDER_OBJECT_FOLDER))) Or (aFolderComponent(O_FOLDER_OBJECT_FOLDER) Is Nothing) Then
		lErrorNumber = OpenFolder(oRequest, aFolderComponent, sErrorDescription)
	End If

	If lErrorNumber = 0 Then
		aFolderComponent(S_DATE_CREATED_FOLDER) = aFolderComponent(O_FOLDER_OBJECT_FOLDER).DateCreated
		aFolderComponent(S_DATE_MODIFIED_FOLDER) = aFolderComponent(O_FOLDER_OBJECT_FOLDER).DateLastModified
		aFolderComponent(S_DATE_ACCESSED_FOLDER) = aFolderComponent(O_FOLDER_OBJECT_FOLDER).DateLastAccessed
		aFolderComponent(N_SIZE_FOLDER) = aFolderComponent(O_FOLDER_OBJECT_FOLDER).Size
		aFolderComponent(S_TYPE_FOLDER) = aFolderComponent(O_FOLDER_OBJECT_FOLDER).Type
		aFolderComponent(S_PARENT_NAME_FOLDER) = aFolderComponent(O_FOLDER_OBJECT_FOLDER).ParentFolder.Name
		Set aFolderComponent(O_FILES_FOLDER) = aFolderComponent(O_FOLDER_OBJECT_FOLDER).Files
		Set aFolderComponent(O_SUBFOLDERS_FOLDER) = aFolderComponent(O_FOLDER_OBJECT_FOLDER).SubFolders
	End If

	GetFolderInformation = lErrorNumber
	Err.Clear
End Function

Function CloseFolder(aFolderComponent, sErrorDescription)
'************************************************************
'Purpose: To close a folder and clean the used objects
'Inputs:  aFolderComponent
'Outputs: aFolderComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CloseFolder"
	Dim lErrorNumber

	aFolderComponent(O_FOLDER_OBJECT_FOLDER).Close()
	lErrorNumber = Err.number
	If lErrorNumber <> 0 Then
		sErrorDescription = "El directorio '" & aFolderComponent(S_PATH_FOLDER) & aFolderComponent(S_NAME_FOLDER) & "' no pudo ser cerrado."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FolderComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	Set aFolderComponent(O_FILE_SYSTEM_OBJECT_FOLDER) = Nothing
	Set aFolderComponent(O_FOLDER_OBJECT_FOLDER) = Nothing
	CloseFolder = lErrorNumber
	Err.Clear
End Function

Function CleanFolderComponent(aFolderComponent)
'************************************************************
'Purpose: To get rid of the objects used in the component and erase the component itself
'Inputs:  aFolderComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CleanFolderComponent"

	Set aFolderComponent(O_FILE_SYSTEM_OBJECT_FOLDER) = Nothing
	Set aFolderComponent(O_FOLDER_OBJECT_FOLDER) = Nothing
	Set aFolderComponent(O_FILES_FOLDER) = Nothing
	Set aFolderComponent(O_SUBFOLDERS_FOLDER) = Nothing
	Erase aFolderComponent

	CleanFolderComponent = lErrorNumber
	Err.Clear
End Function

Function DisplayFolderInformation(oRequest, aFolderComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information related to a given folder
'Inputs:  oRequest, aFolderComponent
'Outputs: aFolderComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayFolderInformation"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aFolderComponent(B_COMPONENT_INITIALIZED_FOLDER)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeFolderComponent(oRequest, aFolderComponent)
	End If

	If (Not IsObject(aFolderComponent(O_FOLDER_OBJECT_FOLDER))) Or (aFolderComponent(O_FOLDER_OBJECT_FOLDER) Is Nothing) Then
		lErrorNumber = GetFolderInformation(oRequest, aFolderComponent, sErrorDescription)
	End If

	If lErrorNumber = 0 Then
		Response.Write "<TABLE BGCOLOR=""#" & S_WIDGET_FRAME_FOR_GUI & """ WIDTH=""400"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""1""><TR><TD>" & vbNewLine
			Response.Write "<TABLE BGCOLOR=""#FFFFFF"" WIDTH=""100%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""3""><TR><FORM><TD>" & vbNewLine
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & vbNewLine
					Response.Write "<IMG SRC=""Images/ClosedFolder.gif"" WIDTH=""32"" HEIGHT=""32"" /><FONT SIZE=""2"">&nbsp;" & aFolderComponent(S_NAME_FOLDER) & "</FONT>" & vbNewLine
					Response.Write "<HR WIDTH=""98%"" />" & vbNewLine
					Response.Write "<B>Location:</B> " & aFolderComponent(S_PATH_FOLDER) & "<BR />" & vbNewLine
					Response.Write "<B>Type:</B> " & aFolderComponent(S_TYPE_FOLDER) & "<BR />" & vbNewLine
					Response.Write "<B>Size:</B> " & aFolderComponent(N_SIZE_FOLDER) & " Bytes<BR />" & vbNewLine
					Response.Write "<B>Parent folder:</B> " & aFolderComponent(S_PARENT_NAME_FOLDER) & vbNewLine
					Response.Write "<HR WIDTH=""98%"" />" & vbNewLine
					Response.Write "<B>Created:</B> " & aFolderComponent(S_DATE_CREATED_FOLDER) & "<BR />" & vbNewLine
					Response.Write "<B>Modified:</B> " & aFolderComponent(S_DATE_MODIFIED_FOLDER) & "<BR />" & vbNewLine
					Response.Write "<B>Accessed:</B> " & aFolderComponent(S_DATE_ACCESSED_FOLDER) & vbNewLine
				Response.Write "</FONT>" & vbNewLine
			Response.Write "</TD></FORM></TR></TABLE>" & vbNewLine
		Response.Write "</TD></TR></TABLE>" & vbNewLine
	End If

	DisplayFolderInformation = lErrorNumber
	Err.Clear
End Function

Function DisplayFolderPath(oRequest, aFolderComponent, sErrorDescription)
'************************************************************
'Purpose: To display the path of a given folder
'Inputs:  oRequest, aFolderComponent
'Outputs: aFolderComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayFolderPath"
	Dim aFolderPath
	Dim sTempFolderName
	Dim sTempFolderPath
	Dim iIndex, iLevel
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aFolderComponent(B_COMPONENT_INITIALIZED_FOLDER)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeFolderComponent(oRequest, aFolderComponent)
	End If

	If (Not IsObject(aFolderComponent(O_FOLDER_OBJECT_FOLDER))) Or (aFolderComponent(O_FOLDER_OBJECT_FOLDER) Is Nothing) Then
		lErrorNumber = GetFolderInformation(oRequest, aFolderComponent, sErrorDescription)
	End If

	If lErrorNumber = 0 Then
		aFolderPath = Split(aFolderComponent(S_PATH_FOLDER), "\", -1, vbBinaryCompare)
		iLevel = 1
		For iIndex = 0 To (UBound(aFolderPath) - 1)
			sTempFolderName = aFolderPath(iIndex)
			sTempFolderPath = Left(aFolderComponent(S_PATH_FOLDER), (InStr(1, aFolderComponent(S_PATH_FOLDER), sTempFolderName) - Len("\\")))
			If (aFolderComponent(N_DISPLAY_LEVEL_FOLDER) < 0) Or (iLevel > aFolderComponent(N_DISPLAY_LEVEL_FOLDER)) Then
				Response.Write "<A "
					If (aFolderComponent(N_START_LEVEL_FOLDER) < 0) Or (iLevel > aFolderComponent(N_START_LEVEL_FOLDER)) Then
						Response.Write "HREF=""" & aFolderComponent(S_TARGET_PAGE_FOLDER) & "?FolderName=" & Server.URLEncode(sTempFolderName) & "&FolderPath=" & Server.URLEncode(sTempFolderPath) & "&FolderFilter=" & Server.URLEncode(aFolderComponent(S_FILTER_FOR_FILES_FOLDER)) & aFolderComponent(S_EXTRA_URL_FOLDER) & """"
					End If
				Response.Write ">" & sTempFolderName & "</A> > " & vbNewLine
			End If
			iLevel = iLevel + 1
		Next
		Response.Write "<B>" & aFolderComponent(S_NAME_FOLDER) & "</B>" & vbNewLine
	End If
	DisplayFolderPath = lErrorNumber
	Err.Clear
End Function

Function DisplayFolderContents(oRequest, bLink, aFolderComponent, sErrorDescription)
'************************************************************
'Purpose: To display the contents of a given folder
'Inputs:  oRequest, bLink, aFolderComponent
'Outputs: aFolderComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayFolderContents"
	Dim oItemInFolder
	Dim sTempItemName
	Dim sTempFileJavascript
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aFolderComponent(B_COMPONENT_INITIALIZED_FOLDER)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeFolderComponent(oRequest, aFolderComponent)
	End If

	If (Not IsObject(aFolderComponent(O_FOLDER_OBJECT_FOLDER))) Or (aFolderComponent(O_FOLDER_OBJECT_FOLDER) Is Nothing) Then
		lErrorNumber = GetFolderInformation(oRequest, aFolderComponent, sErrorDescription)
	End If

	aFolderComponent(B_IS_EMPTY_FOLDER) = True
	If lErrorNumber = 0 Then
		Response.Write "<TABLE WIDTH=""98%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			If aFolderComponent(B_DISPLAY_SUBFOLDERS_FOLDER) Then
				For Each oItemInFolder In aFolderComponent(O_SUBFOLDERS_FOLDER)
					sTempItemName = oItemInFolder.Name
					If Len(sTempItemName) > 0 Then
						Response.Write "<TR>" & vbNewLine
							Response.Write "<TD WIDTH=""20"">" & vbNewLine
								Response.Write "<IMG SRC=""Images/IcnFolder.gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=""0"" />&nbsp;" & vbNewLine
							Response.Write "</TD>" & vbNewLine
							Response.Write "<TD NOWRAP=""1""><FONT FACE=""Arial"" SIZE=""2"">" & vbNewLine
								Response.Write "<A HREF=""" & aFolderComponent(S_TARGET_PAGE_FOLDER) & "?FolderName=" & Server.URLEncode(sTempItemName) & "&FolderPath=" & Server.URLEncode(aFolderComponent(S_PATH_FOLDER) & aFolderComponent(S_NAME_FOLDER)) & """ STYLE=""text-decoration: none"">" & sTempItemName & "</A>" & vbNewLine
							Response.Write "</FONT></TD>" & vbNewLine
						Response.Write "</TR>" & vbNewLine
						aFolderComponent(B_IS_EMPTY_FOLDER) = False
					End If
				Next
				Response.Flush()
			Else
				For Each oItemInFolder In aFolderComponent(O_SUBFOLDERS_FOLDER)
					sTempItemName = oItemInFolder.Name
					If Len(sTempItemName) > 0 Then
						aFolderComponent(B_IS_EMPTY_FOLDER) = False
						Exit For
					End If
				Next
			End If
			If aFolderComponent(B_DISPLAY_FILES_FOLDER) Then
				For Each oItemInFolder In aFolderComponent(O_FILES_FOLDER)
					sTempItemName = oItemInFolder.Name
					If Len(sTempItemName) > 0 Then
						If (Len(aFolderComponent(S_FILTER_FOR_FILES_FOLDER)) = 0) Or (InStr(1, sTempItemName, aFolderComponent(S_FILTER_FOR_FILES_FOLDER), vbTextCompare) > 0) Then
							Response.Write "<TR>" & vbNewLine
								Response.Write "<TD WIDTH=""20"">" & vbNewLine
									Response.Write "<IMG SRC=""Images/IcnFile"
										Select Case LCase(Right(sTempItemName, Len(".txt")))
											Case ".doc", ".htm", ".img", ".mov", ".msg", ".pdf", ".ppt", ".wav", ".xls", ".zip"
												Response.Write UCase(Right(sTempItemName, Len("txt")))
											Case "docx"
												Response.Write "DOC"
											Case "html"
												Response.Write "HTM"
											Case ".bmp", ".gif", "jpeg", ".jpg", ".tif"
												Response.Write "IMG"
											Case ".avi", "mpeg", ".mpg", ".qt", ".wmv"
												Response.Write "MOV"
											Case ".pps"
												Response.Write "PPT"
											Case "pptx"
												Response.Write "PPT"
											Case ".au", ".mp3", ".wma"
												Response.Write "WAV"
											Case "xlsx"
												Response.Write "XLS"
										End Select
									Response.Write ".gif"" WIDTH=""16"" HEIGHT=""16"" BORDER=""0"" />&nbsp;" & vbNewLine
								Response.Write "</TD>" & vbNewLine
								Response.Write "<TD NOWRAP=""1""><FONT FACE=""Arial"" SIZE=""2"">" & vbNewLine
									If Len(aFolderComponent(S_FILE_JAVASCRIPT_FOLDER)) = 0 Then
										Response.Write "<A HREF=""" & aFolderComponent(S_FILE_TARGET_PAGE_FOLDER) & "?FileName=" & Server.URLEncode(sTempItemName) & "&FilePath=" & Server.URLEncode(aFolderComponent(S_PATH_FOLDER) & aFolderComponent(S_NAME_FOLDER)) & aFolderComponent(S_EXTRA_URL_FOLDER) & """ STYLE=""text-decoration: none"">" & sTempItemName & " (" & DisplayFileSize(oItemInFolder.Size) & ")</A>" & vbNewLine
									Else
										sTempFileJavascript = Replace(Replace(Replace(Replace(aFolderComponent(S_FILE_JAVASCRIPT_FOLDER), "<FILE_NAME>", sTempItemName, 1, -1, vbBinaryCompare), "<FOLDER_NAME>", Replace(Server.URLEncode(aFolderComponent(S_NAME_FOLDER)), "+", " "), 1, -1, vbBinaryCompare), "<FILE_PATH>", Replace(Server.URLEncode(aFolderComponent(S_PATH_FOLDER) & aFolderComponent(S_NAME_FOLDER)), "+", " "), 1, -1, vbBinaryCompare), "%5C", "\\", 1, -1, vbBinaryCompare)
										Response.Write "<A HREF=""javascript: " & sTempFileJavascript & """ STYLE=""text-decoration: none"">" & sTempItemName & " (" & DisplayFileSize(oItemInFolder.Size) & ")</A>" & vbNewLine
									End If
								Response.Write "</FONT></TD>" & vbNewLine
								If bLink Then Response.Write "<TD><A HREF=""" & GetASPFileName("") & "?RemoveFile=1&FileName=" & Server.URLEncode(sTempItemName) & "&FilePath=" & Server.URLEncode(aFolderComponent(S_PATH_FOLDER) & aFolderComponent(S_NAME_FOLDER)) & "&" & RemoveParameterFromURLString(RemoveParameterFromURLString(oRequest, "FileName"), "FilePath") & """><IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar archivo"" BORDER=""0"" /></A></TD>" & vbNewLine
							Response.Write "</TR>" & vbNewLine
							aFolderComponent(B_IS_EMPTY_FOLDER) = False
							Response.Flush()
						Else
							aFolderComponent(B_IS_EMPTY_FOLDER) = False
						End If
					End If
				Next
			Else
				For Each oItemInFolder In aFolderComponent(O_FILES_FOLDER)
					sTempItemName = oItemInFolder.Name
					If Len(sTempItemName) > 0 Then
						aFolderComponent(B_IS_EMPTY_FOLDER) = False
						Exit For
					End If
				Next
			End If
		Response.Write "</TABLE>" & vbNewLine
	End If
	DisplayFolderContents = lErrorNumber
	Err.Clear
End Function

Function DisplayFolderContentsAsList(oRequest, aFolderComponent, sErrorDescription)
'************************************************************
'Purpose: To display the contents of a given folder in a coma
'         separated list
'Inputs:  oRequest, aFolderComponent
'Outputs: aFolderComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayFolderContentsAsList"
	Dim oItemInFolder
	Dim sTempItemName
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aFolderComponent(B_COMPONENT_INITIALIZED_FOLDER)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeFolderComponent(oRequest, aFolderComponent)
	End If

	If (Not IsObject(aFolderComponent(O_FOLDER_OBJECT_FOLDER))) Or (aFolderComponent(O_FOLDER_OBJECT_FOLDER) Is Nothing) Then
		lErrorNumber = GetFolderInformation(oRequest, aFolderComponent, sErrorDescription)
	End If

	aFolderComponent(B_IS_EMPTY_FOLDER) = True
	If lErrorNumber = 0 Then
		If aFolderComponent(B_DISPLAY_SUBFOLDERS_FOLDER) Or aFolderComponent(B_DISPLAY_FILES_FOLDER) Then
			If aFolderComponent(B_DISPLAY_SUBFOLDERS_FOLDER) Then
				For Each oItemInFolder In aFolderComponent(O_SUBFOLDERS_FOLDER)
					sTempItemName = oItemInFolder.Name
					If Len(sTempItemName) > 0 Then
						Response.Write "'" & sTempItemName & "', "
						aFolderComponent(B_IS_EMPTY_FOLDER) = False
					End If
				Next
			End If
			If aFolderComponent(B_DISPLAY_FILES_FOLDER) Then
				For Each oItemInFolder In aFolderComponent(O_FILES_FOLDER)
					sTempItemName = oItemInFolder.Name
					If Len(sTempItemName) > 0 Then
						If (Len(aFolderComponent(S_FILTER_FOR_FILES_FOLDER)) = 0) Or (InStr(1, sTempItemName, aFolderComponent(S_FILTER_FOR_FILES_FOLDER), vbTextCompare) > 0) Then
							Response.Write "'" & sTempItemName & "', "
							aFolderComponent(B_IS_EMPTY_FOLDER) = False
						End If
					End If
				Next
			End If
		End If
	End If
	DisplayFolderContentsAsList = lErrorNumber
	Err.Clear
End Function

Function DisplayFolderContentsAsText(oRequest, aFolderComponent, sErrorDescription)
'************************************************************
'Purpose: To display the contents of a given folder
'Inputs:  oRequest, aFolderComponent
'Outputs: aFolderComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayFolderContentsAsText"
	Dim oItemInFolder
	Dim sTempItemName
	Dim sTempFileJavascript
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aFolderComponent(B_COMPONENT_INITIALIZED_FOLDER)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeFolderComponent(oRequest, aFolderComponent)
	End If

	If (Not IsObject(aFolderComponent(O_FOLDER_OBJECT_FOLDER))) Or (aFolderComponent(O_FOLDER_OBJECT_FOLDER) Is Nothing) Then
		lErrorNumber = GetFolderInformation(oRequest, aFolderComponent, sErrorDescription)
	End If

	aFolderComponent(B_IS_EMPTY_FOLDER) = True
	If lErrorNumber = 0 Then
		Response.Write "<FONT FACE=""Arial"" SIZE=""2"">"
			If aFolderComponent(B_DISPLAY_SUBFOLDERS_FOLDER) Then
				Response.Write "<B>"
					For Each oItemInFolder In aFolderComponent(O_SUBFOLDERS_FOLDER)
						sTempItemName = oItemInFolder.Name
						If Len(sTempItemName) > 0 Then
							Response.Write "- " & sTempItemName & "<BR />" & vbNewLine
							aFolderComponent(B_IS_EMPTY_FOLDER) = True
						End If
					Next
				Response.Write "</B>"
				Response.Flush()
			End If
			If aFolderComponent(B_DISPLAY_FILES_FOLDER) Then
				For Each oItemInFolder In aFolderComponent(O_FILES_FOLDER)
					sTempItemName = oItemInFolder.Name
					If Len(sTempItemName) > 0 Then
						If (Len(aFolderComponent(S_FILTER_FOR_FILES_FOLDER)) = 0) Or (InStr(1, sTempItemName, aFolderComponent(S_FILTER_FOR_FILES_FOLDER), vbTextCompare) > 0) Then
							Response.Write sTempItemName & "<BR />"
							aFolderComponent(B_IS_EMPTY_FOLDER) = True
						End If
					End If
				Next
				Response.Flush()
			End If
		Response.Write "</FONT>"
	End If
	DisplayFolderContentsAsText = lErrorNumber
	Err.Clear
End Function

Function DisplayFolderContentsInList(oRequest, aFolderComponent, sErrorDescription)
'************************************************************
'Purpose: To display the contents of a given folder in a form
'         element
'Inputs:  oRequest, aFolderComponent
'Outputs: aFolderComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayFolderContentsInList"
	Dim oItemInFolder
	Dim sTempItemName
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aFolderComponent(B_COMPONENT_INITIALIZED_FOLDER)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeFolderComponent(oRequest, aFolderComponent)
	End If

	If (Not IsObject(aFolderComponent(O_FOLDER_OBJECT_FOLDER))) Or (aFolderComponent(O_FOLDER_OBJECT_FOLDER) Is Nothing) Then
		lErrorNumber = GetFolderInformation(oRequest, aFolderComponent, sErrorDescription)
	End If

	If lErrorNumber = 0 Then
		If aFolderComponent(B_DISPLAY_SUBFOLDERS_FOLDER) Or aFolderComponent(B_DISPLAY_FILES_FOLDER) Then
			aFolderComponent(B_IS_EMPTY_FOLDER) = True
			If aFolderComponent(B_DISPLAY_SUBFOLDERS_FOLDER) Then
				For Each oItemInFolder In aFolderComponent(O_SUBFOLDERS_FOLDER)
					sTempItemName = oItemInFolder.Name
					If Len(sTempItemName) > 0 Then
						Response.Write "<OPTION VALUE=""" & sTempItemName & """>"
							Response.Write sTempItemName
						Response.Write "</OPTION>" & vbNewLine
						aFolderComponent(B_IS_EMPTY_FOLDER) = False
					End If
				Next
				Response.Flush()
			End If
			If aFolderComponent(B_DISPLAY_FILES_FOLDER) Then
				For Each oItemInFolder In aFolderComponent(O_FILES_FOLDER)
					sTempItemName = oItemInFolder.Name
					If Len(sTempItemName) > 0 Then
						If (Len(aFolderComponent(S_FILTER_FOR_FILES_FOLDER)) = 0) Or (InStr(1, sTempItemName, aFolderComponent(S_FILTER_FOR_FILES_FOLDER), vbTextCompare) > 0) Then
							Response.Write "<OPTION VALUE=""" & sTempItemName & """>"
								Response.Write sTempItemName
							Response.Write "</OPTION>" & vbNewLine
							aFolderComponent(B_IS_EMPTY_FOLDER) = False
							Response.Flush()
						End If
					End If
				Next
			End If
			If aFolderComponent(B_IS_EMPTY_FOLDER) Then
				Response.Write "<OPTION VALUE="""">-----------</OPTION>" & vbNewLine
			End If
		End If
	End If
	DisplayFolderContentsInList = lErrorNumber
	Err.Clear
End Function
%>