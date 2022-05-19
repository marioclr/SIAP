<%
Class FileUploader
	Public oFiles
	Private oForm
	Private sOriginalName

	Private Sub Class_Initialize()
	'************************************************************
	'Purpose: To initialize the File Uploader variables
	'************************************************************
		On Error Resume Next

		Set oFiles = Server.CreateObject("Scripting.Dictionary")
		Set oForm = Server.CreateObject("Scripting.Dictionary")
	End Sub

	Private Sub Class_Terminate()
	'************************************************************
	'Purpose: To clean the File Uploader variables
	'************************************************************
		On Error Resume Next

		If IsObject(oFiles) Then
			oFiles.RemoveAll
			Set oFiles = Nothing
		End If
		If IsObject(oForm) Then
			oForm.RemoveAll
			Set oForm = Nothing
		End If
	End Sub

	Public Property Get GetFormValues(sFieldName)
	'************************************************************
	'Purpose: To get the values from the oForm variable
	'Inputs:  sFieldName
	'************************************************************
		If oForm.Exists(sFieldName) Then
			GetFormValues = oForm.Item(sFieldName)
		Else
			GetFormValues = ""
		End If
	End Property

	Public Property Get GetOriginalName()
	'************************************************************
	'Purpose: To get the values from the oForm variable
	'Inputs:  sFieldName
	'************************************************************
			GetOriginalName = sOriginalName
	End Property

	Public Sub Upload()
	'************************************************************
	'Purpose: To get the data of the upload file 
	'************************************************************
		On Error Resume Next
		Dim oBinaryRequest
		Dim sControl
		Dim iStartPosition
		Dim iEndPosition
		Dim iPosition
		Dim sLimitByte
		Dim iPositionFile
		Dim iLimitPosition
		Dim oFileTemp
		Dim sName
		Dim sFormContent

		oBinaryRequest = Request.BinaryRead(Request.TotalBytes)

		iStartPosition = 1
		iEndPosition = InStrB(iStartPosition, oBinaryRequest, ConvertStringToBinary(Chr(13)))
		If (iEndPosition - iStartPosition) <= 0 Then Exit Sub

		sLimitByte = MidB(oBinaryRequest, iStartPosition, (iEndPosition - iStartPosition))
		iLimitPosition = InStrB(1, oBinaryRequest, sLimitByte)

		Do Until iLimitPosition = 0 'InStrB(iLimitPosition, oBinaryRequest, (sLimitByte & ConvertStringToBinary("--")))
			iPosition = InStrB(iLimitPosition, oBinaryRequest, ConvertStringToBinary("Content-Disposition"))
			iPosition = InStrB(iPosition, oBinaryRequest, ConvertStringToBinary("name="))
			iStartPosition = iPosition + 6'Len("name=""")
			iEndPosition = InStrB(iStartPosition, oBinaryRequest, ConvertStringToBinary(Chr(34)))
			sControl = ConvertBinaryToString(MidB(oBinaryRequest, iStartPosition, (iEndPosition - iStartPosition)))
			iPositionFile =InStrB(iLimitPosition, oBinaryRequest, ConvertStringToBinary("filename="))
			'iLimitPosition = InStrB(iEndPosition, oBinaryRequest, sLimitByte)
			
			If (iPositionFile <> 0) And (iPositionFile < InStrB(iEndPosition, oBinaryRequest, sLimitByte)) Then
				Set oFileTemp = New File

				iStartPosition = iPositionFile + 10'Len("filename= ")
				iEndPosition = InStrB(iStartPosition, oBinaryRequest, ConvertStringToBinary(Chr(34)))
				sName = ConvertBinaryToString(MidB(oBinaryRequest, iStartPosition, (iEndPosition - iStartPosition)))
				sOriginalName = Right(sName, (Len(sName) - InStrRev(sName, "\")))
				oFileTemp.sName = Right(sName, (Len(sName) - InStrRev(sName, "\")))
				iPosition = InStrB(iEndPosition, oBinaryRequest, ConvertStringToBinary("Content-Type:"))
				iStartPosition = iPosition + 14'Len("Content-Type: ")
				iEndPosition = InStrB(iStartPosition, oBinaryRequest, ConvertStringToBinary(Chr(13)))
				oFileTemp.sContentType = ConvertBinaryToString(MidB(oBinaryRequest, iStartPosition, (iEndPosition - iStartPosition)))
				iStartPosition = iEndPosition + 4'Len(vbNewLine & vbNewLine)
				iEndPosition = InStrB(iStartPosition, oBinaryRequest, sLimitByte) - 2
				oFileTemp.sData = MidB(oBinaryRequest, iStartPosition, (iEndPosition - iStartPosition))
				If oFileTemp.Size > 0 Then
					Call oFiles.Add(sControl, oFileTemp)
				End If
			Else
				iPosition = InStrB(iPosition, oBinaryRequest, ConvertStringToBinary(Chr(13)))
				iStartPosition = iPosition + 4'Len(vbNewLine & vbNewLine)
				iEndPosition = InStrB(iStartPosition, oBinaryRequest, sLimitByte) - 2
				sFormContent = ConvertBinaryToString(MidB(oBinaryRequest, iStartPosition, (iEndPosition - iStartPosition)))
				If Not oForm.Exists(sControl) Then
					Call oForm.Add(sControl, sFormContent)
				Else
					oForm.Item(sControl) = oForm.Item(sControl) & "," & sFormContent
				End If
			End If

			iLimitPosition = InStrB((iLimitPosition + LenB(sLimitByte)), oBinaryRequest, sLimitByte)
		Loop
	End Sub

	Private Function ConvertStringToBinary(sStringToConvert)
	'************************************************************
	'Purpose: To convert ANSI character data to binary data
	'Inputs:  sStringToConvert
	'************************************************************	
		On Error Resume Next
		Dim iIndex
		Dim sConversion

		For iIndex = 1 To Len(sStringToConvert)
			sConversion = sConversion & ChrB(AscB(Mid(sStringToConvert, iIndex, 1)))
			If Err.Number <> 0 Then Exit For
		Next

		ConvertStringToBinary = sConversion
	End Function

	Private Function ConvertBinaryToString(sBinaryToConvert)
	'************************************************************
	'Purpose: To convert binary data to ANSI character data
	'Inputs:  sBinaryToConvert
	'************************************************************
		On Error Resume Next
		Dim iIndex
		Dim asConversion()

'Response.Write "1. " & Err.Number & ": " & Err.Description & "<BR />"
		Redim asConversion(LenB(sBinaryToConvert) - 1)
'Response.Write "2. " & Err.Number & ": " & Err.Description & "<BR />"
		For iIndex = 1 To LenB(sBinaryToConvert)
			asConversion(iIndex-1) = Chr(AscB(MidB(sBinaryToConvert, iIndex, 1)))
			If Err.Number <> 0 Then Exit For
		Next
'Response.Write "3. " & Err.Number & ": " & Err.Description & "<BR />"

		ConvertBinaryToString = Join(asConversion, "")
	End Function
End Class

Class File
	Public sName
	Public sContentType
	Public sData

	Public Property Get Size()
	'***************************************************************************
	'Purpose: To get the size of a file
	'***************************************************************************
		Size = LenB(sData)
	End Property

	Public Sub SaveFileAs(sPathFile, sFileName, lErrorNumber, sErrorDescription)
	'***************************************************************************
	'Purpose: To save a file using the specified name and path
	'Inputs:  sPathFile, sFileName
	'Outputs: lErrorNumber, sErrorDescription
	'***************************************************************************	
		Dim oFileSystem
		Dim oFile
		Dim iIndex

		If (Len(sPathFile) = 0) Or (Len(sFileName) = 0) Then Exit Sub

		If StrComp(Mid(sPathFile, Len(sPathFile)), "\", vbBinaryCompare) <> 0 Then
			sPathFile = sPathFile & "\"
		End If

		Set oFileSystem = Server.CreateObject("Scripting.FileSystemObject")
		lErrorNumber = Err.number
		If lErrorNumber <> 0 Then
			sErrorDescription = "No se pudo crear una instancia del objeto 'Scripting.FileSystemObject' porque el archivo 'scrrun.dll' no se encuentra correctamente registrado en el servidor de Web. Favor de contactar al Administrador."
			If Len(Err.description) > 0 Then
				sErrorDescription = sErrorDescription & "<BR />Error del servidor de Web: " & Err.description
			End If
		Else
			If Not oFileSystem.FolderExists(sPathFile) Then Exit Sub
			Set oFile = oFileSystem.CreateTextFile(sPathFile & sFileName, True)
			lErrorNumber = Err.number
			If lErrorNumber <> 0 Then
				sErrorDescription = "No se pudo crear el archivo '" & sFileName & "'. Favor de contactar al Administrador."
				If Len(Err.description) > 0 Then
					sErrorDescription = sErrorDescription & "<BR />Error del servidor de Web: " & Err.description
				End If
			Else
				For iIndex = 1 To LenB(sData)
					oFile.Write Chr(AscB(MidB(sData, iIndex, 1)))
				Next 
			End If
		End If
		
		oFile.Close
		Set oFileSystem = Nothing
	End Sub
End Class
%>