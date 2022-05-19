<!-- #include file="DateLibrary.asp" -->
<%
Const S_HEXADECIMAL_DIGITS = "0123456789ABCDEF"

Function BlockTheSystem()
'************************************************************
'Purpose: To block the system
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BlockTheSystem"
	Dim oFileSystem 
	Dim oTextFile 

	Set oFileSystem = CreateObject("Scripting.FileSystemObject")
	If Err.number = 0 Then
		Set oTextFile = oFileSystem.OpenTextFile(Server.MapPath("Logs\Block.txt"), 2, True)
		If Err.number = 0 Then
			oTextFile.Write ""
			If Err.number = 0 Then
				oTextFile.Close
			End If
		End If
	End If
	Application.Contents("SIAP_Block") = "1"
	Call LogErrorInXMLFile(L_ERR_BLOCKED_SYSTEM, "¡El sistema ha sido bloqueado!", 000, "GeneralLibrary.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)

	ReDim aEmailComponent(N_EMAIL_COMPONENT_SIZE)
	aEmailComponent(S_FROM_EMAIL) = S_ADMIN_EMAIL_ACCOUNT
	aEmailComponent(S_TO_EMAIL) = GetAdminOption(aAdminOptionsComponent, SYSTEM_BLOCKED_RECIPIENTS_OPTION)
	aEmailComponent(S_CC_EMAIL) = "victor@jibda.com"
	aEmailComponent(S_SUBJECT_EMAIL) = "SIAP ha sido bloqueado"
	aEmailComponent(S_BODY_EMAIL) = "<FONT FACE=""Arial"" SIZE=""2"">"
		aEmailComponent(S_BODY_EMAIL) = aEmailComponent(S_BODY_EMAIL) & "Este mensaje ha sido enviado por el Sistema de Administración del Personal (SIAP) ya que se han detectado " & GetAdminOption(aAdminOptionsComponent, LOGIN_FAILURES_OPTION) & " intentos fallidos de entrar al sistema desde la misma máquina.<BR /><BR />"
		aEmailComponent(S_BODY_EMAIL) = aEmailComponent(S_BODY_EMAIL) & "Clave de acceso: " & aLoginComponent(S_ACCESS_KEY_LOGIN) & "<BR />"
		aEmailComponent(S_BODY_EMAIL) = aEmailComponent(S_BODY_EMAIL) & "Dirección IP: " & Request.ServerVariables("REMOTE_ADDR") & "<BR />"
		aEmailComponent(S_BODY_EMAIL) = aEmailComponent(S_BODY_EMAIL) & "IP del Sistema: " & SYSTEM_IP & "<BR />"
		aEmailComponent(S_BODY_EMAIL) = aEmailComponent(S_BODY_EMAIL) & "URL del Sistema: " & SYSTEM_PATH & "<BR />"
	aEmailComponent(S_BODY_EMAIL) = aEmailComponent(S_BODY_EMAIL) & "</FONT>"
	lErrorNumber = SendEmail(oRequest, aEmailComponent, sErrorDescription)

    Set oTextFile = Nothing
    Set oFileSystem = Nothing
    BlockTheSystem = Err.number
	Err.Clear
End Function

Function BuildList(sSource, sSeparator, iCounter)
'************************************************************
'Purpose: To create a list repeating iCounter times the source
'         string
'Inputs:  sSource, sSeparator, iCounter
'Outputs: A list
'************************************************************
	Const S_FUNCTION_NAME = "BuildList"
	Dim iIndex

	BuildList = ""
	For iIndex = 1 To iCounter
		BuildList = BuildList & sSource & sSeparator
	Next
	If Len(BuildList) > 0 Then BuildList = Left(BuildList, (Len(BuildList) - Len(sSeparator)))
End Function

Function CheckBlocked()
'************************************************************
'Purpose: To check if the system is blocked
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckBlocked"
	Dim oFileSystem
	Dim oTextFile

	Set oFileSystem = CreateObject("Scripting.FileSystemObject")
	If Err.number = 0 Then
		Set oTextFile = oFileSystem.OpenTextFile(Server.MapPath("Logs/Block.txt"))
		CheckBlocked = ((Err.Number = 0) Or (Len(Application.Contents("SIAP_Block")) > 0))
		Err.Clear
	Else
		CheckBlocked = (Len(Application.Contents("SIAP_Block")) > 0)
	End If

	Err.Clear
End Function

Function CleanDontExport(sFormContents)
'************************************************************
'Purpose: To remove the <DONT_EXPORT> tags
'Inputs:  sFormContents
'Outputs: sFormContents
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CleanDontExport"
	Dim iStartPos
	Dim iEndPos

	iStartPos = InStr(1, sFormContents, "<DONT_EXPORT>", vbBinaryCompare)
	Do While (iStartPos > 0)
		iStartPos = iStartPos - Len("<")
		iEndPos = InStr(iStartPos, sFormContents, "</DONT_EXPORT>", vbBinaryCompare)
		If iEndPos > 0 Then
			iEndPos = iEndPos + Len("</DONT_EXPORT")
			sFormContents = Left(sFormContents, iStartPos) & Right(sFormContents, (Len(sFormContents) - iEndPos))
		End If
		iStartPos = InStr(1, sFormContents, "<DONT_EXPORT>", vbBinaryCompare)
		If Err.number <> 0 Then Exit Do
	Loop

	CleanDontExport = Err.number
	Err.Clear
End Function

Function CleanStringForAttribute(sStringToChange)
'************************************************************
'Purpose: To replace <BR /> and " in a string
'Inputs:  sStringToChange
'Outputs: A string that can be used as HTML text without breaking any tag
'************************************************************
	CleanStringForAttribute = Replace(Replace(Replace(sStringToChange, """", "&#34;"), "<BR />", "&#13;"), "<BR>", "&#13;")
End Function

Function CleanStringForFolderName(sStringToChange)
'************************************************************
'Purpose: To replace \, /, :, *, <, >, |, Á, É, Í, Ó, Ú, Ñ, á, é, í, ó, ú, ñ in a string
'Inputs:  sStringToChange
'Outputs: A string that can be used as HTML text without breaking any tag
'************************************************************
	CleanStringForFolderName = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(sStringToChange, "\", ""), "/", ""), ":", ""), "*", ""), "<", ""), ">", ""), "|", ""), "Á", "A"), "É", "E"), "Í", "I"), "Ó", "O"), "Ú", "U"), "Ñ", "N"), "á", "a"), "é", "e"), "í", "i"), "ó", "o"), "ú", "u"), "ñ", "n"))
End Function

Function CleanStringForHTML(sStringToChange)
'************************************************************
'Purpose: To replace &, <, and " in a string
'Inputs:  sStringToChange
'Outputs: A string that can be used as HTML text without breaking any tag
'************************************************************
	CleanStringForHTML = Replace(Replace(Replace(Replace(Replace(sStringToChange, "&", "&#38;"), "<", "&#60;"), """", "&#34;"), "&#60;BR />", "<BR />"), vbNewLine, "<BR />")
	If StrComp(Right(CleanStringForHTML, Len(";")), ";", vbBinaryCompare) = 0 Then CleanStringForHTML = CleanStringForHTML & " "
End Function

Function CleanStringForJavaScript(sStringToChange)
'************************************************************
'Purpose: To replace \, ', ", vbNewLine, & and = in a string
'Inputs:  sStringToChange
'Outputs: A string that can be used as HTML text without breaking any tag
'************************************************************
	CleanStringForJavaScript = Replace(Replace(Replace(Replace(Replace(Replace(Replace(sStringToChange, "\", "\\"), "/", "\/"), "'", "\'"), """", ""), vbNewLine, "\n"), "&", "and"), "=", "equals")
End Function

Function CleanStringForJavaScriptName(sStringToChange)
'************************************************************
'Purpose: To replace #, (, ), +, -, ., / and \ in a string
'Inputs:  sStringToChange
'Outputs: A string that can be used as HTML text without breaking any tag
'************************************************************
	CleanStringForJavaScriptName = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(sStringToChange, " ", "__32__"), "#", "__35__"), "(", "__40__"), ")", "__41__"), "+", "__43__"), "-", "__45__"), ".", "__46__"), "/", "__47__"), "\", "__92__")
End Function

Function CleanStringFromJavaScriptName(sStringToChange)
'************************************************************
'Purpose: To replace scape codes with #, (, ), +, -, ., / and \ in a string
'Inputs:  sStringToChange
'Outputs: A string that can be used as HTML text without breaking any tag
'************************************************************
	CleanStringFromJavaScriptName = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(sStringToChange, "__32__", " "), "__35__", "#"), "__40__", "("), "__41__", ")"), "__43__", "+"), "__45__", "-"), "__46__", "."), "__47__", "/"), "__92__", "\")
End Function

Function CleanStringForReportField(sStringToChange)
'************************************************************
'Purpose: To replace \, /, :, *, <, >, |, Á, É, Í, Ó, Ú, Ñ, á, é, í, ó, ú, ñ in a string
'Inputs:  sStringToChange
'Outputs: A string that can be used as HTML text without breaking any tag
'************************************************************
	CleanStringForReportField = Trim(Replace(sStringToChange, "<BR />", ""))
End Function

Function CountStringAppearance(sInputString, sStringToFind)
'************************************************************
'Purpose: To count how many times sStringToFind appears
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CountStringAppearance"
	Dim iIndex

	CountStringAppearance = 0
	iIndex = 1
	Do While (iIndex > 0)
		iIndex = InStr(iIndex, sInputString, sStringToFind, vbBinaryCompare)
		If (iIndex > 0) Then
			CountStringAppearance = CountStringAppearance + 1
			iIndex = iIndex + Len(sStringToFind)
			If Err.number <> 0 Then Exit Do
		End If
	Loop
End Function

Function DisplayInstructionsMessage(sTitle, sInstructionMessage)
'************************************************************
'Purpose: To display the instructions in a box
'Inputs:  sTitle, sInstructionMessage
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayInstructionsMessage"

	Response.Write "<TABLE BGCOLOR=""#" & S_WIDGET_FRAME_FOR_GUI & """ WIDTH=""98%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""1""><TR><TD>"
		Response.Write "<TABLE BGCOLOR=""#FFFFFF"" WIDTH=""100%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""3""><TR><TD>"
			If Len(sTitle) > 0 Then
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"" COLOR=""#" & S_INSTRUCTIONS_FOR_GUI & """><B>"
					Response.Write UCase(sTitle)
					If Len(sInstructionMessage) > 0 Then
						Response.Write "<BR /><BR />"
					End If
				Response.Write "</B></FONT>"
			End If
			Response.Write "<IMG SRC=""Images/IcnInformation.gif"" WIDTH=""32"" HEIGHT=""32"" ALIGN=""LEFT"" VALIGN=""ABSMIDDLE"" HSPACE=""5"" />"
			Response.Write "<FONT FACE=""Arial"" SIZE=""2"" COLOR=""#000000"">"
				Response.Write sInstructionMessage
			Response.Write "</FONT>"
		Response.Write "</TD></TR></TABLE>"
	Response.Write "</TD></TR></TABLE>"

	DisplayInstructionsMessage = Err.number
	Err.Clear
End Function

Function DisplayErrorMessage(sTitle, sErrorMessage)
'************************************************************
'Purpose: To display an error message in a box
'Inputs:  sTitle, sErrorMessage
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayErrorMessage"

	Response.Write "<TABLE BGCOLOR=""#" & S_WIDGET_FRAME_FOR_GUI & """ WIDTH=""98%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""1""><TR><TD>"
		Response.Write "<TABLE BGCOLOR=""#FFFFFF"" WIDTH=""100%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""3""><TR><TD>"
			If Len(sTitle) > 0 Then
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"" COLOR=""#" & S_SECOND_TITLE_FOR_GUI & """><B>"
					Response.Write UCase(sTitle)
					If Len(sErrorMessage) > 0 Then
						Response.Write "<BR /><BR />"
					End If
				Response.Write "</B></FONT>"
			End If
			Response.Write "<IMG SRC=""Images/IcnExclamation.gif"" WIDTH=""32"" HEIGHT=""32"" ALIGN=""LEFT"" VALIGN=""ABSMIDDLE"" HSPACE=""5"" VSPACE=""1"" />"
			Response.Write "<FONT FACE=""Arial"" SIZE=""2"" COLOR=""#000000"">"
				Response.Write sErrorMessage
			Response.Write "</FONT>"
		Response.Write "</TD></TR></TABLE>"
	Response.Write "</TD></TR></TABLE>"

	DisplayErrorMessage = Err.number
	Err.Clear
End Function

Function DisplayErrorMessageForPPC(sTitle, sErrorMessage)
'************************************************************
'Purpose: To display an error message in a box
'Inputs:  sTitle, sErrorMessage
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayErrorMessageForPPC"

	Response.Write "<TABLE BGCOLOR=""#" & S_WIDGET_FRAME_FOR_GUI & """ WIDTH=""98%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""1""><TR><TD>"
		Response.Write "<TABLE BGCOLOR=""#FFFFFF"" WIDTH=""100%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""3""><TR><TD>"
			If Len(sTitle) > 0 Then
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"" COLOR=""#" & S_SECOND_TITLE_FOR_GUI & """><B>"
					Response.Write UCase(sTitle)
					If Len(sErrorMessage) > 0 Then
						Response.Write "<BR /><BR />"
					End If
				Response.Write "</B></FONT>"
			End If
			Response.Write "<FONT FACE=""Arial"" SIZE=""1"" COLOR=""#000000"">"
				Response.Write sErrorMessage
			Response.Write "</FONT>"
		Response.Write "</TD></TR></TABLE>"
	Response.Write "</TD></TR></TABLE>"

	DisplayErrorMessageForPPC = Err.number
	Err.Clear
End Function

Function DisplayErrorMessageAsTooltip(sTitle, sErrorMessage)
'************************************************************
'Purpose: To display an error message in a box
'Inputs:  sTitle, sErrorMessage
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayErrorMessage"

	Response.Write "<TABLE BGCOLOR=""#" & S_WIDGET_FRAME_FOR_GUI & """ WIDTH=""98%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""1""><TR><TD>"
		Response.Write "<TABLE BGCOLOR=""#FFFFFF"" WIDTH=""100%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""1""><TR><TD>"
			Response.Write "<IMG SRC=""Images/IcnExclamationSmall.gif"" WIDTH=""11"" HEIGHT=""10"" ALIGN=""LEFT"" VALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
			If Len(sTitle) > 0 Then
				Response.Write "<FONT FACE=""Verdana"" SIZE=""1"" COLOR=""#" & S_SECOND_TITLE_FOR_GUI & """><B>"
					Response.Write UCase(sTitle)
					If Len(sErrorMessage) > 0 Then
						Response.Write ":&nbsp;"
					End If
				Response.Write "</B></FONT>"
			End If
			Response.Write "<FONT FACE=""Verdana"" SIZE=""1"" COLOR=""#000000"">"
				Response.Write sErrorMessage
			Response.Write "</FONT>"
		Response.Write "</TD></TR></TABLE>"
	Response.Write "</TD></TR></TABLE>"

	DisplayErrorMessageAsTooltip = Err.number
	Err.Clear
End Function

Function DisplayErrorMessageInPlainText(sTitle, sErrorMessage, sNewLine)
'************************************************************
'Purpose: To display an error message in plain text
'Inputs:  sTitle, sErrorMessage, sNewLine
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayErrorMessageInPlainText"

	Response.Write "<FONT FACE=""Arial"" SIZE=""2"">"
		Response.Write Replace(Replace(sTitle, "&nbsp;", " ", 1, -1, vbBinaryCompare), "<BR />", sNewLine, 1, -1, vbBinaryCompare)
		If (Len(sTitle) > 0) And (Len(sErrorMessage) > 0) Then
			Response.Write sNewLine & sNewLine
		End If
		Response.Write Replace(Replace(sErrorMessage, "&nbsp;", " ", 1, -1, vbBinaryCompare), "<BR />", sNewLine, 1, -1, vbBinaryCompare)
	Response.Write "</FONT>"

	DisplayErrorMessageInPlainText = Err.number
	Err.Clear
End Function

Function DisplayFileSize(lSize)
'************************************************************
'Purpose: To display the size of a file in GB, MB, KB or B
'Inputs:  lSize
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayFileSize"

	If lSize >= 1073741824 Then
		DisplayFileSize = CInt(lSize / 1073741824) & " GB"
	ElseIf lSize >= 1048576 Then
		DisplayFileSize = CInt(lSize / 1048576) & " MB"
	ElseIf lSize >= 1024 Then
		DisplayFileSize = CInt(lSize / 1024) & " KB"
	Else
		DisplayFileSize = lSize & " b"
	End If

	Err.Clear
End Function

Function DisplayURLParametersAsHiddenValues(vURL)
'************************************************************
'Purpose: To display the URL parameters and their values as
'         hidden fields
'Inputs:  vURL
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayURLParametersAsHiddenValues"
	Dim sURL
	Dim iPosition
	Dim asParameterValuePairs
	Dim asParameterValue
	Dim iIndex

	sURL = CStr(vURL)
	iPosition = InStr(1, sURL, "?", vbBinaryCompare)
	If iPosition > 0 Then
		sURL = Right(sURL, (Len(sURL) - iPosition))
	End If
	asParameterValuePairs = Split(sURL, "&", -1, vbBinaryCompare)
	For iIndex = 0 To UBound(asParameterValuePairs)
		asParameterValue = Split(asParameterValuePairs(iIndex), "=", -1, vbBinaryCompare)
		If UBound(asParameterValue) = 1 Then
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""" & asParameterValue(0) & """ ID=""" & asParameterValue(0) & "Hdn"" VALUE="""
			If (InStr(1, asParameterValue(1), "%", vbBinaryCompare) > 0) Or (InStr(1, asParameterValue(1), "+", vbBinaryCompare) > 0) Then
				Response.Write UnEncode(asParameterValue(1))
			Else
				Response.Write asParameterValue(1)
			End If
			Response.Write """ />" & vbNewLine
		End If
	Next

	DisplayURLParametersAsHiddenValues = Err.number
	Err.Clear
End Function

Function DoubleSplit(sStringToSplit, asDoubleArray)
'************************************************************
'Purpose: To build a two-dimension array based on the input
'         string
'Inputs:  sStringToSplit
'Outputs: asDoubleArray
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DoubleSplit"
	Dim aTemp()
	Dim aTemp2
	Dim aTemp3()
	Dim iMaxBound
	Dim iIndex
	Dim jIndex

	aTemp2 = Split(sStringToSplit, LIST_SEPARATOR, -1, vbBinaryCompare)
	Redim aTemp3(UBound(aTemp2))
	For iIndex = 0 To UBound(aTemp2)
		aTemp3(iIndex) = Split(aTemp2(iIndex), SECOND_LIST_SEPARATOR, -1, vbBinaryCompare)
	Next
	asDoubleArray = aTemp3
	iMaxBound = 0
	For iIndex = 0 To UBound(asDoubleArray)
		If iMaxBound < UBound(asDoubleArray(iIndex)) Then
			iMaxBound = UBound(asDoubleArray(iIndex))
		End If
	Next
	For iIndex = 0 To UBound(asDoubleArray)
		If iMaxBound > UBound(asDoubleArray(iIndex)) Then
			Redim aTemp(iMaxBound)
			For jIndex = 0 To UBound(asDoubleArray(iIndex))
				aTemp(jIndex) = asDoubleArray(iIndex)(jIndex)
			Next
			For jIndex = (jIndex + 1) To iMaxBound
				aTemp(jIndex) = 0
			Next
			asDoubleArray(iIndex) = aTemp
		End If
	Next

	DoubleSplit = Err.number
	Err.Clear
End Function

Function FormatFloat(dNumber)
'*******************************************************************************
'Purpose: To return a float replaceing "," with "."
'Inputs:  dNumber
'Outputs: A string with the new float value
'*******************************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "FormatFloat"

	FormatFloat = Replace(CStr(dNumber), ",", ".")
	Err.Clear
End Function

Function FormatNumberAsText(dNumber, bAddSufix)
'************************************************************
'Purpose: To transform a number into a text
'Inputs:  dNumber, bAddSufix
'Outputs: The number as text
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "FormatNumberAsText"
	Dim asNumberNames
	Dim asNumber
	Dim sInteger
	Dim sDecimals
	Dim iIndex
	Dim sChar
	Dim iPrevious
	Dim sTemp
	Dim asTemp

	sInteger = ""
	sDecimals = ""
	asNumberNames = Split(";" & _
						  ",uno,dos,tres,cuatro,cinco,seis,siete,ocho,nueve,diez,once,doce,trece,catorce,quince,dieciséis,diecisiete,dieciocho,diecinueve;" & _
						  ",diez,veinte.veinti,treinta,cuarenta,cincuenta,sesenta,setenta,ochenta,noventa;" & _
						  ",cien.ciento,doscientos,trescientos,cuatrocientos,quinientos,seiscientos,setecientos,ochocientos,novecientos;" & _
						  ", mil,dos mil,tres mil,cuatro mil,cinco mil,seis mil,siete mil,ocho mil,nueve mil;" & _
						  ",un millón.millones,un billón.billones,un trillón.trillones,un cuatrillón.cuatrillones,un quintillón.quintillones" & _
	"", ";", -1, vbBinaryCompare)
	For iIndex = 0 To UBound(asNumberNames)
		asNumberNames(iIndex) = Split(asNumberNames(iIndex), ",", -1, vbBinaryCompare)
	Next
	asNumber = Split(FormatNumber(dNumber, 2, True, False, False), ".", -1, vbBinaryCompare)
	If UBound(asNumber) = 1 Then
		If Len(asNumber(1)) > 0 Then
			sDecimals = " " & asNumber(1) & "/100"
		Else
			sDecimals = " 00/100"
		End If
	End If
	For iIndex = 1 To Len(asNumber(0))
		sChar = Mid(asNumber(0), iIndex, Len("0"))
		If IsNumeric(sChar) Then sInteger = sInteger & sChar
	Next
	iPrevious = 0
	FormatNumberAsText = ""
	For iIndex = 1 To Len(sInteger)
		sChar = Mid(sInteger, (Len(sInteger) + 1 - iIndex), Len("0"))
		If Len(sChar) > 0 Then
			Select Case iIndex
				Case 1, 4
					sTemp = ""
					If iIndex = 4 Then sTemp = " mil "
					If ((iIndex = 1) And (Len(sInteger) > 1)) Or ((iIndex = 4) And (Len(sInteger) > 4)) Then
						If CInt(Mid(sInteger, (Len(sInteger) - iIndex), Len("0"))) = 1 Then
							FormatNumberAsText = asNumberNames(1)(CInt(Mid(sInteger, (Len(sInteger) - iIndex), Len("00")))) & sTemp & FormatNumberAsText
							iIndex = iIndex + 1
						Else
							'If iIndex = 4 Then
								FormatNumberAsText = Replace(asNumberNames(1)(CInt(sChar)), "uno", "un") & sTemp & FormatNumberAsText
							'Else
							'	FormatNumberAsText = Replace(asNumberNames(1)(CInt(sChar)), "uno", "") & sTemp & FormatNumberAsText
							'End If
						End If
					Else
						FormatNumberAsText = Replace(asNumberNames(1)(CInt(sChar)), "uno", "") & sTemp & FormatNumberAsText
					End If
				Case 2, 5
					If InStr(1, asNumberNames(2)(CInt(sChar)), ".", vbBinaryCompare) > 0 Then
						asTemp = Split(asNumberNames(2)(CInt(sChar)), ".")
						If iPrevious = 0 Then
							FormatNumberAsText = asTemp(0) & FormatNumberAsText
						Else
							FormatNumberAsText = asTemp(1) & Replace(FormatNumberAsText, "seis", "séis")
						End If
					Else
						If iPrevious = 0 Then
							FormatNumberAsText = asNumberNames(2)(CInt(sChar)) & FormatNumberAsText
						ElseIf CInt(sChar) = 0 Then
							FormatNumberAsText = asNumberNames(2)(CInt(sChar)) & FormatNumberAsText
						Else
							FormatNumberAsText = asNumberNames(2)(CInt(sChar)) & " y " & FormatNumberAsText
						End If
					End If
				Case 3, 6
					If InStr(1, asNumberNames(3)(CInt(sChar)), ".", vbBinaryCompare) > 0 Then
						asTemp = Split(asNumberNames(3)(CInt(sChar)), ".")
						If (CLng(sInteger) Mod 100) = 0 Then
							FormatNumberAsText = asTemp(0) & sTemp
						Else
							FormatNumberAsText = asTemp(1) & " " & FormatNumberAsText
						End If
					Else
						FormatNumberAsText = asNumberNames(3)(CInt(sChar)) & " " & FormatNumberAsText
					End If
				Case 40
					FormatNumberAsText = asNumberNames(4)(CInt(sChar)) & " " & FormatNumberAsText
				Case 50, 60
					If InStr(1, asNumberNames(iIndex - 3)(CInt(sChar)), ".", vbBinaryCompare) > 0 Then
						asTemp = Split(asNumberNames(iIndex - 3)(CInt(sChar)), ".")
						If CInt(Mid(sInteger, (Len(sInteger) + 2 - iIndex), Len("0"))) = 0 Then
							FormatNumberAsText = asTemp(0) & " " & FormatNumberAsText
						Else
							FormatNumberAsText = asTemp(0) & asNumberNames(4)(1) & " " & FormatNumberAsText
						End If
					Else
						If CInt(Mid(sInteger, (Len(sInteger) + 2 - iIndex), Len("0"))) = 0 Then
							FormatNumberAsText = asNumberNames(iIndex - 3)(CInt(sChar)) & " " & FormatNumberAsText
						Else
							FormatNumberAsText = asNumberNames(iIndex - 3)(CInt(sChar)) & asNumberNames(4)(1) & " " & FormatNumberAsText
						End If
					End If
			End Select
			iPrevious = CInt(sChar)
		End If
		If iIndex >= 6 Then Exit For
	Next
	If Len(sInteger) > 6 Then
		sInteger = Left(sInteger, (Len(sInteger) - Len("000000")))
		For iIndex = 1 To Len(sInteger) Step 6
			If CLng(Right(sInteger, Len("000000"))) = 1 Then
				asTemp = Split(asNumberNames(5)(Int(iIndex/6) + 1), ".")
				FormatNumberAsText = asTemp(0) & " " & FormatNumberAsText
			Else
				asTemp = Split(asNumberNames(5)(Int(iIndex/6) + 1), ".")
				FormatNumberAsText = Replace(FormatNumberAsText(Right(sInteger, Len("000000"))), " pesos 00/100 M.N.", "") & " " & asTemp(1) & " " & FormatNumberAsText
			End If
			sInteger = Left(sInteger, (Len(sInteger) - Len("000000")))
		Next
	End If
	If bAddSufix Then
		FormatNumberAsText = FormatNumberAsText & " pesos" & sDecimals & " M.N."
	ElseIf Len(sDecimals) > 0 Then
		FormatNumberAsText = FormatNumberAsText' & " con " & sDecimals
	End If

	Err.Clear
End Function

Function GenerateRandomCharactersSecuence(iCount)
'************************************************************
'Purpose: To generate a secuence of random numbers
'Inputs:  iCount
'Outputs: A secuence of random numbers
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GenerateRandomNumbersSecuence"
	Dim sRandomNumbers
	Dim iIndex
	Dim iRandomNumber
	Randomize

	sRandomNumbers = ""
	For iIndex = 1 To iCount
		iRandomNumber = Int(75 * Rnd) + 47
		If InStr(1, ",58,59,60,61,62,63,64,91,92,93,94,95,96,", "," & iRandomNumber & ",", vbBinaryCompare) = 0 Then
			sRandomNumbers = sRandomNumbers & Chr(iRandomNumber)
		Else
			iIndex = iIndex - 1
		End If
	Next

	GenerateRandomCharactersSecuence = sRandomNumbers
	Err.Clear
End Function

Function GenerateRandomHexadecimalSecuence(iCount)
'************************************************************
'Purpose: To generate a secuence of random numbers
'Inputs:  iCount
'Outputs: A secuence of random numbers
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GenerateRandomNumbersSecuence"
	Dim sRandomNumbers
	Dim iIndex
	Randomize

	sRandomNumbers = ""
	For iIndex = 1 To iCount
		sRandomNumbers = sRandomNumbers & Mid(S_HEXADECIMAL_DIGITS, (Int(16 * Rnd) + 1), 1)
	Next

	GenerateRandomHexadecimalSecuence = sRandomNumbers
	Err.Clear
End Function

Function GenerateRandomNumbersArray(iLimit, iCount, bAllDifferent)
'************************************************************
'Purpose: To generate an array of random numbers
'Inputs:  iLimit, iCount, bAllDifferent
'Outputs: An array of random numbers
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GenerateRandomNumbersArray"
	Dim sRandomNumbers
	Dim iIndex
	Dim iRandomNumber
	Randomize

	If iLimit < iCount Then
		bAllDifferent = False
	End If

	sRandomNumbers = ""
	For iIndex = 1 To iCount
		iRandomNumber = Int(iLimit * Rnd)
		If Not bAllDifferent Or (InStr(1, ("," & sRandomNumbers), ("," & iRandomNumber & ","), vbBinaryCompare) = 0) Then
			sRandomNumbers = sRandomNumbers & iRandomNumber & ","
		Else
			iIndex = iIndex - 1
		End If
	Next
	sRandomNumbers = Left(sRandomNumbers, (Len(sRandomNumbers) - Len(",")))

	GenerateRandomNumbersArray = Split(sRandomNumbers, ",", -1, vbBinaryCompare)
	Err.Clear
End Function

Function GenerateRandomNumbersSecuence(iCount)
'************************************************************
'Purpose: To generate a secuence of random numbers
'Inputs:  iCount
'Outputs: A secuence of random numbers
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GenerateRandomNumbersSecuence"
	Dim sRandomNumbers
	Dim iIndex
	Randomize

	sRandomNumbers = ""
	For iIndex = 1 To iCount
		sRandomNumbers = sRandomNumbers & Int(10 * Rnd)
	Next

	GenerateRandomNumbersSecuence = sRandomNumbers
	Err.Clear
End Function

Function GetASPFileName(sPathInfoServerVariable)
'************************************************************
'Purpose: To get the name of the ASP file (removing the
'         parameters and the web server name from the url)
'Inputs:  sPathInfoServerVariable
'Outputs: The name of the ASP file
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetASPFileName"
	Dim sEndOfPageName

	If Len(sPathInfoServerVariable) = 0 Then
		sPathInfoServerVariable = Request.ServerVariables("PATH_INFO")
	End If
	GetASPFileName = Right(sPathInfoServerVariable, Len(sPathInfoServerVariable) - (InStrRev(sPathInfoServerVariable, "/")))

	sEndOfPageName = InStr(1, GetASPFileName, "?", vbBinaryCompare)
	If sEndOfPageName > 0 Then
		GetASPFileName = Left(GetASPFileName, (sEndOfPageName - Len("?")))
	End If

	Err.Clear
End Function

Function GetDomainURL(sPathInfoServerVariable)
'************************************************************
'Purpose: To get the name of the ASP file (removing the
'         parameters and the web server name from the url)
'Inputs:  sPathInfoServerVariable
'Outputs: The name of the ASP file
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetDomainURL"
	Dim sEndOfPageName

	If Len(sPathInfoServerVariable) = 0 Then
		sPathInfoServerVariable = Request.ServerVariables("PATH_INFO")
	End If
	GetDomainURL = S_HTTP & Request.ServerVariables("SERVER_NAME") & Left(sPathInfoServerVariable, InStrRev(sPathInfoServerVariable, "/"))
	
	Err.Clear
End Function

Function GetParameterFromURLString(sURL, sParameterToGet)
'************************************************************
'Purpose: To get a parameter from a URL
'Inputs:  vURL, sParameterToGet
'Outputs: The value of the given parameter
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetParameterFromURLString"
	Dim asURL
	Dim sValue
	Dim iIndex

	asURL = Split(sURL, "&", -1, vbBinaryCompare)
	For iIndex = 0 To UBound(asURL)
		If InStr(1, (asURL(iIndex)), (sParameterToGet & "="), vbTextCompare) = 1 Then
			sValue = sValue & Right(asURL(iIndex), (Len(asURL(iIndex)) - Len((sParameterToGet & "=")))) & ","
		End If
	Next
	If Len(sValue) > 0 Then sValue = Left(sValue, (Len(sValue) - Len(",")))

	GetParameterFromURLString = sValue
	Err.Clear
End Function

Function HexToLng(sHexValue)
'************************************************************
'Purpose: To get the long value of an hexadecimal value
'Inputs:  sHexValue
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "HexToLng"
	Dim lLngValue
	Dim iIndex
	Dim iPower

	lLngValue = 0
	iPower = 1
	For iIndex = Len(sHexValue) To iIndex > 1 Step -1
		lLngValue = lLngValue + ((InStr(1, S_HEXADECIMAL_DIGITS, Mid(sHexValue, iIndex, 1), vbTextCompare) - 1) * iPower)
		iPower = iPower * 16
	Next

	HexToLng = lLngValue
	Err.Clear
End Function

Function JoinLists(sList1, sList2, sSeparator)
'************************************************************
'Purpose: To join two lists using the given separator
'Inputs:  sList1, sList2, sSeparator
'Outputs: A list
'************************************************************
	If Len(sList1) = 0 Then
		JoinLists = sList2
	ElseIf Len(sList2) = 0 Then
		JoinLists = sList1
	ElseIf StrComp((Right(sList1, Len(sSeparator))), sSeparator, vbBinaryCompare) = 0 Then
		JoinLists = sList1 & sList2
	ElseIf StrComp((Left(sList2, Len(sSeparator))), sSeparator, vbBinaryCompare) = 0 Then
		JoinLists = sList1 & sList2
	Else
		JoinLists = sList1 & sSeparator & sList2
	End If
End Function

Function RemoveHTMLFromString(sStringToChange)
'************************************************************
'Purpose: To remove any HTML tag from the String
'Inputs:  sStringToChange
'Outputs: The new string
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveHTMLFromString"
	Dim sTextToClean
	Dim iStart
	Dim iEnd

	sTextToClean = sStringToChange
	iStart = InStr(1, sTextToClean, "<", vbBinaryCompare)
	iEnd = InStr(iStart, sTextToClean, ">", vbBinaryCompare)

	If (iStart * iEnd) > 0 Then
		sTextToClean = Left(sTextToClean, (iStart - Len("<"))) + Right(sTextToClean, (Len(sTextToClean) - iEnd))
		sTextToClean = RemoveHTMLFromString(sTextToClean)
	End If

	RemoveHTMLFromString = sTextToClean
	Err.Clear
End Function

Function RemoveEmptyParametersFromURLString(vURL)
'************************************************************
'Purpose: To remove all the empty parameters from a URL
'Inputs:  vURL
'Outputs: The URL string without the empty parameters
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveEmptyParametersFromURLString"
	Dim sTempURL
	Dim aParameters
	Dim aParameterValue
	Dim iIndex

	sTempURL = CStr(vURL)
	If InStr(1, sTempURL, "&", vbBinaryCompare) = 1 Then
		sTempURL = Right(sTempURL, Len(sTempURL) - Len("&"))
	End If
	If InStr(1, sTempURL, "?", vbBinaryCompare) = 1 Then
		sTempURL = Right(sTempURL, Len(sTempURL) - Len("?"))
	End If

	aParameters = Split(sTempURL, "&", -1, vbBinaryCompare)
	For iIndex = 0 To UBound(aParameters)
		aParameterValue = Split(aParameters(iIndex), "=", -1, vbBinaryCompare)
		If Len(aParameterValue(1)) = 0 Then
			sTempURL = RemoveParameterFromURLString(sTempURL, aParameterValue(0))
		Else
			If StrComp(Right(aParameterValue(0), Len("Year")), "Year", vbTextCompare) = 0 Then
				If CInt(aParameterValue(1)) <= 0 Then sTempURL = RemoveParameterFromURLString(sTempURL, aParameterValue(0))
			End If
			If StrComp(Right(aParameterValue(0), Len("Month")), "Month", vbTextCompare) = 0 Then
				If CInt(aParameterValue(1)) <= 0 Then sTempURL = RemoveParameterFromURLString(sTempURL, aParameterValue(0))
			End If
			If StrComp(Right(aParameterValue(0), Len("Day")), "Day", vbTextCompare) = 0 Then
				If CInt(aParameterValue(1)) <= 0 Then sTempURL = RemoveParameterFromURLString(sTempURL, aParameterValue(0))
			End If
		End If
	Next

	RemoveEmptyParametersFromURLString = sTempURL
	Err.Clear
End Function

Function RemoveParameterFromURLString(vURL, sParameterToRemove)
'************************************************************
'Purpose: To remove a parameter from a URL
'Inputs:  vURL, sParameterToRemove
'Outputs: The URL string without the given parameter and its value
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveParameterFromURLString"
	Dim iInitialPos
	Dim iFinalPos
	Dim sTempURL

	sTempURL = CStr(vURL)
	If InStr(1, sTempURL, "&", vbBinaryCompare) = 1 Then
		sTempURL = Right(sTempURL, Len(sTempURL) - Len("&"))
	End If
	iInitialPos = InStr(1, sTempURL, "&" & sParameterToRemove & "=", vbTextCompare)
	If iInitialPos = 0 Then
		iInitialPos = InStr(1, sTempURL, "?" & sParameterToRemove & "=", vbTextCompare)
	End If
	If iInitialPos = 0 Then
		iInitialPos = InStr(1, sTempURL, sParameterToRemove & "=", vbTextCompare)
		If iInitialPos <> 1 Then
			iInitialPos = 0
		End If
	End If
	If iInitialPos > 0 Then
		If iInitialPos > 1 Then
			iInitialPos = iInitialPos + Len("&")
		End If
		iFinalPos = InStr(iInitialPos, sTempURL, "&", vbTextCompare)
		If iFinalPos > 0 Then
			iFinalPos = iFinalPos + Len("&")
		End If
		If iInitialPos = 1 Then
			If iFinalPos > 0 Then
				sTempURL = Mid(sTempURL, iFinalPos)
			Else
				sTempURL = ""
			End If
		Else
			sTempURL = Mid(sTempURL, 1, iInitialPos - 1)
			If iFinalPos > 0 Then
				sTempURL = sTempURL & Mid(CStr(vURL), iFinalPos)
			End If
		End If
	End If
	If StrComp(Right(sTempURL, Len("&")), "&", vbBinaryCompare) = 0 Then
		sTempURL = Left(sTempURL, Len(sTempURL) - Len("&"))
	End If
	If ((InStr(1, sTempURL, "&") = 0) And (InStr(1, sTempURL, "=") = 0) And (InStr(1, sTempURL, "?") = 0)) And Len(sTempURL) > 0 Then
		sTempURL = sTempURL & "?"
	End If


	iInitialPos = InStr(1, sTempURL, "&" & sParameterToRemove & "=", vbTextCompare)
	If iInitialPos = 0 Then
		iInitialPos = InStr(1, sTempURL, "?" & sParameterToRemove & "=", vbTextCompare)
	End If
	If iInitialPos = 0 Then
		iInitialPos = InStr(1, sTempURL, sParameterToRemove & "=", vbTextCompare)
		If iInitialPos <> 1 Then
			iInitialPos = 0
		End If
	End If
	If iInitialPos > 0 Then
		sTempURL = RemoveParameterFromURLString(sTempURL, sParameterToRemove)
	End If
	RemoveParameterFromURLString = sTempURL
	Err.Clear
End Function

Function RepeatString(sStringToDisplay, lCount)
'************************************************************
'Purpose: To display lCount times the value of sStringToDisplay
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RepeatString"
	Dim iIndex

	RepeatString = ""
	For iIndex = 1 To lCount
		RepeatString = RepeatString & sStringToDisplay
	Next
End Function

Function ReplaceValueInURLString(vURL, sParameterToChange, sNewValue)
'************************************************************
'Purpose: To replace the value of a parameter in a URL string
'         with the given value
'Inputs:  vURL, sParameterToChange, sNewValue
'Outputs: The URL string with the new value for the given parameter
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ReplaceValueInURLString"
	Dim sURL

	sURL = RemoveParameterFromURLString(vURL, sParameterToChange)
	If Len(sURL) > 0 Then
		If StrComp(Right(sURL, Len("?")), "?", vbBinaryCompare) = 0 Then
			sURL = sURL & sParameterToChange & "=" & Server.URLEncode(sNewValue)
		Else
			sURL = sURL & "&" & sParameterToChange & "=" & Server.URLEncode(sNewValue)
		End If
	Else
		sURL = sParameterToChange & "=" & Server.URLEncode(sNewValue)
	End If

	ReplaceValueInURLString = sURL
	Err.Clear
End Function

Function SizeText(sText, cFill, iSize, iOrder)
'************************************************************
'Purpose: To transform the given text cropping or filling to
'         fulfill the given size
'Inputs:  sText, cFill, iSize
'Outputs: The text in the given size
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SizeText"
	Dim iIndex

	SizeText = sText
	If Len(SizeText) < iSize Then
		If StrComp(cFill, " ", vbBinaryCompare) = 0 Then
			If iOrder = 1 Then
				SizeText = SizeText & "                                                                                                                        "
			Else
				SizeText = "                                                                                                                        " & SizeText
			End If
		Else
			If iOrder = 1 Then
				SizeText = SizeText & Replace("                                                                                                                        ", " ", cFill, 1, -1, vbBinaryCompare)
			Else
				SizeText = Replace("                                                                                                                        ", " ", cFill, 1, -1, vbBinaryCompare) & SizeText
			End If
		End If
		For iIndex = (Len(SizeText) + 1) To iSize
			If iOrder = 1 Then
				SizeText = SizeText & cFill
			Else
				SizeText = cFill & SizeText
			End If
		Next
	End If
	If iOrder = 1 Then
		SizeText = Left(SizeText, iSize)
	Else
		SizeText = Right(SizeText, iSize)
	End If
	Err.Clear
End Function

Function SplitString(sSource, iSections, iSize)
'************************************************************
'Purpose: To get the words in the string and  return them in
'         an array.
'Inputs:  sSource, iSections, iSize
'Outputs: An array with elements of iSize characters length
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SplitString"
	Dim asTemp
	Dim sTemp
	Dim asTarget
	Dim sTarget
	Dim iIndex

	asTemp = Split(Replace(sSource, "  ", " "), " ")
	sTemp = ""
	For iIndex = 0 To UBound(asTemp)
		If Len(sTemp) = 0 Then
			sTemp = sTemp & asTemp(iIndex)
		Else
			If Len(sTemp & " " & asTemp(iIndex)) > iSize Then
				sTarget = sTarget & sTemp & LIST_SEPARATOR
				sTemp = asTemp(iIndex)
			Else
				sTemp = sTemp & " " & asTemp(iIndex)
			End If
		End If
	Next
	sTarget = sTarget & sTemp
	asTarget = Split(sTarget, LIST_SEPARATOR)
	For iIndex = 0 To UBound(asTarget)
		asTarget(iIndex) = Left(asTarget(iIndex), iSize)
	Next

	SplitString = asTarget
	Err.Clear
End Function

Function TransformLongIntoBinaryPowerList(lLong)
'************************************************************
'Purpose: To transform a long into a binary power list
'Inputs:  lLong
'Outputs: A coma separated list with the binary powers
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "TransformLongIntoBinaryPowerList"
	Dim lTemp
	Dim sList
	Dim iIndex

	If lLong = -1 Then
		sList = "1,2,4,8,16,32,64,128,256,512,1024,2048,4096,8192,16384,32768,65536,131072,262144,524288,1048576,2097152,4194304,8388608,16777216,33554432,67108864,134217728,268435456,536870912,1073741824,2147483648,"
	Else
		iIndex = 0
		sList = "0,"
		lTemp = 1
		Do While Not (lTemp > lLong)
			If (lTemp And lLong) = lTemp Then sList = sList & lTemp & ","
			iIndex = iIndex + 1
			lTemp = (2^iIndex)
			If Err.number <> 0 Then Exit Do
		Loop
	End If
	If Len(sList) > 0 Then sList = Left(sList, (Len(sList) - Len(",")))

	TransformLongIntoBinaryPowerList = sList
	Err.Clear
End Function

Function UnEncode(sStringToChange)
'************************************************************
'Purpose: To transform the %xx codes into chars
'Inputs:  sStringToChange
'Outputs: A String unencoded
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UnEncode"
	Dim iStartPos
	Dim sChar
	Dim bDone

	iStartPos = 1
	bDone = False
	UnEncode = Replace(sStringToChange, "+", " ")
	Do While Not bDone
		iStartPos = InStr(iStartPos, UnEncode, "%", vbBinaryCompare)
		If (iStartPos > 0) And (iStartPos <= (Len(UnEncode) - 2)) Then
			sChar = Replace(Mid(UnEncode, iStartPos, Len("%00")), "%", "")
			If Len(sChar) > 0 Then
				UnEncode = Left(UnEncode, (iStartPos-1)) & Chr(HexToLng(sChar)) & Right(UnEncode, (Len(UnEncode) - (iStartPos + 2)))
			End If
			iStartPos = iStartPos + 1
		Else
			bDone = True
		End If
		If Err.number <> 0 Then Exit Do
	Loop

	Err.Clear
End Function
%>