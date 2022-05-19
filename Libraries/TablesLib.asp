<%
Function CheckCellsArray(lColumns, sDefaultValue, asCells)
'************************************************************
'Purpose: To check the cells in the array are set before
'         using them
'Inputs:  lColumns, sDefaultValue, asCells
'Outputs: asCells
'************************************************************
	On Error Resume Next
	Dim iIndex

	If Not IsArray(asCellWidths) Then
		Redim asCells(lColumns)
		For iIndex = 0 To lColumns
			asCells(iIndex) = sDefaultValue
		Next
	End If
	Err.Clear
End Function

Function CheckColorsArray(asTableColors)
'************************************************************
'Purpose: To check the colors in the array are set before
'         using them
'Inputs:  asTableColors
'Outputs: asTableColors
'************************************************************
	On Error Resume Next

	If Not IsArray(asTableColors) Then
		asTableColors = Split(",,,,", ",", -1, vbBinaryCompare)
	End If
	Redim Preserve asTableColors(4)
	If Len(asTableColors(0)) = 0 Then asTableColors(0) = S_LIGHT_BGCOLOR_FOR_GUI
	If Len(asTableColors(1)) = 0 Then asTableColors(1) = S_BGCOLOR_FOR_GUI
	If Len(asTableColors(2)) = 0 Then asTableColors(2) = S_DARK_BGCOLOR_FOR_GUI
	If Len(asTableColors(3)) = 0 Then asTableColors(3) = S_BGCOLOR_FOR_GUI
	If Len(asTableColors(4)) = 0 Then asTableColors(4) = S_TABLE_TITLE_FOR_GUI
	Select Case CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION))
		Case 2
			asTableColors(0) = S_BGCOLOR_FOR_GUI
			asTableColors(1) = "FFFFFF"
			asTableColors(2) = ""
			asTableColors(3) = S_BGCOLOR_FOR_GUI
			asTableColors(4) = S_BGCOLOR_FOR_GUI
	End Select
	Err.Clear
End Function

Function CheckColumnsTitlesArray(oRecordset, asColumnsTitles)
'************************************************************
'Purpose: To check the columns titles in the array are set
'         before using them
'Inputs:  oRecordset, asColumnsTitles
'Outputs: asColumnsTitles
'************************************************************
	On Error Resume Next
	Dim sTemp
	Dim iIndex

	sTemp = ""
	If Not IsArray(asColumnsTitles) Then
		For iIndex = 0 To oRecordset.Fields.Count
			sTemp = sTemp & oRecordset.Fields(iIndex).Name & ","
			If Err.number <> 0 Then Exit For
		Next
		If Len(sTemp) > 0 Then sTemp = Left(sTemp, (Len(sTemp) - Len(",")))
		asColumnsTitles = Split(sTemp, ",", -1, vbBinaryCompare)
	End If

	Err.Clear
End Function

Function DisplayRecordsetAsTable(oRecordset, asColumnsTitles, asCellAlignments, asCellWidths, asTableColors, bPlainText, sErrorDescription)
'************************************************************
'Purpose: To display the information from the database in a table
'Inputs:  oRecordset, asColumnsTitles, asCellAlignments, asCellWidths, asTableColors, bPlainText
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayRecordsetAsTable"
	Dim asColumnsTitles2
	Dim sColumnsTitles2
	Dim asRowContents
	Dim sRowContents
	Dim sFieldValue
	Dim sTempRowContents
	Dim iDiference
	Dim iTableWidth 
	Dim iCount
	Dim lErrorNumber
	Dim iRowCounter

	Call CheckColumnsTitlesArray(oRecordset, asColumnsTitles)
	Call CheckCellsArray(UBound(asColumnsTitles), "", asCellAlignments)
	Call CheckCellsArray(UBound(asColumnsTitles), 1, asCellWidths)
	iCount = 0
	iTableWidth = 0
	For iCount = 0 To (UBound(asCellWidths))
		iTableWidth = iTableWidth +	CInt(asCellWidths(iCount))
	Next
	Call CheckColorsArray(asTableColors)
	Response.Write "<TABLE WIDTH=""" & iTableWidth & """ BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
		iDiference = oRecordset.Fields.Count - (UBound(asColumnsTitles) + 1)
		'If UBound(asColumnsTitles) > 0 Then
			sColumnsTitles2 = Join(asColumnsTitles, ",")
			For iCount = 0 To iDiference
				sColumnsTitles2 = sColumnsTitles2 & "," & "&nbsp;"
			Next 
			asColumnsTitles2 = Split(sColumnsTitles2, ",", -1, vbBinaryCompare)
			If bPlainText Then
				lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles2, False, sErrorDescription)
			Else
				lErrorNumber = DisplayTableHeader3D(asColumnsTitles2, asCellWidths, asTableColors, sErrorDescription)
			End If
		'End If
		If Not oRecordset.EOF Then
			iDiference = (UBound(asColumnsTitles) + 1) - oRecordset.Fields.Count
			sTempRowContents = ""
			For iCount = 0 To iDiference
				sTempRowContents = sTempRowContents & "&nbsp;" & TABLE_SEPARATOR
			Next

			iRowCounter = 0
			Do While Not oRecordset.EOF
				sRowContents = ""
				For iCount=0 To oRecordset.Fields.Count-1
					sFieldValue = ""
					sFieldValue = oRecordset.Fields(iCount).Value
					Err.Clear
					If IsNull(sFieldValue) Or IsEmpty(sFieldValue) Then
						sRowContents = sRowContents & "&nbsp;" & TABLE_SEPARATOR
					Else
						sRowContents = sRowContents & CleanStringForHTML(sFieldValue) & TABLE_SEPARATOR
					End If
				Next
				sRowContents = sRowContents & sTempRowContents
				sRowContents = Left(sRowContents, (Len(sRowContents) - Len(TABLE_SEPARATOR)))
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bPlainText Then
					lErrorNumber = DisplayTableRowText(asRowContents, False, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
				iRowCounter = iRowCounter + 1
			Loop
		End If
	Response.Write "</TABLE>" & vbNewLine

	DisplayRecordsetAsTable = lErrorNumber
	Err.Clear
End Function

Function DisplayLine(asColumnsTitles, sLineColor, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display a line
'Inputs:  asColumnsTitles, sLineColor, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayLine"
	Dim iColSpan
	Dim iPos
	Dim iSpan
	Dim iCount

	If bForExport Then
		Response.Write "<TR><TD HEIGHT=""1""></TD></TR>" & vbNewLine
	Else
		iColSpan = 0
		For iCount = 0 To UBound(asColumnsTitles)
			If InStr(1, asColumnsTitles(iCount), "<SPAN", vbBinaryCompare) = 1 Then
				iPos = InStr(1, asColumnsTitles(iCount), "COLS=""", vbBinaryCompare) + Len("COLS=""")
				iSpan = CInt(Mid(asColumnsTitles(iCount), iPos, (InStr(iPos, asColumnsTitles(iCount), """", vbBinaryCompare) - iPos)))
				iColSpan = iColSpan + iSpan
			Else
				iColSpan = iColSpan + 1
			End If
		Next
		iColSpan = ((iColSpan + 1) * 5)
		If Len(sLineColor) = 0 Then sLineColor = S_BGCOLOR_FOR_GUI
		Response.Write "<TR><TD BGCOLOR=""#" & sLineColor & """ COLSPAN=""" & iColSpan & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR>" & vbNewLine
	End If

	DisplayLine = lErrorNumber
	Err.Clear
End Function

Function DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
'************************************************************
'Purpose: To display the table header using the 3D style
'Inputs:  asColumnsTitles, asCellWidths, asTableColors
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayTableHeader3D"
	Dim sWidth
	Dim iColSpan
	Dim iPos
	Dim iSpan
	Dim iCount
	Dim iAddCount
	Dim lErrorNumber
	
	Call CheckColorsArray(asTableColors)
	iColSpan = 0
	For iCount = 0 To UBound(asColumnsTitles)
		If InStr(1, asColumnsTitles(iCount), "<SPAN", vbBinaryCompare) = 1 Then
			iPos = InStr(1, asColumnsTitles(iCount), "COLS=""", vbBinaryCompare) + Len("COLS=""")
			iSpan = CInt(Mid(asColumnsTitles(iCount), iPos, (InStr(iPos, asColumnsTitles(iCount), """", vbBinaryCompare) - iPos)))
			iColSpan = iColSpan + iSpan
		Else
			iColSpan = iColSpan + 1
		End If
	Next
	iColSpan = ((iColSpan) * 5)
	Response.Write "<TR><TD BGCOLOR=""#" & asTableColors(0) & """ COLSPAN=""" & iColSpan & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR>"
	Response.Write "<TR BGCOLOR=""#" & asTableColors(1) & """>"
		iAddCount = 0
	    For iCount = 0 To UBound(asColumnsTitles)
			sWidth = ""
			If Len(asCellWidths(iCount + iAddCount)) = 0 Then sWidth = "WIDTH=""" & asCellWidths(iCount + iAddCount) & """"
			Response.Write "<TD BGCOLOR=""" & asTableColors(0) & """ WIDTH=""1""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
			Response.Write "<TD>&nbsp;&nbsp;</TD>"
			Response.Write "<TD " & sWidth & " ALIGN=""CENTER"""
				If InStr(1, asColumnsTitles(iCount), "<SPAN", vbBinaryCompare) = 1 Then
					iPos = InStr(1, asColumnsTitles(iCount), "COLS=""", vbBinaryCompare) + Len("COLS=""")
					iSpan = (CInt(Mid(asColumnsTitles(iCount), iPos, (InStr(iPos, asColumnsTitles(iCount), """", vbBinaryCompare) - iPos))) - 1)
					iAddCount = iAddCount + iSpan
					iSpan = iSpan * 5 + 1
					Response.Write " COLSPAN=""" & iSpan & """"
				End If
			Response.Write "><NOBR><FONT FACE=""Arial"" SIZE=""2"" COLOR=""#" & asTableColors(4) & """><B>" & asColumnsTitles(iCount) & "</B></FONT></NOBR></TD>"
			Response.Write "<TD>&nbsp;&nbsp;</TD>"
			Response.Write "<TD BGCOLOR=""" & asTableColors(2) & """ WIDTH=""1""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>" & vbNewLine
		Next
		Response.Write "<TD BGCOLOR=""#FFFFFF""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>" & vbNewLine
	Response.Write "</TR>" & vbNewLine
	Response.Write "<TR><TD BGCOLOR=""#" & asTableColors(2) & """ COLSPAN=""" & iColSpan & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR>" & vbNewLine

	DisplayTableHeader3D = lErrorNumber
	Err.Clear
End Function

Function DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
'************************************************************
'Purpose: To display the table header using the plain style
'Inputs:  asColumnsTitles, asCellWidths, asTableColors
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayTableHeaderPlain"
	Dim sWidth
	Dim iPos
	Dim iSpan
	Dim iColSpan
	Dim iCount
	Dim iAddCount
	Dim lErrorNumber
	
	Call CheckColorsArray(asTableColors)
	iColSpan = 0
	For iCount = 0 To UBound(asColumnsTitles)
		If InStr(1, asColumnsTitles(iCount), "<SPAN", vbBinaryCompare) = 1 Then
			iPos = InStr(1, asColumnsTitles(iCount), "COLS=""", vbBinaryCompare) + Len("COLS=""")
			iSpan = CInt(Mid(asColumnsTitles(iCount), iPos, (InStr(iPos, asColumnsTitles(iCount), """", vbBinaryCompare) - iPos)))
			iColSpan = iColSpan + iSpan
		Else
			iColSpan = iColSpan + 1
		End If
	Next
	iColSpan = ((iColSpan) * 5)
	Response.Write "<TR><TD BGCOLOR=""#" & asTableColors(0) & """ COLSPAN=""" & iColSpan & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR>"
	Response.Write "<TR BGCOLOR=""#" & asTableColors(1) & """>"
		iAddCount = 0
	    For iCount = 0 To UBound(asColumnsTitles)
			sWidth = ""
			If Len(asCellWidths(iCount + iAddCount)) = 0 Then sWidth = "WIDTH=""" & asCellWidths(iCount + iAddCount) & """"
			Response.Write "<TD BGCOLOR=""" & asTableColors(0) & """ WIDTH=""1""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
			Response.Write "<TD>&nbsp;&nbsp;</TD>"
			Response.Write "<TD " & sWidth & " ALIGN=""CENTER"""
				If InStr(1, asColumnsTitles(iCount), "<SPAN", vbBinaryCompare) = 1 Then
					iPos = InStr(1, asColumnsTitles(iCount), "COLS=""", vbBinaryCompare) + Len("COLS=""")
					iSpan = (CInt(Mid(asColumnsTitles(iCount), iPos, (InStr(iPos, asColumnsTitles(iCount), """", vbBinaryCompare) - iPos))) - 1)
					iAddCount = iAddCount + iSpan
					iSpan = iSpan * 5 + 1
					Response.Write " COLSPAN=""" & iSpan & """"
				End If
			Response.Write "><NOBR><FONT FACE=""Arial"" SIZE=""2"" COLOR=""#" & asTableColors(4) & """><B>" & asColumnsTitles(iCount) & "</B></FONT></NOBR></TD>"
			Response.Write "<TD>&nbsp;&nbsp;</TD>"
			Response.Write "<TD BGCOLOR="""
				If iCount <> UBound(asColumnsTitles) Then
					Response.Write asTableColors(1)
				Else
					Response.Write asTableColors(0)
				End If
			Response.Write """ WIDTH=""1""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>" & vbNewLine
		Next
		Response.Write "<TD BGCOLOR=""#FFFFFF""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>" & vbNewLine
	Response.Write "</TR>" & vbNewLine
	Response.Write "<TR><TD BGCOLOR=""#" & asTableColors(0) & """ COLSPAN=""" & iColSpan & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR>" & vbNewLine

	DisplayTableHeaderPlain = lErrorNumber
	Err.Clear
End Function

Function DisplayTableHeaderPlainText(asColumnsTitles, bUseHTML, sErrorDescription)
'************************************************************
'Purpose: To display the table header using plain text
'Inputs:  asColumnsTitles, bUseHTML
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayTableHeaderPlainText"
	Dim iPos
	Dim iSpan
	Dim iCount
	Dim sRowBegin
	Dim sRowEnd
	Dim sCellBegin
	Dim sCellSpanBegin
	Dim sCellEnd
	Dim sFontBegin
	Dim sFontEnd
	Dim sBoldBegin
	Dim sBoldEnd
	Dim lErrorNumber
	
	If bUseHTML Then
		sRowBegin = "<TR>"
		sRowEnd = "</TR>"
		sCellBegin = "<TD ALIGN=""CENTER"">"
		sCellSpanBegin = "<TD ALIGN=""CENTER"" COLSPAN=""xxx"">"
		sCellEnd = "</TD>"
		sFontBegin = "<FONT FACE=""Arial"" SIZE=""2"">"
		sFontEnd = "</FONT>"
		sBoldBegin = "<B>"
		sBoldEnd = "</B>"
	Else
		sRowEnd = vbNewLine
		sCellEnd = vbTab
	End If
	Response.Write sRowBegin
	    For iCount = 0 To UBound(asColumnsTitles)
			If InStr(1, asColumnsTitles(iCount), "<SPAN", vbBinaryCompare) = 1 Then
				iPos = InStr(1, asColumnsTitles(iCount), "COLS=""", vbBinaryCompare) + Len("COLS=""")
				iSpan = CInt(Mid(asColumnsTitles(iCount), iPos, (InStr(iPos, asColumnsTitles(iCount), """", vbBinaryCompare) - iPos)))
				Response.Write Replace(sCellSpanBegin, "xxx", iSpan) & sFontBegin & sBoldBegin & Mid(asColumnsTitles(iCount), (InStr(iPos, asColumnsTitles(iCount), ">", vbBinaryCompare) + Len(">"))) & sBoldEnd & sFontEnd & sCellEnd
			Else
				Response.Write sCellBegin & sFontBegin & sBoldBegin & asColumnsTitles(iCount) & sBoldEnd & sFontEnd & sCellEnd
			End If
		Next
	Response.Write sRowEnd

	DisplayTableHeaderPlainText = lErrorNumber
	Err.Clear
End Function

Function GetTableHeaderPlainText(asColumnsTitles, bUseHTML, sErrorDescription)
'************************************************************
'Purpose: To display the table header using plain text
'Inputs:  asColumnsTitles, bUseHTML
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetTableHeaderPlainText"
	Dim iPos
	Dim iSpan
	Dim iCount
	Dim sRowBegin
	Dim sRowEnd
	Dim sCellBegin
	Dim sCellSpanBegin
	Dim sCellEnd
	Dim sFontBegin
	Dim sFontEnd
	Dim sBoldBegin
	Dim sBoldEnd
	Dim lErrorNumber
	
	If bUseHTML Then
		sRowBegin = "<TR>"
		sRowEnd = "</TR>"
		sCellBegin = "<TD ALIGN=""CENTER"">"
		sCellSpanBegin = "<TD ALIGN=""CENTER"" COLSPAN=""xxx"">"
		sCellEnd = "</TD>"
		sFontBegin = "<FONT FACE=""Arial"" SIZE=""2"">"
		sFontEnd = "</FONT>"
		sBoldBegin = "<B>"
		sBoldEnd = "</B>"
	Else
		sRowEnd = vbNewLine
		sCellEnd = vbTab
	End If
	GetTableHeaderPlainText = sRowBegin
	    For iCount = 0 To UBound(asColumnsTitles)
			If InStr(1, asColumnsTitles(iCount), "<SPAN", vbBinaryCompare) = 1 Then
				iPos = InStr(1, asColumnsTitles(iCount), "COLS=""", vbBinaryCompare) + Len("COLS=""")
				iSpan = CInt(Mid(asColumnsTitles(iCount), iPos, (InStr(iPos, asColumnsTitles(iCount), """", vbBinaryCompare) - iPos)))
				GetTableHeaderPlainText = GetTableHeaderPlainText & Replace(sCellSpanBegin, "xxx", iSpan) & sFontBegin & sBoldBegin & Mid(asColumnsTitles(iCount), (InStr(iPos, asColumnsTitles(iCount), ">", vbBinaryCompare) + Len(">"))) & sBoldEnd & sFontEnd & sCellEnd
			Else
				GetTableHeaderPlainText = GetTableHeaderPlainText & sCellBegin & sFontBegin & sBoldBegin & asColumnsTitles(iCount) & sBoldEnd & sFontEnd & sCellEnd
			End If
		Next
	GetTableHeaderPlainText = GetTableHeaderPlainText & sRowEnd

	Err.Clear
End Function

Function GetTableHeaderPlainTextWidth(asColumnsTitles, asCellWidths, bUseHTML, iFontSize, sErrorDescription)
'************************************************************
'Purpose: To display the table header using plain text
'Inputs:  asColumnsTitles, bUseHTML
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetTableHeaderPlainTextWidth"
	Dim iPos
	Dim iSpan
	Dim iCount
	Dim sRowBegin
	Dim sRowEnd
	Dim sCellBegin
    Dim sCellWidthBegin
	Dim sCellSpanBegin
	Dim sCellEnd
	Dim sFontBegin
	Dim sFontEnd
	Dim sBoldBegin
	Dim sBoldEnd
	Dim lErrorNumber
    Dim sWidth
	
	If bUseHTML Then
        sWidth = ""
		sRowBegin = "<TR>"
		sRowEnd = "</TR>"
		sCellBegin = "<TD ALIGN=""CENTER"">"
		sCellSpanBegin = "<TD ALIGN=""CENTER"" COLSPAN=""xxx"">"
        sCellWidthBegin = "<TD ALIGN=""CENTER"" YYY>"
		sCellEnd = "</TD>"
		sFontBegin = "<FONT FACE=""Arial"" SIZE=""" & iFontSize & """>"
		sFontEnd = "</FONT>"
		sBoldBegin = "<B>"
		sBoldEnd = "</B>"
	Else
		sRowEnd = vbNewLine
		sCellEnd = vbTab
	End If
	GetTableHeaderPlainTextWidth = sRowBegin
	    For iCount = 0 To UBound(asColumnsTitles)
            If Len(asCellWidths(iCount)) > 0 Then sWidth = "WIDTH=""" & asCellWidths(iCount) & """"
			If InStr(1, asColumnsTitles(iCount), "<SPAN", vbBinaryCompare) = 1 Then
				iPos = InStr(1, asColumnsTitles(iCount), "COLS=""", vbBinaryCompare) + Len("COLS=""")
				iSpan = CInt(Mid(asColumnsTitles(iCount), iPos, (InStr(iPos, asColumnsTitles(iCount), """", vbBinaryCompare) - iPos)))
				GetTableHeaderPlainTextWidth = GetTableHeaderPlainTextWidth & Replace(sCellSpanBegin, "xxx", iSpan) & sFontBegin & sBoldBegin & Mid(asColumnsTitles(iCount), (InStr(iPos, asColumnsTitles(iCount), ">", vbBinaryCompare) + Len(">"))) & sBoldEnd & sFontEnd & sCellEnd
			Else
				GetTableHeaderPlainTextWidth = GetTableHeaderPlainTextWidth & Replace(sCellWidthBegin, "YYY", sWidth) & sFontBegin & sBoldBegin & asColumnsTitles(iCount) & sBoldEnd & sFontEnd & sCellEnd
			End If
		Next
	GetTableHeaderPlainTextWidth = GetTableHeaderPlainTextWidth & sRowEnd

	Err.Clear
End Function

Function DisplayRTFRow(asRowContents, asCellAlignments, asCellWidths, bBorder, sFileName, sErrorDescription)
'************************************************************
'Purpose: To display a table row using RTF comands
'Inputs:  asRowContents, asCellAlignments, asCellWidths, bBorder, sFileName
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayRTFRow"
	Dim sWidth
	Dim sAlign
	Dim iCount
	Dim sRowContents
	Dim lErrorNumber

	If bBorder Then
		sRowContents = RTF_ROW_BEGIN_BORDER & " "
	Else
		sRowContents = RTF_ROW_BEGIN & " "
	End If
	lErrorNumber = AppendTextToFile(sFileName, sRowContents, sErrorDescription)

	For iCount = 0 To UBound(asRowContents)
		sWidth = ""
		sAlign = ""
		If Len(asCellWidths(iCount)) > 0 Then
			sWidth = asCellWidths(iCount)
			If bBorder Then
				sRowContents = RTF_CELL_BEGIN_BORDER & sWidth
			Else
				sRowContents = RTF_CELL_BEGIN & sWidth
			End If
			lErrorNumber = AppendTextToFile(sFileName, sRowContents, sErrorDescription)
		End If
	Next

	For iCount = 0 To UBound(asRowContents)
		sAlign = RTF_LEFT
		Select Case asCellAlignments(iCount)
			Case "CENTER"
				sAlign = RTF_CENTER
			Case "RIGHT"
				sAlign = RTF_RIGHT
			Case Else
				sAlign = RTF_LEFT
		End Select
		sRowContents = sAlign & " " & asRowContents(iCount) & " " & RTF_CELL_END
		lErrorNumber = AppendTextToFile(sFileName, sRowContents, sErrorDescription)
	Next

	If bBorder Then
		sRowContents = RTF_ROW_END_BORDER & " "
	Else
		sRowContents = RTF_ROW_END & " "
	End If
	lErrorNumber = AppendTextToFile(sFileName, sRowContents, sErrorDescription)

	DisplayRTFRow = lErrorNumber
	Err.Clear
End Function

Function DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, sRowBGColor, sLineColor, sRowID, sStyle, sErrorDescription)
'************************************************************
'Purpose: To display a table row using HTML
'Inputs:  asRowContents, asCellAlignments, asCellWidths, sRowBGColor, sLineColor, sRowID, sStyle
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayTableRow"
	Dim sWidth
	Dim sAlign
	Dim iColSpan
	Dim iPos
	Dim iSpan
	Dim iCount
	Dim iAddCount
	Dim sFontBegin
	Dim bBanding
	Dim sBGColor
	Dim lErrorNumber

	iColSpan = 0
	For iCount = 0 To UBound(asRowContents)
		If InStr(1, asRowContents(iCount), "<SPAN", vbBinaryCompare) = 1 Then
			iPos = InStr(1, asRowContents(iCount), "COLS=""", vbBinaryCompare) + Len("COLS=""")
			iSpan = CInt(Mid(asRowContents(iCount), iPos, (InStr(iPos, asRowContents(iCount), """", vbBinaryCompare) - iPos)))
			iColSpan = iColSpan + iSpan
		Else
			iColSpan = iColSpan + 1
		End If
	Next
	iColSpan = ((iColSpan + 1) * 5)
	If Len(sRowBGColor) = 0 Then sRowBGColor = "FFFFFF"
	If Len(sLineColor) = 0 Then sLineColor = S_BGCOLOR_FOR_GUI
	If StrComp(GetASPFileName(""), "Export.asp", vbTextCompare) = 0 Then
		sFontBegin = "<FONT FACE=""Arial"" SIZE=""2"">"
		bBanding = False
	Else
		sFontBegin = "<FONT FACE=""Verdana"" SIZE=""1"">"
		bBanding = (InStr(1, sStyle, "<BANDING />", vbBinaryCompare) > 0)
	End If
	sStyle = Replace(sStyle, "<BANDING />", "")
	Response.Write "<TR BGCOLOR=""#" & sRowBGColor & """"
		If Len(sRowID) > 0 Then Response.Write " NAME=""" & sRowID & """ ID=""" & sRowID & """"
	Response.Write " onMouseOver=""SwitchItemBGColor(this, '" & S_SELECTED_BGCOLOR_MENU & "')"" onMouseOut=""SwitchItemBGColor(this, '" & sRowBGColor & "')"" " & sStyle & ">"
		iAddCount = 0
		For iCount = 0 To UBound(asRowContents)
			sWidth = ""
			sAlign = ""
			sBGColor = ""
			If bBanding And ((iCount Mod 2) = 1) Then sBGColor = " BGCOLOR=""#CCCCCC"" "
			If Len(asCellWidths(iCount + iAddCount)) > 0 Then sWidth = "WIDTH=""" & asCellWidths(iCount + iAddCount) & """"
			If Len(asCellAlignments(iCount + iAddCount)) > 0 Then sAlign = "ALIGN=""" & asCellAlignments(iCount + iAddCount) & """"
			Response.Write "<TD WIDTH=""1""" & sBGColor & "><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
			Response.Write "<TD" & sBGColor & ">&nbsp;&nbsp;</TD>"
			If InStr(1, asRowContents(iCount), "<SPAN", vbBinaryCompare) = 1 Then
				iPos = InStr(1, asRowContents(iCount), "COLS=""", vbBinaryCompare) + Len("COLS=""")
				iSpan = (CInt(Mid(asRowContents(iCount), iPos, (InStr(iPos, asRowContents(iCount), """", vbBinaryCompare) - iPos))) - 1)
				iAddCount = iAddCount + iSpan
				iSpan = iSpan * 5 + 1
				Response.Write "<TD " & sBGColor & sAlign & " VALIGN=""TOP"" COLSPAN=""" & iSpan & """"
			Else
				Response.Write "<TD " & sBGColor & sWidth & " " & sAlign & " VALIGN=""TOP"""
			End If
			Response.Write ">" & sFontBegin & Replace(asRowContents(iCount), vbNewLine, "<BR />", 1, -1, vbBinaryCompare) & "</FONT></TD>"
			Response.Write "<TD" & sBGColor & ">&nbsp;&nbsp;</TD>"
			Response.Write "<TD WIDTH=""1""" & sBGColor & "><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>" & vbNewLine
		Next
	Response.Write "</TR>" & vbNewLine
    If StrComp(sLineColor, "<NO_LINE />", vbBinaryCompare) <> 0 Then
		Response.Write "<TR"
			If Len(sRowID) > 0 Then Response.Write " NAME=""" & sRowID & "Ln"" ID=""" & sRowID & "Ln"""
		Response.Write " " & sStyle & "><TD BGCOLOR=""#" & sLineColor & """ COLSPAN=""" & iColSpan & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR>" & vbNewLine
	End If
    Response.Flush()

	DisplayTableRow = lErrorNumber
	Err.Clear
End Function

Function DisplayTableRowText(asRowContents, bUseHTML, sErrorDescription)
'************************************************************
'Purpose: To display a table row using plain text
'Inputs:  asRowContents, bUseHTML
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayTableRowText"
	Dim iSpan
	Dim iPos
	Dim iCount
	Dim sRowBegin
	Dim sRowEnd
	Dim sCellBegin
	Dim sCellSpanBegin
	Dim sCellEnd
	Dim sFontBegin
	Dim sFontEnd
	Dim lErrorNumber
	
	If bUseHTML Then
		sRowBegin = "<TR>"
		sRowEnd = "</TR>"
		sCellBegin = "<TD VALIGN=""TOP"">"
		sCellSpanBegin = "<TD COLSPAN=""xxx"">"
		sCellEnd = "</TD>"
		sFontBegin = "<FONT FACE=""Arial"" SIZE=""2"">"
		sFontEnd = "</FONT>"
	Else
		sRowEnd = vbNewLine
		sCellEnd = vbTab
	End If
	Response.Write sRowBegin
		For iCount = 0 To UBound(asRowContents)
			If InStr(1, asRowContents(iCount), "<SPAN", vbBinaryCompare) = 1 Then
				iPos = InStr(1, asRowContents(iCount), "COLS=""", vbBinaryCompare) + Len("COLS=""")
				iSpan = CInt(Mid(asRowContents(iCount), iPos, (InStr(iPos, asRowContents(iCount), """", vbBinaryCompare) - iPos)))
				Response.Write Replace(sCellSpanBegin, "xxx", iSpan) & sFontBegin & Mid(asRowContents(iCount), (InStr(iPos, asRowContents(iCount), ">", vbBinaryCompare) + Len(">"))) & sFontEnd & sCellEnd
			Else
				Response.Write sCellBegin & sFontBegin & asRowContents(iCount) & sFontEnd & sCellEnd
			End If
		Next
	Response.Write sRowEnd
	Response.Flush()

	DisplayTableRowText = lErrorNumber
	Err.Clear
End Function

Function GetTableRow(asRowContents, asCellAlignments, asCellWidths, sLineColor, sRowID, sStyle, sErrorDescription)
'************************************************************
'Purpose: To display a table row using HTML
'Inputs:  asRowContents, asCellAlignments, asCellWidths, sLineColor, sRowID, sStyle
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetTableRow"
	Dim sWidth
	Dim sAlign
	Dim iColSpan
	Dim iPos
	Dim iSpan
	Dim iCount
	Dim iAddCount
	Dim sFontBegin
	Dim bBanding
	Dim sBGColor

	iColSpan = 0
	For iCount = 0 To UBound(asRowContents)
		If InStr(1, asRowContents(iCount), "<SPAN", vbBinaryCompare) = 1 Then
			iPos = InStr(1, asRowContents(iCount), "COLS=""", vbBinaryCompare) + Len("COLS=""")
			iSpan = CInt(Mid(asRowContents(iCount), iPos, (InStr(iPos, asRowContents(iCount), """", vbBinaryCompare) - iPos)))
			iColSpan = iColSpan + iSpan
		Else
			iColSpan = iColSpan + 1
		End If
	Next
	iColSpan = ((iColSpan + 1) * 5)
	If Len(sLineColor) = 0 Then sLineColor = S_BGCOLOR_FOR_GUI
	If StrComp(GetASPFileName(""), "Export.asp", vbTextCompare) = 0 Then
		sFontBegin = "<FONT FACE=""Arial"" SIZE=""2"">"
		bBanding = False
	Else
		sFontBegin = "<FONT FACE=""Verdana"" SIZE=""1"">"
		bBanding = (InStr(1, sStyle, "<BANDING />", vbBinaryCompare) > 0)
	End If
	sStyle = Replace(sStyle, "<BANDING />", "")
	GetTableRow = "<TR"
		If Len(sRowID) > 0 Then GetTableRow = GetTableRow & " NAME=""" & sRowID & """ ID=""" & sRowID & """"
	GetTableRow = GetTableRow & " onMouseOver=""SwitchItemBGColor(this, '" & S_SELECTED_BGCOLOR_MENU & "')"" onMouseOut=""SwitchItemBGColor(this, '')"" " & sStyle & ">"
		iAddCount = 0
		For iCount = 0 To UBound(asRowContents)
			sWidth = ""
			sAlign = ""
			sBGColor = ""
			If bBanding And ((iCount Mod 2) = 1) Then sBGColor = " BGCOLOR=""#CCCCCC"" "
			If Len(asCellWidths(iCount + iAddCount)) > 0 Then sWidth = "WIDTH=""" & asCellWidths(iCount + iAddCount) & """"
			If Len(asCellAlignments(iCount + iAddCount)) > 0 Then sAlign = "ALIGN=""" & asCellAlignments(iCount + iAddCount) & """"
			GetTableRow = GetTableRow & "<TD WIDTH=""1""" & sBGColor & "><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
			GetTableRow = GetTableRow & "<TD" & sBGColor & ">&nbsp;&nbsp;</TD>"
			If InStr(1, asRowContents(iCount), "<SPAN", vbBinaryCompare) = 1 Then
				iPos = InStr(1, asRowContents(iCount), "COLS=""", vbBinaryCompare) + Len("COLS=""")
				iSpan = (CInt(Mid(asRowContents(iCount), iPos, (InStr(iPos, asRowContents(iCount), """", vbBinaryCompare) - iPos))) - 1)
				iAddCount = iAddCount + iSpan
				iSpan = iSpan * 5 + 1
				GetTableRow = GetTableRow & "<TD " & sBGColor & sAlign & " VALIGN=""TOP"" COLSPAN=""" & iSpan & """"
			Else
				GetTableRow = GetTableRow & "<TD " & sBGColor & sWidth & " " & sAlign & " VALIGN=""TOP"""
			End If
			GetTableRow = GetTableRow & ">" & sFontBegin & Replace(asRowContents(iCount), vbNewLine, "<BR />", 1, -1, vbBinaryCompare) & "</FONT></TD>"
			GetTableRow = GetTableRow & "<TD" & sBGColor & ">&nbsp;&nbsp;</TD>"
			GetTableRow = GetTableRow & "<TD WIDTH=""1""" & sBGColor & "><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>" & vbNewLine
		Next
	GetTableRow = GetTableRow & "</TR>" & vbNewLine
    GetTableRow = GetTableRow & "<TR"
		If Len(sRowID) > 0 Then GetTableRow = GetTableRow & " NAME=""" & sRowID & "Ln"" ID=""" & sRowID & "Ln"""
	GetTableRow = GetTableRow & " " & sStyle & "><TD BGCOLOR=""#" & sLineColor & """ COLSPAN=""" & iColSpan & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR>" & vbNewLine

	Err.Clear
End Function

Function GetTableRowText(asRowContents, bUseHTML, sErrorDescription)
'************************************************************
'Purpose: To display a table row using plain text
'Inputs:  asRowContents, bUseHTML
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetTableRowText"
	Dim iSpan
	Dim iPos
	Dim iCount
	Dim sRowBegin
	Dim sRowEnd
	Dim sCellBegin
	Dim sCellSpanBegin
	Dim sCellEnd
	Dim sFontBegin
	Dim sFontEnd
	
	If bUseHTML Then
		sRowBegin = "<TR>"
		sRowEnd = "</TR>"
		sCellBegin = "<TD VALIGN=""TOP"">"
		sCellSpanBegin = "<TD COLSPAN=""xxx"">"
		sCellEnd = "</TD>"
		sFontBegin = "<FONT FACE=""Arial"" SIZE=""2"">"
		sFontEnd = "</FONT>"
	Else
		sRowEnd = vbNewLine
		sCellEnd = vbTab
	End If
	GetTableRowText = sRowBegin
		For iCount = 0 To UBound(asRowContents)
			If InStr(1, asRowContents(iCount), "<SPAN", vbBinaryCompare) = 1 Then
				iPos = InStr(1, asRowContents(iCount), "COLS=""", vbBinaryCompare) + Len("COLS=""")
				iSpan = CInt(Mid(asRowContents(iCount), iPos, (InStr(iPos, asRowContents(iCount), """", vbBinaryCompare) - iPos)))
				GetTableRowText = GetTableRowText & Replace(sCellSpanBegin, "xxx", iSpan) & sFontBegin & Mid(asRowContents(iCount), (InStr(iPos, asRowContents(iCount), ">", vbBinaryCompare) + Len(">"))) & sFontEnd & sCellEnd
			Else
				GetTableRowText = GetTableRowText & sCellBegin & sFontBegin & asRowContents(iCount) & sFontEnd & sCellEnd
			End If
		Next
	GetTableRowText = GetTableRowText & sRowEnd

	Err.Clear
End Function
%>