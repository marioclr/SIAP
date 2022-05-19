<%
Const N_HEIGHT_GRAPH = 0
Const N_BARS_WIDTH_GRAPH = 1
Const N_SEPARATION_WIDTH_GRAPH = 2
Const N_SERIES_SEPARATION_WIDTH_GRAPH = 3
Const N_SCALE_GRAPH = 4
Const S_SCALE_IMAGE_GRAPH = 5
Const N_SEPARATORS_GRAPH = 6
Const S_BASE_IMAGE_GRAPH = 7
Const AS_BARS_COLOR_GRAPH = 8
Const N_START_COLOR_GRAPH = 9
Const A_VALUES_GRAPH = 10
Const A_HEADERS_GRAPH = 11
Const A_LEGEND_GRAPH = 12
Const S_DESCRIPTION_GRAPH = 13
Const B_SIDE_BY_SIDE_TEXT_GRAPH = 14
Const S_BGCOLOR_GRAPH = 15
Const S_FRAME_COLOR_GRAPH = 16
Const N_FRAME_SEPARATION_GRAPH = 17
Const N_FRAME_HEIGHT_GRAPH = 18
Const N_SCALE_FOR_BARS_GRAPH = 19
Const B_INVERT_SERIES_GRAPH = 20
Const B_COMPONENT_INITIALIZED_GRAPH = 21

Const N_GRAPH_COMPONENT_SIZE = 21

Dim S_COLOR_PALETTE
'If StrComp(GetOption(aOptionsComponent, COLORS_IN_GRAPHS_OPTION), "1", vbBinaryCompare) = 0 Then
	S_COLOR_PALETTE = "D20000,004F5C,314884,788431,F4CA1E,C09600,C00050,970027,E0E2DD,8C268E,F250A1,009DB6,C3C3B8,9CC0DD,BB9CC7,DA8888,FF5A00,6D2A38,F6082A,00F3B6,909090,5608F6,44F608,00DEFE,B9F9F9,384467,38B800,846631,629A84,F6B289,4185FF,1E6600,1F5B59,4D2525"
'Else
'	S_COLOR_PALETTE = "000000,000000,000000,000000,000000,000000,000000,000000,000000,000000,000000,000000,000000,000000,000000,000000,000000,000000,000000,000000,000000,000000,000000,000000,000000,000000,000000,000000,000000,000000,000000,000000,000000,000000"
'End If

Dim aGraphComponent()
ReDim aGraphComponent(N_GRAPH_COMPONENT_SIZE)

Function InitializeGraphComponent(oRequest, aGraphComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Graph Component
'         using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aGraphComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeGraphComponent"
	Dim aTemp()
	Dim sTemp
	Dim iIndex
	Dim jIndex
	Redim Preserve aGraphComponent(N_GRAPH_COMPONENT_SIZE)

	If IsEmpty(aGraphComponent(N_HEIGHT_GRAPH)) Then
		If Len(oRequest("GraphHeight").Item) > 0 Then
			aGraphComponent(N_HEIGHT_GRAPH) = CLng(oRequest("GraphHeight").Item)
		Else
			aGraphComponent(N_HEIGHT_GRAPH) = 300
		End If
	End If

	If IsEmpty(aGraphComponent(N_BARS_WIDTH_GRAPH)) Then
		If Len(oRequest("GraphBarsWidth").Item) > 0 Then
			aGraphComponent(N_BARS_WIDTH_GRAPH) = CLng(oRequest("GraphBarsWidth").Item)
		Else
			aGraphComponent(N_BARS_WIDTH_GRAPH) = 3
		End If
	End If

	If IsEmpty(aGraphComponent(N_SEPARATION_WIDTH_GRAPH)) Then
		If Len(oRequest("GraphBarsSeparation").Item) > 0 Then
			aGraphComponent(N_SEPARATION_WIDTH_GRAPH) = CLng(oRequest("GraphBarsSeparation").Item)
		Else
			aGraphComponent(N_SEPARATION_WIDTH_GRAPH) = 3
		End If
	End If

	If IsEmpty(aGraphComponent(N_SERIES_SEPARATION_WIDTH_GRAPH)) Then
		If Len(oRequest("GraphBarsSeriesSeparation").Item) > 0 Then
			aGraphComponent(N_SERIES_SEPARATION_WIDTH_GRAPH) = CLng(oRequest("GraphBarsSeriesSeparation").Item)
		Else
			aGraphComponent(N_SERIES_SEPARATION_WIDTH_GRAPH) = 5
		End If
	End If

	If IsEmpty(aGraphComponent(N_SCALE_GRAPH)) Then
		If Len(oRequest("GraphScale").Item) > 0 Then
			aGraphComponent(N_SCALE_GRAPH) = CLng(oRequest("GraphScale").Item)
		Else
			aGraphComponent(N_SCALE_GRAPH) = 100
		End If
	End If

	If IsEmpty(aGraphComponent(S_SCALE_IMAGE_GRAPH)) Then
		If Len(oRequest("GraphScaleImage").Item) > 0 Then
			aGraphComponent(S_SCALE_IMAGE_GRAPH) = oRequest("GraphScaleImage").Item
		Else
			aGraphComponent(S_SCALE_IMAGE_GRAPH) = "Scale100.gif"
		End If
	End If

	If IsEmpty(aGraphComponent(N_SEPARATORS_GRAPH)) Then
		If Len(oRequest("GraphSeparators").Item) > 0 Then
			aGraphComponent(N_SEPARATORS_GRAPH) = CLng(oRequest("GraphSeparators").Item)
		Else
			aGraphComponent(N_SEPARATORS_GRAPH) = 5
		End If
	End If

	If IsEmpty(aGraphComponent(S_BASE_IMAGE_GRAPH)) Then
		If Len(oRequest("GraphBaseImage").Item) > 0 Then
			aGraphComponent(S_BASE_IMAGE_GRAPH) = oRequest("GraphBaseImage").Item
		Else
			aGraphComponent(S_BASE_IMAGE_GRAPH) = "Transparent.gif"
		End If
	End If

	If IsEmpty(aGraphComponent(AS_BARS_COLOR_GRAPH)) Then
		If Len(oRequest("GraphBarColor").Item) > 0 Then
			aGraphComponent(AS_BARS_COLOR_GRAPH) = Split(Replace(oRequest("GraphBarColor").Item, "#", ""), LIST_SEPARATOR, -1, vbBinaryCompare)
		Else
			aGraphComponent(AS_BARS_COLOR_GRAPH) = Split(S_COLOR_PALETTE, ",", -1, vbBinaryCompare)
		End If
	End If

	If IsEmpty(aGraphComponent(N_START_COLOR_GRAPH)) Then
		If Len(oRequest("GraphStartColor").Item) > 0 Then
			aGraphComponent(N_START_COLOR_GRAPH) = CInt(oRequest("GraphStartColor").Item)
		Else
			aGraphComponent(N_START_COLOR_GRAPH) = 0
		End If
	End If

	If IsEmpty(aGraphComponent(A_VALUES_GRAPH)) Then
		If Len(oRequest("GraphValues").Item) > 0 Then
			Call DoubleSplit(oRequest("GraphValues").Item, aGraphComponent(A_VALUES_GRAPH))
		Else
			Randomize
			Redim aTemp(Int(5 * Rnd))
			jIndex = Int(25 * Rnd) + 5
			For iIndex = 0 To UBound(aTemp)
				aTemp(iIndex) = GenerateRandomNumbersArray(aGraphComponent(N_SCALE_GRAPH), jIndex, False)
			Next
			aGraphComponent(A_VALUES_GRAPH) = aTemp
		End If
	End If

	If IsEmpty(aGraphComponent(A_LEGEND_GRAPH)) Then
		If Len(oRequest("GraphLegend").Item) > 0 Then
			aGraphComponent(A_LEGEND_GRAPH) = Split(oRequest("GraphLegend").Item, LIST_SEPARATOR, -1, vbBinaryCompare)
		Else
			sTemp = ""
			For iIndex = 0 To UBound(aGraphComponent(A_VALUES_GRAPH))
				sTemp = sTemp & iIndex & LIST_SEPARATOR
			Next
			sTemp = Left(sTemp, (Len(sTemp) - Len(LIST_SEPARATOR)))
			aGraphComponent(A_LEGEND_GRAPH) = Split(sTemp, LIST_SEPARATOR, -1, vbBinaryCompare)
		End If
	End If

	If IsEmpty(aGraphComponent(S_DESCRIPTION_GRAPH)) Then
		If Len(oRequest("GraphDescription").Item) > 0 Then
			aGraphComponent(S_DESCRIPTION_GRAPH) = CleanStringForHTML(oRequest("GraphDescription").Item)
		Else
			aGraphComponent(S_DESCRIPTION_GRAPH) = "Bla, bla, bla, bla, bla, bla, bla<BR />bla, bla, bla, bla, bla, bla, bla<BR />bla, bla, bla, bla, bla, bla, bla<BR />bla, bla, bla, bla, bla, bla, bla"
		End If
	End If

	If IsEmpty(aGraphComponent(B_SIDE_BY_SIDE_TEXT_GRAPH)) Then
		aGraphComponent(B_SIDE_BY_SIDE_TEXT_GRAPH) = (Len(oRequest("GraphSideBySideText").Item) > 0)
	End If

	If IsEmpty(aGraphComponent(S_BGCOLOR_GRAPH)) Then
		If Len(oRequest("GraphBGColor").Item) > 0 Then
			aGraphComponent(S_BGCOLOR_GRAPH) = Replace(oRequest("GraphBGColor").Item, "#", "")
		Else
			aGraphComponent(S_BGCOLOR_GRAPH) = "FFFFFF"
		End If
	End If

	If IsEmpty(aGraphComponent(S_FRAME_COLOR_GRAPH)) Then
		If Len(oRequest("GraphFrameColor").Item) > 0 Then
			aGraphComponent(S_FRAME_COLOR_GRAPH) = Replace(oRequest("GraphFrameColor").Item, "#", "")
		Else
			aGraphComponent(S_FRAME_COLOR_GRAPH) = "000000"
		End If
	End If

	If IsEmpty(aGraphComponent(N_FRAME_SEPARATION_GRAPH)) Then
		If Len(oRequest("GraphFrameSeparation").Item) > 0 Then
			aGraphComponent(N_FRAME_SEPARATION_GRAPH) = CInt(oRequest("GraphFrameSeparation").Item)
		Else
			aGraphComponent(N_FRAME_SEPARATION_GRAPH) = 20
		End If
	End If

	aGraphComponent(N_FRAME_HEIGHT_GRAPH) = aGraphComponent(N_HEIGHT_GRAPH) + 2
	aGraphComponent(N_SCALE_FOR_BARS_GRAPH) = CDbl(aGraphComponent(N_HEIGHT_GRAPH) / aGraphComponent(N_SCALE_GRAPH))

	aGraphComponent(B_COMPONENT_INITIALIZED_GRAPH) = True
	InitializeGraphComponent = Err.number
	Err.Clear
End Function

Function SetScaleForGraph(oRequest, aGraphComponent, sErrorDescription)
'************************************************************
'Purpose: To set the scale and the image for the graph given
'         the values
'Inputs:  oRequest, aGraphComponent
'Outputs: aGraphComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetScaleForGraph"
	Dim iIndex
	Dim jIndex
	Dim aGraphHeights
	Dim bComponentInitialized

'	bComponentInitialized = aGraphComponent(B_COMPONENT_INITIALIZED_GRAPH)
'	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
'		Call InitializeGraphComponent(oRequest, aGraphComponent)
'	End If

	aGraphHeights = Split("71,71,71,99,129,159,187,216,245,274,303", ",", -1, vbBinaryCompare)
	aGraphComponent(N_SCALE_GRAPH) = 0
	For iIndex = 0 To UBound(aGraphComponent(A_VALUES_GRAPH))
		If IsArray(aGraphComponent(A_VALUES_GRAPH)(iIndex)) Then
			For jIndex = 0 To UBound(aGraphComponent(A_VALUES_GRAPH)(iIndex))
				If aGraphComponent(N_SCALE_GRAPH) < CInt(aGraphComponent(A_VALUES_GRAPH)(iIndex)(jIndex)) Then aGraphComponent(N_SCALE_GRAPH) = CInt(aGraphComponent(A_VALUES_GRAPH)(iIndex)(jIndex))
			Next
		Else
			If aGraphComponent(N_SCALE_GRAPH) < CInt(aGraphComponent(A_VALUES_GRAPH)(iIndex)) Then aGraphComponent(N_SCALE_GRAPH) = CInt(aGraphComponent(A_VALUES_GRAPH)(iIndex))
		End If
	Next
	aGraphComponent(N_SCALE_GRAPH) = CInt(aGraphComponent(N_SCALE_GRAPH) / 10) * 10 + 10
	If aGraphComponent(N_SCALE_GRAPH) < 20 Then
		aGraphComponent(N_SCALE_GRAPH) = 20
		aGraphComponent(N_HEIGHT_GRAPH) = aGraphHeights(2)
		aGraphComponent(N_SEPARATORS_GRAPH) = 5
	ElseIf (aGraphComponent(N_SCALE_GRAPH) > 100) And (aGraphComponent(N_SCALE_GRAPH) <= 1000) Then
		aGraphComponent(N_SCALE_GRAPH) = 1000
		aGraphComponent(N_HEIGHT_GRAPH) = 303
		aGraphComponent(N_SEPARATORS_GRAPH) = 5
	ElseIf (aGraphComponent(N_SCALE_GRAPH) > 1000) And (aGraphComponent(N_SCALE_GRAPH) <= 10000) Then
		aGraphComponent(N_SCALE_GRAPH) = 10000
		aGraphComponent(N_HEIGHT_GRAPH) = 303
		aGraphComponent(N_SEPARATORS_GRAPH) = 5
	ElseIf (aGraphComponent(N_SCALE_GRAPH) > 10000) And (aGraphComponent(N_SCALE_GRAPH) <= 100000) Then
		aGraphComponent(N_SCALE_GRAPH) = 100000
		aGraphComponent(N_HEIGHT_GRAPH) = 303
		aGraphComponent(N_SEPARATORS_GRAPH) = 5
	Else
		aGraphComponent(N_HEIGHT_GRAPH) = aGraphHeights(aGraphComponent(N_SCALE_GRAPH) / 10)
		aGraphComponent(N_SEPARATORS_GRAPH) = aGraphComponent(N_SCALE_GRAPH) / 5
	End If
	aGraphComponent(S_SCALE_IMAGE_GRAPH) = "Scale" & aGraphComponent(N_SCALE_GRAPH) & ".gif"

	Erase aGraphHeights
	SetScaleForGraph = Err.number
End Function

Function DisplayGraph(oRequest, bForExport, aGraphComponent, sErrorDescription)
'************************************************************
'Purpose: To initialize the empty elements of the Graph Component
'         using the URL parameters or default values
'Inputs:  oRequest, bForExport, aGraphComponent
'Outputs: aGraphComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayGraph"
	Dim iIndex, jIndex
	Dim bComponentInitialized

	bComponentInitialized = aGraphComponent(B_COMPONENT_INITIALIZED_GRAPH)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeGraphComponent(oRequest, aGraphComponent)
	End If

	If bForExport Then
		Response.Write "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			Response.Write "<TR>"
				Response.Write "<TD>&nbsp;</TD>" & vbNewLine
				For iIndex = 0 To UBound(aGraphComponent(A_LEGEND_GRAPH))
					Response.Write "<TD>" & aGraphComponent(A_LEGEND_GRAPH)(iIndex) & "</TD>" & vbNewLine
				Next
			Response.Write "</TR>"

			iIndex = UBound(aGraphComponent(A_VALUES_GRAPH)(0))
			If Err.number <> 0 Then
				Err.Clear
				For iIndex = 0 To UBound(aGraphComponent(A_VALUES_GRAPH))
					Response.Write "<TR>" & vbNewLine
						Response.Write "<TD>" & aGraphComponent(A_HEADERS_GRAPH)(iIndex) & "</TD>" & vbNewLine
						Call DisplayBar(bForExport, aGraphComponent, iIndex)
					Response.Write "</TR>" & vbNewLine
				Next
			Else
				For iIndex = 0 To UBound(aGraphComponent(A_VALUES_GRAPH)(0))
					Response.Write "<TR>" & vbNewLine
						Response.Write "<TD>" & aGraphComponent(A_HEADERS_GRAPH)(iIndex) & "</TD>" & vbNewLine
						For jIndex = 0 To UBound(aGraphComponent(A_VALUES_GRAPH))
							Call DisplayBarIJ(bForExport, aGraphComponent, iIndex, jIndex)
						Next
					Response.Write "</TR>" & vbNewLine
				Next
			End If
		Response.Write "</TABLE>" & vbNewLine
		If Not aGraphComponent(B_SIDE_BY_SIDE_TEXT_GRAPH) Then
			Response.Write "<BR />" & aGraphComponent(S_DESCRIPTION_GRAPH)
		End If
	Else
		Response.Write "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			Response.Write "<TR>" & vbNewLine
				Response.Write "<TD VALIGN=""BOTTOM"">" & vbNewLine
					If Len(aGraphComponent(S_SCALE_IMAGE_GRAPH)) > 0 Then
						Response.Write "<IMG SRC=""Images/" & aGraphComponent(S_SCALE_IMAGE_GRAPH) & """ /><BR /><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" />"
					Else
						Call DisplayVerticalScale(aGraphComponent(N_HEIGHT_GRAPH), aGraphComponent(N_SCALE_GRAPH), aGraphComponent(N_SEPARATORS_GRAPH))
					End If
				Response.Write "</TD>" & vbNewLine
				Response.Write "<TD VALIGN=""BOTTOM"">" & vbNewLine
					Response.Write "<TABLE BGCOLOR=""#" & aGraphComponent(S_FRAME_COLOR_GRAPH) & """ HEIGHT=""" & aGraphComponent(N_FRAME_HEIGHT_GRAPH) & """ BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""1""><TR>" & vbNewLine
						Response.Write "<TD VALIGN=""TOP"">" & vbNewLine
							Response.Write "<TABLE BGCOLOR=""#" & aGraphComponent(S_BGCOLOR_GRAPH)& """ WIDTH=""100%"" HEIGHT=""100%"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0""><TR>" & vbNewLine
								Response.Write "<TD><IMG SRC=""Images/Transparent.gif"" WIDTH=""" & aGraphComponent(N_FRAME_SEPARATION_GRAPH) & """ HEIGHT=""" & aGraphComponent(N_HEIGHT_GRAPH) & """ /></TD>" & vbNewLine
								Err.clear
								iIndex = UBound(aGraphComponent(A_VALUES_GRAPH)(0))
								If Err.number <> 0 Then
									Err.Clear
									For iIndex = 0 To UBound(aGraphComponent(A_VALUES_GRAPH))
										Call DisplayBar(bForExport, aGraphComponent, iIndex)
									Next
								Else
									For iIndex = 0 To UBound(aGraphComponent(A_VALUES_GRAPH)(0))
										For jIndex = 0 To UBound(aGraphComponent(A_VALUES_GRAPH))
											Call DisplayBarIJ(bForExport, aGraphComponent, iIndex, jIndex)
										Next
										Response.Write "<TD><IMG SRC=""Images/Transparent.gif"" WIDTH=""" & aGraphComponent(N_SERIES_SEPARATION_WIDTH_GRAPH) & """ HEIGHT=""1"" /></TD>" & vbNewLine
									Next
								End If
								Response.Write "<TD><IMG SRC=""Images/Transparent.gif"" WIDTH=""" & aGraphComponent(N_FRAME_SEPARATION_GRAPH) & """ HEIGHT=""1"" /></TD>" & vbNewLine
							Response.Write "</TR></TABLE>" & vbNewLine
						Response.Write "</TD>" & vbNewLine
						Response.Write "<TD BGCOLOR=""#" & aGraphComponent(S_BGCOLOR_GRAPH)& """ VALIGN=""TOP"">" & vbNewLine
							Response.Write "<TABLE BORDER=""0"" CELLSPACING=""5"" CELLPADDING=""0"">" & vbNewLine
								For iIndex = 0 To UBound(aGraphComponent(A_LEGEND_GRAPH))
									Response.Write "<TR>"
										Response.Write "<TD WIDTH=""3"" BGCOLOR=""#" & aGraphComponent(AS_BARS_COLOR_GRAPH)((iIndex + aGraphComponent(N_START_COLOR_GRAPH)) Mod (UBound(aGraphComponent(AS_BARS_COLOR_GRAPH)) + 1)) & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""3"" HEIGHT=""3"" /></TD>" & vbNewLine
										Response.Write "<TD><FONT FACE=""Verdana"" SIZE=""1"">" & aGraphComponent(A_LEGEND_GRAPH)(iIndex) & "</FONT></TD>" & vbNewLine
									Response.Write "</TR>"
									Response.Write "<TR>"
										Response.Write "<TD COLSPAN=""2""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""3"" /></TD>" & vbNewLine
									Response.Write "</TR>"
								Next
								If Not aGraphComponent(B_SIDE_BY_SIDE_TEXT_GRAPH) Then
									Response.Write "<TR><TD></TD><TD COLSPAN=""1""><FONT FACE=""Verdana"" SIZE=""1"">" & aGraphComponent(S_DESCRIPTION_GRAPH) & "</FONT></TD></TR>"
								End If
							Response.Write "</TABLE>" & vbNewLine
						Response.Write "</TD>" & vbNewLine
						If aGraphComponent(B_SIDE_BY_SIDE_TEXT_GRAPH) Then
							Response.Write "<TD BGCOLOR=""#" & aGraphComponent(S_BGCOLOR_GRAPH)& """><IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""1"" /></TD>" & vbNewLine
							Response.Write "<TD BGCOLOR=""#" & aGraphComponent(S_BGCOLOR_GRAPH)& """ VALIGN=""TOP"">" & vbNewLine
								Response.Write "<FONT FACE=""Verdana"" SIZE=""1"">" & aGraphComponent(S_DESCRIPTION_GRAPH) & "</FONT>"
							Response.Write "</TD>" & vbNewLine
						End If
					Response.Write "</TR></TABLE>" & vbNewLine
				Response.Write "</TD>" & vbNewLine
			Response.Write "</TR>" & vbNewLine
			Response.Write "<TR><TD>&nbsp;</TD><TD><IMG SRC=""Images/" & aGraphComponent(S_BASE_IMAGE_GRAPH) & """></TD></TR>" & vbNewLine
		Response.Write "</TABLE>" & vbNewLine
	End If

	DisplayGraph = Err.number
	Err.Clear
End Function

Function DisplayBar(bForExport, aGraphComponent, iIndex)
'************************************************************
'Purpose: To display a bar given the indexes
'Inputs:  bForExport, aGraphComponent, iIndex
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayBar"
	Dim iTopHeight
	Dim iBottomHeight

	If bForExport Then
		Response.Write "<TD>" & aGraphComponent(A_VALUES_GRAPH)(iIndex) & "</TD>" & vbNewLine
	Else
		iTopHeight = Int((aGraphComponent(N_SCALE_GRAPH) - aGraphComponent(A_VALUES_GRAPH)(iIndex)) * aGraphComponent(N_SCALE_FOR_BARS_GRAPH))
		iBottomHeight = Int(aGraphComponent(A_VALUES_GRAPH)(iIndex) * aGraphComponent(N_SCALE_FOR_BARS_GRAPH))
		Response.Write "<TD VALIGN=""BOTTOM"">" & vbNewLine
			Response.Write "<TABLE HEIGHT=""" & aGraphComponent(N_HEIGHT_GRAPH) & """ BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If aGraphComponent(A_VALUES_GRAPH)(iIndex) > 0 Then
					Response.Write "<TR><TD HEIGHT=""" & iTopHeight & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""" & aGraphComponent(N_BARS_WIDTH_GRAPH) & """ HEIGHT=""" & iTopHeight & """ ALT=""" & aGraphComponent(A_VALUES_GRAPH)(iIndex) & """ /></TD></TR>" & vbNewLine
					Response.Write "<TR><TD HEIGHT=""" & iBottomHeight & """ BGCOLOR=""#" & aGraphComponent(AS_BARS_COLOR_GRAPH)(0) & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""" & aGraphComponent(N_BARS_WIDTH_GRAPH) & """ HEIGHT=""" & iBottomHeight & """ ALT=""" & aGraphComponent(A_VALUES_GRAPH)(iIndex) & """ /></TD></TR>" & vbNewLine
				Else
					Response.Write "<TR><TD HEIGHT=""" & iTopHeight + iBottomHeight & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""" & aGraphComponent(N_BARS_WIDTH_GRAPH) & """ HEIGHT=""" & iTopHeight + iBottomHeight & """ ALT=""" & aGraphComponent(A_VALUES_GRAPH)(iIndex) & """ /></TD></TR>" & vbNewLine
				End If
			Response.Write "</TABLE>" & vbNewLine
		Response.Write "</TD>" & vbNewLine
		Response.Write "<TD><IMG SRC=""Images/Transparent.gif"" WIDTH=""" & aGraphComponent(N_SEPARATION_WIDTH_GRAPH) & """ HEIGHT=""1"" /></TD>" & vbNewLine
	End If

	DisplayBar = Err.number
End Function

Function DisplayBarIJ(bForExport, aGraphComponent, iIndex, jIndex)
'************************************************************
'Purpose: To display a bar given the indexes
'Inputs:  bForExport, aGraphComponent, iIndex, jIndex
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayBarIJ"
	Dim iTopHeight
	Dim iBottomHeight

	If bForExport Then
		Response.Write "<TD>" & aGraphComponent(A_VALUES_GRAPH)(jIndex)(iIndex) & "</TD>" & vbNewLine
	Else
		iTopHeight = Int((aGraphComponent(N_SCALE_GRAPH) - aGraphComponent(A_VALUES_GRAPH)(jIndex)(iIndex)) * aGraphComponent(N_SCALE_FOR_BARS_GRAPH))
		iBottomHeight = Int(aGraphComponent(A_VALUES_GRAPH)(jIndex)(iIndex) * aGraphComponent(N_SCALE_FOR_BARS_GRAPH))
		Response.Write "<TD VALIGN=""BOTTOM"">" & vbNewLine
			Response.Write "<TABLE HEIGHT=""" & aGraphComponent(N_HEIGHT_GRAPH) & """ BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If aGraphComponent(A_VALUES_GRAPH)(jIndex)(iIndex) > 0 Then
					Response.Write "<TR><TD HEIGHT=""" & iTopHeight & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""" & aGraphComponent(N_BARS_WIDTH_GRAPH) & """ HEIGHT=""" & iTopHeight & """ ALT=""" & aGraphComponent(A_VALUES_GRAPH)(jIndex)(iIndex) & """ /></TD></TR>" & vbNewLine
					Response.Write "<TR><TD HEIGHT=""" & iBottomHeight & """ BGCOLOR=""#" & aGraphComponent(AS_BARS_COLOR_GRAPH)((jIndex + aGraphComponent(N_START_COLOR_GRAPH)) Mod (UBound(aGraphComponent(AS_BARS_COLOR_GRAPH)) + 1)) & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""" & aGraphComponent(N_BARS_WIDTH_GRAPH) & """ HEIGHT=""" & iBottomHeight & """ ALT=""" & aGraphComponent(A_VALUES_GRAPH)(jIndex)(iIndex) & """ /></TD></TR>" & vbNewLine
				Else
					Response.Write "<TR><TD HEIGHT=""" & iTopHeight + iBottomHeight & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""" & aGraphComponent(N_BARS_WIDTH_GRAPH) & """ HEIGHT=""" & iTopHeight + iBottomHeight & """ ALT=""" & aGraphComponent(A_VALUES_GRAPH)(jIndex)(iIndex) & """ /></TD></TR>" & vbNewLine
				End If
			Response.Write "</TABLE>" & vbNewLine
		Response.Write "</TD>" & vbNewLine
		Response.Write "<TD><IMG SRC=""Images/Transparent.gif"" WIDTH=""" & aGraphComponent(N_SEPARATION_WIDTH_GRAPH) & """ HEIGHT=""1"" /></TD>" & vbNewLine
	End If

	DisplayBarIJ = Err.number
End Function

Function DisplayVerticalScale(lScaleHeight, lMaximumValue, lSeparators)
'************************************************************
'Purpose: To display a vertical scale
'Inputs:  lScaleHeight, lMaximumValue, lSeparators
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayVerticalScale"
	Dim lHeight
	Dim iIndex

	lHeight = CInt(lScaleHeight / lSeparators)
	lHeight = lHeight - 1
	Response.Write "<TABLE WIDTH=""3"" HEIGHT=""" & lScaleHeight & """ BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
		Response.Write "<TR>"
			Response.Write "<TD BGCOLOR=""#FFFFFF"" VALIGN=""TOP"" ROWSPAN=""" & ((lSeparators * 2) + 1) & """><FONT FACE=""Verdana"" SIZE=""1"">"
				Call DisplayTextAsImage(FormatNumber(lMaximumValue, 2, True, False, True), True)
			Response.Write "</FONT></TD>"
			Response.Write "<TD BGCOLOR=""#000000""><IMG SRC=""Images/Transparent.gif"" WIDTH=""3"" HEIGHT=""1"" /></TD>"
		Response.Write "</TR>" & vbNewLine
		For iIndex = 1 To lSeparators
			Response.Write "<TR><TD BGCOLOR=""#FFFFFF""><IMG SRC=""Images/Transparent.gif"" WIDTH=""3"" HEIGHT=""" & lHeight & """ /></TD></TR>" & vbNewLine
			Response.Write "<TR><TD BGCOLOR=""#000000""><IMG SRC=""Images/Transparent.gif"" WIDTH=""3"" HEIGHT=""1"" /></TD></TR>" & vbNewLine
		Next
	Response.Write "</TABLE>" & vbNewLine

	DisplayVerticalScale = Err.number
	Err.Clear
End Function

Function DisplayTextAsImage(lMaximumValue, bHorizontal)
'************************************************************
'Purpose: To display a string using small images
'Inputs:  lMaximumValue, bHorizontal
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayTextAsImage"
	Dim sTemp
	Dim sEnter
	Dim sChar
	Dim iIndex
	
	sTemp = CStr(lMaximumValue)
	sEnter = ""
	If Not bHorizontal Then sEnter = "<BR />"
	Response.Write "<NOBR>"
		For iIndex = 0 To Len(sTemp)
			sChar = Mid(sTemp, iIndex, Len("V"))
			Response.Write "<IMG SRC=""Images/Chr_" & sChar & ".gif"" WIDTH=""3"" HEIGHT=""5"" BORDER=""0"" /><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""6"" BORDER=""0"" />" & sEnter
			If StrComp(sChar, " ", vbBinaryCompare) = 0 Then Response.Write "</NOBR><NOBR>"
		Next
	Response.Write "</NOBR>"

	DisplayTextAsImage = Err.number
	Err.Clear
End Function
%>