<%
Const S_TITLE_MENU = 0
Const A_ELEMENTS_MENU = 1
Const B_USE_DIV_MENU = 2
Const S_POPUP_DIV_NAME_MENU = 3
Const N_LEFT_FOR_DIV_MENU = 4
Const N_TOP_FOR_DIV_MENU = 5
Const N_WIDTH_FOR_DIV_MENU = 6
Const B_CLOSE_DIV_AFTER_ACTION_MENU = 7
Const S_TITLE_ACTION_MENU = 8
Const B_COMPONENT_INITIALIZED_MENU = 9

Const N_MENU_COMPONENT_SIZE = 9

Dim aMenuComponent()
ReDim aMenuComponent(N_MENU_COMPONENT_SIZE)
Dim aOptionsMenuComponent()
ReDim aOptionsMenuComponent(N_MENU_COMPONENT_SIZE)

Function InitializeMenuComponent(aMenuComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Menu Component
'         using the URL parameters or default values
'Outputs: aMenuComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeMenuComponent"
	Redim Preserve aMenuComponent(N_MENU_COMPONENT_SIZE)

	If IsEmpty(aMenuComponent(S_TITLE_MENU)) Then
		aMenuComponent(S_TITLE_MENU) = "Opciones..."
	End If

	If IsEmpty(aMenuComponent(B_USE_DIV_MENU)) Then
		aMenuComponent(B_USE_DIV_MENU) = False
	End If

	If IsEmpty(aMenuComponent(S_POPUP_DIV_NAME_MENU)) Then
		aMenuComponent(S_POPUP_DIV_NAME_MENU) = "Pp" & GetSerialNumberForDate("") & GenerateRandomNumbersSecuence(10) & "Div"
	End If

	If IsEmpty(aMenuComponent(N_LEFT_FOR_DIV_MENU)) Then
		aMenuComponent(N_LEFT_FOR_DIV_MENU) = 0
	End If

	If IsEmpty(aMenuComponent(N_TOP_FOR_DIV_MENU)) Then
		aMenuComponent(N_TOP_FOR_DIV_MENU) = 0
	End If

	If IsEmpty(aMenuComponent(N_WIDTH_FOR_DIV_MENU)) Then
		aMenuComponent(N_WIDTH_FOR_DIV_MENU) = 200
	End If

	If IsEmpty(aMenuComponent(B_CLOSE_DIV_AFTER_ACTION_MENU)) Then
		aMenuComponent(B_CLOSE_DIV_AFTER_ACTION_MENU) = True
	End If

	If IsEmpty(aMenuComponent(S_TITLE_ACTION_MENU)) Then
		aMenuComponent(S_TITLE_ACTION_MENU) = ""
	End If

	aMenuComponent(B_COMPONENT_INITIALIZED_MENU) = True
	InitializeMenuComponent = Err.number
	Err.Clear
End Function

Function DisplayMenuInTwoColumns(aMenuComponent)
'************************************************************
'Purpose: To display the menu items in a two columns table
'Inputs:  aMenuComponent
'Outputs: aMenuComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayMenuInTwoColumns"
	Dim bWasLine
	Dim iIndex
	Dim jIndex
	Dim bFirstItem
	Dim iItemCounter
	Dim bComponentInitialized

	bComponentInitialized = aMenuComponent(B_COMPONENT_INITIALIZED_MENU)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeMenuComponent(aMenuComponent)
	End If

	bFirstItem = True
	iItemCounter = 0
	bWasLine = True
	Response.Write vbNewLine & "<TR>"
	For iIndex = 0 To UBound(aMenuComponent(A_ELEMENTS_MENU))
		If InStr(1, aMenuComponent(A_ELEMENTS_MENU)(iIndex)(0), "<TITLE />", vbBinaryCompare) > 0 Then
			Response.Write "</TR><TR><TD BGCOLOR=""#" & S_WIZARD_TITLE_FOR_GUI & """ COLSPAN=""7""><FONT FACE=""Arial"" SIZE=""2"" COLOR=""#" & S_MENU_LINK_FOR_GUI & """>&nbsp;<B>" & UCase(Replace(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(0), "<TITLE />", "")) & "</B></FONT></TD></TR><TR><TD><BR /></TD></TR><TR>"
			iItemCounter = 0
			bWasLine = True
		ElseIf StrComp(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(0), "<LINE />", vbBinaryCompare) <> 0 Then
			If aMenuComponent(A_ELEMENTS_MENU)(iIndex)(4) <> False Then
			If ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And CLng(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(4))) > 0) Or (CLng(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(4)) = -1) Then
				If Len(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(2)) = 0 Then aMenuComponent(A_ELEMENTS_MENU)(iIndex)(2) = "Images/Transparent.gif"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""64"">"
					Response.Write "<A"
						If Len(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(3)) > 0 Then Response.Write " HREF=""" & aMenuComponent(A_ELEMENTS_MENU)(iIndex)(3) & """"
					Response.Write "><IMG SRC=""" & aMenuComponent(A_ELEMENTS_MENU)(iIndex)(2) & """ WIDTH=""64"" HEIGHT=""64"" ALT=""" & RemoveHTMLFromString(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(0)) & """ BORDER=""0"" /></A>"
				Response.Write "<BR /><BR /></TD>"
				Response.Write "<TD WIDTH=""5"">&nbsp;</TD>"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""286"">"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B><A"
						If Len(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(3)) > 0 Then Response.Write " HREF=""" & aMenuComponent(A_ELEMENTS_MENU)(iIndex)(3) & """"
					Response.Write " CLASS=""SpecialLink"">" & aMenuComponent(A_ELEMENTS_MENU)(iIndex)(0) & "</A></B><BR /></FONT>"
					If aMenuComponent(B_USE_DIV_MENU) Then Response.Write "<DIV CLASS=""MenuOverflow"">"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & aMenuComponent(A_ELEMENTS_MENU)(iIndex)(1) & "</FONT>"
					If aMenuComponent(B_USE_DIV_MENU) Then Response.Write "</DIV>"
				Response.Write "<BR /><BR /></TD>"
				bFirstItem = False

				iItemCounter = iItemCounter + 1
				If (iItemCounter Mod 2) = 0 Then
					Response.Write "</TR><TR>"
				Else
					Response.Write "<TD>&nbsp;</TD>"
				End If
				bWasLine = False
			End If
			End If
		Else
			If (Not bFirstItem) And (iIndex < UBound(aMenuComponent(A_ELEMENTS_MENU))) Then
				If aMenuComponent(A_ELEMENTS_MENU)(iIndex)(4) Then
					If Not bWasLine Then
						For jIndex = (iIndex + 1) To UBound(aMenuComponent(A_ELEMENTS_MENU))
							If aMenuComponent(A_ELEMENTS_MENU)(jIndex)(4) <> False Then
							If ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And CLng(aMenuComponent(A_ELEMENTS_MENU)(jIndex)(4))) > 0) Or (CLng(aMenuComponent(A_ELEMENTS_MENU)(jIndex)(4)) = -1) Then
								Response.Write "</TR><TR><TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """ COLSPAN=""7""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR><TR><TD><BR /></TD></TR><TR>"
								iItemCounter = 0
								bWasLine = True
								Exit For
							End If
							End If
						Next
					End If
				End If
			End If
		End If
	Next
	Response.Write "</TR>" & vbNewLine
	DisplayMenuInTwoColumns = Err.number
	Err.Clear
End Function

Function DisplayMenuInThreeSmallColumns(aMenuComponent)
'************************************************************
'Purpose: To display the menu items in a two columns table
'Inputs:  aMenuComponent
'Outputs: aMenuComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayMenuInThreeSmallColumns"
	Dim bWasLine
	Dim iIndex
	Dim jIndex
	Dim bFirstItem
	Dim iItemCounter
	Dim bComponentInitialized

	bComponentInitialized = aMenuComponent(B_COMPONENT_INITIALIZED_MENU)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeMenuComponent(aMenuComponent)
	End If

	bFirstItem = True
	iItemCounter = 0
	bWasLine = True
	Response.Write vbNewLine & "<TR>"
	For iIndex = 0 To UBound(aMenuComponent(A_ELEMENTS_MENU))
		If aMenuComponent(A_ELEMENTS_MENU)(iIndex)(4) <> False Then
		If ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And CLng(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(4))) > 0) Or (CLng(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(4)) = -1) Then
			If InStr(1, aMenuComponent(A_ELEMENTS_MENU)(iIndex)(0), "<TITLE />", vbBinaryCompare) > 0 Then
				Response.Write "</TR><TR><TD BGCOLOR=""#" & S_WIZARD_TITLE_FOR_GUI & """ COLSPAN=""12""><FONT FACE=""Arial"" SIZE=""2"" COLOR=""#" & S_MENU_LINK_FOR_GUI & """>&nbsp;<B>" & UCase(Replace(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(0), "<TITLE />", "")) & "</B></FONT></TD></TR><TR><TD><BR /></TD></TR><TR>"
				iItemCounter = 0
				bWasLine = True
			ElseIf StrComp(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(0), "<LINE />", vbBinaryCompare) <> 0 Then
				If Len(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(2)) = 0 Then aMenuComponent(A_ELEMENTS_MENU)(iIndex)(2) = "Images/Transparent.gif"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""32"">"
					Response.Write "<A"
						If Len(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(3)) > 0 Then Response.Write " HREF=""" & aMenuComponent(A_ELEMENTS_MENU)(iIndex)(3) & """"
					Response.Write "><IMG SRC=""" & aMenuComponent(A_ELEMENTS_MENU)(iIndex)(2) & """ WIDTH=""32"" HEIGHT=""32"" ALT=""" & RemoveHTMLFromString(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(0)) & """ BORDER=""0"" /></A>"
				Response.Write "<BR /></TD>"
				Response.Write "<TD WIDTH=""3"">&nbsp;</TD>"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""290"">"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B><A"
						If Len(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(3)) > 0 Then Response.Write " HREF=""" & aMenuComponent(A_ELEMENTS_MENU)(iIndex)(3) & """"
					Response.Write " CLASS=""SpecialLink"">" & aMenuComponent(A_ELEMENTS_MENU)(iIndex)(0) & "</A></B><BR /></FONT>"
					If aMenuComponent(B_USE_DIV_MENU) Then Response.Write "<DIV CLASS=""MenuOverflow"">"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & aMenuComponent(A_ELEMENTS_MENU)(iIndex)(1) & "</FONT>"
					If aMenuComponent(B_USE_DIV_MENU) Then Response.Write "</DIV>"
				Response.Write "<BR /></TD>"
				Response.Write "<TD>&nbsp;</TD>"
				bFirstItem = False

				iItemCounter = iItemCounter + 1
				If (iItemCounter Mod 3) = 0 Then
					Response.Write "</TR><TR>"
				End If
				bWasLine = False
			Else
				If (Not bFirstItem) And (iIndex < UBound(aMenuComponent(A_ELEMENTS_MENU))) Then
					If aMenuComponent(A_ELEMENTS_MENU)(iIndex)(4) Then
						If Not bWasLine Then
							For jIndex = (iIndex + 1) To UBound(aMenuComponent(A_ELEMENTS_MENU))
								If aMenuComponent(A_ELEMENTS_MENU)(jIndex)(4) <> False Then
								If ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And CLng(aMenuComponent(A_ELEMENTS_MENU)(jIndex)(4))) > 0) Or (CLng(aMenuComponent(A_ELEMENTS_MENU)(jIndex)(4)) = -1) Then
									Response.Write "</TR><TR><TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """ COLSPAN=""12""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR><TR><TD><BR /></TD></TR><TR>"
									iItemCounter = 0
									bWasLine = True
									Exit For
								End If
								End If
							Next
						End If
					End If
				End If
			End If
		End If
		End If
	Next
	Response.Write "</TR>" & vbNewLine
	DisplayMenuInThreeSmallColumns = Err.number
	Err.Clear
End Function

Function DisplayMenuPopup(bShowDropDownBox, aMenuComponent)
'************************************************************
'Purpose: To display the menu items as a popup menu
'Inputs:  bShowDropDownBox, aMenuComponent
'Outputs: aMenuComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayMenuPopup"
	Dim bFirstItem
	Dim sBGColor
	Dim bWasLine
	Dim iIndex
	Dim jIndex
	Dim bComponentInitialized

	bComponentInitialized = aMenuComponent(B_COMPONENT_INITIALIZED_MENU)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeMenuComponent(aMenuComponent)
	End If

	If bShowDropDownBox Then
		Response.Write vbNewLine & "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD BGCOLOR=""#" & S_BORDER_MENU & """ COLSPAN=""4""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD BGCOLOR=""#" & S_BORDER_MENU & """ WIDTH=""1""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
				Response.Write "<TD WIDTH=""1"">"
					Response.Write "<A HREF=""javascript: ToogleImage('MenuArrow" & aMenuComponent(S_POPUP_DIV_NAME_MENU) & "Img', 'Images/BtnArrRight.gif', 'Images/BtnArrDown.gif'); TogglePopupMenu('" & aMenuComponent(S_POPUP_DIV_NAME_MENU) & "', document." & aMenuComponent(S_POPUP_DIV_NAME_MENU) & ", false);"
						If Len(aMenuComponent(S_TITLE_ACTION_MENU)) > 0 Then
							Response.Write " " & aMenuComponent(S_TITLE_ACTION_MENU)
						End If
					Response.Write """>"
						Response.Write "<IMG SRC=""Images/BtnArrRight.gif"" WIDTH=""13"" HEIGHT=""13"" ALT="""" BORDER=""0"" NAME=""MenuArrow" & aMenuComponent(S_POPUP_DIV_NAME_MENU) & "Img"" />"
					Response.Write "</A>"
				Response.Write "</TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"" STYLE=""font-size: 13px""><NOBR>&nbsp;"
					Response.Write "<A HREF=""javascript: ToogleImage('MenuArrow" & aMenuComponent(S_POPUP_DIV_NAME_MENU) & "Img', 'Images/BtnArrRight.gif', 'Images/BtnArrDown.gif'); TogglePopupMenu('" & aMenuComponent(S_POPUP_DIV_NAME_MENU) & "', document." & aMenuComponent(S_POPUP_DIV_NAME_MENU) & ", false);"
						If Len(aMenuComponent(S_TITLE_ACTION_MENU)) > 0 Then
							Response.Write " " & aMenuComponent(S_TITLE_ACTION_MENU)
						End If
					Response.Write """ CLASS=""SpecialLink"" STYLE=""width: 100%"">" & aMenuComponent(S_TITLE_MENU) & "</A>"
				Response.Write "&nbsp;&nbsp;</FONT></NOBR></TD>"
				Response.Write "<TD BGCOLOR=""#" & S_BORDER_MENU & """ WIDTH=""1""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD BGCOLOR=""#" & S_BORDER_MENU & """ COLSPAN=""4""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>" & vbNewLine
	End If

	bFirstItem = True
	Response.Write vbNewLine & "<DIV ID=""" & aMenuComponent(S_POPUP_DIV_NAME_MENU) & """ CLASS=""ClassPopupItem"" STYLE=""left: " & aMenuComponent(N_LEFT_FOR_DIV_MENU) & "px; top: " & aMenuComponent(N_TOP_FOR_DIV_MENU) & "px;"""
	If Not bShowDropDownBox Then
		Response.Write " onMouseOver=""ShowPopupItem('" & aMenuComponent(S_POPUP_DIV_NAME_MENU) & "', document." & aMenuComponent(S_POPUP_DIV_NAME_MENU) & ")"""
		Response.Write " onMouseOut=""HidePopupItem('" & aMenuComponent(S_POPUP_DIV_NAME_MENU) & "', document." & aMenuComponent(S_POPUP_DIV_NAME_MENU) & ")"""
	End If
	Response.Write ">"
		bWasLine = False
		Response.Write "<TABLE BGCOLOR=""#FFFFFF"" BORDER=""0"" WIDTH=""" & aMenuComponent(N_WIDTH_FOR_DIV_MENU) & """ CELLPADDING=""0"" CELLSPACING=""0"" STYLE=""filter: progid:DXImageTransform.Microsoft.Shadow(color='#666666', Direction='135', Strength='3');"">"
			Response.Write "<TR>"
				Response.Write "<TD BGCOLOR=""#" & S_BORDER_MENU & """ COLSPAN=""5""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
			Response.Write "</TR>"
			For iIndex = 0 To UBound(aMenuComponent(A_ELEMENTS_MENU))
				If aMenuComponent(A_ELEMENTS_MENU)(iIndex)(4) <> False Then
'				If ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And CLng(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(4))) = CLng(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(4))) Or (CLng(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(4)) = -1) Then
				If ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And CLng(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(4))) > 0) Or (CLng(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(4)) = -1) Then
					If StrComp(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(0), "<LINE />", vbBinaryCompare) <> 0 Then
						If Len(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(2)) = 0 Then aMenuComponent(A_ELEMENTS_MENU)(iIndex)(2) = "Images/Transparent.gif"
						Response.Write "<TR"
							If Len(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(3)) > 0 Then
								Response.Write " onMouseOver=""SwitchItemBGColor(this, '" & S_SELECTED_BGCOLOR_MENU & "')"" onMouseOut=""SwitchItemBGColor(this, '')"""
							End If
						Response.Write ">"
							Response.Write "<TD BGCOLOR=""#" & S_BORDER_MENU & """ WIDTH=""1""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
							Response.Write "<TD BGCOLOR=""#" & S_LIGHT_BGCOLOR_MENU & """ WIDTH=""1"">"
								Response.Write "<A"
									If Len(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(3)) > 0 Then
										Response.Write " HREF=""" & aMenuComponent(A_ELEMENTS_MENU)(iIndex)(3) & """"
										If aMenuComponent(B_CLOSE_DIV_AFTER_ACTION_MENU) Then
											Response.Write " onClick=""SwapImage('MenuArrow" & aMenuComponent(S_POPUP_DIV_NAME_MENU) & "Img', 'Images/BtnArrRight.gif'); HidePopupItem('" & aMenuComponent(S_POPUP_DIV_NAME_MENU) & "', document." & aMenuComponent(S_POPUP_DIV_NAME_MENU) & ")"""
										End If
									End If
								Response.Write ">"
									If Len(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(2)) > 0 Then
										Response.Write "<IMG SRC=""" & aMenuComponent(A_ELEMENTS_MENU)(iIndex)(2) & """ WIDTH=""10"" HEIGHT=""8"" ALT=""" & RemoveHTMLFromString(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(0)) & """ BORDER=""0"" />"
									Else
										Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""8"" HEIGHT=""16"" BORDER=""0"" />"
									End If
								Response.Write "</A>"
							Response.Write "</TD>"
							sBGColor = ""
							If (StrComp(GetASPFileName(""), GetASPFileName(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(3)), vbBinaryCompare) = 0) And (InStr(1, Replace((GetASPFileName("") & "?" & oRequest & "&"), "?&", "&", 1, -1, vbBinaryCompare), Replace((aMenuComponent(A_ELEMENTS_MENU)(iIndex)(3) & "&"), "?&", "&", 1, -1, vbBinaryCompare), vbTextCompare) > 0) Then
								sBGColor = " BGCOLOR=""#" & S_LIGHT_BGCOLOR_MENU & """"
							End If
							Response.Write "<TD" & sBGColor & "><FONT SIZE=""1"" STYLE=""font-size: 10px"">&nbsp;</FONT></TD>"
							Response.Write "<TD" & sBGColor & "><NOBR>&nbsp;"
								Response.Write "<A"
									If Len(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(3)) > 0 Then
										Response.Write " HREF=""" & aMenuComponent(A_ELEMENTS_MENU)(iIndex)(3) & """ CLASS=""SpecialLink"""
										If aMenuComponent(B_CLOSE_DIV_AFTER_ACTION_MENU) Then
											Response.Write " onClick=""SwapImage('MenuArrow" & aMenuComponent(S_POPUP_DIV_NAME_MENU) & "Img', 'Images/BtnArrRight.gif'); HidePopupItem('" & aMenuComponent(S_POPUP_DIV_NAME_MENU) & "', document." & aMenuComponent(S_POPUP_DIV_NAME_MENU) & ")"""
										End If
									End If
								Response.Write " STYLE=""width:100%;""><FONT FACE=""Arial"" SIZE=""2"""
									If ((Len(aMenuComponent(A_ELEMENTS_MENU)(iIndex)(3)) = 0) Or (InStr(1, aMenuComponent(A_ELEMENTS_MENU)(iIndex)(2), "Dis.", vbTextCompare) > 1)) Then
										Response.Write " COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """"
									End If
								Response.Write " STYLE=""font-size: 13px"">" & aMenuComponent(A_ELEMENTS_MENU)(iIndex)(0) & "</FONT></A>"
							Response.Write "&nbsp;&nbsp;</NOBR></TD>"
							Response.Write "<TD BGCOLOR=""#" & S_BORDER_MENU & """ WIDTH=""1""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
						Response.Write "</TR>"
						bFirstItem = False
						bWasLine = False
					Else
						If (Not bFirstItem) And (iIndex < UBound(aMenuComponent(A_ELEMENTS_MENU))) Then
							If Not bWasLine Then
								For jIndex = (iIndex + 1) To UBound(aMenuComponent(A_ELEMENTS_MENU))
									If aMenuComponent(A_ELEMENTS_MENU)(jIndex)(4) <> False Then
									If ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And CLng(aMenuComponent(A_ELEMENTS_MENU)(jIndex)(4))) > 0) Or (CLng(aMenuComponent(A_ELEMENTS_MENU)(jIndex)(4)) = -1) Then
										Response.Write "<TR>"
											Response.Write "<TD BGCOLOR=""#" & S_BORDER_MENU & """ WIDTH=""1""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
											Response.Write "<TD BGCOLOR=""#" & S_LIGHT_BGCOLOR_MENU & """ WIDTH=""1""><FONT SIZE=""1"" STYLE=""font-size: 10px"">&nbsp;</FONT></TD>"
											Response.Write "<TD><FONT SIZE=""1"" STYLE=""font-size: 10px"">&nbsp;</FONT></TD>"
											Response.Write "<TD><IMG SRC=""Images/DotGray.gif"" WIDTH=""100%"" HEIGHT=""1"" /></TD>"
											Response.Write "<TD BGCOLOR=""#" & S_BORDER_MENU & """ WIDTH=""1""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
										Response.Write "</TR>"
										bWasLine = True
										Exit For
									End If
									End If
								Next
							End If
							bWasLine = True
						End If
					End If
				End If
				End If
			Next
			Response.Write "<TR>"
				Response.Write "<TD BGCOLOR=""#" & S_BORDER_MENU & """ COLSPAN=""5""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	Response.Write "</DIV>" & vbNewLine

	DisplayMenuPopup = Err.number
	Err.Clear
End Function
%>