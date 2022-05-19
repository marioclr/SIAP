<%
Function DisplayHelpMenu(iSelectedSection, sErrorDescription)
'************************************************************
'Purpose: To display the sections for the online help.
'Inputs:  iSelectedSection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayHelpMenu"
	Dim sHelpContents
	Dim sBoldBegin
	Dim sBoldEnd
	Dim sTab
	Dim oXML
	Dim oXMLNode
	Dim oXMLSectionNode
	Dim bDisplay
	Dim iIndex
	Dim lErrorNumber

	sHelpContents = GetFileContents(Server.MapPath("Help\Help.xml"), sErrorDescription)
	If lErrorNumber = 0 Then
		lErrorNumber = CreateXMLDOMObject(oXML, sErrorDescription)
		If lErrorNumber = 0 Then
			lErrorNumber = LoadXMLToObject(sHelpContents, oXML, sErrorDescription)
			If lErrorNumber = 0 Then
				For Each oXMLNode In oXML.documentElement.selectNodes("/HELP/TOPIC")
					bDisplay = True
					If (CLng(oXMLNode.getAttribute("PERMISSION")) = -1) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And CLng(oXMLNode.getAttribute("PERMISSION"))) > 0) Then
					If bDisplay Then
						If CInt(oXMLNode.getAttribute("ID")) = iSelectedSection Then
							sBoldBegin = "<B>"
							sBoldEnd = "</B>"
						Else
							sBoldBegin = ""
							sBoldEnd = ""
						End If
						If oXMLNode.hasChildNodes() Then
							Response.Write "<IMG SRC=""Images/BtnExpandSmall.gif"" WIDTH=""9"" HEIGHT=""9"" />"
						Else
							Response.Write "<IMG SRC=""Images/Bullet.gif"" WIDTH=""9"" HEIGHT=""9"" />"
						End If
						sTab = ""
						If Not IsNull(oXMLNode.getAttribute("TAB")) Then
							For iIndex = 0 To CInt(oXMLNode.getAttribute("TAB"))
								sTab = sTab & "&nbsp;&nbsp;&nbsp;"
							Next
						End If
						Response.Write sBoldBegin & sTab & "<A HREF=""Help.asp?HelpSection=" & oXMLNode.getAttribute("ID") & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & """>" & oXMLNode.getAttribute("TITLE") & "</A><BR />" & sBoldEnd
						If oXMLNode.hasChildNodes() Then
							Set oXMLSectionNode = oXMLNode.firstChild
							Do While Not (oXMLSectionNode Is Nothing)
								If (CLng(oXMLSectionNode.getAttribute("PERMISSION")) = -1) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And CLng(oXMLSectionNode.getAttribute("PERMISSION"))) > 0) Then
									If CInt(oXMLSectionNode.getAttribute("ID")) = iSelectedSection Then
										sBoldBegin = "<B>"
										sBoldEnd = "</B>"
									Else
										sBoldBegin = ""
										sBoldEnd = ""
									End If
									sTab = ""
									If Not IsNull(oXMLSectionNode.getAttribute("TAB")) Then
										For iIndex = 0 To CInt(oXMLSectionNode.getAttribute("TAB"))
											sTab = sTab & "&nbsp;&nbsp;&nbsp;"
										Next
									End If
									Response.Write "&nbsp;&nbsp;&nbsp;<IMG SRC=""Images/Bullet.gif"" WIDTH=""9"" HEIGHT=""9"" />"
									Response.Write sBoldBegin & sTab & "<A HREF=""Help.asp?HelpSection=" & oXMLSectionNode.getAttribute("ID") & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & """>" & oXMLSectionNode.getAttribute("TITLE") & "</A><BR />" & sBoldEnd
								End If
								Set oXMLSectionNode = oXMLSectionNode.nextSibling
								If Err.number <> 0 Then Exit Do
							Loop
						End If
					End If
					End If
					If Err.number <> 0 Then Exit For
				Next
			End If
		End If
	End If

	DisplayHelpMenu = lErrorNumber
	Err.Clear
End Function

Function DisplayHelpSection(iHelpSection, sErrorDescription)
'************************************************************
'Purpose: To display the section for the online help.
'Inputs:  iHelpSection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayHelpSection"
	Dim sHelpContents
	Dim lErrorNumber

	If FileExists(Server.MapPath("Help\Help_" & Right(("000" & iHelpSection), Len("000")) & ".htm"), sErrorDescription) Then
		sHelpContents = GetFileContents(Server.MapPath("Help\Help_" & Right(("000" & iHelpSection), Len("000")) & ".htm"), sErrorDescription)
		If lErrorNumber = 0 Then
			Response.Write Replace(sHelpContents, "<ACCESS_KEY />", aLoginComponent(S_ACCESS_KEY_LOGIN))
		End If
	End If

	DisplayHelpSection = lErrorNumber
	Err.Clear
End Function
%>