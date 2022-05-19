<%
Function LaunchIntro()
'************************************************************
'Purpose: To launch the intro
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "LaunchIntro"
	Dim sCookie

	If B_DISPLAY_INTRO Then
		sCookie = ""
		sCookie = Request.Cookies("SIAP_Intro")
		Err.Clear
		If StrComp(sCookie, "1", vbBinaryCompare) <> 0 Then
			Response.Cookies("SIAP_Intro") = "1"
			Response.Cookies("SIAP_Intro").Expires = #1/1/2038#
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "OpenNewWindow('Curso\\registroPNE.htm', '', 'Intro', 994, 668, 'no', 'no')" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
		End If
		Response.Write "<SPAN CLASS=""HelpIcon""><B>"
			Response.Write "<IMG SRC=""Images/IcnHelpBig.gif"" WIDTH=""32"" HEIGHT=""32"" BORDER=""0"" ALIGN=""ABSMIDDLE"" />"
			Response.Write "<A HREF=""javascript: OpenNewWindow('Curso\\ayuda.htm', '', 'Intro', 994, 668, 'no', 'no')""><FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>Deseo ver la introducción al sistema</B></FONT></A>" & vbNewLine
		Response.Write "</B></SPAN><BR /><BR />"
		'Response.Write "<IMG SRC=""Images/IcnHelp.gif"" WIDTH=""16"" HEIGHT=""16"" />&nbsp;"
		'Response.Write "<A HREF=""javascript: OpenNewWindow('Curso\\registroPNE.htm', '', 'Intro', 994, 668, 'no', 'no')""><FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>Deseo ver la introducción al sistema</B></FONT></A><BR /><BR />" & vbNewLine
	End If
	Response.Write "<SPAN CLASS=""HelpIcon""><B>"
		Response.Write "<IMG SRC=""Images/IcnHelpBig.gif"" WIDTH=""32"" HEIGHT=""32"" BORDER=""0"" ALIGN=""ABSMIDDLE"" />"
		Response.Write "<A HREF=""javascript: OpenNewWindow('Help.asp?HelpSection=" & iHelpSection & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "', null, 'Help', 720, 720, 'no', 'no')""><FONT COLOR=""#" & S_SELECTED_LINK_FOR_GUI & """><B>Deseo ver la ayuda en línea</B></FONT></A>" & vbNewLine
	Response.Write "</B></SPAN>"
	'Response.Write "<IMG SRC=""Images/IcnHelp.gif"" WIDTH=""16"" HEIGHT=""16"" />&nbsp;"
	'Response.Write "<A HREF=""javascript: OpenNewWindow('Help.asp?HelpSection=" & iHelpSection & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "', null, 'Help', 720, 720, 'no', 'no')""><FONT COLOR=""#" & S_SELECTED_LINK_FOR_GUI & """><B>Deseo ver la ayuda en línea</B></FONT></A>" & vbNewLine

	LaunchIntro = Err.number
	Err.Clear
End Function

Function LogCommonErrors(oRequest, sErrorDescription)
'************************************************************
'Purpose: To check if user is already logged and notify if
'         there was any error while trying to connect the user
'Inputs:  oRequest
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ValidateUserConnection"
	Dim lErrorNumber

	lErrorNumber = 0
	If Len(oRequest("InvalidLicense").Item) > 0 Then
		lErrorNumber = -1
		sErrorDescription = "La licencia para el uso de este sistema no está correcta. Favor de contactar al dueño del producto."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "DefaultLib.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	ElseIf Len(oRequest("ExpiredLicense").Item) > 0 Then
		lErrorNumber = -1
		sErrorDescription = "La licencia para el uso de este sistema ha expirado. Favor de contactar al dueño del producto."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "DefaultLib.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	ElseIf Len(oRequest("InvalidUser").Item) > 0 Then
		lErrorNumber = -1
		sErrorDescription = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>Se perdió la conexión con el sistema.</B></FONT> Favor de introducir la clave de acceso y la contraseña."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "DefaultLib.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
		Response.Cookies("SIAP_CurrentAccessKey").Item = ""
	ElseIf Len(oRequest("SessionExpired").Item) > 0 Then
		lErrorNumber = -1
		sErrorDescription = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """><B>La sesión expiró.</B></FONT> Favor de conectarse de nuevo."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "DefaultLib.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	End If

	LogCommonErrors = lErrorNumber
	Err.Clear
End Function

Function ShowFlashPlayerMessage()
'************************************************************
'Purpose: To launch the intro
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ShowFlashPlayerMessage"

	Response.Write "<SPAN CLASS=""HelpIcon"" STYLE=""background-color: #FFFFFF""><IMG SRC=""Images/IcnHelp.gif"" WIDTH=""16"" HEIGHT=""16"" /> <B>¿Ya tiene Flash Player instalado?</B><BR />"
	Response.Write "<TABLE BORDER=""0"" WIDTH=""250"" CELLSPACING=""0"" CELLPADDING=""0""><TR>"
		Response.Write "<TD>"
'			Response.Write "<OBJECT CLASSID=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" CODEBASE="" & S_FLASH_PLAYER_URL & "" WIDTH=""60"" HEIGHT=""30"" BGCOLOR=""#FFFFFF"">"
'				Response.Write "<PARAM NAME=""MOVIE"" VALUE=""SWF/FlashPlayer.swf"">"
'				Response.Write "<PARAM NAME=""QUALITY"" VALUE=""HIGH"">"
'				Response.Write "<PARAM NAME=""BGCOLOR"" VALUE=""#FFFFFF"">"
'				Response.Write "<PARAM NAME=""MENU"" VALUE=""FALSE"">"
'				Response.Write "<EMBED SRC=""SWF/FlashPlayer.swf"" QUALITY=""HIGH"" PLUGINSPAGE=""http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"" TYPE=""application/x-shockwave-flash"" WIDTH=""60"" HEIGHT=""30"" MENU=""FALSE"" BGCOLOR=""#FFFFFF"">"
'				Response.Write "</EMBED>"
'			Response.Write "</OBJECT>"
		Response.Write "</TD>"
		Response.Write "<TD>&nbsp;&nbsp;&nbsp;</TD>"
'		Response.Write "<TD><FONT FACE=""Arial"" SIZE=""1"">Si usted no ve la animación, <A HREF=""http://www.adobe.com/flashplayer"" TARGET=""FlashPlayer"">presione aquí</A>.</FONT></TD>"
		Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Si no es así, <A HREF=""http://www.adobe.com/flashplayer"" TARGET=""FlashPlayer"">presione aquí</A>.</FONT></TD>"
	Response.Write "</TR></TABLE></SPAN>"

	ShowFlashPlayerMessage = Err.number
	Err.Clear
End Function
%>