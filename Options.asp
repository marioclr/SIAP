<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
%>
<!-- #include file="Libraries/GlobalVariables.asp" -->
<!-- #include file="Libraries/LoginComponent.asp" -->
<%
Dim sTitle
Dim bAdminOptions
Dim sFileContents

bAdminOptions = Len(oRequest("Admin").Item) > 0
If bAdminOptions Then
	If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_TOOLS_PERMISSIONS Then
	Else
		Response.Redirect "AccessDenied.asp?Permission=" & N_TOOLS_PERMISSIONS
	End If
End If

If Len(oRequest("LoadDefaultValues").Item) > 0 Then
	aAdminOptionsComponent(A_OPTIONS) = Split(DEFAULT_ADMIN_OPTIONS, LIST_SEPARATOR, -1, vbBinaryCompare)
	aOptionsComponent(A_OPTIONS) = Split(DEFAULT_OPTIONS, LIST_SEPARATOR, -1, vbBinaryCompare)
ElseIf (lErrorNumber = 0) And (Len(oRequest("SaveOptions").Item) > 0) Then
	If bAdminOptions Then
		If (StrComp(GetAdminOption(aAdminOptionsComponent, UPDATE_OPTION), oRequest("P0007").Item, vbBinaryCompare) <> 0) Or (StrComp(GetAdminOption(aAdminOptionsComponent, DELETE_OPTION), oRequest("P0008").Item, vbBinaryCompare) <> 0) Or (StrComp(GetAdminOption(aAdminOptionsComponent, INSERT_OPTION), oRequest("P0009").Item, vbBinaryCompare) <> 0) Then
			sFileContents = GetFileContents(Server.MapPath("Libraries\DatabaseLibrary.asp"), sErrorDescription)
			If Len(sFileContents) > 0 Then
				If CInt(oRequest("P0007").Item) = 0 Then
					sFileContents = Replace(sFileContents, "'False|UPDATE_OPTION", "'True|UPDATE_OPTION")
					sFileContents = Replace(sFileContents, "And True Then 'UPDATE_OPTION", "And False Then 'UPDATE_OPTION")
				Else
					sFileContents = Replace(sFileContents, "'True|UPDATE_OPTION", "'False|UPDATE_OPTION")
					sFileContents = Replace(sFileContents, "And False Then 'UPDATE_OPTION", "And True Then 'UPDATE_OPTION")
				End If

				If CInt(oRequest("P0008").Item) = 0 Then
					sFileContents = Replace(sFileContents, "'False|DELETE_OPTION", "'True|DELETE_OPTION")
					sFileContents = Replace(sFileContents, "And True Then 'DELETE_OPTION", "And False Then 'DELETE_OPTION")
				Else
					sFileContents = Replace(sFileContents, "'True|DELETE_OPTION", "'False|DELETE_OPTION")
					sFileContents = Replace(sFileContents, "And False Then 'DELETE_OPTION", "And True Then 'DELETE_OPTION")
				End If

				If CInt(oRequest("P0009").Item) = 0 Then
					sFileContents = Replace(sFileContents, "'False|INSERT_OPTION", "'True|INSERT_OPTION")
					sFileContents = Replace(sFileContents, "And True Then 'INSERT_OPTION", "And False Then 'INSERT_OPTION")
				Else
					sFileContents = Replace(sFileContents, "'True|INSERT_OPTION", "'False|INSERT_OPTION")
					sFileContents = Replace(sFileContents, "And False Then 'INSERT_OPTION", "And True Then 'INSERT_OPTION")
				End If

				lErrorNumber = SaveTextToFile(Server.MapPath("Libraries\DatabaseLibrary.asp"), sFileContents, sErrorDescription)
			End If
		End If
		sTitle = "Ocurrió un error en las opciones del sistema"
		lErrorNumber = SetOptions(oRequest, aAdminOptionsComponent, sErrorDescription)
	Else
		sTitle = "Ocurrió un error en las preferencias del usuario"
		lErrorNumber = SetOptions(oRequest, aOptionsComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		If bAdminOptions Then
			lErrorNumber = ModifyAdminOptions(oRequest, oADODBConnection, aAdminOptionsComponent, sErrorDescription)
		Else
			lErrorNumber = ModifyOptions(oRequest, oADODBConnection, aOptionsComponent, sErrorDescription)
		End If
		If lErrorNumber = 0 Then
			sTitle = "Confirmación"
			If Len(sOptionsErrorDescription) > 0 Then sOptionsErrorDescription = "<B>" & sOptionsErrorDescription & "</B><BR /><BR />"
			If bAdminOptions Then
				sErrorDescription = sOptionsErrorDescription & "Las opciones del sistema fueron guardadas correctamente."
			Else
				sErrorDescription = sOptionsErrorDescription & "Las preferencias del usuario fueron guardadas correctamente."
			End If
		End If
	End If
End If

aHeaderComponent(L_SELECTED_OPTION_HEADER) = TOOLS_TOOLBAR
If bAdminOptions Then
	aHeaderComponent(S_TITLE_NAME_HEADER) = "Opciones del Sistema"
Else
	aHeaderComponent(S_TITLE_NAME_HEADER) = "Preferencias del Usuario"
End If
Response.Cookies("SoS_SectionID") = 201
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<!-- #include file="_Header.asp" -->
		<%Response.Write "Usted se encuentra aquí:&nbsp;<A HREF=""Main.asp"">Inicio</A> > <B>" & aHeaderComponent(S_TITLE_NAME_HEADER) & "</B><BR /><BR />"
		If Len(sErrorDescription) > 0 Then
			Call DisplayErrorMessage(sTitle, sErrorDescription)
		End If%>
		<FORM NAME="OptionsFrm" ID="OptionsFrm" ACTION="Options.asp" METHOD="POST">
			<%If bAdminOptions Then%>
				<!-- #include file="OptionsAdmin.asp" -->
			<%Else%>
				<!-- #include file="OptionsUser.asp" -->
			<%End If%>
			<INPUT TYPE="SUBMIT" NAME="SaveOptions" VALUE="Guardar Preferencias" CLASS="Buttons" />
			<IMG SRC="Images/Transparent.gif" WIDTH="100" HEIGHT="1" />
			<INPUT TYPE="BUTTON" VALUE="Leer Valores Originales" CLASS="Buttons" onClick="window.location.href = 'Options.asp?LoadDefaultValues=1&<%Response.Write oRequest%>'" />
			<IMG SRC="Images/Transparent.gif" WIDTH="100" HEIGHT="1" />
			<INPUT TYPE="BUTTON" VALUE="Cancelar" CLASS="RedButtons" onClick="window.location.href = 'Tools.asp'" />
		</FORM>
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>
<%
Set oADODBConnection = Nothing
%>