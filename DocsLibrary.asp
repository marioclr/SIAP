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
<!-- #include file="Libraries/CatalogComponent.asp" -->
<!-- #include file="Libraries/CatalogsLib.asp" -->
<!-- #include file="Libraries/Main_ISSSTELib.asp" -->
<%
Dim sNames
Dim bAction
Dim bSearchForm
Dim bShowForm
Dim bError

aCatalogComponent(S_TABLE_NAME_CATALOG) = "DocsLibrary"
lErrorNumber = DoCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, bSearchForm, bShowForm, sErrorDescription)
bError = (lErrorNumber <> 0)

aHeaderComponent(L_SELECTED_OPTION_HEADER) = DOCS_TOOLBAR
aHeaderComponent(S_TITLE_NAME_HEADER) = "Normateca"
bWaitMessage = False
Response.Cookies("SoS_SectionID") = 197
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%If ((aLoginComponent(N_USER_PERMISSIONS4_LOGIN) And N_31_PERMISSIONS4) = N_31_PERMISSIONS4) Then
			aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Agregar un nuevo documento",_
					  "",_
					  "", "DocsLibrary.asp?New=1", True)_
			)
			aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 783
			aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
			aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 210
		End If%>
		<!-- #include file="_Header.asp" -->
		<%Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> > <B>Normateca</B><BR />"
		If Len(sLastReport) > 0 Then
			Response.Write "Último reporte: <A HREF=""" & sLastReport & """>" & GetReportNameByConstant(GetParameterFromURLString(sLastReport, "ReportID")) & "</A><BR />"
		End If
		Response.Write "<BR /><BR />"

		If bShowForm Then
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "function IsFileReady(oForm) {" & vbNewLine
					Response.Write "if (oForm.FilePath.value == '') {" & vbNewLine
						Response.Write "alert('No se ha guardado el documento');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "} else {" & vbNewLine
						Response.Write "oForm.FileType.value = oForm.FilePath.value.substr(oForm.FilePath.value.search(/\./gi) + 1);" & vbNewLine
					Response.Write "}" & vbNewLine

					Response.Write "return true;" & vbNewLine
				Response.Write "} // End of IsFileReady" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
						
			If Len(oRequest("Delete").Item) = 0 Then Response.Write "<IFRAME SRC=""BrowserFileForInfo.asp?Action=DocsLibrary&UserID=" & aLoginComponent(N_USER_ID_LOGIN) & """ NAME=""UploadInfoIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""60""></IFRAME><BR />"
			Response.Write "<B>Registro de documentos</B><BR />"
			lErrorNumber = DisplayCatalogForm(oRequest, oADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)
		Else
			lErrorNumber = Display371SearchResults(oRequest, oADODBConnection, True, sErrorDescription)
			If lErrorNumber <> 0 Then
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
				lErrorNumber = 0
				Response.Write "<BR />"
			End If
		End If%>
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>