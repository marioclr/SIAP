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
<!-- #include file="Libraries/AreasLib.asp" -->
<!-- #include file="Libraries/AreaComponent.asp" -->
<!-- #include file="Libraries/JobsLib.asp" -->
<!-- #include file="Libraries/JobComponent.asp" -->
<%
Dim iSelectedTab
Dim bAction
Dim bError

Call InitializeAreaComponent(oRequest, aAreaComponent)
Call InitializeJobComponent(oRequest, aJobComponent)
Call GetAreasURLValues(oRequest, iSelectedTab, bAction, aAreaComponent(S_QUERY_CONDITION_AREA))

bError = False
If bAction Then
	lErrorNumber = DoAreasAction(oRequest, oADODBConnection, oRequest("Action").Item, sErrorDescription)
	If lErrorNumber <> 0 Then
		bError = True
	Else
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "parent.location.reload();" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If
End If
If lErrorNumber = 0 Then
	If aAreaComponent(N_ID_AREA) > -1 Then
		lErrorNumber = GetArea(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
	End If
	If aJobComponent(N_ID_JOB) > -1 Then
		lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
	End If
End If
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%If lErrorNumber <> 0 Then
			Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
			lErrorNumber = 0
			Response.Write "<BR />"
		End If
		Select Case oRequest("Action").Item
			Case "Areas"
				Call DisplayAreasTabs(oRequest, bError, sErrorDescription)
				Response.Write "<BR />"
				Select Case iSelectedTab
					Case 2
						lErrorNumber = DisplayAreaPositionsForm(oRequest, oADODBConnection, GetASPFileName(""), aAreaComponent, sErrorDescription)
					Case 3
						lErrorNumber = DisplayAreaHistoryList(oRequest, oADODBConnection, False, aAreaComponent, sErrorDescription)
					Case 4
						lErrorNumber = DisplayAreaPositionsHistoryList(oRequest, oADODBConnection, False, aAreaComponent, sErrorDescription)
					Case Else
						If Len(oRequest("ShowInfo").Item) > 0 Then
							lErrorNumber = DisplayArea(oRequest, oADODBConnection, False, aAreaComponent, sErrorDescription)
						Else
							lErrorNumber = DisplayAreaForm(oRequest, oADODBConnection, GetASPFileName(""), aAreaComponent, sErrorDescription)
						End If
				End Select
			Case "Jobs"
				If Len(oRequest("ShowInfo").Item) > 0 Then
					lErrorNumber = DisplayJob(oRequest, oADODBConnection, False, aJobComponent, sErrorDescription)
				Else
					Call DisplayJobsTabs(oRequest, bError, sErrorDescription)
					lErrorNumber = DisplayJobForm(oRequest, oADODBConnection, GetASPFileName(""), aJobComponent, sErrorDescription)
				End If
		End Select
		If lErrorNumber <> 0 Then
			Response.Write "<BR /><BR />"
			Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
			lErrorNumber = 0
			Response.Write "<BR />"
		End If%>
	</BODY>
</HTML>