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
<!-- #include file="Libraries/BudgetLib.asp" -->
<!-- #include file="Libraries/BudgetComponent.asp" -->
<%
Dim sSection
Dim bAction
Dim bError
Dim sNames
Dim aPath
Dim aTempBudgetComponent()

Call InitializeBudgetComponent(oRequest, aBudgetComponent)
Call GetBudgetURLValues(oRequest, sSection, bAction, aBudgetComponent(S_QUERY_CONDITION_BUDGET))

bError = False
If bAction Then
	lErrorNumber = DoBudgetAction(oRequest, oADODBConnection, sSection, oRequest("Action").Item, sErrorDescription)
	bError = (lErrorNumber <> 0)
End If
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
		<SCRIPT LANGUAGE="JavaScript"><!--
			var bReady = false;
		//--></SCRIPT>
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<DIV ID="WaitDiv" CLASS="ClassPopupItem" STYLE="top: 200px; visibility: visible;">
			<TABLE WIDTH="100%" BORDER="0" CELLPADDING="0" CELLSPACING="0"><TR><TD ALIGN="CENTER">
				<IMG SRC="Images/AniWait.gif" WIDTH="100" HEIGHT="100" ALT="Cargando información..." /><BR /><BR />
				<FONT FACE="Arial" SIZE="2"><B>Cargando información...</B></FONT>
			</TD></TR></TABLE>
		</DIV>
		<%Response.Flush()
		If (lErrorNumber <> 0) Or bError Then
			Response.Write "<BR />"
			Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
			lErrorNumber = 0
			Response.Write "<BR />"
		End If

		If (Len(oRequest("View").Item) = 0) And (Len(oRequest("Modify").Item) = 0) Then
			lErrorNumber = DisplayMoneySearchForm(oRequest, oADODBConnection, GetASPFileName(""), aBudgetComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
					Response.Write "<DIV NAME=""EntriesDiv"" ID=""EntriesDiv"" STYLE=""width: 460px; height: 300px; overflow: auto;"">"
						lErrorNumber = DisplayMoneyTable(oRequest, oADODBConnection, (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)), aBudgetComponent, sErrorDescription)
					Response.Write "</DIV>"
				End If
			End If
		Else
			lErrorNumber = DisplayMoneyForm(oRequest, oADODBConnection, GetASPFileName(""), aBudgetComponent, sErrorDescription)
		End If
		If lErrorNumber <> 0 Then
			Response.Write "<BR /><BR />"
			Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
			lErrorNumber = 0
			Response.Write "<BR />"
		End If%>
		<SCRIPT LANGUAGE="JavaScript"><!--
			HidePopupItem('WaitDiv', document.WaitDiv);
			<%If (Len(oRequest("View").Item) > 0) Or (Len(oRequest("Modify").Item) > 0) Then%>
			bReady = true;
			<%End If%>
		//--></SCRIPT>
	</BODY>
</HTML>