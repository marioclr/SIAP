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
<!-- #include file="Libraries/XMLLibrary.asp" -->
<%
Dim iIndex

Call InitializeErrorLogComponent(oRequest, aErrorLogComponent)

aHeaderComponent(L_SELECTED_OPTION_HEADER) = TOOLS_TOOLBAR
aHeaderComponent(S_TITLE_NAME_HEADER) = "Bitácora de Errores"
Response.Cookies("SoS_SectionID") = 198
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
		<SCRIPT LANGUAGE="JavaScript"><!--
			function BuildDateStringForFile(oForm) {
				var sMonth = '';
				sMonth = '0' + (parseInt(GetSelectedItems(oForm.MonthForFile)) + 1);
				sMonth = sMonth.substr((sMonth.length - 2));
				oForm.ErrorLogDate.value = oForm.YearForFile.value + sMonth + oForm.DayForFile.value;
			}
		//--></SCRIPT>
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<!-- #include file="_Header.asp" -->
		<!-- BEGIN: CONTENTS -->
		Usted se encuentra aquí: <A HREF="Main.asp">Inicio</A> > <A HREF="Tools.asp">Herramientas</A> > <B>Bitácora de errores</B><BR /><BR />
		<TABLE WIDTH="350" BORDER="0" CELLPADDING="1" CELLSPACING="0"><TR><TD BGCOLOR="#<%Response.Write S_WIDGET_FRAME_FOR_GUI%>">
			<TABLE WIDTH="100%" BORDER="0" CELLPADDING="0" CELLSPACING="0"><TR>
				<TD BGCOLOR="#<%Response.Write S_WIDGET_BGCOLOR_FOR_GUI%>" WIDTH="1">
					<A HREF="javascript: ToogleImage('ErrorLogFormImg', 'Images/BtnArrRight.gif', 'Images/BtnArrDown.gif'); TogglePopupMenu('ErrorLogFormDiv', document.ErrorLogFormDiv, false)" CLASS="SpecialLink"><IMG SRC="Images/BtnArrRight.gif" WIDTH="13" HEIGHT="13" BORDER="0" NAME="ErrorLogFormImg" /></A>
				</TD>
				<TD BGCOLOR="#<%Response.Write S_WIDGET_BGCOLOR_FOR_GUI%>">
					<FONT FACE="Arial" SIZE="2">&nbsp;<A HREF="javascript: ToogleImage('ErrorLogFormImg', 'Images/BtnArrRight.gif', 'Images/BtnArrDown.gif'); TogglePopupMenu('ErrorLogFormDiv', document.ErrorLogFormDiv, false)" CLASS="SpecialLink">Bitácora del día <%Response.Write Mid(aErrorLogComponent(S_DATE_ERROR_LOG), Len("YYYYMMD"), Len("DD")) & " de " & asMonthNames_es(Mid(aErrorLogComponent(S_DATE_ERROR_LOG), Len("YYYYM"), Len("MM"))) & " de " & Left(aErrorLogComponent(S_DATE_ERROR_LOG), Len("YYYY"))%></A></FONT>
				</TD>
			</TR></TABLE>
		</TD></TR></TABLE>

		<DIV ID="ErrorLogFormDiv" CLASS="ClassPopupItem" STYLE="z-index: 99;">
			<TABLE WIDTH="350" BORDER="0" CELLPADDING="1" CELLSPACING="0"><TR><TD BGCOLOR="#<%Response.Write S_WIDGET_FRAME_FOR_GUI%>">
				<TABLE WIDTH="100%" BORDER="0" CELLPADDING="0" CELLSPACING="0"><TR><TD BGCOLOR="#FFFFFF">
					<FORM NAME="ErrorLogFrm" ID="ErrorLogFrm" ACTION="ErrorLog.asp" METHOD="GET" onSubmit="BuildDateStringForFile(this)">
						<INPUT TYPE="HIDDEN" NAME="ErrorLogDate" ID="ErrorLogDateHdn" VALUE="" />
						<INPUT TYPE="HIDDEN" NAME="LogFolder" ID="LogFolderHdn" VALUE="<%Response.Write aErrorLogComponent(S_FOLDER_LOG)%>" />
						<TABLE WIDTH="100%" BORDER="0" CELLPADDING="0" CELLSPACING="0">
							<TR>
								<TD ROWSPAN="6"><IMG SRC="Images/Transparent.gif" WIDTH="20" HEIGHT="1" /></TD>
								<TD BGCOLOR="#<%Response.Write S_WIDGET_BGCOLOR_FOR_GUI%>" COLSPAN="2"><FONT FACE="Arial" SIZE="2">Fecha&nbsp;</FONT></TD>
							</TR>
							<TR>
								<TD><IMG SRC="Images/Transparent.gif" WIDTH="20" HEIGHT="1" /></TD>
								<TD><%
									Response.Write DisplayDateCombos(CInt(Mid(aErrorLogComponent(S_DATE_ERROR_LOG), 1, 4)), CInt(Mid(aErrorLogComponent(S_DATE_ERROR_LOG), 5, 2)), CInt(Mid(aErrorLogComponent(S_DATE_ERROR_LOG), 7, 2)), "YearForFile", "MonthForFile", "DayForFile", N_START_YEAR, Year(Date()), True, False)
									If Not bIsNetscape Then
										Response.Write "<FONT SIZE=""1""><BR /><BR /></FONT>"
										Response.Write "<IFRAME SRC=""BrowserMonth.asp?LogFolder=" & aErrorLogComponent(S_FOLDER_LOG) & "&ErrorLogDate=" & aErrorLogComponent(S_DATE_ERROR_LOG) & """ NAME=""BrowserMonthIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""112""></IFRAME>"
									End If
								%></TD>
							</TR>
							<TR><TD COLSPAN="2"><IMG SRC="Images/Transparent.gif" WIDTH="1" HEIGHT="5" /></TD></TR>
							<TR><TD BGCOLOR="#000000" COLSPAN="2"><IMG SRC="Images/Transparent.gif" WIDTH="1" HEIGHT="1" /></TD></TR>
							<TR><TD BGCOLOR="#<%Response.Write S_WIDGET_BGCOLOR_FOR_GUI%>" COLSPAN="2"><FONT FACE="Arial" SIZE="2">Tipo de mensajes&nbsp;</FONT></TD></TR>
							<TR>
								<TD><IMG SRC="Images/Transparent.gif" WIDTH="20" HEIGHT="1" /></TD>
								<TD><FONT FACE="Arial" SIZE="2"><%
									For iIndex = 0 To UBound(aMessagesTypes)
										Response.Write "<INPUT TYPE=""CHECKBOX"" NAME=""ErrorLogLevel"" HIDDEN=""ErrorLogLevelHdn"" VALUE=""" & (2 ^ iIndex) & """"
											If (aErrorLogComponent(N_SHOW_LEVEL_ERROR_LOG) And (2 ^ iIndex)) <> 0 Then Response.Write " CHECKED=""1"""
										Response.Write " />"
										Response.Write "<IMG SRC=""Images/IcnErrorLevel" & (2 ^ iIndex) & ".gif"" WIDTH=""16"" HEIGHT=""16"" /> " & aMessagesTypes(iIndex) & "<BR />"
									Next
								%></FONT></TD>
							</TR>
						</TABLE><BR />&nbsp;&nbsp;&nbsp;
						<INPUT TYPE="SUBMIT" VALUE="Ver Bitácora" CLASS="Buttons" />
					</FORM>
				</TD></TR></TABLE>
			</TD></TR></TABLE>
		</DIV><BR />

		<%lErrorNumber = DisplayLogFile(aErrorLogComponent, sErrorDescription)
		If lErrorNumber <> 0 Then
			Call DisplayErrorMessage("Bitácora de errores", sErrorDescription)
		End If%>
		<BR />
		<!-- END: CONTENTS -->
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>