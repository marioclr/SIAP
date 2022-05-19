	</FONT></TD>
</TR></TABLE>
<!-- BEGIN: FOOTER -->
<FONT SIZE="2"><BR /><BR /></FONT><TABLE WIDTH="994" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<TR><TD ALIGN="CENTER" COLSPAN="2"><IMG SRC="Images/DotBlue.gif" WIDTH="980" HEIGHT="1" /></TD></TR>
	<TR>
		<TD VALIGN="BOTTOM">&nbsp;</TD>
		<TD ALIGN="RIGHT" VALIGN="BOTTOM">
			<FONT FACE="Arial" SIZE="2"><A HREF="About.asp" STYLE="text-decoration: none"><FONT COLOR="#<%Response.Write S_MAIN_COLOR_FOR_GUI%>">Derechos Reservados &#174; <%Response.Write Year(Now())%></FONT></A>&nbsp;&nbsp;</FONT>
		</TD>
	</TR>
</TABLE>
<%If bWaitMessage Then%>
<SCRIPT LANGUAGE="JavaScript"><!--
	HidePopupItem('WaitDiv', document.WaitDiv)
//--></SCRIPT>
<%End If%>
<!-- END: FOOTER -->