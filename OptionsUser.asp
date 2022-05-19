		<FONT FACE="Arial" SIZE="2" COLOR="#<%Response.Write S_INSTRUCTIONS_FOR_GUI%>"><B>Modifique los valores de las preferencias de acuerdo a sus gustos y necesidades.</B></FONT><BR /><BR />
		<!-- PÁGINA DE INICIO -->
		<B>Generales</B><BR />
		<IMG SRC="Images/Transparent.gif" WIDTH="20" HEIGHT="62" ALIGN="LEFT" />
		<FONT FACE="Arial" SIZE="2">Página de inicio: </FONT>
		<SELECT NAME="P0001" SIZE="1" CLASS="Lists">
			<OPTION VALUE="Main.asp">Página principal</OPTION>
			If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REPORTS_PERMISSIONS Then Response.Write "<OPTION VALUE=""Reports.asp"">Reportes</OPTION>"%>
		</SELECT>
		<SCRIPT LANGUAGE="JavaScript"><!--
			SelectItemByValue('<%Response.Write GetOption(aOptionsComponent, START_PAGE_OPTION)%>', false, document.OptionsFrm.P0001);
		//--></SCRIPT>
		<BR /><BR />
		<!-- ESTILO DE LAS TABLAS -->
		<FONT FACE="Arial" SIZE="2">Estilo de las tablas: </FONT>
		<SELECT NAME="P0002" SIZE="1" CLASS="Lists">
			<OPTION VALUE="1">Tablas en 3D</OPTION>
			<OPTION VALUE="2">Tablas sencillas</OPTION>
		</SELECT>
		<SCRIPT LANGUAGE="JavaScript"><!--
			SelectItemByValue('<%Response.Write GetOption(aOptionsComponent, TABLE_STYLE_OPTION)%>', false, document.OptionsFrm.P0002);
		//--></SCRIPT>
		<BR /><BR />

		<B>Impresión</B><BR />
		<IMG SRC="Images/Transparent.gif" WIDTH="20" HEIGHT="34" ALIGN="LEFT" />
		<INPUT TYPE="CHECKBOX" NAME="Dummy0000" ID="Dummy0000Chk"<%
			If CInt(GetOption(aOptionsComponent, SHOW_PRINT_INFO_OPTION)) = 1 Then Response.Write " CHECKED=""1"""
		%> onClick="SetHiddenValueForCheckBox(this.checked, this.form.P0000)" />
		<INPUT TYPE="HIDDEN" NAME="P0000" ID="P0000Hdn" VALUE="<%Response.Write GetOption(aOptionsComponent, SHOW_PRINT_INFO_OPTION)%>" />
		<FONT FACE="Arial" SIZE="2">
			Mostrar las instrucciones para imprimir.
		</FONT><BR /><BR />

		<B>Reportes</B><BR />
		<IMG SRC="Images/Transparent.gif" WIDTH="20" HEIGHT="60" ALIGN="LEFT" />
		<FONT FACE="Arial" SIZE="2">Número de renglones que se mostrarán en el listado de trámites: </FONT>
		<INPUT TYPE="TEXT" NAME="P0004" ID="P0004Chk" SIZE="3" VALUE="<%
			Response.Write GetOption(aOptionsComponent, REPORT_ROWS_OPTION)
		%>" CLASS="TextFields" />
		<FONT FACE="Arial" SIZE="2">
			[10-200]
		</FONT><BR /><BR />

		<INPUT TYPE="CHECKBOX" NAME="Dummy0003" ID="Dummy0003Chk"<%
			If CInt(GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)) = 1 Then Response.Write " CHECKED=""1"""
		%> onClick="SetHiddenValueForCheckBox(this.checked, this.form.P0003)" />
		<INPUT TYPE="HIDDEN" NAME="P0003" ID="P0003Hdn" VALUE="<%Response.Write GetOption(aOptionsComponent, EXPORT_FILTER_OPTION)%>" />
		<FONT FACE="Arial" SIZE="2">
			Exportar e imprimir la información del filtro junto con el reporte.
		</FONT><BR /><BR />

		<B>Tablero de control</B><BR />
		<IMG SRC="Images/Transparent.gif" WIDTH="20" HEIGHT="24" ALIGN="LEFT" />
		<FONT FACE="Arial" SIZE="2">Estilo de los semáforos: </FONT>
		<SELECT NAME="P0009" SIZE="1" CLASS="Lists">
			<OPTION VALUE="1">Iconos</OPTION>
			<OPTION VALUE="2">Gráfica de barras</OPTION>
		</SELECT>
		<SCRIPT LANGUAGE="JavaScript"><!--
			SelectItemByValue('<%Response.Write GetOption(aOptionsComponent, TRESHOLD_STYLE_OPTION)%>', false, document.OptionsFrm.P0009);
		//--></SCRIPT>
		<BR /><BR />
		<%
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""P0005"" ID=""P0005Hdn"" VALUE=""" & GetOption(aOptionsComponent, EMPLOYEE_ORDER_OPTION) & """ />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""P0006"" ID=""P0006Hdn"" VALUE=""" & GetOption(aOptionsComponent, EMPLOYEE_SORT_OPTION) & """ />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""P0007"" ID=""P0007Hdn"" VALUE=""" & GetOption(aOptionsComponent, PAYMENT_ORDER_OPTION) & """ />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""P0008"" ID=""P0008Hdn"" VALUE=""" & GetOption(aOptionsComponent, PAYMENT_SORT_OPTION) & """ />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""P0010"" ID=""P0010Hdn"" VALUE=""" & GetOption(aOptionsComponent, FULL_PROJECT_OPTION) & """ />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""P0011"" ID=""P0011Hdn"" VALUE=""" & GetOption(aOptionsComponent, CHECKS_LEFT_MARGIN1_OPTION) & """ />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""P0012"" ID=""P0012Hdn"" VALUE=""" & GetOption(aOptionsComponent, CHECKS_TOP_MARGIN1_OPTION) & """ />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""P0013"" ID=""P0013Hdn"" VALUE=""" & GetOption(aOptionsComponent, CHECKS_LEFT_MARGIN2_OPTION) & """ />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""P0014"" ID=""P0014Hdn"" VALUE=""" & GetOption(aOptionsComponent, CHECKS_TOP_MARGIN2_OPTION) & """ />" & vbNewLine
		%>