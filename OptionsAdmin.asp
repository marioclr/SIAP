<%
Dim iIndexForAdminOptions
Dim jIndexForAdminOptions
%>
			<INPUT TYPE="HIDDEN" NAME="Admin" ID="AdminHdn" VALUE="1" />
			<%Call DisplayErrorMessage("", "<FONT COLOR=""#" & S_INSTRUCTIONS_FOR_GUI & """><B>Estos valores afectan a todo el sistema. Modifique con precauci�n estas<BR />opciones de acuerdo a las necesidades de administraci�n del sistema.</B></FONT>")%>
			<BR />

			<DIV CLASS="TitleBar"><FONT FACE="Arial" SIZE="2" COLOR="#FFFFFF"><B>&nbsp;OPCIONES GENERALES</B></FONT></DIV><BR />
			<!-- INFORMACI�N DEL CONTACTO -->
			<B>Informaci�n del contacto:</B><BR />
			<IMG SRC="Images/Transparent.gif" WIDTH="20" HEIGHT="72" ALIGN="LEFT" />
			Esta informaci�n le permitir� al usuario ponerse en contacto con los administradores del sistema en caso de surgir alguna duda o problema.<BR />
			Nombre: <INPUT TYPE="TEXT" NAME="P0004" ID="P0004Txt" SIZE="30" MAXLENGTH="30" VALUE="<%Response.Write GetAdminOption(aAdminOptionsComponent, CONTACT_NAME_OPTION)%>" CLASS="TextFields" /><BR />
			Tel�fono: <INPUT TYPE="TEXT" NAME="P0005" ID="P0005Txt" SIZE="30" MAXLENGTH="30" VALUE="<%Response.Write GetAdminOption(aAdminOptionsComponent, CONTACT_PHONE_OPTION)%>" CLASS="TextFields" /><BR />
			Correo electr�nico: <INPUT TYPE="TEXT" NAME="P0006" ID="P0006Txt" SIZE="30" MAXLENGTH="30" VALUE="<%Response.Write GetAdminOption(aAdminOptionsComponent, CONTACT_EMAIL_OPTION)%>" CLASS="TextFields" /><BR />
			<BR /><BR />

			<DIV CLASS="TitleBar"><FONT FACE="Arial" SIZE="2" COLOR="#FFFFFF"><B>&nbsp;BIT�CORA DE ERRORES</B></FONT></DIV><BR />
			<!-- REGISTRO DE OPERACIONES SOBRE LA BASE DE DATOS -->
			<B>Registro de operaciones sobre la base de datos:</B><BR />
			<IMG SRC="Images/Transparent.gif" WIDTH="20" HEIGHT="60" ALIGN="LEFT" />
			<INPUT TYPE="CHECKBOX" NAME="Dummy0007" ID="Dummy0007Chk"<%
				If CInt(GetAdminOption(aAdminOptionsComponent, UPDATE_OPTION)) = 1 Then Response.Write " CHECKED=""1"""
			%> onClick="SetHiddenValueForCheckBox(this.checked, this.form.P0007)" />
			<INPUT TYPE="HIDDEN" NAME="P0007" ID="P0007Hdn" VALUE="<%Response.Write GetAdminOption(aAdminOptionsComponent, UPDATE_OPTION)%>" />
			Registrar las operaciones de actualizaci�n de registros de la base de datos.<BR />

			<INPUT TYPE="CHECKBOX" NAME="Dummy0008" ID="Dummy0008Chk"<%
				If CInt(GetAdminOption(aAdminOptionsComponent, DELETE_OPTION)) = 1 Then Response.Write " CHECKED=""1"""
			%> onClick="SetHiddenValueForCheckBox(this.checked, this.form.P0008)" />
			<INPUT TYPE="HIDDEN" NAME="P0008" ID="P0008Hdn" VALUE="<%Response.Write GetAdminOption(aAdminOptionsComponent, DELETE_OPTION)%>" />
			Registrar las operaciones de eliminaci�n de registros de la base de datos.<BR />

			<INPUT TYPE="CHECKBOX" NAME="Dummy0009" ID="Dummy0009Chk"<%
				If CInt(GetAdminOption(aAdminOptionsComponent, INSERT_OPTION)) = 1 Then Response.Write " CHECKED=""1"""
			%> onClick="SetHiddenValueForCheckBox(this.checked, this.form.P0009)" />
			<INPUT TYPE="HIDDEN" NAME="P0009" ID="P0009Hdn" VALUE="<%Response.Write GetAdminOption(aAdminOptionsComponent, INSERT_OPTION)%>" />
			Registrar las operaciones de inserci�n de registros de la base de datos.<BR />
			<BR /><BR />

			<DIV CLASS="TitleBar"><FONT FACE="Arial" SIZE="2" COLOR="#FFFFFF"><B>&nbsp;FONAC</B></FONT></DIV><BR />
			<!-- VALORES PARA EL C�LCULO DEL FONAC -->
			<IMG SRC="Images/Transparent.gif" WIDTH="20" HEIGHT="76" ALIGN="LEFT" />
			Aportaci�n del empleado al FONAC: <INPUT TYPE="TEXT" NAME="P0016" ID="P0016Txt" SIZE="10" MAXLENGTH="10" VALUE="<%Response.Write GetAdminOption(aAdminOptionsComponent, FONAC_01_OPTION)%>" CLASS="TextFields" /><BR />
			Aportaci�n de la Instituci�n al FONAC: <INPUT TYPE="TEXT" NAME="P0017" ID="P0017Txt" SIZE="10" MAXLENGTH="10" VALUE="<%Response.Write GetAdminOption(aAdminOptionsComponent, FONAC_02_OPTION)%>" CLASS="TextFields" /><BR />
			Aportaci�n de la dependencia al FONAC: <INPUT TYPE="TEXT" NAME="P0018" ID="P0018Txt" SIZE="10" MAXLENGTH="10" VALUE="<%Response.Write GetAdminOption(aAdminOptionsComponent, FONAC_03_OPTION)%>" CLASS="TextFields" /><BR />
			Factor del sindicato para el FONAC: <INPUT TYPE="TEXT" NAME="P0019" ID="P0019Txt" SIZE="10" MAXLENGTH="10" VALUE="<%Response.Write GetAdminOption(aAdminOptionsComponent, FONAC_04_OPTION)%>" CLASS="TextFields" /><BR />
			<BR />

			<DIV CLASS="TitleBar"><FONT FACE="Arial" SIZE="2" COLOR="#FFFFFF"><B>&nbsp;SEGURIDAD</B></FONT></DIV><BR />
			<!-- CAMBIO DE CONTRASE�A -->
			<B>Cambio de contrase�a:</B><BR />
			<IMG SRC="Images/Transparent.gif" WIDTH="20" HEIGHT="22" ALIGN="LEFT" />
			Los usuarios deber�n cambiar su contrase�a cada <INPUT TYPE="TEXT" NAME="P0003" ID="P0003Txt" SIZE="3" MAXLENGTH="3" VALUE="<%Response.Write GetAdminOption(aAdminOptionsComponent, PASSWORDS_DAYS_OPTION)%>" CLASS="TextFields" /> d�as. [30 - 365]
			<BR /><BR />

			<!-- BLOQUEAR EL SISTEMA -->
			<B>Bloquear el Sistema:</B><BR />
			<IMG SRC="Images/Transparent.gif" WIDTH="20" HEIGHT="22" ALIGN="LEFT" />
			Bloquear el sistema si se detectan <INPUT TYPE="TEXT" NAME="P0001" ID="P0001Hdn" SIZE="3" MAXLENGTH="3" VALUE="<%Response.Write GetAdminOption(aAdminOptionsComponent, LOGIN_FAILURES_OPTION)%>" CLASS="TextFields" /> intentos fallidos para entrar al sistema desde la misma m�quina.  [3 - 100]<BR />
			Enviar un correo electr�nico a <INPUT TYPE="TEXT" NAME="P0002" ID="P0002Hdn" SIZE="40" MAXLENGTH="100" VALUE="<%Response.Write GetAdminOption(aAdminOptionsComponent, SYSTEM_BLOCKED_RECIPIENTS_OPTION)%>" CLASS="TextFields" /> al bloquear el sistema.
			<BR /><BR />

			<!-- TABLERO DE CONTROL -->
			<DIV CLASS="TitleBar"><FONT FACE="Arial" SIZE="2" COLOR="#FFFFFF"><B>&nbsp;TABLERO DE CONTROL</B></FONT></DIV><BR />
			<B>Colores del sem�foro:</B><BR />
			<IMG SRC="Images/Transparent.gif" WIDTH="20" HEIGHT="152" ALIGN="LEFT" />
			<SPAN STYLE="background-color: #<%Response.Write GetAdminOption(aAdminOptionsComponent, RED_COLOR_OPTION)%>"><IMG SRC="Images/IcnTreshold.gif" WIDTH="16" HEIGHT="16"></SPAN>
			Sem�foro rojo: <INPUT TYPE="TEXT" NAME="P0010" ID="P0010Txt" SIZE="6" MAXLENGTH="6" VALUE="<%Response.Write GetAdminOption(aAdminOptionsComponent, RED_COLOR_OPTION)%>" CLASS="TextFields" /><BR />
			Este color aparecer� cuando el porcentaje de avance sea: <%Response.Write GetAdminOption(aAdminOptionsComponent, RED_TRESHOLD_OPTION)%>%<BR /><BR />

			<SPAN STYLE="background-color: #<%Response.Write GetAdminOption(aAdminOptionsComponent, YELLOW_COLOR_OPTION)%>"><IMG SRC="Images/IcnTreshold.gif" WIDTH="16" HEIGHT="16"></SPAN>
			Sem�foro amarillo: <INPUT TYPE="TEXT" NAME="P0011" ID="P0011Txt" SIZE="6" MAXLENGTH="6" VALUE="<%Response.Write GetAdminOption(aAdminOptionsComponent, YELLOW_COLOR_OPTION)%>" CLASS="TextFields" /><BR />
			Este color aparecer� cuando el porcentaje de avance sea: <INPUT TYPE="TEXT" NAME="P0014" ID="P0014Txt" SIZE="2" MAXLENGTH="2" VALUE="<%Response.Write GetAdminOption(aAdminOptionsComponent, YELLOW_TRESHOLD_OPTION)%>" CLASS="TextFields" />%<BR /><BR />

			<SPAN STYLE="background-color: #<%Response.Write GetAdminOption(aAdminOptionsComponent, GREEN_COLOR_OPTION)%>"><IMG SRC="Images/IcnTreshold.gif" WIDTH="16" HEIGHT="16"></SPAN>
			Sem�foro verde: <INPUT TYPE="TEXT" NAME="P0012" ID="P0012Txt" SIZE="6" MAXLENGTH="6" VALUE="<%Response.Write GetAdminOption(aAdminOptionsComponent, GREEN_COLOR_OPTION)%>" CLASS="TextFields" /><BR />
			Este color aparecer� cuando el porcentaje de avance sea: <%Response.Write GetAdminOption(aAdminOptionsComponent, GREEN_TRESHOLD_OPTION)%>%
			<BR /><BR />