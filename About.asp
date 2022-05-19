<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
%>
<!-- #include file="Libraries/GlobalVariables.asp" -->
<%
Dim aTemp
Dim lCourseID
Dim lEvaluationNumber
Dim sNames

aHeaderComponent(S_TITLE_NAME_HEADER) = "Acerca de SIAP"
aHeaderComponent(L_SELECTED_OPTION_HEADER) = NO_TOOLBAR
Response.Cookies("SoS_SectionID") = 187
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<!-- #include file="_Header.asp" -->
		<BR />
		<TABLE WIDTH="760" BORDER="0" CELLSPACING="0" CELLPADDING="0">
			<TR>
				<TD WIDTH="1"><IMG SRC="Images/Transparent.gif" WIDTH="1" HEIGHT="1" /></TD>
				<TD WIDTH="379" VALIGN="TOP"><FONT FACE="Arial" SIZE="2">
					<!-- BEGIN: CONTENTS -->
					<B>Desarrollado por xxx</B><BR />
					<BR />
					Este desarrollo comenzó el <%Response.Write DisplayDateFromSerialNumber(L_SIAP_DATE, -1, -1, -1)%> y se encuentra protegido por las leyes del registro de autor.<BR />
					<BR />
					<B>Versión: </B>1.0.0.<%Response.Write DateDiff("d", GetDateFromSerialNumber(L_SIAP_DATE), Now()) & "&nbsp;(" &  Now()%>)<BR /><BR />
						<DIV STYLE="width: 100%; height: 212px; overflow: auto;">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Reportes relacionados<BR />
						</DIV>
					<BR />
					<!-- END: CONTENTS -->
				</FONT></TD>
				<TD WIDTH="5"><IMG SRC="Images/Transparent.gif" WIDTH="5" HEIGHT="1" /></TD>
				<TD WIDTH="1" VALIGN="MIDDLE" ROWSPAN="2"><IMG SRC="Images/DotTeal.gif" WIDTH="1" HEIGHT="300" /></TD>
				<TD WIDTH="5"><IMG SRC="Images/Transparent.gif" WIDTH="5" HEIGHT="1" /></TD>
				<TD WIDTH="369" VALIGN="TOP"><!--
					<FONT FACE="Arial" SIZE="2"><B>Versiones anteriores:</B></FONT>
					<DIV STYLE="width: 100%; height: 305px; overflow: auto;">
						<!- - BEGIN: CONTENTS - - >
						<FONT FACE="Arial" SIZE="2">
							&nbsp;&nbsp;&nbsp;-&nbsp;<B>Versión 1.1.0.121</B><BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Componentes<BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Core engine<BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Diseño de la base de datos<BR />
							&nbsp;&nbsp;&nbsp;-&nbsp;<B>Versión 1.2.0.324</B><BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Conexión con la agenda<BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Nueva interfaz<BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;<FONT COLOR="#D20000">Wizards</FONT><BR />
							&nbsp;&nbsp;&nbsp;-&nbsp;<B>Versión 1.3.0.393</B><BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Usuarios conectados<BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Registro de entradas al sistema<BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Reportes graficados<BR />
							&nbsp;&nbsp;&nbsp;-&nbsp;<B>Versión 1.4.0.419</B><BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Nuevos reportes<BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Registro de entradas a las evaluaciones<BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Registro de entradas a los cursos<BR />
							&nbsp;&nbsp;&nbsp;-&nbsp;<B>Versión 2.0.0.784</B><BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Bitácora de errores<BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Carga rápida de registros ("Fast upload")<BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Continuación del curso después de una evaluación<BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Descripción para los cursos<BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Mejora en los Wizards<BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;<FONT COLOR="#D20000">Nueva interfaz</FONT><BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Perfiles<BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Selección de página de entrada<BR />
							&nbsp;&nbsp;&nbsp;-&nbsp;<B>Versión 2.1.0.999</B><BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Creación de cuentas por parte de los usuarios<BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;<FONT COLOR="#D20000">Encuestas</FONT><BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Mejora en los reportes<BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Notas para los tutores<BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Selección de estilo para las tablas<BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Solicitud de contraseña<BR />
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#183;&nbsp;Soporte de cursos vía http<BR />
						</FONT>
						<!- - END: CONTENTS - - >
					</DIV>
				--></TD>
			</TR>
		</TABLE>
		<FORM><INPUT TYPE="BUTTON" VALUE="Regresar" CLASS="Buttons" onClick="window.history.go(-1)" /></FORM>
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>