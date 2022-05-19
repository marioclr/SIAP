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
aHeaderComponent(S_TITLE_NAME_HEADER) = "Manuales"
bWaitMessage = False
Response.Cookies("SoS_SectionID") = 207
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<!-- #include file="_Header.asp" -->
		Usted se encuentra aquí: <A HREF="Main.asp">Inicio</A> > <B>Manuales</B><BR />
		<BR /><BR /><TABLE WIDTH="720" BORDER="0" CELLPADDING="0" CELLSPACING="0">
			<%aMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("<TITLE />Personal",_
					  "",_
					  "", "", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 1))),_
				Array("0. Introducción",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Personal/0. Personal.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 1))),_
				Array("1. Plazas",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Personal/1. Personal. Plazas.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 1))),_
				Array("2. Asignación de No. de empleado",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Personal/2. Personal. Asignacion No Empleado.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 1))),_
				Array("3. Consulta de personal",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Personal/3. Personal. Consulta Personal.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 1))),_
				Array("4. Administración de personal",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Personal/4. Personal. Admon Personal.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 1))),_
				Array("5. Agüinaldos",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Personal/5. Personal. Aguinaldos.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 1))),_
				Array("6. Acumulados",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Personal/6. Personal. Acumulados.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 1))),_
				Array("7. SIAE",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Personal/7. Personal. SIAE.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 1))),_
				Array("8. Reclamos",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Personal/8. Personal. Reclamos.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 1))),_
				Array("9. Reportes",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Personal/9. Personal. Reportes.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 1))),_
				Array("10. Catálogos",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Personal/10. Personal. Catalogos.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 1))),_
				Array("11. Alta de Honorarios",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Personal/11. PERSONAL. ALTA DE HONORARIOS.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 1))),_
				Array("12. Baja de Personal de Honorarios",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Personal/12. PERSONAL. BAJA DE PERSONAL DE HONORARIOS.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 1))),_
				Array("13. Cambio de Honorarios (Cambio de Importe de Honorarios)",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Personal/13. PERSONAL. CAMBIO DE HONORARIOS (CAMBIO DE IMPORTE DE HONORARIOS).PDF"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 1))),_
				


				Array("<TITLE />Prestaciones",_
					  "",_
					  "", "", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 2))),_
				Array("0. Introducción",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Prestaciones/0. Prestaciones.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 2))),_
				Array("1. Consulta de personal",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Prestaciones/1. Prestaciones. Consulta Personal.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 2))),_
				Array("2. Prestaciones",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Prestaciones/2. Desconcentrados. Prestaciones.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 2))),_
				Array("3. Informática",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Prestaciones/3. Desconcentrados. Informatica.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 2))),_
				Array("4. Pension Alimenticia",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Prestaciones/4. Prestaciones. Pension Alimenticia.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 2))),_

				Array("<TITLE />Desarrollo Humano",_
					  "",_
					  "", "", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 3))),_
				Array("00. Indice",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Desarrollo Humano/00. Desarrollo Humano. Indice.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 3))),_
				Array("0. Introducción",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Desarrollo Humano/0. Desarrollo Humano. Introduccion.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 3))),_
				Array("1. Estructuras ocupacionales. Catálogos",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Desarrollo Humano/1. Desarrollo Humano. Estructuras Ocupacionales. Catalogos.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 3))),_
				Array("2. Estructuras ocupacionales. Carga de tabuladores",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Desarrollo Humano/2. Desarrollo Humano. Estructuras Ocupacionales. Carga Tabuladores.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 3))),_
				Array("3.1. Estructuras ocupacionales. Tabuladores",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Desarrollo Humano/3.1. Desarrollo Humano. Estructuras Ocupacionales. Tabuladores.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 3))),_
				Array("3. Estructuras ocupacionales. Carga UNIMED",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Desarrollo Humano/3. Desarrollo Humano. Estructuras Ocupacionales. Carga UNIMED.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 3))),_
				Array("4. Reportes",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Desarrollo Humano/4. Desarrollo Humano. Reportes.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 3))),_
				Array("5. Consulta de personal",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Desarrollo Humano/5. Personal. Consulta Personal.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 3))),_
				Array("6. Desarrollo humano. Capacitación",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Desarrollo Humano/6. Desarrollo Humano. Desarrollo Humano. Capacitacion.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 3))),_
				Array("7. Planeación de recursos humanos",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Desarrollo Humano/7. Desarrollo Humano. Planeacion Recursos Humanos.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 3))),_
				Array("8. Búsqueda de centros de trabajo y centros de pago",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Desarrollo Humano/8. Desarrollo Humano. Busqueda Centros Trabajo Centros Pago.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 3))),_
				Array("9. Reportes",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Desarrollo Humano/9. Personal. Reportes.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 3))),_
				Array("10. Catálogos",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Desarrollo Humano/10. Personal. Catalogos.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 3))),_
				Array("11. Agregar una Nueva Plaza",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Desarrollo Humano/11. DESARROLLO HUMANO. AGREGAR UNA NUEVA PLAZA.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 3))),_


				Array("<TITLE />Informática",_
					  "",_
					  "", "", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 4))),_
				Array("0. Introducción",_
					  "",_
                      "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Informatica/0. Informatica.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 4))),_
				Array("1. Conceptos Pago",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Informatica/1. Informatica. Conceptos Pago.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 4))),_
				Array("2. Empleados",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Informatica/2. Informatica. Empleados.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 4))),_
				Array("3. Nueva nómina",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Informatica/3. Informatica. Nueva Nomina.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 4))),_
				Array("4. Prenómina",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Informatica/4. Informatica. Prenomina.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 4))),_
				Array("5. Cerrar nómina",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Informatica/5. Informatica. Cerrar Nomina.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 4))),_
				Array("6. Nóminas especiales",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Informatica/6. Informatica. Nominas Especiales.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 4))),_
				Array("7. Cheques",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Informatica/7. Informatica. Cheques.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 4))),_
				Array("8. Apertura y cierre de registros",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Informatica/8. Informatica. Apertura Registros.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 4))),_
				Array("9. Reportes",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Informatica/9. Informatica. Reportes.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 4))),_
				Array("10. Catálogos",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Informatica/10. Informatica. Catalogos.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 4))),_
                Array("11. Guía rápida de impresión de cheques",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Informatica/11. Informatica. Guia rapida de impresion de cheques.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 4))),_
                 Array("12. Guía rápida de impresión de depósitos",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Informatica/12. Informatica. Guia rapida de impresion de depositos.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 4))),_
				Array("<TITLE />Desconcentrados",_
					  "",_
					  "", "", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 5))),_
				Array("1. Personal",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Desconcentrados/1. Desconcentrados. Personal.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 5))),_
				Array("2. Prestaciones",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Desconcentrados/2. Desconcentrados. Prestaciones.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 5))),_
				Array("3. Informática",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Desconcentrados/3. Desconcentrados. Informatica.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 5))),_
				Array("4. Incidencias",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Desconcentrados/4. Desconcentrados. Incidencias.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 5))),_
				Array("5. Informática. Nómina",_
					  "",_
					  "Images/MnDocument.gif", "Uploaded Files/DocsLibrary/Desconcentrados/5. Desconcentrados. Informatica. Nomina.pdf"" TARGET=""_blank", ((aLoginComponent(N_PROFILE_ID_LOGIN) = 0) Or (aLoginComponent(N_PROFILE_ID_LOGIN) = 5))),_

				Array("",_
					  "",_
					  "", "", False)_
			)
			aMenuComponent(B_USE_DIV_MENU) = True
			Call DisplayMenuInTwoColumns(aMenuComponent)%>
		</TABLE>
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>