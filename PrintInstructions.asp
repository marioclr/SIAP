<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
%>
<!-- #include file="Libraries/GlobalVariables.asp" -->
<HTML>
	<HEAD>
		<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1" />
		<TITLE>Sistema Integral de Administraci�n de Personal del ISSSTE</TITLE>
		<!-- #include file="_JavaScript.asp" -->
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" onLoad="window.focus()">
		<FONT FACE="Verdana" SIZE="2">
			<FONT COLOR="#<%Response.Write S_WARNING_FOR_GUI%>"><B>Instrucciones</B></FONT><BR /><BR />
			<B>Para eliminar los encabezados en la impresi�n:</B>
			<FONT SIZE="1"><OL>
				<LI>Cancele la impresi�n.</LI>
				<LI>En el men� <B>Archivo > Configuraci�n de p�gina</B> (File > Page Setup) borre el texto que se encuentra en los campos de texto <B>Encabezado</B> (Header) y <B>Pie de p�gina</B> (Footer).</LI>
				<LI>Imprima el documento en <B>Archivo > Imprimir</B> (File > Print).</LI>
			</OL></FONT>
			<B>Para revisar que la informaci�n salga completa en la impresi�n:</B>
			<FONT SIZE="1"><OL>
				<LI>Cancele la impresi�n.</LI>
				<LI>En el men� <B>Archivo > Vista Preeliminar</B> (File > Print Preview) revise que el reporte salga completo.</LI>
				<LI>Si las columnas no se muestran completas dentro de la hoja, en el men� <B>Archivo > Configuraci�n de p�gina</B> (File > Page Setup) cambie la orientaci�n de la impresi�n a horizontal.</LI>
				<LI>Imprima el documento en <B>Archivo > Imprimir</B> (File > Print).</LI>
			</OL></FONT>
			<B>Para imprimir los colores que se muestran en el documento:</B>
			<FONT SIZE="1"><OL>
				<LI>Cancele la impresi�n.</LI>
				<LI>En el men� <B>Herramientas > Opciones de Internet</B> (Tools > Internet Options) seleccione la pesta�a <B>Avanzado</B> (Advanced).</LI>
				<LI>Marque la opci�n <B>Imprimir colores e im�genes de fondo</B> (Print background colors and images) ubicado en la secci�n Impresi�n (Printing) de la lista de opciones.</LI>
				<LI>Imprima el documento en <B>Archivo > Imprimir</B> (File > Print).</LI>
			</OL></FONT>
		</FONT>
	</BODY>
</HTML>