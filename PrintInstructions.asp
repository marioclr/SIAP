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
		<TITLE>Sistema Integral de Administración de Personal del ISSSTE</TITLE>
		<!-- #include file="_JavaScript.asp" -->
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" onLoad="window.focus()">
		<FONT FACE="Verdana" SIZE="2">
			<FONT COLOR="#<%Response.Write S_WARNING_FOR_GUI%>"><B>Instrucciones</B></FONT><BR /><BR />
			<B>Para eliminar los encabezados en la impresión:</B>
			<FONT SIZE="1"><OL>
				<LI>Cancele la impresión.</LI>
				<LI>En el menú <B>Archivo > Configuración de página</B> (File > Page Setup) borre el texto que se encuentra en los campos de texto <B>Encabezado</B> (Header) y <B>Pie de página</B> (Footer).</LI>
				<LI>Imprima el documento en <B>Archivo > Imprimir</B> (File > Print).</LI>
			</OL></FONT>
			<B>Para revisar que la información salga completa en la impresión:</B>
			<FONT SIZE="1"><OL>
				<LI>Cancele la impresión.</LI>
				<LI>En el menú <B>Archivo > Vista Preeliminar</B> (File > Print Preview) revise que el reporte salga completo.</LI>
				<LI>Si las columnas no se muestran completas dentro de la hoja, en el menú <B>Archivo > Configuración de página</B> (File > Page Setup) cambie la orientación de la impresión a horizontal.</LI>
				<LI>Imprima el documento en <B>Archivo > Imprimir</B> (File > Print).</LI>
			</OL></FONT>
			<B>Para imprimir los colores que se muestran en el documento:</B>
			<FONT SIZE="1"><OL>
				<LI>Cancele la impresión.</LI>
				<LI>En el menú <B>Herramientas > Opciones de Internet</B> (Tools > Internet Options) seleccione la pestaña <B>Avanzado</B> (Advanced).</LI>
				<LI>Marque la opción <B>Imprimir colores e imágenes de fondo</B> (Print background colors and images) ubicado en la sección Impresión (Printing) de la lista de opciones.</LI>
				<LI>Imprima el documento en <B>Archivo > Imprimir</B> (File > Print).</LI>
			</OL></FONT>
		</FONT>
	</BODY>
</HTML>