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
<!-- #include file="Libraries/EmployeesLib.asp" -->
<!-- #include file="Libraries/EmployeeComponent.asp" -->
<%
aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
aHeaderComponent(S_TITLE_NAME_HEADER) = "Personal"
bWaitMessage = False
Response.Cookies("SoS_SectionID") = 1001
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/Export.js"></SCRIPT>
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<!-- #include file="_Header.asp" -->
		Usted se encuentra aquí: <A HREF="Main.asp">Inicio</A> > <B>Personal</B><BR /><BR />
		<%Response.Write "<BR /><TABLE WIDTH=""720"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			aMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Centros de trabajo",_
					  "Administre la estructura de áreas generadoras y centros de trabajo y defina las plazas para cada uno de ellos.",_
					  "Images/MnAreas.gif", "Areas.asp", N_AREAS_PERMISSIONS),_
				Array("Puestos",_
					  "Defina los puestos y sus conceptos de pago.",_
					  "Images/MnPositions.gif", "Positions.asp", N_POSITIONS_PERMISSIONS),_
				Array("Plazas",_
					  "Busque las plazas que desea administrar.",_
					  "Images/MnJobs.gif", "Jobs.asp", N_JOBS_PERMISSIONS),_
				Array("Empleados",_
					  "Administre la información personal de los empleados, defina sus conceptos de pago y revise su historial de pagos.",_
					  "Images/MnEmployees.gif", "Employees.asp", N_EMPLOYEES_PERMISSIONS),_
				Array("Desarrollo Humano",_
					  "Bolsa de trabajo, calendario de cursos, empleados de nuevo ingreso",_
					  "Images/MnSADE.gif", "SADE.asp", N_SADE_PERMISSIONS)_
			)
			aMenuComponent(B_USE_DIV_MENU) = True
			Call DisplayMenuInTwoColumns(aMenuComponent)
		Response.Write "</TABLE>"%>
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>