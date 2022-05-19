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
If aLoginComponent(N_PROFILE_ID_LOGIN) > 0 Then
	Response.Redirect "Main_ISSSTE.asp?SectionID=" & aLoginComponent(N_PROFILE_ID_LOGIN)
End If

aHeaderComponent(L_SELECTED_OPTION_HEADER) = HOME_TOOLBAR
Select Case Hour(Time())
	Case 5,6,7,8,9,10,11
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Buenos d�as "
	Case 12,13,14,15,16,17
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Buenas tardes "
	Case 18
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Buenas tardes "
	Case 0,1,2,3,4,19,20,21,22,23
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Buenas noches "
	Case Else
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Bienvenido "
End Select
aHeaderComponent(S_TITLE_NAME_HEADER) = aHeaderComponent(S_TITLE_NAME_HEADER) & CleanStringForHTML(aLoginComponent(S_USER_NAME_LOGIN))
bWaitMessage = False
Response.Cookies("SoS_SectionID") = 187
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<!-- #include file="_Header.asp" -->
		Usted se encuentra aqu�: <B>Inicio</B><BR />
		<BR /><BR /><TABLE WIDTH="720" BORDER="0" CELLPADDING="0" CELLSPACING="0">
			<%If B_ISSSTE Then
				aMenuComponent(A_ELEMENTS_MENU) = Array(_
					Array("Personal",_
						  "Administre las �reas y sus plazas, los puestos y sus conceptos de pago, y a los empleados que conforman la plantilla del personal.",_
						  "Images/MnHumanResources.gif", "Main_ISSSTE.asp?SectionID=1", True),_
					Array("Prestaciones",_
						  "Administre los terceros institucionales, las antig�edades, la pensi�n alimenticia, el fondo de ahorro capitalizable, el sistema de ahorro para el retiro y el seguro de separaci�n individualizado.",_
						  "Images/MnSection2.gif", "Main_ISSSTE.asp?SectionID=2", True),_
					Array("Desarrollo Humano",_
						  "Revise las estructuras ocupacionales, los tabuladores, ejecute reportes, realice la selecci�n de personal, administre la capacitaci�n y la planeaci�n de recursos humanos.",_
						  "Images/MnSection3.gif", "Main_ISSSTE.asp?SectionID=3", True),_
					Array("Inform�tica",_
						  "Administre los conceptos de pago, los empleados, calcule la n�mina, genere y cancele cheques y ejecute reportes.",_
						  "Images/MnSection4.gif", "Main_ISSSTE.asp?SectionID=4", True),_
					Array("Presupuesto",_
						  "Administre las estructuras program�ticas y el clasificador por objeto del gasto, y ejecute reportes relacionados al presupuesto.",_
						  "Images/MnBudget.gif", "Main_ISSSTE.asp?SectionID=5", True),_
					Array("Departamento t�cnico",_
						  "Atienda la ventanilla �nica, revise los tr�mites que tiene pendientes, emita licencias por comisi�n sindical y controle los procesos del tablero de control",_
						  "Images/MnSection6.gif", "Main_ISSSTE.asp?SectionID=6", True),_
					Array("Desconcentrados",_
						  "Administre a los empleados, sus n�minas y ejecute reportes.",_
						  "Images/MnSection7.gif", "Main_ISSSTE.asp?SectionID=7", True),_
					Array("Herramientas",_
						  "Administraci�n de cat�logos, bit�cora de errores, cambio de su contrase�a, preferencias, etc.",_
						  "Images/MnTools.gif", "Tools.asp", True),_
					Array("Normateca",_
						  "Manuales de procedimientos, glosarios, anexos, etc.",_
						  "Images/MnManuals.gif", "DocsLibrary.asp", True)_
				)
			Else
				aMenuComponent(A_ELEMENTS_MENU) = Array(_
					Array("Presupuesto",_
						  "Defina el presupuesto para cada �rea.",_
						  "Images/MnBudget.gif", "Budget.asp", N_BUDGET_PERMISSIONS),_
					Array("Personal",_
						  "Administre las �reas y sus plazas, los puestos y sus conceptos de pago, y a los empleados que conforman la plantilla del personal.",_
						  "Images/MnHumanResources.gif", "HumanResources.asp", (N_AREAS_PERMISSIONS + N_POSITIONS_PERMISSIONS + N_EMPLOYEES_PERMISSIONS)),_
					Array("N�mina",_
						  "Calcule el pago de la n�mina para la plantilla de personal.",_
						  "Images/MnPayroll.gif", "Payroll.asp", True),_
					Array("Cheques",_
						  "Alta y b�squeda de los cheques para el pago de los empleados.",_
						  "Images/MnPayments.gif", "Payments.asp", (N_PAYMENTS_PERMISSIONS)),_
					Array("Reportes",_
						  "Obtenga reportes estad�sticos en relaci�n a la informaci�n del personal y el c�lculo de la n�mina.",_
						  "Images/MnReports.gif", "Reports.asp", (N_REPORTS_PERMISSIONS)),_
					Array("Herramientas",_
						  "Administraci�n de cat�logos, bit�cora de errores, cambio de su contrase�a, preferencias, etc.",_
						  "Images/MnTools.gif", "Tools.asp", True),_
					Array("Normateca",_
						  "Manuales de procedimientos, glosarios, anexos, etc.",_
						  "Images/MnManuals.gif", "DocsLibrary.asp", True)_
				)
			End If
			aMenuComponent(B_USE_DIV_MENU) = True
			Call DisplayMenuInTwoColumns(aMenuComponent)%>
		</TABLE>
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>