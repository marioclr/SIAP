<%
Const ALL_TOOLBAR = -1
Const NO_TOOLBAR = 0
Const HOME_TOOLBAR = 1
Const BUDGET_TOOLBAR = 2
Const HUMAN_RESOURCES_TOOLBAR = 4
Const PAYROLL_TOOLBAR = 8
Const PAYMENTS_TOOLBAR = 16
Const REPORTS_TOOLBAR = 32
Const CATALOGS_TOOLBAR = 64
Const TOOLS_TOOLBAR = 128
Const DOCS_TOOLBAR = 256
Const LOGOUT_TOOLBAR = 512
Const DUMMY_TOOLBAR = 4096

Const L_SELECTED_OPTION_HEADER = 0
Const L_LINKED_OPTION_HEADER = 1
Const S_TITLE_NAME_HEADER = 2
Const S_WINDOW_TITLE_HEADER = 3

Const N_HEADER_COMPONENT_SIZE = 3

Dim aHeaderComponent()
Call ReceiveHeaderRequest(aHeaderComponent)

Function ReceiveHeaderRequest(aHeaderComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Header Component
'         using default values
'Outputs: aHeaderComponent
'************************************************************
	On Error Resume Next
	Redim Preserve aHeaderComponent(N_HEADER_COMPONENT_SIZE)

	aHeaderComponent(L_SELECTED_OPTION_HEADER) = DUMMY_TOOLBAR
	aHeaderComponent(L_LINKED_OPTION_HEADER) = ALL_TOOLBAR
	aHeaderComponent(S_TITLE_NAME_HEADER) = ""
	aHeaderComponent(S_WINDOW_TITLE_HEADER) = "Sistema Integral de Administración de Personal del ISSSTE"

	ReceiveHeaderRequest = Err.number
	Err.Clear
End Function
%>