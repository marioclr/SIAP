<%
Dim iSIAPTACOConnectionType
Dim oSIAPTACOADODBConnection
iSIAPTACOConnectionType = iConnectionType
Call CreateADODBConnection(TACO_DATABASE_PATH, TACO_DATABASE_USERNAME, TACO_DATABASE_PASSWORD, iSIAPTACOConnectionType, oSIAPTACOADODBConnection, "")
%>