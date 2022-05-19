<%
Dim objZip
Set objZip = Server.CreateObject("XStandard.Zip")
objZip.Pack "d:\Temp\uno.asp", "d:\Temp\uno.zip"
objZip.Pack "d:\Temp\dos.asp", "d:\Temp\dos.zip"
Set objZip = Nothing
%>