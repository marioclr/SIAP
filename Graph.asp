<!-- #include file="Libraries/GlobalVariables.asp" -->
<!-- #include file="Libraries/GraphComponent.asp" -->
<HTML>
	<HEAD>
		<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1" />
	</HEAD>
	<BODY TEXT="#000000" LINK="#000099" ALINK="#0000FF" VLINK="#000099" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<CENTER>
			<%Call DisplayGraph(oRequest, False, aGraphComponent, sErrorDescription)%>
		</CENTER>
	</BODY>
</HTML>