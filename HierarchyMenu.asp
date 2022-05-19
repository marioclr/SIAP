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
Dim sAction
Dim sCondition
sAction = oRequest("Action").Item
sCondition = ""
If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
	sCondition = sCondition & " And (Areas.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))"
End If
%>
<HTML>
	<HEAD>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/HTMLLists.js"></SCRIPT>
		<LINK REL="STYLESHEET" TYPE="text/css" HREF="Styles/SIAP.css" />
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<FORM NAME="HierarchyMenuFrm" ID="HierarchyMenuFrm"><FONT FACE="Arial" SIZE="2">
			<%Select Case sAction
				Case "SubAreas"
					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
					Response.Write "Centros de trabajo:<BR />"
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""SubAreaID"" ID=""SubAreaIDLst"""
						If Len(oRequest("Size").Item) > 0 Then
							Response.Write " SIZE=""1"""
						Else
							Response.Write " SIZE=""5"" MULTIPLE=""1"""
						End If
					Response.Write " CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;} parent.window.document." & oRequest("TargetField").Item & ".value = GetSelectedValues(this).replace(/;;;/gi, ',');"">"
						If Len(oRequest("AreaID").Item) > 0 Then
							Response.Write "<OPTION VALUE=""" & oRequest("AreaID").Item & """>Todos</OPTION>"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(ParentID=" & oRequest("AreaID").Item & ") And (AreaID>-1)" & sCondition, "AreaShortName, AreaName", oRequest("SubAreaID").Item, "", sErrorDescription)
						Else
							Response.Write "<OPTION VALUE=""-1"">Todos</OPTION>"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(ParentID=-2) And (AreaID>-1)" & sCondition, "AreaShortName, AreaName", oRequest("SubAreaID").Item, "", sErrorDescription)
						End If
					Response.Write "</SELECT>"
					If Len(oRequest("AreaID").Item) = 0 Then
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							Response.Write "parent.window.document." & oRequest("TargetField").Item & ".value='-1';" & vbNewLine
						Response.Write "//--></SCRIPT>" & vbNewLine
					End If
				Case "Zones"
					Response.Write "<SELECT NAME=""ZoneID"" ID=""ZoneIDLst"" SIZE=""1"" CLASS=""Lists"" onChange="""
						If CInt(oRequest("PathLevel").Item) = 2 Then
							Response.Write "parent.window.document." & oRequest("TargetField").Item & ".location.href='HierarchyMenu.asp?Action=Zones&TargetField=" & oRequest("SecondTargetField").Item & "&ParentID=' + this.value + '&PathLevel=3';"
						Else
							Response.Write "parent.window.document." & oRequest("TargetField").Item & ".value=this.value;"
						End If
					Response.Write """>"
						If Len(oRequest("ParentID").Item) > 0 Then
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Zones", "ZoneID", "ZoneCode, ZoneName", "(ParentID=" & oRequest("ParentID").Item & ") And (ZoneID>-1)", "ZoneCode, ZoneName", oRequest("ZoneID").Item, "", sErrorDescription)
						Else
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Zones", "ZoneID", "ZoneCode, ZoneName", "(ParentID=-2) And (ZoneID>-1)", "ZoneCode, ZoneName", oRequest("ZoneID").Item, "", sErrorDescription)
						End If
					Response.Write "</SELECT>"
					Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
						If CInt(oRequest("PathLevel").Item) = 2 Then
							Response.Write "parent.window.document." & oRequest("TargetField").Item & ".location.href='HierarchyMenu.asp?Action=Zones&TargetField=" & oRequest("SecondTargetField").Item & "&ParentID=' + document.HierarchyMenuFrm.ZoneID.value + '&PathLevel=3';" & vbNewLine
						Else
							Response.Write "parent.window.document." & oRequest("TargetField").Item & ".value=document.HierarchyMenuFrm.ZoneID.value;"
						End If
					Response.Write "//--></SCRIPT>" & vbNewLine
				Case Else
			End Select%>
		</FONT></FORM>
	</BODY>
</HTML>
<SCRIPT LANGUAGE="JavaScript"><!--
	//HidePopupItem('WaitSmallDiv', document.WaitSmallDiv)
//--></SCRIPT>