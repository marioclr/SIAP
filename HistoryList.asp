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
Dim lPaperworkID
Dim lOwnerID
Dim lRecordID
Dim bClosed
Dim bOwner
Dim oRecordset

sAction = oRequest("Action").Item
lPaperworkID = CLng(oRequest("PaperworkID").Item)
lOwnerID = 0
If IsNumeric(oRequest("OwnerID").Item) Then lOwnerID = CLng(oRequest("OwnerID").Item)
Select Case sAction
	Case "Paperworks"
		bClosed = False
		bOwner = False
		sErrorDescription = "No se pudo obtener la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EndDate From Paperworks Where (PaperworkID=" & lPaperworkID & ")", "HistoryList.asp", "_root", 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				bClosed = (CLng(oRecordset.Fields("EndDate").Value) > 0)
                'bClosed = False
			End If
			oRecordset.Close
		End If
		sErrorDescription = "No se pudo obtener la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaperworkOwnersLKP.OwnerID From PaperworkOwnersLKP, PaperworkOwners Where (PaperworkOwnersLKP.OwnerID=PaperworkOwners.OwnerID) And (PaperworkOwnersLKP.PaperworkID=" & lPaperworkID & ") And (PaperworkOwners.EmployeeID=" & lOwnerID & ")", "HistoryList.asp", "_root", 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			bOwner = (Not oRecordset.EOF)
			oRecordset.Close
		End If
		If Len(oRequest("AddComment").Item) > 0 Then
			sErrorDescription = "No se pudo obtener un identificador para el nuevo registro."
			lErrorNumber = GetNewIDFromTable(oADODBConnection, "PaperworkComments", "RecordID", "(PaperworkID=" & lPaperworkID & ")", 1, lRecordID, sErrorDescription)
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudo obtener el historial de comentarios."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PaperworkComments (PaperworkID, RecordID, OwnerID, CommentDate, CommentHour, Comments) Values (" & lPaperworkID & ", " & lRecordID & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & Mid(GetSerialNumberForDate(""), Len("000000000"), Len("0000")) & ", '" & Replace(oRequest("Comments").Item, "'", "´") & "')", "HistoryList.asp", "_root", 000, sErrorDescription, oRecordset)
			End If
		ElseIf Len(oRequest("CloseOwner").Item) > 0 Then
			sErrorDescription = "No se pudo obtener el historial de comentarios."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update PaperworkOwnersLKP Set EndDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", ClosingNumber='" & Replace(oRequest("ClosingNumber").Item, "'", "´") & "' Where (PaperworkID=" & lPaperworkID & ") And (OwnerID=" & lOwnerID & ")", "HistoryList.asp", "_root", 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudo obtener el historial de comentarios."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select OwnerID From PaperworkOwnersLKP Where (PaperworkID=" & lPaperworkID & ") And (EndDate=0)", "HistoryList.asp", "_root", 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					sErrorDescription = "No se pudo obtener el historial de comentarios."
					If oRecordset.EOF Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Paperworks Set EndDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", DocClassification='" & Replace(oRequest("ClosingNumber").Item, "'", "´") & "', StatusID=3 Where (PaperworkID=" & lPaperworkID & ")", "HistoryList.asp", "_root", 000, sErrorDescription, oRecordset)
					Else
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Paperworks Set EndDate=0, DocClassification='', StatusID=0 Where (PaperworkID=" & lPaperworkID & ")", "HistoryList.asp", "_root", 000, sErrorDescription, oRecordset)
					End If
					oRecordset.Close
				End If
			End If
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "parent.window.location.href='EmployeeSupport.asp?PaperworkID=" & lPaperworkID & "&Change=1';" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
		End If
End Select
%>
<HTML>
	<HEAD>
		<LINK REL="STYLESHEET" TYPE="text/css" HREF="Styles/SIAP.css" />
		<SCRIPT LANGUAGE="JavaScript"><!--
			function CheckHistoryListFields(oForm) {
				if (oForm) {
					if (oForm.Comments.value == '') {
						alert('Favor de introducir un comentario');
						oForm.Comments.focus();
						return false;
					}
				}

				return true;
			} // End of CheckHistoryListFields

			function CheckCloseOwnerFields(oForm) {
				if (oForm) {
					if (oForm.ClosingNumber.value == '') {
						alert('Favor de introducir el número de oficio');
						oForm.ClosingNumber.focus();
						return false;
					}
				}
				return true;
			} // End of CheckCloseOwnerFields
		//--></SCRIPT>
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%Select Case sAction
			Case "Paperworks"
				Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
					If Not bClosed Then
						Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"
							Response.Write "<FORM NAME=""HistoryListFrm"" ID=""HistoryListFrm"" ACTION=""HistoryList.asp"" METHOD=""POST"" onSubmit=""return CheckHistoryListFields(this)"">"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""Paperworks"" />"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PaperworkID"" ID=""PaperworkIDHdn"" VALUE=""" & lPaperworkID & """ />"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OwnerID"" ID=""OwnerIDHdn"" VALUE=""" & lOwnerID & """ />"
								Response.Write "Introducir Comentario:<BR />"
								Response.Write "<TEXTAREA NAME=""Comments"" ID=""CommentsTxtArea"" ROWS=""6"" COLS=""40"" MAXLENGTH=""2000"" CLASS=""TextFields""></TEXTAREA><BR /><BR />"
								Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""AddComment"" ID=""AddCommentBtn"" VALUE=""Agregar Comentario"" CLASS=""Buttons"" />"
							Response.Write "</FORM><BR />"
							If bOwner And False Then
								Response.Write "<IMG SRC=""Images/DotBlue.gif"" WIDTH=""320"" HEIGHT=""1"" /><BR /><BR />"
								Response.Write "<FORM NAME=""CloseOwnerFrm"" ID=""CloseOwnerFrm"" ACTION=""HistoryList.asp"" METHOD=""GET"" onSubmit=""return CheckCloseOwnerFields(this)"">"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""Paperworks"" />"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PaperworkID"" ID=""PaperworkIDHdn"" VALUE=""" & lPaperworkID & """ />"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OwnerID"" ID=""OwnerIDHdn"" VALUE=""" & lOwnerID & """ />"
									Response.Write "Oficio:<BR />"
									Response.Write "<INPUT NAME=""ClosingNumber"" ID=""ClosingNumberTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE="""" CLASS=""TextFields"" /><BR /><BR />"
									Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""CloseOwner"" ID=""CloseOwnerBtn"" VALUE=""Guardar descargo"" CLASS=""Buttons"" />"
								Response.Write "</FORM>"
							End If
						Response.Write "</FONT></TD>"
						Response.Write "<TD>&nbsp;&nbsp;&nbsp;</TD>"
					End If
					Response.Write "<TD VALIGN=""TOP"">Seguimiento/Comentarios<DIV CLASS=""HistoryList""><FONT FACE=""Arial"" SIZE=""2"">"
						sErrorDescription = "No se pudo obtener el historial de comentarios."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaperworkComments.*, UserName, UserLastName From PaperworkComments, Users Where (PaperworkComments.OwnerID=Users.UserID) And (PaperworkComments.PaperworkID=" & lPaperworkID & ") Order By CommentDate, CommentHour, RecordID", "HistoryList.asp", "_root", 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								Do While Not oRecordset.EOF
									Response.Write "<B>" & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("CommentDate").Value), Int(CInt(oRecordset.Fields("CommentHour").Value) / 100), (CInt(oRecordset.Fields("CommentHour").Value) Mod 100), -1) & ". " & CleanStringForHTML(CStr(oRecordset.Fields("UserName").Value) & " " & CStr(oRecordset.Fields("UserLastName").Value)) & ":</B><BR />"
									Response.Write CleanStringForHTML(CStr(oRecordset.Fields("Comments").Value)) & "<BR /><BR />"
									oRecordset.MoveNext
									If Err.number <> 0 Then Exit Do
								Loop
							Else
								Response.Write "<B>No existen comentarios registrados.</B>"
							End If
						End If
					Response.Write "</FONT></DIV></TD>"
				Response.Write "</TR></TABLE><BR /><BR />"
		End Select%>
	</BODY>
</HTML>
<SCRIPT LANGUAGE="JavaScript"><!--
	//HidePopupItem('WaitSmallDiv', document.WaitSmallDiv)
//--></SCRIPT>
<%
Set oRecordset = Nothing
%>