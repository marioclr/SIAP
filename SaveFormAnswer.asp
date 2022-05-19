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
<!-- #include file="Libraries/FormComponent.asp" -->
<%
If Len(oRequest) > 0 Then
	If Len(oRequest("TextAnswer").Item) > 0 Then
		Response.Write "<FORM NAME=""SaveFrm"" ID=""SaveFrm"" ACTION=""SaveFormAnswer.asp"" METHOD=""POST"">"
			Call DisplayURLParametersAsHiddenValues(RemoveParameterFromURLString(oRequest, "TextAnswer"))
			Response.Write "<TEXTAREA NAME=""Answer"" ID=""AnswerTxtArea"" ROWS=""5"" COLS=""60""></TEXTAREA>"
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "document.SaveFrm.Answer.value = parent.document." & oRequest("FormName").Item & ".FF__" & oRequest("FormID").Item & "__" & oRequest("FormFieldID").Item & ".value;" & vbNewLine
			Response.Write "document.SaveFrm.submit();" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	Else
		lErrorNumber = SaveUserAnswer(oADODBConnection, oRequest("FormID").Item, oRequest("FormFieldID").Item, oRequest("AnswerID").Item, oRequest("Answer").Item, sErrorDescription)
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			If lErrorNumber <> 0 Then
				Response.Write "alert('Error al guardar los cambios\n\n" & CleanStringForJavaScript(sErrorDescription) & "');" & vbNewLine
			Else
				If CLng(oRequest("bOnlyOne").Item) = 0 Then
					Response.Write "parent.SaveAnswers(parent.document." & oRequest("FormName").Item & ", " & (oRequest("ItemID").Item + 1) & ", false, '');" & vbNewLine
				End If
			End If
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If
End If
%>