<%
Function AddNodeToXML(sNodePath, sCondition, vNewNode, oXML, sErrorDescription)
'************************************************************
'Purpose: Append a new node into the XML DOM object, at the
'         end of the given path
'Inputs:  sNodePath, sCondition, vNewNode
'Outputs: oXML, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddNodeToXML"
	Dim oTempXMLForNewNode
	Dim lErrorNumber

	If Len(sCondition) > 0 Then
		If InStr(1, sCondition, "[", vbTextCompare) <> 1 Then
			sCondition = "[" & sCondition & "]"
		End If
	End If

	If Not IsObject(vNewNode) Then
		lErrorNumber = CreateXMLDOMObject(oTempXMLForNewNode, sErrorDescription)
		If lErrorNumber = 0 Then
			If StrComp(Right(vNewNode, Len(".xml")), ".xml", vbTextCompare) = 0 Then
				oTempXMLForNewNode.load(Server.MapPath(vNewNode))
				If Err.number <> 0 Then
					Err.Clear 
					oTempXMLForNewNode.load(vNewNode)
				End If
			Else
				oTempXMLForNewNode.loadXML(CStr(vNewNode))
			End If
			lErrorNumber = Err.number
			If lErrorNumber <> 0 Then
				sErrorDescription = "La estructura XML no pudo ser cargada en el objeto XML DOM. Puede que la estructura contenga errores de sintaxis. Favor de contactar al Administrador."
				If Len(Err.description) > 0 Then
					sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
				End If
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "XMLLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
			End If
		End If
	Else
		Set oTempXMLForNewNode = vNewNode
	End If

	If lErrorNumber = 0 Then
		oXML.selectSingleNode(sNodePath & sCondition).appendChild(oTempXMLForNewNode)
		lErrorNumber = Err.number
		If lErrorNumber <> 0 Then
			sErrorDescription = "El nodo no pudo ser agregado en la estructura XML. Puede que la posición indicada no exista o que la condición esté mal."
			If Len(Err.description) > 0 Then
				sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
			End If
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "XMLLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
		End If
	End If

	Set oNewElement = Nothing
	Set oNewAttribute = Nothing
	AddNodeToXML = lErrorNumber
	Err.Clear
End Function

Function CleanStringForXML(sStringToChange, bCleanQuotes)
'************************************************************
'Purpose: To replace <BR />, á, é, í, ó, ú, ñ, Á, É, Í, Ó, Ú, Ñ, ¿, ¡, ü, " in a string
'Inputs:  sStringToChange, bCleanQuotes
'Outputs: A string that can be used as XML text without breaking any tag
'************************************************************
	CleanStringForXML = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(sStringToChange, "<BR />", "&#60;BR />"), "á", "&#225;"), "é", "&#233;"), "í", "&#237;"), "ó", "&#243;"), "ú", "&#250;"), "ñ", "&#241;"), "Á", "&#193;"), "É", "&#201;"), "Í", "&#205;"), "Ó", "&#211;"), "Ú", "&#218;"), "Ñ", "&#209;"), "¿", "&#191;"), "¡", "&#161;"), "ü", "&#252;")
	If bCleanQuotes Then
		CleanStringForXML = Replace(CleanStringForXML, """", "&#34;")
	End If
End Function

Function CreateXMLDOMObject(oXMLDOM, sErrorDescription)
'************************************************************
'Purpose: To create an XML DOM Object
'Outputs: oXMLDOM, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CreateXMLDOMObject"
	Dim lErrorNumber

	Set oXMLDOM = Server.CreateObject("Microsoft.XMLDOM")
	lErrorNumber = Err.Number
	If lErrorNumber <> 0 Then
		sErrorDescription = "El objeto XML DOM no pudo ser creado. El archivo MSXML.dll puede no estar registrado en el servidor Web. Favor de contactar al Administrador."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "XMLLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	ElseIf oXMLDOM.parseError.errorCode <> 0 Then
		lErrorNumber = oXMLDOM.parseError.errorCode
		sErrorDescription = "La estructura XML no pudo ser cargada en el objeto XML DOM."
		If Len(oXMLDOM.parseError.reason) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & oXMLDOM.parseError.reason
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "XMLLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CreateXMLDOMObject = lErrorNumber
	Err.Clear
End Function

Function LoadXMLToObject(vXML, oXML, sErrorDescription)
'************************************************************
'Purpose: To load an XML into the given object
'Inputs:  vXML
'Outputs: oXML, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "LoadXMLToObject"
	Dim lErrorNumber

	If Not IsObject(vXML) Then
		lErrorNumber = CreateXMLDOMObject(oXML, sErrorDescription)
		If lErrorNumber = 0 Then
			If StrComp(Right(vXML, Len(".xml")), ".xml", vbTextCompare) = 0 Then
				oXML.load(Server.MapPath(vXML))
				If Err.number <> 0 Then
					Err.Clear 
					oXML.load(vXML)
					If oXML.parseError.errorCode = -1072896760 Then
						lErrorNumber = CreateXMLDOMObject(oXML, sErrorDescription)
						If lErrorNumber = 0 Then
							oXML.loadXML(CleanStringForXML(GetFileContents(vXML, sErrorDescription), False))
						End If
					End If
				End If
			Else
				oXML.loadXML(CStr(vXML))
				If oXML.parseError.errorCode = -1072896760 Then
					Err.clear
					lErrorNumber = CreateXMLDOMObject(oXML, sErrorDescription)
					If lErrorNumber = 0 Then
						oXML.loadXML(CleanStringForXML(CStr(vXML), False))
					End If
				End If
			End If
			lErrorNumber = Err.number
			If lErrorNumber <> 0 Then
				sErrorDescription = "La estructura XML no pudo ser cargada en el objeto XML DOM. Puede que la estructura contenga errores de sintaxis. Favor de contactar al Administrador."
				If Len(Err.description) > 0 Then
					sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
				End If
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "XMLLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
			ElseIf oXML.parseError.errorCode <> 0 Then
				lErrorNumber = oXML.parseError.errorCode
				sErrorDescription = "La estructura XML no pudo ser cargada en el objeto XML DOM."
				If Len(oXML.parseError.reason) > 0 Then
					sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & oXML.parseError.reason
				End If
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "XMLLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
			End If
		End If
	Else
		Set oXML = vXML
	End If

	LoadXMLToObject = lErrorNumber
	Err.Clear
End Function

Function MergeXMLWithXSL(vXML, vXSL, sResult, sErrorDescription)
'************************************************************
'Purpose: To send the data stored in an XML to an XSL and get
'         the text generated by this file
'Inputs:  vXML, vXSL
'Outputs: sResult, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "MergeXMLWithXSL"
	Dim oTempXML
	Dim oTempXSL
	Dim lErrorNumber

	If Not IsObject(vXML) Then
		lErrorNumber = CreateXMLDOMObject(oTempXML, sErrorDescription)
		If lErrorNumber = 0 Then
			If StrComp(Right(vXML, Len(".xml")), ".xml", vbTextCompare) = 0 Then
				oTempXML.load(Server.MapPath(vXML))
				If Err.number <> 0 Then
					Err.Clear 
					oTempXML.load(vXML)
				End If
			Else
				oTempXML.loadXML(CStr(vXML))
			End If
			lErrorNumber = Err.number
			If lErrorNumber <> 0 Then
				sErrorDescription = "La estructura XML no pudo ser cargada en el objeto XML DOM. Puede que la estructura contenga errores de sintaxis. Favor de contactar al Administrador."
				If Len(Err.description) > 0 Then
					sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
				End If
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "XMLLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
			ElseIf oTempXML.parseError.errorCode <> 0 Then
				lErrorNumber = oTempXML.parseError.errorCode
				sErrorDescription = "La estructura XML no pudo ser cargada en el objeto XML DOM."
				If Len(oTempXML.parseError.reason) > 0 Then
					sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & oTempXML.parseError.reason
				End If
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "XMLLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
			End If
		End If
	Else
		Set oTempXML = vXML
	End If
	
	If lErrorNumber = 0 Then
		If Not IsObject(vXSL) Then
			lErrorNumber = CreateXMLDOMObject(oTempXSL, sErrorDescription)
			If lErrorNumber = 0 Then
				If StrComp(Right(vXSL, Len(".xsl")), ".xsl", vbTextCompare) = 0 Then
					oTempXSL.load(Server.MapPath(vXSL))
					If Err.number <> 0 Then
						Err.Clear 
						oTempXSL.load(vXSL)
					End If
				Else
					oTempXSL.loadXML(CStr(vXSL))
				End If
				lErrorNumber = Err.number
				If lErrorNumber <> 0 Then
					sErrorDescription = "La estructura XSL no pudo ser cargada en el objeto XML DOM. Puede que la estructura contenga errores de sintaxis. Favor de contactar al Administrador."
					If Len(Err.description) > 0 Then
						sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
					End If
					Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "XMLLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
				ElseIf oTempXSL.parseError.errorCode <> 0 Then
					lErrorNumber = oTempXSL.parseError.errorCode
					sErrorDescription = "La estructura XML no pudo ser cargada en el objeto XML DOM."
					If Len(oTempXSL.parseError.reason) > 0 Then
						sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & oTempXSL.parseError.reason
					End If
					Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "XMLLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
				End If
			End If
		Else
			Set oTempXSL = vXSL
		End If
	End If

	sResult = oTempXML.transformNode(oTempXSL)
	lErrorNumber = Err.Number
	If lErrorNumber <> 0 Then
		sErrorDescription = "The objeto XML DOM no pudo transformarse usando el XSL provisto."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "XMLLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	ElseIf oTempXML.parseError.errorCode <> 0 Then
		lErrorNumber = oTempXML.parseError.errorCode
		sErrorDescription = "La estructura XML no pudo ser cargada en el objeto XML DOM."
		If Len(oTempXML.parseError.reason) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & oTempXML.parseError.reason
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "XMLLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	Set oTempXML = Nothing
	Set oTempXSL = Nothing
	MergeXMLNodeAndXSL = lErrorNumber
	Err.Clear
End Function

Function RemoveNodeFromXML(sNodePath, sCondition, oXML, sErrorDescription)
'************************************************************
'Purpose: Remove the node located in the given path from the
'         given instance of the XML DOM
'Inputs:  sNodePath
'Outputs: oXML, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveNodeFromXML"
	Dim oNode
	Dim lErrorNumber

	If Len(sCondition) > 0 Then
		sCondition = "[" & sCondition & "]"
	End If

	Set oNode = oXML.selectSingleNode(sNodePath & sCondition)
	lErrorNumber = Err.Number
	If lErrorNumber <> 0 Or oNode Is Nothing Then
		sErrorDescription = "No se pudo remover el nodo de la estructura XML. Puede que no exista o que la condición se encuentre mal."
		If Len(Err.description) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "XMLLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	ElseIf oXML.parseError.errorCode <> 0 Then
		lErrorNumber = oXML.parseError.errorCode
		sErrorDescription = "La estructura XML no pudo ser cargada en el objeto XML DOM."
		If Len(oXML.parseError.reason) > 0 Then
			sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & oXML.parseError.reason
		End If
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "XMLLibrary.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	Else
		Call oXML.documentElement.removeChild(oNode)
	End If

	RemoveNodeFromXML = lErrorNumber
	Err.Clear
End Function
%>