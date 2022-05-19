<%
Function GetProjectsURLValues(oRequest, iSelectedTab, bAction)
'************************************************************
'Purpose: To initialize the global variables using the URL
'Inputs:  oRequest
'Outputs: iSelectedTab, bAction
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetProjectsURLValues"
	Dim sCondition

	iSelectedTab = 1
	bAction = (Len(oRequest("UpdateTaskStatus").Item) > 0) Or (Len(oRequest("RemoveFile").Item) > 0) Or (StrComp(oRequest("Action").Item, "HistoryList", vbBinaryCompare) = 0)
	bShowSearchForm = (Len(oRequest("DoSearch").Item) = 0) And (Len(oRequest("ShowInfo").Item) = 0) And Not bShowForm And Not bAction

	GetProjectsURLValues = Err.number
	Err.Clear
End Function

Function DoProjectsAction(oRequest, oADODBConnection, iSelectedTab, sErrorDescription)
'************************************************************
'Purpose: To add, change or delete the project information
'Inputs:  oRequest, oADODBConnection, iSelectedTab
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DoProjectsAction"
	Dim aPath
	Dim iIndex
	Dim sAnswerID
	Dim bClean
	Dim oRecordset
	Dim lErrorNumber

	If Len(oRequest("RemoveFile").Item) > 0 Then
		If FileExists(oRequest("FilePath").Item & "\" & oRequest("FileName").Item, sErrorDescription) Then
			lErrorNumber = DeleteFile(oRequest("FilePath").Item & "\" & oRequest("FileName").Item, sErrorDescription)
		End If
	ElseIf Len(oRequest("UpdateTaskStatus").Item) > 0 Then
		lErrorNumber = UpdateTaskStatus(oRequest, oADODBConnection, False, aTaskComponent, sErrorDescription)
		If (lErrorNumber = 0) And (Len(oRequest("FromParent").Item) > 0) Then
			aPath = Split(aTaskComponent(S_PATH_TASK), ",", -1, vbBinaryCompare)
			If UBound(aPath) < 2 Then
				Response.Redirect "Projects.asp?View=" & oRequest("View").Item & "&ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & "&AreaID=" & oRequest("AreaID").Item & "&CategoryID=" & oRequest("CategoryID").Item & "&UserID=" & oRequest("UserID").Item & "&Blend=1"
			Else
				Response.Redirect "Projects.asp?View=" & oRequest("View").Item & "&ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & "&TaskID=" & aTaskComponent(N_PARENT_ID_TASK) & "&ParentID=" & aPath(UBound(aPath) - 2) & "&TaskPath=" & Left(aTaskComponent(S_PATH_TASK), (InStrRev(aTaskComponent(S_PATH_TASK), ",") - Len(","))) & "&AreaID=" & oRequest("AreaID").Item & "&CategoryID=" & oRequest("CategoryID").Item & "&UserID=" & oRequest("UserID").Item & "&Blend=1"
			End If
		End If
	ElseIf StrComp(oRequest("Action").Item, "HistoryList", vbBinaryCompare) = 0 Then
		If Len(oRequest("Add").Item) > 0 Then
			lErrorNumber = AddEventToHistoryList(oRequest, oADODBConnection, aHistoryComponent, sErrorDescription)
			bClean = (lErrorNumber = 0)

			If (lErrorNumber = 0) And B_USE_SMTP Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ProjectNumber, ProjectName, TaskNumber, UserEmail From " & TACO_PREFIX & "Projects, " & TACO_PREFIX & "Tasks, " & TACO_PREFIX & "TaskUsersLKP, Users Where (" & TACO_PREFIX & "Projects.ProjectID=" & TACO_PREFIX & "Tasks.ProjectID) And (" & TACO_PREFIX & "Tasks.ProjectID=" & TACO_PREFIX & "TaskUsersLKP.ProjectID) And (" & TACO_PREFIX & "Tasks.TaskID=" & TACO_PREFIX & "TaskUsersLKP.TaskID) And (" & TACO_PREFIX & "TaskUsersLKP.UserID=Users.UserID) And (" & TACO_PREFIX & "Tasks.ProjectID=" & aHistoryComponent(N_PROJECT_ID_HISTORY) & ") And (" & TACO_PREFIX & "Tasks.TaskID=" & aHistoryComponent(N_TASK_ID_HISTORY) & ") Order By UserLastName, UserName", "TaCoProjectsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						ReDim aEmailComponent(N_EMAIL_COMPONENT_SIZE)
						aEmailComponent(S_TO_EMAIL) = ""

						aEmailComponent(S_FROM_EMAIL) = aLoginComponent(S_USER_E_MAIL_LOGIN)
						aEmailComponent(S_SUBJECT_EMAIL) = "Nuevo comentario para la actividad " & CleanStringForHTML(CStr(oRecordset.Fields("TaskNumber").Value)) & " del proceso " & CleanStringForHTML(CStr(oRecordset.Fields("ProjectNumber").Value) & " " & CStr(oRecordset.Fields("ProjectName").Value))

						aEmailComponent(S_BODY_EMAIL) = GetFileContents(Server.MapPath("Template_NewComment.htm"), sErrorDescription)
						If (Len(aEmailComponent(S_FROM_EMAIL)) > 0) And (Len(aEmailComponent(S_TO_EMAIL)) > 0) And (Len(aEmailComponent(S_BODY_EMAIL)) > 0) Then
							aEmailComponent(S_BODY_EMAIL) = Replace(aEmailComponent(S_BODY_EMAIL), "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
							aEmailComponent(S_BODY_EMAIL) = Replace(aEmailComponent(S_BODY_EMAIL), "<SYSTEM_IP />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
							aEmailComponent(S_BODY_EMAIL) = Replace(aEmailComponent(S_BODY_EMAIL), "<EXT_SYSTEM_URL />", S_HTTP & EXT_SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
							aEmailComponent(S_BODY_EMAIL) = Replace(aEmailComponent(S_BODY_EMAIL), "<EXT_SYSTEM_IP />", S_HTTP & EXT_SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
							aEmailComponent(S_BODY_EMAIL) = Replace(aEmailComponent(S_BODY_EMAIL), "<TASK_NUMBER />", CleanStringForHTML(CStr(oRecordset.Fields("TaskNumber").Value)))
							aEmailComponent(S_BODY_EMAIL) = Replace(aEmailComponent(S_BODY_EMAIL), "<PROYECT_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("ProjectNumber").Value) & " " & CStr(oRecordset.Fields("ProjectName").Value)))
							aEmailComponent(S_BODY_EMAIL) = Replace(aEmailComponent(S_BODY_EMAIL), "<USER_NAME />", CleanStringForHTML(aLoginComponent(S_USER_NAME_LOGIN) & " " & aLoginComponent(S_USER_LAST_NAME_LOGIN)))
							aEmailComponent(S_BODY_EMAIL) = Replace(aEmailComponent(S_BODY_EMAIL), "<COMMENTS />", aHistoryComponent(S_DESCRIPTION_HISTORY))
							aEmailComponent(S_BODY_EMAIL) = Replace(aEmailComponent(S_BODY_EMAIL), "<TASK_URL />", "ProjectID=" & aHistoryComponent(N_PROJECT_ID_HISTORY) & "&TaskID=" & aHistoryComponent(N_TASK_ID_HISTORY) & "&ParentID=" & oRequest("ParentID").Item & "&TaskPath=" & oRequest("TaskPath").Item)
							lErrorNumber = SendEmail(oRequest, aEmailComponent, sErrorDescription)
						End If

						Do While Not oRecordset.EOF
							aEmailComponent(S_TO_EMAIL) = aEmailComponent(S_TO_EMAIL) & CStr(oRecordset.Fields("UserEmail").Value) & ","
							oRecordset.MoveNext
						Loop
						aEmailComponent(S_TO_EMAIL) = Left(aEmailComponent(S_TO_EMAIL), (Len(aEmailComponent(S_TO_EMAIL)) - Len(",")))
						If (Len(aEmailComponent(S_FROM_EMAIL)) > 0) And (Len(aEmailComponent(S_TO_EMAIL)) > 0) And (Len(aEmailComponent(S_BODY_EMAIL)) > 0) Then
							lErrorNumber = SendEmail(oRequest, aEmailComponent, sErrorDescription)
						End If
					End If
					oRecordset.Close
				End If
			End If
		ElseIf Len(oRequest("Modify").Item) > 0 Then
			lErrorNumber = ModifyEventOnHistoryList(oRequest, oADODBConnection, aHistoryComponent, sErrorDescription)
			bClean = (lErrorNumber = 0)
		ElseIf Len(oRequest("Remove").Item) > 0 Then
			lErrorNumber = RemoveEventOnHistoryList(oRequest, oADODBConnection, aHistoryComponent, sErrorDescription)
			bClean = (lErrorNumber = 0)
		End If
		If bClean Then
			aHistoryComponent(N_RECORD_ID_HISTORY) = -1
			aHistoryComponent(N_DATE_HISTORY) = 0
			aHistoryComponent(N_HOUR_HISTORY) = 0
			aHistoryComponent(N_MINUTE_HISTORY) = 0
			aHistoryComponent(S_DESCRIPTION_HISTORY) = ""
		End If
	End If

	Set oRecordset = Nothing
	DoProjectsAction = lErrorNumber
	Err.Clear
End Function

Function DisplayAreasOrCategoriesForProject(oADODBConnection, lProjectID, sView, sErrorDescription)
'************************************************************
'Purpose: To display the project advance
'Inputs:  oADODBConnection, lProjectID, sView
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayAreasOrCategoriesForProject"
	Dim sCondition
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	sCondition = ""
	sErrorDescription = ""
	Select Case sView
		Case "Areas"
			'sCondition = " And (" & TACO_PREFIX & "TaskAreasLKP.AreaID In ())"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct " & TACO_PREFIX & "Areas.AreaID, AreaName, CompanyName From " & TACO_PREFIX & "Areas, " & TACO_PREFIX & "Companies, " & TACO_PREFIX & "TaskAreasLKP Where (" & TACO_PREFIX & "Areas.CompanyID=" & TACO_PREFIX & "Companies.CompanyID) And (" & TACO_PREFIX & "Areas.AreaID=" & TACO_PREFIX & "TaskAreasLKP.AreaID) And (" & TACO_PREFIX & "TaskAreasLKP.ProjectID=" & lProjectID & ")" & sCondition & " Order By CompanyName, AreaName", "TaCoProjectsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			asColumnsTitles = Split("Compañía,Área", ",", -1, vbBinaryCompare)
			asCellWidths = Split("200,600,", ",", -1, vbBinaryCompare)
		Case "Categories"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct " & TACO_PREFIX & "Categories.CategoryID, CategoryName From " & TACO_PREFIX & "Categories, " & TACO_PREFIX & "TaskCategoriesLKP Where (" & TACO_PREFIX & "Categories.CategoryID=" & TACO_PREFIX & "TaskCategoriesLKP.CategoryID) And (" & TACO_PREFIX & "TaskCategoriesLKP.ProjectID=" & lProjectID & ")" & sCondition & " Order By CategoryName", "TaCoProjectsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			asColumnsTitles = Split("Categoría", ",", -1, vbBinaryCompare)
			asCellWidths = Split("800,", ",", -1, vbBinaryCompare)
		Case "Users"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct Users.UserID, UserName, UserLastName From Users, " & TACO_PREFIX & "TaskUsersLKP Where (Users.UserID=" & TACO_PREFIX & "TaskUsersLKP.UserID) And (" & TACO_PREFIX & "TaskUsersLKP.ProjectID=" & lProjectID & ")" & sCondition & " Order By UserLastName, UserName", "TaCoProjectsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			asColumnsTitles = Split("Responsable", ",", -1, vbBinaryCompare)
			asCellWidths = Split("800,", ",", -1, vbBinaryCompare)
	End Select
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE WIDTH=""800"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If
				asCellAlignments = Split(",", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					Select Case sView
						Case "Areas"
							sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("CompanyName").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""Projects.asp?View=" & sView & "&ProjectID=" & lProjectID & "&AreaID=" & CStr(oRecordset.Fields("AreaID").Value) & """>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value)) & "</A>"
						Case "Categories"
							sRowContents = "<A HREF=""Projects.asp?View=" & sView & "&ProjectID=" & lProjectID & "&CategoryID=" & CStr(oRecordset.Fields("CategoryID").Value) & """>" & CleanStringForHTML(CStr(oRecordset.Fields("CategoryName").Value)) & "</A>"
						Case "Users"
							sRowContents = "<A HREF=""Projects.asp?View=" & sView & "&ProjectID=" & lProjectID & "&UserID=" & CStr(oRecordset.Fields("UserID").Value) & """>" & CleanStringForHTML(CStr(oRecordset.Fields("UserName").Value) & " " & CStr(oRecordset.Fields("UserLastName").Value)) & "</A>"
					End Select
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE>"
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			Select Case sView
				Case "Areas"
					sErrorDescription = "Las actividades de este proceso no están clasificadas por áreas."
				Case "Categories"
					sErrorDescription = "Las actividades de este proceso no están agrupadas por categorías."
				Case "Users"
					sErrorDescription = "Las actividades de este proceso no están clasificadas por usuario."
			End Select
		End If
	End If

	Set oRecordset = Nothing
	DisplayAreasOrCategoriesForProject = lErrorNumber
	Err.Clear
End Function

Function DisplayProject(oADODBConnection, lProjectID, sErrorDescription)
'************************************************************
'Purpose: To display the project information
'Inputs:  oADODBConnection, lProjectID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayProject"
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = ""
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From " & TACO_PREFIX & "Projects Where (ProjectID=" & lProjectID & ")", "TaCoProjectsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<FONT FACE=""Arial"" SIZE=""2"">"
				Response.Write "<B>Descripción del proceso: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("ProjectDescription").Value)) & "<BR /><BR />"
				Response.Write "<B>Objetivo del proceso: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("ProjectObjective").Value)) & "<BR />"
			Response.Write "</FONT>"
		End If
	End If

	Set oRecordset = Nothing
	DisplayProject = lErrorNumber
	Err.Clear
End Function

Function DisplayProjects(oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the projects in the system
'Inputs:  oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayProjects"
	Dim oRecordset
	Dim iCounter
	Dim lErrorNumber

	sErrorDescription = ""
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From " & TACO_PREFIX & "Projects Where ((ProjectFile Like '" & S_WILD_CHAR & iGlobalSectionID & S_WILD_CHAR & "') Or (ProjectFile Like '" & S_WILD_CHAR & "-1" & S_WILD_CHAR & "')) And (Active=1) Order By ProjectNumber, ProjectName", "TaCoProjectsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If (lErrorNumber = 0) Then
		Response.Write "<TABLE WIDTH=""789"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			Response.Write "<TR>" & vbNewLine
				iCounter = 0
				Do While Not oRecordset.EOF
					Response.Write "<TD WIDTH=""260"" VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"
						If Not bIsNetscape Then
							Response.Write "<DIV ID=""WidgetInfo" & iCounter & "Div"" CLASS=""WidgetInfo"" onMouseOver=""ShowWidget('" & iCounter & "')"" onMouseOut=""HideWidget('" & iCounter & "')"">"
								Response.Write "<TABLE BGCOLOR=""#" & S_LIGHT_BGCOLOR & """ WIDTH=""256"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
									Response.Write "<TR>"
										Response.Write "<TD WIDTH=""4""><IMG SRC=""Images/CrnWidgetTpLf.gif"" WIDTH=""4"" HEIGHT=""4"" /></TD>"
										Response.Write "<TD BACKGROUND=""Images/FrmWidgetTp.gif"" WIDTH=""70""><IMG SRC=""Images/Transparent.gif"" WIDTH=""70"" HEIGHT=""4"" /></TD>"
										Response.Write "<TD BACKGROUND=""Images/FrmWidgetTp.gif"" WIDTH=""178""><IMG SRC=""Images/Transparent.gif"" WIDTH=""178"" HEIGHT=""4"" /></TD>"
										Response.Write "<TD WIDTH=""4""><IMG SRC=""Images/CrnWidgetTpRg.gif"" WIDTH=""4"" HEIGHT=""4"" /></TD>"
									Response.Write "</TR>"
									Response.Write "<TR>"
										Response.Write "<TD BACKGROUND=""Images/FrmWidgetLf.gif""><IMG SRC=""Images/Transparent.gif"" WIDTH=""4"" HEIGHT=""1"" /></TD>"
										Response.Write "<TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2""><B>"
											Response.Write "<A HREF=""Projects.asp?ProjectID=" & CStr(oRecordset.Fields("ProjectID").Value) & """ STYLE=""text-decoration: none"">"
												Response.Write CleanStringForHTML(CStr(oRecordset.Fields("ProjectNumber").Value) & ". " & CStr(oRecordset.Fields("ProjectName").Value))
											Response.Write "</A>"
										Response.Write "</B></FONT></TD>"
										Response.Write "<TD BACKGROUND=""Images/FrmWidgetRg.gif""><IMG SRC=""Images/Transparent.gif"" WIDTH=""4"" HEIGHT=""1"" /></TD>"
									Response.Write "</TR>"
									Response.Write "<TR>"
										Response.Write "<TD BACKGROUND=""Images/FrmWidgetLf.gif""><IMG SRC=""Images/Transparent.gif"" WIDTH=""4"" HEIGHT=""1"" /></TD>"
										Response.Write "<TD VALIGN=""TOP"">"
											Response.Write "<A HREF=""Projects.asp?ProjectID=" & CStr(oRecordset.Fields("ProjectID").Value) & """>"
												Response.Write "<IMG SRC=""Images/MnProjects0" & (iCounter Mod 9) & ".gif"" WIDTH=""64"" HEIGHT=""64"" ALT=""" & CStr(oRecordset.Fields("ProjectNumber").Value) & ". " & CStr(oRecordset.Fields("ProjectName").Value) & """ BORDER=""0"" />"
											Response.Write "</A><BR />"
											Response.Write "<FONT FACE=""Verdana"" SIZE=""1""><BR />"
												Response.Write "<A HREF=""Projects.asp?ProjectID=" & CStr(oRecordset.Fields("ProjectID").Value) & """ CLASS=""SpecialLink"">Entrar&nbsp;"
												Response.Write "<IMG SRC=""Images/BtnArrSmallRight.gif"" WIDTH=""11"" HEIGHT=""10"" ALT=""Entrar"" BORDER=""0"" /></A>"
											Response.Write "</FONT>"
										Response.Write "</TD>"
										Response.Write "<TD VALIGN=""TOP""><DIV STYLE=""width: 178px; height: 200px; overflow: auto;""><FONT FACE=""Verdana"" SIZE=""1"">"
											Response.Write "<B>Sección: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("ProjectOwner").Value)) & "<BR /><BR />"
											Response.Write "<B>Descripción: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("ProjectDescription").Value)) & "<BR /><BR />"
											Response.Write "<B>Objetivo: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("ProjectObjective").Value)) & "<BR />"
										Response.Write "</FONT></DIV></TD>"
										Response.Write "<TD BACKGROUND=""Images/FrmWidgetRg.gif""><IMG SRC=""Images/Transparent.gif"" WIDTH=""4"" HEIGHT=""1"" /></TD>"
									Response.Write "</TR>"
									Response.Write "<TR>"
										Response.Write "<TD><IMG SRC=""Images/CrnWidgetBtLf.gif"" WIDTH=""4"" HEIGHT=""4"" /></TD>"
										Response.Write "<TD BACKGROUND=""Images/FrmWidgetBt.gif"" COLSPAN=""2""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""4"" /></TD>"
										Response.Write "<TD><IMG SRC=""Images/CrnWidgetBtRg.gif"" WIDTH=""4"" HEIGHT=""4"" /></TD>"
									Response.Write "</TR>"
								Response.Write "</TABLE>"
							Response.Write "</DIV>"
							Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""4"" /><BR /><IMG SRC=""Images/Transparent.gif"" WIDTH=""4"" HEIGHT=""150"" ALIGN=""LEFT"" HSPACE=""0"" />"
						End If
						Response.Write "<A HREF=""Projects.asp?ProjectID=" & CStr(oRecordset.Fields("ProjectID").Value) & """"
							If Not bIsNetscape Then Response.Write " onMouseOver=""ShowWidget('" & iCounter & "')"""
						Response.Write " STYLE=""text-decoration: none"">"
							Response.Write "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("ProjectNumber").Value) & ". " & CStr(oRecordset.Fields("ProjectName").Value)) & "</B><BR />"
							Response.Write "<IMG SRC=""Images/MnProjects0" & (iCounter Mod 9) & ".gif"" WIDTH=""64"" HEIGHT=""64"" ALT=""" & CStr(oRecordset.Fields("ProjectNumber").Value) & ". " & CStr(oRecordset.Fields("ProjectName").Value) & """ BORDER=""0"" />"
						Response.Write "</A><BR />"
						Response.Write "<FONT FACE=""Verdana"" SIZE=""1""><BR />"
							If Not bIsNetscape Then
								Response.Write "<A HREF=""Projects.asp?ProjectID=" & CStr(oRecordset.Fields("ProjectID").Value) & """ CLASS=""SpecialLink"">Entrar&nbsp;"
								Response.Write "<IMG SRC=""Images/BtnArrSmallRight.gif"" WIDTH=""11"" HEIGHT=""10"" ALT=""Entrar"" BORDER=""0"" /></A>"
							Else
								Response.Write "<B>Sección: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("ProjectOwner").Value)) & "<BR /><BR />"
								Response.Write "<B>Descripción: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("ProjectDescription").Value)) & "<BR /><BR />"
								Response.Write "<B>Objetivo: </B>" & CleanStringForHTML(CStr(oRecordset.Fields("ProjectObjective").Value)) & "<BR />"
							End If
						Response.Write "</FONT>"
					Response.Write "</FONT></TD>" & vbNewLine
					Response.Write "<TD>&nbsp;&nbsp;</TD>" & vbNewLine
					iCounter = iCounter + 1
					If (iCounter Mod 3) = 0 Then Response.Write "</TR><TR>" & vbNewLine
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			Response.Write "</TR>" & vbNewLine
		Response.Write "</TABLE>" & vbNewLine
	End If

	Set oRecordset = Nothing
	DisplayProjects = lErrorNumber
	Err.Clear
End Function

Function GetTaskAdvanceAsRowContents(oRequest, oRecordset, lParentID, iAggregationType, aTaskComponent, sView, lRecordID, sTresholdColor)
'************************************************************
'Purpose: To build the HTML for the row contents
'Inputs:  oRequest, oRecordset, lParentID, iAggregationType, aTaskComponent, sView, lRecordID, sTresholdColor
'Outputs: sTresholdColor. A String containing the row information
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetTaskAdvanceAsRowContents"
	Dim sURL
	Dim bHasChildren

	bHasChildren = TaskHasChildren(oADODBConnection, CLng(oRecordset.Fields("ProjectID").Value), CLng(oRecordset.Fields("TaskID").Value), sErrorDescription)
	sURL = "Projects.asp?View="
	If Len(sView) = 0 Then
		sURL = sURL & oRequest("View").Item
	Else
		sURL = sURL & sView
	End If
	sURL = sURL & "&ProjectID=" & CStr(oRecordset.Fields("ProjectID").Value) & "&TaskID=" & CStr(oRecordset.Fields("TaskID").Value) & "&ParentID=" & lParentID & "&TaskPath=" & aTaskComponent(S_PATH_TASK) & "," & CStr(oRecordset.Fields("TaskID").Value)
	Select Case sView
		Case "Areas"
			sURL = sURL & "&AreaID=" & lRecordID
		Case "Categories"
			sURL = sURL & "&CategoryID=" & lRecordID
		Case "Users"
			sURL = sURL & "&UserID=" & lRecordID
		Case Else
			sURL = sURL & "&AreaID=" & oRequest("AreaID").Item & "&CategoryID=" & oRequest("CategoryID").Item & "&UserID=" & oRequest("UserID").Item
	End Select
	GetTaskAdvanceAsRowContents = "<A"
		If bHasChildren Then GetTaskAdvanceAsRowContents = GetTaskAdvanceAsRowContents & " HREF=""" & sURL & """"
	GetTaskAdvanceAsRowContents = GetTaskAdvanceAsRowContents & ">" & CleanStringForHTML(CStr(oRecordset.Fields("TaskNumber").Value)) & "</A>"
	GetTaskAdvanceAsRowContents = GetTaskAdvanceAsRowContents & TABLE_SEPARATOR & "<A"
		If bHasChildren Then GetTaskAdvanceAsRowContents = GetTaskAdvanceAsRowContents & " HREF=""" & sURL & """"
	GetTaskAdvanceAsRowContents = GetTaskAdvanceAsRowContents & ">" & CleanStringForHTML(CStr(oRecordset.Fields("TaskName").Value)) & "</A>"
	If Len(sView) = 0 Then
		If iAggregationType = 1 Then
			GetTaskAdvanceAsRowContents = GetTaskAdvanceAsRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("TargetValue").Value), 2, True, False, True) & "%"
		Else
			GetTaskAdvanceAsRowContents = GetTaskAdvanceAsRowContents & TABLE_SEPARATOR & FormatNumber((CDbl(oRecordset.Fields("TaskPercentage").Value) * 100), 2, True, False, True) & "%"
		End If
	End If
	GetTaskAdvanceAsRowContents = GetTaskAdvanceAsRowContents & TABLE_SEPARATOR & FormatNumber((CDbl(oRecordset.Fields("TaskStatusPercentage").Value) * 100), 2, True, False, True) & "%"
	Call GetColorForTreshold(CDbl(oRecordset.Fields("TaskStatusPercentage").Value), sTresholdColor)
	sURL = "TaskPath=" & aTaskComponent(S_PATH_TASK) & "," & CStr(oRecordset.Fields("TaskID").Value) & "&TaskID=" & CStr(oRecordset.Fields("TaskID").Value) & "&TaskStatusPercentage=" & FormatNumber((CDbl(oRecordset.Fields("TaskStatusPercentage").Value) * 100), 2, True, False, True)
	Select Case CInt(GetOption(aOptionsComponent, TRESHOLD_STYLE_OPTION))
		Case 1
			GetTaskAdvanceAsRowContents = GetTaskAdvanceAsRowContents & TABLE_SEPARATOR & "<SPAN STYLE=""background-color: #" & sTresholdColor & """"
				If Not bHasChildren Then GetTaskAdvanceAsRowContents = GetTaskAdvanceAsRowContents & " onClick=""SendURLValuesToForm('" & sURL & "', document.NewTaskStatusFrm); MovePopupItem('WidgetInfoDiv', document.all['WidgetInfoDiv'], event.x-20, event.y-30); ShowWidget('');"""
			GetTaskAdvanceAsRowContents = GetTaskAdvanceAsRowContents & "><IMG SRC=""Images/IcnTreshold.gif"" WIDTH=""16"" HEIGHT=""16""></SPAN>"
			sTresholdColor = ""
		Case 2
			GetTaskAdvanceAsRowContents = GetTaskAdvanceAsRowContents & TABLE_SEPARATOR
			If bHasChildren Then 
				GetTaskAdvanceAsRowContents = GetTaskAdvanceAsRowContents & GetPercentageBar(CDbl(oRecordset.Fields("TaskStatusPercentage").Value), sTresholdColor, "", "", "")
			Else
				GetTaskAdvanceAsRowContents = GetTaskAdvanceAsRowContents & GetPercentageBar(CDbl(oRecordset.Fields("TaskStatusPercentage").Value), sTresholdColor, "SendURLValuesToForm('" & sURL & "', document.NewTaskStatusFrm); MovePopupItem('WidgetInfoDiv', document.all['WidgetInfoDiv'], event.x-20, event.y-30); ShowWidget('');", "", "")
			End If
			sTresholdColor = ""
		Case 3
	End Select

	Err.Clear
End Function

Function DisplayFullProjectAdvance(oADODBConnection, lProjectID, lParentID, sErrorDescription)
'************************************************************
'Purpose: To display the project advance
'Inputs:  oADODBConnection, lProjectID, lParentID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayFullProjectAdvance"

	Err.Clear
End Function

Function DisplayProjectAdvance(oADODBConnection, lProjectID, lParentID, sErrorDescription)
'************************************************************
'Purpose: To display the project advance
'Inputs:  oADODBConnection, lProjectID, lParentID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayProjectAdvance"
	Dim sTresholdColor
	Dim iAggregationType
	Dim dTotal
	Dim dAdvance
	Dim oRecordset
	Dim lCurrentTaskID
	Dim sURL
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	sTresholdColor = "FFFFFF"
	sErrorDescription = ""
	iAggregationType = 0
	sErrorDescription = "No se pudo obtener el avance de las actividades."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AggregationTypeID From " & TACO_PREFIX & "Tasks Where (ProjectID=" & lProjectID & ") And (TaskID=" & lParentID & ")", "TaCoProjectsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then iAggregationType = CInt(oRecordset.Fields("AggregationTypeID").Value)
		oRecordset.Close
	End If

	sErrorDescription = "No se pudo obtener el avance de las actividades."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select " & TACO_PREFIX & "Tasks.ProjectID, " & TACO_PREFIX & "Tasks.TaskID, TaskNumber, TaskName, " & TACO_PREFIX & "TasksLKP.TargetValue, " & TACO_PREFIX & "TasksLKP.TaskPercentage, " & TACO_PREFIX & "TasksStatusLKP.TaskPercentage As TaskStatusPercentage, LabelName From " & TACO_PREFIX & "Tasks, " & TACO_PREFIX & "TasksLKP, " & TACO_PREFIX & "TasksStatusLKP, " & TACO_PREFIX & "Labels Where (" & TACO_PREFIX & "Tasks.ProjectID=" & TACO_PREFIX & "TasksLKP.ProjectID) And (" & TACO_PREFIX & "Tasks.TaskID=" & TACO_PREFIX & "TasksLKP.TaskID) And (" & TACO_PREFIX & "TasksLKP.ProjectID=" & TACO_PREFIX & "TasksStatusLKP.ProjectID) And (" & TACO_PREFIX & "TasksLKP.TaskID=" & TACO_PREFIX & "TasksStatusLKP.TaskID) And (" & TACO_PREFIX & "TasksLKP.ParentID=" & TACO_PREFIX & "TasksStatusLKP.ParentID) And (" & TACO_PREFIX & "Tasks.LabelID=" & TACO_PREFIX & "Labels.LabelID) And (" & TACO_PREFIX & "Tasks.ProjectID=" & lProjectID & ") And (" & TACO_PREFIX & "TasksLKP.ParentID=" & lParentID & ") Order By TaskNumber, TaskName, StatusDate Desc", "TaCoProjectsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE WIDTH=""800"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				If CInt(GetOption(aOptionsComponent, TRESHOLD_STYLE_OPTION)) = 3 Then
					asColumnsTitles = Split("Clave," & CleanStringForHTML(CStr(oRecordset.Fields("LabelName").Value)) & ",Participación,Avance", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,500,100,100,", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("Clave," & CleanStringForHTML(CStr(oRecordset.Fields("LabelName").Value)) & ",Participación,Avance,Semáforo", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,400,100,100,100", ",", -1, vbBinaryCompare)
				End If
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If
				asCellAlignments = Split(",,RIGHT,RIGHT,CENTER", ",", -1, vbBinaryCompare)
				dTotal = 0
				dAdvance = 0
				lCurrentDate = -1
				Do While Not oRecordset.EOF
					If lCurrentTaskID <> CLng(oRecordset.Fields("TaskID").Value) Then
						sRowContents = GetTaskAdvanceAsRowContents(oRequest, oRecordset, lParentID, iAggregationType, aTaskComponent, "", -1, sTresholdColor)
						If iAggregationType = 1 Then
							dTotal = dTotal + CDbl(oRecordset.Fields("TargetValue").Value)
							dAdvance = dAdvance + CDbl(oRecordset.Fields("TaskStatusPercentage").Value)
						Else
							dTotal = dTotal + CDbl(oRecordset.Fields("TaskPercentage").Value)
							dAdvance = dAdvance + (CDbl(oRecordset.Fields("TaskPercentage").Value) * CDbl(oRecordset.Fields("TaskStatusPercentage").Value))
						End If
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, sTresholdColor, "", "", "", sErrorDescription)
						lCurrentTaskID = CLng(oRecordset.Fields("TaskID").Value)
					End If
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				If iAggregationType <> 1 Then dTotal = dTotal * 100
				sRowContents = TABLE_SEPARATOR & TABLE_SEPARATOR & "<B>" & FormatNumber(dTotal, 2, True, False, True) & "%</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber((dAdvance * 100), 2, True, False, True) & "%</B>"
				Call GetColorForTreshold(dAdvance, sTresholdColor)
				Select Case CInt(GetOption(aOptionsComponent, TRESHOLD_STYLE_OPTION))
					Case 1
						sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN STYLE=""background-color: #" & sTresholdColor & """ onClick=""ShowTaskStatusFrm();""><IMG SRC=""Images/IcnTreshold.gif"" WIDTH=""16"" HEIGHT=""16""></SPAN>"
						sTresholdColor = ""
					Case 2
						sRowContents = sRowContents & TABLE_SEPARATOR & GetPercentageBar(dAdvance, sTresholdColor, "ShowTaskStatusFrm();", "", "")
						sTresholdColor = ""
					Case 3
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				End Select
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, sTresholdColor, "", "", "", sErrorDescription)
			Response.Write "</TABLE>"
		Else
			lErrorNumber = L_ERR_NO_RECORDS
		End If
	End If

	Set oRecordset = Nothing
	DisplayProjectAdvance = lErrorNumber
	Err.Clear
End Function

Function DisplayProjectAdvanceByAreasOrCategories(oADODBConnection, lProjectID, lParentID, lRecordID, sView, sErrorDescription)
'************************************************************
'Purpose: To display the project advance by area
'Inputs:  oADODBConnection, lProjectID, lParentID, lRecordID, sView
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayProjectAdvanceByAreasOrCategories"
	Dim sTresholdColor
	Dim sTables
	Dim sCondition
	Dim oRecordset
	Dim lCurrentTaskID
	Dim sURL
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	sTresholdColor = "FFFFFF"
	sTables = ""
	sCondition = ""
	If lParentID <> -1 Then
		sCondition = sCondition & " And (" & TACO_PREFIX & "TasksLKP.ParentID=" & lParentID & ")"
	ElseIf lRecordID <> -1 Then
		Select Case sView
			Case "Areas"
				sTables = ", " & TACO_PREFIX & "TaskAreasLKP"
				sCondition = sCondition & " And (" & TACO_PREFIX & "Tasks.ProjectID=" & TACO_PREFIX & "TaskAreasLKP.ProjectID) And (" & TACO_PREFIX & "Tasks.TaskID=" & TACO_PREFIX & "TaskAreasLKP.TaskID) And (" & TACO_PREFIX & "TaskAreasLKP.AreaID=" & lRecordID & ")"
			Case "Categories"
				sTables = ", " & TACO_PREFIX & "TaskCategoriesLKP"
				sCondition = sCondition & " And (" & TACO_PREFIX & "Tasks.ProjectID=" & TACO_PREFIX & "TaskCategoriesLKP.ProjectID) And (" & TACO_PREFIX & "Tasks.TaskID=" & TACO_PREFIX & "TaskCategoriesLKP.TaskID) And (" & TACO_PREFIX & "TaskCategoriesLKP.CategoryID=" & lRecordID & ")"
			Case "Users"
				sTables = ", " & TACO_PREFIX & "TaskUsersLKP"
				sCondition = sCondition & " And (" & TACO_PREFIX & "Tasks.ProjectID=" & TACO_PREFIX & "TaskUsersLKP.ProjectID) And (" & TACO_PREFIX & "Tasks.TaskID=" & TACO_PREFIX & "TaskUsersLKP.TaskID) And (" & TACO_PREFIX & "TaskUsersLKP.UserID=" & lRecordID & ")"
		End Select
	End If
	sErrorDescription = "No se pudo obtener la información de las actividades."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select " & TACO_PREFIX & "Tasks.ProjectID, " & TACO_PREFIX & "Tasks.TaskID, TaskNumber, TaskName, " & TACO_PREFIX & "TasksStatusLKP.TaskPercentage As TaskStatusPercentage, LabelName From " & TACO_PREFIX & "Tasks, " & TACO_PREFIX & "TasksLKP, " & TACO_PREFIX & "TasksStatusLKP, " & TACO_PREFIX & "Labels" & sTables & " Where (" & TACO_PREFIX & "Tasks.ProjectID=" & TACO_PREFIX & "TasksLKP.ProjectID) And (" & TACO_PREFIX & "Tasks.TaskID=" & TACO_PREFIX & "TasksLKP.TaskID) And (" & TACO_PREFIX & "TasksLKP.ProjectID=" & TACO_PREFIX & "TasksStatusLKP.ProjectID) And (" & TACO_PREFIX & "TasksLKP.TaskID=" & TACO_PREFIX & "TasksStatusLKP.TaskID) And (" & TACO_PREFIX & "TasksLKP.ParentID=" & TACO_PREFIX & "TasksStatusLKP.ParentID) And (" & TACO_PREFIX & "Tasks.LabelID=" & TACO_PREFIX & "Labels.LabelID) And (" & TACO_PREFIX & "Tasks.ProjectID=" & lProjectID & ")" & sCondition & " Order By TaskNumber, TaskName, StatusDate Desc", "TaCoProjectsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE WIDTH=""800"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				If CInt(GetOption(aOptionsComponent, TRESHOLD_STYLE_OPTION)) = 3 Then
					asColumnsTitles = Split("Clave," & CleanStringForHTML(CStr(oRecordset.Fields("LabelName").Value)) & ",Avance", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,600,100,100,", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("Clave," & CleanStringForHTML(CStr(oRecordset.Fields("LabelName").Value)) & ",Avance,Semáforo", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,500,100,100", ",", -1, vbBinaryCompare)
				End If
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If
				asCellAlignments = Split(",,RIGHT,CENTER", ",", -1, vbBinaryCompare)
				lCurrentTaskID = -1
				Do While Not oRecordset.EOF
					If lCurrentTaskID <> CLng(oRecordset.Fields("TaskID").Value) Then
						sRowContents = GetTaskAdvanceAsRowContents(oRequest, oRecordset, -1, 0, aTaskComponent, sView, lRecordID, sTresholdColor)
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, sTresholdColor, "", "", "", sErrorDescription)
						lCurrentTaskID = CLng(oRecordset.Fields("TaskID").Value)
					End If
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE>"
		Else
			lErrorNumber = L_ERR_NO_RECORDS
		End If
	End If

	Set oRecordset = Nothing
	DisplayProjectAdvanceByAreasOrCategories = lErrorNumber
	Err.Clear
End Function

Function DisplayTaskTabs(oADODBConnection, lProjectID, lParentID, sErrorDescription)
'************************************************************
'Purpose: To display the tabs for the tasks HTML forms
'Inputs:  oRequest, lProjectID, lParentID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayTaskTabs"
	Dim iSelectedTab
	Dim asTitles
	Dim iIndex
	Dim sURL
	Dim lErrorNumber

	iSelectedTab = 1
	asTitles = Split(",Descripción,Información adicional,Formulario,Expediente,Mediateca,Reporte", ",")
	If (Len(aTaskComponent(S_FILE_TASK)) = 0) And (Len(aTaskComponent(S_PROJECT_FILE_TASK)) = 0) Then asTitles(2) = ""
	If aTaskComponent(N_FORM_TASK) = -1 Then asTitles(3) = ""
	If aTaskComponent(N_PUCO_SECTION_ID_TASK) = -1 Then asTitles(5) = ""
	If Len(aTaskComponent(S_REPORT_URL_TASK)) = 0 Then asTitles(6) = ""
	Response.Write "<SCRIPT LANGUAGE=""Javascript""><!--" & vbNewLine
		 Response.Write "var iCurrentDiv = " & iSelectedTab & ";" & vbNewLine
		 Response.Write "function ShowTaskTab(iTab) {" & vbNewLine
			Response.Write "HideDisplay(document.all['Task' + iCurrentDiv + 'Div']);" & vbNewLine
			Response.Write "document.all['Task' + iCurrentDiv + 'LfDiv'].style.backgroundColor = '#CCCCCC';" & vbNewLine
			Response.Write "document.all['Task' + iCurrentDiv + 'CtDiv'].style.backgroundColor = '#CCCCCC';" & vbNewLine
			Response.Write "document.all['Task' + iCurrentDiv + 'RgDiv'].style.backgroundColor = '#CCCCCC';" & vbNewLine
			Response.Write "iCurrentDiv = iTab;" & vbNewLine
			Response.Write "ShowDisplay(document.all['Task' + iCurrentDiv + 'Div']);" & vbNewLine
			Response.Write "document.all['Task' + iCurrentDiv + 'LfDiv'].style.backgroundColor = '#" & S_TAB_BGCOLOR_FOR_GUI & "';" & vbNewLine
			Response.Write "document.all['Task' + iCurrentDiv + 'CtDiv'].style.backgroundColor = '#" & S_TAB_BGCOLOR_FOR_GUI & "';" & vbNewLine
			Response.Write "document.all['Task' + iCurrentDiv + 'RgDiv'].style.backgroundColor = '#" & S_TAB_BGCOLOR_FOR_GUI & "';" & vbNewLine
		 Response.Write "} // End of ShowTaskTab" & vbNewLine
	Response.Write "//--></SCRIPT>" & vbNewLine

	Response.Write "<TABLE BORDER=""0"" WIDTH=""98%"" CELLPADDING=""0"" CELLSPACING=""0""><TR>" & vbNewLine
		For iIndex = 1 To UBound(asTitles)
			If Len(asTitles(iIndex)) > 0 Then
				Response.Write "<TD BGCOLOR=""#"
					If iSelectedTab = iIndex Then
						Response.Write S_TAB_BGCOLOR_FOR_GUI
					Else
						Response.Write "CCCCCC"
					End If
				Response.Write """ WIDTH=""5"" NAME=""Task" & iIndex & "LfDiv"" ID=""Task" & iIndex & "LfDiv""><IMG SRC=""Images/TbLf.gif"" WIDTH=""5"" HEIGHT=""21"" /></TD>"
				Response.Write "<TD BACKGROUND=""Images/TbBg.gif"" BGCOLOR=""#"
					If iSelectedTab = iIndex Then
						Response.Write S_TAB_BGCOLOR_FOR_GUI
					Else
						Response.Write "CCCCCC"
					End If
				Response.Write """ WIDTH=""160"" ALIGN=""CENTER"" NAME=""Task" & iIndex & "CtDiv"" ID=""Task" & iIndex & "CtDiv""><NOBR>"
				Response.Write "<A HREF=""javascript: ShowTaskTab(" & iIndex & ")"" CLASS=""SpecialLink"" STYLE=""width: 100%""><FONT FACE=""Arial"" SIZE=""2"" COLOR=""#"
					If iSelectedTab = iIndex Then
						Response.Write S_TAB_TEXT_FOR_GUI
					Else
						Response.Write "000000"
					End If
				Response.Write """><B>&nbsp;&nbsp;&nbsp;" & asTitles(iIndex) & "&nbsp;&nbsp;&nbsp;</B></A></NOBR></FONT></TD>"
				Response.Write "<TD BGCOLOR=""#"
					If iSelectedTab = iIndex Then
						Response.Write S_TAB_BGCOLOR_FOR_GUI
					Else
						Response.Write "CCCCCC"
					End If
				Response.Write """ WIDTH=""5"" NAME=""Task" & iIndex & "RgDiv"" ID=""Task" & iIndex & "RgDiv""><IMG SRC=""Images/TbRg.gif"" WIDTH=""5"" HEIGHT=""21"" /></TD>" & vbNewLine
			End If
		Next
		Response.Write "<TD BACKGROUND=""Images/TbBgDot.gif"" WIDTH=""*""><IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""21"" /></TD>" & vbNewLine
	Response.Write "</TR></TABLE><BR />" & vbNewLine

	DisplayTaskTabs = lErrorNumber
	Err.Clear
End Function

Function DisplayTaskForms(oADODBConnection, aTaskComponent, sErrorDescription)
'************************************************************
'Purpose: To display the tabs for the tasks HTML forms
'Inputs:  oRequest, lProjectID, lParentID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayTaskForms"

'	Response.Write "<TABLE BORDER=""0"" WIDTH=""100%"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
'		Response.Write "<TD VALIGN=""TOP"">"
'			Response.Write "<DIV ID=""PicTaskPanelDiv"" STYLE=""width: 43px"">"
'				Response.Write "<IMG SRC=""Images/PicTaskPanel.gif"" WIDTH=""43"" HEIGHT=""112"" />"
'			Response.Write "</DIV>"
'
'			Response.Write "<DIV ID=""TaskPanelDiv"" STYLE=""display: none; width: 100%"">"
'				If aTaskComponent(N_ID_TASK) <> -1 Then
'					Call DisplayTaskTabs(oADODBConnection, aTaskComponent(N_PROJECT_ID_TASK), aTaskComponent(N_ID_TASK), "")
'				End If
'
'				Response.Write "<DIV ID=""Task1Div"" STYLE=""width: 98%;"">"
'					If aTaskComponent(N_ID_TASK) = -1 Then
'						lErrorNumber = DisplayProject(oADODBConnection, aTaskComponent(N_PROJECT_ID_TASK), sErrorDescription)
'					Else
'						lErrorNumber = DisplayTask(oRequest, oADODBConnection, True, aTaskComponent, sErrorDescription)
'					End If
'					If lErrorNumber <> 0 Then
'						Call DisplayErrorMessage("Error en los procesos", sErrorDescription)
'						lErrorNumber = 0
'						sErrorDescription = ""
'						Response.Write "<BR />"
'					End If
'				Response.Write "</DIV>"

'				If aTaskComponent(N_ID_TASK) <> -1 Then
'					Response.Write "<DIV ID=""Task2Div"" STYLE=""display: none; width: 98%;"">"
'						If Len(aTaskComponent(S_PROJECT_FILE_TASK)) > 0 Then
'							Response.Write GetFileContents(Server.MapPath(UPLOADED_PHYSICAL_PATH & aTaskComponent(S_PROJECT_FILE_TASK)), sErrorDescription)
'							If Len(sErrorDescription) > 0 Then
'								If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_PROJECTS_PERMISSIONS) = N_PROJECTS_PERMISSIONS Then sErrorDescription = sErrorDescription & "<BR /><BR /><A HREF=""Catalogs.asp?Action=Projects&ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & "&Change=1&Highlight=ProjectFile""><B>Revise la ruta del archivo</B></A> desde la administración de procesos."
'								Call DisplayErrorMessage("Error en el proceso", sErrorDescription)
'								sErrorDescription = ""
'							End If
'						End If
'						If Len(aTaskComponent(S_FILE_TASK)) > 0 Then
'							If Len(aTaskComponent(S_PROJECT_FILE_TASK)) > 0 Then Response.Write "<BR /><HR /><BR />"
'							If InStr(1, aTaskComponent(S_FILE_TASK), "http", vbBinaryCompare) = 0 Then
'								Response.Write GetFileContents(Server.MapPath(UPLOADED_PHYSICAL_PATH & aTaskComponent(S_FILE_TASK)), sErrorDescription)
'							Else
'								Response.Write "<IFRAME SRC=""" & aTaskComponent(S_FILE_TASK) & """ NAME=""TaskFileIFrame"" FRAMEBORDER=""0"" WIDTH=""98%"" HEIGHT=""400""></IFRAME>"
'							End If
'						End If
'					Response.Write "</DIV>"
'
'					Response.Write "<DIV ID=""Task3Div"" STYLE=""display: none; width: 98%;""><FONT FACE=""Arial"" SIZE=""2"">"
'						Response.Write "<FORM NAME=""TaskFrm"" ID=""TaskFrm"" ACTION=""Projects.asp"" METHOD=""POST"" onSumbit="""">"
'							If aTaskComponent(N_FORM_TASK) > -1 Then
'								lErrorNumber = DisplayFormForTask(oADODBConnection, "TaskFrm", True, False, aFormComponent, aTaskComponent, sErrorDescription)
'								If lErrorNumber <> 0 Then
'									Call DisplayErrorMessage("Error en los formularios", sErrorDescription)
'									lErrorNumber = 0
'									sErrorDescription = ""
'									Response.Write "<BR />"
'								End If
'							End If
'						Response.Write "</FORM>"
'					Response.Write "</FONT></DIV>"
'
'					Response.Write "<DIV ID=""Task4Div"" STYLE=""display: none; width: 98%;"">"
'						Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
'							Response.Write "<TR>"
'								Response.Write "<TD WIDTH=""300"" VALIGN=""TOP"">"
'									lErrorNumber = DisplayHistoryListForm(oRequest, oADODBConnection, GetASPFileName(""), aHistoryComponent, sErrorDescription)
'									If lErrorNumber <> 0 Then
'										Response.Write "<BR />"
'										Call DisplayErrorMessage("Error en el hsitorial de comentarios", sErrorDescription)
'										lErrorNumber = 0
'										sErrorDescription = ""
'									End If
'								Response.Write "</TD>"
'								Response.Write "<TD><IMG SRC=""Images/Transparent.gif"" WIDTH=""8"" HEIGHT=""1"" /></TD>"
'								Response.Write "<TD BGCOLOR=""#" & S_MAIN_COLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
'								Response.Write "<TD><IMG SRC=""Images/Transparent.gif"" WIDTH=""8"" HEIGHT=""1"" /></TD>"
'								Response.Write "<TD WIDTH=""300"" VALIGN=""TOP"">"
'									lErrorNumber = DisplayHistoryList(oRequest, oADODBConnection, aHistoryComponent, sErrorDescription)
'									If lErrorNumber <> 0 Then
'										Response.Write "<BR />"
'										Call DisplayErrorMessage("Error en el hsitorial de comentarios", sErrorDescription)
'										lErrorNumber = 0
'										sErrorDescription = ""
'									End If
'								Response.Write "</TD>"
'								Response.Write "<TD><IMG SRC=""Images/Transparent.gif"" WIDTH=""8"" HEIGHT=""1"" /></TD>"
'								Response.Write "<TD BGCOLOR=""#" & S_MAIN_COLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
'								Response.Write "<TD><IMG SRC=""Images/Transparent.gif"" WIDTH=""8"" HEIGHT=""1"" /></TD>"
'								Response.Write "<TD WIDTH=""300"" VALIGN=""TOP"">"
'									Response.Write "<IFRAME SRC=""BrowserFile.asp?ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & "&TaskID=" & aTaskComponent(N_ID_TASK) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & """ NAME=""TaskFilesIFrame"" FRAMEBORDER=""0"" WIDTH=""300"" HEIGHT=""348""></IFRAME>"
'								Response.Write "</TD>"
'							Response.Write "</TR>"
'						Response.Write "</TABLE>"
'					Response.Write "</DIV>"
'
'					Response.Write "<DIV ID=""Task5Div"" STYLE=""display: none; width: 98%;"">"
'						If aTaskComponent(N_PUCO_SECTION_ID_TASK) > -1 Then
'							Response.Write "<IFRAME SRC=""/PuCo/ShowDocuments.asp?ParentID=" & aTaskComponent(N_PUCO_SECTION_ID_TASK) & "&TaskID=" & aTaskComponent(N_ID_TASK) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & """ NAME=""TaskFilesIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""348""></IFRAME>"
'						End If
'					Response.Write "</DIV>"
'
'					Response.Write "<DIV ID=""Task6Div"" STYLE=""display: none; width: 98%;"">"
'						If Len(aTaskComponent(S_REPORT_URL_TASK)) > 0 Then
'							Response.Write "<IFRAME SRC=""" & aTaskComponent(S_REPORT_URL_TASK) & "&FromTaCo=1&Accesskey=vac2&Password=vac2"" NAME=""TaskReportIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""348""></IFRAME>"
'						End If
'					Response.Write "</DIV>"
'				End If
'			Response.Write "</DIV>"
'		Response.Write "</TD>"
'		Response.Write "<TD WIDTH=""5"" VALIGN=""TOP""><IMG SRC=""Images/Transparent.gif"" WIDTH=""5"" HEIGHT=""1"" /></TD>"
''		Response.Write "<TD WIDTH=""11"" VALIGN=""TOP"">"
''			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""60"" /><BR />"
''			Response.Write "<A HREF=""javascript: ShowTaskStatusFrm(false);""><IMG SRC=""Images/ArrExpandRg" & sImageSuffix & ".gif"" WIDTH=""11"" HEIGHT=""40"" ALT="""" BORDER=""0"" NAME=""ExpandArrowImg"" /></A>"
''		Response.Write "</TD>"
'		Response.Write "<TD ID=""ExpandArrowDiv"" BACKGROUND=""Images/BGLnExpand.gif"" WIDTH=""7"" VALIGN=""TOP"">"
'			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""60"" /><BR />"
'			Response.Write "<A HREF=""javascript: ShowTaskStatusFrm(false);""><IMG SRC=""Images/ArrExpandRg" & sImageSuffix & ".gif"" WIDTH=""7"" HEIGHT=""49"" ALT="""" BORDER=""0"" NAME=""ExpandArrowImg"" /></A><BR />"
'			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""60"" />"
'		Response.Write "</TD>"
'		Response.Write "<TD WIDTH=""5"" VALIGN=""TOP""><IMG SRC=""Images/Transparent.gif"" WIDTH=""5"" HEIGHT=""1"" /></TD>"
'		Response.Write "<TD VALIGN=""TOP"">"
'			If lErrorNumber = 0 Then
'				If Not bIsNetscape Then
'					Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
'						Response.Write "var sCurrentDiv = '';" & vbNewLine
'
'						Response.Write "function ShowTaskWidget(sDivID) {" & vbNewLine
'							Response.Write "HideTaskWidget(sCurrentDiv);" & vbNewLine
'
'							Response.Write "var oTaskDiv = document.all['TaskInfo' + sDivID + 'Div'];" & vbNewLine
'							Response.Write "if (oTaskDiv) {" & vbNewLine
'								Response.Write "ShowPopupItem('TaskInfo' + sDivID + 'Div', oTaskDiv);" & vbNewLine
'							Response.Write "}" & vbNewLine
'
'							Response.Write "sCurrentDiv = sDivID;" & vbNewLine
'						Response.Write "} // End of ShowTaskWidget" & vbNewLine
'
'						Response.Write "function HideTaskWidget(sDivID) {" & vbNewLine
'							Response.Write "var oTaskDiv = document.all['TaskInfo' + sDivID + 'Div'];" & vbNewLine
'							Response.Write "if (oTaskDiv)" & vbNewLine
'								Response.Write "HidePopupItem('TaskInfo' + sDivID + 'Div', oTaskDiv);" & vbNewLine
'						Response.Write "} // End of HideTaskWidget" & vbNewLine
'					Response.Write "//--></SCRIPT>" & vbNewLine
'				End If
'
'				Response.Write "<DIV ID=""PicAdvancePanelDiv"" STYLE=""display: none; width: 43px"">"
'					Response.Write "<IMG SRC=""Images/PicAdvancePanel.gif"" WIDTH=""43"" HEIGHT=""112"" />"
'				Response.Write "</DIV>"
'
'				Response.Write "<DIV ID=""AdvancePanelDiv"" CLASS=""AdvanceSection"">"
					If Not bIsNetscape Then
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							Response.Write "function CheckNewTaskValue(oForm) {" & vbNewLine
								Response.Write "if (oForm) {" & vbNewLine
									Response.Write "if (! CheckFloatValue(oForm.TaskStatusPercentage, 'el porcentaje', N_BOTH_FLAG, N_CLOSED_FLAG, 0, 100))" & vbNewLine
										Response.Write "return false;" & vbNewLine
								Response.Write "}" & vbNewLine

								Response.Write "return true;" & vbNewLine
							Response.Write "} // End of CheckNewTaskValue" & vbNewLine
						Response.Write "//--></SCRIPT>" & vbNewLine
								
						Response.Write "<DIV ID=""WidgetInfoDiv"" CLASS=""WidgetInfo""><FORM NAME=""NewTaskStatusFrm"" ID=""NewTaskStatusFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""POST"" onSubmit=""return CheckNewTaskValue(this);"">"
							Response.Write "<TABLE BGCOLOR=""#" & S_LIGHT_BGCOLOR & """ WIDTH=""158"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">"
								Response.Write "<TR>"
									Response.Write "<TD WIDTH=""4""><IMG SRC=""Images/CrnWidgetTpLf.gif"" WIDTH=""4"" HEIGHT=""4"" /></TD>"
									Response.Write "<TD BACKGROUND=""Images/FrmWidgetTp.gif"" WIDTH=""150"" COLSPAN=""2""><IMG SRC=""Images/Transparent.gif"" WIDTH=""150"" HEIGHT=""4"" /></TD>"
									Response.Write "<TD WIDTH=""4""><IMG SRC=""Images/CrnWidgetTpRg.gif"" WIDTH=""4"" HEIGHT=""4"" /></TD>"
								Response.Write "</TR>"
								Response.Write "<TR>"
									Response.Write "<TD BACKGROUND=""Images/FrmWidgetLf.gif""><IMG SRC=""Images/Transparent.gif"" WIDTH=""4"" HEIGHT=""1"" /></TD>"
									Response.Write "<TD VALIGN=""BOTTOM""><FONT FACE=""Verdana"" SIZE=""1""><B>Avance:</B></FONT></TD>"
									Response.Write "<TD ALIGN=""RIGHT"">"
										Response.Write "<A HREF=""javascript: HideWidget('')"">"
											Response.Write "<IMG SRC=""Images/BtnClose.gif"" WIDTH=""11"" HEIGHT=""10"" ALT=""Cancelar"" BORDER=""0"" />"
										Response.Write "</A>"
										Response.Write "<BR /><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""6"" />"
									Response.Write "</TD>"
									Response.Write "<TD BACKGROUND=""Images/FrmWidgetRg.gif""><IMG SRC=""Images/Transparent.gif"" WIDTH=""4"" HEIGHT=""1"" /></TD>"
								Response.Write "</TR>"
								Response.Write "<TR>"
									Response.Write "<TD BACKGROUND=""Images/FrmWidgetLf.gif""><IMG SRC=""Images/Transparent.gif"" WIDTH=""4"" HEIGHT=""1"" /></TD>"
									Response.Write "<TD><FONT FACE=""Verdana"" SIZE=""1"">&nbsp;&nbsp;&nbsp;"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""View"" ID=""ViewHdn"" VALUE=""" & oRequest("View").Item & """ />"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProjectID"" ID=""ProjectIDHdn"" VALUE=""" & aTaskComponent(N_PROJECT_ID_TASK) & """ />"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaskID"" ID=""TaskIDHdn"" VALUE="""" />"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ParentID"" ID=""ParentIDHdn"" VALUE=""" & aTaskComponent(N_ID_TASK) & """ />"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaskPath"" ID=""TaskPathHdn"" VALUE="""" />"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ValueForAllParents"" ID=""ValueForAllParentsHdn"" VALUE=""1"" />"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FromParent"" ID=""FromParentHdn"" VALUE=""1"" />"
										Response.Write "<INPUT TYPE=""TEXT"" NAME=""TaskStatusPercentage"" ID=""TaskStatusPercentageTxt"" VALUE="""" SIZE=""6"" MAXLENGTH=""6"" CLASS=""TextFields"" />"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaID"" ID=""AreaIDHdn"" VALUE=""" & oRequest("AreaID").Item & """ />"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CategoryID"" ID=""CategoryIDHdn"" VALUE=""" & oRequest("CategoryID").Item & """ />"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""UserID"" ID=""UserIDHdn"" VALUE=""" & oRequest("UserID").Item & """ />"
									Response.Write "%</FONT></TD>"
									Response.Write "<TD ALIGN=""RIGHT""><INPUT TYPE=""SUBMIT"" NAME=""UpdateTaskStatus"" ID=""UpdateTaskStatusBtn"" VALUE=""  Ok  "" CLASS=""Buttons"" /></TD>"
									Response.Write "<TD BACKGROUND=""Images/FrmWidgetRg.gif""><IMG SRC=""Images/Transparent.gif"" WIDTH=""4"" HEIGHT=""1"" /></TD>"
								Response.Write "</TR>"
								Response.Write "<TR>"
									Response.Write "<TD><IMG SRC=""Images/CrnWidgetBtLf.gif"" WIDTH=""4"" HEIGHT=""4"" /></TD>"
									Response.Write "<TD BACKGROUND=""Images/FrmWidgetBt.gif"" COLSPAN=""2""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""4"" /></TD>"
									Response.Write "<TD><IMG SRC=""Images/CrnWidgetBtRg.gif"" WIDTH=""4"" HEIGHT=""4"" /></TD>"
								Response.Write "</TR>"
							Response.Write "</TABLE>"
						Response.Write "</DIV>"
					End If
					If (aTaskComponent(N_ID_TASK) = -1) And (Len(sView) > 0) Then
						Select Case sView
							Case "Areas"
								If Len(oRequest("AreaID").Item) > 0 Then lRecordID = CLng(oRequest("AreaID").Item)
								If lRecordID = -1 Then
									lErrorNumber = DisplayAreasOrCategoriesForProject(oADODBConnection, aTaskComponent(N_PROJECT_ID_TASK), sView, sErrorDescription)
								Else
									lErrorNumber = DisplayProjectAdvanceByAreasOrCategories(oADODBConnection, aTaskComponent(N_PROJECT_ID_TASK), aTaskComponent(N_ID_TASK), lRecordID, sView, sErrorDescription)
								End If
							Case "Categories"
								If Len(oRequest("CategoryID").Item) > 0 Then lRecordID = CLng(oRequest("CategoryID").Item)
								If lRecordID = -1 Then
									lErrorNumber = DisplayAreasOrCategoriesForProject(oADODBConnection, aTaskComponent(N_PROJECT_ID_TASK), sView, sErrorDescription)
								Else
									lErrorNumber = DisplayProjectAdvanceByAreasOrCategories(oADODBConnection, aTaskComponent(N_PROJECT_ID_TASK), aTaskComponent(N_ID_TASK), lRecordID, sView, sErrorDescription)
								End If
							Case "Users"
								If Len(oRequest("UserID").Item) > 0 Then lRecordID = CLng(oRequest("UserID").Item)
								If lRecordID = -1 Then
									lErrorNumber = DisplayAreasOrCategoriesForProject(oADODBConnection, aTaskComponent(N_PROJECT_ID_TASK), sView, sErrorDescription)
								Else
									lErrorNumber = DisplayProjectAdvanceByAreasOrCategories(oADODBConnection, aTaskComponent(N_PROJECT_ID_TASK), aTaskComponent(N_ID_TASK), lRecordID, sView, sErrorDescription)
								End If
						End Select
					Else
						If CInt(GetOption(aOptionsComponent, FULL_PROJECT_OPTION)) = 0 Then
							lErrorNumber = DisplayProjectAdvance(oADODBConnection, aTaskComponent(N_PROJECT_ID_TASK), aTaskComponent(N_ID_TASK), sErrorDescription)
							If lErrorNumber = L_ERR_NO_RECORDS Then
								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
									Response.Write "ShowTaskStatusFrm(false);" & vbNewLine
									Response.Write "ToogleDiv('TaskParameters');" & vbNewLine
									Response.Write "HideDisplay(document.all['ExpandArrowDiv']);" & vbNewLine
									Response.Write "HideDisplay(document.all['PicAdvancePanelDiv']);" & vbNewLine
								Response.Write "//--></SCRIPT>" & vbNewLine
								lErrorNumber = 0
							End If
						Else
							lErrorNumber = DisplayFullProjectAdvance(oADODBConnection, aTaskComponent(N_PROJECT_ID_TASK), sErrorDescription)
						End If
					End If
					If lErrorNumber <> 0 Then
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						sErrorDescription = ""
						Response.Write "<BR />"
					End If
'				Response.Write "</DIV>"
'			End If
'		Response.Write "</TD>"
'	Response.Write "</TR></TABLE>"
'	If StrComp(oRequest("Action").Item, "HistoryList", vbBinaryCompare) = 0 Then
'		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
'			Response.Write "if (IsDisplayed(document.all['ExpandArrowDiv']))" & vbNewLine
'				Response.Write "ShowTaskStatusFrm(false);" & vbNewLine
'			Response.Write "ShowTaskTab(4);" & vbNewLine
'		Response.Write "//--></SCRIPT>" & vbNewLine
'	End If

	DisplayTaskForms = lErrorNumber
	Err.Clear
End Function

Function TaskHasChildren(oADODBConnection, lProjectID, lTaskID, sErrorDescription)
'************************************************************
'Purpose: To check if the given task has children
'Inputs:  oRequest, lProjectID, lTaskID
'Outputs: sErrorDescription. A boolean
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "TaskHasChildren"
	Dim oRecordset
	Dim lErrorNumber

	TaskHasChildren = False
	sErrorDescription = "No se pudo revisar si la actividad especificada está conformada por más actividades."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select TaskID From " & TACO_PREFIX & "TasksLKP Where (ProjectID=" & lProjectID & ") And (ParentID=" & lTaskID & ")", "TaCoProjectsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		TaskHasChildren = (Not oRecordset.EOF)
	End If

	Err.Clear
End Function
%>