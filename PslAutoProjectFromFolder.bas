'' This macro adds all executables found in a
'' directory to a new project

'THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND
'EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES
'OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE


Sub Main
	Begin Dialog UserDialog 540,161,"Bulk Project", _
        .BulkProject ' %GRID:10,7,1,1
		Text 10,14,130,14,"Name of the Project:",.Text1
		Text 10,42,200,14,"Directory for Project:",.Text2
		Text 10,70,170,14,"Directory with files:",.Text3
		TextBox 220,7,270,21,.Project
		TextBox 220,35,270,21,.PSLDir
		TextBox 220,63,270,21,.SourceDir
		PushButton 500,35,30,21,"...",.GetPSLDir
		PushButton 500,63,30,21,"...",.GetSourceDir
		OKButton 430,105,90,21
		CancelButton 430,133,90,21
	End Dialog
	Dim dlg As UserDialog
	If Dialog(dlg) = 0 Then Exit Sub

	' Create PASSOLO project
	Dim prj As PslProject
	Set prj = PSL.Projects.Add(dlg.Project, dlg.PSLDir)

	If prj Is Nothing Then
		MsgBox ("Failed to add project")
		Exit Sub
	End If

	' Go to source directory and scan all exes
	ChDir(dlg.SourceDir)
	Dim file As String
	file = Dir$("*.exe")
	While file <> ""
		prj.SourceLists.Add(dlg.SourceDir & "\" & file, _
            file, pslLangEnglishUSA)
		file = Dir$()
	Wend

	' same for DLLs
	file = Dir$("*.dll")
	While file <> ""
		prj.SourceLists.Add(dlg.SourceDir & "\" & file, _
            file, pslLangEnglishUSA)
		file = Dir$()
	Wend
End Sub

Private Function BulkProject(DlgItem$, Action%, SuppValue&) _
As Boolean                                                  
	Dim folder As String

	' We only wnat to check button clicks
	If Action% <> 2 Then Exit Function

	' Let user select project directory
	If DlgItem$ = "GetPSLDir" Then
		If PSL.SelectFolder(folder, _
            "Select Directory for Project") Then
			DlgText "PSLDir", folder
		End If
		BulkProject = True
	End If

	' Let user select source directory
	If DlgItem$ = "GetSourceDir" Then
		If PSL.SelectFolder(folder, _
            "Select Directory of Files") Then
			DlgText "SourceDir", folder
		End If
		BulkProject = True
	End If

	' Only OK if user has enter all data
	If DlgItem$ = "OK" Then
		If DlgText("SourceDir") = "" Then
			MsgBox("Please select directory of files")
			BulkProject = True
			Exit Function
		End If
		If DlgText("PSLDir") = "" Then
			MsgBox("Please select directory for project")
			BulkProject = True
			Exit Function
		End If
		If DlgText("SourceDir") = "" Then
			MsgBox("Please enter project name")
			BulkProject = True
			Exit Function
		End If
	End If
End Function
