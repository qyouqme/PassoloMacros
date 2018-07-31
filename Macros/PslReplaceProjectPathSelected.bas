''Search and replaces partial strings in source and target file names

'THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND
'EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES
'OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE

Option Explicit
Sub Main
	' Get PASSOLO project
    Dim prj As PslProject
    Set prj = PSL.ActiveProject

    ' Do we have an open project?
    If prj Is Nothing Then
		MsgBox("No active PASSOLO project.")
		Exit Sub
	End If

	' Define dialog for Search and Replace
	Begin Dialog UserDialog 530,105,"Change project path" ' %GRID:10,7,1,1
		Text 20,20,110,14,"Search for:",.Text1
		Text 20,50,110,14,"Replace with:",.Text2
		TextBox 140,14,370,21,.Search
		TextBox 140,44,370,21,.Replace
		OKButton 320,77,90,21
		CancelButton 420,77,90,21
	End Dialog

	' Show dialog for Search and Replace
	Dim dlg As UserDialog
	Dim rts As Integer
	rts = Dialog(dlg)

	' Stop macro if Cancel pressed
	If rts = 0 Then
		Exit Sub
	End If

	' PASSOLO source list
	Dim src As PslSourceList

	' Browse thru all source lists in the project
	For Each src In prj.SourceLists
		If src.Selected Then
			' Find search string in source path
			If InStr(src.SourceFile, dlg.Search) > 0 Then
				' Replace search string with replace string
				src.SourceFile = Replace( src.SourceFile, dlg.Search, dlg.Replace, 1, 1)
			End If
		End If
	Next src

	' PASSOLO translation list
	Dim trn As PslTransList

	' Browse thru all source lists in the project
	For Each trn In prj.TransLists
		If trn.Selected Then
			If InStr(trn.TargetFile, dlg.Search) > 0 Then
				' Replace search string with replace string
				trn.TargetFile = Replace( trn.TargetFile, dlg.Search, dlg.Replace, 1, 1)
			End If
		End If
	Next trn
End Sub
