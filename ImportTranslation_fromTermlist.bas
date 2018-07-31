'#Reference {2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.5#0#C:\Program Files\Common Files\Microsoft Shared\OFFICE14\MSO.DLL#Microsoft Office 14.0 Object Library
'#Reference {00020813-0000-0000-C000-000000000046}#1.7#0#D:\ToolSoftware\Office2010\Office14\EXCEL.EXE#Microsoft Excel 14.0 Object Library
'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\system32\scrrun.dll#Microsoft Scripting Runtime
Dim prj As PslProject
Dim langlst As PslLanguages
Dim trnlsts As PslTransLists
Dim int_totalTransList As Integer
Dim regxP As Object

Sub Main
	Dim filePath As String
	Dim srcDict As Object
	Dim fo_path As Folder
	Dim fso As New FileSystemObject

	Begin Dialog UserDialog 860,140,"Import Translation from Termlist",.TMAutotrans ' %GRID:10,7,1,1
		Text 20,14,90,21,"Enter Path:",.lab_path
		TextBox 20,42,750,21,.tb_path
		OKButton 240,91,90,28,.btn_ok
		CancelButton 530,91,90,28,.btn_cancel
		PushButton 790,42,60,21,"...",.puBtn_browse
	End Dialog
	Dim dlg As UserDialog
	If Dialog(dlg)=0 Then Exit Sub

    filePath = dlg.tb_path

    If Right(filePath,1)<>"\" Then filePath=filePath+"\"

    PSL.Output "Start - Import Translations from Termlist"

    Set prj = PSL.ActiveProject
    Set langlst = PSL.ActiveProject.Languages
    Set trnlsts = prj.TransLists
    int_totalTransList = trnlsts.Count

    'MsgBox CStr(int_totalTransList)
    
    Set regxP = CreateObject("vbscript.regexp")
    regxP.Pattern = "^[a-z]{3}$"
    regxP.Global = True

    Set srcDict = CreateObject("Scripting.Dictionary")
    Set fo_path = fso.GetFolder(filePath)
    processfolder fo_path, srcDict

    PSL.Output "End - Import Translations from Termlist"
    MsgBox "DONE!"
End Sub

Sub processfolder(fold As Folder, obj As Object)
	Dim fi As File
	Dim fisysobj As New FileSystemObject
	Dim subfold As Folder

	For Each fi In fold.Files
	 	If LCase(fisysobj.GetExtensionName (fi.Name)) = "xlsx" Then
	 		InputTrans fi.Path, obj
	 	End If
	Next

	For Each subfold In fold.SubFolders
		processfolder subfold, obj
	Next
End Sub

Sub InputTrans(fPath As String, obj As Object)
	Dim excApp As Excel.Application
	Dim wb_cur As Excel.Workbook
	Dim ws_cur As Worksheet
	Dim str_LCode As String
	Dim int_maxrow As Integer
	Dim trnlst As PslTransList
	Dim trnstr As PslTransString
	Dim int_transStr As Integer
	Dim var_findID As Variant
	Dim int_FindIdx As Long

	PSL.Output "Get translations from: " + fPath
	Set excApp = CreateObject("Excel.Application")
	Set wb_cur = excApp.Workbooks.Open(fPath)
	For Each ws_cur In wb_cur.Sheets
		str_LCode = LCase(ws_cur.Name)
		If regxP.Test(str_LCode) = True Then
			For Eachlst = 1 To int_totalTransList Step 1
				Set trnlst = trnlsts(Eachlst)
				If trnlst.Language.LangCode = str_LCode And trnlst.Selected = True And trnlst.ExportFile = "" Then
					With ws_cur
						int_maxrow = .Range("A65535").End(xlUp).Row
						For idx_title = 2 To int_maxrow Step 1
							If .Range("A" + CStr(idx_title)) = trnlst.Title Then
								var_findID = .Range("D" + CStr(idx_title))
								int_FindIdx = 1
								While True
									Set trnstr = trnlst.FindID(var_findID, int_FindIdx)
									If trnstr Is Nothing Then Exit While

									If trnstr.Number = .Range("C" + CStr(idx_title)) And trnstr.SourceText = .Range("E" + CStr(idx_title)) And trnstr.State(pslStateReadOnly) = False And trnstr.State(pslStateReview) = False And trnstr.State(pslStateTranslated) = False Then
										trnstr.Text = .Range("F" + CStr(idx_title))
										trnstr.State(pslStateReview) = True
									End If

									int_FindIdx = int_FindIdx + 1
								Wend
							End If
						Next idx_title
					End With
				End If
				trnlst.Save
			Next Eachlst
		End If
	Next
	wb_cur.Close
	excApp.Quit
End Sub



Private Function TMAutotrans(DlgItem$, Action%, SuppValue&) _
As Boolean
	Dim Folder As String
	' We only wnat to check button clicks
	If Action% <> 2 Then Exit Function


	' Let user select TM file
	If DlgItem$ = "puBtn_browse" Then
	    Folder=PSL.ActiveProject.Location
	    If PSL.SelectFolder(Folder,"Select folder") Then
			DlgText "tb_path", Folder
		End If
		TMAutotrans = True
	End If
End Function
