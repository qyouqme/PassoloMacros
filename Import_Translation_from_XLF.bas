'#Reference {F5078F18-C551-11D3-89B9-0000F81FE221}#6.0#0#C:\Windows\System32\msxml6.dll#Microsoft XML, v6.0
'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\system32\scrrun.dll#Microsoft Scripting Runtime
''Import translation of updated stringlist from translated XLF file.
Sub Main
	Dim inPath As String
	Dim proj As PslProject
	Dim tranlst As PslTransList
	Dim inFolder As Folder
	Dim filesDict As Object
	Dim fso As New FileSystemObject

	Begin Dialog UserDialog 930,140,"Import Translation from XLF",.BrowseFolderPath ' %GRID:10,7,1,1
		Text 20,14,90,21,"Path Input:",.lab_inputPath
		TextBox 20,42,780,21,.tb_inputPath
		OKButton 230,84,150,35,.btn_OK
		CancelButton 580,84,150,35,.btn_cancel
		PushButton 820,35,90,28,"Browse...",.puBtn_browse
	End Dialog
	Dim dlg As UserDialog
	If Dialog(dlg)=0 Then Exit Sub

	inPath = dlg.tb_inputPath
	If Right(inPath,1)<>"\" Then inPath=inPath + "\"

	Set proj = PSL.ActiveProject
	Set filesDict = CreateObject("Scripting.Dictionary")

	PSL.Output "********** Start - Import Translation from XLF **********"
	Set inFolder = fso.GetFolder(inPath)
	getAllFiles inFolder, filesDict
	'PSL.Output CStr(filesDict.Count)

	For Each tranlst In proj.TransLists
		If tranlst.Selected = True Then
			Dim filename As String
			filename = CStr(tranlst.ListID) + "_" + fso.GetFileName(tranlst.SourceList.SourceFile) + ".xlf"
			If filesDict.Exists(LCase(filename)) Then
				Dim filepath As String
				Dim xlfDoc As New DOMDocument60
				Dim fileNd As IXMLDOMElement
				Dim origfile, sourceLang, targetLang As String

				filepath = filesDict.Item(LCase(filename))
				xlfDoc.async = False
				xlfDoc.validateOnParse = False
				xlfDoc.preserveWhiteSpace = True
				xlfDoc.resolveExternals = False
				xlfDoc.load(filepath)
				xlfDoc.setProperty("SelectionLanguage", "XPath")
				xlfDoc.setProperty("SelectionNamespaces", "xmlns:xlfns='" + xlfDoc.documentElement.namespaceURI + "'")

				Set fileNd = xlfDoc.selectSingleNode("//xlfns:file")
				origfile = fileNd.getAttribute("original")
				sourceLang = fileNd.getAttribute("source-language")
				targetLang = fileNd.getAttribute("target-language")
				If origfile = fso.GetFileName(tranlst.SourceList.SourceFile) And sourceLang = PSL.GetLangCode(tranlst.SourceList.LangID, 8) And targetLang = PSL.GetLangCode(tranlst.Language.LangID, 8) Then
					PSL.Output "Import Translation >> " + PSL.GetLangCode(tranlst.Language.LangID, 8) + " - " + fso.GetFileName(tranlst.SourceList.SourceFile)
					ImportTranslation tranlst, fileNd
				End If
			End If
		End If
	Next

	PSL.Output "********** End - Import Translation from XLF **********"
	MsgBox "DONE!", vbOkOnly, "Import Translation from XLF"
End Sub
Private Function BrowseFolderPath(DlgItem$, Action%, SuppValue&) _
As Boolean
	Dim Folder As String
	' We only want to check button clicks
	If Action% <> 2 Then Exit Function


	' Let user select TM file
	If DlgItem$ = "puBtn_browse" Then
	    Folder=PSL.ActiveProject.Location
	    If PSL.SelectFolder(Folder,"Select folder") Then
			DlgText "tb_inputPath", Folder
		End If
		BrowseFolderPath = True
	End If
End Function
Sub getAllFiles(inFolder As Folder, filesDict As Object)
	Dim fso As New FileSystemObject
	Dim xlfFile As File
	Dim subFold As Folder
	For Each xlfFile In inFolder.Files
		If LCase(fso.GetExtensionName (xlfFile.Name))="xlf" Then
		'PSL.Output xlfFile.Name
		'PSL.Output xlfFile.Path
		filesDict.Add(LCase(xlfFile.Name), xlfFile.Path)
		End If
	Next

	For Each subFold In inFolder.SubFolders
		getAllFiles subFold, filesDict
	Next
End Sub

Sub ImportTranslation(tranlst As PslTransList, fileNd As IXMLDOMElement)
	Dim tuNd As IXMLDOMElement
	Dim tuNds As Variant

	Set tuNds = fileNd.selectNodes(".//xlfns:trans-unit")
	For Each tuNd In tuNds
		Dim strNum As String
		Dim transtr As PslTransString

		strNum = tuNd.getAttribute("id")
		Set transtr = tranlst.String(CLng(strNum), 3)
		If transtr.State(pslStateReadOnly) = False And transtr.State(pslStateLocked) = False And transtr.State(pslStateHidden) = False And transtr.State(pslStateTranslated) = False And transtr.State(pslStateReview) = False Then
			Dim sourceNd, targetNd, noteNd As IXMLDOMElement
			Dim sourceStr, targetStr, noteStr As String

			Set sourceNd = tuNd.selectSingleNode("./xlfns:source")
			sourceStr = sourceNd.Text

			Set targetNd = tuNd.selectSingleNode("./xlfns:target")
			targetStr = targetNd.Text

			Set noteNd = tuNd.selectSingleNode("./xlfns:note")
			noteStr = noteNd.Text

			If sourceStr= transtr.SourceText And noteStr = transtr.ID Then
				transtr.Text = targetStr
				transtr.State(pslStateReview) = True
			End If
		End If
	Next

	tranlst.Save
End Sub
