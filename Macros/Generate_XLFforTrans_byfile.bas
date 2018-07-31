'#Reference {F5078F18-C551-11D3-89B9-0000F81FE221}#6.0#0#C:\Windows\System32\msxml6.dll#Microsoft XML, v6.0
'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\system32\scrrun.dll#Microsoft Scripting Runtime
''Export the updated stringlist as XLF file to translate.
Dim fso As New FileSystemObject
Sub Main
	Dim outPath As String
	Dim proj As PslProject
	Dim tranlst As PslTransList

	Begin Dialog UserDialog 930,140,"Generate XLF to Translate",.BrowseFolderPath ' %GRID:10,7,1,1
		Text 20,14,90,21,"Path Output:",.lab_outputPath
		TextBox 20,42,780,21,.tb_outputPath
		OKButton 230,84,150,35,.btn_OK
		CancelButton 580,84,150,35,.btn_cancel
		PushButton 820,35,90,28,"Browse...",.puBtn_browse
	End Dialog
	Dim dlg As UserDialog
	If Dialog(dlg)=0 Then Exit Sub

	outPath = dlg.tb_outputPath
	If Right(outPath,1)<>"\" Then outPath=outPath + "\"

	PSL.Output "********** Start - Generate XLF to Translate **********"
	Set proj = PSL.ActiveProject

	For Each tranlst In proj.TransLists
		If tranlst.Selected = True Then
			Dim outFile As String
			outFile = outPath + PSL.GetLangCode(tranlst.Language.LangID, 8) + "\" + CStr(tranlst.ListID) + "_" + fso.GetFileName(tranlst.SourceList.SourceFile) + ".xlf"
			If fso.FileExists(outFile) = True Then fso.DeleteFile(outFile)

			PSL.Output PSL.GetLangCode(tranlst.Language.LangID, 8) + " - " + fso.GetFileName(tranlst.SourceList.SourceFile)
			GenerateXLF tranlst, outFile
		End If
	Next
	PSL.Output "********** End - Generate XLF to Translate **********"
	MsgBox "DONE!", vbOkOnly, "Generate XLF to Translate"
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
			DlgText "tb_outputPath", Folder
		End If
		BrowseFolderPath = True
	End If
End Function
Sub GenerateXLF(tslst As PslTransList,pathOut As String)
	Dim xlfDoc As New DOMDocument60
	Dim rootNd, fileNd, bodyNd As IXMLDOMElement
	Dim rootNdSpace, fileNdSpace, bodyNdSpace, tuNdSpace, tuChildNdSpace As IXMLDOMNode
	Dim strcount As Integer

	xlfDoc.async = False
	xlfDoc.validateOnParse = False
	xlfDoc.preserveWhiteSpace = True
	xlfDoc.resolveExternals = False
	xlfDoc.loadXML("<?xml version=""1.0"" encoding=""UTF-8""?><xliff xmlns=""urn:oasis:names:tc:xliff:document:1.1"" version=""1.1""></xliff>")

	Set rootNdSpace = xlfDoc.createTextNode(vbLf)
	Set fileNdSpace = xlfDoc.createTextNode(vbLf + "  ")
	Set bodyNdSpace = xlfDoc.createTextNode(vbLf + "    ")
	Set tuNdSpace = xlfDoc.createTextNode(vbLf + "      ")
	Set tuChildNdSpace = xlfDoc.createTextNode(vbLf + "        ")

	Set rootNd = xlfDoc.documentElement
	rootNd.appendChild(fileNdSpace.cloneNode(True))

	Set fileNd = xlfDoc.createNode(1, "file", xlfDoc.documentElement.namespaceURI)
	fileNd.setAttribute("original", fso.GetFileName(tslst.SourceList.SourceFile))
	fileNd.setAttribute("source-language", PSL.GetLangCode(tslst.SourceList.LangID, 8))
	fileNd.setAttribute("datatype", "plaintext")
	fileNd.setAttribute("target-language", PSL.GetLangCode(tslst.Language.LangID, 8))
	fileNd.appendChild(bodyNdSpace.cloneNode(True))

	Set bodyNd = xlfDoc.createNode(1, "body", xlfDoc.documentElement.namespaceURI)

	strcount = 0
	For idxstr = 1 To tslst.StringCount Step 1
		Dim transtr As PslTransString
		Set transtr = tslst.String(idxstr)
		If transtr.State(pslStateReadOnly) = False And transtr.State(pslStateLocked) = False And transtr.State(pslStateHidden) = False And transtr.State(pslStateTranslated) = False And transtr.State(pslStateReview) = False Then
			Dim tuNd, sourceNd, targetNd, noteNd As IXMLDOMElement

			Set tuNd = xlfDoc.createNode(1, "trans-unit", xlfDoc.documentElement.namespaceURI)
			tuNd.setAttribute("xml:space", "preserve")
			tuNd.setAttribute("id", CStr(transtr.Number))
			tuNd.setAttribute("approved", "yes")

			Set sourceNd = xlfDoc.createNode(1, "source", xlfDoc.documentElement.namespaceURI)
			sourceNd.Text = transtr.SourceText

			Set targetNd = xlfDoc.createNode(1, "target", xlfDoc.documentElement.namespaceURI)
			targetNd.setAttribute("state", "in translation")

			Set noteNd = xlfDoc.createNode(1, "note", xlfDoc.documentElement.namespaceURI)
			noteNd.setAttribute("from", "Context")
			noteNd.Text = transtr.ID

			tuNd.appendChild(tuChildNdSpace.cloneNode(True))
			tuNd.appendChild(sourceNd)
			tuNd.appendChild(tuChildNdSpace.cloneNode(True))
			tuNd.appendChild(targetNd)
			tuNd.appendChild(tuChildNdSpace.cloneNode(True))
			tuNd.appendChild(noteNd)
			tuNd.appendChild(tuNdSpace.cloneNode(True))

			bodyNd.appendChild(tuNdSpace.cloneNode(True))
			bodyNd.appendChild(tuNd)
			strcount = strcount + 1
		End If
	Next idxstr

	If strcount > 0 Then
		bodyNd.appendChild(bodyNdSpace.cloneNode(True))
		fileNd.appendChild(bodyNd)
		fileNd.appendChild(fileNdSpace.cloneNode(True))
		rootNd.appendChild(fileNd)
		rootNd.appendChild(rootNdSpace.cloneNode(True))

		If fso.FolderExists(fso.GetParentFolderName(pathOut)) = False Then fso.CreateFolder(fso.GetParentFolderName(pathOut))
		xlfDoc.Save(pathOut)
	End If
End Sub
