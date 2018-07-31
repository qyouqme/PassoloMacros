''Export the update [en->target] stringlist for each sync Loop
Sub Main
	start_time = Timer()

'defind related var stirngs
Dim prj As PslProject
Dim trnlst As PslTransList
Dim srclst As PslSourceList
Dim srcTxt As String
Dim tagTxt As String
Dim  i, j, n, m As Long


Set prj = PSL.ActiveProject

For i = 1 To prj.TransLists.Count		'翻译列表的个数
Set trnlst = prj.TransLists(i)		'取得某个翻译列表

If  trnlst.Language.LangCode = "jpn" Then

		For j = 1 To trnlst.StringCount 				'遍历字串列表



			If trnlst.String(j).State(pslStateTranslated) Then

			tagTxt = trnlst.String(j).Text
			srcTxt = trnlst.String(j).SourceText

				If tagTxt = srcTxt Or tagTxt = "&quot;" + srcTxt + "&quot;" Then


					For n = 1 To prj.SourceLists.Count
					Set srclst = prj.SourceLists(n)

					For m = 1 To srclst.StringCount
						If srclst.String(m).Text = srcTxt Then
						srclst.String(m).State(pslStateReadOnly) = True
						srclst.Save
						End If
					Next m
					Next n


				End If


			End If

	Next j

End If

Next i


MsgBox "Done!"

End Sub
