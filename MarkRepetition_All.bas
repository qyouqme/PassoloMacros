Sub Main
	Dim prj As PslProject
	Dim langlst As PslLanguages
	Dim trnlsts As PslTransLists
	Dim TotTrnlsts, StrTot, WTmpStr As Long
	Dim srcDict, dictRepStr, regxP As Object
	Dim trnlst As PslTransList
	Dim fname, szLang, szStrTmp, strTransTxt As String
	Dim trnstr As PslTransString
	Dim StrTmpCnt As PslStringCounter

	PSL.Output "Start - Mark Repetition"

	Set prj = PSL.ActiveProject
	Set langlst = PSL.ActiveProject.Languages
	Set trnlsts = prj.TransLists
	Set regxP = CreateObject("vbscript.regexp")

	TotTrnlsts = prj.TransLists.Count
	regxP.Pattern = "[a-zA-Z\d_]+\.h"
	regxP.Global = True

	For Each lang In langlst
		Set srcDict = CreateObject("Scripting.Dictionary")
		Set dictRepStr = CreateObject("Scripting.Dictionary")

		For CountTrnlsts = 1 To TotTrnlsts
			Set trnlst = trnlsts(CountTrnlsts)
			fname=LCase(trnlst.TargetFile)

			If trnlst.Selected =True And trnlst.ExportFile ="" Then
				szLang = trnlst.Language.LangCode
				If szLang = lang.LangCode Then
					StrTot = trnlst.StringCount
					For EachStr = 1 To StrTot
						Set trnstr = trnlst.String(EachStr)
						WTmpStr = 0
						szStrTmp = trnstr.SourceText
						Set StrTmpCnt = PSL.GetTextCounts(szStrTmp)
						WTmpStr = StrTmpCnt.WordCount
						Dim b_repet As Boolean
						b_hFile = regxP.Test(szStrTmp)
						If WTmpStr > 0 And b_hFile = False Then
							If Not srcDict.Exists(szStrTmp) Then
								Dim data(2) As Variant
								PSL.Output szStrTmp
								data(0) = szStrTmp
								data(1) = "repit"
								data(2) = "repit"
								srcDict.Add(szStrTmp, data)
							Else
								If dictRepStr.Exists(szStrTmp) Then
									dictRepStr(szStrTmp) = dictRepStr(szStrTmp) + 1
								Else
									dictRepStr(szStrTmp) = 1
								End If

								strTransTxt = trnstr.Text
								strTransTxt = "$" + dictRepStr.Item(szStrTmp) + strTransTxt
								trnstr.Text = strTransTxt
							End If
						End If
					Next EachStr
					trnlst.Save
				End If
			End If
		Next CountTrnlsts
		Set srcDict=Nothing
		Set dictRepStr = Nothing
	Next lang
	PSL.Output "End - Mark Repetition"
	MsgBox "DONE!"
End Sub
