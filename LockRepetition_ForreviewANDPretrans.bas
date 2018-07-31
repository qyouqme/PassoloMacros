''- Autodesk customized statistics
''- Strings/Words count for Translation Lists
''- Version 6.22


Option Explicit

Dim TotToTranS, TotToTranW, TotUpdToTranS, TotUpdToTranW, TotToReviewS, TotToReviewW, TotRepS, TotRepW As Long
Dim outputBuffer As String

Function addToBuffer(buffer As String) As String
	outputBuffer = outputBuffer + buffer + "~"
End Function

' Add or update an item in the dictionary with the corresponding counters
Sub AddUpdDictItem(dict As Object, key As String, WRepCount As Long, RepNew As Long, RepUpd As Long)

	' if the item does not exist in the dictionary already
	If Not dict.Exists(key) Then
		' create new item
		Dim data(2) As Variant
		data(0) = WRepCount
		data(1) = RepNew
		data(2) = RepUpd
		' add new item to the dictionary
		dict.Add(key, data)
	Else
		'get existing item
		Dim stringIdent() As Variant
		stringIdent = dict.Item(key)
		'prepare updated item
		If RepNew = 1 Then
			stringIdent(1) = stringIdent(1) + 1
		ElseIf RepUpd = 1 Then
			stringIdent(2) = stringIdent(2) + 1
		End If
		' update item in the dictionary
		dict.Item(key) =  stringIdent
	End If

End Sub

' Repetition counter
Sub RepCounter(RepDict As Object)

	Dim item As Variant

	For Each item In RepDict
		Dim RepString() As Variant
		RepString = RepDict(item)
		If ( RepString(1) + RepString(2) ) > 1 Then
			'DEBUG: Print #1, item & " " & RepString(0) & " " & RepString(1) & " " & RepString(2)
			If RepString(1) > 0 Then
				RepString(1) = RepString(1) - 1
			Else
				RepString(2) = RepString(2) - 1
			End If
   			TotToTranS = TotToTranS - RepString(1)
			TotToTranW = TotToTranW - ( RepString(1) * RepString(0) )
    		TotUpdToTranS = TotUpdToTranS - RepString(2)
   			TotUpdToTranW = TotUpdToTranW - ( RepString(2) * RepString(0) )
			TotRepS = TotRepS + RepString(1) + RepString(2)
			TotRepW = TotRepW + ( RepString(1) * RepString(0) ) + ( RepString(2) * RepString(0) )
		End If
	Next

End Sub

Sub Main

	Dim prj As PslProject
	Dim trnlst As PslTransList
	Dim trnlsts As PslTransLists
	Dim szLang, szFullName, szFileName, szStrTmp, szPrj As String
	Dim ToTran, WToTran, UpdToTran, WUpdToTran, RepToTran, WRepToTran, ToReview, WToReview As Long
	Dim TotTrnlsts, CountTrnlsts, pos, posfromend, stringlen, PrintIt, StrTot, EachStr, WTmpStr As Long
	Dim langlst As PslLanguages
	Dim lang As PslLanguage
	Dim trnstr As PslTransString
	Dim StrTmpCnt As PslStringCounter
	Dim srcDict As Object
	Dim fname As String




	PSL.Output "Start - Lock Repition"

	'Print #1, "STATISTICS - version 6.22"

	Set prj = PSL.ActiveProject
	Set langlst = PSL.ActiveProject.Languages
	Set trnlsts = prj.TransLists

	TotTrnlsts = prj.TransLists.Count

	For Each lang In langlst

      Set srcDict = CreateObject("Scripting.Dictionary")

	  For CountTrnlsts = 1 To TotTrnlsts
			Set trnlst = trnlsts(CountTrnlsts)
			fname=LCase(trnlst.TargetFile)

            'If trnlst.Property ("M:mark")="drop3" Then 'Only process the current files.
			If trnlst.Selected =True And trnlst.ExportFile ="" Then 'And InStr(1,fname,"\form\")=0 And InStr(1,fname,"\iform\")=0 Then

			szLang = trnlst.Language.LangCode
			If szLang = lang.LangCode Then


				' Go throu all strings
				StrTot = trnlst.StringCount
				For EachStr = 1 To StrTot

					Set trnstr = trnlst.String(EachStr)

					' If resource type is not version & type is not font & not translated & not to review & not read only
					' then string (words) is evaluated for NEW/UPDATED/REPEATED logic.
					' Also adds string to repetition dictionary: SourceText StrCount WordCount RepNew RepUpd where:
					' SourceText = English string
					' WordCount = words count
					' RepNew = repetitions counter for new strings
					' RepUpd = repetitions counter for updated strings

					If trnstr.ResType <> "Version" And trnstr.Type <> "DialogFont" And trnstr.State(pslStateTranslated) = True And trnstr.State(pslStateReview) = True And _
					  trnstr.State(pslStateReadOnly) = False And  trnstr.State(pslStateHidden) = False And trnstr.Resource.State(pslStateReadOnly) = False Then

						' Word count for the string
						WTmpStr = 0
						szStrTmp = trnstr.SourceText
						Set StrTmpCnt = PSL.GetTextCounts(szStrTmp)
						WTmpStr = StrTmpCnt.WordCount

						If WTmpStr > 0 Then


							If Not srcDict.Exists(szStrTmp) Then
								' create new item
								Dim data(2) As Variant
								PSL.Output szStrTmp
								data(0) = szStrTmp
								data(1) = "repit"
								data(2) = "repit"
								' add new item to the dictionary
								srcDict.Add(szStrTmp, data)
							Else
								trnstr.State(pslStateReview)=False
                                trnstr.State(pslStateTranslated)=True
                                trnstr.Property ("HiSoftLock") ="Locked20110310"
                                trnstr.State(pslStateLocked)=True
							End If

						End If

					End If

				Next EachStr

              trnlst.Save
			End If

			End If

		Next CountTrnlsts

    Set srcDict=Nothing

	Next lang


	PSL.Output "End - Repition Lock"



End Sub
