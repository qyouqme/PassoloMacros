'#Reference {73280B25-59E4-11D1-8F65-00A0C9101CA1}#3.0#0#E:\Program Files\TRADOS\Team-3\TW4Win\TW4Win.tlb#TRADOS Translator's Workbench Type Library
'' Look up translations in TRADOS Workbench
'' Implements call back PSL_OnCheckString

'THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND
'EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES
'OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE

Option Explicit

Dim TW3App As TW4Win.Application
Dim TW3Mem As TW4Win.TranslationMemory

Public Sub PSL_OnAutoTranslate(Translations As PslTranslations, _
                       ByVal MinMatch As Long, ByVal MaxCount As Long)
    Dim TUnit As TranslationUnit, i As Integer

    If TW3Mem Is Nothing Or Translations.SourceString = "" Then Exit Sub

    TW3Mem.Search (Translations.SourceString)
    Set TUnit = TW3Mem.TranslationUnit

    For i = 1 To TW3Mem.HitCount
      Translations.Add(TUnit.Target, TUnit.Source, TUnit.Score, "TW")
      TUnit.Next
    Next i
End Sub

Public Sub PSL_OnStartScript()
	Set TW3App = CreateObject("TW4Win.Application")

	If TW3App Is Nothing Then
		MsgBox("Can't open the TRADOS Workbench")
		Exit Sub
	End If

	Set TW3Mem = TW3App.TranslationMemory
	TW3Mem.Open("c:\ddata\demo\trados\psldemoed.tmw", "User")
End Sub
