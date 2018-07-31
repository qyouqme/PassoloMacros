'#Reference {3F4DACA7-160D-11D2-A8E9-00104B365C9F}#5.5#0#C:\WINNT\System32\Vbscript.dll\3#Microsoft VBScript Regular Expressions 5.5
'' Checks string for matching 'C' format specifiers
'' Implements call back PSL_OnCheckString

'THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND
'EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES
'OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE

Option Explicit

Public Sub PSL_OnCheckString(Ctl As PslTransString)
	Dim i As Integer
	Dim srccoll As MatchCollection
	Dim trgcoll As MatchCollection
	Dim regexpr As RegExp
	Dim srcstring As String
	Dim trgstring As String

	' If we dont have to deal with string list
	If Ctl.ResType <> "String Table" Then Exit Sub

	' No C-format specifiers
	If InStr(Ctl.SourceText, "%") = 0 And _
       InStr(Ctl.Text, "%") = 0 Then Exit Sub

	Set regexpr = New RegExp
	regexpr.Pattern = _
       "([^%]|^)%-?\d*\.?\d*(ld|lo|lx|d|o|x|h|c|s|f|n)"
	regexpr.IgnoreCase = True
	regexpr.Global=True
	Set srccoll = regexpr.Execute(Ctl.SourceText)
	Set trgcoll = regexpr.Execute(Ctl.Text)

	If srccoll.Count <> trgcoll.Count Then
		Ctl.OutputError _
            "Different number of format specifier"
		Exit Sub
	End If

	For i = 0 To srccoll.Count - 1
		srcstring = srccoll.Item(i)
		trgstring = trgcoll.Item(i)

		If Left(srcstring,1) <> "%" Then
			srcstring = Mid(srcstring, 2)
		End If

		If Left(trgstring,1) <> "%" Then
			trgstring = Mid(trgstring,2)
		End If

		If srcstring <> trgstring Then
			Ctl.OutputError _
                CStr(i + 1) + ". Format specifier different"
			Exit Sub
		End If
	Next
End Sub
