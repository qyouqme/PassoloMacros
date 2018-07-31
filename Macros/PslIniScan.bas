'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\WINNT\System32\scrrun.dll#Microsoft Scripting Runtime
'' Allows processing of INI files
'' Implements call back PSL_OnProcessUserFile

'THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND
'EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES
'OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE

' If INI-section, return section name otherwise emtpy string
Function GetHeader(ln As String) As String
	ln = Trim(ln)

	If Left(ln,1) <> "[" Or Right(ln, 1) <> "]" Then Exit Function

	GetHeader = Mid(ln, 2, Len(ln) - 2)
End Function

' Get Position of "=" if not found or comment return 0
Function GetDataPos(ln As String) As Long
	GetDataPos = 0

	' Comment?
	If Left(ln,1) = ";" Then Exit Function

	GetDataPos = InStr(ln, "=")
End Function

' Update translation list and generate target file
Function IniUpdate(rd As PslResData) As Long
    Dim fso As Scripting.FileSystemObject
    Dim tsi As Scripting.TextStream
    Dim tso As Scripting.TextStream
    Dim s As String, header As String, i As Integer

	Set fso = New Scripting.FileSystemObject
    Set tsi = fso.OpenTextFile(rd.SourceFile)

	' Generating target file
	If rd.Action = pslResUpdGenerate Then
		Set tso = fso.CreateTextFile(rd.TargetFile)
	End If

	IniUpdate = 0

	On Error GoTo ende
	While True
		s = tsi.ReadLine
		' Handle new section header
		If GetHeader(s) <> "" Then
			header = GetHeader(s)
			rd.ProcessResource "INI-Section", header
		Else
			' Only if section header is there and line contains data
			If header <> "" Then
				i = GetDataPos(s)
				If i <> 0 Then
					rd.SetEntryData(pslResUpdId, Left(s, i - 1))
					rd.SetEntryData(pslResUpdText, Mid(s, i + 1))
					rd.ProcessEntry

					' Generate changed line
					If rd.Action = pslResUpdGenerate Then
						s = Left(s, i - 1) & "=" & _
                            rd.GetEntryData(pslResUpdText)
					End If
				End If
			End If
		End If

		' Write data to target file
		If rd.Action = 1 Then
			tso.WriteLine(s)
		End If
	Wend
ende:
End Function

Function IniListContent(rd As PslResData) As Long
    Dim fso As Scripting.FileSystemObject
    Dim tsi As Scripting.TextStream
    Dim s As String, header As String

    Set fso = New Scripting.FileSystemObject
    Set tsi = fso.OpenTextFile(rd.SourceFile)

	IniListContent = 0
	On Error GoTo ende

	' Add all sections
	While True
		s = tsi.ReadLine
		header = GetHeader(s)
		If header <> "" Then
			rd.ListResource "INI-Section", header, 0
		End If
	Wend
ende:
End Function


Function IniCheckFiles(rd As PslResData) As Long
    Dim fso As Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    Dim s As String
    Set fso = New Scripting.FileSystemObject

	' No INI file
	If UCase(Right(rd.SourceFile, 3)) <> "INI" Then GoTo failure

	' File does not exists
    On Error GoTo failure
    Set ts = fso.OpenTextFile(rd.SourceFile)

	' target file, should not be the same as source file
	If rd.TargetFile <> "" Then
    	If UCase(rd.SourceFile) = UCase(rd.TargetFile) Then _
    	    GoTo failure
	End If

	IniCheckFiles = 0
	Exit Function

failure:
	IniCheckFiles = 1

End Function

Function IniGetLanguages(rd As PslResData) As Long
	' Only Language neutral
	IniGetLanguages = 0
	rd.AddLanguage(0)
End Function

Public Sub PSL_OnProcessUserFile(rd As PslResData)
	rd.Error = 0

	If rd.Action = pslResUpdGetFileExtensions Then
		rd.AddFileExtension("INI Files (*.ini)", "*.ini")
		Exit Sub
	End If

	' No Ini file
	If UCase(Right(rd.SourceFile, 3)) <> "INI" Then Exit Sub

	Select Case rd.Action
		Case pslResUpdScanTarget
			rd.Error = IniUpdate(rd)
		Case pslResUpdGenerate
			rd.Error = IniUpdate(rd)
		Case pslResUpdScanData
			rd.Error = IniUpdate(rd)
		Case pslResUpdListContent
			rd.Error = IniListContent(rd)
		Case pslResUpdCheckSourceTarget
			rd.Error = IniCheckFiles(rd)
		Case pslResUpdGetLanguages
			rd.Error = IniGetLanguages(rd)
		Case Else
			rd.Error = 1
	End Select
End Sub


