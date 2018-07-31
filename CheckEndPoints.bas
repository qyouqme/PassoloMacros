'#Reference {C1D8C091-AC66-4159-B738-E70A12B983A4}#1.0#0#C:\Program Files\Apple Software Update\ScriptingObjectModel.dll#ScriptingObjectModel
''[de,es,fr,it,sc,tc,ja,ko]Check the number of points in the source text if equal the number in the targer text.
Sub Main

start_time = Timer()

'������ر���
Dim prj As PslProject
Dim trnlst As PslTransList
Dim  i, j, arrcount As Long
Dim trans_str As PslTransString
Dim myLang As PslLanguage
Dim endPoint As String
Dim srcPoint As String
Dim tgtPoint As String
Dim csvName As String
Dim csvHeader As String
Dim fso
Dim flag As Boolean



Set prj = PSL.ActiveProject

flag = False

'������־�ļ�
csvName =prj.Location
If Right(csvName, 1) <> "\" Then csvName = csvName & "\"
csvName = csvName & prj.Name & ".csv"

Set fso = CreateObject("Scripting.FileSystemObject")
'�����־�ļ����ڣ���ɾ��
If Dir(csvName)<>"" Then fso.DeleteFile (csvName)
Set myfile = fso.CreateTextFile (csvName,True,True)
' Make header for output file
csvHeader = "Language" & Chr(9) & "FileName" & Chr(9) & "Source String" & Chr(9) & "Target String" & Chr(9)
myfile.WriteLine csvHeader

'����ÿ���ִ�
For i = 1 To prj.TransLists.Count
    Set trnlst = prj.TransLists(i)
    For j = 1 To trnlst.StringCount
        Set trans_str = trnlst.String(j)
        '�ִ���Ӧ������
        Set myLang = trnlst.Language

		'�������Զ���������
        Select Case CStr(myLang.LangCode)
	        Case "deu","esn","fra","ita"
	            endPoint = "."
	        Case "chs","cht","jpn","kor"
	            endPoint = "��"
	        Case Else
	            endPoint = ""
    	End Select
		'��ȡsource�ִ���target�ִ������һ���ַ�
		srcPoint =Right(Trim(trans_str.SourceText),1)
		tgtPoint =Right(Trim(trans_str.Text),1)


		'��������ڼ������ԣ�����������ѭ��
		If(endPoint="") Then GoTo NextString
		'�ж�source�ִ������һ���ַ��ǲ���.��
		If(srcPoint = ".") Then
			GoTo HavePoinit
		Else
			GoTo NoPoint
		End If

		'���target�ִ������һ���ַ����ڽ�����㣬����������ѭ��
		HavePoinit:
		If(tgtPoint =endPoint) Then
			GoTo NextString
		Else
			GoTo writeToLog
		End If

		'���target�ִ������һ���ַ����ڽ�����㣬����������ѭ��
		NoPoint:
		If(tgtPoint <>endPoint) Then
		GoTo NextString
		Else
			GoTo writeToLog
		End If

		writeToLog:
		'д����־�ļ�
		Dim fileName As String
		Dim src As String
		Dim tgt As String
		Dim myLine As String

		fileName = GetTitle(trnlst.TargetFile)
		src = trans_str.SourceText
		tgt = trans_str.Text
		myLine = CStr(myLang.LangCode) & Chr(9) & fileName & Chr(9) & src & Chr(9) & tgt & Chr(9)
		myfile.WriteLine myLine

		flag = True

    NextString:
    Next j
Next i

myFile.Close

PSL.Output("Done in '" & 	Timer() - start_time & " secs")

If(flag) Then
	MsgBox("Please see "&CStr(csvName)&" for details")
Else
	MsgBox("No errors find")
End If





End Sub
Private Function GetTitle(path As String) As String
  Dim s As String, ss As String
  Dim n As Integer

  s = Trim(path)
  Do
    n = InStr(s, "\")
    If n < 1 Then
      Exit Do
    End If
    s = Mid(s, n + 1)
  Loop
  n = InStr(s, ".")
  If n < 1 Then
    ss = s
    s = ""
  End If
  Do Until s = ""
    n = InStr(s, ".")
    If n < 1 Then
      Exit Do
    End If
    If ss = "" Then
      ss = Left(s, n - 1)
    Else
      ss = ss & "." & Left(s, n - 1)
    End If
    s = Mid(s, n + 1)
  Loop
  GetTitle = ss
End Function
