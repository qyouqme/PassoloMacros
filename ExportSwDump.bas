' TransDump.bas
' Version 1.0 Created 13/04/2005 by Hidenori Yoshizumi
'
''This macro exports English and Japanese strings to CSV file.
' It is possible to select optional field such as "Comment", "Translation Date", "Resource"
'
' Version 1.1 05/10/2005 by Hidenori Yoshizumi
'  - Resouce is default and placed in the second row.


Option Explicit

'Define Constant
Dim CsvName As String      'CSV file name
Dim bComment As Boolean
Dim bTransDate As Boolean
Dim bResource As Boolean
Dim errorcount As Integer

Function DialogFunc(DlgItem As String, Action As Integer, SuppValue As Integer) As Boolean
	Dim i As Integer

	Select Case Action
	Case 1 ' Dialog box initialization
		bComment = False
		bTransDate = False
		bTransDate = False
	Case 2 ' Value changing or button pressed
		Select Case DlgItem
			Case "chkComment"
				Select Case DlgValue("chkComment")
				Case 0
					bComment = False
				Case 1
					bComment = True
				End Select
			Case "chkTranslationDate"
				Select Case DlgValue("chkTranslationDate")
				Case 0
					bTransDate = False
				Case 1
					bTransDate = True
				End Select
		End Select
	End Select

End Function

Sub Main()
  Dim prj As PslProject
  Dim trnlst As PslTransList
  Dim header As String
  Dim title As String, s1 As String, s2 As String, s As String
  Dim filePath As String
  Dim Comment As String
  Dim TransDate As Date
  Dim ResourceName As String
  Dim i As Long
  Dim fso
  Dim myfile

  Dim dumy As Integer

  		Begin Dialog UserDialog 450,150,"Translation Dump.",.dialogfunc ' %GRID:10,7,1,1

			GroupBox 10,10,430,90,"Please select optional fields",.GroupBox1
			CheckBox 30,30,160,21,"Comment",.chkComment
			CheckBox 30,50,160,21,"Translation Date",.chkTranslationDate
			OKButton 270,120,80,21
			CancelButton 360,120,80,21
		End Dialog

		Dim dlg As UserDialog

		If Dialog(dlg) = 0 Then Exit Sub
		errorcount = 0

  If PSL.Projects.Count < 1 Then
    MsgBox "No project is opened.", vbCritical, "Warning - PageOne"
    Exit Sub
  End If

  Set prj = PSL.ActiveProject

  ' Put put file path
  filePath = prj.Location
  If Right(filePath, 1) <> "\" Then
    filePath = filePath & "\"
  End If
  filePath = filePath & prj.Name & ".csv"

  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myfile = fso.CreateTextFile(filePath, True, True)

  ' Make header for output file
  header = "Title" & Chr(9) & "Resource" & Chr(9) & "Number" & Chr(9) & "ID" & Chr(9) & "English" & Chr(9) & "Localized"
  If bComment Then
	header = header & Chr(9) & "Comment"
  End If
  If bTransDate Then
	header = header & Chr(9) & "Translation Date"
  End If

  myfile.WriteLine header & Chr(9) &  "Status"


    For Each trnlst In prj.TransLists
      title = GetTitle(trnlst.TargetFile)
      'MsgBox CStr(trnlst.StringCount)
      For i = 1 To trnlst.StringCount
        s1 = Trim(trnlst.String(i).SourceText)
        s2 = Trim(trnlst.String(i).Text)
        If s1 <> "" Or s2 <> "" Then
          ResourceName = trnlst.String(i).Resource.Type & " " & trnlst.String(i).Resource.ID

		  s = Refine(title) & Chr(9) & Refine(ResourceName) & Chr(9) & Refine(Str(trnlst.String(i).Number)) & Chr(9)
		  If trnlst.String(i).IDName = "" Then
			s = s & Refine(CStr(trnlst.String(i).ID))
 		  Else
			s = s & Refine(CStr(trnlst.String(i).IDName))
          End If
		  s = s & Chr(9) & Refine(trnlst.String(i).SourceText) & Chr(9) & Refine(trnlst.String(i).Text)

		  '### Additional Fields
		  Comment   = Refine(trnlst.String(i).Comment)
		  TransDate = trnlst.String(i).DateTranslated

		  If bComment Then
		  	s = s & Chr(9) & Comment
		  End If
		  If bTransDate Then
		  	s = s & Chr(9) & Month(TransDate) & "/" & Day(TransDate) & "/" & Year(TransDate)
		  End If
		  '###

          If trnlst.String(i).State(pslStateReadOnly) Then
			s = s & Chr(9) & "Locked"
          ElseIf trnlst.String(i).State(pslStateReview) Then
			s = s & Chr(9) & "Review"
          ElseIf Not trnlst.String(i).State(pslStateTranslated) Then
			s = s & Chr(9) & "Not Translated"
          ElseIf trnlst.String(i).State(pslStateBookmark) Then
			s = s & Chr(9) & "Bookmark"
          End If

          myfile.WriteLine s
        End If
      Next i
    Next

  myfile.Close
  MsgBox "Done"
End Sub

Private Function Refine(s As String) As String
  Dim ss As String

  ss = Trim(s)
  ss = Replace(ss, Chr(10), "\n")
  ss = Replace(ss, Chr(13), "\r")
  ss = Replace(ss, Chr(9), "\t")
  ss = Replace(ss, """", """""")
  Refine = ss
End Function

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
