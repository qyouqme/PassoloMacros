Attribute VB_Name = "passolo_macro_extract"
'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\WINDOWS\system32\scrrun.dll#Microsoft Scripting Runtime
Option Explicit

'Define Constant
Dim CsvName As String      'CSV file name
Dim bComment As Boolean
Dim bTransDate As Boolean
Dim bResource As Boolean
Dim errorcount As Integer
Sub Main()
  Dim prj As PslProject
  Dim trnlst As PslTransList
  Dim header As String
  Dim title As String, s1 As String, s2 As String, s As String
  Dim filePath As String, filePath1 As String
  Dim Comment As String
  Dim TransDate As Date
  Dim ResourceName As String
  Dim i As Integer
  Dim fso
  Dim myfile
  Dim strnum As String
  Dim strid As String, sourcestr As String, targetstr As String, filename As String
  Dim fsys As New FileSystemObject
  Dim dumy As Integer

        

  Set prj = PSL.ActiveProject


  ' Put put file path
  filePath = prj.Location
  If Right(filePath, 1) <> "\" Then
    filePath = filePath & "\"
  End If

   filePath1 = filePath & "exported.txt"

   If fsys.FileExists(filePath1) =True Then fsys.DeleteFile filePath1
   Open filePath1 For Binary As #11
   output_bom 11
   For Each trnlst In prj.TransLists

      If trnlst.Selected =True Then



      filename = Refine(trnlst.SourceList.SourceFile)



      For i = 1 To trnlst.StringCount

        If trnlst.String(i).ResType <> "Version" And trnlst.String(i).State(pslStateReadOnly)=False Then
	        s1 = Trim(trnlst.String(i).SourceText)
	        s2 = Trim(trnlst.String(i).Text)
	        If s1 <> "" Or s2 <> "" Then

	            Replace s1,Chr$(13)," "
	            Replace s1,Chr$(10)," "
	            Replace s2,Chr$(13)," "
	            Replace s2,Chr$(10)," "
	            Replace s1,"&nbsp;"," "
				Replace s2, "&nbsp;"," "
	            Replace s1,"&apos;","'"
				Replace s2, "&apos;","'"

	            output_str 11, s1 &  "~"+ s2+"~"+trnlst.String(i).Type+"~"+trnlst.String(i).ID  +"~"+filename

	        End If
        End If
      Next i



      End If 'Selected = True

    Next

  Close #11

  
End Sub

Private Function Refine(s As String) As String
  Dim ss As String

  ss = s
  Replace ss, "&", "&amp;"
  Replace ss, "<", "&lt;"
  Replace ss, ">", "&gt;"
  
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

Sub output_bom(fnum As Integer)
Dim bb As Byte

bb = 255
Put #fnum, , bb
bb = 254
Put #fnum, , bb
End Sub

Sub output_str(fnum As Integer, Str)
Dim bserice() As Byte
Dim bb As Byte
Dim i As Long

bserice = Str

For i = 1 To LenB(Str)

  Put #fnum, , bserice(i - 1)

Next i

bb = 13
  Put #fnum, , bb
  bb = 0
  Put #fnum, , bb
  bb = 10
  Put #fnum, , bb
  bb = 0
  Put #fnum, , bb

End Sub
Sub Replace(tmpstr, ss, tt)
Dim pos1 As Long

pos1 = 1

Do


pos1 = InStr(pos1, tmpstr, ss)

If pos1 <> 0 Then

   tmpstr = Left(tmpstr, pos1 - 1) + tt + Right(tmpstr, Len(tmpstr) - pos1 - Len(ss) + 1)
   pos1 = pos1 + Len(tt)
   
End If

Loop Until pos1 = 0


End Sub

