'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\WINDOWS\system32\scrrun.dll#Microsoft Scripting Runtime
''- Autodesk customized statistics
''- Strings/Words count for Translation Lists
''- update log path fname, fnamevar, fnamenotrans, log files generated in PSL LPU location.
''- Output error in Passolo message
''- Version 6.23


Option Explicit

Dim TotToTranS, TotToTranW, TotUpdToTranS, TotUpdToTranW, TotToReviewS, TotToReviewW, TotRepS, TotRepW As Long
Dim outputBuffer As String
Dim varlist(200) As String
Dim var_len As Integer
Sub initvarlist()
' define the var list, make sure change the var_len=[?] if you add something.
	varlist(1)="\r"
    varlist(2)="\t"
    varlist(3)="\n"
    varlist(4)="%r"
    varlist(5)="%n"
    varlist(6)="\\"
    varlist(7)="%"
    varlist(8)=Chr$(13)
    varlist(9)=Chr$(10)
    varlist(10)="%s"
    varlist(11)="%u"
    varlist(12)="%p"
    varlist(13)="%d"
    varlist(14)=""""
    varlist(15)="<"
    varlist(16)=">"
    varlist(17)="&"
    varlist(18)="&gt;"
    varlist(19)="&lt;"
    varlist(20)="&amp;"
    var_len=20

End Sub


Function addToBuffer(buffer As String) As String
	outputBuffer = outputBuffer + buffer + "~"
End Function

' Add or update an item in the dictionary with the corresponding counters
Function AddDictItem(dict As Object, key As String, trStr As String,projname As String , filename As String,lgcy As String  ) As Boolean

	' if the item does not exist in the dictionary already
	If Not dict.Exists(key) Then
		' create new item
		Dim data(4) As Variant
		data(0) = "tr"
		data(1) = trStr
		data(2)=projname
		data(3)=filename
        data(4)=lgcy

		' add new item to the dictionary
		dict.Add(key, data)

        AddDictItem = True
	Else
		'get existing item
		Dim stringIdent() As Variant
		stringIdent = dict.Item(key)

        trStr=stringIdent(1)
        projname=stringIdent(2)
        filename=stringIdent(3)
        lgcy=stringIdent(4)

        AddDictItem = False
	End If

End Function

Function LookUpDictItem(dict As Object, key As String, trStr As String ) As Boolean

	' if the item does not exist in the dictionary already
	If Not dict.Exists(key) Then

        LookUpDictItem = False
	Else
		'get existing item
		Dim stringIdent() As Variant
		stringIdent = dict.Item(key)

        trStr=stringIdent(1)
        LookUpDictItem = True
	End If

End Function
Sub Main

	Dim prj As PslProject
	Dim trnlst As PslTransList
	Dim trnlsts As PslTransLists
	Dim szLang, szFullName, szFileName, szStrTmp, trStrTmp, OldStrTmp, szPrj As String
	Dim ToTran, WToTran, UpdToTran, WUpdToTran, RepToTran, WRepToTran, ToReview, WToReview As Long
	Dim TotTrnlsts, CountTrnlsts, pos, posfromend, stringlen, PrintIt, StrTot, EachStr, WTmpStr As Long
	Dim langlst As PslLanguages
	Dim lang As PslLanguage
	Dim trnstr As PslTransString
	Dim StrTmpCnt As PslStringCounter
	Dim srcDict As Object
	Dim fname As String
    Dim fnamevar As String
    Dim fnamenotrans As String
    Dim fsys As New FileSystemObject
    Dim k As Integer
    Dim lg1 As String ,lg2 As String

	'fname="D:\!Temp\tt\"+PSL.ActiveProject.Name & "_inconsis.txt"
    fname=PSL.ActiveProject.Location + "\"+ PSL.ActiveProject.Name & "_inconsis.txt"
'   fnamevar=PSL.ActiveProject.Location + "\" + PSL.ActiveProject.Name & "_var.txt"
'   fnamenotrans=PSL.ActiveProject.Location + "\" + PSL.ActiveProject.Name & "_notrans.txt"

    If fsys.FileExists (fname)= True Then fsys.DeleteFile fname
'   If fsys.FileExists (fnamevar)= True Then fsys.DeleteFile fnamevar
'   If fsys.FileExists (fnamenotrans)= True Then fsys.DeleteFile fnamenotrans

    Open fname For Binary As #22
'   Open fnamevar For Binary As #23
'   Open fnamenotrans For Binary As #24

    output_bom 22
'   output_bom 23
'   output_bom 24

    'Create a dictionary
    Set srcDict = CreateObject("Scripting.Dictionary")

    'Initialize the varlist
    initvarlist

	PSL.Output "Start - Check All"
    'For k=1 To PSL.Projects.Count

	Set prj = PSL.ActiveProject

	Set langlst = PSL.ActiveProject.Languages
	Set trnlsts = prj.TransLists

	TotTrnlsts = prj.TransLists.Count

    PSL.Output "Phase I - Put the validated translations into Dict"

	For Each lang In langlst

	  For CountTrnlsts = 1 To TotTrnlsts
			Set trnlst = trnlsts(CountTrnlsts)
			fname=LCase(trnlst.TargetFile)

			If trnlst.Selected =True   Then

			szLang = trnlst.Language.LangCode
			If szLang = lang.LangCode Then


				' Go throu all strings
				StrTot = trnlst.StringCount
				For EachStr = 1 To StrTot

					Set trnstr = trnlst.String(EachStr)

					'Only put validated translations into the dictionary

 					If trnstr.ResType <> "Version" And trnstr.Type <> "DialogFont" And _
					   trnstr.State(pslStateReadOnly) = False And  trnstr.State(pslStateHidden) = False And _
                       trnstr.Resource.State(pslStateReadOnly) = False And  _
                       trnstr.State(pslStateLocked)=False And _
 					   trnstr.State(pslStateTranslated) = True And trnstr.State(pslStateReview) = True  Then

						' Word count for the string
						WTmpStr = 0
						'szStrTmp = LCase(Trim(trnstr.SourceText))
						szStrTmp = trnstr.SourceText
						Set StrTmpCnt = PSL.GetTextCounts(szStrTmp)
						WTmpStr = StrTmpCnt.WordCount
                        trStrTmp=trnstr.Text
                        OldStrTmp=trStrTmp

                        'Check Inconsistency
                        If trnstr.Comment ="[OLD]" Then
                           lg1="y"
                        Else
                           lg1="n"
                        End If

						If WTmpStr > 0 Then
						    Dim filename As String, prjname As String
						    filename=trnlst.TargetFile
						    prjname=prj.Name


                            lg2=lg1
						    If AddDictItem(srcDict,szStrTmp+szLang, trStrTmp,filename,prjname,lg2) = False Then

						       'Log the inconsistency
                               If OldStrTmp<>trStrTmp  Then
                                  PSL.Output szStrTmp & " is not consistent: " & OldStrTmp & " - " & trStrTmp
                                  Replace szStrTmp,Chr$(13),"\r"
                                  Replace szStrTmp,Chr$(10),"\n"
                                  Replace OldStrTmp,Chr$(13),"\r"
                                  Replace OldStrTmp,Chr$(10),"\n"
                                  Replace trStrTmp,Chr$(13),"\r"
                                  Replace trStrTmp,Chr$(10),"\n"

                                  output_str 22, lg1 & lg2+"~"+szStrTmp & "~"+prj.Name & "~" & OldStrTmp & "~" & trnlst.TargetFile & "~" & prjname & "~" & trStrTmp & "~" & filename
                               End If

							End If

						End If

						'Check No Trans
'						If szStrTmp=trStrTmp And WTmpStr > 0 Then
'						   trnstr.OutputError ("Not Trans - " + trnstr.Text)
'						   output_str 24, lg1 & "~"+szStrTmp & "~"+trStrTmp & "~" & trnlst.TargetFile
						End If

                        'Check the variables
'                        If WTmpStr > 0 Then CheckVariable szStrTmp, trStrTmp,trnlst.TargetFile,lg1

                        'Check the var variables
'                        If WTmpStr > 0 Then CheckVVariable szStrTmp, trStrTmp,trnlst.TargetFile,lg1

					End If

				Next EachStr

			End If

			End If

		Next CountTrnlsts


	Next lang


    'Next
	PSL.Output Chr(10)
	PSL.Output "End - Check - Please check the log file below"
	PSL.Output fnamevar
	PSL.Output fnamenotrans

    Close #22
'   Close #23
'   Close #24


End Sub
Sub CheckVVariable(in1, in2, filename,lecays)
	Dim srsstr As String
	Dim tgtstr As String
	Dim s1 As String
	Dim t1 As String
	Dim vl As String

    Dim svlist(200) As String
    Dim tvlist(200) As String
    Dim slen As Integer
    Dim tlen As Integer
    Dim varval As String
    Dim k1 As Integer ,k2 As Integer

    Dim m,n As Integer
    Dim orderchanged As Integer
    Dim oups As String


	srsstr=in1
	tgtstr=in2

    slen=0
	While fval(srsstr, varval,"[","]")<>0
       slen=slen+1
       svlist(slen)=varval
	Wend

    srsstr=in1
	While fval(srsstr, varval,"{","}")<>0
       slen=slen+1
       svlist(slen)=varval
	Wend

    srsstr=in1
	While fval(srsstr, varval,"#","#")<>0
       slen=slen+1
       svlist(slen)=varval
	Wend
	
    tlen=0
	While fval(tgtstr, varval,"[","]")<>0
       tlen=tlen+1
       tvlist(tlen)=varval
	Wend

    tgtstr=in2
	While fval(tgtstr, varval,"{","}")<>0
       tlen=tlen+1
       tvlist(tlen)=varval
	Wend

    tgtstr=in2
	While fval(tgtstr, varval,"#","#")<>0
       tlen=tlen+1
       tvlist(tlen)=varval
	Wend
	
    orderchanged=0

    For m=1 To slen
       For n=1 To tlen

         If svlist(m)=tvlist(n) Then
            svlist(m)=""
            tvlist(n)=""
            If n<>m Then orderchanged=1
            Exit For
         End If
       Next n
    Next m

    k1=0
    For m=1 To slen
        If svlist(m)<>"" Then
           If k1=0 Then oups="The following variables are removed from the translation: "
           oups=oups+""""+svlist(m)+""""
           k1=1
        End If
    Next m

    k2=0
    For m=1 To tlen
        If tvlist(m)<>"" Then
           If k2=0 Then oups=oups+"The following variables are added to the translation: "
           oups=oups+""""+tvlist(m)+""""
           k2=1
        End If
    Next m

    If k1<>0 Or k2<>0 Then
       s1=in1
       t1=in2
       Replace s1,Chr$(13),"<chr$(13)>"
	   Replace s1,Chr$(10),"<chr$(10)>"
       Replace t1,Chr$(13),"<chr$(13)>"
	   Replace t1,Chr$(10),"<chr$(10)>"

      output_str 23, "********************************************************"
      output_str 23, lecays & "~"+s1 & "~"+t1 & "~" & filename
      output_str 23,oups
      'output_str 23, "********************************************************" & Chr$(13) & Chr$(10)
      output_str 23, Chr$(13) & Chr$(10)

    End If

    If k1=0 And k2=0 And orderchanged=1 Then
       s1=in1
       t1=in2
       Replace s1,Chr$(13),"<chr$(13)>"
	   Replace s1,Chr$(10),"<chr$(10)>"
       Replace t1,Chr$(13),"<chr$(13)>"
	   Replace t1,Chr$(10),"<chr$(10)>"
	   output_str 23, "********************************************************"
       output_str 23, lecays & "~"+s1 & "~"+t1 & "~" & filename
       output_str 23, "------------------------ Order changed -----------------"
       output_str 23, Chr$(13) & Chr$(10)

    End If

End Sub
Function fval(s1 As String,op As String,ls As String ,rs As String ) As Integer
Dim pos1 As Long
Dim pos2 As Long

fval=0
op=""
pos1=InStr(1,s1,ls)
If pos1<>0 Then

   pos2=InStr(pos1+1,s1,rs)
   If pos2>pos1 Then
      op=Mid(s1,pos1+1,pos2-pos1-1)
      s1=Right(s1,Len(s1)-pos2)
      op=ls & op & rs
      fval=1
   Else
      s1=""
      op=""
   End If
End If

End Function

Sub CheckVariable(in1, in2, filename,lecays)
	Dim srsstr As String
	Dim tgtstr As String
	Dim s1 As String
	Dim t1 As String
	Dim vl As String

    Dim c_srs As Integer , c_tgt As Integer
    Dim k As Integer

	srsstr=in1
	tgtstr=in2

    For k=1 To var_len
        c_srs=CountAstr(srsstr,varlist(k))
        c_tgt=CountAstr(tgtstr,varlist(k))
        If c_srs<>c_tgt Then
           s1=srsstr
           t1=tgtstr
           Replace s1,Chr$(13),"<chr$(13)>"
		   Replace s1,Chr$(10),"<chr$(10)>"
           Replace t1,Chr$(13),"<chr$(13)>"
		   Replace t1,Chr$(10),"<chr$(10)>"
           vl=varlist(k)
           Replace vl,Chr$(13),"<chr$(13)>"
		   Replace vl,Chr$(10),"<chr$(10)>"
		   output_str 23, "********************************************************"
           output_str 23, vl & " - not equal" & "~" & lecays & "~"+s1 & "~"+t1 & "~" & filename
           output_str 23, Chr$(13) & Chr$(10)
        End If

    Next k
End Sub
Function CountAstr(srs, spstr) As Integer
Dim tmpstr As String
Dim pos1 As Integer

CountAstr = 0

tmpstr = srs

While InStr(1, tmpstr, spstr) <> 0

  pos1 = InStr(1, tmpstr, spstr)
  CountAstr = CountAstr + 1
  tmpstr = Right(tmpstr, Len(tmpstr) - pos1 - Len(spstr) + 1)
Wend

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
