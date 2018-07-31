'#Reference {F5078F18-C551-11D3-89B9-0000F81FE221}#4.0#0#C:\WINDOWS\system32\msxml4.dll#Microsoft XML, v4.0
'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\WINDOWS\system32\SCRRUN.DLL#Microsoft Scripting Runtime
''- Autodesk customized statistics
''- Strings/Words count for Translation Lists
''- Version 6.22


Option Explicit

Type fileinfo
  Name As String
  changedtime As String
End Type

Dim fall(6000) As fileinfo

Dim TotToTranS, TotToTranW, TotUpdToTranS, TotUpdToTranW, TotToReviewS, TotToReviewW, TotRepS, TotRepW As Long
Dim outputBuffer As String

Function addToBuffer(buffer As String) As String
	outputBuffer = outputBuffer + buffer + "~"
End Function

' Add or update an item in the dictionary with the corresponding counters
Function AddDictItem(dict As Object, key As String, trStr As String ) As Boolean

	' if the item does not exist in the dictionary already
	If Not dict.Exists(key) Then
		' create new item
		Dim data(1) As Variant
		data(0) = "tr"
		data(1) = trStr

		' add new item to the dictionary
		dict.Add(key, data)

        AddDictItem = True
	Else
		'get existing item
		Dim stringIdent() As Variant
		stringIdent = dict.Item(key)

        trStr=stringIdent(1)
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
    Dim fsys As New FileSystemObject
    Dim k As Integer
    Dim lg1 As String ,lg2 As String
    Dim fd As Folder
    Dim fl As File
    Dim sbfd As Folder
    Dim filePath As String


	Begin Dialog UserDialog 800,196,"Export to XML",.TMAutotrans ' %GRID:10,7,1,1
		Text 20,35,260,21,"Please specify destination:",.Text1
		TextBox 30,63,650,21,.tmname
		OKButton 300,119,60,21,.d
		CancelButton 440,119,60,21
		PushButton 700,63,50,21,"...",.SelectFile
	End Dialog

   Dim dlg As UserDialog

  If Dialog(dlg)=0 Then Exit Sub

  filePath=dlg.tmname

   If Right(filePath,1)<>"\" Then filePath=filePath+"\"

    fname=filePath

    'Open fname+"log.txt" For Output As #1

    'Write into Dictionary
    Set srcDict = CreateObject("Scripting.Dictionary")

    Set fd=fsys.GetFolder (fname)

    processfolder fd,srcDict

    Autotrans srcDict

   ' Close #1

End Sub
Sub processfolder(fd As Folder,srcDict As Object  )
	Dim fl As File
	Dim fsys As New FileSystemObject
    Dim sbfd As Folder

	For Each fl In fd.Files

      If LCase(fsys.GetExtensionName (fl.Name))="xml" Then
         PutinDict fl.Path, srcDict
      End If
	Next

    For Each sbfd In fd.SubFolders
        processfolder sbfd,srcDict
    Next

End Sub
Sub PutinDict(filename As String, srcDict As Object )
Dim xmldoc As New DOMDocument40
Dim rootnode As IXMLDOMNode
Dim fileslist As IXMLDOMNodeList
Dim filenode As IXMLDOMNode
Dim tmpnode As IXMLDOMNode

Dim srslist As IXMLDOMNodeList, tgtlist As IXMLDOMNodeList,idlist As IXMLDOMNodeList,parserlist As IXMLDOMNodeList
Dim strlist As IXMLDOMNodeList
Dim ttnode1 As IXMLDOMNode
Dim ttnode2 As IXMLDOMNode
Dim ttnode3 As IXMLDOMNode
Dim ttnode4 As IXMLDOMNode
Dim ttnode5 As IXMLDOMNode
Dim ttnode As IXMLDOMNode

xmldoc.async = False
xmldoc.validateOnParse = False
xmldoc.preserveWhiteSpace = True
xmldoc.resolveExternals = False

xmldoc.load (filename)

If xmldoc.parseError.errorCode <>0 Then MsgBox filename

If xmldoc.parseError.errorCode <>0 Then
   MsgBox "File not well Formed"
   End
End If

Dim projectname As String, filen As String

Set rootnode=xmldoc.selectSingleNode ("dictionary/project/projectname")
projectname=rootnode.Text

Dim languagename As String
Set rootnode=xmldoc.selectSingleNode ("dictionary/project/language/languagename")
languagename=rootnode.Text

Set fileslist=xmldoc.getElementsByTagName ("file")

Dim m As Integer, tmpstr As String
Dim parsertype As Integer




For m=1 To fileslist.length

   Set tmpnode=fileslist.Item(m-1)
   Set rootnode=tmpnode.selectSingleNode ("filename")
   filen=rootnode.Text

   'Set srslist=tmpnode.selectNodes ("StringLists/StringList/source")
   'Set tgtlist=tmpnode.selectNodes ("StringLists/StringList/target")
   'Set idlist=tmpnode.selectNodes ("StringLists/StringList/id")
   'Set parserlist=tmpnode.selectNodes ("StringLists/StringList/parser")

   Set strlist=tmpnode.selectNodes ("StringLists/StringList")

   Dim k As Integer



   For k=1 To strlist.length

       Set ttnode=strlist.Item (k-1)

       Set ttnode1=ttnode.selectSingleNode ("source")
       Set ttnode2=ttnode.selectSingleNode ("target")
       Set ttnode3=ttnode.selectSingleNode ("id")
       Set ttnode4=ttnode.selectSingleNode ("parser")


       tmpstr=projectname & languagename & filen

       Dim tmps As String

       If ttnode1 Is Nothing Or ttnode2 Is Nothing Or ttnode3 Is Nothing Or ttnode4 Is Nothing  Then

       Else

       tmps=ttnode3.Text 'id

       tmpstr=tmpstr+CStr(tmps)

         tmps=ttnode1.Text 'source

          Replace tmps,"\\r","hzwrrr"
          Replace tmps,"\\n","hzwnnn"
          'Replace tmps,"\\","\"
          Replace tmps,"\r",Chr$(13)
          Replace tmps,"\n",Chr$(10)
          Replace tmps,"hzwrrr","\r"
          Replace tmps,"hzwnnn","\n"

       If ttnode4.Text ="1" Then 'Parser node

             '    Replace tmps,"&","&amp;"
              '   Replace tmps,"<","&lt;"
              '   Replace tmps,">","&gt;"
       End If

       tmpstr=tmpstr+tmps

       If InStr(1,tmps,"VoiceFolder")<>0 Then

          'MsgBox "SDS"
       End If

          tmps=ttnode2.Text 'Target node
		  Replace tmps,"\""","\\"""
          Replace tmps,"\\r","hzwrrr"
          Replace tmps,"\\n","hzwnnn"
          'Replace tmps,"\\","\"
          Replace tmps,"\r",Chr$(13)
          Replace tmps,"\n",Chr$(10)
          Replace tmps,"hzwrrr","\r"
          Replace tmps,"hzwnnn","\n"

       Dim rst As Boolean

       parsertype=CInt(ttnode4.Text)

       Replace tmps,"&apos;","hzwsingle"
       Replace tmps,"&quot;","hzwdouble"

       If parsertype And 4 Then 'Parser node

            Replace tmps,"&","&amp;"
       End If

       If parsertype And 1 Then 'Parser node

            Replace tmps,"<","&lt;"
       End If

       If parsertype And 2 Then 'Parser node

            Replace tmps,">","&gt;"
       End If

       Replace tmps,"hzwsingle","&apos;"
       Replace tmps,"hzwdouble","&quot;"


       If ttnode4.Text="8" Then 'Parser node

                 Replace tmps,"&hisoft","&amp;"
				 Replace tmps,"</hisoft","&lt;/"
                 Replace tmps,"<hisoft","&lt;"
                 Replace tmps,"hisoft/>","/&gt;"
                 Replace tmps,"hisoft>","&gt;"
       End If

       rst=AddDictItem(srcDict, tmpstr,tmps)

    '   Print #1,tmpstr

       End If

   Next k
Next m
End Sub

Sub Autotrans(srcDict As Object )

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

	Dim fname As String
    Dim rst As Boolean

    Dim rrs As String

    Dim rootnode As IXMLDOMNode, snode As IXMLDOMNode,tnode As IXMLDOMNode,langnode As IXMLDOMNode
    Dim m As Long
    Dim nodelist As IXMLDOMNodeList
    Dim filePath As String
    Dim szStrTmpv As String



	Set prj = PSL.ActiveProject
	Set langlst = PSL.ActiveProject.Languages
	Set trnlsts = prj.TransLists

	TotTrnlsts = prj.TransLists.Count

    Dim prjname As String,langname As String


    prjname=prj.Name


	For Each lang In langlst


    langname=lang.LangCode

    PSL.Output "Phase II - Get the translation"

	'For Each lang In langlst

	  For CountTrnlsts = 1 To TotTrnlsts

			Set trnlst = trnlsts(CountTrnlsts)


			fname=trnlst.SourceList.SourceFile

			Dim tpstr As String

            tpstr=prjname+langname+fname

			If trnlst.Selected =True  Then
            'If trnlst.Property ("M:mark")="drop3" Then

			szLang = trnlst.Language.LangCode
			If szLang = lang.LangCode Then


				' Go throu all strings
				StrTot = trnlst.StringCount
				For EachStr = 1 To StrTot

					Set trnstr = trnlst.String(EachStr)

                    tpstr=prjname+langname+fname
					tpstr=tpstr+CStr(EachStr)

                    'Update the Tu which the source is empty and the translation is not empty
					If Trim(trnstr.SourceText)="" And Trim(trnstr.Text)<>"" Then
						trnstr.Text=trnstr.SourceText
					End If

					'Only put validated translations into the dictionary
 					If trnstr.ResType <> "Version" And trnstr.Type <> "DialogFont" And _
					   trnstr.State(pslStateReadOnly) = False And  trnstr.State(pslStateHidden) = False And _
                       trnstr.Resource.State(pslStateReadOnly) = False  And  _
                       Not (trnstr.State(pslStateTranslated) = True) And trnstr.SourceText<>"" Then 'And trnstr.State(pslStateReview) = False)   Then

						' Word count for the string
						WTmpStr = 0
						szStrTmp = tpstr+trnstr.SourceText
						Set StrTmpCnt = PSL.GetTextCounts(szStrTmp)
						WTmpStr = StrTmpCnt.WordCount
                        trStrTmp=trnstr.Text
                        OldStrTmp=trStrTmp

                        'Print #1,szStrTmp

						If WTmpStr > 0 Then
						    If LookUpDictItem(srcDict,szStrTmp, trStrTmp) = False Then
						       If Trim(trnstr.SourceText)<>"" Then

                                  PSL.Output "The translation for " & szStrTmp & " can not be found
                               Else
                               trnstr.Text= trnstr.SourceText
                               'trnstr.TransComment  ="Updated On 06-17" 'trnstr.TransComment+"Autotrans on 090903"
                               trnstr.State(pslStateTranslated) = True
                               trnstr.State(pslStateReview) = True
                               End If
                            Else
                               trnstr.Text= trStrTmp
                               'trnstr.TransComment  ="Updated On 06-17" 'trnstr.TransComment+"Autotrans on 090903"
                               trnstr.State(pslStateTranslated) = True
                               trnstr.State(pslStateReview) = True

							End If

						End If

					End If

				Next EachStr

              trnlst.Save
			End If

			End If

		Next CountTrnlsts


	Next lang

	PSL.Output "End - Autotrans"



End Sub

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
Private Function TMAutotrans(DlgItem$, Action%, SuppValue&) _
As Boolean
	Dim Folder As String

	' We only wnat to check button clicks
	If Action% <> 2 Then Exit Function


	' Let user select TM file
	If DlgItem$ = "SelectFile" Then
	    Folder=PSL.ActiveProject.Location
	    If PSL.SelectFolder(Folder,"Specify folder") Then
			DlgText "tmname", Folder
		End If
		TMAutotrans = True
	End If


End Function
