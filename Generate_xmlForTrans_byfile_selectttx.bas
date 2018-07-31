Attribute VB_Name = "passolo_macro_extract"
'#Reference {F5078F18-C551-11D3-89B9-0000F81FE221}#4.0#0#C:\WINDOWS\system32\msxml4.dll#Microsoft XML, v4.0
'#Reference {F5078F18-C551-11D3-89B9-0000F81FE221}#5.0#0#C:\Program Files\Common Files\Microsoft Shared\OFFICE11\MSXML5.DLL#Microsoft XML, v5.0
'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\WINDOWS\system32\scrrun.dll#Microsoft Scripting Runtime
Option Explicit

'Define Constant
Dim CsvName As String      'CSV file name
Dim bComment As Boolean
Dim bTransDate As Boolean
Dim bResource As Boolean
Dim errorcount As Integer

Dim xmldoc As New DOMDocument40
  Dim prj As PslProject
  Dim trnlst As PslTransList
  Dim header As String
  Dim title As String, s1 As String, s2 As String, s As String
  Dim filePath As String, filePath1 As String
  Dim Comment As String
  Dim TransDate As Date
  Dim ResourceName As String
  Dim i As Long
  Dim fso
  Dim myfile
  Dim strnum As String
  Dim strid As String, sourcestr As String, targetstr As String, filename As String,filename1 As String
  Dim fsys As New FileSystemObject
  Dim dumy As Long
  Dim rootnode As IXMLDOMNode, snode As IXMLDOMNode, filenamenode As IXMLDOMNode,tnode As IXMLDOMNode,langnode As IXMLDOMNode,txtnode As IXMLDOMNode
  Dim idnode As IXMLDOMNode
  Dim kl As Long
  Dim oupath As String
  Dim filenode As IXMLDOMNode
  Dim m As Long
  Dim projnode As IXMLDOMNode
  Dim projnamenode As IXMLDOMNode
   Dim langnamenode As IXMLDOMNode
   Dim filesnode As IXMLDOMNode
        Dim strsnode As IXMLDOMNode

        Dim filenum As Long
        Dim itemnum As Long
Dim option1 As Boolean
Dim option2 As Boolean
Dim option3 As Boolean
Dim option4 As Boolean
Dim option5 As Boolean
Dim option6 As Boolean
Dim option7 As Boolean
Dim cond1 As Boolean
Dim cond2 As Boolean
Dim cond3 As Boolean
Dim cond4 As Boolean
Dim cond5 As Boolean

Dim tuvnode1 As IXMLDOMNode
Dim tuvnode2 As IXMLDOMNode
Dim tunode As IXMLDOMNode
Dim nodelist As IXMLDOMNodeList



Sub Main()

	Begin Dialog UserDialog 800,245,"Export to XML",.TMAutotrans ' %GRID:10,7,1,1
		Text 20,35,260,21,"Please specify destination:",.Text1
		TextBox 30,63,650,21,.tmname
		OKButton 300,189,70,21,.d
		CancelButton 430,189,90,21
		PushButton 700,63,50,21,"...",.SelectFile
		OptionGroup .Group1
			OptionButton 100,105,160,14,"Source",.OptionButton1
			OptionButton 100,126,200,21,"Target",.OptionButton2
		GroupBox 290,91,470,91,"Translation Setting",.GroupBox1
		OptionGroup .Group2
			OptionButton 310,112,200,14,"Only Validated Units",.OptionButton3
			OptionButton 310,133,220,14,"Only For Review Units",.OptionButton4
			OptionButton 520,105,220,21,"Validated and For Review",.OptionButton5
			OptionButton 520,133,160,14,"New and Update",.OptionButton6
			OptionButton 310,154,210,14,"Only Updated String",.OptionButton7
	End Dialog

   Dim dlg As UserDialog
   Dim extname As String
   Dim parsertype As Long
   Dim tempstr1 As String

  If Dialog(dlg)=0 Then Exit Sub

  filePath=dlg.tmname

   If Right(filePath,1)<>"\" Then filePath=filePath+"\"



  Set prj=PSL.ActiveProject

  For kl=1 To prj.Languages.Count


   For Each trnlst In prj.TransLists



     If trnlst.Language.LangCode =prj.Languages.Item(kl).LangCode And trnlst.Selected =True  Then

      filenum=1
      CreateHead
      itemnum=0

      For i = 1 To trnlst.StringCount


        s1 = Trim(trnlst.String(i).SourceText)
        s2 = Trim(trnlst.String(i).Text)
        If InStr(1,s1,"This feature will be uninstalled completely")=1 Then
            s1 = Trim(trnlst.String(i).SourceText)
        End If
        If s1 <> "" Then

          cond1=option3=True And trnlst.String(i).State(pslStateTranslated)=True And trnlst.String(i).State(pslStateReview)=False
          cond2=option4=True And trnlst.String(i).State(pslStateTranslated)=True And trnlst.String(i).State(pslStateReview)=True
          cond3=option5=True And trnlst.String(i).State(pslStateTranslated)=True
          cond4=option6=True And trnlst.String(i).State(pslStateTranslated)=False
          cond5=option7=True And trnlst.String(i).OldText<>""

          If trnlst.String(i).ResType <> "Version" And trnlst.String(i).ResType <>"DialogFont" And trnlst.String(i).State(pslStateReadOnly)=False And trnlst.String(i).Resource.State(pslStateReadOnly)=False And _
          (cond1 Or cond2 Or cond3 Or cond4 Or cond5)  And _
          trnlst.String(i).State(pslStateHidden)=False And trnlst.String(i).State(pslStateLocked)=False Then

              itemnum=itemnum+1

              If itemnum>1000 Then

                 CreateTail
                 xmldoc.Save oupath
                 Set nodelist=xmldoc.getElementsByTagName("Tu")
			       If nodelist.length <> 0 Then
			          Open oupath For Binary As #22
			          output_bom 22
			          output_str 22,xmldoc.xml
			          Close #22
			       End If


                 itemnum=0
                 filenum=filenum+1
                 CreateHead

              End If

              Dim tmpnode As IXMLDOMNode
              Set tmpnode=xmldoc.selectSingleNode(".//UserSettings")
              tmpnode.Attributes.getNamedItem("SourceLanguage").Text=PSL.GetLangCode ( trnlst.SourceList.LangID,8)

              Dim langcd As String
              langcd=PSL.GetLangCode (trnlst.Language.LangID,8)
              If LCase(langcd)="he" Then langcd="IW"

              tmpnode.Attributes.getNamedItem("TargetLanguage").Text=langcd

			  tmpnode.Attributes.getNamedItem("TargetDefaultFont").Text="Arial Unicode MS"

    	      sourcestr = trnlst.String(i).SourceText
              If cond5 Then sourcestr = trnlst.String(i).OldText

    	      Replace sourcestr,"\r","\\r"
    	      Replace sourcestr,"\n","\\n"
    	      Replace sourcestr,Chr$(13),"\r"
              Replace sourcestr,Chr$(10),"\n"

			  If option1=True Then

        	  	targetstr =trnlst.String(i).SourceText

        	  Else
				targetstr =trnlst.String(i).Text
        	  End If

    	      Replace targetstr,"\r","\\r"
    	      Replace targetstr,"\n","\\n"
    	      Replace targetstr,Chr$(13),"\r"
              Replace targetstr,Chr$(10),"\n"


			  rootnode.appendChild CreateUTnodeNew(xmldoc,"<StringList>")
              rootnode.appendChild xmldoc.createTextNode (Chr$(13)+Chr$(10))

			  rootnode.appendChild CreateUTnodeNew(xmldoc,"<source>"+Refine(sourcestr)+"</source>")
              rootnode.appendChild xmldoc.createTextNode (Chr$(13)+Chr$(10))




              Dim parsev As String
              parsev="0"
              extname=LCase(fsys.GetExtensionName(trnlst.SourceList.SourceFile))



			  'Preprocess source string
			  sourcestr=PreprocessStr(sourcestr,parsev,extname)

			  'Preprocess target string
			  targetstr=PreprocessStr(targetstr,parsev,extname)

			  rootnode.appendChild CreateUTnodeNew(xmldoc,"<target><![CDATA[")

              Set tuvnode1=xmldoc.createElement ("Tuv")
			  Set tuvnode2=xmldoc.createElement ("Tuv")

			  Dim attr As IXMLDOMAttribute

			  Set attr=xmldoc.createAttribute("Lang")
			  attr.Text=PSL.GetLangCode ( trnlst.SourceList.LangID,8)
			  tuvnode1.Attributes.setNamedItem attr.cloneNode(True )

              langcd=PSL.GetLangCode (trnlst.Language.LangID,8)
              If LCase(langcd)="he" Then langcd="IW"

              attr.Text=langcd
		   	  tuvnode2.Attributes.setNamedItem attr.cloneNode(True )

              Set tunode=xmldoc.createElement ("Tu")

              Replace sourcestr,"&apos;","'"
              Replace targetstr,"&apos;","'"

              tuvnode1.appendChild xmldoc.createTextNode (sourcestr)

			  Set tmpnode=xmldoc.createElement("df")
			  Set attr=xmldoc.createAttribute("Font")
			  attr.Text ="Arial Unicode MS"
			  tmpnode.Attributes.setNamedItem attr

			  tmpnode.appendChild xmldoc.createTextNode (targetstr)

              tuvnode2.appendChild tmpnode

              tunode.appendChild tuvnode1
              tunode.appendChild tuvnode2

			  rootnode.appendChild tunode

  			  rootnode.appendChild CreateUTnodeNew(xmldoc,"]]></target>")
              rootnode.appendChild xmldoc.createTextNode (Chr$(13)+Chr$(10))

  			  rootnode.appendChild CreateUTnodeNew(xmldoc,"<parser>"+parsev+"</parser>")
              rootnode.appendChild xmldoc.createTextNode (Chr$(13)+Chr$(10))

			  rootnode.appendChild CreateUTnodeNew(xmldoc,"<id>"+CStr(i)+"</id>")
              rootnode.appendChild xmldoc.createTextNode (Chr$(13)+Chr$(10))

			  s1=trnlst.String(i).ID & " " & trnlst.String(i).ResType & " " & trnlst.String(i).Type
			  rootnode.appendChild CreateUTnodeNew(xmldoc,"<reference>"+Refine(s1)+"</reference>")
              rootnode.appendChild xmldoc.createTextNode (Chr$(13)+Chr$(10))

              rootnode.appendChild CreateUTnodeNew(xmldoc,"</StringList>")
              rootnode.appendChild xmldoc.createTextNode (Chr$(13)+Chr$(10))

          End If
        End If
      Next i

       CreateTail
       xmldoc.Save oupath
       fsys.DeleteFile oupath

       Set nodelist=xmldoc.getElementsByTagName("Tu")
       If nodelist.length <> 0 Then
          Open oupath For Binary As #22
          output_bom 22
          output_str 22,xmldoc.xml
          Close #22
       End If




      End If



    Next


Next kl
MsgBox "Done"

End Sub
Function PreprocessStr(sourcestrorg,parsev,extname) As String
Dim parsertype As Long
Dim sourcestr As String
Dim tempstr1 As String

sourcestr=sourcestrorg

  If InStr(1,sourcestr,"&lt;")<>0 Or InStr(1,sourcestr,"&amp;")<>0 Or InStr(1,sourcestr,"&gt;")<>0  Then

 If (InStr(1,sourcestr,"<")<>0 Or InStr(1,sourcestr,">")<>0 Or InStr(1,sourcestr,"&")<>0) And (extname="htm" Or extname="html")  Then

	Replace sourcestr,"&lt;/","</hisoft"
 	Replace sourcestr,"&lt;","<hisoft"
 	Replace sourcestr,"/&gt;","hisoft/>"
 	Replace sourcestr,"&gt;","hisoft>"
 	'Replace sourcestr,"&quot;",""""
 	Replace sourcestr,"&apos;","'"
 	Replace sourcestr,"&amp;","&hisoft"
 	parsev="8"

 End If

 If InStr(1,sourcestr,"&lt;")<>0 Or InStr(1,sourcestr,"&amp;")<>0 Or InStr(1,sourcestr,"&gt;")<>0  Then

    Replace sourcestr,"&apos;","hzwsingle"
    Replace sourcestr,"&quot;","hzwdouble"

	parsertype=0
	tempstr1=sourcestr
 	If InStr(1,sourcestr,"&lt;")<>0 And InStr(1,sourcestr,"<")=0 Then Replace sourcestr,"&lt;","<"
 	If tempstr1<>sourcestr Then
		parsertype=parsertype+1
 	End If

	tempstr1=sourcestr
 	If InStr(1,sourcestr,"&gt;")<>0 And InStr(1,sourcestr,">")=0 Then Replace sourcestr,"&gt;",">"
 	If tempstr1<>sourcestr Then
		parsertype=parsertype+2
 	End If

	tempstr1=sourcestr

	Dim tempstr2 As String
	tempstr2=sourcestr
	Replace tempstr2,"&amp;",""

 	If InStr(1,tempstr2,"&lt;")=0 And InStr(1,tempstr2,"&gt;")=0 And InStr(1,tempstr2,"&")=0 Then Replace sourcestr,"&amp;","&"


 	If tempstr1<>sourcestr Then
		parsertype=parsertype+4
 	End If
 	parsertype=parsertype And 7
 	parsev=CStr(parsertype)

    Replace sourcestr,"hzwsingle","&apos;"
    Replace sourcestr,"hzwdouble","&quot;"

 End If

End If

Replace sourcestr,"<![CDATA[","<!\[CDATA\["
Replace sourcestr,"]]>","\]\]>"


Replace sourcestr,ChrW(160),"&nbsp;"

PreprocessStr=sourcestr
End Function

Sub CreateHead()

Dim tmplatepath As String
Dim projectname As String


tmplatepath="C:\Users\Public\Documents\Passolo 2011\Macros\template.ttx"
If fsys.FileExists(tmplatepath)=False Then

  MsgBox "Please modify the path to the file template.ttx"
  End
End If

If fsys.FolderExists(filePath+prj.Languages.Item(kl).LangCode)=False Then fsys.CreateFolder filePath+prj.Languages.Item(kl).LangCode

  oupath=filePath+prj.Languages.Item(kl).LangCode+"\"+prj.Name +"_"+ CStr(trnlst.ListID)+"_"+fsys.GetFileName (trnlst.SourceList.SourceFile) +"_"+CStr(filenum)+"_"+prj.Languages.Item(kl).LangCode+".xml"+".ttx"

  If fsys.FileExists(oupath)=True Then

	MsgBox "Duplicate Files"
	End
  End If

fsys.CopyFile tmplatepath,oupath

  xmldoc.async = False
  xmldoc.validateOnParse = False
  xmldoc.preserveWhiteSpace = True
  xmldoc.resolveExternals = False

  xmldoc.load tmplatepath

  If xmldoc.parseError.errorCode <>0 Then
     MsgBox "The template file was corrupted."
     End
  End If

  Set rootnode=xmldoc.selectSingleNode (".//Raw")





  For m=rootnode.childNodes.length To 1 Step -1


  rootnode.removeChild rootnode.childNodes.Item(m-1)

  Next m

 ' For Each prj In PSL.Projects
 '"<?xml version=""1.0"" encoding=""utf-8""?><dictionary></dictionary>"

 rootnode.appendChild CreateUTnodeNew(xmldoc,"<?xml version=""1.0"" encoding=""utf-8""?>")
 rootnode.appendChild xmldoc.createTextNode (Chr$(13)+Chr$(10))

 rootnode.appendChild CreateUTnodeNew(xmldoc,"<dictionary>")
 rootnode.appendChild xmldoc.createTextNode (Chr$(13)+Chr$(10))

 rootnode.appendChild CreateUTnodeNew(xmldoc,"<project>")
 rootnode.appendChild xmldoc.createTextNode (Chr$(13)+Chr$(10))

 projectname="<projectname>"+prj.Name+"</projectname>"
 rootnode.appendChild CreateUTnodeNew(xmldoc,projectname)
 rootnode.appendChild xmldoc.createTextNode (Chr$(13)+Chr$(10))

 rootnode.appendChild CreateUTnodeNew(xmldoc,"<language>")
 rootnode.appendChild xmldoc.createTextNode (Chr$(13)+Chr$(10))

 projectname="<languagename>"+CStr(prj.Languages.Item(kl).LangCode )+"</languagename>"
 rootnode.appendChild CreateUTnodeNew(xmldoc,projectname)
 rootnode.appendChild xmldoc.createTextNode (Chr$(13)+Chr$(10))

 rootnode.appendChild CreateUTnodeNew(xmldoc,"<files>")
 rootnode.appendChild xmldoc.createTextNode (Chr$(13)+Chr$(10))

 projectname="<file><filename>"+Refine(trnlst.SourceList.SourceFile)+"</filename>"
 rootnode.appendChild CreateUTnodeNew(xmldoc,projectname)
 rootnode.appendChild xmldoc.createTextNode (Chr$(13)+Chr$(10))

 rootnode.appendChild CreateUTnodeNew(xmldoc,"<StringLists>")



End Sub
Sub CreateTail()
 rootnode.appendChild CreateUTnodeNew(xmldoc,"</StringLists>")
 rootnode.appendChild xmldoc.createTextNode (Chr$(13)+Chr$(10))

 rootnode.appendChild CreateUTnodeNew(xmldoc,"</file>")
 rootnode.appendChild xmldoc.createTextNode (Chr$(13)+Chr$(10))

 rootnode.appendChild CreateUTnodeNew(xmldoc,"</files>")
 rootnode.appendChild xmldoc.createTextNode (Chr$(13)+Chr$(10))

 rootnode.appendChild CreateUTnodeNew(xmldoc,"</language>")
 rootnode.appendChild xmldoc.createTextNode (Chr$(13)+Chr$(10))

 rootnode.appendChild CreateUTnodeNew(xmldoc,"</project>")
 rootnode.appendChild xmldoc.createTextNode (Chr$(13)+Chr$(10))

 rootnode.appendChild CreateUTnodeNew(xmldoc,"</dictionary>")
 rootnode.appendChild xmldoc.createTextNode (Chr$(13)+Chr$(10))

End Sub

Function CreateUTnodeNew(xmldoc As DOMDocument40, tmpstr) As IXMLDOMNode
Dim utnode As IXMLDOMNode
Dim dfg As IXMLDOMAttribute

Set utnode = xmldoc.createElement("ut")
Set dfg = xmldoc.createAttribute("Style")
dfg.Text = "external"
utnode.Attributes.setNamedItem dfg

utnode.Text = tmpstr

Set CreateUTnodeNew = utnode

End Function
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
  Dim n As Long

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

Sub output_bom(fnum As Long)
Dim bb As Byte

bb = 255
Put #fnum, , bb
bb = 254
Put #fnum, , bb
End Sub

Sub output_str(fnum As Long, Str)
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
If Action=1 Then
	DlgValue "Group1", 0
    DlgValue "Group2", 0
    option1=True
	option3=True
	option4=False
	option5=False
	option6=False

ElseIf Action=2 Then

  If DlgItem = "Group1" Then

		Select Case DlgValue("Group1")

		Case 0
			option1=True
		Case 1
			option1=False

       End Select

	ElseIf DlgItem = "Group2" Then

            option7=False
			option3=False
			option4=False
			option5=False
			option6=False

		Select Case DlgValue("Group2")
		Case 0
			option3=True

		Case 1
			option4=True
		Case 2
			option5=True
		Case 3
			option6=True
        Case 4
            option7=True
       End Select

	End If

End If


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
