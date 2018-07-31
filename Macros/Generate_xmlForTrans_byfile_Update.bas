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
Dim errorcount As Long

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


Sub Main()

	Begin Dialog UserDialog 800,196,"Export to XML",.TMAutotrans ' %GRID:10,7,1,1
		Text 20,35,260,21,"Please specify destination:",.Text1
		TextBox 30,63,650,21,.tmname
		OKButton 300,119,60,21,.d
		CancelButton 440,119,60,21
		PushButton 700,63,50,21,"...",.SelectFile
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

        If s1 <> "" Then
          
          If trnlst.String(i).ResType <> "Version" And trnlst.String(i).ResType <>"DialogFont" And trnlst.String(i).Resource.State(pslStateReadOnly)=False  And  trnlst.String(i).State(pslStateReadOnly)=False And _
          trnlst.String(i).State(pslStateTranslated)=False  And _
          trnlst.String(i).State(pslStateHidden)=False Then

              itemnum=itemnum+1

              If itemnum>1000 Then

                 projnode.appendChild langnode
                 projnode.appendChild txtnode.cloneNode (True )
                 xmldoc.Save oupath

                 itemnum=0
                 filenum=filenum+1
                 CreateHead

              End If

    	      sourcestr = trnlst.String(i).SourceText
              If sourcestr="manual" Then
                 sourcestr="manual"
              End If 
    	      Replace sourcestr,"\r","\\r"
    	      Replace sourcestr,"\n","\\n"
    	      Replace sourcestr,Chr$(13),"\r"
              Replace sourcestr,Chr$(10),"\n"

        	  targetstr =trnlst.String(i).SourceText 'trnlst.String(i).Text
    	      Replace targetstr,"\r","\\r"
    	      Replace targetstr,"\n","\\n"
    	      Replace targetstr,Chr$(13),"\r"
              Replace targetstr,Chr$(10),"\n"


        	  Dim temppnode As IXMLDOMNode

        	  Set temppnode=xmldoc.createElement ("StringList")
        	  Set snode=xmldoc.createElement ("source")
              Set tnode=xmldoc.createElement ("target")

              Dim parsernode As IXMLDOMNode

              Set parsernode=xmldoc.createElement ("parser")

              Set idnode=xmldoc.createElement ("id")

              Dim refnode As IXMLDOMNode

              Set refnode=xmldoc.createElement ("reference")

              refnode.Text = trnlst.String(i).ID & " " & trnlst.String(i).ResType & " " & trnlst.String(i).Type
              snode.Text =sourcestr
              'tnode.Text=targetstr
              idnode.Text =CStr(i)

              Dim parsev As String
              parsev="0"
              extname=LCase(fsys.GetExtensionName(trnlst.SourceList.SourceFile))

              If InStr(1,targetstr,"&lt;")<>0 Or InStr(1,targetstr,"&amp;")<>0 Or InStr(1,targetstr,"&gt;")<>0  Then

              	 If (InStr(1,targetstr,"<")<>0 Or InStr(1,targetstr,">")<>0 Or InStr(1,targetstr,"&")<>0) And (extname="htm" Or extname="html")  Then

					Replace targetstr,"&lt;/","</hisoft"
                 	Replace targetstr,"&lt;","<hisoft"
                 	Replace targetstr,"/&gt;","hisoft/>"
                 	Replace targetstr,"&gt;","hisoft>"
                 	'Replace targetstr,"&quot;",""""
                 	'Replace targetstr,"&apos;","'"
                 	Replace targetstr,"&amp;","&hisoft"
                 	parsev="8"

              	 End If

              	 If InStr(1,targetstr,"&lt;")<>0 Or InStr(1,targetstr,"&amp;")<>0 Or InStr(1,targetstr,"&gt;")<>0  Then

                    Replace targetstr,"&apos;","hzwsingle"
                    Replace targetstr,"&quot;","hzwdouble"

					parsertype=0
					tempstr1=targetstr
                 	If InStr(1,targetstr,"&lt;")<>0 And InStr(1,targetstr,"<")=0 Then Replace targetstr,"&lt;","<"
                 	If tempstr1<>targetstr Then
                		parsertype=parsertype+1
                 	End If

					tempstr1=targetstr
                 	If InStr(1,targetstr,"&gt;")<>0 And InStr(1,targetstr,">")=0 Then Replace targetstr,"&gt;",">"
                 	If tempstr1<>targetstr Then
						parsertype=parsertype+2
                 	End If

					tempstr1=targetstr

					Dim tempstr2 As String
					tempstr2=targetstr
					Replace tempstr2,"&amp;",""

                 	If InStr(1,tempstr2,"&lt;")=0 And InStr(1,tempstr2,"&gt;")=0 And InStr(1,tempstr2,"&")=0 Then Replace targetstr,"&amp;","&"


                 	If tempstr1<>targetstr Then
						parsertype=parsertype+4
                 	End If
                 	parsertype=parsertype And 7
                 	parsev=CStr(parsertype)

                    Replace targetstr,"hzwsingle","&apos;"
                    Replace targetstr,"hzwdouble","&quot;"

                 End If

              End If

              Replace targetstr,"<![CDATA[","<!\[CDATA\["
			  Replace targetstr,"]]>","\]\]>"

              tnode.appendChild xmldoc.createCDATASection (targetstr)

              parsernode.Text=parsev

              temppnode.appendChild txtnode.cloneNode (True )
              temppnode.appendChild snode
              temppnode.appendChild txtnode.cloneNode (True )
              temppnode.appendChild tnode
              temppnode.appendChild txtnode.cloneNode (True )
              temppnode.appendChild parsernode
              temppnode.appendChild txtnode.cloneNode (True )


              temppnode.appendChild idnode
              temppnode.appendChild txtnode.cloneNode (True )
              temppnode.appendChild refnode
              temppnode.appendChild txtnode.cloneNode (True )

              strsnode.appendChild temppnode
              strsnode.appendChild txtnode.cloneNode (True )

          End If
        End If
      Next i

       projnode.appendChild langnode
       projnode.appendChild txtnode.cloneNode (True )

       Dim nodelist As IXMLDOMNodeList

       Set nodelist=xmldoc.getElementsByTagName("StringList")

       If nodelist.length <>0 Then  xmldoc.Save oupath


      End If



    Next


Next kl
MsgBox "Done"

End Sub
Sub CreateHead()


  If fsys.FolderExists(filePath+prj.Languages.Item(kl).LangCode)=False Then fsys.CreateFolder filePath+prj.Languages.Item(kl).LangCode

  Dim gyh As String
  Dim gyh_Arry() As String
  Dim Arry_length As Integer
  Dim gyh_int As Integer

  gyh_int = 0
  gyh = fsys.GetFileName (trnlst.SourceList.SourceFile)
  gyh_int = InStr (gyh_int+1, gyh, ";")

  If gyh_int <>0 Then

  gyh_Arry = Split(gyh,";")
  Arry_length = UBound(gyh_Arry) - LBound(gyh_Arry) + 1
  'gyh_int = InStr (gyh_int+1, gyh, ";")

  If Arry_length > 3 Then
  gyh = gyh_Arry(UBound(gyh_Arry)-2) +";"+ gyh_Arry(UBound(gyh_Arry)-1)+";"+gyh_Arry(UBound(gyh_Arry))

  End If

  End If


  oupath=filePath+prj.Languages.Item(kl).LangCode+"\"+ CStr(trnlst.ListID)+"_"+ gyh +"_"+CStr(filenum)+"_"+prj.Languages.Item(kl).LangCode+".xml"

  xmldoc.async = False
  xmldoc.validateOnParse = False
  xmldoc.preserveWhiteSpace = True
  xmldoc.resolveExternals = False

     xmldoc.loadXML "<?xml version=""1.0"" encoding=""utf-8""?><dictionary></dictionary>"


  Set txtnode=xmldoc.createTextNode  (Chr$(13)+Chr$(10))

  If xmldoc.parseError.errorCode <>0 Then
     MsgBox "File not well Formed"
     End
  End If

  Set rootnode=xmldoc.selectSingleNode ("dictionary")



  For m=rootnode.childNodes.length To 1 Step -1


  rootnode.removeChild rootnode.childNodes.Item(m-1)

  Next m

 ' For Each prj In PSL.Projects


   Set projnode=xmldoc.createElement ("project")

   rootnode.appendChild txtnode.cloneNode (True )
   rootnode.appendChild projnode
   rootnode.appendChild txtnode.cloneNode (True )

   Set projnamenode=xmldoc.createElement ("projectname")
   projnamenode.Text =prj.Name

   projnode.appendChild txtnode.cloneNode (True )
   projnode.appendChild projnamenode
   projnode.appendChild txtnode.cloneNode (True )

   Set langnode=xmldoc.createElement ("language")
   langnode.appendChild txtnode.cloneNode (True )



   Set langnamenode=xmldoc.createElement ("languagename")

   langnamenode.Text =CStr(prj.Languages.Item(kl).LangCode )


   langnode.appendChild langnamenode
   langnode.appendChild txtnode.cloneNode (True )


   Set filesnode=xmldoc.createElement ("files")

   langnode.appendChild filesnode
   langnode.appendChild txtnode.cloneNode (True )

        Set filenode=xmldoc.createElement ("file")

        filesnode.appendChild txtnode.cloneNode (True )
        filesnode.appendChild filenode
        filesnode.appendChild txtnode.cloneNode (True )

        Set filenamenode=xmldoc.createElement ("filename")
        filenamenode.Text=trnlst.SourceList.SourceFile


        filenode.appendChild filenamenode
        filenode.appendChild txtnode.cloneNode (True )


        Set strsnode=xmldoc.createElement ("StringLists")

        filenode.appendChild strsnode
        filenode.appendChild txtnode.cloneNode (True )



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
