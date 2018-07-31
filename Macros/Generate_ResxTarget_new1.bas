Attribute VB_Name = "passolo_macro_extract"
'#Reference {F5078F18-C551-11D3-89B9-0000F81FE221}#5.0#0#C:\Program Files\Common Files\Microsoft Shared\OFFICE11\MSXML5.DLL#Microsoft XML, v5.0
'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\WINDOWS\system32\scrrun.dll#Microsoft Scripting Runtime
Option Explicit

'Define Constant
Dim CsvName As String      'CSV file name
Dim bComment As Boolean
Dim bTransDate As Boolean
Dim bResource As Boolean
Dim errorcount As Integer
Dim fnout As Integer
Dim proplist(15) As String
Dim proplist_len As Integer

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
  Dim strid As String, sourcestr As String, targetstr As String, filename As String, sourcefilename As String
  Dim fsys As New FileSystemObject
  Dim dumy As Integer
  Dim x1 As Long,y1 As Long , w1 As Long ,h1 As Long
  Dim x2 As Long,y2 As Long , w2 As Long ,h2 As Long
  Dim idsize As String ,idlocation As String
  Dim xmldoc As New DOMDocument50
  Dim tgtdoc As New DOMDocument50
  Dim nodelist As IXMLDOMNodeList
  Dim snodelist As IXMLDOMNodeList
  Dim connstr As String
  Dim foldn As String
  Dim m As Integer
  Dim n As Integer
  Dim bnode As IXMLDOMNode, tnode As IXMLDOMNode
  Dim node As IXMLDOMNode
  Dim svnode As IXMLDOMNode
  Dim tvnode As IXMLDOMNode
  Dim sna As String ,na As String, tpre As String , spre As String

  Set prj = PSL.ActiveProject
  Dim nn As Long , mm As Long , zquot As String

  ' Put put file path
  filePath = prj.Location
  If Right(filePath, 1) <> "\" Then
    filePath = filePath & "\"
  End If

  InitPropList

  Open filePath & "text.log" For Output As #2

  Open filePath & "xml.log" For Output As #1

  Open filePath & "Att.log" For Output As #3

  Open filePath & "addnouse.log" For Output As #4

  Open filePath & "adduseful.log" For Output As #5

  Open filePath & "tablelayou.log" For Output As #6

  Open filePath & "sizeexcep.log" For Output As #7

  Open filePath & "generated.log" For Output As #8

   For Each trnlst In prj.TransLists

      'Only process the selected files
      If trnlst.Selected =True  Then
'And trnlst.Language.LangCode = "jpn"

      'If trnlst.Property ("M:mark")="drop3" Then


      filename = trnlst.TargetFile
      fnout=0

      'Only process resx files
      If LCase(fsys.GetExtensionName (filename))<>"resx" Then

         'Generate the target file
         If trnlst.GenerateTarget=False Then
            Print #8,filename +" is not generated!"
         End If

         If LCase(fsys.GetExtensionName (filename))="vmsg" Then
            process_vmsg filename
         End If

      Else

      'Generate the target file
      trnlst.GenerateTarget

      'Get the source file name
      sourcefilename=trnlst.SourceList.SourceFile

      'Copy the source file to the localized file
      'fsys.CopyFile sourcefilename,filename


      'Load the translated XML file
      xmldoc.async = False
      xmldoc.validateOnParse = False
      xmldoc.preserveWhiteSpace = True
      xmldoc.resolveExternals = False
      xmldoc.load (filename)

      If xmldoc.parseError.errorCode <>0 Then
         Print #1, filename
         GoTo skip
      End If

      'Load the source XML file
      tgtdoc.async = False
      tgtdoc.validateOnParse = False
      tgtdoc.preserveWhiteSpace = True
      tgtdoc.resolveExternals = False
      tgtdoc.load (sourcefilename)

      tgtdoc.Save filename+".srs"

      If xmldoc.parseError.errorCode <>0 Then
         Print #1, sourcefilename
         GoTo skip
      End If

      'Get the node list whose parent is root and which has a attribute named "name"
      'Set the namespace

      sna = tgtdoc.documentElement.namespaceURI
      If sna <> "" Then
          na = "xmlns:na='" + sna + "'"
          tgtdoc.setProperty "SelectionNamespaces", na
          tpre = "na:"
      Else
          tpre = ""
      End If


      sna = xmldoc.documentElement.namespaceURI
      If sna <> "" Then
          na = "xmlns:na='" + sna + "'"
          xmldoc.setProperty "SelectionNamespaces", na
          spre = "na:"
      Else
         spre = ""
      End If

      connstr = "/" + tpre + "root/*[@name !='']"

      Set nodelist = tgtdoc.selectNodes(connstr)

      'Log the file which contains tablelayout control
      IdentifyTable nodelist, filename


      For n= nodelist.length To 1 Step -1

          Set node=nodelist.Item (n-1)

          'Get the value of the name attibute
          Dim tmpstr As String
          tmpstr=node.Attributes.getNamedItem ("name").Text


          'Only process the node whose name attribute's value is not equal to "$this.Language"
          If tmpstr<>"$this.Language" And tmpstr<>"$this.RightToLeft" Then 'Modified on 6-18-2009

          'Get the node list from the target file
          connstr = "/" + spre + "root/*[@name =""" & tmpstr & """]"
          Set snodelist = xmldoc.selectNodes(connstr)

          If snodelist.length <>0 Then

             'Log the duplicated node
             If snodelist.length <>1 Then

                w1=0

                For m=1 To snodelist.length
                   Set node=snodelist.Item (m-1)
                   Print #2, "Duplicate Node" & " " & node.nodeName
                Next m


             Else

             'A unique node is found

               Set bnode=snodelist.Item (0)

               FindOutException bnode, node, spre, tpre,filename

               x1=bnode.Attributes.length
               x2=node.Attributes.length

               Dim type1 As String ,type2 As String

               'Get the value of the attribute type if possible
               type1=""
               For m=1 To x1

                  If LCase(bnode.Attributes.Item(m-1).nodeName)="type" Then type1=bnode.Attributes.Item(m-1).Text
               Next m

               type2=""

               For m=1 To x2

                  If LCase(node.Attributes.Item(m-1).nodeName)="type" Then type2=node.Attributes.Item(m-1).Text
               Next m

               If type1<>type2 Then
                  'Log the element which type attribute does not match

                  Print #3, filename

                  s1=""
                  For m=1 To x1
                      s1=s1 & bnode.Attributes.Item(m-1).nodeName & "=" & bnode.Attributes.Item(m-1).Text & " "
                  Next m
                  Print #3,s1

                  s1=""
                  For m=1 To x2
                      s1=s1 & node.Attributes.Item(m-1).nodeName & "=" & node.Attributes.Item(m-1).Text & " "
                  Next m
                  Print #3, s1

                  Print #3,""


               Else
                'Copy the value node from the target to source

                h1=InStr(1,node.nodeName,"xsd:")

                If h1=0 And node.nodeName <>"resheader" Then

                   'Get the value element, and remove all of its children
                   Set svnode=node.selectSingleNode (tpre + "value")
                   Dim cdv As Integer

                   cdv=0
                   If Not(svnode Is Nothing) Then

                      If InStr(1,svnode.xml,"![CDATA[")<>0 Then

                         If svnode.childNodes.length =1 Then

                            If svnode.childNodes.Item(0).nodeType =4  Then
                               Set svnode=svnode.childNodes.Item (0)
                               cdv=1
                            End If
                         Else
                            MsgBox "Exceptional CDATA"
                         End If
                      End If

                      For m=svnode.childNodes.length To 1 Step -1
                          svnode.removeChild svnode.childNodes.Item (m-1)
                      Next m

                   End If

                    Set tvnode=bnode.selectSingleNode (spre + "value")

                    'Copy translated node to target file
 				    If Not(tvnode Is Nothing ) And Not(svnode Is Nothing) Then
 				       If InStr(1,tvnode.xml,"![CDATA[")<>0 Then
 				          MsgBox "Error with CDATA
 				          cdv=0
 				       End If

 				       If cdv=1 Then
                          svnode.Text =tvnode.Text
 				       Else

                       For m=1 To tvnode.childNodes.length

                           svnode.appendChild tvnode.childNodes.Item(m-1).cloneNode (True )

                       Next m

                       End If
				    End If
				ElseIf node.nodeName <>"resheader" Then
				   MsgBox "Other node is ignored
                End If

                'Remove the processed node from the translated file
                bnode.parentNode.removeChild bnode

             End If

            End If



          End If

        End If

      Next n

      'Process the added node in the translated file
      connstr = "/" + spre + "root/*[@name !='']"
      Set nodelist = xmldoc.selectNodes(connstr)


      For n=1 To nodelist.length

          Set node=nodelist.Item (n-1)

          'Dim tmpstr As String

          tmpstr=node.Attributes.getNamedItem ("name").Text

          'Look for the same node in the target file
          connstr = "//*[@name ='" & tmpstr & "']"
          Set snodelist = tgtdoc.selectNodes(connstr)

          If snodelist.length =0 Then

             If node.nodeName <>"data" Then
                Print #4, filename
                Print #4, node.xml
             Else

               'Maybe append the node to the target file
                Print #5, filename
                Print #5, node.xml
             End If
          End If

      Next n

      UpdateLinkArea trnlst,tgtdoc

      tgtdoc.Save (filename)

      End If



      End If 'Selected = True

skip:

    Next


Close #1
Close #2
Close #3
Close #4
Close #5
Close #6
Close #7
Close #8

Open filePath & "generated.log" For Binary As #8

Dim flenas As Integer

flenas=LOF(8)
Close #8

If flenas<>0 Then
   MsgBox "some files are not generated! Please see "+ filePath & "generated.log"
End If

Set fsys=Nothing

MsgBox "Done"

End Sub
Sub HasLinkArea(tgtdoc As DOMDocument50,idval As String,vals As String  )

Dim sna As String ,na As String, tpre As String , spre As String
Dim snodelist As IXMLDOMNodeList
Dim node As IXMLDOMNode
Dim connstr As String

Dim bnode As IXMLDOMNode



      sna = tgtdoc.documentElement.namespaceURI
      If sna <> "" Then
          na = "xmlns:na='" + sna + "'"
          tgtdoc.setProperty "SelectionNamespaces", na
          tpre = "na:"
      Else
          tpre = ""
      End If

      connstr = "//*[@name ='" & idval & "']"
      Set snodelist = tgtdoc.selectNodes(connstr)

      If snodelist.length =1 Then
         Set node=snodelist.Item(0)
         Set bnode=node.selectSingleNode (tpre+"value")
         bnode.Text =vals

      End If






End Sub
Function UpdateLinkArea(trnlst As PslTransList,tgtdoc As DOMDocument50  )
Dim i As Integer
Dim m As Integer
Dim idval As String
Dim vals As String
Dim mm As Integer
Dim props As String

	For i = 1 To trnlst.StringCount
          'For m=1 To trnlst.String(i).Properties.Count
        For mm=1 To proplist_len

          props=proplist(mm)
          vals=trnlst.String(i).Property("M:"+props)
          If vals<>"" Then
             idval=trnlst.String(i).ID +"."+props
             HasLinkArea tgtdoc,idval,vals
          End If
        Next mm


    Next i



End Function

Sub FindOutException (bnode As IXMLDOMNode , node As IXMLDOMNode, spre As String, tpre As String, filename As String   )
	Dim tmpstr1,tmpstr2 As String


	tmpstr1=bnode.Attributes.getNamedItem("name").Text
	tmpstr2=node.Attributes.getNamedItem("name").Text

	If (InStr(1,LCase(tmpstr1),".location")<>0 And InStr(1,LCase(tmpstr2),".location""")<>0) Then
		DecideValue bnode,node,spre,tpre, filename,1,tmpstr1
	ElseIf  InStr(1,LCase(tmpstr1),".size")<>0 And InStr(1,LCase(tmpstr2),".size""")<>0 Then
        DecideValue bnode,node,spre,tpre,filename,0,tmpstr1
    End If
End Sub
Sub DecideValue(bnode As IXMLDOMNode , node As IXMLDOMNode, spre As String, tpre As String ,filename As String, flag As Integer, namev As String  )
	Dim tvnode As IXMLDOMNode
	Dim svnode As IXMLDOMNode
	Dim taxis() As String
	Dim saxis() As String
	Dim t1,s1 As String
    Dim tlen, slen As Integer


    Set tvnode=bnode.selectSingleNode (spre + "value")
    Set svnode=node.selectSingleNode (spre + "value")
    s1=svnode.Text
    t1=tvnode.Text

    taxis=Split(t1,",")
    saxis=Split(s1,",")

    tlen=UBound(taxis)
    slen=UBound(saxis)

    If slen=tlen Then

       If flag=1 Then
          If Abs(CInt(taxis(0))-CInt(saxis(0)))> 10 Or Abs(CInt(taxis(0))-CInt(saxis(0)))> 5 Then
             If fnout=0 Then
                Print #7, filename
                Print #7,"----------------------------------------------------------------------------"
                fnout=1
             End If
             Print #7,namev & ": The location Is changed from " & CStr(saxis(0))+"," & CStr(saxis(1)) & " To " & CStr(taxis(0))+"," & CStr(taxis(1))
             Print #7,""

          End If

		  'If Abs(CInt(taxis(0))-CInt(saxis(0)))> 5 Then Print #7, filename +" The change in top is more than 5"
       Else
          If Abs(CInt(taxis(0))-CInt(saxis(0)))> 30 Then
             If fnout=0 Then
                Print #7, filename
                Print #7,"----------------------------------------------------------------------------"
                fnout=1
             End If
             Print #7,namev & ": The size is changed from " & CStr(saxis(0))+"," & CStr(saxis(1)) & " to " & CStr(taxis(0))+"," & CStr(taxis(1))
             Print #7,""
          End If

		  'If Abs(CInt(taxis(0))-CInt(saxis(0)))> 2 Then Print #7, filename +" The change in height is more than 2"
       End If

    End If

End Sub

Sub IdentifyTable (nodelist As IXMLDOMNodeList , filename As String )
Dim i As Integer
Dim node As IXMLDOMNode
Dim tmpstr As String

For i=1 To nodelist.length

   Set node=nodelist.Item (i-1)

   tmpstr=node.Text

   If InStr(1,LCase(tmpstr),LCase(".TableLayoutPanel"))<>0 Then
      Print #6, filename
      Exit Sub
   End If
Next i

End Sub

Sub CreateF(foldn As String)
Dim tmpstr As String
Dim pos1 As Long
Dim fsys As New FileSystemObject
Do

  pos1=InStr(1,foldn,"\")
  If pos1<>0 Then
     tmpstr=tmpstr+Left(foldn,pos1)
     foldn=Mid(foldn,pos1+1)
     If fsys.FolderExists (tmpstr)= False Then fsys.CreateFolder tmpstr
  Else
     'tmpstr=tmpstr+foldn
     'If fsys.FolderExists (tmpstr)= False Then fsys.CreateFolder tmpstr
  End If

Loop Until pos1=0

End Sub

Function updatenode(xmldoc As DOMDocument50,idsize As String, Val As String  )
Dim nodelist As IXMLDOMNodeList
Dim connstr As String
Dim nm As Long
Dim nn As Long

Dim bnode As IXMLDOMNode, tnode As IXMLDOMNode
Dim node As IXMLDOMNode
Dim t1 As String
Dim pos1 As Integer

t1=idsize
pos1=InStr(1,t1,".")
If pos1<>0 Then
   t1=Left(t1,pos1-1)
   t1=t1 & ".ClientSize"
End If

If t1<>"" Then
   connstr = "//*[@name ='" + idsize + "' or @name ='" + idsize + ".Text'  or @name ='" + t1 + "' ]"
Else
  connstr = "//*[@name ='" + idsize + "' or @name ='" + idsize + ".Text']"
End If
Set nodelist = xmldoc.selectNodes(connstr)

If nodelist.length =1 Then

   For nm=1 To nodelist.length

       Set bnode=nodelist.Item (nm-1)
       Set tnode=bnode.selectSingleNode ("value")
       If tnode Is Nothing Then MsgBox "Value element is not found"
       If tnode.childNodes.length>1 Or tnode.childNodes.length=0 Then MsgBox "exception"

       For nn=1 To tnode.childNodes.length
           Set node=tnode.childNodes.Item (nn-1)
           If node.nodeType=4 Then
              node.Text =Val
           Else
              node.Text =Val
           End If
       Next nn

   Next nm
 ElseIf nodelist.length =2 Then
     MsgBox "Duplicate"
 Else
     MsgBox "Not found"
 End If
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

Sub process_vmsg(filename As String )
	Dim fsys As New FileSystemObject
	Dim bb(3) As Byte

  If fsys.FileExists (filename)=False Then
     MsgBox "Parser info is missed! Please restore and try again!"
     End
  End If
	Open filename For Binary As #41
    Get #41,,bb(1)
    Get #41,,bb(2)
    Get #41,,bb(3)
    If bb(1)=239 And bb(2)=187  And bb(3)=191 Then
       If fsys.FileExists (filename+"1")= True Then fsys.DeleteFile filename+"1"
       Open filename+"1" For Binary As #51
       While Loc(41)<LOF(41)

         Get #41,,bb(1)
         Put #51,,bb(1)
       Wend
       Close #51
       Close #41

       fsys.CopyFile filename+"1",filename
       fsys.DeleteFile filename+"1"
    Else
    	Close #41
    End If


End Sub


Sub InitPropList()
proplist_len=2

proplist(1)="LinkArea"
proplist(2)="Width"
'proplist(3)="Location"
'proplist(4)="Size"



End Sub
