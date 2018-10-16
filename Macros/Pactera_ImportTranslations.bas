Attribute VB_Name = "passolo_macro_import"
'#Reference {F5078F18-C551-11D3-89B9-0000F81FE221}#6.0#0#C:\Windows\System32\msxml6.dll#Microsoft XML, v6.0
'#Reference {F5078F18-C551-11D3-89B9-0000F81FE221}#4.0#0#C:\Windows\SysWow64\msxml4.dll#Microsoft XML, v4.0
'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\SysWOW64\scrrun.dll#Microsoft Scripting Runtime
Option Explicit

'variables for Dialog
Dim optionXml As Boolean
Dim optionXlf As Boolean
'fso variables
Dim fsys As New FileSystemObject
Dim mainFolderPath As String, folderPath As String
Dim fd As Folder, fl As File
Dim extName As String

'Passolo variables
Dim prj As PslProject
Dim srcLangId As Long, tgtLangId As Long
Dim srcLangCode$, tgtLangCode$  'lang-country code, e.g. zh-CN
Dim trnLst As PslTransList
Dim trnStr As PslTransString, trnLstId As Long

Dim srcStr As String, tgtStr As String, srcFilePath$
Dim prjName$,srcFile$, strNum$
Dim isReadOnly As Boolean, isValidated As Boolean
Dim isLocked As Boolean, isInAttr As Boolean

'xml variables
Dim xmlDoc As New DOMDocument40
Dim xmlxmlRootNode As IXMLDOMElement, xmlProjNode As IXMLDOMElement
Dim xmlStrNode As IXMLDOMElement

'pactXliff variables
Dim xlfDoc As New DOMDocument40
Const xlfNS As String = "urn:oasis:names:tc:xliff:document:1.2"
Const pactNS As String = "pgs.pactera.com"
Dim xlfTransUnitNode As IXMLDOMElement

'Dictionary varibles
Dim dict As New Dictionary
Dim key$, value$

'other variables
Dim delimeter As String

Sub Main()
  delimeter = ChrW(9478)

   Begin Dialog UserDialog 770,245,"Import from XML/Xliff",.MainDialog ' %GRID:10,7,1,1
      Text 20,20,180,14,"Please specify translated Folder:",.Text1
      TextBox 20,42,650,21,.FolderPath
      OKButton 200,200,70,21
      CancelButton 400,200,90,21
      PushButton 700,42,50,21,"...",.SelectFile
      GroupBox 640,80,110,80,"File Format",.GroupBox3
      OptionGroup .Group3
        OptionButton 660,110,60,14,"Xml",.OptionXml
        OptionButton 660,130,80,14,"pactXliff",.OptionXlf
   End Dialog
  
  Dim dlg As UserDialog
  If Dialog(dlg) = 0 Then Exit Sub

  mainFolderPath = dlg.FolderPath
  Set fd = fsys.GetFolder(mainFolderPath)
  PSL.Output "Phase I: Loading translated files from: " + mainFolderPath
  processFolder fd
  Autotrans
End Sub

Private Sub processFolder(fd As Folder)
  Dim sbFd As Folder
  For Each fl In fd.Files
    extName = fsys.GetExtensionName(fl.Name)
    If optionXml And LCase(extName)="xml" Then
      PutXmlIntoDict fl.Path
    ElseIf optionXlf And LCase(extName)="pactxliff" Then 
      PutXlfIntoDict fl.Path
    End If
  Next fl

  For Each sbFd In fd.SubFolders
    processFolder sbFd
  Next sbFd
End Sub

Sub PutXmlIntoDict(filePath As String)
  xmlDoc.async = False
  xmlDoc.validateOnParse = False
  xmlDoc.preserveWhiteSpace = True
  xmlDoc.resolveExternals = False

  xmlDoc.load filePath

  If xmlDoc.parseError.errorCode <>0 Then
    PSL.Output filePath + " is not well-formed."
    End
  End If

  'get KEY
  prjName = xmlDoc.selectSingleNode("./project/@name").Text
  srcFile = xmlDoc.selectSingleNode(".//file/@srcFile").Text
  tgtLangCode = xmlDoc.selectSingleNode(".//file/@target-language").Text

  For Each xmlStrNode In xmlDoc.selectNodes(".//string")
    strNum = xmlStrNode.selectSingleNode("./@strNum").Text
    isInAttr = CBool(xmlStrNode.selectSingleNode("./@InXmlAttribute").Text)
    srcStr = xmlStrNode.selectSingleNode("./source").Text
    tgtStr = xmlStrNode.selectSingleNode("./target").Text
    
    If isInAttr Then
      srcStr = PostProcess(srcStr)
      tgtStr = PostProcess(tgtStr)
    End If

    key = GetKey()
    AddItemToDict key, tgtStr
  Next xmlStrNode
End Sub

Sub PutXlfIntoDict(filePath As String)
  xlfDoc.async = False
  xlfDoc.validateOnParse = False
  xlfDoc.preserveWhiteSpace = True
  xlfDoc.resolveExternals = False
  xlfDoc.setProperty "SelectionNamespaces", "xmlns:pact=""" + pactNS + """ xmlns:xlf=""" + xlfNS + """"

  xlfDoc.load filePath

  If xlfDoc.parseError.errorCode <>0 Then
    PSL.Output filePath + " is not well-formed."
    End
  End If

  'get KEY
  prjName = xlfDoc.selectSingleNode(".//xlf:file/@pact:prjName").Text
  srcFile = xlfDoc.selectSingleNode(".//xlf:file/@pact:srcFile").Text
  tgtLangCode = xlfDoc.selectSingleNode(".//xlf:file/@target-language").Text

  For Each xlfTransUnitNode In xlfDoc.selectNodes(".//xlf:trans-unit")
    strNum = xlfTransUnitNode.selectSingleNode("./@pact:strNum").Text
    isInAttr = CBool(xlfTransUnitNode.selectSingleNode("./@pact:InXmlAttribute").Text)
    srcStr = xlfTransUnitNode.selectSingleNode("./xlf:source").Text
    tgtStr = xlfTransUnitNode.selectSingleNode("./xlf:target").Text
    
    If isInAttr Then
      srcStr = PostProcess(srcStr)
      tgtStr = PostProcess(tgtStr)
    End If

    key = GetKey()
    AddItemToDict key, tgtStr
  Next xlfTransUnitNode
End Sub

Private Function PostProcess(src$) As String
  If isInAttr Then
    src = Replace(src, "&lt;", "<")
    src = Replace(src, "&gt;", ">")
    src = Replace(src, "&quot;", """")
    src = Replace(src, "&apos;", "'")
    src = Replace(src, "&amp;", "&")
  End If
  PostProcess = src
End Function

' Add or update an item in the dictionary
Private Sub AddItemToDict(key As String, value As String)
  ' if the item does not exist in the dictionary already
  If Not dict.Exists(key) Then
    ' add new item to the dictionary
    dict.Add(key, value)
  Else
    'update
    dict.Item(key) = value
  End If
End Sub


Function LookUpDictItem(key As String) As Boolean
  LookUpDictItem = dict.Exists(key)
  If LookUpDictItem Then value = dict.Item(key)
End Function

Sub Autotrans()
  PSL.Output "Phase II - Get the translation"
  Set prj = PSL.ActiveProject
  prjName = prj.Name
  For Each trnLst In prj.TransLists
    If trnLst.Selected = False Then GoTo NextTrnLst

    tgtLangId = trnLst.Language.LangID
    tgtLangCode = PSL.GetLangCode(tgtLangId, pslCodeLangRgn)
    srcFilePath = Replace(trnLst.SourceList.SourceFile, prj.Location, "")

    For trnLstId = 1 To trnLst.StringCount
      Set trnStr = trnLst.String(trnLstId)
      srcStr = Trim(trnStr.SourceText)
      tgtStr = Trim(trnStr.Text)
      If srcStr = "" And tgtStr <> "" Then
        trnStr.Text = trnStr.SourceText

        GoTo NextTrnStr
      End If

      isReadOnly = trnStr.State(pslStateReadOnly)
      isValidated = trnStr.State(pslStateTranslated) And Not(trnStr.State(pslStateReview))
      isLocked = trnStr.State(pslStateLocked)
      If isReadOnly Or isValidated Or isLocked Or _
         trnStr.State(pslStateHidden) Then GoTo NextTrnStr
      
      strNum = CStr(trnStr.Number)

      key = GetKey()
      If LookUpDictItem(key) Then
        trnStr.Text= value
        trnStr.State(pslStateTranslated) = True
        trnStr.State(pslStateReview) = True
      Else
        PSL.Output "The translation for " & key & " can not be found"
      End If
NextTrnStr:
    Next trnLstId
NextTrnLst:
  trnLst.Save
  Next trnLst
  PSL.Output "End - Autotrans"
End Sub

Private Function GetKey() As String
  GetKey = prjName + delimeter + srcFile + delimeter + strNum + delimeter + srcStr + tgtLangCode
End Function

Private Function MainDialog(DlgItem$, Action%, SuppValue&) As Boolean
  Dim dlgFolderPath As String
  If Action = 1 Then   'Initialization
    DlgValue "Group3", 1
  ElseIf Action = 2 And DlgItem = "OK" Then 'When "OK" button clicked
    optionXml = DlgValue("Group3") = 0
    optionXlf = DlgValue("Group3") = 1

  ' Let user select folder
  ElseIf DlgItem$ = "SelectFile" Then
    dlgFolderPath = PSL.ActiveProject.Location
    If PSL.SelectFolder(dlgFolderPath,"Specify Folder") Then
      DlgText "FolderPath", dlgFolderPath
    End If
    MainDialog = True
  End If
End Function
