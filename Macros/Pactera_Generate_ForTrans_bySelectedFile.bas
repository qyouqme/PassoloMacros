Attribute VB_Name = "passolo_macro_extract"
'#Reference {F5078F18-C551-11D3-89B9-0000F81FE221}#6.0#0#C:\Windows\System32\msxml6.dll#Microsoft XML, v6.0
'#Reference {F5078F18-C551-11D3-89B9-0000F81FE221}#4.0#0#C:\Windows\SysWow64\msxml4.dll#Microsoft XML, v4.0
'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\SysWOW64\scrrun.dll#Microsoft Scripting Runtime
Option Explicit

'variables for Dialog
Dim optionSource As Boolean
Dim optionReadOnly As Boolean
Dim optionValidated As Boolean
Dim optionForReview As Boolean
Dim optionNewUpdate As Boolean
Dim optionXml As Boolean
Dim optionXlf As Boolean
Dim optionThresholder As Integer

'fso variables
Dim fsys As New FileSystemObject
Dim mainFolderPath As String, folderPath As String
Dim outFilePath As String
Dim extName As String
Dim fileNum As Integer, itemNum As Integer

'Passolo variables
Dim prj As PslProject
Dim srcLangId As Long, tgtLangId As Long
Dim srcLangCode$, tgtLangCode$  'lang-country code, e.g. zh-CN
Dim trnLst As PslTransList
Dim trnStr As PslTransString

Dim trnLstId As Long
Dim srcStr As String, tgtStr As String
Dim strComment$, strId$, resType$, strType$, strNum As Long
Dim isReadOnly As Boolean, isValidated As Boolean
Dim isForReview As Boolean, isNewUpdate As Boolean
Dim isLocked As Boolean, isAT As Boolean, isInAttr As Boolean
Dim srcFilePath$

'xml variables
Dim xmlFilePath$
Dim xmlDoc As New DOMDocument40
Dim xmlRootNode As IXMLDOMElement, xmlProjNode As IXMLDOMElement
Dim xmlFileNode As IXMLDOMElement
Dim xmlStrLstNode As IXMLDOMElement, xmlSrcNode As IXMLDOMElement, xmlTgtNode As IXMLDOMElement
Dim xmlCmtNode As IXMLDOMElement

'pactXliff variables
Dim xlfFilePath$
Dim xlfDoc As New DOMDocument40
Const xlfNS As String = "urn:oasis:names:tc:xliff:document:1.2"
Const pactNS As String = "pgs.pactera.com"
Dim xlfBodyNode As IXMLDOMElement
Dim xlfFileNode As IXMLDOMElement, xlfTransUnitNode As IXMLDOMElement
Dim xlfSrcNode As IXMLDOMElement, xlfTgtNode As IXMLDOMElement
Dim xlfNoteNode As IXMLDOMElement
Dim parserAgain As String

'other variables
Dim CR$, LF$, CRLF$, delimeter$

Sub Main()
  'breaks
  CR = Chr$(13)
  LF = Chr$(10)
  CRLF = CR + LF

  delimeter = ChrW(9478)

   Begin Dialog UserDialog 770,245,"Export to XML/Xliff",.MainDialog ' %GRID:10,7,1,1
      Text 20,20,180,14,"Please specify destination:",.Text1
      TextBox 20,42,650,21,.FolderPath
      OKButton 200,200,70,21
      CancelButton 400,200,90,21
      PushButton 700,42,50,21,"...",.SelectFile
      GroupBox 20,80,110,80,"Target Data",.GroupBox1
      OptionGroup .Group1
         OptionButton 40,110,70,14,"Source",.OptionSource
         OptionButton 40,130,70,14,"Target",.OptionTarget
      GroupBox 160,80,450,80,"Extract Strings Setting",.GroupBox2
      CheckBox 180,110,200,14,"Read Only Units (Grey)",.OptionReadOnly
      CheckBox 390,110,200,14,"Validated Units (Black)",.OptionValidated
      CheckBox 180,130,200,14,"For Review Units (Blue)",.OptionForReview
      CheckBox 390,130,200,14,"New or Update Units (Red)",.OptionNewUpdate
      GroupBox 640,80,110,80,"File Format",.GroupBox3
      CheckBox 660,110,60,14,"Xml",.OptionXml
      CheckBox 660,130,80,14,"pactXliff",.OptionXlf
      Text 20,170,250,14,"Split file when reach maxiumn units: "
      TextBox 270,170,60,14, .OptionThresholder
      Text 350,170,200,14,"(0 means no split.)"
   End Dialog

  Dim dlg As UserDialog
  If Dialog(dlg) = 0 Then Exit Sub

  mainFolderPath = dlg.FolderPath
  If Right(mainFolderPath,1) <> "\" Then mainFolderPath = mainFolderPath + "\"

  Set prj = PSL.ActiveProject
  For Each trnLst In prj.TransLists
    If trnLst.Selected = False Then GoTo NextTrnLst

    srcLangId = trnLst.SourceList.LangID
    srcLangCode = PSL.GetLangCode(srcLangId, pslCodeLangRgn) 'pslCodeLangRgn,8, zh-CN, or pslCodeTrados,15, ZH-CN
    tgtLangId = trnLst.Language.LangID
    tgtLangCode = PSL.GetLangCode(tgtLangId, pslCodeLangRgn)

    srcFilePath = Replace(trnLst.SourceList.SourceFile, prj.Location, "")
    extName = LCase(fsys.GetExtensionName(srcFilePath))
    outFilePath = mainFolderPath + prj.Name + "\" + tgtLangCode + srcFilePath
    
    If parserAgain = "" Then
      If extName <> "htm" Or extName <> "html" Or extName <> "xml" Then
        resType = trnLst.String(1).ResType
        If resType = "Html.Document" Or resType = "XML" Then
          parserAgain = "Html/XML"
        Else
          parserAgain = "Text"
        End If
      End If
    End If
    
    fileNum = 1
    CreateFile
    itemNum = 0
    For trnLstId = 1 To trnLst.StringCount
      Set trnStr = trnLst.String(trnLstId)
      srcStr = Trim(trnStr.SourceText)
      tgtStr = Trim(trnStr.Text)

      If srcStr = "" Then GoTo NextTrnStr

      isReadOnly = optionReadOnly = True And trnStr.State(pslStateReadOnly)
      isValidated = optionValidated = True And trnStr.State(pslStateTranslated) And Not(trnStr.State(pslStateReview))
      isForReview = optionForReview = True And trnStr.State(pslStateTranslated) And trnStr.State(pslStateReview)
      isNewUpdate = optionNewUpdate And Not(trnStr.State(pslStateTranslated))

      If Not(isReadOnly Or isValidated Or isForReview Or isNewUpdate) Or trnStr.State(pslStateHidden) Then GoTo NextTrnStr
        
      itemNum = itemNum + 1
      If optionThresholder > 0 And itemNum > optionThresholder Then
        SaveFile

        itemNum = 0
        fileNum = fileNum + 1
        CreateFile
      End If

      strComment = trnStr.Comment
      strId = trnStr.ID
      resType = trnStr.ResType
      strType = trnStr.Type
      strNum = trnStr.Number

      isLocked = trnStr.State(pslStateLocked)
      isAT = trnStr.State(pslStateAutoTranslated)
      isInAttr = trnStr.Properties.Value(1) = srcStr

      srcStr = PreprocessText(srcStr)
      If optionSource = True Then
        tgtStr = srcStr
      Else
        tgtStr = PreprocessText(tgtStr)
      End If

      If optionXml Then AppendXmlTransUnit
      If optionXlf Then AppendXlfTransUnit
NextTrnStr:
    Next trnLstId
    SaveFile
NextTrnLst:
  Next trnLst
  MsgBox "Done"
End Sub

Private Function PreprocessText(src$) As String
  'consider \r to &#13;
  If isInAttr Then
    src = Replace(src, "&", "&amp;")
    src = Replace(src, "<", "&lt;")
    src = Replace(src, ">", "&gt;")
    src = Replace(src, """", "&quot;")
    src = Replace(src, "'", "&apos;")
  End If
  PreprocessText = src
End Function

Private Sub CreateFile()
  If optionXml Then CreateXmlFile
  If optionXlf Then CreateXlfFile
End Sub

Private Sub CreateXmlFile()
  folderPath = fsys.GetParentFolderName(outFilePath)
  MyCreateFolder folderPath
  xmlFilePath = outFilePath + "_" + tgtLangCode + "_" + Format(fileNum, "00") + ".xml"

  xmlDoc.async = False
  xmlDoc.validateOnParse = False
  xmlDoc.preserveWhiteSpace = True
  xmlDoc.resolveExternals = False

  xmlDoc.loadXML "<?xml version = ""1.0"" encoding = ""utf-8""?><PassoloStrings>" + CRLF + "</PassoloStrings>"
  Set xmlRootNode = xmlDoc.selectSingleNode("PassoloStrings")
  Set xmlProjNode = AppendNewElement(xmlRootNode, "project")
  Set xmlFileNode = AppendNewElement(xmlProjNode,"file")

  xmlProjNode.setAttribute "name", prj.Name
  xmlFileNode.setAttribute "srcFile", srcFilePath
  xmlFileNode.setAttribute "source-language", srcLangCode
  xmlFileNode.setAttribute "target-language", tgtLangCode
End Sub

Private Sub CreateXlfFile()
  folderPath = fsys.GetParentFolderName(outFilePath)
  MyCreateFolder folderPath
  xlfFilePath = outFilePath + "_" + tgtLangCode + "_" + Format(fileNum, "00") + ".pactXliff"

  xlfDoc.async = False
  xlfDoc.validateOnParse = False
  xlfDoc.preserveWhiteSpace = True
  xlfDoc.resolveExternals = False
  xlfDoc.setProperty "SelectionNamespaces", "xmlns:pact=""" + pactNS + """ xmlns:xlf=""" + xlfNS + """"

  xlfDoc.loadXML "<?xml version=""1.0"" encoding=""utf-8""?>" + CRLF + _
  "<xliff xmlns:pact=""" + pactNS + """ xmlns=""" + xlfNS + """ version=""1.2"" pact:version=""1.0"">" + CRLF + _
  "<file original=""" + outFilePath + """ datatype=""x-passolo-xml"" source-language=""" + srcLangCode + """ target-language=""" + tgtLangCode + """>" + CRLF + _
  "<body>" + CRLF + "</body>" + CRLF + _
  "</file>" + CRLF + _
  "</xliff>"

  Set xlfFileNode = xlfDoc.selectSingleNode(".//xlf:file")
  Set xlfBodyNode = xlfDoc.selectSingleNode(".//xlf:body")
  xlfFileNode.setAttribute("pact:prjName", prj.Name)
  xlfFileNode.setAttribute("pact:srcFile", srcFilePath)
  xlfFileNode.setAttribute("pact:parserAgain", parserAgain)
End Sub

Private Sub SaveFile()
  If optionXml Then
    If xmlFileNode.childNodes.length > 0 Then xmlDoc.Save xmlFilePath
  End If
  If optionXlf Then
    If xlfBodyNode.childNodes.length > 0 Then xlfDoc.Save xlfFilePath
  End If
End Sub

Private Sub AppendTransUnit
  If optionXml Then AppendXmlTransUnit
  If optionXlf Then AppendXlfTransUnit
End Sub

Private Sub AppendXmlTransUnit()
  Set xmlStrLstNode = AppendNewElement(xmlFileNode, "string")
  Set xmlSrcNode = AppendNewElement(xmlStrLstNode, "source", True, False)
  Set xmlTgtNode = AppendNewElement(xmlStrLstNode, "target", True, False)
  Set xmlCmtNode = AppendNewElement(xmlStrLstNode, "strComment", True, False)

  xmlStrLstNode.setAttribute "strId", strId + delimeter + CStr(strNum)
  xmlStrLstNode.setAttribute "strNum", strNum
  xmlStrLstNode.setAttribute "resType", resType
  xmlStrLstNode.setAttribute "type", strType
  xmlStrLstNode.setAttribute "validated", CStr(isValidated)
  xmlStrLstNode.setAttribute "locked", CStr(isLocked)
  xmlStrLstNode.setAttribute "AT", CStr(isAT)
  xmlStrLstNode.setAttribute "InXmlAttribute", CStr(isInAttr)

  AppendNewCDATASection xmlSrcNode, srcStr
  AppendNewCDATASection xmlTgtNode, tgtStr
  xmlCmtNode.Text = strComment
End Sub

Private Sub AppendXlfTransUnit()
  Set xlfTransUnitNode = AppendNewElementNS(xlfBodyNode, xlfNS, "trans-unit")
  Set xlfSrcNode = AppendNewElementNS(xlfTransUnitNode, xlfNS, "source", True, False)
  Set xlfTgtNode = AppendNewElementNS(xlfTransUnitNode, xlfNS, "target", True, False)
  Set xlfNoteNode = AppendNewElementNS(xlfTransUnitNode, xlfNS, "note",True, False)

  xlfTransUnitNode.setAttribute "id", strId + delimeter + CStr(strNum)
  xlfTransUnitNode.setAttribute "pact:strNum", strNum
  xlfTransUnitNode.setAttribute "pact:resType", CStr(resType)
  xlfTransUnitNode.setAttribute "pact:type", CStr(strType)
  xlfTransUnitNode.setAttribute "pact:validated", CStr(isValidated)
  xlfTransUnitNode.setAttribute "pact:locked", CStr(isLocked)
  xlfTransUnitNode.setAttribute "pact:AT", CStr(isAT)
  xlfTransUnitNode.setAttribute "pact:InXmlAttribute", CStr(isInAttr)

  If isLocked Or isReadOnly Or isValidated Then xlfTransUnitNode.setAttribute("translate", "no")

  AppendNewCDATASection xlfSrcNode, srcStr
  AppendNewCDATASection xlfTgtNode, tgtStr
  xlfNoteNode.Text = strComment
End Sub

Private Function MainDialog(DlgItem$, Action%, SuppValue&) As Boolean
  Dim dlgFolderPath As String
  If Action = 1 Then   'Initialization
    DlgValue "Group1", 1
    DlgValue "OptionNewUpdate", 1
    DlgValue "OptionXlf", 1
  ElseIf Action = 2 And DlgItem = "OK" Then 'When "OK" button clicked
    optionSource = DlgValue("Group1") = 0
    optionReadOnly = DlgValue("OptionReadOnly")
    optionValidated = DlgValue("OptionValidated")
    optionForReview = DlgValue("OptionForReview")
    optionNewUpdate = DlgValue("OptionNewUpdate")
    optionXml = DlgValue("OptionXml")
    optionXlf = DlgValue("OptionXlf")
    optionThresholder =Val( DlgText("OptionThresholder"))

  ' Let user select folder
  ElseIf DlgItem$ = "SelectFile" Then
    dlgFolderPath = PSL.ActiveProject.Location
    If PSL.SelectFolder(dlgFolderPath,"Specify Folder") Then
      DlgText "FolderPath", dlgFolderPath
    End If
    MainDialog = True
  End If
End Function

'xml common operations
Private Function AppendNewElement(srcNode As IXMLDOMElement, nodeName$, Optional appendBreak As Boolean = True, Optional childBreak As Boolean = True) As IXMLDOMElement
  Dim doc As DOMDocument40
  Set doc = srcNode.ownerDocument
  Set AppendNewElement = doc.createElement(nodeName)
  srcNode.appendChild AppendNewElement
  If appendBreak Then srcNode.appendChild doc.createTextNode(CRLF)
  If childBreak Then AppendNewElement.appendChild doc.createTextNode(CRLF)
End Function

Private Function AppendNewElementNS(srcNode As IXMLDOMElement, nsURI$, nodeName$, Optional appendBreak As Boolean = True, Optional childBreak As Boolean = True) As IXMLDOMElement
  Dim doc As DOMDocument40
  Set doc = srcNode.ownerDocument
  Set AppendNewElementNS = doc.createNode(1, nodeName, nsURI)
  srcNode.appendChild AppendNewElementNS
  If appendBreak Then srcNode.appendChild doc.createTextNode(CRLF)
  If childBreak Then AppendNewElementNS.appendChild doc.createTextNode(CRLF)
End Function

Private Sub appendBreak(node As IXMLDOMElement)
  Dim doc As DOMDocument40
  Set doc = node.ownerDocument
  node.appendChild doc.createTextNode(CRLF)
End Sub

Private Sub AppendBreak2(node As IXMLDOMNode)
  Dim doc As DOMDocument40
  Set doc = node.ownerDocument
  node.appendChild doc.createTextNode(CRLF)
End Sub

Private Sub AppendNewCDATASection(srcNode As IXMLDOMElement, src$)
  Dim doc As DOMDocument40
  Set doc = srcNode.ownerDocument
  srcNode.appendChild doc.createCDATASection(src)
  'Replace tgtStr,"<![CDATA[","<!\[CDATA["
  'Replace tgtStr,"]]>","]\]>"
End Sub

Private Sub MyCreateFolder(folderPath$)
  Dim subFolderPath$
  If Not(fsys.FolderExists(folderPath)) Then
    subFolderPath = fsys.GetParentFolderName(folderPath)
    MyCreateFolder subFolderPath
    fsys.CreateFolder folderPath
  End If
End Sub
