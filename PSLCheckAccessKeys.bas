''This is to check or modify or delete if the access keys
''at end of tranlation strings are the same as in
''source strings and if not change to the same.
''Please use it only when you exactly know what it does! 
''(c) 2007-2008 by gnatix (Last modified on 2008.02.24)
''Modified by wanfu (Last modified on 2010.11.11)

Public trn As PslTransList,TransString As PslTransString,OSLanguage As String

Public SpaceTrn As String,acckeyTrn As String,ExpStringTrn As String,EndStringTrn As String
Public acckeySrc As String,EndStringSrc As String,ShortcutSrc As String,ShortcutTrn As String
Public PreStringTrn As String,EndSpaceSrc As String,EndSpaceTrn As String

Public AllCont As Integer,AccKey As Integer,EndChar As Integer,Acceler As Integer,iVo As Integer
Public srcLineNum As Integer,trnLineNum As Integer,srcAccKeyNum As Integer,trnAccKeyNum As Integer
Public AddedCount As Integer,ModifiedCount As Integer,WarningCount As Integer

Public DefaultCheckList() As String,AppRepStr As String,PreRepStr As String
Public cWriteLoc As String,cSelected() As String,cUpdateSet() As String,cUpdateSetBak() As String
Public CheckList() As String,CheckListBak() As String,CheckDataList() As String,CheckDataListBak() As String
Public tempCheckList() As String,tempCheckDataList() As String

Private Const Version = "2010.11.11"
Private Const ToUpdateCheckVersion = "2010.09.25"
Private Const CheckRegKey = "HKCU\Software\VB and VBA Program Settings\AccessKey\"
Private Const CheckFilePath = MacroDir & "\Data\PSLCheckAccessKeys.dat"
Private Const JoinStr = vbBack
Private Const SubJoinStr = Chr$(1)
Private Const rSubJoinStr = Chr$(19) & Chr$(20)
Private Const LngJoinStr = "|"

Private Const DefaultObject = "Microsoft.XMLHTTP"
Private Const updateAppName = "PSLCheckAccessKeys"
Private Const updateMainFile = "PSLCheckAccessKeys.bas"
Private Const updateINIFile = "PSLMacrosUpdates.ini"
Private Const updateMethod = "GET"
Private Const updateINIMainUrl = "ftp://hhdown:0011@czftp.hanzify.org/update/PSLMacrosUpdates.ini"
Private Const updateINIMinorUrl = "http://www.wanfutrade.com/software/hanhua/PSLMacrosUpdates.ini"
Private Const updateMainUrl = "ftp://hhdown:0011@czftp.hanzify.org/download/PslAccessKey_Modified_wanfu.rar"
Private Const updateMinorUrl = "ftp://hhdown:0011@222.76.212.240:121/downloads/PslAccessKey_Modified_wanfu.rar"
Private Const updateAsync = "False"


'字串处理默认设置
Function CheckSettings(DataName As String,OSLang As String) As String
	Dim ExCr As String,LnSp As String,ChkBkt As String,KpPair As String,ChkEnd As String
	Dim NoTrnEnd As String,TrnEnd As String,Short As String,Key As String,KpKey As String
	If DataName = DefaultCheckList(0) Then
		If OSLang = "0404" Then
			ExCr = "&&,& ,&#,&amp;,&nbsp;,&quot;,&lt;,&gt;,&apos;"
			LnSp = "\r,\n,\r\n,|,&"
			ChkBkt = "(),,[],,<>,≌≈,"
			KpPair = "(),,[],,{},,<>,≌≈,,,,,,'',,？＝,～～,,々—,"""""
			AsiaKey = "1"

			ChkEnd = ". .. ... .... ..... ......  :  ;  !  ?  ,   > >> -> ] } + -"
			TrnEnd = ",| .| ;| !| ?|"
			NoTrnEnd = "*[*],*{*},*<*>,*(?,*[?,*],*{?,*},*<?,*>,*%?!?!,*%?!??!"

			Short = "\t,\b"
			Key = "BackSpace,Tab,Clear,Enter,Shift,Ctrl,Alt,Pause,Caps Lock,Escape,Esc,Space,Break," & _
					"Page Up,Page Down,PgUp,PgDn,End,Home,Left,Up,Right,Down,Left Arrow,Up Arrow," & _
					"Right Arrow,Down Arrow,Select,Print,Ins,Delete,Del,Help,Num Lock,F#,F1#,%?,&?,?"
			KpKey = "?,?,オ?,?,絙?,絙?,オ絙?,絙?," & _
					"龄,龄,オ龄,龄,絙繷,絙繷,オ絙繷,絙繷,◆,□,■,△"
			PreStr = "\001|\n001,\u00A9|(C),\u00AE|(R)"
			AppStr = "&lt;|<,&gt;|>,&quot;|"",&apos;|',&#9;|\t,&#10;|\n,&#13;|\r,&amp;|&,&#39;|'"
		Else
			ExCr = "&&,& ,&#,&amp;,&nbsp;,&quot;,&lt;,&gt;,&apos;"
			LnSp = "\r,\n,\r\n,|,&"
			ChkBkt = "(),（）,[],［］,<>,＜＞,〈〉"
			KpPair = "(),（）,[],［］,{},｛｝,<>,＜＞,〈〉,《》,【】,「」,『』,'',‘’,ˋˊ,‵‵,“”,〝〞,"""""
			AsiaKey = "1"

			ChkEnd = ". .. ... .... ..... ...... 。 : ： ; ； ! ！ ? ？ , ， 、 > >> -> ] } + -"
			TrnEnd = ",|， .|。 ;|； !|！ ?|？"
			NoTrnEnd = "*[*],*{*},*<*>,*(?,*[?,*],*{?,*},*<?,*>,*%?!?!,*%?!??!"

			Short = "\t,\b"
			Key = "BackSpace,Tab,Clear,Enter,Shift,Ctrl,Alt,Pause,Caps Lock,Escape,Esc,Space,Break," & _
					"Page Up,Page Down,PgUp,PgDn,End,Home,Left,Up,Right,Down,Left Arrow,Up Arrow," & _
					"Right Arrow,Down Arrow,Select,Print,Ins,Delete,Del,Help,Num Lock,F#,F1#,%?,&?,?"
			KpKey = "向上键,向下键,向左键,向右键,上箭头,下箭头,左箭头,右箭头," & _
					"向上鍵,向下鍵,向左鍵,向右鍵,上箭頭,下箭頭,左箭頭,右箭頭,↑,↓,←,→"
			PreStr = "\001|\n001,\u00A9|(C),\u00AE|(R)"
			AppStr = "&lt;|<,&gt;|>,&quot;|"",&apos;|',&#9;|\t,&#10;|\n,&#13;|\r,&amp;|&,&#39;|'"
		End If
	ElseIf DataName = DefaultCheckList(1) Then
		If OSLang = "0404" Then
			ExCr = "&&,& ,&#,&amp;,&nbsp;,&quot;,&lt;,&gt;,&apos;"
			LnSp = "\r,\n,\r\n,|,&"
			ChkBkt = "(),,[],,<>,≌≈,"
			KpPair = "(),,[],,{},,<>,≌≈,,,,,,'',,？＝,～～,,々—,"""""
			AsiaKey = "0"

			ChkEnd = ". .. ... .... ..... ......  :  ;  !  ?  ,   > >> -> ] } + -"
			TrnEnd = "|, |. |; |! |?"
			NoTrnEnd = "*[*],*{*},*<*>,*(?,*[?,*],*{?,*},*<?,*>,*%?!?!,*%?!??!"

			Short = "\t,\b"
			Key = "BackSpace,Tab,Clear,Enter,Shift,Ctrl,Alt,Pause,Caps Lock,Escape,Esc,Space,Break," & _
					"Page Up,Page Down,PgUp,PgDn,End,Home,Left,Up,Right,Down,Left Arrow,Up Arrow," & _
					"Right Arrow,Down Arrow,Select,Print,Ins,Delete,Del,Help,Num Lock,F#,F1#,%?,&?,?," & _
					"?,?,オ?,?,絙?,絙?,オ絙?,絙?," & _
					"龄,龄,オ龄,龄,絙繷,絙繷,オ絙繷,絙繷,◆,□,■,△"
			KpKey = "Up,Right,Down,Left Arrow,Up Arrow,Right Arrow,Down Arrow"
			PreStr = "\001|\n001,\u00A9|(C),\u00AE|(R)"
			AppStr = "&lt;|<,&gt;|>,&quot;|"",&apos;|',&#9;|\t,&#10;|\n,&#13;|\r,&amp;|&,&#39;|'"
		Else
			ExCr = "&&,& ,&#,&amp;,&nbsp;,&quot;,&lt;,&gt;,&apos;"
			LnSp = "\r,\n,\r\n,|,&"
			ChkBkt = "(),（）,[],［］,<>,＜＞,〈〉"
			KpPair = "(),（）,[],［］,{},｛｝,<>,＜＞,〈〉,《》,【】,「」,『』,'',‘’,ˋˊ,‵‵,“”,〝〞,"""""
			AsiaKey = "0"

			ChkEnd = ". .. ... .... ..... ...... 。 : ： ; ； ! ！ ? ？ , ， 、 > >> -> ] } + -"
			TrnEnd = "，|, 。|. ；|; ！|! ？|?"
			NoTrnEnd = "*[*],*{*},*<*>,*(?,*[?,*],*{?,*},*<?,*>,*%?!?!,*%?!??!"

			Short = "\t,\b"
			Key = "BackSpace,Tab,Clear,Enter,Shift,Ctrl,Alt,Pause,Caps Lock,Escape,Esc,Space,Break," & _
					"Page Up,Page Down,PgUp,PgDn,End,Home,Left,Up,Right,Down,Left Arrow,Up Arrow," & _
					"Right Arrow,Down Arrow,Select,Print,Ins,Delete,Del,Help,Num Lock,F#,F1#,%?,&?,?," & _
					"向上键,向下键,向左键,向右键,上箭头,下箭头,左箭头,右箭头," & _
					"向上鍵,向下鍵,向左鍵,向右鍵,上箭頭,下箭頭,左箭頭,右箭頭,↑,↓,←,→""
			KpKey = "Up,Right,Down,Left Arrow,Up Arrow,Right Arrow,Down Arrow"
			PreStr = "\001|\n001,\u00A9|(C),\u00AE|(R)"
			AppStr = "&lt;|<,&gt;|>,&quot;|"",&apos;|',&#9;|\t,&#10;|\n,&#13;|\r,&amp;|&,&#39;|'"
		End If
	End If
	If InStr(PreStr,SubJoinStr) Then PreStr = Replace(PreStr,SubJoinStr,rSubJoinStr)
	CheckSettings = ExCr & SubJoinStr & LnSp & SubJoinStr & ChkBkt & SubJoinStr & KpPair & _
					SubJoinStr & AsiaKey & SubJoinStr & ChkEnd & SubJoinStr & NoTrnEnd & _
					SubJoinStr & TrnEnd & SubJoinStr & Short & SubJoinStr & Key & _
					SubJoinStr & KpKey & SubJoinStr & PreStr & SubJoinStr & AppStr
End Function


'请务必查看对话框帮助主题以了解更多信息。
Private Function MainDlgFunc%(DlgItem$, Action%, SuppValue&)
	Dim TypeMsg As Integer,CountMsg As Integer,HeaderID As Integer,Header As String
	If OSLanguage = "0404" Then
		Msg29 = "岿粇"
		Msg30 = "叫匡璶矪瞶﹃摸"
		Msg31 = "叫匡璶矪瞶﹃ず甧"
		Msg32 = "叫匡璶矪瞶﹃摸㎝ず甧"
		Msg36 = "礚猭纗叫浪琩琌Τ糶竚舦:" & vbCrLf & vbCrLf
	Else
		Msg29 = "错误"
		Msg30 = "请选择要处理的字串类型！"
		Msg31 = "请选择要处理的字串内容！"
		Msg32 = "请选择要处理的字串类型和内容！"
		Msg36 = "无法保存！请检查是否有写入下列位置的权限:" & vbCrLf & vbCrLf
	End If

	'获取目标语言
	trnLng = PSL.GetLangCode(trn.Language.LangID,pslCode639_1)
	If trnLng = "zh" Or trnLng = "ja" Or trnLng = "ko" Then TranLang = "Asia"
	If trnLng = "zh" Then
		trnLng = PSL.GetLangCode(trn.Language.LangID,pslCodeLangRgn)
		If trnLng = "zh-CHS" Or trnLng = "zh-SG" Then trnLng = "zh-CN"
		If trnLng = "zh-CHT" Or trnLng = "zh-HK" Or trnLng = "zh-MO" Then trnLng = "zh-TW"
	End If

	Select Case Action%
	Case 1
		If Join(cSelected) <> "" Then
			AutoSet = cSelected(0)
			CheckSet = cSelected(1)
			AutoChk = StrToInteger(cSelected(2))
			AutoSele = StrToInteger(cSelected(3))
			miVo = StrToInteger(cSelected(4))
			mAllType = StrToInteger(cSelected(5))
			mMenu = StrToInteger(cSelected(6))
			mDialog = StrToInteger(cSelected(7))
			mString = StrToInteger(cSelected(8))
			mAccTable = StrToInteger(cSelected(9))
			mVer = StrToInteger(cSelected(10))
			mOther = StrToInteger(cSelected(11))
			mSelOnly = StrToInteger(cSelected(12))
			mAllCont = StrToInteger(cSelected(13))
			mAccKey = StrToInteger(cSelected(14))
			mEndChar = StrToInteger(cSelected(15))
			mAcceler = StrToInteger(cSelected(16))
			VerTag = StrToInteger(cSelected(17))
			SetTag = StrToInteger(cSelected(18))
			DateTag = StrToInteger(cSelected(19))
			StateTag = StrToInteger(cSelected(20))
			AllTag = StrToInteger(cSelected(21))
			NoChkTag = StrToInteger(cSelected(22))
			NoChgSta = StrToInteger(cSelected(23))
			repStr = StrToInteger(cSelected(24))
			KeepSet = StrToInteger(cSelected(25))
		End If

		DlgText "AutoSetList",AutoSet
		DlgText "CheckSetList",CheckSet
		If DlgText("AutoSetList") = "" Then DlgValue "AutoSetList",0
		If DlgText("CheckSetList") = "" Then DlgValue "CheckSetList",0
		DlgValue "AutoChk",AutoChk
		DlgValue "AutoSelection",AutoSele
		DlgValue "Options",miVo
		DlgValue "AllType",mAllType
		DlgValue "Menu",mMenu
		DlgValue "Dialog",mDialog
		DlgValue "Strings",mString
		DlgValue "AccTable",mAccTable
		DlgValue "Versions",mVer
		DlgValue "Other",mOther
		DlgValue "Seleted",mSelOnly
		DlgValue "AllCont",mAllCont
		DlgValue "AccKey",mAccKey
		DlgValue "EndChar",mEndChar
		DlgValue "Acceler",mAcceler
		DlgValue "TagVer",VerTag
		DlgValue "TagSet",SetTag
		DlgValue "TagDate",DateTag
		DlgValue "TagState",StateTag
		DlgValue "TagAll",AllTag
		DlgValue "NoCheckTag",NoChkTag
		DlgValue "NoChangeState",NoChgSta
		DlgValue "AutoRepStr",repStr
		DlgValue "KeepSet",KeepSet

		TypeMsg = mAllType + mMenu + mDialog + mString + mAccTable + mVer + mOther + mSelOnly
		CountMsg = mAllCont + mAccKey + mEndChar + mAcceler
		TagMsg = VerTag + SetTag + DateTag + StateTag
		If TypeMsg = 0 Then DlgValue "AllType",1
		If CountMsg = 0 Then DlgValue "AllCont",1
		If TagMsg <> 0 Then DlgValue "TagAll",0
		DlgVisible "TagState",False
		If trn.IsOpen = False Then
			DlgEnable "Seleted",False
			DlgValue "Seleted",0
		End If
		If AutoSele = 1 Then
			HeaderID = getCheckID(CheckDataList,trnLng,TranLang)
			DlgValue "AutoSetList",HeaderID
			DlgValue "CheckSetList",HeaderID
			DlgEnable "AutoSetList",False
			DlgEnable "CheckSetList",False
			DlgEnable "AutoChk",False
		Else
			If AutoChk = 1 Then	DlgEnable "AutoSetList",False
		End If
		If CheckNullData(CheckSet,CheckDataList,"4,11",0) = True Then DlgEnable "OKButton",False
		DlgEnable "SaveButton",False
	Case 2 ' 数值更改或者按下了按钮
		AutoSet = DlgText("AutoSetList")
		CheckSet = DlgText("CheckSetList")
		AutoChk = DlgValue("AutoChk")
		AutoSele = DlgValue("AutoSelection")
		miVo = DlgValue("Options")
		mAllType = DlgValue("AllType")
		mMenu = DlgValue("Menu")
		mDialog = DlgValue("Dialog")
		mString =  DlgValue("Strings")
		mAccTable = DlgValue("AccTable")
		mVer = DlgValue("Versions")
		mOther = DlgValue("Other")
		mSelOnly = DlgValue("Seleted")
		mAllCont = DlgValue("AllCont")
		mAccKey = DlgValue("AccKey")
		mEndChar = DlgValue("EndChar")
		mAcceler = DlgValue("Acceler")
		VerTag = DlgValue("TagVer")
		SetTag = DlgValue("TagSet")
		DateTag = DlgValue("TagDate")
		StateTag = DlgValue("TagState")
		AllTag = DlgValue("TagAll")
		NoChkTag = DlgValue("NoCheckTag")
		NoChgSta = DlgValue("NoChangeState")
		repStr = DlgValue("AutoRepStr")
		KeepSet = DlgValue("KeepSet")
		nSelected = AutoSet  & JoinStr & CheckSet & JoinStr & AutoChk & JoinStr & AutoSele & _
					JoinStr & miVo & JoinStr & mAllType & JoinStr & mMenu & JoinStr & mDialog & _
					JoinStr & mString & JoinStr & mAccTable & JoinStr & mVer & JoinStr & mOther & _
					JoinStr & mSelOnly & JoinStr & mAllCont & JoinStr & mAccKey & JoinStr & _
					mEndChar & JoinStr & mAcceler & JoinStr & VerTag & JoinStr & SetTag & _
					JoinStr & DateTag & JoinStr & StateTag & JoinStr & AllTag & JoinStr & _
					NoChkTag & JoinStr & NoChgSta & JoinStr & repStr & JoinStr & KeepSet
		'检测字串类型和内容选择是否同时选定全部和其他单项
		If DlgItem$ = "Menu" Or DlgItem$ = "Dialog" Or DlgItem$ = "Strings" Or DlgItem$ = "AccTable" Or _
			DlgItem$ = "Versions" Or DlgItem$ = "Other" Then
			If mMenu + mDialog + mString + mAccTable + mVer + mOther + mSelOnly = 0 Then
				DlgValue "AllType",1
			Else
				DlgValue "AllType",0
			End If
			If mMenu + mDialog + mString + mAccTable + mVer + mOther <> 0 Then DlgValue "Seleted",0
		End If
		If DlgItem$ = "AllType" And mAllType = 1 Then
			DlgValue "Menu",0
			DlgValue "Dialog",0
			DlgValue "Strings",0
			DlgValue "AccTable",0
			DlgValue "Versions",0
			DlgValue "Other",0
			DlgValue "Seleted",0
		End If
		If DlgItem$ = "Seleted" Then
			If mSelOnly = 1 Then
				DlgValue "AllType",0
				DlgValue "Menu",0
				DlgValue "Dialog",0
				DlgValue "Strings",0
				DlgValue "AccTable",0
				DlgValue "Versions",0
				DlgValue "Other",0
			Else
				DlgValue "AllType",1
			End If
		End If
		If DlgItem$ = "AccKey" Or DlgItem$ = "EndChar" Or DlgItem$ = "Acceler" Then
			If mAccKey = 1 Or mEndChar = 1 Or mAcceler = 1 Then
				DlgValue "AllCont",0
			End If
			If mAccKey + mEndChar + mAcceler = 0 Then
				DlgValue "AllCont",1
			End If
		End If
		If DlgItem$ = "AllCont" Then
			If mAllCont = 1 Then
				DlgValue "AccKey",0
				DlgValue "EndChar",0
				DlgValue "Acceler",0
			End If
			If mAccKey + mEndChar + mAcceler = 0 Then
				DlgValue "AllCont",1
			End If
		End If
		If DlgItem$ = "TagVer" Or DlgItem$ = "TagSet" Or DlgItem$ = "TagDate" Or DlgItem$ = "TagState" Then
			If VerTag = 1 Or SetTag = 1 Or DateTag = 1 Or StateTag = 1 Then
				DlgValue "TagAll",0
				DlgValue "NoCheckTag",0
			End If
		End If
		If DlgItem$ = "TagAll" Then
			If AllTag = 1 Then
				DlgValue "TagVer",0
				DlgValue "TagSet",0
				DlgValue "TagDate",0
				DlgValue "TagState",0
			End If
		End If
		If DlgItem$ = "NoCheckTag" Then
			If NoChkTag = 1 Then
				DlgValue "TagVer",0
				DlgValue "TagSet",0
				DlgValue "TagDate",0
				DlgValue "TagState",0
				DlgValue "TagAll",1
			End If
		End If
		If DlgItem$ = "SetButton" Then
			CheckListBak = CheckList
			CheckDataListBak = CheckDataList
			cUpdateSetBak = cUpdateSet
			AutoSetID = DlgValue("AutoSetList")
			HeaderID = DlgValue("CheckSetList")
			Call Settings(HeaderID)
			DlgListBoxArray "AutoSetList",CheckList()
			DlgListBoxArray "CheckSetList",CheckList()
			If DlgValue("AutoSelection") = 1 Then
				HeaderID = getCheckID(CheckDataList,trnLng,TranLang)
				DlgValue "AutoSetList",HeaderID
				DlgValue "CheckSetList",HeaderID
			Else
				DlgValue "AutoSetList",AutoSetID
				DlgValue "CheckSetList",HeaderID
			End If
			If DlgText("AutoSetList") = "" Then DlgValue "AutoSetList",0
			If DlgText("CheckSetList") = "" Then DlgValue "CheckSetList",0
		End If
		If DlgItem$ = "AutoSetList" Or DlgItem$ = "CheckSetList" Then
			If cSelected(0) = "" Then AutoSet = CheckList(0) Else AutoSet = cSelected(0)
			If cSelected(25) = "" Then KeepSet = 0 Else KeepSet = cSelected(25)
			If DlgValue("AutoChk") = 1 Then DlgText "AutoSetList",DlgText("CheckSetList")
		End If
		If DlgItem$ = "AutoChk" Or DlgItem$ = "AutoSelection" Then
			If cSelected(0) = "" Then AutoSet = CheckList(0) Else AutoSet = cSelected(0)
			If cSelected(1) = "" Then CheckSet = CheckList(0) Else CheckSet = cSelected(1)
			If cSelected(25) = "" Then KeepSet = 0 Else KeepSet = cSelected(25)
			If DlgItem$ = "AutoChk" Then
				If DlgValue("AutoChk") = 1 Then
					DlgText "AutoSetList",DlgText("CheckSetList")
					DlgEnable "AutoSetList",False
				Else
					DlgText "AutoSetList",AutoSet
					If AutoSele = 1 Then DlgEnable "AutoSetList",False
					If AutoSele = 0 Then DlgEnable "AutoSetList",True
				End If
			Else
				If DlgValue("AutoSelection") = 1 Then
					HeaderID = getCheckID(CheckDataList,trnLng,TranLang)
					DlgValue "AutoSetList",HeaderID
					DlgValue "CheckSetList",HeaderID
					DlgEnable "AutoSetList",False
					DlgEnable "CheckSetList",False
					DlgEnable "AutoChk",False
				Else
					DlgText "AutoSetList",AutoSet
					DlgText "CheckSetList",CheckSet
					If AutoChk = 1 Then DlgEnable "AutoSetList",False
					If AutoChk = 0 Then DlgEnable "AutoSetList",True
					DlgEnable "CheckSetList",True
					DlgEnable "AutoChk",True
				End If
			End If
		End If
		If CheckNullData(CheckSet,CheckDataList,"4,11",0) = True Then
			DlgEnable "OKButton",False
		Else
			DlgEnable "OKButton",True
		End If
		If DlgItem$ = "OKButton" Then
			'检测字串类型和内容选择是否为空
			TypeMsg = mAllType + mMenu + mDialog + mString + mAccTable + mVer + mOther + mSelOnly
			CountMsg = mAllCont + mAccKey + mEndChar + mAcceler
			If TypeMsg = 0 And CountMsg <> 0 Then
				MsgBox(Msg30,vbOkOnly+vbInformation,Msg29)
				MainDlgFunc% = True ' 防止按下按钮关闭对话框窗口
			ElseIf TypeMsg <> 0 And CountMsg = 0 Then
				MsgBox(Msg31,vbOkOnly+vbInformation,Msg29)
				MainDlgFunc% = True ' 防止按下按钮关闭对话框窗口
			ElseIf TypeMsg = 0 And CountMsg = 0 Then
				MsgBox(Msg32,vbOkOnly+vbInformation,Msg29)
				MainDlgFunc% = True ' 防止按下按钮关闭对话框窗口
			End If
			If Join(cSelected) = nSelected Then Exit Function
			cSelected = Split(nSelected,JoinStr)
			If CheckWrite(CheckDataList,cWriteLoc,"Main") = False Then
				If WriteLocation = CheckFilePath Then Msg36 = Msg36 & CheckFilePath
				If WriteLocation = CheckRegKey Then Msg36 = Msg36 & CheckRegKey
				If MsgBox(Msg36,vbYesNo+vbInformation,Msg29) = vbNo Then Exit Function
				MainDlgFunc% = True ' 防止按下按钮关闭对话框窗口
			End If
		End If
		If DlgItem$ = "SaveButton" And Join(cSelected) <> nSelected Then
			cSelected = Split(nSelected,JoinStr,-1)
			cSelected(25) = "1"
			If CheckWrite(CheckDataList,cWriteLoc,"Main") = False Then
				If WriteLocation = CheckFilePath Then Msg36 = Msg36 & CheckFilePath
				If WriteLocation = CheckRegKey Then Msg36 = Msg36 & CheckRegKey
				MsgBox(Msg36,vbOkOnly+vbInformation,Msg29)
			Else
				DlgEnable "SaveButton",False
			End If
			DlgValue "KeepSet",1
		Else
			If Join(cSelected) <> nSelected Then
				DlgEnable "SaveButton",True
			Else
				DlgEnable "SaveButton",False
			End If
		End If
		If DlgItem$ = "HelpButton" Then Call CheckHelp("MainHelp")
		If DlgItem$ = "AboutButton" Then Call CheckHelp("About")
		If DlgItem$ <> "OKButton" And DlgItem$ <> "CancelButton" Then
			Header = DlgText("AutoSetList")
			If Header <> AutoSet Then
				DlgValue "KeepSet",1
				DlgEnable "KeepSet",False
			Else
				DlgValue "KeepSet",KeepSet
				DlgEnable "KeepSet",True
			End If
			MainDlgFunc% = True ' 防止按下按钮关闭对话框窗口
		End If
	End Select
End Function


' 主程序
Sub Main
	Dim i As Integer,j As Integer,CheckVer As String,CheckSet As String,CheckID As Integer
	Dim srcString As String,trnString As String,OldtrnString As String,NewtrnString As String
	Dim LineNumErrCount As Integer,accKeyNumErrCount As Integer,ErrorCount As Integer
	Dim TransListOpen As Boolean,CheckDate As Date,TranDate As Date,CheckState As String
	Dim TranLang As String

	On Error GoTo SysErrorMsg
	'检测系统语言
	Dim strKeyPath As String, WshShell As Object
	Set WshShell = CreateObject("WScript.Shell")
	If WshShell Is Nothing Then
		MsgBox("Your system is missing the Windows Script Host (WSH) object!")
		Exit Sub
	End If
	strKeyPath = "HKLM\SYSTEM\CurrentControlSet\Control\Nls\Language\Default"
	OSLanguage = WshShell.RegRead(strKeyPath)
	If OSLanguage = "" Then
		strKeyPath = "HKLM\SYSTEM\CurrentControlSet\Control\Nls\Language\InstallLanguage"
		OSLanguage = WshShell.RegRead(strKeyPath)
	End If
	Set WshShell = Nothing

	If OSLanguage = "0404" Then
		Msg00 = "(c) 2007-2008 by gnatix, 2009-2010 by wanfu  セ: " & Version
		Msg01 = "獽倍龄沧ゎ才㎝硉竟浪琩エ栋"
		Msg02 = "セ祘Αノ浪琩э㎝埃赣陆亩睲虫い┮Τ陆亩﹃" & _
				"獽倍龄沧ゎ才㎝硉竟叫浪琩硋兵秈︽狡琩"
		Msg03 = "陆亩睲虫: "

		Msg04 = "砞﹚匡"
		Msg05 = "笆エ栋:"
		Msg06 = "浪琩エ栋:"
		Msg07 = "笆エ栋㎝浪琩エ栋(&U)"
		Msg08 = "笆匡(&X)"

		Msg09 = "浪琩夹癘"
		Msg10 = "┛菠セ(&B)"
		Msg11 = "┛菠砞﹚(&G)"
		Msg12 = "┛菠ら戳(&T)"
		Msg13 = "┛菠篈(&Z)"
		Msg14 = "场┛菠(&Q)"

		Msg15 = "巨匡"
		Msg16 = "度浪琩(&O)"
		Msg17 = "浪琩タ(&M)"
		Msg18 = "埃獽倍龄(&R)"

		Msg19 = "﹃摸"
		Msg20 = "场(&A)"
		Msg21 = "匡虫(&M)"
		Msg22 = "癸杠よ遏(&D)"
		Msg23 = "﹃(&S)"
		Msg24 = "硉竟(&A)"
		Msg25 = "セ(&V)"
		Msg26 = "ㄤ(&O)"
		Msg27 = "度匡拒(&L)"

		Msg28 = "﹃ず甧"
		Msg29 = "场(&F)"
		Msg30 = "獽倍龄(&K)"
		Msg31 = "沧ゎ才(&E)"
		Msg32 = "硉竟(&P)"

		Msg33 = "铬筁﹃"
		Msg34 = "ㄑ滦糵(&K)"
		Msg35 = "喷靡(&E)"
		Msg36 = "ゼ陆亩(&N)"

		Msg37 = "ㄤ匡兜"
		Msg38 = "膥尿笆纗┮Τ匡(&V)"
		Msg39 = "ぃミ┪埃浪琩夹癘(&K)"
		Msg40 = "ぃ跑﹍陆亩篈(&Y)"
		Msg41 = "笆蠢传疭﹚じ(&L)"

		Msg42 = "闽(&A)"
		Msg43 = "弧(&H)"
		Msg44 = "砞﹚(&S)"
		Msg45 = "纗匡(&L)"

		Msg50 = "絋粄"
		Msg51 = "癟"
		Msg52 = "岿粇"
		Msg53 =	"眤 Passolo セびセエ栋度続ノ Passolo 6.0 のセ叫どㄏノ"
		Msg54 = "叫匡陆亩睲虫"
		Msg55 = "タミ㎝穝陆亩睲虫..."
		Msg56 = "礚猭ミ㎝穝陆亩睲虫叫浪琩眤盡砞﹚"
		Msg57 = "赣睲虫ゼ砆秨币篈浪琩岿粇礚猭э﹃" & vbCrLf & _
				"眤惠璶琵╰参笆秨币赣陆亩睲虫盾" & vbCrLf & vbCrLf & _
				"秨币浪琩狦т岿粇盢ぃ穦笆纗э" & vbCrLf & _
				"闽超陆亩睲虫玥盢笆闽超陆亩睲虫"
		Msg58 = "タ秨币陆亩睲虫..."
		Msg59 = "礚猭秨币陆亩睲虫叫浪琩眤盡砞﹚"
		Msg60 = "赣睲虫矪秨币篈篈秈︽岿粇浪琩㎝э盢ㄏ" & vbCrLf & _
				"眤ゼ纗陆亩礚猭临╰参盢纗眤陆" & vbCrLf & _
				"亩礛秈︽浪琩㎝э" & vbCrLf & vbCrLf & _
				"眤絋﹚璶琵╰参笆纗眤陆亩盾"
		Msg62 = "タ浪琩惠璶碭だ牧叫祔..."
		Msg65 = "亩ゅ: "
		Msg66 = "璸ノ: "
		Msg67 = "hh  mm だ ss "
		Msg70 = "璣ゅいゅ"
		Msg71 = "いゅ璣ゅ"
		Msg72 = "M:祘Αセ"
		Msg73 = "M:砞﹚嘿"
		Msg74 = "M:浪琩篈"
		Msg75 = "M:浪琩ら戳"
		Msg76 = "Τ岿粇"
		Msg77 = "礚岿粇"
		Msg78 = "タ"
		Msg79 = "yyyymるdら hh:mm:ss"
	Else
		Msg00 = "(c) 2007-2008 by gnatix, 2009-2010 by wanfu  版本: " & Version
		Msg01 = "快捷键、终止符和加速器检查宏"
		Msg02 = "本程序用于检查、修改和删除该翻译列表中所有翻译字串" & _
				"的快捷键、终止符和加速器。请在检查后逐条进行人工复查。"
		Msg03 = "翻译列表: "

		Msg04 = "配置选择"
		Msg05 = "自动宏:"
		Msg06 = "检查宏:"
		Msg07 = "自动宏和检查宏相同(&U)"
		Msg08 = "自动选择(&X)"

		Msg09 = "检查标记"
		Msg10 = "忽略版本(&B)"
		Msg11 = "忽略配置(&G)"
		Msg12 = "忽略日期(&T)"
		Msg13 = "忽略状态(&Z)"
		Msg14 = "全部忽略(&Q)"

		Msg15 = "操作选择"
		Msg16 = "仅检查(&O)"
		Msg17 = "检查并修正(&M)"
		Msg18 = "删除快捷键(&R)"

		Msg19 = "字串类型"
		Msg20 = "全部(&A)"
		Msg21 = "菜单(&M)"
		Msg22 = "对话框(&D)"
		Msg23 = "字串表(&S)"
		Msg24 = "加速器表(&A)"
		Msg25 = "版本(&V)"
		Msg26 = "其他(&O)"
		Msg27 = "仅选定(&L)"

		Msg28 = "字串内容"
		Msg29 = "全部(&F)"
		Msg30 = "快捷键(&K)"
		Msg31 = "终止符(&E)"
		Msg32 = "加速器(&P)"

		Msg33 = "跳过字串"
		Msg34 = "供复审(&K)"
		Msg35 = "已验证(&E)"
		Msg36 = "未翻译(&N)"

		Msg37 = "其他选项"
		Msg38 = "继续时自动保存所有选择(&V)"
		Msg39 = "不创建或删除检查标记(&K)"
		Msg40 = "不更改原始翻译状态(&Y)"
		Msg41 = "自动替换特定字符(&L)"

		Msg42 = "关于(&A)"
		Msg43 = "帮助(&H)"
		Msg44 = "设置(&S)"
		Msg45 = "保存选择(&L)"

		Msg50 = "确认"
		Msg51 = "信息"
		Msg52 = "错误"
		Msg53 =	"您的 Passolo 版本太低，本宏仅适用于 Passolo 6.0 及以上版本，请升级后再使用。"
		Msg54 = "请选择一个翻译列表！"
		Msg55 = "正在创建和更新翻译列表..."
		Msg56 = "无法创建和更新翻译列表，请检查您的方案设置。"
		Msg57 = "该列表未被打开。此状态下可以检查错误但无法修改字串。" & vbCrLf & _
				"您需要让系统自动打开该翻译列表吗？" & vbCrLf & vbCrLf & _
				"为了安全，打开检查后，如果找到错误将不会自动保存修改" & vbCrLf & _
				"并关闭翻译列表。否则将自动关闭翻译列表。"
		Msg58 = "正在打开翻译列表..."
		Msg59 = "无法打开翻译列表，请检查您的方案设置。"
		Msg60 = "该列表已处于打开状态。此状态下进行错误检查和修改将使" & vbCrLf & _
				"您未保存的翻译无法还原。为了安全，系统将先保存您的翻" & vbCrLf & _
				"译，然后进行检查和修改。" & vbCrLf & vbCrLf & _
				"您确定要让系统自动保存您的翻译吗？"
		Msg62 = "正在检查，可能需要几分钟，请稍候..."
		Msg65 = "原译文: "
		Msg66 = "合计用时: "
		Msg67 = "hh 小时 mm 分 ss 秒"
		Msg70 = "英文到中文"
		Msg71 = "中文到英文"
		Msg72 = "M:程序版本"
		Msg73 = "M:配置名称"
		Msg74 = "M:检查状态"
		Msg75 = "M:检查日期"
		Msg76 = "有错误"
		Msg77 = "无错误"
		Msg78 = "已修正"
		Msg79 = "yyyy年m月d日 hh:mm:ss"
	End If

	If PSL.Version < 600 Then
		MsgBox Msg53,vbOkOnly+vbInformation,Msg51
		Exit Sub
	End If

	Set trn = PSL.ActiveTransList
	'检测翻译列表是否被选择
	If trn Is Nothing Then
		MsgBox Msg54,vbOkOnly+vbInformation,Msg51
		Exit Sub
	End If

	'初始化数组
	ReDim DefaultCheckList(1),CheckList(0),CheckDataList(0)
	DefaultCheckList(0) = Msg70
	DefaultCheckList(1) = Msg71

	'读取字串处理设置
	If CheckGet("",CheckList,CheckDataList,"") = False Then
		For i = LBound(DefaultCheckList) To UBound(DefaultCheckList)
			CheckName = DefaultCheckList(i)
			LangPair = Join(LangCodeList(CheckName,OSLanguage,1,107),SubJoinStr)
			Data = CheckName & JoinStr & CheckSettings(CheckName,OSLanguage) & JoinStr & LangPair
			CreateArray(CheckName,Data,CheckList,CheckDataList)
		Next i
	Else
		For i = LBound(DefaultCheckList) To UBound(DefaultCheckList)
			CheckName = DefaultCheckList(i)
			If InStr(Join(DefaultCheckList,JoinStr),CheckName) = 0 Then
				LangPair = Join(LangCodeList(CheckName,OSLanguage,1,107),SubJoinStr)
				Data = CheckName & JoinStr & CheckSettings(CheckName,OSLanguage) & JoinStr & LangPair
				CreateArray(CheckName,Data,CheckList,CheckDataList)
			End If
		Next i
	End If

	'获取更新数据并检查新版本
	If Join(cUpdateSet) <> "" Then
		updateMode = cUpdateSet(0)
		updateUrl = cUpdateSet(1)
		CmdPath = cUpdateSet(2)
		CmdArg = cUpdateSet(3)
		updateCycle = cUpdateSet(4)
		updateDate = cUpdateSet(5)
		If updateMode = "" Then
			cUpdateSet(0) = "1"
			updateMode = "1"
		End If
		If updateUrl = "" Then cUpdateSet(1) = updateMainUrl & vbCrLf & updateMinorUrl
		If CmdPath = "" Or (CmdPath <> "" And Dir(CmdPath) = "") Then
			CmdPathArgList = Split(getCMDPath(".rar","",""),JoinStr)
			cUpdateSet(2) = CmdPathArgList(0)
			cUpdateSet(3) = CmdPathArgList(1)
		End If
	Else
		updateMode = "1"
		updateUrl = updateMainUrl & vbCrLf & updateMinorUrl
		updateCycle = "7"
		Temp = updateMode & JoinStr & updateUrl & JoinStr & getCMDPath(".rar","","") & _
				JoinStr & updateCycle & JoinStr & updateDate
		cUpdateSet = Split(Temp,JoinStr)
	End If
	If updateMode <> "" And updateMode <> "2" Then
		If updateDate <> "" Then
			i = CInt(DateDiff("d",CDate(updateDate),Date))
			m = StrComp(Format(Date,"yyyy-MM-dd"),updateDate)
			If updateCycle <> "" Then n = i - CInt(updateCycle)
		End If
		If updateDate = "" Or (m = 1 And n >= 0) Then
			If Download(updateMethod,updateUrl,updateAsync,updateMode) = True Then
				cUpdateSet(5) = Format(Date,"yyyy-MM-dd")
				CheckWrite(CheckDataList,cWriteLoc,"Update")
				GoTo ExitSub
			Else
				cUpdateSet(5) = Format(Date,"yyyy-MM-dd")
				CheckWrite(CheckDataList,cWriteLoc,"Update")
			End If
		End If
	End If

	'对话框
	Msg03 = Msg03 & trn.Title & " - " & PSL.GetLangCode(trn.Language.LangID,pslCodeText)
	Begin Dialog UserDialog 620,462,Msg01,.MainDlgFunc ' %GRID:10,7,1,1
		Text 20,7,580,14,Msg00,.Text1,2
		Text 20,28,580,28,Msg02,.Text2
		Text 20,63,580,14,Msg03,.Text3,2

		GroupBox 20,84,390,98,Msg04,.Configuration
		Text 40,105,70,14,Msg05,.Text4
		Text 40,133,70,14,Msg06,.Text5
		DropListBox 120,101,270,21,CheckList(),.AutoSetList
		DropListBox 120,129,270,21,CheckList(),.CheckSetList
		CheckBox 40,157,230,14,Msg07,.AutoChk
		CheckBox 280,157,120,14,Msg08,.AutoSelection

		GroupBox 430,84,170,98,Msg09,.CheckOption
		CheckBox 450,105,140,14,Msg10,.TagVer
		CheckBox 450,123,140,14,Msg11,.TagSet
		CheckBox 450,143,140,14,Msg12,.TagDate
		CheckBox 450,154,140,14,Msg13,.TagState
		CheckBox 450,161,140,14,Msg14,.TagAll

		GroupBox 20,189,580,49,Msg15,.OptionSelection
		OptionGroup .Options
			OptionButton 80,210,160,14,Msg16,.CheckOnly
			OptionButton 250,210,160,14,Msg17,.CheckAndModiy
			OptionButton 420,210,170,14,Msg18,.RemoveKey

		GroupBox 20,245,580,70,Msg19,.StrTypeSelection
		CheckBox 40,266,130,14,Msg20,.AllType
		CheckBox 180,266,130,14,Msg21,.Menu
		CheckBox 330,266,130,14,Msg22,.Dialog
		CheckBox 470,266,120,14,Msg23,.Strings
		CheckBox 40,287,130,14,Msg24,.AccTable
		CheckBox 180,287,130,14,Msg25,.Versions
		CheckBox 330,287,130,14,Msg26,.Other
		CheckBox 470,287,120,14,Msg27,.Seleted

		GroupBox 20,322,580,49,Msg28,.StrContSelection
		CheckBox 40,343,130,14,Msg29,.AllCont
		CheckBox 180,343,130,14,Msg30,.AccKey
		CheckBox 330,343,130,14,Msg31,.EndChar
		CheckBox 470,343,120,14,Msg32,.Acceler
		CheckBox 40,399,260,14,Msg38,.KeepSet
		CheckBox 330,378,260,14,Msg39,.NoCheckTag
		CheckBox 40,378,260,14,Msg40,.NoChangeState
		CheckBox 330,399,260,14,Msg41,.AutoRepStr

		PushButton 20,434,90,21,Msg42,.AboutButton
		PushButton 110,434,90,21,Msg43,.HelpButton
		PushButton 200,434,90,21,Msg44,.SetButton
		PushButton 290,434,110,21,Msg45,.SaveButton
		OKButton 420,434,90,21,.OKButton
		CancelButton 510,434,90,21,.CancelButton '6 取消
	End Dialog
	Dim dlg As UserDialog
	If Dialog(dlg) = 0 Then GoTo ExitSub
	CheckID = dlg.CheckSetList
	iVo = dlg.Options
	AllCont = dlg.AllCont
	AccKey = dlg.AccKey
	EndChar = dlg.EndChar
	Acceler = dlg.Acceler
	AutoMacroSet = cSelected(0)
	CheckMacroSet = cSelected(1)

	'获取字串类型组合
	If dlg.Menu = 1 Then StrTypes = "|Menu|"
	If dlg.Dialog = 1 Then StrTypes = StrTypes & "|Dialog|"
	If dlg.Strings = 1 Then StrTypes = StrTypes & "|StringTable|"
	If dlg.AccTable = 1 Then StrTypes = StrTypes & "|AcceleratorTable|"
	If dlg.Versions = 1 Then StrTypes = StrTypes & "|Version|"

	'提示打开关闭的翻译列表，如果需要修改和删除
	TransListOpen = False
	If trn.IsOpen = False And iVo <> 0 Then
		Massage = MsgBox(Msg03 & vbCrLf & vbCrLf & Msg57,vbYesNoCancel,Msg50)
		If Massage = vbYes Then
			PSL.Output Msg58
			If trn.Open = False Then
				MsgBox Msg59,vbOkOnly+vbInformation,Msg51
				GoTo ExitSub
			Else
				TransListOpen = True
			End If
		End If
		If Massage = vbNo Then iVo = 0
		If Massage = vbCancel Then GoTo ExitSub
	End If

	'提示保存打开的翻译列表，以免处理后数据不可恢复
	If trn.IsOpen = True And iVo <> 0 Then
		Massage = MsgBox(Msg03 & vbCrLf & vbCrLf & Msg60,vbYesNoCancel,Msg50)
 		If Massage = vbYes Then trn.Save
		If Massage = vbCancel Then GoTo ExitSub
	End If

	'如果翻译列表的更改时间晚于原始列表，自动更新
	If trn.SourceList.LastChange > trn.LastChange Then
		PSL.Output Msg55
		If trn.Update = False Then
			MsgBox Msg56,vbOkOnly+vbInformation,Msg51
			GoTo ExitSub
		End If
	End If

	'设置检查宏专用的用户定义属性
	If Not trn Is Nothing Then
		If trn.Property(19980) <> "CheckAccessKeys" Then
			trn.Property(19980) = "CheckAccessKeys"
		End If
	End If

	'获取PSL的目标语言代码
	trnLng = PSL.GetLangCode(trn.Language.LangID,pslCode639_1)
	If trnLng = "zh" Or trnLng = "ja" Or trnLng = "ko" Then TranLang = "Asia"

	'释放不再使用的动态数组所使用的内存
	Erase CheckListBak,CheckDataListBak,tempCheckList,tempCheckDataList

	'根据是否选择 "仅选定字串" 项设置要检查的字串数
	If dlg.Seleted = 0 Then
		StringCount = trn.StringCount
	Else
		StringCount = trn.StringCount(pslSelection)
	End If

	'开始字串操作
	StartTimes = Timer
	ErrorCount = 0
	AddedCount = 0
	ModifiedCount = 0
	WarningCount = 0
	LineNumErrCount = 0
	accKeyNumErrCount = 0
	PSL.OutputWnd.Clear
	PSL.Output Msg62
	For j = 1 To StringCount
		'根据是否选择 "仅选定字串" 项设置要翻译的字串
		If dlg.Seleted = 0 Then Set TransString = trn.String(j)
		If dlg.Seleted = 1 Then Set TransString = trn.String(j,pslSelection)

		'消息和字串初始化并获取原文和翻译字串
		Massage = ""
		LineMsg = ""
		AcckeyMsg = ""
		ChangeMsg = ""
		srcString = TransString.SourceText
		trnString = TransString.Text
		OldtrnString = trnString

		'删除检查标记
		If dlg.NoCheckTag = 1 Then TransString.Properties.RemoveAll

		'字串类型处理
		If dlg.AllType = 0 And dlg.Seleted = 0 Then
			If InStr("Menu|Dialog|StringTable|AcceleratorTable|Version",TransString.ResType) Then
				If InStr(StrTypes,TransString.ResType) = 0 Then GoTo Skip
			Else
				If dlg.Other = 0 Then GoTo Skip
			End If
		End If

		'跳过未翻译、已锁定和只读的字串
		If iVo <> 2 Then
			If srcString = trnString Then GoTo Skip
			If TransString.State(pslStateTranslated) = False Then GoTo Skip
		End If
		If TransString.State(pslStateLocked) = True Then GoTo Skip
		If TransString.State(pslStateReadOnly) = True Then GoTo Skip
		If Trim(srcString) = "" Then GoTo Skip

		'跳过已检查无错误并且检查日期晚于翻译日期的字串
		If iVo <> 2 And dlg.NoCheckTag = 0 And dlg.TagAll = 0 Then
			TranDate = TransString.DateTranslated
			If dlg.TagVer = 0 Then CheckVer = TransString.Property(Msg72)
			If dlg.TagVer = 1 Then CheckVer = Version
			If dlg.TagSet = 0 Then CheckSet = TransString.Property(Msg73)
			If dlg.TagSet = 1 Then CheckSet = CheckMacroSet
			If dlg.TagDate = 0 Then CheckDate = TransString.Property(Msg75)
			If dlg.TagDate = 1 Then CheckDate = TranDate
			If dlg.TagState = 0 Then CheckState = TransString.Property(Msg74)
			If dlg.TagState = 1 Then CheckState = ""
			If CheckState = Msg77 And CheckDate >= TranDate Then
				If CheckVer = Version And CheckSet = CheckMacroSet Then GoTo Skip
			End If
		End If

		'开始处理字串
		NewtrnString = CheckHanding(CheckID,srcString,trnString,TranLang)
		If dlg.AutoRepStr = 1 Then NewtrnString = ReplaceStr(CheckID,NewtrnString,0)

		'调用消息输出
		If iVo <> 2 Then
			If srcLineNum <> trnLineNum Then
				LineMsg = LineErrMassage(srcLineNum,trnLineNum,LineNumErrCount)
			End If
			If srcAccKeyNum <> trnAccKeyNum And (AllCont = 1 Or AccKey = 1) Then
				AcckeyMsg = AccKeyErrMassage(srcAccKeyNum,trnAccKeyNum,accKeyNumErrCount)
			End If
			'If NewtrnString <> OldtrnString Then
				ChangeMsg = ReplaceMassage(OldtrnString,NewtrnString)
			'End If
			Massage = ChangeMsg & AcckeyMsg & LineMsg
		ElseIf iVo = 2 And NewtrnString <> OldtrnString Then
			Massage = DelAccKeyMassage(OldtrnString,NewtrnString)
		End If
		If Massage <> "" Then TransString.OutputError(Massage)

		'替换或附加、删除翻译字串
		If NewtrnString <> OldtrnString Then
			If iVo = 1 Then TransString.OutputError(Msg65 & OldtrnString)
			If iVo <> 0 Then TransString.Text = NewtrnString
		End If

		'更改和原文不一致的译文字串状态,以方便查看
		If dlg.NoChangeState = 0 Then
			If NewtrnString <> OldtrnString Or srcLineNum <> trnLineNum Or srcAccKeyNum <> trnAccKeyNum Then
				TransString.State(pslStateReview) = True
			Else
				TransString.State(pslStateReview) = False
			End If
		End If

		'设置字串的自定义检查属性和检查日期
		If dlg.NoCheckTag = 0 Then
			TransString.Property(Msg72) = Version
			TransString.Property(Msg73) = CheckMacroSet
			TransString.Property(Msg75) = Format(Now,Msg79)
			If iVo = 0 Then
				If NewtrnString <> OldtrnString Or srcLineNum <> trnLineNum Or srcAccKeyNum <> trnAccKeyNum Then
					TransString.Property(Msg74) = Msg76
				Else
					TransString.Property(Msg74) = Msg77
				End If
			Else
				If NewtrnString <> OldtrnString Then
					TransString.Property(Msg74) = Msg78
				Else
					If srcLineNum <> trnLineNum Or srcAccKeyNum <> trnAccKeyNum Then
						TransString.Property(Msg74) = Msg76
					Else
						TransString.Property(Msg74) = Msg77
					End If
				End If
			End If
		End If
		Skip:
	Next j

	'错误计数及消息输出
	ErrorCount = ModifiedCount + AddedCount + WarningCount + LineNumErrCount + accKeyNumErrCount
	PSL.Output CountMassage(ErrorCount,LineNumErrCount,accKeyNumErrCount)
	If ErrorCount = 0 And TransListOpen = True Then trn.Close
	EndTimes = Timer
	PSL.Output Msg66 & Format(DateAdd("s",EndTimes - StartTimes,0),Msg67)

	'取消检查宏专用的用户定义属性的设置
	ExitSub:
	If Not trn Is Nothing Then
		If trn.Property(19980) = "CheckAccessKeys" Then
			trn.Property(19980) = ""
		End If
	End If
	On Error GoTo 0
	Exit Sub

	'显示程序错误消息
	SysErrorMsg:
	Call sysErrorMassage(Err)
	GoTo ExitSub
End Sub


'检测并下载新版本
Function Download(Method As String,Url As String,Async As String,Mode As String) As Boolean
	Dim i As Integer,n As Integer,m As Integer,k As Integer,updateINI As String
	Dim TempPath As String,File As String,OpenFile As Boolean,Body As Variant
	Dim xmlHttp As Object,UrlList() As String,Stemp As Boolean
	If OSLanguage = "0404" Then
		Msg01 = "岿粇"
		Msg02 = "穝ア毖"
		Msg03 = "代刚ア毖"
		Msg04 = "╰参⊿Τ杆 RAR 秆溃罽甅ノ祘Α礚猭秆溃罽更郎"
		Msg05 = "ぶゲ璶把计叫浪琩砞﹚い笆穝把计砞﹚"
		Msg06 = "秨币穝呼ア毖叫浪琩呼琌タ絋┪"
		Msg08 = "礚猭莉癟叫浪琩呼琌タ絋┪"
		Msg09 = "礚猭秆溃罽郎叫浪琩秆溃祘Α┪秆溃把计琌タ絋"
		Msg10 = "祘Α嘿: "
		Msg11 = "猂隔畖: "
		Msg12 = "磅︽把计: "
		Msg13 = "RAR 秆溃罽祘Αゼт琌祘Α隔畖岿粇┪砆簿埃"
		Msg14 = "笆穝代刚"
		Msg15 = "代刚Θ穝呼㎝秆溃祘Α把计タ絋"
		Msg16 = "ヘ玡呼セ %s"
		Msg17 = "祇瞷Τノ穝セ - %s琌更穝"
		Msg18 = "癟"
		Msg19 = "絋粄"
		Msg20 = "穝Θ祘Α盢挡挡叫穝币笆祘Α"
		Msg21 = "莉癟ぃセ癟叫浪琩叫浪琩呼琌タ絋"
		Msg22 = "穝ず甧:"
		Msg23 = "琌惠璶穝更穝"
		Msg24 = "タ代刚穝呼㎝秆溃祘Α把计叫祔..."
		Msg25 = "タ浪琩穝セ叫祔..."
		Msg26 = "タ更穝セ叫祔..."
		Msg27 = "タ秆溃罽叫祔..."
		Msg28 = "タ絋粄更祘Αセ叫祔..."
		Msg29 = "タ杆穝セ叫祔..."
		Msg30 = "眤╰参ぶ Microsoft.XMLHTTP ン礚猭穝"
	Else
		Msg01 = "错误"
		Msg02 = "更新失败！"
		Msg03 = "测试失败！"
		Msg04 = "系统没有安装 RAR 解压缩应用程序！无法解压缩下载文件。"
		Msg05 = "缺少必要的参数，请检查配置中的自动更新参数设置。"
		Msg06 = "打开更新网址失败！请检查网址是否正确或可访问。"
		Msg08 = "无法获取信息！请检查网址是否正确或可访问。"
		Msg09 = "无法解压缩文件！请检查解压程序或解压参数是否正确。"
		Msg10 = "程序名称: "
		Msg11 = "解析路径: "
		Msg12 = "运行参数: "
		Msg13 = "RAR 解压缩程序未找到！可能是程序路径错误或已被卸载。"
		Msg14 = "自动更新测试"
		Msg15 = "测试成功！更新网址和解压程序参数正确。"
		Msg16 = "目前网上的版本为 %s。"
		Msg17 = "发现有可用的新版本 - %s，是否下载更新？"
		Msg18 = "消息"
		Msg19 = "确认"
		Msg20 = "更新成功！程序将退出，退出后请重新启动程序。"
		Msg21 = "获取的信息不包含版本信息！请检查请检查网址是否正确。"
		Msg22 = "更新内容:"
		Msg23 = "是否需要重新下载并更新？"
		Msg24 = "正在测试更新网址和解压程序参数，请稍候..."
		Msg25 = "正在检查新版本，请稍候..."
		Msg26 = "正在下载新版本，请稍候..."
		Msg27 = "正在解压缩，请稍候..."
		Msg28 = "正在确认下载的程序版本，请稍候..."
		Msg29 = "正在安装新版本，请稍候..."
		Msg30 = "您的系统缺少 Microsoft.XMLHTTP 对象，无法更新！"
	End If
	Download = False
	OpenFile = False
	If Join(cUpdateSet) <> "" Then
		If Mode = "" Then Mode = cUpdateSet(0)
		If Url = "" Then Url = cUpdateSet(1)
		ExePath = cUpdateSet(2)
		Argument = cUpdateSet(3)
	End If
	If Mode = "2" Then Exit Function
	PSL.OutputWnd.Clear
	If Mode = "4" Then PSL.Output Msg24
	If Mode <> "4" Then PSL.Output Msg25
	If ExePath = "" Then
		If Mode <> "4" Then MsgBox(Msg02 & vbCrLf & Msg04,vbOkOnly+vbInformation,Msg01)
		If Mode = "4" Then MsgBox(Msg03 & vbCrLf & Msg04,vbOkOnly+vbInformation,Msg01)
		Exit Function
	End If
	If Url = "" Or Mode = "" Or Argument = "" Then
		If Mode <> "4" Then MsgBox(Msg02 & vbCrLf & Msg05,vbOkOnly+vbInformation,Msg01)
		If Mode = "4" Then MsgBox(Msg03 & vbCrLf & Msg05,vbOkOnly+vbInformation,Msg01)
		Exit Function
	End If
	TempPath = MacroDir & "\temp\"
	Set xmlHttp = CreateObject(DefaultObject)
	If xmlHttp Is Nothing Then Set xmlHttp = CreateObject("Msxml2.XMLHTTP")
	If Not xmlHttp Is Nothing Then
		If Mode = "4" Then GoTo getFile
		Stemp = False
		updateINIUrl = updateINIMainUrl & vbCrLf & updateINIMinorUrl
		UrlList = Split(updateINIUrl,vbCrLf)
		ReGetUpdateINIFile:
		For i = LBound(UrlList) To UBound(UrlList)
			updateINI = ""
			If UrlList(i) <> "" Then
				On Error GoTo SkipINIUrl
				xmlHttp.Open Method,UrlList(i),Async,User,Password
				xmlHttp.send()
				If xmlHttp.readyState = 4 Then updateINI = BytesToBstr(xmlHttp.responseBody,"utf-8")
				On Error GoTo 0
			End If
			If updateINI <> "" Then
				If InStr(LCase(updateINI),LCase(updateAppName)) Then
					Exit For
				Else
					updateINI = ""
				End If
			End If
			SkipINIUrl:
		Next i
		If Stemp = False And Url <> "" Then
			UrlList = Split(Url,vbCrLf)
			For i = LBound(UrlList) To UBound(UrlList)
				updateUrl = UrlList(i)
				n = InStrRev(LCase(updateUrl),"/download/")
				If n = 0 Then n = InStrRev(updateUrl,"/")
				If n <> 0 Then
					UrlList(i) = Left(updateUrl,n) & "update/" & updateINIFile
					Stemp = True
				End If
			Next i
			If Stemp = True Then GoTo ReGetUpdateINIFile
		End If
		xmlHttp.Abort
		If updateINI <> "" Then
			UrlList = Split(updateINI,vbCrLf)
			For i = LBound(UrlList) To UBound(UrlList)
				L$ = UrlList(i)
				If Left(Trim(L$),1) = "[" And Right(Trim(L$),1) = "]" Then
					Header = Mid(Trim(L$),2,Len(Trim(L$))-2)
				End If
				If L$ <> "" And Header = updateAppName Then
					If InStr(L$,"=") Then setPreStr = Trim(Left(L$,InStr(L$,"=")-1))
					If InStr(L$,"=") Then setAppStr = Mid(L$,InStr(L$,"=")+1)
					If setPreStr = "Version" Then NewVersion = Trim(setAppStr)
					If InStr(setPreStr,"URL_") Then
						Site = Trim(setAppStr)
						If UpdateSite <> "" Then UpdateSite = UpdateSite & vbCrLf & Site
						If UpdateSite = "" Then UpdateSite = Site
					End If
					If OSLanguage = "0404" And InStr(setPreStr,"Des_cht") Then
						Des = setAppStr
						If UpdateDes <> "" Then UpdateDes = UpdateDes & vbCrLf & Des
						If UpdateDes = "" Then UpdateDes = Des
					ElseIf OSLanguage <> "0404" And InStr(setPreStr,"Des_chs") Then
						Des = setAppStr
						If UpdateDes <> "" Then UpdateDes = UpdateDes & vbCrLf & Des
						If UpdateDes = "" Then UpdateDes = Des
					End If
				End If
				If Header <> updateAppName And NewVersion <> "" Then Exit For
			Next i
			If NewVersion <> "" Then
				If StrComp(NewVersion,Version) = 1 Then
					If Mode = "1" Or Mode = "3" Then
						Msg = Replace(Msg17,"%s",NewVersion)
						Msg = Msg & vbCrLf & vbCrLf & Msg22 & vbCrLf & UpdateDes
						OKMsg = MsgBox(Msg,vbYesNo+vbInformation,Msg19)
					End If
					If Mode <> "0" And ((Mode = "1" Or Mode = "3") And OKMsg <> vbYes) Then
						NewVersion = ""
					End If
				Else
					If Mode = "3" Then
						Msg = Replace(Msg16,"%s",NewVersion)
						OKMsg = MsgBox(Msg & vbCrLf & Msg23,vbYesNo+vbInformation,Msg19)
						If OKMsg = vbNo Then NewVersion = ""
					Else
						NewVersion = ""
					End If
				End If
			End If
		Else
			If Mode = "1" Or Mode = "3" Then MsgBox(Msg06,vbOkOnly+vbInformation,Msg01)
		End If
		If NewVersion = "" Then
			Set xmlHttp = Nothing
			Exit Function
		End If

		getFile:
		If Mode <> "4" Then PSL.Output Msg26
		If UpdateSite = "" Then UrlList = Split(Url,vbCrLf)
		If UpdateSite <> "" Then UrlList = ClearArray(Split(Url & vbCrLf & UpdateSite,vbCrLf))
		m = 0
		n = 0
		For i = LBound(UrlList) To UBound(UrlList)
			Body = ""
			On Error GoTo Skip
			xmlHttp.Open Method,UrlList(i),Async,User,Password
			xmlHttp.send()
			If xmlHttp.readyState = 4 Then Body = xmlHttp.responseBody
			On Error GoTo 0
			If LenB(Body) <= 0 Then m = m + 1
			Skip:
			If Err.Number <> 0 Then n = n + 1
			If n = UBound(UrlList) + 1 Then
				If Mode <> "4" Then MsgBox(Msg02 & vbCrLf & Msg06,vbOkOnly+vbInformation,Msg01)
				If Mode = "4" Then MsgBox(Msg03 & vbCrLf & Msg06,vbOkOnly+vbInformation,Msg01)
				xmlHttp.Abort
				Set xmlHttp = Nothing
				Exit Function
			End If
			If m = UBound(UrlList) + 1 Then
				If Mode <> "4" Then MsgBox(Msg02 & vbCrLf & Msg08,vbOkOnly+vbInformation,Msg01)
				If Mode = "4" Then MsgBox(Msg03 & vbCrLf & Msg08,vbOkOnly+vbInformation,Msg01)
				xmlHttp.Abort
				Set xmlHttp = Nothing
				Exit Function
			End If
			If LenB(Body) > 0 Then Exit For
		Next i
		xmlHttp.Abort
		Set xmlHttp = Nothing
		If Url <> Join(UrlList,vbCrLf) Then cUpdateSet(1) = Join(UrlList,vbCrLf)
	Else
		If Mode = "3" Then MsgBox(Msg30,vbOkOnly+vbInformation,Msg01)
		Exit Function
	End If

	File = TempPath & "temp.rar"
	If LenB(Body) > 0 Then
		On Error Resume Next
		If Dir(TempPath & "*.*") = "" Then MkDir TempPath
		On Error GoTo 0
		BytesToFile(Body,File)
	End If

	If Dir(File) <> "" Then
		If ExePath <> "" Then
			If UBound(Split(ExePath,"%",-1)) >= 2 Then
				AppExePath = Mid(ExePath,InStr(ExePath,"%")+1)
				Env = Left(AppExePath,InStr(AppExePath,"%")-1)
				ExePath = Replace(ExePath,"%" & Env & "%",Environ(Env),,1)
			End If
			ExePath = RemoveBackslash(ExePath,"""","""",1)
		End If
		If Argument <> "" Then
			If InStr(Argument,"""%1""") Then Argument = Replace(Argument,"%1",File)
			If InStr(Argument,"""%2""") Then Argument = Replace(Argument,"%2","*.bas")
			If InStr(Argument,"""%3""") Then Argument = Replace(Argument,"%3",TempPath)
			If InStr(Argument,"%1") Then Argument = Replace(Argument,"%1","" & File & "")
			If InStr(Argument,"%2") Then Argument = Replace(Argument,"%2","*.bas")
			If InStr(Argument,"%3") Then Argument = Replace(Argument,"%3","" & TempPath & "")
		End If
		If ExePath <> "" And Dir(ExePath) <> "" Then
			If Mode <> "4" Then PSL.Output Msg27
			Set WshShell = CreateObject("WScript.Shell")
			Return = WshShell.Run("""" & ExePath & """ " & Argument,0,True)
			Set WshShell = Nothing
			If Return = 0 Then OpenFile = True
			If Return <> 0 Then
				If Mode <> "4" Then MsgBox(Msg02 & vbCrLf & Msg09,vbOkOnly+vbInformation,Msg01)
				If Mode = "4" Then MsgBox(Msg03 & vbCrLf & Msg09,vbOkOnly+vbInformation,Msg01)
			End If
		ElseIf ExePath <> "" And Dir(ExePath) = "" Then
			Msg = Msg10 & ExeName & vbCrLf & Msg11 & ExePath & vbCrLf & Msg12 & _
					Argument & vbCrLf & vbCrLf & Msg13
			MsgBox(Msg,vbOkOnly+vbInformation,Msg01)
		End If
	End If

	If OpenFile = True Then
		If Mode <> "4" Then PSL.Output Msg28
		File = TempPath & updateMainFile
		Open File For Input As #1
		Do While Not EOF(1)
			Line Input #1,L$
			FindStr = "Private Const Version = "
			n = InStr(L$,FindStr)
			If n <> 0 Then
				WebVersion = Mid(L$,n+Len(FindStr)+1,10)
				Exit Do
			End If
		Loop
		Close #1

		If WebVersion <> "" Then
			If Mode = "4" Then
				Msg = Replace(Msg16,"%s",WebVersion)
				MsgBox(Msg15 & vbCrLf & Msg,vbOkOnly+vbInformation,Msg14)
				Download = True
			ElseIf StrComp(WebVersion,Version) = 1 Or (Mode = "3" And OKMsg = vbYes) Then
				If (Mode = "1" Or Mode = "3") And NewVersion = "" Then
					Msg = Replace(Msg17,"%s",WebVersion)
					OKMsg = MsgBox(Msg,vbYesNo+vbInformation,Msg19)
				End If
				If Mode = "0" Or ((Mode = "1" Or Mode = "3") And OKMsg = vbYes) Then
					If Mode <> "4" Then PSL.Output Msg29
					File = Dir$(TempPath & "*.bas")
					Do While File <> ""
						FileCopy TempPath & File,MacroDir & "\" & File
						Kill TempPath & File
						File = Dir$(TempPath & "*.bas")
					Loop
					MsgBox(Msg20,vbOkOnly+vbInformation,Msg19)
					Download = True
				End If
			End If
		Else
			MsgBox(Msg21,vbOkCancel+vbInformation,Msg01)
		End If
	End If

	File = Dir$(TempPath & "*.*")
	If File <> "" Then
		Do While File <> ""
			Kill TempPath & File
			File = Dir$(TempPath & "*.*")
		Loop
		If Dir(TempPath & "*.*") = "" Then RmDir TempPath
	End If
End Function


'从注册表中获取 RAR 扩展名的默认程序
Function getCMDPath(ExtName As String,CmdPath As String,Argument As String) As String
	Dim ExePathStr As String,ExeArg As String,WshShell As Object
	Set WshShell = CreateObject("WScript.Shell")
	If Not WshShell Is Nothing Then
		On Error Resume Next
		strKeyPath = "HKCR\" & ExtName & "\"
		ExtCmdStr = WshShell.RegRead(strKeyPath)
		If ExtCmdStr <> "" Then
			ExePathStr = WshShell.RegRead("HKCR\" & ExtCmdStr & "\shell\open\command\")
		End If
		On Error GoTo 0
		Set WshShell = Nothing
		If ExePathStr <> "" Then
			If UBound(Split(ExePathStr,"%",-1)) >= 2 Then
				AppExePathStr = Mid(ExePathStr,InStr(ExePathStr,"%")+1)
				Env = Left(AppExePathStr,InStr(AppExePathStr,"%")-1)
				ExePathStr = Replace(ExePathStr,"%" & Env & "%",Environ(Env),,1)
			End If

			PreExePath = Left(ExePathStr,InStrRev(ExePathStr,"\")+1)
			Temp = Mid(ExePathStr,Len(PreExePath)+1)
			If InStr(Temp," ") Then
				AppExePath = Left(Temp,InStr(Temp," ")-1)
			Else
				AppExePath = Temp
			End If
			ExePath = RemoveBackslash(PreExePath & AppExePath,"""","""",1)
			ExeArg = Mid(ExePathStr,Len(PreExePath & AppExePath)+1)

			If InStr(ExePath,"\") = 0 Then
				If Dir(Environ("SystemRoot") & "\system32\" & ExePath) <> "" Then
					ExePath = Environ("SystemRoot") & "\system32\" & ExePath
				ElseIf Dir(Environ("SystemRoot") & "\" & ExePath) <> "" Then
					ExePath = Environ("SystemRoot") & "\" & ExePath
				End If
			End If

			If InStr(LCase(ExePath),"winrar.exe") Then
				If ExeArg <> "" Then
					If InStr(ExeArg,"""%1""") Then
						ExeArg = "e " & Replace(ExeArg,"""%1""","""%1"" %2 ""%3""")
					ElseIf InStr(ExeArg,"%1") Then
						ExeArg = "e " & Replace(ExeArg,"%1","""%1"" %2 ""%3""")
					Else
						ExeArg = "e ""%1"" %2 ""%3"" " & ExeArg
					End If
				Else
					ExeArg = "e ""%1"" %2 ""%3"""
				End If
			ElseIf InStr(LCase(ExePath),"winzip.exe") Then
				ExePath = Replace(LCase(ExePath),"winzip.exe","wzunzip.exe")
				If ExeArg <> "" Then
					If InStr(ExeArg,"""%1""") Then
						ExeArg = Replace(ExeArg,"""%1""","""%1"" %2 ""%3""")
					ElseIf InStr(ExeArg,"%1") Then
						ExeArg = Replace(ExeArg,"%1","""%1"" %2 ""%3""")
					Else
						ExeArg = """%1"" %2 ""%3"" " & ExeArg
					End If
				Else
					ExeArg = " ""%1"" %2 ""%3"""
				End If
			ElseIf InStr(LCase(ExePath),"7z.exe") Then
				If ExeArg <> "" Then
					If InStr(ExeArg,"""%1""") Then
						ExeArg = "e " & Replace(ExeArg,"""1%""","""%1"" -o""%3"" %2")
					ElseIf InStr(ExeArg,"%1") Then
						ExeArg = "e " & Replace(ExeArg,"1%","""%1"" -o""%3"" %2")
					Else
						ExeArg = "e ""%1"" -o""%3"" %2 " & ExeArg
					End If
				Else
					ExeArg = "e ""%1"" -o""%3"" %2"
				End If
			End If

			If ExePath <> "" And Dir(ExePath) <> "" Then
				CmdPath = ExePath
				Argument = ExeArg
				getCMDPath = ExePath & JoinStr & ExeArg
			End If
		End If
	End If
End Function


'转换二进制数据为指定编码格式的字符
Function BytesToBstr(strBody As Variant,outCode As String) As String
    Dim objStream As Object
    Set objStream = CreateObject("Adodb.Stream")
    With objStream
    	.Type = 1
    	.Mode = 3
    	.Open
    	.Write strBody
    	.Position = 0
    	.Type = 2
    	.Charset = outCode
    	BytesToBstr = .ReadText
    	.Close
    End With
    Set objStream = Nothing
End Function


'写入二进制数据到文件
Function BytesToFile(strBody As Variant,File As String) As String
	Dim objStream As Object
	Set objStream = CreateObject("ADODB.Stream")
	If Not objStream Is Nothing Then
		objStream.Type = 1
		objStream.Mode = 3
		objStream.Open
		objStream.Write(strBody)
		objStream.Position = 0
		objStream.SaveToFile File,2
		objStream.Flush
		objStream.Close
		Set objStream = Nothing
	End If
End Function


'检查字串是否仅包含数字和符号
Function CheckStr(textStr As String,AscRange As String) As Boolean
	Dim i As Integer,j As Integer,n As Integer,InpAsc As Long
	Dim Pos As Integer,Min As Long,Max As Long
	CheckStr = False
	If Len(Trim(textStr)) = 0 Then Exit Function
	n = 0
	For i = 1 To Len(textStr)
		InpAsc = AscW(Mid(textStr,i,1))
		AscValue = Split(AscRange,",",-1)
		For j = 0 To UBound(AscValue)
			Pos = InStr(AscValue(j),"-")
			If Pos <> 0 Then
				Min = CLng(Left(AscValue(j),Pos-1))
				Max = CLng(Mid(AscValue(j),Pos+1))
			Else
				Min = CLng(AscValue(j))
				Max = CLng(AscValue(j))
			End If
			If InpAsc >= Min And InpAsc <= Max Then n = n + 1
		Next j
	Next i
	If n = Len(textStr) Then CheckStr = True
End Function


'替换特定字符
Function ReplaceStr(CheckID As Integer,trnStr As String,fType As Integer) As String
	Dim i As Integer,BaktrnStr As String
	'获取选定配置的参数
	TempArray = Split(CheckDataList(CheckID),JoinStr)
	SetsArray = Split(TempArray(1),SubJoinStr)
	If InStr(SetsArray(11),rSubJoinStr) Then
		PreStr = Replace(SetsArray(11),rSubJoinStr,SubJoinStr)
	Else
		PreStr = SetsArray(11)
	End If
	If fType <> 0 Then AutoRepChar = PreStr
	If fType = 0 Then AutoRepChar = SetsArray(12)
	BaktrnStr = trnStr
	PreRepStr = ""
	AppRepStr = ""
	If AutoRepChar <> "" Then
		FindStrArr = Split(Convert(AutoRepChar),",",-1)
		For i = LBound(FindStrArr) To UBound(FindStrArr)
			FindStr = FindStrArr(i)
			If InStr(FindStr,"|") Then
				TempArray = Split(FindStr,"|")
				If fType < 2 Then
					PreStr = TempArray(0)
					AppStr = TempArray(1)
				Else
					PreStr = TempArray(1)
					AppStr = TempArray(0)
				End If
			End If
			If PreStr <> "" And InStr(BaktrnStr,PreStr) Then
				BaktrnStr = Replace(BaktrnStr,PreStr,AppStr)
				If fType = 0 Then
					If PreRepStr <> "" Then
						If InStr(PreRepStr,PreStr) = 0 Then PreRepStr = PreRepStr & JoinStr & PreStr
					Else
						PreRepStr = PreStr
					End If
					If AppRepStr <> "" Then
						If InStr(AppRepStr,AppStr) = 0 Then AppRepStr = AppRepStr & JoinStr & AppStr
					Else
						AppRepStr = AppStr
					End If
				End If
			End If
		Next i
	End If
	ReplaceStr = BaktrnStr
End Function


'检查修正快捷键、终止符和加速器
Function CheckHanding(CheckID As Integer,srcStr As String,trnStr As String,TranLang As String) As String
	Dim i As Integer,BaksrcStr As String,BaktrnStr As String,srcStrBak As String,trnStrBak As String
	Dim srcNum As Integer,trnNum As Integer,srcSplitNum As Integer,trnSplitNum As Integer
	Dim FindStrArr() As String,srcStrArr() As String,trnStrArr() As String,LineSplitArr() As String
	Dim posinSrc As Integer,posinTrn As Integer

	'获取选定配置的参数
	TempArray = Split(CheckDataList(CheckID),JoinStr)
	SetsArray = Split(TempArray(1),SubJoinStr)
	ExcludeChar = SetsArray(0)
	LineSplitChar = SetsArray(1)
	KeepCharPair = SetsArray(3)

	'按字符长度排序
	If LineSplitChar <> "" Then
		FindStrArr = Split(LineSplitChar,",",-1)
		FindStrArr = SortArray(FindStrArr,0,"Lenght",">")
		LineSplitChar = Join(FindStrArr,",")
	End If

	'参数初始化
	srcNum = 0
	trnNum = 0
	srcSplitNum = 0
	trnSplitNum = 0
	srcLineNum = 0
	trnLineNum = 0
	srcAccKeyNum = 0
	trnAccKeyNum = 0
	BaksrcStr = srcStr
	BaktrnStr = trnStr
	posinSrc = InStrRev(srcStr,"&")
	posinTrn = InStrRev(trnStr,"&")

	'排除字串中的非快捷键
	If (posinSrc <> 0 Or posinTrn <> 0) And ExcludeChar <> "" Then
		FindStrArr = Split(Convert(ExcludeChar),",",-1)
		For i = LBound(FindStrArr) To UBound(FindStrArr)
			FindStr = LTrim(FindStrArr(i))
			If FindStr <> "" Then
				BaksrcStr = Replace(BaksrcStr,FindStr,"*a" & i & "!N!" & i & "d*")
				BaktrnStr = Replace(BaktrnStr,FindStr,"*a" & i & "!N!" & i & "d*")
			End If
		Next i
	End If

	'过滤不是快捷键的快捷键
	If (posinSrc <> 0 Or posinTrn <> 0) And KeepCharPair <> "" Then
		FindStrArr = Split(Convert(KeepCharPair),",",-1)
		For i = LBound(FindStrArr) To UBound(FindStrArr)
			FindStr = Trim(FindStrArr(i))
			If FindStr <> "" Then
				LFindStr = Trim(Left(FindStr,1))
				RFindStr = Trim(Right(FindStr,1))
				ToRepStr = LFindStr & "&" & RFindStr
				BeRepStr = LFindStr & "*!N!" & i & "!M!" & i & "!N!*" & RFindStr
				BaksrcStr = Replace(BaksrcStr,ToRepStr,BeRepStr)
				BaktrnStr = Replace(BaktrnStr,ToRepStr,BeRepStr)
			End If
		Next i
	End If

	'用替换法拆分字串
	If LineSplitChar <> "" Then
		srcStrBak = BaksrcStr
		trnStrBak = BaktrnStr
		FindStrArr = Split(Convert(LineSplitChar),",",-1)
		For i = UBound(FindStrArr) To LBound(FindStrArr) Step -1
			FindStr = Trim(FindStrArr(i))
			If InStr(BaksrcStr,FindStr) Then srcNum = UBound(Split(BaksrcStr,FindStr,-1))
			If InStr(BaktrnStr,FindStr) Then trnNum = UBound(Split(BaktrnStr,FindStr,-1))
			If srcNum = trnNum And srcNum <> 0 And trnNum <> 0 Then
				If FindStr <> "&" Then
					srcStrBak = Replace(srcStrBak,FindStr,"*c!N!g*")
					trnStrBak = Replace(trnStrBak,FindStr,"*c!N!g*")
				ElseIf FindStr = "&" Then
					srcStrBak = Insert(srcStrBak,FindStr,"*c!N!g*",0)
					trnStrBak = Insert(trnStrBak,FindStr,"*c!N!g*",1)
				End If
			End If
		Next i
	End If
	srcStrArr = Split(srcStrBak,"*c!N!g*",-1)
	trnStrArr = Split(trnStrBak,"*c!N!g*",-1)

	'字串处理
	srcSplitNum = UBound(srcStrArr)
	trnSplitNum = UBound(trnStrArr)
	If srcSplitNum = 0 And trnSplitNum = 0 Then
		BaktrnStr = StringReplace(CheckID,BaksrcStr,BaktrnStr,TranLang)
	ElseIf srcSplitNum <> 0 Or trnSplitNum <> 0 Then
		LineSplitArr = MergeArray(srcStrArr,trnStrArr)
		BaktrnStr = ReplaceStrSplit(CheckID,BaktrnStr,LineSplitArr,TranLang)
	End If

	'计算行数
	LineSplitChars = "\r\n,\r,\n"
	FindStrArr = Split(Convert(LineSplitChars),",",-1)
	For i = LBound(FindStrArr) To UBound(FindStrArr)
		FindStr = Trim(FindStrArr(i))
		If InStr(BaksrcStr,FindStr) Then srcLineNum = UBound(Split(BaksrcStr,FindStr,-1))
		If InStr(BaktrnStr,FindStr) Then trnLineNum = UBound(Split(BaktrnStr,FindStr,-1))
	Next i

	'计算快捷键数
	If InStr(BaksrcStr,"&") Then srcAccKeyNum = UBound(Split(BaksrcStr,"&",-1))
	If InStr(BaktrnStr,"&") Then trnAccKeyNum = UBound(Split(BaktrnStr,"&",-1))

	'还原不是快捷键的快捷键
	If (posinSrc <> 0 Or posinTrn <> 0) And KeepCharPair <> "" Then
		FindStrArr = Split(Convert(KeepCharPair),",",-1)
		For i = LBound(FindStrArr) To UBound(FindStrArr)
			FindStr = Trim(FindStrArr(i))
			If FindStr <> "" Then
				LFindStr = Trim(Left(FindStr,1))
				RFindStr = Trim(Right(FindStr,1))
				ToRepStr = LFindStr & "*!N!" & i & "!M!" & i & "!N!*" & RFindStr
				BeRepStr = LFindStr & "&" & RFindStr
				BaksrcStr = Replace(BaksrcStr,ToRepStr,BeRepStr)
				BaktrnStr = Replace(BaktrnStr,ToRepStr,BeRepStr)
			End If
		Next i
	End If

	'还原字串中被排除的非快捷键
	If (posinSrc <> 0 Or posinTrn <> 0) And ExcludeChar <> "" Then
		FindStrArr = Split(Convert(ExcludeChar),",",-1)
		For i = LBound(FindStrArr) To UBound(FindStrArr)
			FindStr = LTrim(FindStrArr(i))
			If FindStr <> "" Then
				BaksrcStr = Replace(BaksrcStr,"*a" & i & "!N!" & i & "d*",FindStr)
				BaktrnStr = Replace(BaktrnStr,"*a" & i & "!N!" & i & "d*",FindStr)
			End If
		Next i
	End If
	CheckHanding = BaktrnStr
End Function


' 按行获取字串的各个字段并替换翻译字符串
Function StringReplace(CheckID As Integer,srcStr As String,trnStr As String,TranLang As String) As String
	Dim posinSrc As Integer,posinTrn As Integer,StringSrc As String,StringTrn As String
	Dim accesskeySrc As String,accesskeyTrn As String,Temp As String
	Dim ShortcutPosSrc As Integer,ShortcutPosTrn As Integer,PreTrn As String
	Dim EndStringPosSrc As Integer,EndStringPosTrn As Integer,AppTrn As String
	Dim preKeyTrn As String,appKeyTrn As String,Stemp As Boolean,FindStrArr() As String
	Dim i As Integer,j As Integer,x As Integer,y As Integer,m As Integer,n As Integer

	'获取选定配置的参数
	TempArray = Split(CheckDataList(CheckID),JoinStr)
	SetsArray = Split(TempArray(1),SubJoinStr)
	CheckBracket = SetsArray(2)
	AsiaKey = SetsArray(4)
	CheckEndChar = SetsArray(5)
	NoTrnEndChar = SetsArray(6)
	AutoTrnEndChar = SetsArray(7)
	CheckShortChar = SetsArray(8)
	CheckShortKey = SetsArray(9)
	KeepShortKey = SetsArray(10)

	'按字符长度排序
	If CheckEndChar <> "" Then
		FindStrArr = Split(CheckEndChar,-1)
		FindStrArr = SortArray(FindStrArr,0,"Lenght",">")
		CheckEndChar = Join(FindStrArr)
	End If

	'参数初始化
	acckeySrc = ""
	acckeyTrn = ""
	ShortcutSrc = ""
	ShortcutTrn = ""
	EndStringSrc = ""
	EndStringTrn = ""
	ExpStringTrn = ""
	SpaceTrn = ""
	PreStringTrn = ""
	AppStringTrn = ""
	EndSpaceSrc = ""
	EndSpaceTrn = ""
	ShortcutPosSrc = 0
	ShortcutPosTrn = 0
	EndStringPosSrc = 0
	EndStringPosTrn = 0

	'提取字串末尾空格
	EndSpaceSrc = Space(Len(srcStr) - Len(RTrim(srcStr)))
	EndSpaceTrn = Space(Len(trnStr) - Len(RTrim(trnStr)))

	'获取加速器
	If CheckShortChar <> "" Then
		FindStrArr = Split(Convert(CheckShortChar),",",-1)
		For i = LBound(FindStrArr) To UBound(FindStrArr)
			FindStr = Trim(FindStrArr(i))
			If FindStr <> "" Then
				For n = 0 To 1
					If n = 0 Then Temp = srcStr
					If n = 1 Then Temp = trnStr
					x = UBound(Split(Temp,FindStr,-1))
					y = 0
					Shortcut = ""
					ShortcutPos = y
					If x = 1 Then
						If Left(LTrim(Temp),Len(FindStr)) <> FindStr Then
							y = InStrRev(Temp,FindStr)
						End If
					ElseIf x > 1 Then
						y = InStrRev(Temp,FindStr)
					End If
					If y <> 0 Then
						ShortcutKey = Trim(Mid(Temp,y+1))
						If Trim(ShortcutKey) = "+" Then
							If CheckKeyCode(ShortcutKey,CheckShortKey) <> 0 Then
								Shortcut = Trim(Mid(Temp,y))
								ShortcutPos = y
							End If
						Else
							x = 0
							KeyArr = Split(ShortcutKey,"+",-1)
							For j = LBound(KeyArr) To UBound(KeyArr)
								x = x + CheckKeyCode(KeyArr(j),CheckShortKey)
							Next j
							If x <> 0 And x >= UBound(KeyArr) Then
								Shortcut = Trim(Mid(Temp,y))
								ShortcutPos = y
							End If
						End If
						If n = 0 And ShortcutSrc = "" Then
							ShortcutSrc = Shortcut
							ShortcutPosSrc = ShortcutPos
						ElseIf n = 1 And ShortcutTrn = "" Then
							ShortcutTrn = Shortcut
							ShortcutPosTrn = ShortcutPos
						End If
					End If
				Next n
			End If
			If ShortcutSrc <> "" And ShortcutTrn <> "" Then Exit For
		Next i
	End If

	'获取终止符，下列字符均会被检查
	If CheckEndChar <> "" Then
		FindStrArr = Split(Convert(CheckEndChar),-1)
		For i = UBound(FindStrArr) To LBound(FindStrArr) Step -1
			FindStr = Trim(FindStrArr(i))
			If FindStr <> "" Then
				If InStr(FindStr,"*") Or InStr(FindStr,"?") Or InStr(FindStr,"#") Then
					PreFindStr = Left(FindStr,1)
				Else
					PreFindStr = FindStr
				End If
				If PreFindStr <> "" Then
					For n = 0 To 1
						If n = 0 Then
							Temp = srcStr
							EndSpace = EndSpaceSrc
							Shortcut = ShortcutSrc
						ElseIf n = 1 Then
							Temp = trnStr
							EndSpace = EndSpaceTrn
							Shortcut = ShortcutTrn
						End If
						x = Len(Shortcut & EndSpace)
						If Len(Temp) > x Then Temp = Left(Temp,Len(Temp) - x)
						y = InStrRev(Temp,PreFindStr)
						EndString = ""
						EndStringPos = 0
						If y <> 0 Then
							PreStr = Left(Temp,y - 1)
							AppStr = Mid(Temp,y)
							x = Len(PreStr) - Len(RTrim(PreStr))
							If AppStr <> "" And Trim(AppStr) Like FindStr Then
								EndString = Space(x) & AppStr
								EndStringPos = y - x
								If n = 0 And EndStringSrc = "" Then
									EndStringSrc = EndString
									EndStringPosSrc = EndStringPos
								ElseIf n = 1 And EndStringTrn = "" Then
									EndStringTrn = EndString
									EndStringPosTrn = EndStringPos
								End If
							End If
						End If
					Next n
				End If
			End If
			If EndStringSrc <> "" And EndStringTrn <> "" Then Exit For
		Next i
	End If

	'获取原文和翻译的快捷键位置
	posinSrc = InStrRev(srcStr,"&")
	posinTrn = InStrRev(trnStr,"&")

	'获取原文和翻译的快捷键 (包括快捷键符前后的字符)
	If posinSrc <> 0 Then accesskeySrc = Mid(srcStr,posinSrc + 1,1)
	If posinTrn <> 0 Then accesskeyTrn = Mid(trnStr,posinTrn + 1,1)
	If (posinSrc <> 0 Or posinTrn <> 0) And CheckBracket <> "" Then
		FindStrArr = Split(Convert(CheckBracket),",",-1)
		For i = LBound(FindStrArr) To UBound(FindStrArr)
			FindStr = Trim(FindStrArr(i))
			If FindStr <> "" Then
				LFindStr = Trim(Left(FindStr,1))
				RFindStr = Trim(Right(FindStr,1))
				For n = 0 To 1
					If n = 0 Then
						Temp = srcStr
						j = posinSrc
						EndSpace = EndSpaceSrc
						Shortcut = ShortcutSrc
						EndString = EndStringSrc
					ElseIf n = 1 Then
						Temp = trnStr
						j = posinTrn
						EndSpace = EndSpaceTrn
						Shortcut = ShortcutTrn
						EndString = EndStringTrn
					End If
					AccessKey = ""
					If j > 1 Then
						If InStrRev(Temp,LFindStr & Mid(Temp,j,2) & RFindStr) Then
							j = j - 1
							AccessKey = Mid(Temp,j,4)
						Else
							PreStr = Left(Temp,j-1)
							AppStr = Mid(Temp,j+2)
							x = Len(PreStr) - Len(RTrim(PreStr))
							y = Len(AppStr) - Len(LTrim(AppStr))
							If Right(RTrim(PreStr),1) = LFindStr And Left(LTrim(AppStr),1) = RFindStr Then
								j = j-x-1
								AccessKey = Mid(Temp,j,x+y+4)
							Else
								AccessKey = Mid(Temp,j,1)
							End If
						End If
					Else
						If j <> 0 Then AccessKey = Mid(Temp,j,1)
					End If
					If n = 0 And acckeySrc = "" Then
						posinSrc = j
						acckeySrc = AccessKey
					ElseIf n = 1 And acckeyTrn = "" Then
						posinTrn = j
						acckeyTrn = AccessKey
					End If
				Next n
			End If
			If acckeySrc <> "" And acckeyTrn <> "" Then Exit For
		Next i
	End If

	'获取快捷键后面的非终止符和非加速器的字符，这些字符将被移动到快捷键前
	If posinTrn <> 0 Then
		x = Len(EndStringTrn & ShortcutTrn & EndSpaceTrn)
		If InStr(ShortcutTrn,"&") Then x = Len(EndSpaceTrn)
		If InStr(EndStringTrn,"&") Then x = Len(ShortcutTrn & EndSpaceTrn)
		If Len(trnStr) > x Then
			Temp = Left(trnStr,Len(trnStr) - x)
			ExpStringTrn = Mid(Temp,posinTrn + Len(acckeyTrn))
		End If
	End If

	'去除快捷键或终止符或加速器前面的空格
	x = Len(acckeyTrn & ExpStringTrn & EndStringTrn & ShortcutTrn & EndSpaceTrn)
	If InStr(EndStringTrn & ShortcutTrn,"&") Then
		x = Len(ExpStringTrn & EndStringTrn & ShortcutTrn & EndSpaceTrn)
	End If
	If Len(trnStr) > x Then
		Temp = Left(trnStr,Len(trnStr) - x)
		y = Len(Temp) - Len(RTrim(Temp))
		SpaceTrn = Space(y)
		If y > 0 And ExpStringTrn <> "" Then
			If CheckStr(Left(Trim(ExpStringTrn),1),"0-255") = True Then SpaceTrn = Space(y-1)
		End If
	End If

	'获取翻译中快捷键前的终止符
	If (acckeySrc <> "" Or acckeyTrn <> "") And CheckEndChar <> "" Then
		x = Len(SpaceTrn & acckeyTrn & ExpStringTrn & EndStringTrn & ShortcutTrn & EndSpaceTrn)
		If InStr(EndStringTrn & ShortcutTrn,"&") Then
			x = Len(SpaceTrn & ExpStringTrn & EndStringTrn & ShortcutTrn & EndSpaceTrn)
		End If
		If Len(trnStr) > x Then
			Temp = Left(trnStr,Len(trnStr) - x)
			FindStrArr = Split(Convert(CheckEndChar),-1)
			For i = UBound(FindStrArr) To LBound(FindStrArr) Step -1
				FindStr = Trim(FindStrArr(i))
				If InStr(FindStr,"*") Or InStr(FindStr,"?") Or InStr(FindStr,"#") Then
					PreFindStr = Left(FindStr,1)
				Else
					PreFindStr = FindStr
				End If
				y = InStrRev(Temp,PreFindStr)
				If y <> 0 And PreFindStr <> "" Then
					PreStr = Left(Temp,y - 1)
					AppStr = Mid(Temp,y)
					x = Len(PreStr) - Len(RTrim(PreStr))
					If AppStr <> "" And Trim(AppStr) Like FindStr Then
						PreStringTrn = Space(x) & AppStr
					End If
				End If
				If PreStringTrn <> "" Then Exit For
			Next i
		End If
	End If

	'自动翻译符合条件的终止符
	If (EndStringSrc And AutoTrnEndChar) <> "" Then
		FindStrArr = Split(Convert(AutoTrnEndChar),-1)
		For i = LBound(FindStrArr) To UBound(FindStrArr)
			FindStr = Trim(FindStrArr(i))
			If InStr(FindStr,"|") Then
				TempArray = Split(FindStr,"|")
				LFindStr = TempArray(0)
				RFindStr = TempArray(1)
				If Trim(EndStringSrc) = LFindStr Then
					EndStringSrc = RFindStr
					Exit For
				End If
			End If
		Next i
	End If

	'要被保留的终止符组合
	If EndStringSrc <> "" And NoTrnEndChar <> "" Then
		PreEndStringSrc = Left(srcStr,EndStringPosSrc + Len(RTrim(EndStringSrc)))
		If CheckKeyCode(PreEndStringSrc,NoTrnEndChar) = 1 Then
			EndStringSrc = ""
		End If
	End If
	If EndStringTrn <> "" And NoTrnEndChar <> "" Then
		PreEndStringTrn = Left(trnStr,EndStringPosTrn + Len(RTrim(EndStringTrn)))
		If CheckKeyCode(PreEndStringTrn,NoTrnEndChar) = 1 Then
			EndStringTrn = ""
		End If
	End If

	'保留符合条件的加速器翻译
	If ShortcutSrc <> "" And ShortcutTrn <> "" And KeepShortKey <> "" Then
		srcNum = 0
		trnNum = 0
		If InStr(ShortcutSrc,"+") And Trim(ShortcutSrc) <> "+" Then
			SrcKeyArr = Split(ShortcutSrc,"+",-1)
			srcNum = UBound(SrcKeyArr)
		End If
		If InStr(ShortcutTrn,"+") And Trim(ShortcutTrn) <> "+" Then
			TrnKeyArr = Split(ShortcutTrn,"+",-1)
			trnNum = UBound(TrnKeyArr)
		End If
		If srcNum > trnNum Then Num = srcNum
		If srcNum < trnNum Then Num = trnNum
		If srcNum = trnNum Then Num = srcNum
		If srcNum <> 0 And trnNum <> 0 Then
			For i = 0 To Num
				SrcKey = ""
				TrnKey = ""
				If i <= srcNum Then SrcKey = Trim(SrcKeyArr(i))
				If i <= trnNum Then TrnKey = Trim(TrnKeyArr(i))
				If SrcKey <> "" And TrnKey <> "" Then
					If CheckKeyCode(TrnKey,KeepShortKey) <> 0 Then
						ShortcutSrc = Replace(ShortcutSrc,SrcKey,TrnKey)
					End If
				End If
			Next i
		End If
	End If

	'字串内容选择处理
	If AllCont <> 1 Then
		If AccKey = 1 And EndChar = 1 And Acceler <> 1 Then
			ShortcutSrc = ShortcutTrn
		ElseIf AccKey = 1 And EndChar <> 1 And Acceler = 1 Then
			EndStringSrc = EndStringTrn
		ElseIf AccKey <> 1 And EndChar = 1 And Acceler = 1 Then
			SpaceTrn = ""
			ExpStringTrn = ""
			PreStringTrn = ""
			acckeySrc = acckeyTrn
		ElseIf AccKey = 1 And EndChar <> 1 And Acceler <> 1 Then
			EndStringSrc = EndStringTrn
			ShortcutSrc = ShortcutTrn
		ElseIf AccKey <> 1 And EndChar = 1 And Acceler <> 1 Then
			SpaceTrn = ""
			ExpStringTrn = ""
			PreStringTrn = ""
			acckeySrc = acckeyTrn
			ShortcutSrc = ShortcutTrn
		ElseIf AccKey <> 1 And EndChar <> 1 And Acceler = 1 Then
			SpaceTrn = ""
			ExpStringTrn = ""
			PreStringTrn = ""
			acckeySrc = acckeyTrn
			EndStringSrc = EndStringTrn
		End If
	End If

	'数据集成
	If AsiaKey = "1" And iVo <> 2 Then
		StringSrc = acckeySrc & EndStringSrc & ShortcutSrc & EndSpaceSrc
		If InStr(EndStringTrn & ShortcutTrn,"&") Then
			ExpStringTrn = ""
			StringTrn = PreStringTrn & SpaceTrn & EndStringTrn & ShortcutTrn & EndSpaceTrn
		Else
			StringTrn = PreStringTrn & SpaceTrn & acckeyTrn & ExpStringTrn & EndStringTrn & ShortcutTrn & EndSpaceTrn
		End If

		If acckeySrc <> "" Then acckeySrc = "(" & "&" & UCase(accesskeySrc) & ")"
		If acckeyTrn = "&" Then acckeyTrn = "(" & "&" & UCase(accesskeyTrn) & ")"
		If InStr(ShortcutSrc,"&") Then ShortcutSrc = Replace(ShortcutSrc,"&","")
		If Trim(PreStringTrn) <> "" And EndStringTrn = "" Then EndStringTrn = PreStringTrn

		If acckeySrc <> "" And EndStringSrc <> "" And ShortcutSrc <> "" Then
			NewStringTrn = Trim(ExpStringTrn) & acckeySrc & EndStringSrc & ShortcutSrc & EndSpaceSrc
		ElseIf acckeySrc = "" And EndStringSrc <> "" And ShortcutSrc <> "" Then
			NewStringTrn = Trim(ExpStringTrn) & acckeyTrn & EndStringSrc & ShortcutSrc & EndSpaceSrc
		ElseIf acckeySrc <> "" And EndStringSrc = "" And ShortcutSrc <> "" Then
			NewStringTrn = Trim(ExpStringTrn) & acckeySrc & RTrim(EndStringTrn) & ShortcutSrc & EndSpaceSrc
		ElseIf acckeySrc <> "" And EndStringSrc <> "" And ShortcutSrc = "" Then
			NewStringTrn = Trim(ExpStringTrn) & acckeySrc & EndStringSrc & ShortcutTrn & EndSpaceSrc

		ElseIf acckeySrc = "" And EndStringSrc = "" And ShortcutSrc <> "" Then
			NewStringTrn = Trim(ExpStringTrn) & acckeyTrn & RTrim(EndStringTrn) & ShortcutSrc & EndSpaceSrc
		ElseIf acckeySrc = "" And EndStringSrc <> "" And ShortcutSrc = "" Then
			NewStringTrn = Trim(ExpStringTrn) & acckeyTrn & EndStringSrc & ShortcutTrn & EndSpaceSrc
		ElseIf acckeySrc <> "" And EndStringSrc = "" And ShortcutSrc = "" Then
			NewStringTrn = Trim(ExpStringTrn) & acckeySrc & RTrim(EndStringTrn) & ShortcutTrn & EndSpaceSrc

		ElseIf acckeySrc = "" And EndStringSrc = "" And ShortcutSrc = "" Then
			NewStringTrn = Trim(ExpStringTrn) & acckeyTrn & RTrim(EndStringTrn) & ShortcutTrn & EndSpaceSrc
		End If
	ElseIf AsiaKey <> "1" And iVo <> 2 Then
		If InStr(EndStringTrn & ShortcutTrn,"&") Then
			StringSrc = EndStringSrc & ShortcutSrc & EndSpaceSrc
			StringTrn = SpaceTrn & EndStringTrn & ShortcutTrn & EndSpaceTrn
			If EndStringSrc <> "" And ShortcutSrc <> "" Then
				NewStringTrn = EndStringSrc & ShortcutSrc & EndSpaceSrc
			ElseIf EndStringSrc = "" And ShortcutSrc <> "" Then
				NewStringTrn = RTrim(EndStringTrn) & ShortcutSrc & EndSpaceSrc
			ElseIf EndStringSrc <> "" And ShortcutSrc = "" Then
				NewStringTrn = EndStringSrc & ShortcutTrn & EndSpaceSrc
			ElseIf EndStringSrc = "" And ShortcutSrc = "" Then
				NewStringTrn = RTrim(EndStringTrn) & ShortcutTrn & EndSpaceSrc
			End If
		Else
			StringSrc = acckeySrc & EndStringSrc & ShortcutSrc & EndSpaceSrc
			StringTrn = SpaceTrn & acckeyTrn & ExpStringTrn & EndStringTrn & ShortcutTrn & EndSpaceTrn
			If InStr(ShortcutSrc,"&") Then ShortcutSrc = Replace(ShortcutSrc,"&","")
			If acckeySrc <> "" And EndStringSrc <> "" And ShortcutSrc <> "" Then
				NewStringTrn = acckeyTrn & Trim(ExpStringTrn) & EndStringSrc & ShortcutSrc & EndSpaceSrc
			ElseIf acckeySrc = "" And EndStringSrc <> "" And ShortcutSrc <> "" Then
				NewStringTrn = acckeyTrn & Trim(ExpStringTrn) & EndStringSrc & ShortcutSrc & EndSpaceSrc
			ElseIf acckeySrc <> "" And EndStringSrc = "" And ShortcutSrc <> "" Then
				NewStringTrn = acckeyTrn & Trim(ExpStringTrn) & RTrim(EndStringTrn) & ShortcutSrc & EndSpaceSrc
			ElseIf acckeySrc <> "" And EndStringSrc <> "" And ShortcutSrc = "" Then
				NewStringTrn = acckeyTrn & Trim(ExpStringTrn) & EndStringSrc & ShortcutTrn & EndSpaceSrc

			ElseIf acckeySrc = "" And EndStringSrc = "" And ShortcutSrc <> "" Then
				NewStringTrn = acckeyTrn & Trim(ExpStringTrn) & RTrim(EndStringTrn) & ShortcutSrc & EndSpaceSrc
			ElseIf acckeySrc = "" And EndStringSrc <> "" And ShortcutSrc = "" Then
				NewStringTrn = acckeyTrn & Trim(ExpStringTrn) & EndStringSrc & ShortcutTrn & EndSpaceSrc
			ElseIf acckeySrc <> "" And EndStringSrc = "" And ShortcutSrc = "" Then
				NewStringTrn = acckeyTrn & Trim(ExpStringTrn) & RTrim(EndStringTrn) & ShortcutTrn & EndSpaceSrc

			ElseIf acckeySrc = "" And EndStringSrc = "" And ShortcutSrc = "" Then
				NewStringTrn = acckeyTrn & Trim(ExpStringTrn) & RTrim(EndStringTrn) & ShortcutTrn & EndSpaceSrc
			End If
		End If
		If acckeySrc <> "" Then acckeySrc = "&" & accesskeySrc
		If acckeyTrn = "&" Then acckeyTrn = "&" & accesskeyTrn
		'If LCase(acckeySrc) = LCase(acckeyTrn) Then acckeySrc = acckeyTrn
		If InStr(ShortcutSrc,"&") Then ShortcutSrc = Replace(ShortcutSrc,"&","")
		If InStr(ShortcutTrn,"&") Then ShortcutTrn = Replace(ShortcutTrn,"&","")
		ExpStringTrn = ""
	ElseIf acckeyTrn <> "" And iVo = 2 Then
		StringSrc = ""
		StringTrn = acckeyTrn
		If acckeyTrn = "&" & accesskeyTrn Then
			NewStringTrn = accesskeyTrn
		Else
			NewStringTrn = ""
		End If
		acckeySrc = "&" & UCase(accesskeySrc)
		acckeyTrn = "&" & UCase(accesskeyTrn)
	End If

	'PSL.Output "------------------------------ "      '调试用
	'PSL.Output "srcStr = " & srcStr                   '调试用
	'PSL.Output "trnStr = " & trnStr                   '调试用
	'PSL.Output "SpaceTrn = " & SpaceTrn               '调试用
	'PSL.Output "acckeySrc = " & acckeySrc             '调试用
	'PSL.Output "acckeyTrn = " & acckeyTrn             '调试用
	'PSL.Output "EndStringSrc = " & EndStringSrc       '调试用
	'PSL.Output "EndStringTrn = " & EndStringTrn       '调试用
	'PSL.Output "ShortcutSrc = " & ShortcutSrc         '调试用
	'PSL.Output "ShortcutTrn = " & ShortcutTrn         '调试用
	'PSL.Output "ExpStringTrn = " & ExpStringTrn       '调试用
	'PSL.Output "StringSrc = " & StringSrc             '调试用
	'PSL.Output "StringTrn = " & StringTrn             '调试用
	'PSL.Output "NewStringTrn = " & NewStringTrn       '调试用
	'PSL.Output "PreStringTrn = " & PreStringTrn       '调试用

	'字串替换
	Temp = trnStr
	If StringSrc <> StringTrn Then
		If StringTrn <> "" And StringTrn <> NewStringTrn Then
			x = InStrRev(Temp,StringTrn)
			If x <> 0 Then
				PreTrn = Left(Temp,x - 1)
				AppTrn = Mid(Temp,x)
				'PSL.Output "PreTrn = " & PreTrn       '调试用
				'PSL.Output "AppTrn = " & AppTrn       '调试用
				Temp = PreTrn & Replace(AppTrn,StringTrn,NewStringTrn)
			Else
				Temp = Replace(Temp,StringTrn,NewStringTrn)
			End If
		ElseIf StringTrn = "" And NewStringTrn <> "" Then
			Temp = RTrim(Temp) & NewStringTrn
		End If
	End If
	If AsiaKey <> "1" And iVo <> 2 Then
		If acckeyTrn <> "" And acckeyTrn <> "&" & accesskeyTrn Then
			Temp = Replace(Temp,acckeyTrn,"")
		End If
		If acckeySrc <> "" And acckeySrc <> acckeyTrn Then
			posinTrn = InStr(Temp,accesskeySrc)
			If posinTrn = 0 Then posinTrn = InStr(UCase(Temp),UCase(accesskeySrc))
			If posinTrn <> 0 Then
				keyTrn = Mid(Temp,posinTrn,1)
				keySrc = "&" & keyTrn
				If acckeyTrn = "" Then Temp = Replace(Temp,keyTrn,keySrc,,1)
				If acckeyTrn <> "" And acckeyTrn <> keySrc Then
					Temp = Replace(Temp,acckeyTrn,accesskeyTrn)
					Temp = Replace(Temp,keyTrn,keySrc,,1)
				End If
			Else
				If acckeyTrn = "" Then acckeyTrn = acckeySrc
			End If
		End If
	End If
	StringReplace = Temp
End Function


' 修改消息输出
Function ReplaceMassage(OldtrnString As String,NewtrnString As String) As String
	Dim AcckeyMsg As String,EndStringMsg As String,ShortcutMsg As String,Tmsg1 As String
	Dim Tmsg2 As String,Fmsg As String,Smsg As String,Massage1 As String,Massage2 As String
	Dim Massage3 As String,Massage4 As String,n As Integer,x As Integer,y As Integer

	If OSLanguage = "0404" And iVo = 0 Then
		Msg00 = "獽倍龄籔ゅぃ"
		Msg01 = "沧ゎ才籔ゅぃ┪惠璶陆亩"
		Msg02 = "硉竟籔ゅぃ"
		Msg03 = "獽倍龄㎝沧ゎ才籔ゅぃ"
		Msg04 = "沧ゎ才㎝硉竟籔ゅぃ"
		Msg05 = "獽倍龄㎝硉竟籔ゅぃ"
		Msg06 = "獽倍龄沧ゎ才㎝硉竟籔ゅぃ"
		Msg07 = "獽倍龄糶ぃ"
		Msg08 = "獽倍龄糶㎝沧ゎ才籔ゅぃ"
		Msg09 = "獽倍龄糶㎝硉竟籔ゅぃ"
		Msg10 = "獽倍龄糶沧ゎ才㎝硉竟籔ゅぃ"

		Msg11 = "亩ゅいぶ獽倍龄"
		Msg12 = "亩ゅいぶ沧ゎ才"
		Msg13 = "亩ゅいぶ硉竟"
		Msg14 = "亩ゅいぶ獽倍龄㎝沧ゎ才"
		Msg15 = "亩ゅいぶ沧ゎ才㎝硉竟"
		Msg16 = "亩ゅいぶ獽倍龄㎝硉竟"
		Msg17 = "亩ゅいぶ獽倍龄沧ゎ才㎝硉竟"

		Msg21 = "度亩ゅいΤ獽倍龄"
		Msg22 = "度亩ゅいΤ沧ゎ才"
		Msg23 = "度亩ゅいΤ硉竟"
		Msg24 = "度亩ゅいΤ獽倍龄㎝沧ゎ才"
		Msg25 = "度亩ゅいΤ沧ゎ才㎝硉竟"
		Msg26 = "度亩ゅいΤ獽倍龄㎝硉竟"
		Msg27 = "度亩ゅいΤ獽倍龄沧ゎ才㎝硉竟"

		Msg31 = "惠簿笆獽倍龄程"
		Msg32 = "惠簿笆獽倍龄沧ゎ才玡"
		Msg33 = "惠簿笆獽倍龄硉竟玡"

		Msg41 = "獽倍龄玡Τ"
		Msg42 = "沧ゎ才玡Τ"
		Msg43 = "硉竟玡Τ"
		Msg44 = "獽倍龄玡Τ沧ゎ才"
		Msg45 = "獽倍龄玡Τ㎝沧ゎ才"

		Msg46 = "沧ゎ才玡ゑゅぶ"
		Msg47 = "沧ゎ才ゑゅぶ"
		Msg48 = "沧ゎ才玡ゑゅぶ"
		Msg49 = "﹃ゑゅぶ"
		Msg50 = "沧ゎ才玡㎝﹃ゑゅぶ"
		Msg51 = "沧ゎ才㎝﹃ゑゅぶ"
		Msg52 = "沧ゎ才玡㎝﹃ゑゅぶ"
		Msg53 = "沧ゎ才玡Τ緇"
		Msg54 = "沧ゎ才Τ緇"
		Msg55 = "沧ゎ才玡Τ緇"
		Msg56 = "﹃Τ緇"
		Msg57 = "沧ゎ才玡㎝﹃Τ緇"
		Msg58 = "沧ゎ才㎝﹃Τ緇"
		Msg59 = "沧ゎ才玡㎝﹃Τ緇"

		Msg61 = ""
		Msg62 = ""
		Msg63 = ""
		Msg64 = ""
		Msg65 = "盢 "
		Msg66 = " 蠢传 "
		Msg67 = "埃 "
	ElseIf OSLanguage = "0404" And iVo = 1 Then
		Msg00 = "э獽倍龄"
		Msg01 = "э沧ゎ才"
		Msg02 = "э硉竟"
		Msg03 = "э獽倍龄㎝沧ゎ才"
		Msg04 = "э沧ゎ才㎝硉竟"
		Msg05 = "э獽倍龄㎝硉竟"
		Msg06 = "э獽倍龄沧ゎ才㎝硉竟"
		Msg07 = "э獽倍龄糶"
		Msg08 = "э獽倍龄糶㎝沧ゎ才"
		Msg09 = "э獽倍龄糶㎝硉竟"
		Msg10 = "э獽倍龄糶沧ゎ才㎝硉竟"

		Msg11 = "穝糤獽倍龄"
		Msg12 = "穝糤沧ゎ才"
		Msg13 = "穝糤硉竟"
		Msg14 = "穝糤獽倍龄㎝沧ゎ才"
		Msg15 = "穝糤沧ゎ才㎝硉竟"
		Msg16 = "穝糤獽倍龄㎝硉竟"
		Msg17 = "穝糤獽倍龄沧ゎ才㎝硉竟"

		Msg21 = "度亩ゅいΤ獽倍龄"
		Msg22 = "度亩ゅいΤ沧ゎ才"
		Msg23 = "度亩ゅいΤ硉竟"
		Msg24 = "度亩ゅいΤ獽倍龄㎝沧ゎ才"
		Msg25 = "度亩ゅいΤ沧ゎ才㎝硉竟"
		Msg26 = "度亩ゅいΤ獽倍龄㎝硉竟"
		Msg27 = "度亩ゅいΤ獽倍龄沧ゎ才㎝硉竟"

		Msg31 = "簿笆獽倍龄程"
		Msg32 = "簿笆獽倍龄沧ゎ才玡"
		Msg33 = "簿笆獽倍龄硉竟玡"

		Msg41 = "埃獽倍龄玡"
		Msg42 = "埃沧ゎ才玡"
		Msg43 = "埃硉竟玡"
		Msg44 = "埃獽倍龄玡沧ゎ才"
		Msg45 = "埃獽倍龄玡㎝沧ゎ才"

		Msg46 = "穝糤沧ゎ才玡ぶ"
		Msg47 = "穝糤沧ゎ才ぶ"
		Msg48 = "穝糤沧ゎ才玡ぶ"
		Msg49 = "穝糤﹃ぶ"
		Msg50 = "穝糤沧ゎ才玡㎝﹃ぶ"
		Msg51 = "穝糤沧ゎ才㎝﹃ぶ"
		Msg52 = "穝糤沧ゎ才玡㎝﹃ぶ"
		Msg53 = "埃沧ゎ才玡緇"
		Msg54 = "埃沧ゎ才緇"
		Msg55 = "埃沧ゎ才玡緇"
		Msg56 = "埃﹃緇"
		Msg57 = "埃沧ゎ才玡㎝﹃緇"
		Msg58 = "埃沧ゎ才㎝﹃緇"
		Msg59 = "埃沧ゎ才玡㎝﹃緇"

		Msg61 = ""
		Msg62 = ""
		Msg63 = ""
		Msg64 = ""
		Msg65 = "盢 "
		Msg66 = " 蠢传 "
		Msg67 = "埃 "
	ElseIf OSLanguage <> "0404" And iVo = 0 Then
		Msg00 = "快捷键与原文不同"
		Msg01 = "终止符与原文不同或者需要翻译"
		Msg02 = "加速器与原文不同"
		Msg03 = "快捷键和终止符与原文不同"
		Msg04 = "终止符和加速器与原文不同"
		Msg05 = "快捷键和加速器与原文不同"
		Msg06 = "快捷键、终止符和加速器与原文不同"
		Msg07 = "快捷键的大小写不同"
		Msg08 = "快捷键的大小写和终止符与原文不同"
		Msg09 = "快捷键的大小写和加速器与原文不同"
		Msg10 = "快捷键的大小写、终止符和加速器与原文不同"

		Msg11 = "译文中缺少快捷键"
		Msg12 = "译文中缺少终止符"
		Msg13 = "译文中缺少加速器"
		Msg14 = "译文中缺少快捷键和终止符"
		Msg15 = "译文中缺少终止符和加速器"
		Msg16 = "译文中缺少快捷键和加速器"
		Msg17 = "译文中缺少快捷键、终止符和加速器"

		Msg21 = "仅译文中有快捷键"
		Msg22 = "仅译文中有终止符"
		Msg23 = "仅译文中有加速器"
		Msg24 = "仅译文中有快捷键和终止符"
		Msg25 = "仅译文中有终止符和加速器"
		Msg26 = "仅译文中有快捷键和加速器"
		Msg27 = "仅译文中有快捷键、终止符和加速器"

		Msg31 = "需移动快捷键到最后"
		Msg32 = "需移动快捷键到终止符前"
		Msg33 = "需移动快捷键到加速器前"

		Msg41 = "快捷键前有空格"
		Msg42 = "终止符前有空格"
		Msg43 = "加速器前有空格"
		Msg44 = "快捷键前有终止符"
		Msg45 = "快捷键前有空格和终止符"

		Msg46 = "终止符前的空格比原文少"
		Msg47 = "终止符后的空格比原文少"
		Msg48 = "终止符前后的空格比原文少"
		Msg49 = "字串后的空格比原文少"
		Msg50 = "终止符前和字串后的空格比原文少"
		Msg51 = "终止符后和字串后的空格比原文少"
		Msg52 = "终止符前后和字串后的空格比原文少"
		Msg53 = "终止符前有多余的空格"
		Msg54 = "终止符后有多余的空格"
		Msg55 = "终止符前后有多余的空格"
		Msg56 = "字串后有多余的空格"
		Msg57 = "终止符前和字串后有多余的空格"
		Msg58 = "终止符后和字串后有多余的空格"
		Msg59 = "终止符前后和字串后有多余的空格"

		Msg61 = "，"
		Msg62 = "并"
		Msg63 = "。"
		Msg64 = "、"
		Msg65 = "已将 "
		Msg66 = " 替换为 "
		Msg67 = "已删除 "
	ElseIf OSLanguage <> "0404" And iVo = 1 Then
		Msg00 = "已修改了快捷键"
		Msg01 = "已修改了终止符"
		Msg02 = "已修改了加速器"
		Msg03 = "已修改了快捷键和终止符"
		Msg04 = "已修改了终止符和加速器"
		Msg05 = "已修改了快捷键和加速器"
		Msg06 = "已修改了快捷键、终止符和加速器"
		Msg07 = "已修改了快捷键的大小写"
		Msg08 = "已修改了快捷键的大小写和终止符"
		Msg09 = "已修改了快捷键的大小写和加速器"
		Msg10 = "已修改了快捷键的大小写、终止符和加速器"

		Msg11 = "已添加了快捷键"
		Msg12 = "已添加了终止符"
		Msg13 = "已添加了加速器"
		Msg14 = "已添加了快捷键和终止符"
		Msg15 = "已添加了终止符和加速器"
		Msg16 = "已添加了快捷键和加速器"
		Msg17 = "已添加了快捷键、终止符和加速器"

		Msg21 = "仅译文中有快捷键"
		Msg22 = "仅译文中有终止符"
		Msg23 = "仅译文中有加速器"
		Msg24 = "仅译文中有快捷键和终止符"
		Msg25 = "仅译文中有终止符和加速器"
		Msg26 = "仅译文中有快捷键和加速器"
		Msg27 = "仅译文中有快捷键、终止符和加速器"

		Msg31 = "已移动快捷键到最后"
		Msg32 = "已移动快捷键到终止符前"
		Msg33 = "已移动快捷键到加速器前"

		Msg41 = "已去除了快捷键前的空格"
		Msg42 = "已去除了终止符前的空格"
		Msg43 = "已去除了加速器前的空格"
		Msg44 = "已去除了快捷键前的终止符"
		Msg45 = "已去除了快捷键前的空格和终止符"

		Msg46 = "已添加了终止符前缺少的空格"
		Msg47 = "已添加了终止符后缺少的空格"
		Msg48 = "已添加了终止符前后缺少的空格"
		Msg49 = "已添加了字串后缺少的空格"
		Msg50 = "已添加了终止符前和字串后缺少的空格"
		Msg51 = "已添加了终止符后和字串后缺少的空格"
		Msg52 = "已添加了终止符前后和字串后缺少的空格"
		Msg53 = "已去除了终止符前多余的空格"
		Msg54 = "已去除了终止符后多余的空格"
		Msg55 = "已去除了终止符前后多余的空格"
		Msg56 = "已去除了字串后多余的空格"
		Msg57 = "已去除了终止符前和字串后多余的空格"
		Msg58 = "已去除了终止符后和字串后多余的空格"
		Msg59 = "已去除了终止符前后和字串后多余的空格"

		Msg61 = "，"
		Msg62 = "并"
		Msg63 = "。"
		Msg64 = "、"
		Msg65 = "已将 "
		Msg66 = " 替换为 "
		Msg67 = "已删除 "
	End If

	If acckeySrc <> "" And acckeyTrn <> "" Then AcckeyMsg = "aM"
	If acckeySrc <> "" And acckeyTrn = "" Then AcckeyMsg = "aA"
	If acckeySrc = "" And acckeyTrn <> "" Then AcckeyMsg = "aN"
	If acckeySrc = acckeyTrn Then AcckeyMsg = "aG"
	If acckeySrc <> acckeyTrn And LCase(acckeySrc) = LCase(acckeyTrn) Then AcckeyMsg = "aC"

	If EndStringSrc <> "" And EndStringTrn <> "" Then EndStringMsg = "eM"
	If EndStringSrc <> "" And EndStringTrn = "" Then EndStringMsg = "eA"
	If EndStringSrc = "" And EndStringTrn <> "" Then EndStringMsg = "eN"
	If Trim(EndStringSrc) = Trim(EndStringTrn) Then EndStringMsg = "eG"

	If ShortcutSrc <> "" And ShortcutTrn <> "" Then ShortcutMsg = "sM"
	If ShortcutSrc <> "" And ShortcutTrn = "" Then ShortcutMsg = "sA"
	If ShortcutSrc = "" And ShortcutTrn <> "" Then ShortcutMsg = "sN"
	If ShortcutSrc = ShortcutTrn Then ShortcutMsg = "sG"

	Tmsg1 = AcckeyMsg & EndStringMsg & ShortcutMsg
	If Tmsg1 = "aMeGsG" Then Massage1 = Msg00 & Msg63
	If Tmsg1 = "aGeMsG" Then Massage1 = Msg01 & Msg63
	If Tmsg1 = "aGeGsM" Then Massage1 = Msg02 & Msg63
	If Tmsg1 = "aMeMsG" Then Massage1 = Msg03 & Msg63
	If Tmsg1 = "aGeMsM" Then Massage1 = Msg04 & Msg63
	If Tmsg1 = "aMeGsM" Then Massage1 = Msg05 & Msg63
	If Tmsg1 = "aMeMsM" Then Massage1 = Msg06 & Msg63
	If Tmsg1 = "aCeGsG" Then Massage1 = Msg07 & Msg63
	If Tmsg1 = "aCeMsG" Then Massage1 = Msg08 & Msg63
	If Tmsg1 = "aCeGsM" Then Massage1 = Msg09 & Msg63
	If Tmsg1 = "aCeMsM" Then Massage1 = Msg10 & Msg63

	If Tmsg1 = "aAeGsG" Then Massage1 = Msg11 & Msg63
	If Tmsg1 = "aGeAsG" Then Massage1 = Msg12 & Msg63
	If Tmsg1 = "aGeGsA" Then Massage1 = Msg13 & Msg63
	If Tmsg1 = "aAeAsG" Then Massage1 = Msg14 & Msg63
	If Tmsg1 = "aGeAsA" Then Massage1 = Msg15 & Msg63
	If Tmsg1 = "aAeGsA" Then Massage1 = Msg16 & Msg63
	If Tmsg1 = "aAeAsA" Then Massage1 = Msg17 & Msg63

	If Tmsg1 = "aNeGsG" Then Massage1 = Msg21 & Msg63
	If Tmsg1 = "aGeNsG" Then Massage1 = Msg22 & Msg63
	If Tmsg1 = "aGeGsN" Then Massage1 = Msg23 & Msg63
	If Tmsg1 = "aNeNsG" Then Massage1 = Msg24 & Msg63
	If Tmsg1 = "aGeNsN" Then Massage1 = Msg25 & Msg63
	If Tmsg1 = "aNeGsN" Then Massage1 = Msg26 & Msg63
	If Tmsg1 = "aNeNsN" Then Massage1 = Msg27 & Msg63

	If Tmsg1 = "aMeAsG" Then Massage1 = Msg00 & Msg61 & Msg62 & Msg12 & Msg63
	If Tmsg1 = "aMeGsA" Then Massage1 = Msg00 & Msg61 & Msg62 & Msg13 & Msg63
	If Tmsg1 = "aAeMsG" Then Massage1 = Msg01 & Msg61 & Msg62 & Msg11 & Msg63
	If Tmsg1 = "aGeMsA" Then Massage1 = Msg01 & Msg61 & Msg62 & Msg13 & Msg63
	If Tmsg1 = "aAeGsM" Then Massage1 = Msg02 & Msg61 & Msg62 & Msg11 & Msg63
	If Tmsg1 = "aGeAsM" Then Massage1 = Msg02 & Msg61 & Msg62 & Msg12 & Msg63
	If Tmsg1 = "aCeAsG" Then Massage1 = Msg07 & Msg61 & Msg62 & Msg12 & Msg63
	If Tmsg1 = "aCeGsA" Then Massage1 = Msg07 & Msg61 & Msg62 & Msg13 & Msg63

	If Tmsg1 = "aAeNsG" Then Massage1 = Msg11 & Msg61 & Msg22 & Msg63
	If Tmsg1 = "aAeGsN" Then Massage1 = Msg11 & Msg61 & Msg23 & Msg63
	If Tmsg1 = "aNeAsG" Then Massage1 = Msg12 & Msg61 & Msg21 & Msg63
	If Tmsg1 = "aGeAsN" Then Massage1 = Msg12 & Msg61 & Msg23 & Msg63
	If Tmsg1 = "aNeGsA" Then Massage1 = Msg13 & Msg61 & Msg21 & Msg63
	If Tmsg1 = "aGeNsA" Then Massage1 = Msg13 & Msg61 & Msg22 & Msg63

	If Tmsg1 = "aMeNsG" Then Massage1 = Msg00 & Msg61 & Msg22 & Msg63
	If Tmsg1 = "aMeGsN" Then Massage1 = Msg00 & Msg61 & Msg23 & Msg63
	If Tmsg1 = "aNeMsG" Then Massage1 = Msg01 & Msg61 & Msg21 & Msg63
	If Tmsg1 = "aGeMsN" Then Massage1 = Msg01 & Msg61 & Msg23 & Msg63
	If Tmsg1 = "aNeGsM" Then Massage1 = Msg02 & Msg61 & Msg21 & Msg63
	If Tmsg1 = "aGeNsM" Then Massage1 = Msg02 & Msg61 & Msg22 & Msg63
	If Tmsg1 = "aCeNsG" Then Massage1 = Msg07 & Msg61 & Msg22 & Msg63
	If Tmsg1 = "aCeGsN" Then Massage1 = Msg07 & Msg61 & Msg23 & Msg63

	If Tmsg1 = "aMeAsN" Then Massage1 = Msg00 & Msg61 & Msg62 & Msg12 & Msg61 & Msg23 & Msg63
	If Tmsg1 = "aMeNsA" Then Massage1 = Msg00 & Msg61 & Msg62 & Msg13 & Msg61 & Msg22 & Msg63
	If Tmsg1 = "aAeMsN" Then Massage1 = Msg01 & Msg61 & Msg62 & Msg11 & Msg61 & Msg23 & Msg63
	If Tmsg1 = "aNeMsA" Then Massage1 = Msg01 & Msg61 & Msg62 & Msg13 & Msg61 & Msg21 & Msg63
	If Tmsg1 = "aAeNsM" Then Massage1 = Msg02 & Msg61 & Msg62 & Msg11 & Msg61 & Msg22 & Msg63
	If Tmsg1 = "aNeAsM" Then Massage1 = Msg02 & Msg61 & Msg62 & Msg12 & Msg61 & Msg21 & Msg63
	If Tmsg1 = "aCeAsN" Then Massage1 = Msg07 & Msg61 & Msg62 & Msg12 & Msg61 & Msg23 & Msg63
	If Tmsg1 = "aCeNsA" Then Massage1 = Msg07 & Msg61 & Msg62 & Msg13 & Msg61 & Msg22 & Msg63

	If Tmsg1 = "aMeAsA" Then Massage1 = Msg00 & Msg61 & Msg62 & Msg15 & Msg63
	If Tmsg1 = "aAeMsA" Then Massage1 = Msg01 & Msg61 & Msg62 & Msg16 & Msg63
	If Tmsg1 = "aAeAsM" Then Massage1 = Msg02 & Msg61 & Msg62 & Msg14 & Msg63
	If Tmsg1 = "aCeAsA" Then Massage1 = Msg07 & Msg61 & Msg62 & Msg15 & Msg63

	If Tmsg1 = "aMeNsN" Then Massage1 = Msg00 & Msg61 & Msg25 & Msg63
	If Tmsg1 = "aNeMsN" Then Massage1 = Msg01 & Msg61 & Msg26 & Msg63
	If Tmsg1 = "aNeNsM" Then Massage1 = Msg02 & Msg61 & Msg24 & Msg63
	If Tmsg1 = "aCeNsN" Then Massage1 = Msg07 & Msg61 & Msg25 & Msg63

	If Tmsg1 = "aAeNsN" Then Massage1 = Msg11 & Msg61 & Msg25 & Msg63
	If Tmsg1 = "aNeAsN" Then Massage1 = Msg12 & Msg61 & Msg26 & Msg63
	If Tmsg1 = "aNeNsA" Then Massage1 = Msg13 & Msg61 & Msg24 & Msg63

	If Tmsg1 = "aMeMsA" Then Massage1 = Msg03 & Msg61 & Msg62 & Msg13 & Msg63
	If Tmsg1 = "aAeMsM" Then Massage1 = Msg04 & Msg61 & Msg62 & Msg11 & Msg63
	If Tmsg1 = "aMeAsM" Then Massage1 = Msg05 & Msg61 & Msg62 & Msg12 & Msg63
	If Tmsg1 = "aCeMsA" Then Massage1 = Msg08 & Msg61 & Msg62 & Msg13 & Msg63
	If Tmsg1 = "aCeAsM" Then Massage1 = Msg09 & Msg61 & Msg62 & Msg12 & Msg63

	If Tmsg1 = "aMeMsN" Then Massage1 = Msg03 & Msg61 & Msg23 & Msg63
	If Tmsg1 = "aNeMsM" Then Massage1 = Msg04 & Msg61 & Msg21 & Msg63
	If Tmsg1 = "aMeNsM" Then Massage1 = Msg05 & Msg61 & Msg22 & Msg63
	If Tmsg1 = "aCeMsN" Then Massage1 = Msg08 & Msg61 & Msg23 & Msg63
	If Tmsg1 = "aCeNsM" Then Massage1 = Msg09 & Msg61 & Msg22 & Msg63

	If Tmsg1 = "aAeAsN" Then Massage1 = Msg14 & Msg61 & Msg23 & Msg63
	If Tmsg1 = "aNeAsA" Then Massage1 = Msg15 & Msg61 & Msg21 & Msg63
	If Tmsg1 = "aAeNsA" Then Massage1 = Msg16 & Msg61 & Msg22 & Msg63

	If PreStringTrn <> "" And SpaceTrn <> "" And ExpStringTrn <> "" Then Fmsg = "PSE"
	If PreStringTrn <> "" And SpaceTrn = "" And ExpStringTrn <> "" Then Fmsg = "PNE"
	If PreStringTrn <> "" And SpaceTrn <> "" And ExpStringTrn = "" Then Fmsg = "PSN"
	If PreStringTrn <> "" And SpaceTrn = "" And ExpStringTrn = "" Then Fmsg = "PNN"
	If PreStringTrn = "" And SpaceTrn <> "" And ExpStringTrn <> "" Then Fmsg = "NSE"
	If PreStringTrn = "" And SpaceTrn = ""  And ExpStringTrn <> "" Then Fmsg = "NNE"
	If PreStringTrn = "" And SpaceTrn <> "" And ExpStringTrn = "" Then Fmsg = "NSN"
	If PreStringTrn = "" And SpaceTrn = "" And ExpStringTrn = "" Then Fmsg = "NNN"

	If acckeyTrn <> "" And EndStringTrn <> "" And ShortcutTrn <> "" Then Smsg = "AES"
	If acckeyTrn <> "" And EndStringTrn <> "" And ShortcutTrn = "" Then Smsg = "AEN"
	If acckeyTrn <> "" And EndStringTrn = "" And ShortcutTrn <> "" Then Smsg = "ANS"
	If acckeyTrn <> "" And EndStringTrn = "" And ShortcutTrn = "" Then Smsg = "ANN"
	If acckeyTrn = "" And EndStringTrn <> "" And ShortcutTrn <> "" Then Smsg = "NES"
	If acckeyTrn = "" And EndStringTrn <> "" And ShortcutTrn = "" Then Smsg = "NEN"
	If acckeyTrn = "" And EndStringTrn = "" And ShortcutTrn <> "" Then Smsg = "NNS"
	If acckeyTrn = "" And EndStringTrn = "" And ShortcutTrn = "" Then Smsg = "NNN"

	Tmsg2 = Fmsg & Smsg
	If Tmsg2 = "NNEANN" Then Massage2 = Msg31 & Msg63
	If Tmsg2 = "NNEAEN" Then Massage2 = Msg32 & Msg63
	If Tmsg2 = "NNEAES" Then Massage2 = Msg32 & Msg63
	If Tmsg2 = "NNEANS" Then Massage2 = Msg33 & Msg63

	If Tmsg2 = "NSNANN" Then Massage2 = Msg41 & Msg63
	If Tmsg2 = "NSNAEN" Then Massage2 = Msg41 & Msg63
	If Tmsg2 = "NSNAES" Then Massage2 = Msg41 & Msg63
	If Tmsg2 = "NSNANS" Then Massage2 = Msg41 & Msg63
	If Tmsg2 = "NSNNEN" Then Massage2 = Msg42 & Msg63
	If Tmsg2 = "NSNNES" Then Massage2 = Msg42 & Msg63
	If Tmsg2 = "NSNNNS" Then Massage2 = Msg43 & Msg63

	If Tmsg2 = "NSEANN" Then Massage2 = Msg41 & Msg61 & Msg62 & Msg31 & Msg63
	If Tmsg2 = "NSEAEN" Then Massage2 = Msg41 & Msg61 & Msg62 & Msg32 & Msg63
	If Tmsg2 = "NSEAES" Then Massage2 = Msg41 & Msg61 & Msg62 & Msg32 & Msg63
	If Tmsg2 = "NSEANS" Then Massage2 = Msg41 & Msg61 & Msg62 & Msg33 & Msg63

	If Tmsg2 = "PNEANN" Then Massage2 = Msg44 & Msg61 & Msg62 & Msg31 & Msg63
	If Tmsg2 = "PNEAEN" Then Massage2 = Msg44 & Msg61 & Msg62 & Msg32 & Msg63
	If Tmsg2 = "PNEAES" Then Massage2 = Msg44 & Msg61 & Msg62 & Msg32 & Msg63
	If Tmsg2 = "PNEANS" Then Massage2 = Msg44 & Msg61 & Msg62 & Msg33 & Msg63

	If Tmsg2 = "PNNANN" Then Massage2 = Msg44 & Msg61 & Msg62 & Msg31 & Msg63
	If Tmsg2 = "PNNAEN" Then Massage2 = Msg44 & Msg61 & Msg62 & Msg32 & Msg63
	If Tmsg2 = "PNNAES" Then Massage2 = Msg44 & Msg61 & Msg62 & Msg32 & Msg63
	If Tmsg2 = "PNNANS" Then Massage2 = Msg44 & Msg61 & Msg62 & Msg33 & Msg63

	If Tmsg2 = "PSEANN" Then Massage2 = Msg45 & Msg61 & Msg62 & Msg31 & Msg63
	If Tmsg2 = "PSEAEN" Then Massage2 = Msg45 & Msg61 & Msg62 & Msg32 & Msg63
	If Tmsg2 = "PSEAES" Then Massage2 = Msg45 & Msg61 & Msg62 & Msg32 & Msg63
	If Tmsg2 = "PSEANS" Then Massage2 = Msg45 & Msg61 & Msg62 & Msg33 & Msg63

	If Tmsg2 = "PSNANN" Then Massage2 = Msg45 & Msg61 & Msg62 & Msg31
	If Tmsg2 = "PSNAEN" Then Massage2 = Msg45 & Msg61 & Msg62 & Msg32 & Msg63
	If Tmsg2 = "PSNAES" Then Massage2 = Msg45 & Msg61 & Msg62 & Msg32 & Msg63
	If Tmsg2 = "PSNANS" Then Massage2 = Msg45 & Msg61 & Msg62 & Msg33 & Msg63

	FFmsg = "Fn"
	FEmsg = "En"
	If EndStringSrc <> EndStringTrn Then
		If EndStringSrc <> "" Or iVo = 0 Then
			x = Len(EndStringSrc) - Len(LTrim(EndStringSrc))
			y = Len(EndStringTrn) - Len(LTrim(EndStringTrn))
			If x > y Then FFmsg = "Fd"
			If x < y Then FFmsg = "Fx"
			If x = y Then FFmsg = "Fn"
		End If
		x = Len(EndStringSrc) - Len(RTrim(EndStringSrc))
		y = Len(EndStringTrn) - Len(RTrim(EndStringTrn))
		If x > y Then FEmsg = "Ed"
		If x < y Then FEmsg = "Ex"
		If x = y Then FEmsg = "En"
	End If

	SEmsg = "Sn"
	If EndSpaceSrc <> EndSpaceTrn Then
		If Len(EndSpaceSrc) > Len(EndSpaceTrn) Then SEmsg = "Sd"
		If Len(EndSpaceSrc) < Len(EndSpaceTrn) Then SEmsg = "Sx"
	End If

	Tmsg3 = FFmsg & FEmsg & SEmsg
	If Tmsg3 = "FdEdSd" Then Massage3 = Msg52 & Msg63
	If Tmsg3 = "FdEdSx" Then Massage3 = Msg48 & Msg61 & Msg62 & Msg56 & Msg63
	If Tmsg3 = "FdEdSn" Then Massage3 = Msg48 & Msg63

	If Tmsg3 = "FdExSd" Then Massage3 = Msg50 & Msg61 & Msg62 & Msg54 & Msg63
	If Tmsg3 = "FdExSx" Then Massage3 = Msg46 & Msg61 & Msg62 & Msg58 & Msg63
	If Tmsg3 = "FdExSn" Then Massage3 = Msg46 & Msg61 & Msg62 & Msg54 & Msg63

	If Tmsg3 = "FdEnSd" Then Massage3 = Msg50 & Msg63
	If Tmsg3 = "FdEnSx" Then Massage3 = Msg46 & Msg61 & Msg62 & Msg56 & Msg63
	If Tmsg3 = "FdEnSn" Then Massage3 = Msg46 & Msg63

	If Tmsg3 = "FxEdSd" Then Massage3 = Msg53 & Msg61 & Msg62 & Msg51 & Msg63
	If Tmsg3 = "FxEdSx" Then Massage3 = Msg57 & Msg61 & Msg62 & Msg47 & Msg63
	If Tmsg3 = "FxEdSn" Then Massage3 = Msg53 & Msg61 & Msg62 & Msg47 & Msg63

	If Tmsg3 = "FxExSd" Then Massage3 = Msg55 & Msg61 & Msg62 & Msg49 & Msg63
	If Tmsg3 = "FxExSx" Then Massage3 = Msg59 & Msg63
	If Tmsg3 = "FxExSn" Then Massage3 = Msg55 & Msg63

	If Tmsg3 = "FxEnSd" Then Massage3 = Msg53 & Msg61 & Msg62 & Msg49 & Msg63
	If Tmsg3 = "FxEnSx" Then Massage3 = Msg57 & Msg63
	If Tmsg3 = "FxEnSn" Then Massage3 = Msg53 & Msg63

	If Tmsg3 = "FnEdSd" Then Massage3 = Msg51 & Msg63
	If Tmsg3 = "FnEdSx" Then Massage3 = Msg47 & Msg61 & Msg62 & Msg56 & Msg63
	If Tmsg3 = "FnEdSn" Then Massage3 = Msg47 & Msg63

	If Tmsg3 = "FnExSd" Then Massage3 = Msg54 & Msg61 & Msg62 & Msg49 & Msg63
	If Tmsg3 = "FnExSx" Then Massage3 = Msg58 & Msg63
	If Tmsg3 = "FnExSn" Then Massage3 = Msg54 & Msg63

	If Tmsg3 = "FnEnSd" Then Massage3 = Msg49 & Msg63
	If Tmsg3 = "FnEnSx" Then Massage3 = Msg56 & Msg63

	If PreRepStr <> "" And AppRepStr <> "" Then
		If InStr(PreRepStr,JoinStr) Then PreRepStr = Replace(PreRepStr,JoinStr,Msg64)
		If InStr(AppRepStr,JoinStr) Then AppRepStr = Replace(AppRepStr,JoinStr,Msg64)
		Massage4 = Msg65 & PreRepStr & Msg66 & AppRepStr & Msg63
	ElseIf PreRepStr <> "" And AppRepStr = "" Then
		If InStr(PreRepStr,JoinStr) Then PreRepStr = Replace(PreRepStr,JoinStr,Msg64)
		Massage4 = Msg67 & PreRepStr & Msg63
	End If

	If Massage1 <> "" And InStr(Tmsg1,"M") Then n = n + 1
	If Massage1 <> "" And InStr(Tmsg1,"A") Then AddedCount = AddedCount + 1
	If Massage1 <> "" And InStr(Tmsg1,"N") Then WarningCount = WarningCount + 1
	If Massage2 <> "" Or Massage3 <> "" Or Massage4 <> "" Then n = n + 1
	If n > 0 Then ModifiedCount = ModifiedCount + 1
	ReplaceMassage = Massage1 & Massage2 & Massage3 & Massage4
End Function


' 删除快捷键消息输出
Function DelAccKeyMassage(OldtrnString As String,NewtrnString As String) As String
	If OSLanguage = "0404" Then
		Msg01 = "埃㎝ゅぃ獽倍龄 (%s)"
		Msg02 = "埃㎝ゅ獽倍龄 (%s)"
		Msg03 = "埃度亩ゅい獽倍龄 (%s)"
	Else
		Msg01 = "已删除和原文不同的快捷键 (%s)"
		Msg02 = "已删除和原文相同的快捷键 (%s)"
		Msg03 = "已删除仅在译文中存在的快捷键 (%s)"
	End If
	If NewtrnString <> OldtrnString And acckeyTrn <> "" Then
		If acckeyTrn <> "" And acckeySrc <> "" And acckeyTrn <> acckeySrc Then
			DelAccKeyMassage = Replace(Msg01,"%s",acckeyTrn)
			ModifiedCount = ModifiedCount + 1
		ElseIf acckeyTrn <> "" And acckeySrc <> "" And acckeyTrn = acckeySrc Then
			DelAccKeyMassage = Replace(Msg02,"%s",acckeyTrn)
			AddedCount = AddedCount + 1
		ElseIf acckeyTrn <> "" And acckeySrc = "" And acckeyTrn <> acckeySrc Then
			DelAccKeyMassage = Replace(Msg03,"%s",acckeyTrn)
			WarningCount = WarningCount + 1
		End If
	End If
End Function


'输出行数错误消息
Function LineErrMassage(srcLineNum As Integer,trnLineNum As Integer,LineNumErrCount As Integer) As String
	If OSLanguage = "0404" Then
		Msg01 = "亩ゅ︽计ゑゅぶ %s ︽"
		Msg02 = "亩ゅ︽计ゑゅ %s ︽"
	Else
		Msg01 = "译文的行数比原文少 %s 行。"
		Msg02 = "译文的行数比原文多 %s 行。"
	End If
	Dim LineNumErr As Integer
	LineNumErr = 0
	If srcLineNum <> trnLineNum Then
		If srcLineNum > trnLineNum Then
			LineNumErr = srcLineNum - trnLineNum
			LineErrMassage = Replace(Msg01,"%s",CStr(LineNumErr))
			LineNumErrCount = LineNumErrCount + 1
		ElseIf srcLineNum < trnLineNum Then
			LineNumErr = trnLineNum - srcLineNum
			LineErrMassage = Replace(Msg02,"%s",CStr(LineNumErr))
			LineNumErrCount = LineNumErrCount + 1
		End If
	End If
End Function


'输出快捷键数错误消息
Function AccKeyErrMassage(srcAccKeyNum As Integer,trnAccKeyNum As Integer,accKeyNumErrCount As Integer) As String
	If OSLanguage = "0404" Then
		Msg01 = "亩ゅ獽倍龄计ゑゅぶ %s "
		Msg02 = "亩ゅ獽倍龄计ゑゅ %s "
	Else
		Msg01 = "译文的快捷键数比原文少 %s 个。"
		Msg02 = "译文的快捷键数比原文多 %s 个。"
	End If
	Dim accKeyNumErr As Integer
	accKeyNumErr = 0
	If srcAccKeyNum <> trnAccKeyNum Then
		If srcAccKeyNum > trnAccKeyNum Then
			accKeyNumErr = srcAccKeyNum - trnAccKeyNum
			AccKeyErrMassage = Replace(Msg01,"%s",CStr(accKeyNumErr))
			accKeyNumErrCount = accKeyNumErrCount + 1
		ElseIf srcAccKeyNum < trnAccKeyNum Then
			accKeyNumErr = trnAccKeyNum - srcAccKeyNum
			AccKeyErrMassage = Replace(Msg02,"%s",CStr(accKeyNumErr))
			accKeyNumErrCount = accKeyNumErrCount + 1
		End If
	End If
End Function


'错误计数消息输出
Function CountMassage(ErrorCount As Integer,LineNumErrCount As Integer,accKeyNumErrCount As Integer) As String
	If OSLanguage = "0404" Then
		Msg01 = "⊿Τт岿粇"
		Msg02 = "т " & ErrorCount & " 岿粇ㄤい: "
		Msg03 = "㎝ゅぃ " & ModifiedCount & " " & _
				"亩ゅいぶ " & AddedCount & " " & _
				"度亩ゅい " & WarningCount & " "
		Msg04 = "ゅ㎝亩ゅい︽计ぃ " & LineNumErrCount & " " & _
				"ゅ㎝亩ゅい獽倍龄计ぃ " & accKeyNumErrCount & " "
		Msg06 = "э " & ModifiedCount & " 穝糤 " & AddedCount & " " & _
				"璶浪琩 " & WarningCount & " ︽计ぃ " & LineNumErrCount & " " & _
				"獽倍龄计ぃ " & accKeyNumErrCount & " "
		Msg07 = "⊿Τ埃ヴ獽倍龄"
		Msg08 = "埃 " & ErrorCount & " 獽倍龄ㄤい: "
		Msg09 = "㎝ゅぃ " & ModifiedCount & " " & _
				"㎝ゅ " & AddedCount & " " & _
				"度亩ゅい " & WarningCount & " "
	Else
		Msg01 = "没有找到错误。"
		Msg02 = "找到 " & ErrorCount & " 个错误。其中: "
		Msg03 = "和原文不同 " & ModifiedCount & " 个，" & _
				"译文中缺少 " & AddedCount & " 个，" & _
				"仅在译文中存在 " & WarningCount & " 个，"
		Msg04 = "原文和译文中行数不同 " & LineNumErrCount & " 个，" & _
				"原文和译文中快捷键数不同 " & accKeyNumErrCount & " 个。"
		Msg06 = "已修改 " & ModifiedCount & " 个，已添加 " & AddedCount & " 个，" & _
				"要检查 " & WarningCount & " 个，行数不同 " & LineNumErrCount & " 个，" & _
				"快捷键数不同 " & accKeyNumErrCount & " 个。"
		Msg07 = "没有删除任何快捷键。"
		Msg08 = "已删除 " & ErrorCount & " 个快捷键。其中: "
		Msg09 = "和原文不同 " & ModifiedCount & " 个，" & _
				"和原文相同 " & AddedCount & " 个，" & _
				"仅在译文中存在 " & WarningCount & " 个。"
	End If

	If iVo = 0 And ErrorCount = 0 Then CountMassage = Msg01
	If iVo = 0 And ErrorCount <> 0 Then
		CountMassage = Msg02 & vbCrLf & Msg03 & vbCrLf & Msg04
	End If

	If iVo = 1 And ErrorCount = 0 Then CountMassage = Msg01
	If iVo = 1 And ErrorCount <> 0 Then
		CountMassage = Msg02 & vbCrLf & Msg06
	End If

	If iVo = 2 And ErrorCount = 0 Then CountMassage = Msg07
	If iVo = 2 And ErrorCount <> 0 Then
		CountMassage = Msg08 & vbCrLf & Msg09
	End If
End Function


'输出程序错误消息
Sub sysErrorMassage(sysError As ErrObject)
	If OSLanguage = "0404" Then
		Msg01 = "岿粇"
		Msg02 = "祇ネ祘Α砞璸岿粇岿粇絏 "
	Else
		Msg01 = "错误"
		Msg02 = "发生程序设计上的错误。错误代码 "
	End If
	MsgBox(msg02 & CStr(sysError.Number) & ", " & sysError.Description,vbOkOnly+vbInformation,Msg01)
End Sub


'在快捷键后插入特定字符并以此拆分字串
Function Insert(SplitString As String,SplitStr As String,InsStr As String,Leng As Integer) As String
	Dim bakString As String,StartNum As Integer,EndNum As Integer
	Dim i As Integer,oldLeng As Integer,newLeng As Integer,accesskeyStr As String
	If UBound(Split(SplitString,SplitStr)) < 2 Then
		Insert = SplitString
		Exit Function
	End If
	bakString = SplitString
	StartNum = InStr(bakString,SplitStr)
	EndNum = InStrRev(bakString,SplitStr)
	If StartNum < EndNum Then
		For i = StartNum To EndNum
			If Mid(bakString,i,Len(SplitStr)) = SplitStr Then
				oldLeng = InStr(Mid(bakString,i+Len(SplitStr)),SplitStr)
				newLeng = InStr(Mid(bakString,i+Len(SplitStr)),InsStr)
				accesskeyStr = ""
				If oldLeng <> 0 And newLeng <> 0 And oldLeng > newLeng Then
					accesskeyStr = ""
				ElseIf oldLeng <> 0 And newLeng <> 0 And oldLeng < newLeng And i > Leng Then
					accesskeyStr = Mid(bakString,i-Leng,oldLeng)
				ElseIf oldLeng = 0 And newLeng <> 0 And i > Leng Then
					accesskeyStr = Mid(bakString,i-Leng,newLeng)
				ElseIf oldLeng <> 0 And newLeng = 0 And i > Leng  Then
					accesskeyStr = Mid(bakString,i-Leng,oldLeng)
				ElseIf oldLeng = 0 And newLeng = 0 And i > Leng  Then
					accesskeyStr = Mid(bakString,i-Leng)
				End If
				If accesskeyStr <> "" Then
					bakString = Replace(bakString,accesskeyStr,accesskeyStr & InsStr)
				End If
				i = i + oldLeng
			End If
		Next i
	End If
	Insert = bakString
	'PSL.Output "bakString = " & bakString       '调试用
End Function


'进行数组合并
Function MergeArray(srcStrArr() As String,trnStrArr() As String) As Variant
	Dim i As Integer,srcNum As Integer,trnNum As Integer
	Dim srcPassNum As Integer,trnPassNum As Integer
	srcNum = UBound(srcStrArr)
	trnNum = UBound(trnStrArr)
	srcPassNum = 0
	trnPassNum = 0
	Dim TempArray() As String
	For i = 0 To (srcNum + trnNum + 1) Step 2
		ReDim Preserve TempArray(i)
		If srcNum >= srcPassNum Then
			TempArray(i) = srcStrArr(srcPassNum)
			srcPassNum = srcPassNum + 1
		ElseIf srcNum < srcPassNum Then
			TempArray(i) = ""
		End If
		ReDim Preserve TempArray(i+1)
		If trnNum >= trnPassNum Then
			TempArray(i+1) = trnStrArr(trnPassNum)
			trnPassNum = trnPassNum + 1
		ElseIf trnNum < trnPassNum Then
			TempArray(i+1) = ""
		End If
	Next i
	MergeArray = TempArray
End Function


'字串常数正向转换
Function Convert(ConverString As String) As String
	Convert = ConverString
	If Convert = "" Then Exit Function
	If InStr(Convert,"\") = 0 Then Exit Function
	If InStr(Convert,"\r\n") Then Convert = Replace(Convert,"\r\n",vbCrLf)
	If InStr(Convert,"\r\n") Then Convert = Replace(Convert,"\r\n",vbNewLine)
	If InStr(Convert,"\r") Then Convert = Replace(Convert,"\r",vbCr)
	If InStr(Convert,"\r") Then Convert = Replace(Convert,"\r",vbNewLine)
	If InStr(Convert,"\n") Then Convert = Replace(Convert,"\n",vbLf)
	If InStr(Convert,"\b") Then Convert = Replace(Convert,"\b",vbBack)
	If InStr(Convert,"\f") Then Convert = Replace(Convert,"\f",vbFormFeed)
	If InStr(Convert,"\v") Then Convert = Replace(Convert,"\v",vbVerticalTab)
	If InStr(Convert,"\t") Then Convert = Replace(Convert,"\t",vbTab)
	If InStr(Convert,"\'") Then Convert = Replace(Convert,"\'","'")
	If InStr(Convert,"\""") Then Convert = Replace(Convert,"\""","""")
	If InStr(Convert,"\?") Then Convert = Replace(Convert,"\?","?")
	If InStr(Convert,"\") Then Convert = ConvertB(Convert)
	If InStr(Convert,"\0") Then Convert = Replace(Convert,"\0",vbNullChar)
	If InStr(Convert,"\\") Then Convert = Replace(Convert,"\\","\")
End Function


'转换八进制或十六进制转义符
Function ConvertB(ConverString As String) As String
	Dim i As Integer,EscStr As String,ConvCode As String
	Dim ConvString As String,Stemp As Boolean
	ConvertB = ConverString
	If ConvertB = "" Then Exit Function
	i = InStr(ConvertB,"\")
	Do While i <> 0
		EscStr = Mid(ConvertB,i,2)
		Stemp = False

		If EscStr = "\x" Then
			ConvCode = Mid(ConvertB,i+2,2)
			Stemp = CheckStr(UCase(ConvCode),"48-57,65-70")
		ElseIf EscStr = "\u" Then
			ConvCode = Mid(ConvertB,i+2,4)
			Stemp = CheckStr(UCase(ConvCode),"48-57,65-70")
		ElseIf EscStr = "\U" Then
			ConvCode = Mid(ConvertB,i+2,8)
			Stemp = CheckStr(UCase(ConvCode),"48-57,65-70")
		ElseIf EscStr <> "" Then
			EscStr = "\"
			ConvCode = Mid(ConvertB,i+1,3)
			Stemp = CheckStr(ConvCode,"48-55")
		End If

		If Stemp = True Then
			If EscStr = "\x" Then ConvString = ChrW(HEXtoDEC(ConvCode))
			If LCase(EscStr) = "\u" Then ConvString = ChrW(HEXtoDEC(ConvCode))
			If EscStr = "\" Then ConvString = ChrW(OCTtoDEC(ConvCode))
			If ConvString <> "" Then
				ConvertB = Replace(ConvertB,EscStr & ConvCode,ConvString)
				i = 0
			End If
		End If

		i = InStr(i+1,ConvertB,"\")
		If i = 0 Then Exit Do
	Loop
End Function


' 用途: 将八进制转化为十进制
' 输入: OctStr(八进制数)
' 输入数据类型：String
' 输出: OCTtoDEC(十进制数)
' 输出数据类型: Long
' 输入最大数为17777777777,输出最大数为2147483647
Public Function OCTtoDEC(ByVal OctStr As String) As Long
	Dim i As Long,B As Long
	For i = 1 To Len(OctStr)
		Select Case Mid(OctStr, Len(OctStr) - i + 1, 1)
			Case "0"
				B = B + 8 ^ (i - 1) * 0
			Case "1"
				B = B + 8 ^ (i - 1) * 1
			Case "2"
				B = B + 8 ^ (i - 1) * 2
			Case "3"
				B = B + 8 ^ (i - 1) * 3
			Case "4"
				B = B + 8 ^ (i - 1) * 4
			Case "5"
				B = B + 8 ^ (i - 1) * 5
			Case "6"
				B = B + 8 ^ (i - 1) * 6
			Case "7"
				B = B + 8 ^ (i - 1) * 7
		End Select
	Next i
	OCTtoDEC = B
End Function


' 用途: 将十六进制转化为十进制
' 输入: HexStr(十六进制数)
' 输入数据类型: String
' 输出: HEXtoDEC(十进制数)
' 输出数据类型: Long
' 输入最大数为7FFFFFFF,输出最大数为2147483647
Public Function HEXtoDEC(ByVal HexStr As String) As Long
	Dim i As Long,B As Long
	HexStr = UCase(HexStr)
	For i = 1 To Len(HexStr)
		Select Case Mid(HexStr, Len(HexStr) - i + 1, 1)
			Case "0"
				B = B + 16 ^ (i - 1) * 0
			Case "1"
				B = B + 16 ^ (i - 1) * 1
			Case "2"
				B = B + 16 ^ (i - 1) * 2
			Case "3"
				B = B + 16 ^ (i - 1) * 3
			Case "4"
				B = B + 16 ^ (i - 1) * 4
			Case "5"
				B = B + 16 ^ (i - 1) * 5
			Case "6"
				B = B + 16 ^ (i - 1) * 6
			Case "7"
				B = B + 16 ^ (i - 1) * 7
			Case "8"
				B = B + 16 ^ (i - 1) * 8
			Case "9"
				B = B + 16 ^ (i - 1) * 9
			Case "A"
				B = B + 16 ^ (i - 1) * 10
			Case "B"
				B = B + 16 ^ (i - 1) * 11
			Case "C"
				B = B + 16 ^ (i - 1) * 12
			Case "D"
				B = B + 16 ^ (i - 1) * 13
			Case "E"
				B = B + 16 ^ (i - 1) * 14
			Case "F"
				B = B + 16 ^ (i - 1) * 15
        End Select
    Next i
    HEXtoDEC = B
End Function


'转换字符为整数数值
Function StrToInteger(mStr As String) As Integer
	If mStr = "" Then mStr = "0"
	StrToInteger = CInt(mStr)
End Function


'读取数组中的每个字串并替换处理
Function ReplaceStrSplit(CheckID As Integer,trnStr As String,StrSplitArr() As String,TranLang As String) As String
	Dim srcStrSplit As String,trnStrSplit As String,trnStrSplitNew As String,TacckeySrc As String
	Dim TEndStringSrc As String,TShortcutSrc As String,TSpaceTrn As String,TacckeyTrn As String
	Dim TExpStringTrn As String,TEndStringTrn As String,TShortcutTrn As String,TStringSrc As String
	Dim TStringTrn As String,TPreStringTrn As String,TEndSpaceSrc As String,TEndSpaceTrn As String
	TacckeySrc = ""
	TEndStringSrc = ""
	TShortcutSrc = ""
	TSpaceTrn = ""
	TacckeyTrn = ""
	TExpStringTrn = ""
	TEndStringTrn = ""
	TShortcutTrn = ""
	TPreStringTrn = ""
	TEndSpaceSrc = ""
	TEndSpaceTrn = ""
	ReplaceStrSplit = trnStr
	For i = 0 To UBound(StrSplitArr) Step 2
		srcStrSplit = StrSplitArr(i)
		trnStrSplit = StrSplitArr(i+1)
		trnStrSplitNew = StringReplace(CheckID,srcStrSplit,trnStrSplit,TranLang)

		'处理在前后行中包含的重复字符
		Dim y As Integer,z As Integer
		For y = 0 To i Step 2
			z = InStr(StrSplitArr(y+1),trnStrSplit)
			If z <> 0 And y < i Then
				ReplaceStrSplit = Replace(ReplaceStrSplit,StrSplitArr(y+1),"*d!N!f*")
				ReplaceStrSplit = Replace(ReplaceStrSplit,trnStrSplit,trnStrSplitNew)
				ReplaceStrSplit = Replace(ReplaceStrSplit,"*d!N!f*",StrSplitArr(y+1))
				Exit For
			ElseIf z <> 0 And y = i Then
				ReplaceStrSplit = Replace(ReplaceStrSplit,trnStrSplit,trnStrSplitNew)
				Exit For
			End If
		Next y

		'对每行的数据进行连接，用于消息输出
		TacckeySrc = TacckeySrc & acckeySrc
		TEndStringSrc = TEndStringSrc & EndStringSrc
		TShortcutSrc = TShortcutSrc & ShortcutSrc
		TSpaceTrn = TSpaceTrn & SpaceTrn
		TacckeyTrn = TacckeyTrn & acckeyTrn
		TExpStringTrn = TExpStringTrn & ExpStringTrn
		TEndStringTrn = TEndStringTrn & EndStringTrn
		TShortcutTrn = TShortcutTrn & ShortcutTrn
		TPreStringTrn = TPreStringTrn & PreStringTrn
		TEndSpaceSrc = TEndSpaceSrc & EndSpaceSrc
		TEndSpaceTrn = TEndSpaceTrn & EndSpaceTrn
	Next i

	'为调用消息输出，用原有变量替换连接后的数据
	acckeySrc = TacckeySrc
	EndStringSrc = TEndStringSrc
	ShortcutSrc = TShortcutSrc
	SpaceTrn = TSpaceTrn
	acckeyTrn = TacckeyTrn
	ExpStringTrn = TExpStringTrn
	EndStringTrn = TEndStringTrn
	ShortcutTrn = TShortcutTrn
	PreStringTrn = TPreStringTrn
	EndSpaceSrc = TEndSpaceSrc
	EndSpaceTrn = TEndSpaceTrn
End Function


'自定义参数
Function Settings(CheckID As Integer) As Integer
	Dim AppLngList() As String,UseLngList() As String

	If OSLanguage = "0404" Then
		Msg01 = "砞﹚"
		Msg02 = "叫匡砞﹚ゅよ遏い块代刚礚粇甅ノ龟悔巨"
		Msg05 = "弄(&R)"
		Msg06 = "穝糤(&A)"
		Msg07 = "跑(&M)"
		Msg08 = "埃(&D)"
		Msg09 = "睲(&C)"
		Msg10 = "代刚(&T)"

		Msg11 = "砞﹚睲虫"
		Msg12 = "纗摸"
		Msg13 = "砞﹚ず甧"
		Msg14 = "郎"
		Msg15 = "爹"
		Msg16 = "蹲砞﹚"
		Msg17 = "蹲砞﹚"
		Msg18 = "獽倍龄"
		Msg19 = "沧ゎ才"
		Msg20 = "硉竟"
		Msg21 = "じ蠢传"

		Msg22 = "璶逼埃 && 才腹獶獽倍龄じ舱 (硆腹だ筳):"
		Msg23 = "﹃だ澄ノ篨夹才 (ノ獽倍龄﹃矪瞶) (硆腹だ筳):"
		Msg24 = "璶浪琩獽倍龄玡珹腹ㄒ [&&F] (硆腹だ筳):"
		Msg25 = "璶玂痙獶獽倍龄才玡Θ癸じㄒ (&&) (硆腹だ筳):"
		Msg26 = "ゅ陪ボ盿珹腹獽倍龄 (硄盽ノㄈ瑆粂ē)"

		Msg27 = "璶浪琩沧ゎ才 (ノ - ボ絛瞅や穿窾ノじ) (だ筳):"
		Msg28 = "璶玂痙沧ゎ才舱 (ノ - ボ絛瞅や穿窾ノじ) (硆腹だ筳):"
		Msg29 = "璶笆蠢传沧ゎ才癸 (ノ | だ筳蠢传玡じ) (だ筳):"

		Msg30 = "璶浪琩硉竟篨夹才ㄒ \t (硆腹だ筳):"
		Msg31 = "璶浪琩硉竟じ (ノ - ボ絛瞅や穿窾ノじ) (硆腹だ筳):"
		Msg32 = "璶玂痙硉竟じ (ノ - ボ絛瞅や穿窾ノじ) (硆腹だ筳):"

		Msg33 = "璶笆蠢传じ (ノ | だ筳蠢传玡じ) (硆腹だ筳):"
		Msg34 = "陆亩璶砆蠢传じ (ノ | だ筳蠢传玡じ) (硆腹だ筳):"

		Msg35 = "弧(&H)"
		Msg37 = "﹃矪瞶"
		Msg54 = ">"
		Msg55 = "..."

		Msg71 = "続ノ粂ē"
		Msg74 = "穝糤  >"
		Msg75 = "场穝糤 >>"
		Msg76 = "<  埃"
		Msg77 = "<< 场埃"
		Msg78 = "穝糤ノ粂ē"
		Msg79 = "絪胯ノ粂ē"
		Msg80 = "埃ノ粂ē"
		Msg81 = "穝糤続ノ粂ē"
		Msg82 = "絪胯続ノ粂ē"
		Msg83 = "埃続ノ粂ē"

		Msg84 = "笆穝"
		Msg85 = "穝よΑ"
		Msg86 = "笆更穝杆(&A)"
		Msg87 = "Τ穝硄иパи∕﹚更杆(&M)"
		Msg88 = "闽超笆穝(&O)"
		Msg89 = "浪琩繵瞯"
		Msg90 = "浪琩丁筳: "
		Msg91 = "ぱ"
		Msg92 = "程浪琩ら戳:"
		Msg93 = "穝呼睲虫 (だ︽块玡纔)"
		Msg94 = "RAR 秆溃祘Α"
		Msg95 = "祘Α隔畖 (や穿吏挂跑计):"
		Msg96 = "秆溃把计 (%1 溃罽郎%2 璶耝郎%3 秆溃隔畖):"
		Msg97 = "浪琩"
	Else
		Msg01 = "配置"
		Msg02 = "请选择配置并在下列文本框中输入值，测试无误后再应用于实际操作。"
		Msg05 = "读取(&R)"
		Msg06 = "添加(&A)"
		Msg07 = "更改(&M)"
		Msg08 = "删除(&D)"
		Msg09 = "清空(&C)"
		Msg10 = "测试(&T)"

		Msg11 = "配置列表"
		Msg12 = "保存类型"
		Msg13 = "配置内容"
		Msg14 = "文件"
		Msg15 = "注册表"
		Msg16 = "导入配置"
		Msg17 = "导出配置"
		Msg18 = "快捷键"
		Msg19 = "终止符"
		Msg20 = "加速器"
		Msg21 = "字符替换"

		Msg22 = "要排除的含 && 符号的非快捷键字符组合 (半角逗号分隔):"
		Msg23 = "字串拆分用标志符 (用于多个快捷键字串的处理) (半角逗号分隔):"
		Msg24 = "要检查的快捷键前后括号，例如 [&&F] (半角逗号分隔):"
		Msg25 = "要保留的非快捷键符前后成对字符，例如 (&&) (半角逗号分隔):"
		Msg26 = "在文本后面显示带括号的快捷键 (通常用于亚洲语言)"

		Msg27 = "要检查的终止符 (用 - 表示范围，支持通配符) (空格分隔):"
		Msg28 = "要保留的终止符组合 (用 - 表示范围，支持通配符) (半角逗号分隔):"
		Msg29 = "要自动替换的终止符对 (用 | 分隔替换前后的字符) (空格分隔):"

		Msg30 = "要检查的加速器标志符，例如 \t (半角逗号分隔):"
		Msg31 = "要检查的加速器字符 (用 - 表示范围，支持通配符) (半角逗号分隔):"
		Msg32 = "要保留的加速器字符 (用 - 表示范围，支持通配符) (半角逗号分隔):"

		Msg33 = "要自动替换的字符 (用 | 分隔替换前后的字符) (半角逗号分隔):"
		Msg34 = "翻译后要被替换的字符 (用 | 分隔替换前后的字符) (半角逗号分隔):"

		Msg35 = "帮助(&H)"
		Msg37 = "字串处理"
		Msg54 = ">"
		Msg55 = "..."

		Msg71 = "适用语言"
		Msg74 = "添加  >"
		Msg75 = "全部添加 >>"
		Msg76 = "<  删除"
		Msg77 = "<< 全部删除"
		Msg78 = "添加可用语言"
		Msg79 = "编辑可用语言"
		Msg80 = "删除可用语言"
		Msg81 = "添加适用语言"
		Msg82 = "编辑适用语言"
		Msg83 = "删除适用语言"

		Msg84 = "自动更新"
		Msg85 = "更新方式"
		Msg86 = "自动下载更新并安装(&A)"
		Msg87 = "有更新时通知我，由我决定下载并安装(&M)"
		Msg88 = "关闭自动更新(&O)"
		Msg89 = "检查频率"
		Msg90 = "检查间隔: "
		Msg91 = "天"
		Msg92 = "最后检查日期:"
		Msg93 = "更新网址列表 (分行输入，前者优先)"
		Msg94 = "RAR 解压程序"
		Msg95 = "程序路径 (支持环境变量):"
		Msg96 = "解压参数 (%1 为压缩文件，%2 为要提取的文件，%3 为解压路径):"
		Msg97 = "检查"
	End If

	Begin Dialog UserDialog 620,462,Msg01,.SetFunc ' %GRID:10,7,1,1
		Text 20,7,580,14,Msg02
		OptionGroup .Options
			OptionButton 180,28,130,14,Msg37,.StrHandle
			OptionButton 360,28,130,14,Msg84,.AutoUpdate

		GroupBox 20,49,330,70,Msg11,.GroupBox1
		DropListBox 40,66,260,21,CheckList(),.CheckList
		PushButton 300,66,30,21,Msg55,.LevelButton
		PushButton 40,91,90,21,Msg06,.AddButton
		PushButton 140,91,90,21,Msg07,.ChangButton
		PushButton 240,91,90,21,Msg08,.DelButton

		GroupBox 370,49,230,70,Msg12,.GroupBox2
		OptionGroup .cWriteType
			OptionButton 390,69,90,14,Msg14,.cWriteToFile
			OptionButton 490,69,90,14,Msg15,.cWriteToRegistry
		PushButton 390,91,90,21,Msg16,.ImportButton
		PushButton 490,91,90,21,Msg17,.ExportButton

		GroupBox 20,147,580,280,Msg13,.GroupBox3
		OptionGroup .SetType
			OptionButton 45,129,100,14,Msg18
			OptionButton 155,129,100,14,Msg19
			OptionButton 265,129,100,14,Msg20
			OptionButton 375,129,100,14,Msg21
			OptionButton 485,129,100,14,Msg71
		Text 40,168,540,14,Msg22,.ExCrBoxTxt
		Text 40,217,530,14,Msg23,.LnSpBoxTxt
		Text 40,266,540,14,Msg24,.ChkBktBoxTxt
		Text 40,329,540,14,Msg25,.KpPairBoxTxt
		TextBox 40,189,540,21,.ExCrBox,1
		TextBox 40,238,540,21,.LnSpBox,1
		TextBox 40,287,540,35,.ChkBktBox,1
		TextBox 40,350,540,49,.KpPairBox,1
		CheckBox 40,406,540,14,Msg26,.AsiaKeyBox

		Text 40,168,540,14,Msg27,.ChkEndBoxTxt
		Text 40,252,530,14,Msg28,.NoTrnEndBoxTxt
		Text 40,336,540,14,Msg29,.AutoTrnEndBoxTxt
		TextBox 40,189,540,56,.ChkEndBox,1
		TextBox 40,273,540,56,.NoTrnEndBox,1
		TextBox 40,357,540,56,.AutoTrnEndBox,1

		Text 40,168,540,14,Msg30,.ShortBoxTxt
		Text 40,217,540,14,Msg31,.ShortKeyBoxTxt
		Text 40,329,540,14,Msg32,.KpShortKeyBoxTxt
		TextBox 40,189,540,21,.ShortBox,1
		TextBox 40,238,540,84,.ShortKeyBox,1
		TextBox 40,350,540,63,.KpShortKeyBox,1

		Text 40,168,540,14,Msg33,.PreRepStrBoxTxt
		Text 40,294,540,14,Msg34,.AutoWebFlagBoxTxt
		TextBox 40,189,540,98,.PreRepStrBox,1
		TextBox 40,189,540,224,.AutoWebFlagBox,1

		Text 40,161,190,14,Msg72,.AppLngText
		Text 390,161,190,14,Msg73,.UseLngText
		ListBox 40,175,190,245,AppLngList(),.AppLngList
		ListBox 390,175,190,245,UseLngList(),.UseLngList
		PushButton 250,175,120,21,Msg74,.AddLangButton
		PushButton 250,196,120,21,Msg75,.AddAllLangButton
		PushButton 250,224,120,21,Msg76,.DelLangButton
		PushButton 250,245,120,21,Msg77,.DelAllLangButton
		PushButton 250,280,120,21,Msg78,.SetAppLangButton
		PushButton 250,301,120,21,Msg79,.EditAppLangButton
		PushButton 250,322,120,21,Msg80,.DelAppLangButton
		PushButton 250,357,120,21,Msg81,.SetUseLangButton
		PushButton 250,378,120,21,Msg82,.EditUseLangButton
		PushButton 250,399,120,21,Msg83,.DelUseLangButton

		GroupBox 20,49,360,91,Msg85,.UpdateSetGroup
		OptionGroup .UpdateSet
			OptionButton 40,70,330,14,Msg86,.AutoButton
			OptionButton 40,91,330,14,Msg87,.ManualButton
			OptionButton 40,112,330,14,Msg88,.OffButton
		GroupBox 400,49,200,91,Msg89,.CheckGroup
		Text 420,70,80,14,Msg90,.UpdateCycleText
		TextBox 510,68,40,18,.UpdateCycleBox
		Text 560,70,30,14,Msg91,.UpdateDatesText
		Text 420,91,160,14,Msg92,.UpdateDateText
		TextBox 420,112,100,18,.UpdateDateBox
		PushButton 530,110,60,21,Msg97,.CheckButton
		GroupBox 20,154,580,133,Msg93,.WebSiteGroup
		TextBox 40,175,540,98,.WebSiteBox,1
		GroupBox 20,301,580,126,Msg94,.CmdGroup
		Text 40,322,510,14,Msg95,.CmdPathBoxText
		Text 40,371,510,14,Msg96,.ArgumentBoxText
		TextBox 40,343,510,21,.CmdPathBox
		TextBox 40,392,510,21,.ArgumentBox
		PushButton 550,343,30,21,Msg55,.ExeBrowseButton
		PushButton 550,392,30,21,Msg54,.ArgumentButton

		PushButton 20,434,90,21,Msg35,.HelpButton
		PushButton 110,434,100,21,Msg05,.ResetButton
		PushButton 300,434,90,21,Msg10,.TestButton
		PushButton 210,434,90,21,Msg09,.CleanButton
		OKButton 430,434,90,21,.OKButton
		CancelButton 520,434,80,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	dlg.CheckList = CheckID
	If Dialog(dlg) = 0 Then Exit Function
	Settings = dlg.CheckList
End Function


'请务必查看对话框帮助主题以了解更多信息。
Private Function SetFunc%(DlgItem$, Action%, SuppValue&)
	Dim Header As String,HeaderID As Integer,NewData As String,Path As String,cStemp As Boolean
	Dim i As Integer,n As Integer,TempArray() As String,Temp As String,LngName As String
	Dim LngID As Integer,LangArray() As String,AppLngList() As String,UseLngList() As String

	If OSLanguage = "0404" Then
		Msg01 = "岿粇"
		Msg02 = "箇砞"
		Msg03 = ""
		Msg04 = "把酚"
		Msg08 = "ゼ"
		Msg11 = "牡"
		Msg12 =	"狦琘ㄇ把计盢ㄏ祘Α磅︽挡狦ぃタ絋" & vbCrLf & _
				"眤絋龟稱璶硂妓暗盾"
		Msg13 = "砞﹚ず甧竒跑琌⊿Τ纗琌惠璶纗"
		Msg14 = "纗摸竒跑琌⊿Τ纗琌惠璶纗"
		Msg18 = "ヘ玡砞﹚いぶΤ兜把计" & vbCrLf
		Msg19 = "┮Τ砞﹚いぶΤ兜把计" & vbCrLf
		Msg21 = "絋粄"
		Msg22 = "絋龟璶埃砞﹚%s盾"
		Msg24 = "絋龟璶埃粂ē%s盾"
		Msg30 = "癟"
		Msg32 = "蹲砞﹚Θ"
		Msg33 = "蹲砞﹚Θ"
		Msg36 = "礚猭纗叫浪琩琌Τ糶竚舦:" & vbCrLf & vbCrLf
		Msg39 = "蹲ア毖叫浪琩琌Τ糶竚舦" & vbCrLf & _
				"┪蹲郎Α琌タ絋:" & vbCrLf & vbCrLf
		Msg40 = "蹲ア毖叫浪琩琌Τ糶竚舦" & vbCrLf & _
				"┪蹲郎Α琌タ絋:" & vbCrLf & vbCrLf
		Msg41 = "纗ア毖叫浪琩琌Τ糶竚舦:" & vbCrLf & vbCrLf
		Msg42 = "匡璶蹲郎"
		Msg43 = "匡璶蹲郎"
		Msg44 = "砞﹚郎 (*.dat)|*.dat|┮Τ郎 (*.*)|*.*||"
		Msg54 = "ノ粂ē:"
		Msg55 = "続ノ粂ē:"
		Msg60 = "匡秆溃祘Α"
		Msg61 = "磅︽郎 (*.exe)|*.exe|┮Τ郎 (*.*)|*.*||"
		Msg62 = "⊿Τ﹚秆溃祘Α叫穝块┪匡"
		Msg63 = "郎把酚把计(%1)"
		Msg64 = "璶耝郎把计(%2)"
		Msg65 = "秆溃隔畖把计(%3)"
	Else
		Msg01 = "错误"
		Msg02 = "默认值"
		Msg03 = "原值"
		Msg04 = "参照值"
		Msg08 = "未知"
		Msg11 = "警告"
		Msg12 =	"如果某些参数为空，将使程序运行结果不正确。" & vbCrLf & _
				"您确实想要这样做吗？"
		Msg13 = "配置内容已经更改但是没有保存！是否需要保存？"
		Msg14 = "保存类型已经更改但是没有保存！是否需要保存？"
		Msg18 = "当前配置中，至少有一项参数为空！" & vbCrLf
		Msg19 = "所有配置中，至少有一项参数为空！" & vbCrLf
		Msg21 = "确认"
		Msg22 = "确实要删除配置“%s”吗？"
		Msg24 = "确实要删除语言“%s”吗？"
		Msg30 = "信息"
		Msg32 = "导出配置成功！"
		Msg33 = "导入配置成功！"
		Msg36 = "无法保存！请检查是否有写入下列位置的权限:" & vbCrLf & vbCrLf
		Msg39 = "导入失败！请检查是否有写入下列位置的权限" & vbCrLf & _
				"或导入文件的格式是否正确:" & vbCrLf & vbCrLf
		Msg40 = "导出失败！请检查是否有写入下列位置的权限" & vbCrLf & _
				"或导出文件的格式是否正确:" & vbCrLf & vbCrLf
		Msg41 = "保存失败！请检查是否有写入下列位置的权限:" & vbCrLf & vbCrLf
		Msg42 = "选择要导入的文件"
		Msg43 = "选择要导出的文件"
		Msg44 = "配置文件 (*.dat)|*.dat|所有文件 (*.*)|*.*||"
		Msg54 = "可用语言:"
		Msg55 = "适用语言:"
		Msg60 = "选择解压程序"
		Msg61 = "可执行文件 (*.exe)|*.exe|所有文件 (*.*)|*.*||"
		Msg62 = "没有指定解压程序！请重新输入或选择。"
		Msg63 = "文件引用参数(%1)"
		Msg64 = "要提取的文件参数(%2)"
		Msg65 = "解压路径参数(%3)"
	End If

	If DlgValue("Options") = 0 Then
		DlgVisible "GroupBox1",True
		DlgVisible "AddButton",True
		DlgVisible "ChangButton",True
		DlgVisible "DelButton",True

		DlgVisible "GroupBox2",True
		DlgVisible "ImportButton",True
		DlgVisible "ExportButton",True
		DlgVisible "CheckList",True
		DlgVisible "LevelButton",True
		DlgVisible "cWriteType",True

		DlgVisible "GroupBox3",True
		DlgVisible "SetType",True
		If DlgValue("SetType") = 0 Then
			DlgVisible "ExCrBoxTxt",True
			DlgVisible "LnSpBoxTxt",True
			DlgVisible "ChkBktBoxTxt",True
			DlgVisible "KpPairBoxTxt",True
			DlgVisible "ExCrBox",True
			DlgVisible "LnSpBox",True
			DlgVisible "ChkBktBox",True
			DlgVisible "KpPairBox",True
			DlgVisible "AsiaKeyBox",True

			DlgVisible "ChkEndBoxTxt",False
			DlgVisible "NoTrnEndBoxTxt",False
			DlgVisible "AutoTrnEndBoxTxt",False
			DlgVisible "ChkEndBox",False
			DlgVisible "NoTrnEndBox",False
			DlgVisible "AutoTrnEndBox",False

			DlgVisible "ShortBoxTxt",False
			DlgVisible "ShortKeyBoxTxt",False
			DlgVisible "KpShortKeyBoxTxt",False
			DlgVisible "ShortBox",False
			DlgVisible "ShortKeyBox",False
			DlgVisible "KpShortKeyBox",False

			DlgVisible "PreRepStrBoxTxt",False
			DlgVisible "AutoWebFlagBoxTxt",False
			DlgVisible "PreRepStrBox",False
			DlgVisible "AutoWebFlagBox",False

			DlgVisible "AppLngText",False
			DlgVisible "UseLngText",False
			DlgVisible "AppLngList",False
			DlgVisible "UseLngList",False
			DlgVisible "AddLangButton",False
			DlgVisible "AddAllLangButton",False
			DlgVisible "DelLangButton",False
			DlgVisible "DelAllLangButton",False
			DlgVisible "SetAppLangButton",False
			DlgVisible "EditAppLangButton",False
			DlgVisible "DelAppLangButton",False
			DlgVisible "SetUseLangButton",False
			DlgVisible "EditUseLangButton",False
			DlgVisible "DelUseLangButton",False
		ElseIf DlgValue("SetType") = 1 Then
			DlgVisible "ExCrBoxTxt",False
			DlgVisible "LnSpBoxTxt",False
			DlgVisible "ChkBktBoxTxt",False
			DlgVisible "KpPairBoxTxt",False
			DlgVisible "ExCrBox",False
			DlgVisible "LnSpBox",False
			DlgVisible "ChkBktBox",False
			DlgVisible "KpPairBox",False
			DlgVisible "AsiaKeyBox",False

			DlgVisible "ChkEndBoxTxt",True
			DlgVisible "NoTrnEndBoxTxt",True
			DlgVisible "AutoTrnEndBoxTxt",True
			DlgVisible "ChkEndBox",True
			DlgVisible "NoTrnEndBox",True
			DlgVisible "AutoTrnEndBox",True

			DlgVisible "ShortBoxTxt",False
			DlgVisible "ShortKeyBoxTxt",False
			DlgVisible "KpShortKeyBoxTxt",False
			DlgVisible "ShortBox",False
			DlgVisible "ShortKeyBox",False
			DlgVisible "KpShortKeyBox",False

			DlgVisible "PreRepStrBoxTxt",False
			DlgVisible "AutoWebFlagBoxTxt",False
			DlgVisible "PreRepStrBox",False
			DlgVisible "AutoWebFlagBox",False

			DlgVisible "AppLngText",False
			DlgVisible "UseLngText",False
			DlgVisible "AppLngList",False
			DlgVisible "UseLngList",False
			DlgVisible "AddLangButton",False
			DlgVisible "AddAllLangButton",False
			DlgVisible "DelLangButton",False
			DlgVisible "DelAllLangButton",False
			DlgVisible "SetAppLangButton",False
			DlgVisible "EditAppLangButton",False
			DlgVisible "DelAppLangButton",False
			DlgVisible "SetUseLangButton",False
			DlgVisible "EditUseLangButton",False
			DlgVisible "DelUseLangButton",False
		ElseIf DlgValue("SetType") = 2 Then
			DlgVisible "ExCrBoxTxt",False
			DlgVisible "LnSpBoxTxt",False
			DlgVisible "ChkBktBoxTxt",False
			DlgVisible "KpPairBoxTxt",False
			DlgVisible "ExCrBox",False
			DlgVisible "LnSpBox",False
			DlgVisible "ChkBktBox",False
			DlgVisible "KpPairBox",False
			DlgVisible "AsiaKeyBox",False

			DlgVisible "ChkEndBoxTxt",False
			DlgVisible "NoTrnEndBoxTxt",False
			DlgVisible "AutoTrnEndBoxTxt",False
			DlgVisible "ChkEndBox",False
			DlgVisible "NoTrnEndBox",False
			DlgVisible "AutoTrnEndBox",False

			DlgVisible "ShortBoxTxt",True
			DlgVisible "ShortKeyBoxTxt",True
			DlgVisible "KpShortKeyBoxTxt",True
			DlgVisible "ShortBox",True
			DlgVisible "ShortKeyBox",True
			DlgVisible "KpShortKeyBox",True

			DlgVisible "PreRepStrBoxTxt",False
			DlgVisible "AutoWebFlagBoxTxt",False
			DlgVisible "PreRepStrBox",False
			DlgVisible "AutoWebFlagBox",False

			DlgVisible "AppLngText",False
			DlgVisible "UseLngText",False
			DlgVisible "AppLngList",False
			DlgVisible "UseLngList",False
			DlgVisible "AddLangButton",False
			DlgVisible "AddAllLangButton",False
			DlgVisible "DelLangButton",False
			DlgVisible "DelAllLangButton",False
			DlgVisible "SetAppLangButton",False
			DlgVisible "EditAppLangButton",False
			DlgVisible "DelAppLangButton",False
			DlgVisible "SetUseLangButton",False
			DlgVisible "EditUseLangButton",False
			DlgVisible "DelUseLangButton",False
		ElseIf DlgValue("SetType") = 3 Then
			DlgVisible "ExCrBoxTxt",False
			DlgVisible "LnSpBoxTxt",False
			DlgVisible "ChkBktBoxTxt",False
			DlgVisible "KpPairBoxTxt",False
			DlgVisible "ExCrBox",False
			DlgVisible "LnSpBox",False
			DlgVisible "ChkBktBox",False
			DlgVisible "KpPairBox",False
			DlgVisible "AsiaKeyBox",False

			DlgVisible "ShortBoxTxt",False
			DlgVisible "ShortKeyBoxTxt",False
			DlgVisible "KpShortKeyBoxTxt",False
			DlgVisible "ShortBox",False
			DlgVisible "ShortKeyBox",False
			DlgVisible "KpShortKeyBox",False

			DlgVisible "ChkEndBoxTxt",False
			DlgVisible "NoTrnEndBoxTxt",False
			DlgVisible "AutoTrnEndBoxTxt",False
			DlgVisible "ChkEndBox",False
			DlgVisible "NoTrnEndBox",False
			DlgVisible "AutoTrnEndBox",False

			DlgVisible "PreRepStrBoxTxt",True
			DlgVisible "AutoWebFlagBoxTxt",False
			DlgVisible "PreRepStrBox",False
			DlgVisible "AutoWebFlagBox",True

			DlgVisible "AppLngText",False
			DlgVisible "UseLngText",False
			DlgVisible "AppLngList",False
			DlgVisible "UseLngList",False
			DlgVisible "AddLangButton",False
			DlgVisible "AddAllLangButton",False
			DlgVisible "DelLangButton",False
			DlgVisible "DelAllLangButton",False
			DlgVisible "SetAppLangButton",False
			DlgVisible "EditAppLangButton",False
			DlgVisible "DelAppLangButton",False
			DlgVisible "SetUseLangButton",False
			DlgVisible "EditUseLangButton",False
			DlgVisible "DelUseLangButton",False
		ElseIf DlgValue("SetType") = 4 Then
			DlgVisible "ExCrBoxTxt",False
			DlgVisible "LnSpBoxTxt",False
			DlgVisible "ChkBktBoxTxt",False
			DlgVisible "KpPairBoxTxt",False
			DlgVisible "ExCrBox",False
			DlgVisible "LnSpBox",False
			DlgVisible "ChkBktBox",False
			DlgVisible "KpPairBox",False
			DlgVisible "AsiaKeyBox",False

			DlgVisible "ShortBoxTxt",False
			DlgVisible "ShortKeyBoxTxt",False
			DlgVisible "KpShortKeyBoxTxt",False
			DlgVisible "ShortBox",False
			DlgVisible "ShortKeyBox",False
			DlgVisible "KpShortKeyBox",False

			DlgVisible "ChkEndBoxTxt",False
			DlgVisible "NoTrnEndBoxTxt",False
			DlgVisible "AutoTrnEndBoxTxt",False
			DlgVisible "ChkEndBox",False
			DlgVisible "NoTrnEndBox",False
			DlgVisible "AutoTrnEndBox",False

			DlgVisible "PreRepStrBoxTxt",False
			DlgVisible "AutoWebFlagBoxTxt",False
			DlgVisible "PreRepStrBox",False
			DlgVisible "AutoWebFlagBox",False

			DlgVisible "AppLngText",True
			DlgVisible "UseLngText",True
			DlgVisible "AppLngList",True
			DlgVisible "UseLngList",True
			DlgVisible "AddLangButton",True
			DlgVisible "AddAllLangButton",True
			DlgVisible "DelLangButton",True
			DlgVisible "DelAllLangButton",True
			DlgVisible "SetAppLangButton",True
			DlgVisible "EditAppLangButton",True
			DlgVisible "DelAppLangButton",True
			DlgVisible "SetUseLangButton",True
			DlgVisible "EditUseLangButton",True
			DlgVisible "DelUseLangButton",True
		End If
		DlgVisible "UpdateSetGroup",False
		DlgVisible "UpdateSet",False
		DlgVisible "AutoButton",False
		DlgVisible "ManualButton",False
		DlgVisible "OffButton",False
		DlgVisible "CheckGroup",False
		DlgVisible "UpdateCycleText",False
		DlgVisible "UpdateCycleBox",False
		DlgVisible "UpdateDatesText",False
		DlgVisible "UpdateDateText",False
		DlgVisible "UpdateDateBox",False
		DlgVisible "CheckButton",False
		DlgVisible "WebSiteGroup",False
		DlgVisible "WebSiteBox",False
		DlgVisible "CmdGroup",False
		DlgVisible "CmdPathBoxText",False
		DlgVisible "ArgumentBoxText",False
		DlgVisible "CmdPathBox",False
		DlgVisible "ArgumentBox",False
		DlgVisible "ExeBrowseButton",False
		DlgVisible "ArgumentButton",False
	ElseIf DlgValue("Options") = 1 Then
		DlgVisible "GroupBox1",False
		DlgVisible "AddButton",False
		DlgVisible "ChangButton",False
		DlgVisible "DelButton",False

		DlgVisible "GroupBox2",False
		DlgVisible "ImportButton",False
		DlgVisible "ExportButton",False

		DlgVisible "CheckList",False
		DlgVisible "LevelButton",False
		DlgVisible "cWriteType",False
		DlgVisible "SetType",False
		DlgVisible "GroupBox3",False

		DlgVisible "ExCrBoxTxt",False
		DlgVisible "LnSpBoxTxt",False
		DlgVisible "ChkBktBoxTxt",False
		DlgVisible "KpPairBoxTxt",False
		DlgVisible "ExCrBox",False
		DlgVisible "LnSpBox",False
		DlgVisible "ChkBktBox",False
		DlgVisible "KpPairBox",False
		DlgVisible "AsiaKeyBox",False

		DlgVisible "ChkEndBoxTxt",False
		DlgVisible "NoTrnEndBoxTxt",False
		DlgVisible "AutoTrnEndBoxTxt",False
		DlgVisible "ChkEndBox",False
		DlgVisible "NoTrnEndBox",False
		DlgVisible "AutoTrnEndBox",False

		DlgVisible "ShortBoxTxt",False
		DlgVisible "ShortKeyBoxTxt",False
		DlgVisible "KpShortKeyBoxTxt",False
		DlgVisible "ShortBox",False
		DlgVisible "ShortKeyBox",False
		DlgVisible "KpShortKeyBox",False

		DlgVisible "PreRepStrBoxTxt",False
		DlgVisible "AutoWebFlagBoxTxt",False
		DlgVisible "PreRepStrBox",False
		DlgVisible "AutoWebFlagBox",False

		DlgVisible "AppLngText",False
		DlgVisible "UseLngText",False
		DlgVisible "AppLngList",False
		DlgVisible "UseLngList",False
		DlgVisible "AddLangButton",False
		DlgVisible "AddAllLangButton",False
		DlgVisible "DelLangButton",False
		DlgVisible "DelAllLangButton",False
		DlgVisible "SetAppLangButton",False
		DlgVisible "EditAppLangButton",False
		DlgVisible "DelAppLangButton",False
		DlgVisible "SetUseLangButton",False
		DlgVisible "EditUseLangButton",False
		DlgVisible "DelUseLangButton",False

		DlgVisible "UpdateSetGroup",True
		DlgVisible "UpdateSet",True
		DlgVisible "AutoButton",True
		DlgVisible "ManualButton",True
		DlgVisible "OffButton",True
		DlgVisible "CheckGroup",True
		DlgVisible "UpdateCycleText",True
		DlgVisible "UpdateCycleBox",True
		DlgVisible "UpdateDatesText",True
		DlgVisible "UpdateDateText",True
		DlgVisible "UpdateDateBox",True
		DlgVisible "CheckButton",True
		DlgVisible "WebSiteGroup",True
		DlgVisible "WebSiteBox",True
		DlgVisible "CmdGroup",True
		DlgVisible "CmdPathBoxText",True
		DlgVisible "ArgumentBoxText",True
		DlgVisible "CmdPathBox",True
		DlgVisible "ArgumentBox",True
		DlgVisible "ExeBrowseButton",True
		DlgVisible "ArgumentButton",True
	End If

	Select Case Action%
	Case 1 ' 对话框窗口初始化
		If Join(CheckList) <> "" Then
			HeaderID = DlgValue("CheckList")
			TempArray = Split(CheckDataList(HeaderID),JoinStr)
			SetsArray = Split(TempArray(1),SubJoinStr)
			If InStr(SetsArray(11),rSubJoinStr) Then
				PreStr = Replace(SetsArray(11),rSubJoinStr,SubJoinStr)
			Else
				PreStr = SetsArray(11)
			End If
			DlgText "ExCrBox",SetsArray(0)
			DlgText "LnSpBox",SetsArray(1)
			DlgText "ChkBktBox",SetsArray(2)
			DlgText "KpPairBox",SetsArray(3)
			DlgValue "AsiaKeyBox",StrToInteger(SetsArray(4))
			DlgText "ChkEndBox",SetsArray(5)
			DlgText "NoTrnEndBox",SetsArray(6)
			DlgText "AutoTrnEndBox",SetsArray(7)
			DlgText "ShortBox",SetsArray(8)
			DlgText "ShortKeyBox",SetsArray(9)
			DlgText "KpShortKeyBox",SetsArray(10)
			DlgText "PreRepStrBox",PreStr
			DlgText "AutoWebFlagBox",SetsArray(12)
			getLngNameList(TempArray(2),AppLngList,UseLngList)
			DlgListBoxArray "AppLngList",AppLngList()
			DlgListBoxArray "UseLngList",UseLngList()
			DlgValue "AppLngList",0
			DlgValue "UseLngList",0
		End If

		If cWriteLoc = CheckFilePath Then DlgValue "cWriteType",0
		If cWriteLoc = CheckRegKey Then DlgValue "cWriteType",1
		If cWriteLoc = "" Then DlgValue "cWriteType",0

		If DlgText("AppLngList") = "" Then
			DlgText "AppLngText",Msg54 & "(0)"
		Else
			DlgText "AppLngText",Msg54 & "(" & UBound(AppLngList)+1 & ")"
		End If
		If DlgText("UseLngList") = "" Then
			DlgText "UseLngText",Msg55 & "(0)"
		Else
			DlgText "UseLngText",Msg55 & "(" & UBound(UseLngList)+1 & ")"
		End If

		If Join(cUpdateSet) <> "" Then
			DlgValue "UpdateSet",StrToInteger(cUpdateSet(0))
			DlgText "WebSiteBox",cUpdateSet(1)
			DlgText "CmdPathBox",cUpdateSet(2)
			DlgText "ArgumentBox",cUpdateSet(3)
			DlgText "UpdateCycleBox",cUpdateSet(4)
			DlgText "UpdateDateBox",cUpdateSet(5)
		End If
		If DlgText("UpdateDateBox") = "" Then DlgText "UpdateDateBox",Msg08
		DlgEnable "UpdateDateBox",False

		If DlgText("AppLngList") = "" Then
			DlgEnable "AddLangButton",False
			DlgEnable "AddAllLangButton",False
			DlgEnable "EditAppLangButton",False
			DlgEnable "DelAppLangButton",False
		End If
		If DlgText("UseLngList") = "" Then
			DlgEnable "DelLangButton",False
			DlgEnable "DelAllLangButton",False
			DlgEnable "EditUseLangButton",False
			DlgEnable "DelUseLangButton",False
		End If
		Header = DlgText("CheckList")
		If InStr(Join(DefaultCheckList,JoinStr),Header) Then
			DlgEnable "ChangButton",False
			DlgEnable "DelButton",False
		Else
			If UBound(CheckList) = 0 Then DlgEnable "DelButton",False
		End If
	Case 2 ' 数值更改或者按下了按钮
		If DlgValue("Options") = 0 Then
			HeaderID = DlgValue("CheckList")
			TempArray = Split(CheckDataList(HeaderID),JoinStr)
			SetsArray = Split(TempArray(1),SubJoinStr)
			getLngNameList(TempArray(2),AppLngList,UseLngList)

			If DlgItem$ = "CheckList" Then
				If InStr(SetsArray(11),rSubJoinStr) Then
					PreStr = Replace(SetsArray(11),rSubJoinStr,SubJoinStr)
				Else
					PreStr = SetsArray(11)
				End If
				DlgText "ExCrBox",SetsArray(0)
				DlgText "LnSpBox",SetsArray(1)
				DlgText "ChkBktBox",SetsArray(2)
				DlgText "KpPairBox",SetsArray(3)
				DlgValue "AsiaKeyBox",StrToInteger(SetsArray(4))
				DlgText "ChkEndBox",SetsArray(5)
				DlgText "NoTrnEndBox",SetsArray(6)
				DlgText "AutoTrnEndBox",SetsArray(7)
				DlgText "ShortBox",SetsArray(8)
				DlgText "ShortKeyBox",SetsArray(9)
				DlgText "KpShortKeyBox",SetsArray(10)
				DlgText "PreRepStrBox",PreStr
				DlgText "AutoWebFlagBox",SetsArray(12)
				DlgListBoxArray "AppLngList",AppLngList()
				DlgListBoxArray "UseLngList",UseLngList()
				DlgValue "AppLngList",0
				DlgValue "UseLngList",0
			End If

			If DlgItem$ = "LevelButton" Then
				Header = DlgText("CheckList")
				If SetLevel(CheckList,CheckDataList) = True Then
					DlgListBoxArray "CheckList",CheckList()
					DlgText "CheckList",Header
				End If
			End If

			If DlgItem$ = "AddButton" Then
				NewData = AddSet(CheckList)
				If NewData <> "" Then
					For i = 0 To UBound(SetsArray)
						SetsArray(i) = ""
					Next i
					Data = Join(SetsArray,SubJoinStr)
					LangPairList = LangCodeList("",OSLanguage,1,107)
					Temp = NewData & JoinStr & Data & JoinStr & Join(LangPairList,SubJoinStr)
					CreateArray(NewData,Temp,CheckList,CheckDataList)
					DlgListBoxArray "CheckList",CheckList()
					DlgText "CheckList",NewData
					DlgText "ExCrBox",""
					DlgText "LnSpBox",""
					DlgText "ChkBktBox",""
					DlgText "KpPairBox",""
					DlgValue "AsiaKeyBox",0
					DlgText "ChkEndBox",""
					DlgText "NoTrnEndBox",""
					DlgText "AutoTrnEndBox",""
					DlgText "ShortBox",""
					DlgText "ShortKeyBox",""
					DlgText "KpShortKeyBox",""
					DlgText "PreRepStrBox",""
					DlgText "AutoWebFlagBox",""
					HeaderID = DlgValue("CheckList")
					TempArray = Split(CheckDataList(HeaderID),JoinStr)
					getLngNameList(TempArray(2),AppLngList,UseLngList)
					DlgListBoxArray "AppLngList",AppLngList()
					DlgListBoxArray "UseLngList",UseLngList()
					DlgValue "AppLngList",0
					DlgValue "UseLngList",0
				End If
			End If

			If DlgItem$ = "ChangButton" Then
				Header = DlgText("CheckList")
				NewData = EditSet(CheckList,Header)
				If NewData <> "" Then
					CheckList(HeaderID) = NewData
					CheckDataList(HeaderID) = NewData & JoinStr & TempArray(1) & JoinStr & TempArray(2)
					DlgListBoxArray "CheckList",CheckList()
					DlgValue "CheckList",HeaderID
				End If
			End If

	    	If DlgItem$ = "DelButton" Then
				Header = DlgText("CheckList")
				Msg = Replace(Msg22,"%s",Header)
				If MsgBox(Msg,vbYesNo+vbInformation,Msg21) = vbYes Then
					If HeaderID > 0 And HeaderID = UBound(CheckList) Then HeaderID = HeaderID - 1
					CheckList = DelArray(Header,CheckList,"",0)
					CheckDataList = DelArray(Header,CheckDataList,JoinStr,0)
					DlgListBoxArray "CheckList",CheckList()
					DlgValue "CheckList",HeaderID
					TempArray = Split(CheckDataList(HeaderID),JoinStr)
					SetsArray = Split(TempArray(1),SubJoinStr)
					If InStr(SetsArray(11),rSubJoinStr) Then
						PreStr = Replace(SetsArray(11),rSubJoinStr,SubJoinStr)
					Else
						PreStr = SetsArray(11)
					End If
					DlgText "ExCrBox",SetsArray(0)
					DlgText "LnSpBox",SetsArray(1)
					DlgText "ChkBktBox",SetsArray(2)
					DlgText "KpPairBox",SetsArray(3)
					DlgValue "AsiaKeyBox",StrToInteger(SetsArray(4))
					DlgText "ChkEndBox",SetsArray(5)
					DlgText "NoTrnEndBox",SetsArray(6)
					DlgText "AutoTrnEndBox",SetsArray(7)
					DlgText "ShortBox",SetsArray(8)
					DlgText "ShortKeyBox",SetsArray(9)
					DlgText "KpShortKeyBox",SetsArray(10)
					DlgText "PreRepStrBox",PreStr
					DlgText "AutoWebFlagBox",SetsArray(12)
					getLngNameList(TempArray(2),AppLngList,UseLngList)
					DlgListBoxArray "AppLngList",AppLngList()
					DlgListBoxArray "UseLngList",UseLngList()
					DlgValue "AppLngList",0
					DlgValue "UseLngList",0
				End If
			End If

			If DlgItem$ = "ResetButton" Then
				Header = DlgText("CheckList")
				ReDim TempArray(1)
				If InStr(Join(DefaultCheckList,JoinStr),Header) Then TempArray(0) = Msg02
				cStemp = CheckNullData(Header,CheckDataListBak,"4",0)
				If cStemp = False Then TempArray(1) = Msg03
				For i = LBound(CheckList) To UBound(CheckList)
					If i <> HeaderID Then
						ReDim Preserve TempArray(i+2)
						TempArray(i+2) = Msg04 & " - " & CheckList(i)
					End If
				Next i
				i = ShowPopupMenu(TempArray,vbPopupUseRightButton)
				If i = 0 Then
					SetsArray = Split(CheckSettings(Header,OSLanguage),SubJoinStr)
					TempArray = LangCodeList(Header,OSLanguage,1,107)
					Temp = Join(TempArray,SubJoinStr)
				ElseIf i = 1 Then
					For n = LBound(CheckDataListBak) To UBound(CheckDataListBak)
						TempArray = Split(CheckDataListBak(n),JoinStr)
						If TempArray(0) = Header Then
							SetsArray = Split(TempArray(1),SubJoinStr)
							Temp = TempArray(2)
							Exit For
						End If
					Next n
				ElseIf i >= 2 Then
					For n = LBound(CheckList) To UBound(CheckList)
						Temp = Mid(TempArray(i),InStr(TempArray(i),Msg04 & " - ") + Len(Msg04 & " - "))
						If Temp = CheckList(n) Then
							HeaderID = n
							Exit For
						End If
					Next n
					TempArray = Split(CheckDataList(HeaderID),JoinStr)
					SetsArray = Split(TempArray(1),SubJoinStr)
					Temp = TempArray(2)
				End If
				If i >= 0 Then
					If InStr(SetsArray(11),rSubJoinStr) Then
						PreStr = Replace(SetsArray(11),rSubJoinStr,SubJoinStr)
					Else
						PreStr = SetsArray(11)
					End If

					If DlgValue("SetType") = 0 Then
						DlgText "ExCrBox",SetsArray(0)
						DlgText "LnSpBox",SetsArray(1)
						DlgText "ChkBktBox",SetsArray(2)
						DlgText "KpPairBox",SetsArray(3)
						DlgValue "AsiaKeyBox",StrToInteger(SetsArray(4))
					ElseIf DlgValue("SetType") = 1 Then
						DlgText "ChkEndBox",SetsArray(5)
						DlgText "NoTrnEndBox",SetsArray(6)
						DlgText "AutoTrnEndBox",SetsArray(7)
					ElseIf DlgValue("SetType") = 2 Then
						DlgText "ShortBox",SetsArray(8)
						DlgText "ShortKeyBox",SetsArray(9)
						DlgText "KpShortKeyBox",SetsArray(10)
					ElseIf DlgValue("SetType") = 3 Then
						DlgText "PreRepStrBox",PreStr
						DlgText "AutoWebFlagBox",SetsArray(12)
					ElseIf DlgValue("SetType") = 4 Then
						getLngNameList(Temp,AppLngList,UseLngList)
						DlgListBoxArray "AppLngList",AppLngList()
						DlgListBoxArray "UseLngList",UseLngList()
						DlgValue "AppLngList",0
						DlgValue "UseLngList",0
					End If
				End If
			End If

	    	If DlgItem$ = "CleanButton" Then
	    		If DlgValue("SetType") = 0 Then
					DlgText "ExCrBox",""
					DlgText "LnSpBox",""
					DlgText "ChkBktBox",""
					DlgText "KpPairBox",""
					DlgValue "AsiaKeyBox",0
				ElseIf DlgValue("SetType") = 1 Then
					DlgText "ChkEndBox",""
					DlgText "NoTrnEndBox",""
					DlgText "AutoTrnEndBox",""
				ElseIf DlgValue("SetType") = 2 Then
					DlgText "ShortBox",""
					DlgText "ShortKeyBox",""
					DlgText "KpShortKeyBox",""
				ElseIf DlgValue("SetType") = 3 Then
					DlgText "PreRepStrBox",""
					DlgText "AutoWebFlagBox",""
				ElseIf DlgValue("SetType") = 4 Then
					ReDim UseLngList(0)
					AppLngList = ChangeLngNameList(TempArray(2),UseLngList)
					DlgListBoxArray "AppLngList",AppLngList()
					DlgListBoxArray "UseLngList",UseLngList()
					DlgValue "AppLngList",0
					DlgValue "UseLngList",0
				End If
			End If

			If DlgItem$ = "AddLangButton" Or DlgItem$ = "DelLangButton" Then
				If DlgItem$ = "AddLangButton" Then
					LngName = DlgText("AppLngList")
					If LngName <> "" Then
						LngID = DlgValue("AppLngList")
						i = UBound(AppLngList)
						AppLngList = DelArray(LngName,AppLngList,"",0)
						UseLngList = ChangeLngNameList(TempArray(2),AppLngList)
					End If
				Else
					LngName = DlgText("UseLngList")
					If LngName <> "" Then
						LngID = DlgValue("UseLngList")
						i = UBound(UseLngList)
						UseLngList = DelArray(LngName,UseLngList,"",0)
						AppLngList = ChangeLngNameList(TempArray(2),UseLngList)
					End If
				End If
				If LngName <> "" Then
					If LngID > 0 And LngID = i Then LngID = LngID - 1
					DlgListBoxArray "AppLngList",AppLngList()
					DlgListBoxArray "UseLngList",UseLngList()
					If DlgItem$ = "AddLangButton" Then
						DlgValue "AppLngList",LngID
						DlgText "UseLngList",LngName
					Else
						DlgText "AppLngList",LngName
						DlgValue "UseLngList",LngID
					End If
				End If
			End If

			If DlgItem$ = "AddAllLangButton" Or DlgItem$ = "DelAllLangButton" Then
				If DlgItem$ = "AddAllLangButton" Then
					ReDim AppLngList(0)
					UseLngList = ChangeLngNameList(TempArray(2),AppLngList)
				Else
					ReDim UseLngList(0)
					AppLngList = ChangeLngNameList(TempArray(2),UseLngList)
				End If
				DlgListBoxArray "AppLngList",AppLngList()
				DlgListBoxArray "UseLngList",UseLngList()
				DlgValue "AppLngList",0
				DlgValue "UseLngList",0
			End If

			If DlgItem$ = "SetAppLangButton" Or DlgItem$ = "SetUseLangButton" Then
				LangArray = Split(TempArray(2),SubJoinStr)
				NewData = SetLang(LangArray,"","")
				If NewData <> "" Then
					n = UBound(LangArray) + 1
					ReDim Preserve LangArray(n)
					LangArray(n) = NewData
					TempArray(2) = Join(LangArray,SubJoinStr)

					LangPairList = Split(NewData,LngJoinStr)
					NewLngName = LangPairList(0)
					NewLngCode = LangPairList(1)
					If DlgItem$ = "SetAppLangButton" Then
						If AppLngList(0) <> "" Then n = UBound(AppLngList) + 1
						If AppLngList(0) = "" Then n = UBound(AppLngList)
						ReDim Preserve AppLngList(n)
						AppLngList(n) = NewLngName
						DlgListBoxArray "AppLngList",AppLngList()
						DlgValue "AppLngList",n
					Else
						If UseLngList(0) <> "" Then  n = UBound(UseLngList) + 1
						If UseLngList(0) = "" Then  n = UBound(UseLngList)
						ReDim Preserve UseLngList(n)
						UseLngList(n) = NewLngName
						DlgListBoxArray "UseLngList",UseLngList()
						DlgValue "UseLngList",n
					End If
				End If
			End If

			If DlgItem$ = "EditAppLangButton" Or DlgItem$ = "EditUseLangButton" Then
				If DlgItem$ = "EditAppLangButton" Then
					LngName = DlgText("AppLngList")
					LngID = DlgValue("AppLngList")
				Else
					LngName = DlgText("UseLngList")
					LngID = DlgValue("UseLngList")
				End If
				If LngName <> "" Then
					n = -1
					LangArray = Split(TempArray(2),SubJoinStr)
					For i = 0 To UBound(LangArray)
						LangPairList = Split(LangArray(i),LngJoinStr)
						If LangPairList(0) = LngName Then
							LngCode = LangPairList(1)
							n = i
						End If
					Next i
					NewData = SetLang(LangArray,LngName,LngCode)
					If NewData <> "" Then
						If n >= 0 Then LangArray(n) = NewData
						If n < 0 Then
							n = UBound(LangArray) + 1
							ReDim Preserve LangArray(n)
							LangArray(n) = NewData
						End If
						TempArray(2) = Join(LangArray,SubJoinStr)

						LangPairList = Split(NewData,LngJoinStr)
						NewLngName = LangPairList(0)
						NewLngCode = LangPairList(1)
						If DlgItem$ = "EditAppLangButton" Then
							AppLngList(LngID) = NewLngName
							DlgListBoxArray "AppLngList",AppLngList()
							DlgValue "AppLngList",LngID
						Else
							UseLngList(LngID) = NewLngName
							DlgListBoxArray "UseLngList",UseLngList()
							DlgValue "UseLngList",LngID
						End If
					End If
				End If
			End If

			If DlgItem$ = "DelAppLangButton" Or DlgItem$ = "DelUseLangButton" Then
				If DlgItem$ = "DelAppLangButton" Then
					LngName = DlgText("AppLngList")
					LngID = DlgValue("AppLngList")
					i = UBound(AppLngList)
				Else
					LngName = DlgText("UseLngList")
					LngID = DlgValue("UseLngList")
					i = UBound(UseLngList)
				End If
				If LngName <> "" Then
					Msg = Replace(Msg24,"%s",LngName)
					If MsgBox(Msg,vbYesNo+vbInformation,Msg21) = vbYes Then
						LangArray = Split(TempArray(2),SubJoinStr)
						LangArray = DelArray(LngName,LangArray,LngJoinStr,0)
						TempArray(2) = Join(LangArray,SubJoinStr)

						If LngID > 0 And LngID = i Then LngID = LngID - 1
						If DlgItem$ = "DelAppLangButton" Then
							AppLngList = DelArray(LngName,AppLngList,"",0)
							DlgListBoxArray "AppLngList",AppLngList()
							DlgValue "AppLngList",LngID
						Else
							UseLngList = DelArray(LngName,UseLngList,"",0)
							DlgListBoxArray "UseLngList",UseLngList()
							DlgValue "UseLngList",LngID
						End If
					End If
				End If
			End If

			If DlgItem$ = "ImportButton" Then
				If PSL.SelectFile(Path,True,Msg26 & Msg44,Msg42) = True Then
					If CheckGet("",CheckList,CheckDataList,Path) = True Then
						'CheckGet("",CheckListBak,CheckDataListBak,Path)
						DlgListBoxArray "CheckList",CheckList()
						HeaderID = UBound(CheckList)
						DlgValue "CheckList",HeaderID
						TempArray = Split(CheckDataList(HeaderID),JoinStr)
						SetsArray = Split(TempArray(1),SubJoinStr)
						If InStr(SetsArray(11),rSubJoinStr) Then
							PreStr = Replace(SetsArray(11),rSubJoinStr,SubJoinStr)
						Else
							PreStr = SetsArray(11)
						End If
						DlgText "ExCrBox",SetsArray(0)
						DlgText "LnSpBox",SetsArray(1)
						DlgText "ChkBktBox",SetsArray(2)
						DlgText "KpPairBox",SetsArray(3)
						DlgValue "AsiaKeyBox",StrToInteger(SetsArray(4))
						DlgText "ChkEndBox",SetsArray(5)
						DlgText "NoTrnEndBox",SetsArray(6)
						DlgText "AutoTrnEndBox",SetsArray(7)
						DlgText "ShortBox",SetsArray(8)
						DlgText "ShortKeyBox",SetsArray(9)
						DlgText "KpShortKeyBox",SetsArray(10)
						DlgText "PreRepStrBox",PreStr
						DlgText "AutoWebFlagBox",SetsArray(12)
						getLngNameList(TempArray(2),AppLngList,UseLngList)
						DlgListBoxArray "AppLngList",AppLngList()
						DlgListBoxArray "UseLngList",UseLngList()
						DlgValue "AppLngList",0
						DlgValue "UseLngList",0
						MsgBox(Msg33,vbOkOnly+vbInformation,Msg30)
					Else
						MsgBox(Msg39 & cWriteLoc,vbOkOnly+vbInformation,Msg01)
					End If
				End If
			End If

			If DlgItem$ = "ExportButton" Then
				If CheckNullData("",CheckDataList,"4,11",1) = True Then
					If MsgBox(Msg19 & Msg12,vbYesNo+vbInformation,Msg11) = vbNo Then
						SetFunc% = True '防止按下按钮关闭对话框窗口
						Exit Function
					End If
				End If
				If PSL.SelectFile(Path,False,Msg26 & Msg44,Msg43) = True Then
					If CheckWrite(CheckDataList,Path,"All") = False Then
						MsgBox(Msg40 & Path,vbOkOnly+vbInformation,Msg01)
					Else
						MsgBox(Msg32,vbOkOnly+vbInformation,Msg30)
					End If
				End If
			End If

			If DlgItem$ <> "CancelButton" And DlgItem$ <> "OKButton" And DlgItem$ <> "ExportButton" Then
				Header = DlgText("CheckList")
				ExCr = DlgText("ExCrBox")
				LnSp = DlgText("LnSpBox")
				ChkBkt = DlgText("ChkBktBox")
				KpPair = DlgText("KpPairBox")
				AsiaKey = DlgValue("AsiaKeyBox")
				ChkEnd = DlgText("ChkEndBox")
				NoTrnEnd = DlgText("NoTrnEndBox")
				TrnEnd = DlgText("AutoTrnEndBox")
				Short = DlgText("ShortBox")
				Key = DlgText("ShortKeyBox")
				KpKey = DlgText("KpShortKeyBox")
				PreStr = DlgText("PreRepStrBox")
				WebFlag = DlgText("AutoWebFlagBox")

				If DlgItem$ = "ResetButton" Then TempArray = Split(CheckDataList(HeaderID),JoinStr)
				LangArray = Split(TempArray(2),SubJoinStr)
				For i = LBound(LangArray) To UBound(LangArray)
					If LangArray(i) <> "" Then
						LangPairList = Split(LangArray(i),LngJoinStr)
						LngName = LangPairList(0)
						LngCode = LangPairList(1)
						cStemp = False
						For n = LBound(UseLngList) To UBound(UseLngList)
							If UseLngList(n) = LngName Then
								LangArray(i) = LngName & LngJoinStr & LngCode & LngJoinStr & LngCode
								cStemp = True
								Exit For
							End If
						Next n
						If cStemp = False Then
							LangArray(i) = LngName & LngJoinStr & LngCode & LngJoinStr
						End If
					End If
				Next i
				LngPair = Join(LangArray,SubJoinStr)

				Temp = ExCr & LnSp & ChkBkt & KpPair & ChkEnd & NoTrnEnd & TrnEnd & _
						Short & Key & KpKey & PreStr & WebFlag & LngPair
				If Temp <> "" Then
					If InStr(PreStr,SubJoinStr) Then PreStr = Replace(PreStr,SubJoinStr,rSubJoinStr)
					Data = Header & JoinStr & ExCr & SubJoinStr & LnSp & SubJoinStr & ChkBkt & _
							SubJoinStr & KpPair & SubJoinStr & AsiaKey & SubJoinStr & ChkEnd & _
							SubJoinStr & NoTrnEnd & SubJoinStr & TrnEnd & SubJoinStr & Short & _
							SubJoinStr & Key & SubJoinStr & KpKey & SubJoinStr & PreStr & _
							SubJoinStr & WebFlag & JoinStr & LngPair
					CreateArray(Header,Data,CheckList,CheckDataList)
				End If

				If DlgText("AppLngList") = "" Then
					DlgEnable "AddLangButton",False
					DlgEnable "AddAllLangButton",False
					DlgEnable "EditAppLangButton",False
					DlgEnable "DelAppLangButton",False
					DlgText "AppLngText",Msg54 & "(0)"
				Else
					DlgEnable "AddLangButton",True
					DlgEnable "AddAllLangButton",True
					DlgEnable "EditAppLangButton",True
					DlgEnable "DelAppLangButton",True
					DlgText "AppLngText", Msg54 & "(" & UBound(AppLngList)+1 & ")"
				End If
				If DlgText("UseLngList") = "" Then
					DlgEnable "DelLangButton",False
					DlgEnable "DelAllLangButton",False
					DlgEnable "EditUseLangButton",False
					DlgEnable "DelUseLangButton",False
					DlgText "UseLngText",Msg55 & "(0)"
				Else
					DlgEnable "DelLangButton",True
					DlgEnable "DelAllLangButton",True
					DlgEnable "EditUseLangButton",True
					DlgEnable "DelUseLangButton",True
					DlgText "UseLngText",Msg55 & "(" & UBound(UseLngList)+1 & ")"
				End If

				If InStr(Join(DefaultCheckList,JoinStr),Header) Then
					DlgEnable "ChangButton",False
					DlgEnable "DelButton",False
				Else
					DlgEnable "ChangButton",True
					If UBound(CheckList) = 0 Then DlgEnable "DelButton",False
					If UBound(CheckList) <> 0 Then DlgEnable "DelButton",True
				End If
			End If

			If DlgItem$ = "TestButton" Then
				If CheckNullData("",CheckDataList,"4,11",1) = True Then
					If MsgBox(Msg19 & Msg12,vbYesNo+vbInformation,Msg11) = vbNo Then
						SetFunc% = True '防止按下按钮关闭对话框窗口
						Exit Function
					End If
				End If
				Call CheckTest(DlgValue("CheckList"),CheckList)
			End If

			If DlgItem$ = "HelpButton" Then Call CheckHelp("SetHelp")
		ElseIf DlgValue("Options") = 1 Then
			If DlgItem$ = "ExeBrowseButton" Then
				If PSL.SelectFile(Path,True,Msg61,Msg60) = True Then DlgText "CmdPathBox",Path
			End If
			If DlgItem$ = "ArgumentButton" Then
				ReDim TempArray(2)
				TempArray(0) = Msg63
				TempArray(1) = Msg64
				TempArray(2) = Msg65
				Temp = DlgText("ArgumentBox")
				x = ShowPopupMenu(TempArray)
				If x = 0 Then DlgText "ArgumentBox",Temp  & " " & """%1"""
				If x = 1 Then DlgText "ArgumentBox",Temp  & " " & """%2"""
				If x = 2 Then DlgText "ArgumentBox",Temp  & " " & """%3"""
			End If
			If DlgItem$ = "ResetButton" Then
				ReDim TempArray(1)
				TempArray(0) = Msg02
				TempArray(1) = Msg03
				i = ShowPopupMenu(TempArray,vbPopupUseRightButton)
				If i = 0 Then
					Temp = updateMainUrl & vbCrLf & updateMinorUrl
					DlgValue "UpdateSet",1
					DlgText "WebSiteBox",Temp
					Call getCMDPath(".rar",Path,Temp)
					DlgText "CmdPathBox",Path
					DlgText "ArgumentBox",Temp
					DlgText "UpdateCycleBox","7"
				ElseIf i = 1 Then
					DlgValue "UpdateSet",StrToInteger(cUpdateSetBak(0))
					DlgText "WebSiteBox",cUpdateSetBak(1)
					DlgText "CmdPathBox",cUpdateSetBak(2)
					DlgText "ArgumentBox",cUpdateSetBak(3)
					DlgText "UpdateCycleBox",cUpdateSetBak(4)
				End If
			End If
	    	If DlgItem$ = "CleanButton" Then
	    		DlgText "WebSiteBox",""
	    		DlgText "CmdPathBox",""
				DlgText "ArgumentBox",""
	    	End If
	    	If DlgItem$ = "CheckButton" Then
	    		DlgText "UpdateDateBox",Format(Date,"yyyy-MM-dd")
	    		UpdateUrl = DlgText("WebSiteBox")
	    		If Download(updateMethod,updateUrl,updateAsync,"3") = True Then
					cUpdateSet(5) = DlgText("UpdateDateBox")
					If DlgValue("cWriteType") = 0 Then cPath = CheckFilePath
					If DlgValue("cWriteType") = 1 Then cPath = CheckRegKey
					If CheckWrite(CheckDataList,cPath,"Update") = False Then
						MsgBox(Msg36 & cPath,vbOkOnly+vbInformation,Msg01)
						SetFunc% = True '防止按下按钮关闭对话框窗口
						Exit Function
					Else
						Exit All
					End If
				End If
			End If
			UpdateMode = DlgValue("UpdateSet")
			UpdateUrl = DlgText("WebSiteBox")
			CmdPath = DlgText("CmdPathBox")
			CmdArg = DlgText("ArgumentBox")
			UpdateCycle = DlgText("UpdateCycleBox")
			UpdateDate = DlgText("UpdateDateBox")
			If UpdateDate = Msg08 Then UpdateDate = ""
			Data = UpdateMode & JoinStr & updateUrl & JoinStr & CmdPath & JoinStr & CmdArg & _
					JoinStr & UpdateCycle & JoinStr & UpdateDate
			cUpdateSet = Split(Data,JoinStr)
			If DlgItem$ = "TestButton" Then Download(updateMethod,UpdateUrl,updateAsync,"4")
			If DlgItem$ = "HelpButton" Then Call UpdateHelp("SetHelp")
		End If

		If DlgItem$ = "OKButton" Then
			If CheckNullData("",CheckDataList,"4,11",1) = True Then
				If MsgBox(Msg19 & Msg12,vbYesNo+vbInformation,Msg11) = vbNo Then
					SetFunc% = True '防止按下按钮关闭对话框窗口
					Exit Function
				End If
			End If
			If DlgValue("cWriteType") = 0 Then cPath = CheckFilePath
			If DlgValue("cWriteType") = 1 Then cPath = CheckRegKey
			If CheckWrite(CheckDataList,cPath,"Sets") = False Then
				MsgBox(Msg36 & cPath,vbOkOnly+vbInformation,Msg01)
				SetFunc% = True '防止按下按钮关闭对话框窗口
				Exit Function
			Else
				CheckListBak = CheckList
				CheckDataListBak = CheckDataList
				cUpdateSetBak = cUpdateSet
			End If
		End If

		If DlgItem$ = "CancelButton" Then
			CheckList = CheckListBak
			CheckDataList = CheckDataListBak
			cUpdateSet = cUpdateSetBak
		End If
		If DlgItem$ <> "OKButton" And DlgItem$ <> "CancelButton" Then
			SetFunc% = True '防止按下按钮关闭对话框窗口
		End If
	Case 3 ' 文本框或者组合框文本被更改
		If DlgValue("Options") = 0 Then
			HeaderID = DlgValue("CheckList")
			Header = DlgText("CheckList")
			ExCr = DlgText("ExCrBox")
			LnSp = DlgText("LnSpBox")
			ChkBkt = DlgText("ChkBktBox")
			KpPair = DlgText("KpPairBox")
			AsiaKey = DlgValue("AsiaKeyBox")
			ChkEnd = DlgText("ChkEndBox")
			NoTrnEnd = DlgText("NoTrnEndBox")
			TrnEnd = DlgText("AutoTrnEndBox")
			Short = DlgText("ShortBox")
			Key = DlgText("ShortKeyBox")
			KpKey = DlgText("KpShortKeyBox")
			PreStr = DlgText("PreRepStrBox")
			WebFlag = DlgText("AutoWebFlagBox")

			TempArray = Split(CheckDataList(HeaderID),JoinStr)
			LngPair = TempArray(2)
			Temp = ExCr & LnSp & ChkBkt & KpPair & ChkEnd & NoTrnEnd & TrnEnd & _
					Short & Key & KpKey & PreStr & WebFlag & LngPair
			If Temp <> "" Then
				If InStr(PreStr,SubJoinStr) Then PreStr = Replace(PreStr,SubJoinStr,rSubJoinStr)
				Data = Header & JoinStr & ExCr & SubJoinStr & LnSp & SubJoinStr & ChkBkt & _
						SubJoinStr & KpPair & SubJoinStr & AsiaKey & SubJoinStr & ChkEnd & _
						SubJoinStr & NoTrnEnd & SubJoinStr & TrnEnd & SubJoinStr & Short & _
						SubJoinStr & Key & SubJoinStr & KpKey & SubJoinStr & PreStr & _
						SubJoinStr & WebFlag & JoinStr & LngPair
				CreateArray(Header,Data,CheckList,CheckDataList)
			End If
		ElseIf DlgValue("Options") = 1 Then
			UpdateMode = DlgValue("UpdateSet")
			UpdateUrl = DlgText("WebSiteBox")
			CmdPath = DlgText("CmdPathBox")
			CmdArg = DlgText("ArgumentBox")
			UpdateCycle = DlgText("UpdateCycleBox")
			UpdateDate = DlgText("UpdateDateBox")
			If UpdateDate = Msg08 Then UpdateDate = ""
			Data = UpdateMode & JoinStr & updateUrl & JoinStr & CmdPath & JoinStr & CmdArg & _
					JoinStr & UpdateCycle & JoinStr & UpdateDate
			cUpdateSet = Split(Data,JoinStr)
		End If
	End Select
End Function


'添加设置名称
Function AddSet(DataArr() As String) As String
	Dim NewHeader As String
	If OSLanguage = "0404" Then
		Msg01 = "穝糤"
		Msg04 = "叫块穝砞﹚嘿:"
		Msg06 = "岿粇"
		Msg07 = "眤⊿Τ块ヴず甧叫穝块"
		Msg08 = "赣嘿竒叫块ぃ嘿"
	Else
		Msg01 = "新建"
		Msg04 = "请输入新配置的名称:"
		Msg06 = "错误"
		Msg07 = "您没有输入任何内容！请重新输入。"
		Msg08 = "该名称已经存在！请输入一个不同的名称。"
	End If
	Begin Dialog UserDialog 310,77,Msg01 ' %GRID:10,7,1,1
		Text 10,7,290,14,Msg04
		TextBox 10,21,290,21,.TextBox
		OKButton 70,49,80,21,.OKButton
		CancelButton 170,49,80,21,.CancelButton
	End Dialog
    Dim dlg As UserDialog
    DataInPutDlg:
    If Dialog(dlg) = 0 Then Exit Function

	Dim Data As String
	NewHeader = Trim(dlg.TextBox)
	Data = ""
	If IsArray(DataArr) Then
		For i = LBound(DataArr) To UBound(DataArr)
			If NewHeader = DataArr(i) Then
				Data = DataArr(i)
				Exit For
			End If
		Next i
	End If

	If NewHeader = "" Then
		MsgBox Msg07,vbOkOnly+vbInformation,Msg06
		GoTo DataInPutDlg
	ElseIf NewHeader = Data Then
		MsgBox Msg08,vbOkOnly+vbInformation,Msg06
		GoTo DataInPutDlg
	End If
	AddSet = NewHeader
End Function


'编辑设置名称
Function EditSet(DataArr() As String,Header As String) As String
	Dim tempHeader As String,NewHeader As String
	If OSLanguage = "0404" Then
		Msg01 = "跑"
		Msg04 = "侣嘿:"
		Msg05 = "穝嘿:"
		Msg06 = "岿粇"
		Msg07 = "眤⊿Τ块ヴず甧叫穝块"
		Msg08 = "赣嘿竒叫块ぃ嘿"
	Else
		Msg01 = "更改"
		Msg04 = "旧名称:"
		Msg05 = "新名称:"
		Msg06 = "错误"
		Msg07 = "您没有输入任何内容！请重新输入。"
		Msg08 = "该名称已经存在！请输入一个不同的名称。"
	End If
	tempHeader = Header
	If InStr(Header,"&") Then
		tempHeader = Replace(Header,"&","&&")
	End If

	Begin Dialog UserDialog 310,126,Msg01  ' %GRID:10,7,1,1
		GroupBox 10,17,290,28,"",.GroupBox1
		Text 10,7,290,14,Msg04,.Text1
		Text 20,28,270,14,tempHeader,.oldNameText
		Text 10,56,290,14,Msg05,.newNameText
		TextBox 10,70,290,21,.TextBox
		OKButton 70,98,80,21,.OKButton
		CancelButton 170,98,80,21,.CancelButton
	End Dialog
    Dim dlg As UserDialog
    DataInPutDlg:
    dlg.TextBox = Header
    If Dialog(dlg) = 0 Then Exit Function

	Dim Data As String
	NewHeader = Trim(dlg.TextBox)
	Data = ""
	If IsArray(DataArr) Then
		For i = LBound(DataArr) To UBound(DataArr)
			If NewHeader = DataArr(i) Then
				Data = DataArr(i)
				Exit For
			End If
		Next i
	End If

	If NewHeader = "" Then
		MsgBox Msg07,vbOkOnly+vbInformation,Msg06
		GoTo DataInPutDlg
	ElseIf NewHeader = Data Then
		MsgBox Msg08,vbOkOnly+vbInformation,Msg06
		GoTo DataInPutDlg
	End If
	EditSet = NewHeader
End Function


'添加或编辑语言对
Function SetLang(DataArr() As String,LangName As String,LangCode As String) As String
	Dim tempHeader As String,NewLangName As String,NewLangCode As String
	If OSLanguage = "0404" Then
		Msg01 = "穝糤"
		Msg02 = "絪胯"
		Msg04 = "粂ē嘿:"
		Msg05 = "Passolo 粂ē絏:"
		Msg10 = "岿粇"
		Msg11 = "眤⊿Τ块ヴず甧叫穝块"
		Msg12 = "粂ē嘿㎝ Passolo 粂ē絏いぶΤ兜ヘ叫浪琩块"
		Msg13 = "赣粂ē嘿竒叫穝块"
		Msg14 = "赣 Passolo 粂ē絏竒叫穝块"
	Else
		Msg01 = "添加"
		Msg02 = "编辑"
		Msg04 = "语言名称:"
		Msg05 = "Passolo 语言代码:"
		Msg10 = "错误"
		Msg11 = "您没有输入任何内容！请重新输入。"
		Msg12 = "语言名称和 Passolo 语言代码中至少有一个项目为空！请检查并输入。"
		Msg13 = "该语言名称已经存在！请重新输入。"
		Msg14 = "该 Passolo 语言代码已经存在！请重新输入。"
	End If
	If LangName = "" Then Msg = Msg01
	If LangName <> "" Then Msg = Msg02
	Begin Dialog UserDialog 390,168,Msg  ' %GRID:10,7,1,1
		Text 10,14,370,14,Msg04
		TextBox 10,35,370,21,.LangName
		Text 10,77,370,14,Msg05
		TextBox 10,98,370,21,.LangCode
		OKButton 90,140,80,21,.OKButton
		CancelButton 220,140,80,21,.CancelButton
	End Dialog
    Dim dlg As UserDialog
    If LangName <> "" Then dlg.LangName = LangName
    If LangCode <> "" Then dlg.LangCode = LangCode
    DataInPutDlg:
    If Dialog(dlg) = 0 Then Exit Function

	NewLangName = Trim(dlg.LangName)
	NewLangCode = Trim(dlg.LangCode)
	OldLangName = ""
	OldLangCode = ""
	If NewLangName <> "" And NewLangCode <> "" Then
		For i = LBound(DataArr) To UBound(DataArr)
			LangPairList = Split(DataArr(i),LngJoinStr)
			If NewLangName = LangPairList(0) Then OldLangName = LangPairList(0)
			If NewLangCode = LangPairList(1) Then OldLangCode = LangPairList(1)
			If OldLangName <> "" And OldLangCode <> "" Then Exit For
		Next i
	End If

	If NewLangName & NewLangCode = "" Then
		MsgBox Msg11,vbOkOnly+vbInformation,Msg10
		GoTo DataInPutDlg
	End If
	If NewLangName = "" Or NewLangCode = "" Then
		MsgBox Msg12,vbOkOnly+vbInformation,Msg10
		GoTo DataInPutDlg
	End If
	If OldLangName <> "" And NewLangName <> LangName Then
		MsgBox Msg13,vbOkOnly+vbInformation,Msg10
		GoTo DataInPutDlg
	End If
	If OldLangCode <> "" And NewLangCode <> LangCode Then
		MsgBox Msg14,vbOkOnly+vbInformation,Msg10
		GoTo DataInPutDlg
	End If
	SetLang = NewLangName & LngJoinStr & NewLangCode & LngJoinStr
End Function


'更改配置优先级
Function SetLevel(HeaderList() As String,DataList() As String) As Boolean
	SetLevel = False
	If OSLanguage = "0404" Then
		Msg01 = "砞﹚纔"
		Msg02 = "砞﹚纔ノ膀砞﹚続ノ粂ē笆匡砞﹚"
		Msg03 = "矗ボ:" & vbCrLf & _
				"- Τ砞﹚続ノ粂ē惠璶砞﹚ㄤ纔" & vbCrLf & _
				"- 続ノ粂ē砞﹚い玡砞﹚砆纔匡ㄏノ"
		Msg04 = "簿(&U)"
		Msg05 = "簿(&D)"
		Msg06 = "砞(&R)"
	Else
		Msg01 = "配置优先级"
		Msg02 = "配置优先级用于基于配置的适用语言的自动选择配置功能。"
		Msg03 = "提示:" & vbCrLf & _
				"- 有多个配置包含了相同的适用语言时，需要设置其优先级。" & vbCrLf & _
				"- 在相同适用语言的配置中，前面的配置被优先选择使用。"
		Msg04 = "上移(&U)"
		Msg05 = "下移(&D)"
		Msg06 = "重置(&R)"
	End If
	tempCheckList = HeaderList
	tempCheckDataList = DataList
	Begin Dialog UserDialog 480,294,Msg01,.SetLevelFunc ' %GRID:10,7,1,1
		Text 10,7,460,14,Msg02
		Text 10,21,460,42,Msg03
		ListBox 10,70,360,189,tempCheckList(),.tempCheckList
		PushButton 380,77,90,21,Msg04,.UpButton
		PushButton 380,105,90,21,Msg05,.DownButton
		PushButton 380,133,90,21,Msg06,.ResetButton
		OKButton 130,266,80,21,.OKButton
		CancelButton 280,266,80,21,.CancelButton
	End Dialog
    Dim dlg As UserDialog
    If Dialog(dlg) = 0 Then Exit Function
   	SetLevel = True
End Function


'更改配置优先级对话框函数
Private Function SetLevelFunc%(DlgItem$, Action%, SuppValue&)
	Dim i As Integer,ID As Integer,Temp As String
	Select Case Action%
	Case 1 ' 对话框窗口初始化
		If UBound(tempCheckList) = 0 Then
			DlgEnable "UpButton",False
			DlgEnable "DownButton",False
		Else
			ID = DlgValue("tempCheckList")
			If ID = 0 Then
				DlgEnable "UpButton",False
				DlgEnable "DownButton",True
			ElseIf ID = UBound(tempCheckList) Then
				DlgEnable "UpButton",True
				DlgEnable "DownButton",False
			Else
				DlgEnable "UpButton",True
				DlgEnable "DownButton",True
			End If
		End If
		DlgEnable "ResetButton",False
	Case 2 ' 数值更改或者按下了按钮
		If DlgItem$ = "tempCheckList" Then
			If UBound(tempCheckList) = 0 Then
				DlgEnable "UpButton",False
				DlgEnable "DownButton",False
			Else
				ID = DlgValue("tempCheckList")
				If ID = 0 Then
					DlgEnable "UpButton",False
					DlgEnable "DownButton",True
				ElseIf ID = UBound(tempCheckList) Then
					DlgEnable "UpButton",True
					DlgEnable "DownButton",False
				Else
					DlgEnable "UpButton",True
					DlgEnable "DownButton",True
				End If
			End If
		End If

		If DlgItem$ = "UpButton" Then
			ID = DlgValue("tempCheckList")
			If ID <> 0 Then
				Temp = tempCheckList(ID)
				tempCheckList(ID) = tempCheckList(ID-1)
				tempCheckList(ID-1) = Temp
				Temp = tempCheckDataList(ID)
				tempCheckDataList(ID) = tempCheckDataList(ID-1)
				tempCheckDataList(ID-1) = Temp
				ID = ID - 1
			End If
		End If

		If DlgItem$ = "DownButton" Then
			ID = DlgValue("tempCheckList")
			If ID <> UBound(tempCheckList) Then
				Temp = tempCheckList(ID)
				tempCheckList(ID) = tempCheckList(ID+1)
				tempCheckList(ID+1) = Temp
				Temp = tempCheckDataList(ID)
				tempCheckDataList(ID) = tempCheckDataList(ID+1)
				tempCheckDataList(ID+1) = Temp
				ID = ID + 1
			End If
		End If

		If DlgItem$ = "ResetButton" Then
			ID = DlgValue("tempCheckList")
			tempCheckList = CheckList
			tempCheckDataList = CheckDataList
		End If

		If DlgItem$ = "OKButton" Then
			CheckList = tempCheckList
			CheckDataList = tempCheckDataList
		End If

		If DlgItem$ <> "CancelButton" And DlgItem$ <> "OKButton" Then
			DlgListBoxArray "tempCheckList",tempCheckList()
			DlgValue "tempCheckList",ID
			ID = DlgValue("tempCheckList")
			If ID = 0 Then
				DlgEnable "UpButton",False
				DlgEnable "DownButton",True
			ElseIf ID = UBound(tempCheckList) Then
				DlgEnable "UpButton",True
				DlgEnable "DownButton",False
			Else
				DlgEnable "UpButton",True
				DlgEnable "DownButton",True
			End If
			If DlgItem$ <> "tempCheckList" Then
				If Join(tempCheckList) = Join(CheckList) Then
					DlgEnable "ResetButton",False
				Else
					DlgEnable "ResetButton",True
				End If
			End If
			SetLevelFunc = True '防止按下按钮关闭对话框窗口
		End If
	End Select
End Function


'获取字串检查设置
Function CheckGet(SelSet As String,HeaderList() As String,DataList() As String,Path As String) As Boolean
	Dim i As Integer,Header As String,HeaderIDs As String,HeaderIDArr() As String,Temp As String
	Dim LangPairList() As String,TempArray() As String
	CheckGet = False
	NewVersion = ToUpdateCheckVersion
	LangPairList = LangCodeList("check",OSLanguage,1,107)

	If Path = CheckRegKey Then GoTo GetFromRegistry
	If Path = "" Then Path = CheckFilePath
	If Dir(Path) = "" Then GoTo GetFromRegistry
	If Path = CheckFilePath Then On Error GoTo GetFromRegistry
	Open Path For Input As #1
	While Not EOF(1)
		Line Input #1,L$
		If Left(Trim(L$),1) = "[" And Right(Trim(L$),1) = "]" Then
			Header = Mid(Trim(L$),2,Len(Trim(L$))-2)
		End If
		If L$ <> "" And Header <> "" Then
			setPreStr = ""
			setAppStr = ""
			Site = ""
			If InStr(L$,"=") Then setPreStr = Trim(Left(L$,InStr(L$,"=")-1))
			If InStr(L$,"=") Then setAppStr = LTrim(Mid(L$,InStr(L$,"=")+1))
			'获取 Option 项和值
			If setPreStr = "Version" Then OldVersion = setAppStr
			If Header = "Option" And SelSet = "" And setPreStr <> "" Then
				If setPreStr = "AutoMacroSet" Then AutoMacroSet = setAppStr
				If setPreStr = "CheckMacroSet" Then CheckMacroSet = setAppStr
				If setPreStr = "AutoMacroCheck" Then AutoMacroChk = setAppStr
				If setPreStr = "AutoSelection" Then AutoSele = setAppStr
				If setPreStr = "SelectedCheck" Then miVo = setAppStr
				If setPreStr = "CheckAllType" Then mAllType = setAppStr
				If setPreStr = "CheckMenu" Then mMenu = setAppStr
				If setPreStr = "CheckDialog" Then mDialog = setAppStr
				If setPreStr = "CheckString" Then mString = setAppStr
				If setPreStr = "CheckAcceleratorTable" Then mAccTable = setAppStr
				If setPreStr = "CheckVersion" Then mVer = setAppStr
				If setPreStr = "CheckOther" Then mOther = setAppStr
				If setPreStr = "CheckSeletedOnly" Then mSelOnly = setAppStr
				If setPreStr = "CheckAllCont" Then mAllCont = setAppStr
				If setPreStr = "CheckAccKey" Then mAccKey = setAppStr
				If setPreStr = "CheckEndChar" Then mEndChar = setAppStr
				If setPreStr = "CheckAcceler" Then mAcceler = setAppStr
				If setPreStr = "IgnoreVerTag" Then VerTag = setAppStr
				If setPreStr = "IgnoreSetTag" Then SetTag = setAppStr
				If setPreStr = "IgnoreDateTag" Then DateTag = setAppStr
				If setPreStr = "IgnoreStateTag" Then StateTag = setAppStr
				If setPreStr = "TgnoreAllTag" Then AllTag = setAppStr
				If setPreStr = "NoCheckTag" Then NoChkTag = setAppStr
				If setPreStr = "NoChangeTrnState" Then NoChgSta = setAppStr
				If setPreStr = "AutoRepString" Then repStr = setAppStr
				If setPreStr = "KeepSettings" Then KeepSet = setAppStr
				If AutoMacroSet = "Default" Then AutoMacroSet = DefaultCheckList(0)
				If CheckMacroSet = "Default" Then CheckMacroSet = DefaultCheckList(0)
			End If
			'获取 Update 项和值
			If Header = "Update" And SelSet = "" And setPreStr <> "" Then
				If setPreStr = "UpdateMode" Then UpdateMode = setAppStr
				If InStr(setPreStr,"Site_") Then Site = setAppStr
				If Site <> "" Then
					If UpdateSite <> "" Then UpdateSite = UpdateSite & vbCrLf & Site
					If UpdateSite = "" Then UpdateSite = Site
				End If
				If setPreStr = "Path" Then CmdPath = setAppStr
				If setPreStr = "Argument" Then CmdArg = setAppStr
				If setPreStr = "UpdateCycle" Then UpdateCycle = setAppStr
				If setPreStr = "UpdateDate" Then UpdateDate = setAppStr
			End If
			'获取 Option 项外的全部项和值
			If Header <> "Option" And Header <> "Update" And setPreStr <> "" Then
				If setPreStr = "ExcludeChar" Then ExCr = setAppStr
				If setPreStr = "LineSplitChar" Then LnSp = setAppStr
				If setPreStr = "CheckBracket" Then ChkBkt = setAppStr
				If setPreStr = "KeepCharPair" Then KpPair = setAppStr
				If setPreStr = "ShowAsiaKey" Then AsiaKey = setAppStr
				If setPreStr = "CheckEndChar" Then ChkEnd = setAppStr
				If setPreStr = "NoTrnEndChar" Then NoTrnEnd = setAppStr
				If setPreStr = "AutoTrnEndChar" Then TrnEnd = setAppStr
				If setPreStr = "CheckShortChar" Then Short = setAppStr
				If setPreStr = "CheckShortKey" Then Key = setAppStr
				If setPreStr = "KeepShortKey" Then KpKey = setAppStr
				If setPreStr = "PreRepString" Then PreStr = setAppStr
				If setPreStr = "AutoRepString" Then AppStr = setAppStr
				If setPreStr = "ApplyLangList" Then LngPair = setAppStr
				If AsiaKey = "" Then AsiaKey = "0"
			End If
		End If
		If (L$ = "" Or EOF(1)) And Header = "Option" Then
			Data = AutoMacroSet & JoinStr & CheckMacroSet & JoinStr & AutoMacroChk & _
					JoinStr & AutoSele & JoinStr & miVo & JoinStr & mAllType & _
					JoinStr & mMenu & JoinStr & mDialog & JoinStr & mString & _
					JoinStr & mAccTable & JoinStr & mVer & JoinStr & mOther & _
					JoinStr & mSelOnly & JoinStr & mAllCont & JoinStr & mAccKey & _
					JoinStr & mEndChar & JoinStr & mAcceler & JoinStr & VerTag & _
					JoinStr & SetTag & JoinStr & DateTag & JoinStr & StateTag & _
					JoinStr & AllTag & JoinStr & NoChkTag & JoinStr & NoChgSta & _
					JoinStr & RepStr & JoinStr & KeepSet
			cSelected = Split(Data,JoinStr)
		End If
		If (L$ = "" Or EOF(1)) And Header = "Update" Then
			Data = UpdateMode & JoinStr & UpdateSite & JoinStr & CmdPath & JoinStr & CmdArg & _
					JoinStr & UpdateCycle & JoinStr & UpdateDate
			cUpdateSet = Split(Data,JoinStr)
		End If
		If (L$ = "" Or EOF(1)) And Header <> "" And Header <> "Option" And Header <> "Update" Then
			Temp = ExCr & LnSp & ChkBkt & KpPair & ChkEnd & NoTrnEnd & TrnEnd & _
					Short & Key & KpKey & PreStr & AppStr & LngPair
			If Temp <> "" Then
				If LngPair <> "" Then
					If Path <> CheckFilePath Then LngPair = LangNameUpdate(Header,LngPair)
					TempArray = Split(LngPair,SubJoinStr)
					LngPair = Join(MergeLngList(LangPairList,TempArray,"check"),SubJoinStr)
				Else
					LngPair = Join(LangCodeList(Header,OSLanguage,1,107),SubJoinStr)
				End If
				Data = Header & JoinStr & ExCr & SubJoinStr & LnSp & SubJoinStr & ChkBkt & _
						SubJoinStr & KpPair & SubJoinStr & AsiaKey & SubJoinStr & ChkEnd & _
						SubJoinStr & NoTrnEnd & SubJoinStr & TrnEnd & SubJoinStr & Short & _
						SubJoinStr & Key & SubJoinStr & KpKey & SubJoinStr & PreStr & _
						SubJoinStr & AppStr & JoinStr & LngPair
				If Path <> CheckFilePath Then Data = CheckSetUpdate(Header,Data)
				'更新旧版的默认配置值
				If InStr(Join(DefaultCheckList,JoinStr),Header) Then
					If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 Then
						Data = CheckDataUpdate(Header,Data)
					End If
				End If
				'保存数据到数组中
				CreateArray(Header,Data,HeaderList,DataList)
				CheckGet = True
			End If
			'数据初始化
			ExCr = ""
			LnSp = ""
			ChkBkt = ""
			KpPair = ""
			AsiaKey = ""
			ChkEnd = ""
			NoTrnEnd = ""
			TrnEnd = ""
			Short = ""
			Key = ""
			KpKey = ""
			PreStr = ""
			AppStr = ""
			LngPair = ""
		End If
	Wend
	Close #1
	If Path = CheckFilePath Then On Error GoTo 0
	'保存更新后的数据到文件
	If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 And Path = CheckFilePath Then
		If Dir(CheckFilePath) <> "" Then CheckWrite(DataList,CheckFilePath,"All")
	End If
	If cWriteLoc = "" Then cWriteLoc = CheckFilePath
	Exit Function

	GetFromRegistry:
	'获取 Option 项和值
	OldVersion = GetSetting("AccessKey","Option","Version","")
	If SelSet = "" Then
		AutoMacroSet = GetSetting("AccessKey","Option","AutoMacroSet","")
		CheckMacroSet = GetSetting("AccessKey","Option","CheckMacroSet","")
		AutoMacroChk = GetSetting("AccessKey","Option","AutoMacroCheck",0)
		AutoSele = GetSetting("AccessKey","Option","AutoSelection",1)
		miVo = GetSetting("AccessKey","Option","SelectedCheck",0)
		mAllType = GetSetting("AccessKey","Option","CheckAllType",0)
		mMenu = GetSetting("AccessKey","Option","CheckMenu",0)
		mDialog = GetSetting("AccessKey","Option","CheckDialog",0)
		mString = GetSetting("AccessKey","Option","CheckString",0)
		mAccTable = GetSetting("AccessKey","Option","CheckAcceleratorTable",0)
		mVer = GetSetting("AccessKey","Option","CheckVersion",0)
		mOther = GetSetting("AccessKey","Option","CheckOther",0)
		mSelOnly = GetSetting("AccessKey","Option","CheckSeletedOnly",1)
		mAllCont = GetSetting("AccessKey","Option","CheckAllCont",0)
		mAccKey = GetSetting("AccessKey","Option","CheckAccKey",0)
		mEndChar = GetSetting("AccessKey","Option","CheckEndChar",0)
		mAcceler = GetSetting("AccessKey","Option","CheckAcceler",0)
		VerTag = GetSetting("AccessKey","Option","IgnoreVerTag",0)
		SetTag = GetSetting("AccessKey","Option","IgnoreSetTag",0)
		DateTag = GetSetting("AccessKey","Option","IgnoreDateTag",0)
		StateTag = GetSetting("AccessKey","Option","IgnoreStateTag",0)
		AllTag = GetSetting("AccessKey","Option","TgnoreAllTag",1)
		NoChkTag = GetSetting("AccessKey","Option","NoCheckTag",1)
		NoChgSta = GetSetting("AccessKey","Option","NoChangeTrnState",0)
		repStr = GetSetting("AccessKey","Option","AutoRepString",0)
		KeepSet = GetSetting("AccessKey","Option","KeepSetting",1)
		If AutoMacroSet = "Default" Then AutoMacroSet = DefaultCheckList(0)
		If CheckMacroSet = "Default" Then CheckMacroSet = DefaultCheckList(0)
		Data = AutoMacroSet & JoinStr & CheckMacroSet & JoinStr & AutoMacroChk & _
				JoinStr & AutoSele & JoinStr & miVo & JoinStr & mAllType & _
				JoinStr & mMenu & JoinStr & mDialog & JoinStr & mString & _
				JoinStr & mAccTable & JoinStr & mVer & JoinStr & mOther & _
				JoinStr & mSelOnly & JoinStr & mAllCont & JoinStr & mAccKey & _
				JoinStr & mEndChar & JoinStr & mAcceler & JoinStr & VerTag & _
				JoinStr & SetTag & JoinStr & DateTag & JoinStr & StateTag & _
				JoinStr & AllTag & JoinStr & NoChkTag & JoinStr & NoChgSta & _
				JoinStr & repStr & JoinStr & KeepSet
		cSelected = Split(Data,JoinStr)
		'获取 Update 项和值
		UpdateMode = GetSetting("AccessKey","Update","UpdateMode",1)
		Count = GetSetting("AccessKey","Update","Count",0)
		For i = 0 To Count
			Site = GetSetting("AccessKey","Update",CStr(i),"")
			If Site <> "" Then
				If i > 0 Then UpdateSite = UpdateSite & vbCrLf & Site
				If i = 0 Then UpdateSite = Site
			End If
		Next i
		CmdPath = GetSetting("AccessKey","Update","Path","")
		CmdArg = GetSetting("AccessKey","Update","Argument","")
		UpdateCycle = GetSetting("AccessKey","Update","UpdateCycle",7)
		UpdateDate = GetSetting("AccessKey","Update","UpdateDate","")
		Data = UpdateMode & JoinStr & UpdateSite & JoinStr & CmdPath & JoinStr & CmdArg & _
					JoinStr & UpdateCycle & JoinStr & UpdateDate
		cUpdateSet = Split(Data,JoinStr)
	End If

	'获取 Option 外的项和值
	HeaderIDs = GetSetting("AccessKey","Option","Headers","")
	If HeaderIDs <> "" Then
		HeaderIDArr = Split(HeaderIDs,";",-1)
		For i = LBound(HeaderIDArr) To UBound(HeaderIDArr)
			HeaderID = HeaderIDArr(i)
			If HeaderID <> "" Then
				'转存旧版的每个项和值
				Header = GetSetting("AccessKey",HeaderID,"Name","")
				If Header = "" Then Header = HeaderID
				If Header <> "" Then
					ExCr = GetSetting("AccessKey",HeaderID,"ExcludeChar","")
					LnSp = GetSetting("AccessKey",HeaderID,"LineSplitChar","")
					ChkBkt = GetSetting("AccessKey",HeaderID,"CheckBracket","")
					KpPair = GetSetting("AccessKey",HeaderID,"KeepCharPair","")
					AsiaKey = GetSetting("AccessKey",HeaderID,"ShowAsiaKey",0)
					ChkEnd = GetSetting("AccessKey",HeaderID,"CheckEndChar","")
					NoTrnEnd = GetSetting("AccessKey",HeaderID,"NoTrnEndChar","")
					TrnEnd = GetSetting("AccessKey",HeaderID,"AutoTrnEndChar","")
					Short = GetSetting("AccessKey",HeaderID,"CheckShortChar","")
					Key = GetSetting("AccessKey",HeaderID,"CheckShortKey","")
					KpKey = GetSetting("AccessKey",HeaderID,"KeepShortKey","")
					PreStr = GetSetting("AccessKey",HeaderID,"PreRepString","")
					AppStr = GetSetting("AccessKey",HeaderID,"AutoRepString","")
					LngPair = GetSetting("AccessKey",HeaderID,"ApplyLangList","")
					Temp = ExCr & LnSp & ChkBkt & KpPair & ChkEnd & NoTrnEnd & TrnEnd & _
							Short & Key & KpKey & PreStr & AppStr & LngPair
					If Temp <> "" Then
						If LngPair <> "" Then
							TempArray = Split(LngPair,SubJoinStr)
							LngPair = Join(MergeLngList(LangPairList,TempArray,"check"),SubJoinStr)
						Else
							LngPair = Join(LangCodeList(Header,OSLanguage,1,107),SubJoinStr)
						End If
						Data = Header & JoinStr & ExCr & SubJoinStr & LnSp & SubJoinStr & ChkBkt & _
								SubJoinStr & KpPair & SubJoinStr & AsiaKey & SubJoinStr & ChkEnd & _
								SubJoinStr & NoTrnEnd & SubJoinStr & TrnEnd & SubJoinStr & Short & _
								SubJoinStr & Key & SubJoinStr & KpKey & SubJoinStr & PreStr & _
								SubJoinStr & AppStr & JoinStr & LngPair
						'更新旧版的默认配置值
						If InStr(Join(DefaultCheckList,JoinStr),Header) Then
							If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 Then
								Data = CheckDataUpdate(Header,Data)
							End If
						End If
						'保存数据到数组中
						CreateArray(Header,Data,HeaderList,DataList)
						CheckGet = True
					End If
					'删除旧版配置值
					On Error Resume Next
					If Header = HeaderID Then DeleteSetting("AccessKey",Header)
					On Error GoTo 0
				End If
			End If
		Next i
	End If
	'保存更新后的数据到注册表
	If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 Then
		If HeaderIDs <> "" Then CheckWrite(DataList,CheckRegKey,"Sets")
	End If
	If cWriteLoc = "" Then cWriteLoc = CheckRegKey
End Function


'写入字串检查设置
Function CheckWrite(DataList() As String,Path As String,WriteType As String) As Boolean
	Dim i As Integer,Header As String,HeaderIDs As String,TempPath As String
	Dim SetsArray() As String,HeaderIDArr() As String
	CheckWrite = False
	KeepSet = cSelected(UBound(cSelected))

	'写入文件
	If Path <> "" And Path <> CheckRegKey Then
   		On Error Resume Next
   		TempPath = Left(Path,InStrRev(Path,"\"))
   		If Dir(TempPath & "*.*") = "" Then MkDir TempPath
		If Dir(Path) <> "" Then SetAttr Path,vbNormal
		On Error GoTo 0
		On Error GoTo ExitFunction
		Open Path For Output As #2
			Print #2,";------------------------------------------------------------"
			Print #2,";Settings for PSLCheckAccessKeys.bas and PslAutoAccessKey.bas"
			Print #2,";------------------------------------------------------------"
			Print #2,""
			Print #2,"[Option]"
			Print #2,"Version = " & Version
			If KeepSet = "1" Then
				Print #2,"AutoMacroSet = " & cSelected(0)
				Print #2,"CheckMacroSet = " & cSelected(1)
				Print #2,"AutoMacroCheck = " & cSelected(2)
				Print #2,"AutoSelection = " & cSelected(3)
				Print #2,"SelectedCheck = " & cSelected(4)
				Print #2,"CheckAllType = " & cSelected(5)
				Print #2,"CheckMenu = " & cSelected(6)
				Print #2,"CheckDialog = " & cSelected(7)
				Print #2,"CheckString = " & cSelected(8)
				Print #2,"CheckAcceleratorTable = " & cSelected(9)
				Print #2,"CheckVersion = " & cSelected(10)
				Print #2,"CheckOther = " & cSelected(11)
				Print #2,"CheckSeletedOnly = " & cSelected(12)
				Print #2,"CheckAllCont = " & cSelected(13)
				Print #2,"CheckAccKey = " & cSelected(14)
				Print #2,"CheckEndChar = " & cSelected(15)
				Print #2,"CheckAcceler = " & cSelected(16)
				Print #2,"IgnoreVerTag = " & cSelected(17)
				Print #2,"IgnoreSetTag = " & cSelected(18)
				Print #2,"IgnoreDateTag = " & cSelected(19)
				Print #2,"IgnoreStateTag = " & cSelected(20)
				Print #2,"TgnoreAllTag = " & cSelected(21)
				Print #2,"NoCheckTag = " & cSelected(22)
				Print #2,"NoChangeTrnState = " & cSelected(23)
				Print #2,"AutoRepString = " & cSelected(24)
				Print #2,"KeepSettings = " & cSelected(25)
			End If
			Print #2,""
			If Join(cUpdateSet) <> "" Then
				UpdateSiteList = Split(cUpdateSet(1),vbCrLf,-1)
				Print #2,"[Update]"
				Print #2,"UpdateMode = " & cUpdateSet(0)
				For i = LBound(UpdateSiteList) To UBound(UpdateSiteList)
					Print #2,"Site_" & CStr(i) & " = " & UpdateSiteList(i)
				Next i
				Print #2,"Path = " & cUpdateSet(2)
				Print #2,"Argument = " & cUpdateSet(3)
				Print #2,"UpdateCycle = " & cUpdateSet(4)
				Print #2,"UpdateDate = " & cUpdateSet(5)
				Print #2,""
			End If
			For i = LBound(DataList) To UBound(DataList)
				TempArray = Split(DataList(i),JoinStr)
				SetsArray = Split(TempArray(1),SubJoinStr)
				Print #2,"[" & TempArray(0) & "]"
				Print #2,"ExcludeChar = " & SetsArray(0)
				Print #2,"LineSplitChar = " & SetsArray(1)
				Print #2,"CheckBracket = " & SetsArray(2)
				Print #2,"KeepCharPair = " & SetsArray(3)
				Print #2,"ShowAsiaKey = " & SetsArray(4)
				Print #2,"CheckEndChar = " & SetsArray(5)
				Print #2,"NoTrnEndChar = " & SetsArray(6)
				Print #2,"AutoTrnEndChar = " & SetsArray(7)
				Print #2,"CheckShortChar = " & SetsArray(8)
				Print #2,"CheckShortKey = " & SetsArray(9)
				Print #2,"KeepShortKey = " & SetsArray(10)
				Print #2,"PreRepString = " & SetsArray(11)
				Print #2,"AutoRepString = " & SetsArray(12)
				Print #2,"ApplyLangList = " & getLngPair(TempArray(2),"check")
				If i <> UBound(DataList) Then Print #2,""
			Next i
		Close #2
		On Error GoTo 0
		CheckWrite = True
		If Path = CheckFilePath Then cWriteLoc = CheckFilePath
		If Path = CheckFilePath Then GoTo RemoveRegKey

	'写入注册表
	ElseIf Path = CheckRegKey Then
		On Error GoTo ExitFunction
		SaveSetting("AccessKey","Option","Version",Version)
		If WriteType = "Main" Or WriteType = "All" Then
			If KeepSet = "1" Then
				SaveSetting("AccessKey","Option","AutoMacroSet",cSelected(0))
				SaveSetting("AccessKey","Option","CheckMacroSet",cSelected(1))
				SaveSetting("AccessKey","Option","AutoMacroCheck",cSelected(2))
				SaveSetting("AccessKey","Option","AutoSelection",cSelected(3))
				SaveSetting("AccessKey","Option","SelectedCheck",cSelected(4))
				SaveSetting("AccessKey","Option","CheckAllType",cSelected(5))
				SaveSetting("AccessKey","Option","CheckMenu",cSelected(6))
				SaveSetting("AccessKey","Option","CheckDialog",cSelected(7))
				SaveSetting("AccessKey","Option","CheckString",cSelected(8))
				SaveSetting("AccessKey","Option","CheckAcceleratorTable",cSelected(9))
				SaveSetting("AccessKey","Option","CheckVersion",cSelected(10))
				SaveSetting("AccessKey","Option","CheckOther",cSelected(11))
				SaveSetting("AccessKey","Option","CheckSeletedOnly",cSelected(12))
				SaveSetting("AccessKey","Option","CheckAllCont",cSelected(13))
				SaveSetting("AccessKey","Option","CheckAccKey",cSelected(14))
				SaveSetting("AccessKey","Option","CheckEndChar",cSelected(15))
				SaveSetting("AccessKey","Option","CheckAcceler",cSelected(16))
				SaveSetting("AccessKey","Option","IgnoreVerTag",cSelected(17))
				SaveSetting("AccessKey","Option","IgnoreSetTag",cSelected(18))
				SaveSetting("AccessKey","Option","IgnoreDateTag",cSelected(19))
				SaveSetting("AccessKey","Option","IgnoreStateTag",cSelected(20))
				SaveSetting("AccessKey","Option","TgnoreAllTag",cSelected(21))
				SaveSetting("AccessKey","Option","NoCheckTag",cSelected(22))
				SaveSetting("AccessKey","Option","NoChangeTrnState",cSelected(23))
				SaveSetting("AccessKey","Option","AutoRepString",cSelected(24))
				SaveSetting("AccessKey","Option","KeepSetting",cSelected(25))
			End If
		End If
		If WriteType = "Sets" Or WriteType = "All" Then
			'删除原配置项
			HeaderIDs = GetSetting("AccessKey","Option","Headers")
			If HeaderIDs <> "" Then
				HeaderIDArr = Split(HeaderIDs,";",-1)
				On Error Resume Next
				For i = LBound(HeaderIDArr) To UBound(HeaderIDArr)
					DeleteSetting("AccessKey",HeaderIDArr(i))
				Next i
				On Error GoTo 0
			End If
			'写入新配置项
			For i = LBound(DataList) To UBound(DataList)
				ReDim Preserve HeaderIDArr(i)
				HeaderID = CStr(i)
				HeaderIDArr(i) = HeaderID
				TempArray = Split(DataList(i),JoinStr)
				SetsArray = Split(TempArray(1),SubJoinStr)
				SaveSetting("AccessKey",HeaderID,"Name",TempArray(0))
				SaveSetting("AccessKey",HeaderID,"ExcludeChar",SetsArray(0))
				SaveSetting("AccessKey",HeaderID,"LineSplitChar",SetsArray(1))
				SaveSetting("AccessKey",HeaderID,"CheckBracket",SetsArray(2))
				SaveSetting("AccessKey",HeaderID,"KeepCharPair",SetsArray(3))
				SaveSetting("AccessKey",HeaderID,"ShowAsiaKey",SetsArray(4))
				SaveSetting("AccessKey",HeaderID,"CheckEndChar",SetsArray(5))
				SaveSetting("AccessKey",HeaderID,"NoTrnEndChar",SetsArray(6))
				SaveSetting("AccessKey",HeaderID,"AutoTrnEndChar",SetsArray(7))
				SaveSetting("AccessKey",HeaderID,"CheckShortChar",SetsArray(8))
				SaveSetting("AccessKey",HeaderID,"CheckShortKey",SetsArray(9))
				SaveSetting("AccessKey",HeaderID,"KeepShortKey",SetsArray(10))
				SaveSetting("AccessKey",HeaderID,"PreRepString",SetsArray(11))
				SaveSetting("AccessKey",HeaderID,"AutoRepString",SetsArray(12))
				SaveSetting("AccessKey",HeaderID,"ApplyLangList",getLngPair(TempArray(2),"check"))
			Next i
			HeaderIDs = Join(HeaderIDArr,";")
			SaveSetting("AccessKey","Option","Headers",HeaderIDs)
		End If
		If WriteType = "Update" Or WriteType = "Sets" Or WriteType = "All" Then
			If Join(cUpdateSet) <> "" Then
				On Error Resume Next
				DeleteSetting("AccessKey","Update")
				On Error GoTo 0
				UpdateSiteList = Split(cUpdateSet(1),vbCrLf,-1)
				SaveSetting("AccessKey","Update","UpdateMode",cUpdateSet(0))
				For i = LBound(UpdateSiteList) To UBound(UpdateSiteList)
					SaveSetting("AccessKey","Update",CStr(i),UpdateSiteList(i))
				Next i
				SaveSetting("AccessKey","Update","Count",UBound(UpdateSiteList))
				SaveSetting("AccessKey","Update","Path",cUpdateSet(2))
				SaveSetting("AccessKey","Update","Argument",cUpdateSet(3))
				SaveSetting("AccessKey","Update","UpdateCycle",cUpdateSet(4))
				SaveSetting("AccessKey","Update","UpdateDate",cUpdateSet(5))
			End If
		End If
		CheckWrite = True
		cWriteLoc = CheckRegKey
		GoTo RemoveFilePath
	'删除所有保存的设置
	ElseIf Path = "" Then
		'删除文件配置项
 		RemoveFilePath:
		On Error Resume Next
		If Dir(CheckFilePath) <> "" Then
			SetAttr CheckFilePath,vbNormal
			Kill CheckFilePath
		End If
		TempPath = Left(CheckFilePath,InStrRev(CheckFilePath,"\"))
		If Dir(TempPath & "*.*") = "" Then RmDir TempPath
		On Error GoTo 0
		If Path = CheckRegKey Then GoTo ExitFunction
		'删除注册表配置项
		RemoveRegKey:
		If GetSetting("AccessKey","Option","Version") <> "" Then
			HeaderIDs = GetSetting("AccessKey","Option","Headers")
			On Error Resume Next
			If HeaderIDs <> "" Then
				HeaderIDArr = Split(HeaderIDs,";",-1)
				For i = LBound(HeaderIDArr) To UBound(HeaderIDArr)
					DeleteSetting("AccessKey",HeaderIDArr(i))
				Next i
			End If
			DeleteSetting("AccessKey","Option")
			DeleteSetting("AccessKey","Update")
			Dim WshShell As Object
			Set WshShell = CreateObject("WScript.Shell")
			WshShell.RegDelete CheckRegKey
			Set WshShell = Nothing
			On Error GoTo 0
		End If
		If Path = CheckFilePath Then GoTo ExitFunction
		'设置写入位置设置为空
		CheckWrite = True
		cWriteLoc = ""
	End If
	ExitFunction:
End Function


'替换导入配置文件中的语言名称为当前系统的语言名称
Function LangNameUpdate(Header As String,Data As String) As String
	Dim uLangArray() As String,dLangArray() As String,oLangArray() As String
	Dim i As Integer,j As Integer,Stemp As Boolean
	If OSLanguage = "0404" Then toOSLanguage = ""
	If OSLanguage <> "0404" Then toOSLanguage = "0404"
	uLangArray = Split(Data,SubJoinStr)
	dLangArray = LangCodeList("",OSLanguage,0,107)
	oLangArray = LangCodeList("",toOSLanguage,0,107)
	For i = LBound(uLangArray) To UBound(uLangArray)
		uLangPairList = Split(uLangArray(i),LngJoinStr)
		Stemp = False
		For j = LBound(dLangArray) To UBound(dLangArray)
			dLangPairList = Split(dLangArray(j),LngJoinStr)
			If uLangPairList(0) = dLangPairList(0) Then
				Stemp = True
				Exit For
			End If
		Next j
		If Stemp = True Then
			LangNameUpdate = Data
			Exit Function
		Else
			For j = LBound(oLangArray) To UBound(oLangArray)
				oLangPairList = Split(oLangArray(j),LngJoinStr)
				dLangPairList = Split(dLangArray(j),LngJoinStr)
				If uLangPairList(1) = oLangPairList(1) Then
					uLangPairList(0) = dLangPairList(0)
					Exit For
				End If
			Next j
			uLangArray(i) = Join(uLangPairList,LngJoinStr)
		End If
	Next i
	LangNameUpdate = Join(uLangArray,SubJoinStr)
End Function


'替换导入检查配置文件中的字符为当前系统的语言字符
Function CheckSetUpdate(Header As String,Data As String) As String
	Dim uSetsArray() As String,dSetsArray() As String,oSetsArray() As String
	Dim i As Integer,j As Integer,m As Integer,n As Integer,uDataList() As String
	CheckSetUpdate = Data
	If OSLanguage = "0404" Then toOSLanguage = ""
	If OSLanguage <> "0404" Then toOSLanguage = "0404"
	TempArray = Split(Data,JoinStr)
	uSetsArray = Split(TempArray(1),SubJoinStr)
	For i = LBound(DefaultCheckList) To UBound(DefaultCheckList)
		dData = CheckSettings(DefaultCheckList(i),OSLanguage)
		If CheckSetUpdate = dData Then Exit Function
		dSetsArray = Split(dData,SubJoinStr)
		oSetsArray = Split(CheckSettings(DefaultCheckList(i),toOSLanguage),SubJoinStr)
		For j = LBound(uSetsArray) To UBound(uSetsArray)
			uV = uSetsArray(j)
			dV = dSetsArray(j)
			oV = oSetsArray(j)
			If j <> 4 And uV <> "" And uV <> dV Then
				If j = 5 Or j = 7 Then
					spStr = " "
				Else
					spStr = ","
				End If
				If j = 7 And InStr(uV,"|") = 0 Then
					uDataList = Split(uV,spStr)
					For m = LBound(uDataList) To UBound(uDataList)
						uV = uDataList(m)
						uDataList(m) = Left(Trim(uV),1) & "|" & Right(Trim(uV),1)
					Next m
					uV = Join(uDataList,spStr)
				End If
				uDataList = Split(uV,spStr)
				dDataList = Split(dV,spStr)
				oDataList = Split(oV,spStr)
				For m = LBound(uDataList) To UBound(uDataList)
					For n = LBound(oDataList) To UBound(oDataList)
						If uDataList(m) = oDataList(n) Then
							uDataList(m) = dDataList(n)
							Exit For
						End If
					Next n
				Next m
				uSetsArray(j) = Join(uDataList,spStr)
			End If
		Next j
	Next i
	UpdatedData = Join(uSetsArray,SubJoinStr)
	CheckSetUpdate = TempArray(0) & JoinStr & UpdatedData & JoinStr & TempArray(2)
End Function


'更新检查旧版本配置值
Function CheckDataUpdate(Header As String,Data As String) As String
	Dim UpdatedData As String,uV As String,dV As String,spStr As String
	Dim i As Integer,m As Integer,uDataList() As String
	TempArray = Split(Data,JoinStr)
	uSetsArray = Split(TempArray(1),SubJoinStr)
	dSetsArray = Split(CheckSettings(Header,OSLanguage),SubJoinStr)
	For i = LBound(uSetsArray) To UBound(uSetsArray)
		uV = uSetsArray(i)
		dV = dSetsArray(i)
		If uV = "" Then uV = dV
		If i <> 4 And uV <> "" And uV <> dV Then
			If i = 5 Or i = 7 Then
				spStr = " "
			Else
				spStr = ","
			End If
			If i = 7 And InStr(uV,"|") = 0 Then
				uDataList = Split(uV,spStr)
				For m = LBound(uDataList) To UBound(uDataList)
					uV = uDataList(m)
					uDataList(m) = Left(Trim(uV),1) & "|" & Right(Trim(uV),1)
				Next m
				uV = Join(uDataList,spStr)
			End If
			uV = Join(ClearArray(Split(uV & spStr & dV,spStr,-1)),spStr)
			uSetsArray(i) = uV
		End If
	Next i
	UpdatedData = Join(uSetsArray,SubJoinStr)
	CheckDataUpdate = TempArray(0) & JoinStr & UpdatedData & JoinStr & TempArray(2)
End Function


'增加或更改数组项目
Function CreateArray(Header As String,Data As String,HeaderList() As String,DataList() As String) As Boolean
	Dim i As Integer,n As Integer
	If HeaderList(0) = "" Then
		HeaderList(0) = Header
		DataList(0) = Data
	Else
		n = 0
		For i = LBound(HeaderList) To UBound(HeaderList)
			If LCase(HeaderList(i)) = LCase(Header) Then
				If DataList(i) <> Data Then DataList(i) = Data
				n = n + 1
				Exit For
			End If
		Next i
		If n = 0 Then
			i = UBound(HeaderList) + 1
			ReDim Preserve HeaderList(i),DataList(i)
			HeaderList(i) = Header
			If DataList(i) <> Data Then DataList(i) = Data
		End If
	End If
End Function


'删除数组项目
Function DelArray(dName As String,dList() As String,Separator As String,Num As Integer) As Variant
	Dim i As Integer,n As Integer,TempList() As String
	ReDim TempList(0)
	n = 0
	For i = LBound(dList) To UBound(dList)
		If Separator <> "" Then
			LangPairList = Split(dList(i),Separator)
			If LangPairList(Num) <> dName Then
				ReDim Preserve TempList(n)
				TempList(n) = dList(i)
				n = n + 1
			End If
		Else
			If dList(i) <> dName Then
				ReDim Preserve TempList(n)
				TempList(n) = dList(i)
				n = n + 1
			End If
		End If
	Next i
	DelArray = TempList
End Function


'拆分数组
Function SplitData(Data As String,NameList() As String,SrcList() As String,TranList() As String) As Boolean
	Dim i As Integer
	SplitData = False
	ReDim NameList(0),SrcList(0),TranList(0)
	LangArray = Split(Data,SubJoinStr)
	For i = LBound(LangArray) To UBound(LangArray)
		LangPairList = Split(LangArray(i),LngJoinStr)
		If LangPairList(0) <> "" Then
			If LangPairList(1) = "" Then
				srcLng = NullValue
			Else
				srcLng = LangPairList(1)
			End If
			If LangPairList(2) = "" Then
				TranLng = NullValue
			Else
				TranLng = LangPairList(2)
			End If
			ReDim Preserve NameList(i),SrcList(i),TranList(i)
			NameList(i) = LangPairList(0)
			SrcList(i) = srcLng
			TranList(i) = TranLng
			SplitData = True
		End If
	Next i
End Function


'创建二个互补的名称数组
Function getLngNameList(Data As String,NameList() As String,SrcList() As String) As Boolean
	Dim i As Integer,j As Integer,n As Integer
	getLngNameList = False
	ReDim NameList(0),SrcList(0)
	LangArray = Split(Data,SubJoinStr)
	n = 0
	j = 0
	For i = LBound(LangArray) To UBound(LangArray)
		If LangArray(i) <> "" Then
			LangPairList = Split(LangArray(i),LngJoinStr)
			If LangPairList(2) = "" Then
				ReDim Preserve NameList(j)
				NameList(j) = LangPairList(0)
				j = j + 1
			Else
				ReDim Preserve SrcList(n)
				SrcList(n) = LangPairList(0)
				n = n + 1
			End If
			getLngNameList = True
		End If
	Next i
End Function


'合并标准和自定义语言列表
Function MergeLngList(LangArray() As String,DataArray() As String,fType As String) As Variant
	Dim i As Integer,j As Integer,n As Integer,Stemp As Boolean,TempList() As String
	TempList = LangArray
	For i = LBound(DataArray) To UBound(DataArray)
		LangDataList = Split(DataArray(i),LngJoinStr)
		Stemp = False
		For j = LBound(LangArray) To UBound(LangArray)
			LangPairList = Split(LangArray(j),LngJoinStr)
			If LCase(LangDataList(1)) = LCase(LangPairList(1)) Then
				If fType = "engine" Then TempList(j) = DataArray(i)
				If fType = "check" Then TempList(j) = DataArray(i) & LngJoinStr & LangDataList(1)
				Stemp = True
				Exit For
			End If
		Next j
		If Stemp = False Then
			n = UBound(TempList) + 1
			ReDim Preserve TempList(n)
			If fType = "engine" Then TempList(n) = DataArray(i)
			If fType = "check" Then TempList(n) = DataArray(i) & LngJoinStr & LangDataList(1)
		End If
	Next i
	MergeLngList = TempList
End Function


'生成去除数据列表中空项后的语言对
Function getLngPair(Data As String,fType As String) As String
	Dim LngNameList() As String,SrcLngList() As String,TranLngList() As String
	Dim n As Integer,TempArray() As String
	getLngPair = Data
	If SplitData(Data,LngNameList,SrcLngList,TranLngList) = True Then
		n = 0
		ReDim TempArray(0)
		For i = LBound(LngNameList) To UBound(LngNameList)
			LngName = LngNameList(i)
			SrcLngCode = SrcLngList(i)
			TranLngCode = TranLngList(i)
			If SrcLngCode = NullValue Then SrcLngCode = ""
			If TranLngCode = NullValue Then TranLngCode = ""
			If TranLngCode <> "" Then
				ReDim Preserve TempArray(n)
				If fType = "engine" Then
					TempArray(n) = LngName & LngJoinStr & SrcLngCode & LngJoinStr & TranLngCode
				ElseIf fType = "check" Then
					TempArray(n) = LngName & LngJoinStr & TranLngCode
				End If
				n = n + 1
			End If
		Next i
		getLngPair = Join(TempArray,SubJoinStr)
	End If
End Function


'互换数组项目
Function ChangeLngNameList(Data As String,UseList() As String) As Variant
	Dim i As Integer,j As Integer,n As Integer,LngNameList() As String,Stemp As Boolean
	ReDim LngNameList(0)
	LangArray = Split(Data,SubJoinStr)
	n = 0
	For i = LBound(LangArray) To UBound(LangArray)
		LangPairList = Split(LangArray(i),LngJoinStr)
		Stemp = False
		For j = LBound(UseList) To UBound(UseList)
			If UseList(j) = LangPairList(0) Then
				Stemp = True
				Exit For
			End If
		Next j
		If Stemp = False Then
			ReDim Preserve LngNameList(n)
			LngNameList(n) = LangPairList(0)
			n = n + 1
		End If
	Next i
	ChangeLngNameList = LngNameList
End Function


'查找指定值是否在数组中
Function getCheckID(DataList() As String,LngCode As String,OldLngCode As String) As Integer
	Dim i As Integer,j As Integer,Stemp As Boolean
	Stemp = False
	For i = LBound(DataList) To UBound(DataList)
		TempArray = Split(DataList(i),JoinStr)
		If TempArray(2) <> "" Then
			LangArray = Split(TempArray(2),SubJoinStr)
			For j = LBound(LangArray) To UBound(LangArray)
				LangPairList = Split(LangArray(j),LngJoinStr)
				If LCase(LangPairList(2)) = LCase(LngCode) Then
					getCheckID = i
					Stemp = True
					Exit For
				End If
			Next j
		End If
		If Stemp = True Then Exit For
	Next i
	If Stemp = False Then
		If OldLngCode = "Asia" Then getCheckID = 0
		If OldLngCode <> "Asia" Then getCheckID = 1
	End If
End Function


'检查数组中是否有空值
'ftype = 0     检查数组项内是否全为空值
'ftype = 1     检查数组项内是否有空值
'Header = ""   检查整个数组
'Header <> ""  检查指定数组项
Function CheckNullData(Header As String,DataList() As String,SkipNum As String,fType As Integer) As Boolean
	Dim i As Integer,j As Integer,x As Integer,m As Integer,n As Integer,Stemp As Boolean,hStemp As Boolean
	CheckNullData = False
	SkipNumArray = Split(SkipNum,",")
	m = 0
	hStemp = False
	For i = LBound(DataList) To UBound(DataList)
		TempArray = Split(DataList(i),JoinStr)
		SetsArray = Split(TempArray(1),SubJoinStr)
		If Header <> "" And TempArray(0) = Header Then hStemp = True
		If Header = "" Then hStemp = True
		If hStemp = True Then
			n = 0
			For j = LBound(SetsArray) To UBound(SetsArray)
				Stemp = False
				For x = 0 To UBound(SkipNumArray)
					If CStr(j) = SkipNumArray(x) Then
						Stemp = True
						Exit For
					End If
				Next x
				If Trim(SetsArray(j)) = "" And Stemp = False Then
					If fType = 0 Then n = n + 1
					If fType = 1 Then
						CheckNullData = True
						Exit For
					End If
				End If
			Next j
			If getLngPair(TempArray(2),"check") = "" Then
				If fType = 0 Then n = n + 1
				If fType = 1 Then CheckNullData = True
			End If
			If fType = 0 Then
				If Header <> "" Then
					If n = UBound(SetsArray)-UBound(SkipNumArray)+1 Then CheckNullData = True
				Else
					If n = UBound(SetsArray)-UBound(SkipNumArray)+1 Then m = m + 1
					If m = UBound(DataList)+1 Then CheckNullData = True
				End If
			End If
			If Header <> "" Then Exit For
			If fType = 1 And CheckNullData = True Then Exit For
		End If
	Next i
	If Header <> "" And hStemp = False Then CheckNullData = True
End Function


'数组排序
Function SortArray(xArray() As String,Comp As Integer,CompType As String,Operator As String) As Variant
	Dim rMin As Integer,rMax As Integer,MaxLng As Integer,Lng As Integer
	Dim fLng As Integer,sLng As Integer,fMargin As Integer,sMargin As Integer
	rMin = LBound(xArray)
	rMax = UBound(xArray)
	If rMax = 0 Or CompType = "" Or Operator = "" Then
		SortArray = xArray
		Exit Function
	End If
    MaxLng = 1
    For x = rMin To rMax
        Lng = Len(Trim(xArray(x)))
        If Lng > MaxLng Then MaxLng = Lng
    Next
	For x = rMax To rMin Step -1
		For y = rMin To rMax - 1
			fLng = Len(Trim(xArray(y)))
			sLng = Len(Trim(xArray(y+1)))
			If CompType = "Size" Then
				fMargin = MaxLng - fLng
				sMargin = MaxLng - sLng
				fValue = String(fMargin,"0") & xArray(y)
				sValue = String(sMargin,"0") & xArray(y+1)
				MyComp = StrComp(fValue,sValue,Comp)
			End If
			If CompType = "Lenght" Then
				If fLng < sLng Then MyComp = -1
				If fLng = sLng Then MyComp = 0
				If fLng > sLng Then MyComp = 1
			End If
			If Operator = ">" Then
				If MyComp > 0 Then
					Mx = xArray(y + 1)
					xArray(y+1) = xArray(y)
					xArray(y) = Mx
				End If
			ElseIf Operator = "<" Then
				If MyComp < 0 Then
					Mx = xArray(y + 1)
					xArray(y+1) = xArray(y)
					xArray(y) = Mx
				End If
			ElseIf Operator = "=" Then
				If MyComp = 0 Then
					Mx = xArray(y + 1)
					xArray(y+1) = xArray(y)
					xArray(y) = Mx
				End If
			End If
		Next
	Next
	SortArray = xArray
End Function


'清理数组中重复的数据
Function ClearArray(xArray() As String) As Variant
	Dim yArray() As String,Stemp As Boolean,i As Integer,j As Integer,y As Integer
	If UBound(xArray) = 0 Then
		ClearArray = xArray
		Exit Function
	End If
	y = 0
	ReDim yArray(0)
	For i = LBound(xArray) To UBound(xArray)
		Stemp = False
		For j = i + 1 To UBound(xArray)
			If xArray(i) = xArray(j) Then
				Stemp = True
				Exit For
			End If
		Next j
		If Stemp = False Then
			ReDim Preserve yArray(y)
			yArray(y) = xArray(i)
			y = y + 1
		End If
	Next i
	ClearArray = yArray
End Function


'通配符查找指定值
Function CheckKeyCode(FindKey As String,CheckKey As String) As Integer
	Dim KeyCode As Boolean,FindStr As String,Key As String,Pos As Integer
	Key = Trim(FindKey)
	CheckKeyCode = 0
	If InStr(Key,"%") Then Key = Replace(Key,"%","x")
	If CheckKey <> "" And Key <> "" Then
		FindStrArr = Split(Convert(CheckKey),",",-1)
		For i = LBound(FindStrArr) To UBound(FindStrArr)
			FindStr = Trim(FindStrArr(i))
			If InStr(FindStr,"%") Then FindStr = Replace(FindStr,"%","x")
			If InStr(FindStr,"-") Then
				If Left(FindStr,1) <> "[" And Right(FindStr,1) <> "]" Then
					FindStr = "[" & FindStr & "]"
				End If
			End If
			If InStr(FindStr,"[") Then
				Pos = InStr(FindStr,"[")
				If Left(FindStr,Pos-1) <> "[" And Right(FindStr,Pos+1) <> "]" Then
					FindStr = Replace(FindStr,"[","[[]")
				End If
			End If
			KeyCode = False
			CheckKeyCode = 0
			'PSL.Output Key & " : " &  FindStr  '调试用
			KeyCode = UCase(Key) Like UCase(FindStr)
			If KeyCode = True Then CheckKeyCode = 1
			If KeyCode = True Then Exit For
		Next i
	ElseIf CheckKey = "" And Key <> "" Then
		CheckKeyCode = 1
	End If
End Function


'测试检查程序
Sub CheckTest(CheckID As Integer,HeaderList() As String)
	Dim TrnList As PslTransList,i As Integer,TrnListDec As String,TrnListArray() As String
	If OSLanguage = "0404" Then
		Msg01 = "把计代刚"
		Msg02 = "沮兵ン舱穓碝陆亩岿粇倒タ [代刚] 秙块挡狦"
		Msg03 = "砞﹚嘿:"
		Msg05 = "陆亩睲虫:"
		Msg06 = "弄︽计:"
		Msg07 = "ず甧:"
		Msg08 = "や穿窾ノじ㎝だ腹だ筳兜"
		Msg09 = "﹃ず甧:"
		Msg10 = "场(&F)"
		Msg11 = "獽倍龄(&K)"
		Msg12 = "沧ゎ才(&E)"
		Msg13 = "硉竟(&P)"
		Msg15 = "代刚(&T)"
		Msg16 = "睲(&C)"
		Msg18 = "弧(&H)"
		Msg19 = "笆蠢传じ(&R)"
	Else
		Msg01 = "参数测试"
		Msg02 = "根据下列条件组合查找翻译错误并给出修正。按 [测试] 按钮输出结果。"
		Msg03 = "配置名称:"
		Msg05 = "翻译列表:"
		Msg06 = "读入行数:"
		Msg07 = "包含内容:"
		Msg08 = "支持通配符和半角分号分隔的多项"
		Msg09 = "字串内容:"
		Msg10 = "全部(&F)"
		Msg11 = "快捷键(&K)"
		Msg12 = "终止符(&E)"
		Msg13 = "加速器(&P)"
		Msg15 = "测试(&T)"
		Msg16 = "清空(&C)"
		Msg18 = "帮助(&H)"
		Msg19 = "自动替换字符(&R)"
	End If

	ReDim TrnListArray(0)
	For i = 1 To trn.Project.TransLists.Count
		Set TrnList = trn.Project.TransLists(i)
		ReDim Preserve TrnListArray(i)
		TrnListDec = TrnList.Title & " - " & PSL.GetLangCode(TrnList.Language.LangID,pslCodeText)
		TrnListArray(i) = TrnListDec
	Next i
	Begin Dialog UserDialog 650,462,Msg01,.CheckTestFunc ' %GRID:10,7,1,1
		GroupBox 10,21,630,119,"",.GroupBox
		Text 10,7,630,14,Msg02
		Text 30,38,80,14,Msg03,.SetNameText
		DropListBox 120,35,340,21,HeaderList(),.SelSetBox
		CheckBox 480,35,140,14,Msg19,.RepStrBox
		Text 30,66,80,14,Msg05
		DropListBox 120,63,340,21,TrnListArray(),.TrnListBox
		Text 480,66,80,14,Msg06
		TextBox 572,63,50,18,.LineNumBox
		Text 30,94,80,14,Msg07
		TextBox 120,91,340,21,.SpecifyTextBox
		Text 480,90,140,28,Msg08
		Text 30,119,80,14,Msg09
		CheckBox 120,119,110,14,Msg10,.AllCheckBox
		CheckBox 240,119,110,14,Msg11,.AcckeyCheckBox
		CheckBox 360,119,110,14,Msg12,.EndSharCheckBox
		CheckBox 480,119,120,14,Msg13,.ShortCheckBox
		TextBox 10,147,630,280,.InTextBox,1
		PushButton 10,434,100,21,Msg18,.HelpButton
		PushButton 120,434,100,21,Msg15,.TestButton
		PushButton 230,434,100,21,Msg16,.ClearButton
		CancelButton 540,434,100,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	dlg.SelSetBox = CheckID
	If Dialog(dlg) = 0 Then Exit Sub
End Sub


'测试对话框函数
Private Function CheckTestFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim ListDec As String,TempDec As String,LineNum As Integer,inText As String,repStr As Integer
	Dim cAllCont As Integer,cAccKey As Integer,cEndChar As Integer,cAcceler As Integer
	Dim SpecifyText As String,CheckID As Integer
	If OSLanguage = "0404" Then
		Msg01 = "タ穓碝岿粇倒タ惠璶碭だ牧叫祔..."
	Else
		Msg01 = "正在查找错误并给出修正，可能需要几分钟，请稍候..."
	End If

	Select Case Action%
	Case 1
		TempDec = trn.Title & " - " & PSL.GetLangCode(trn.Language.LangID,pslCodeText)
		LineNum = 10
		DlgText "TrnListBox",TempDec
		DlgText "LineNumBox",CStr(LineNum)
		DlgValue "AllCheckBox",1
    	inText = DlgText("InTextBox")
    	If inText <> "" Then DlgEnable "ClearButton",True
    	If inText = "" Then DlgEnable "ClearButton",False
	Case 2 ' 数值更改或者按下了按钮
		If DlgItem$ = "AllCheckBox" Then
			cAllCont= DlgValue("AllCheckBox")
			cAccKey = DlgValue("AcckeyCheckBox")
			cEndChar = DlgValue("EndSharCheckBox")
			cAcceler = DlgValue("ShortCheckBox")
			If cAllCont = 1 Then
				DlgValue "AcckeyCheckBox",0
				DlgValue "EndSharCheckBox",0
				DlgValue "ShortCheckBox",0
			End If
			If cAccKey + cEndChar + cAcceler = 0 Then
				DlgValue "AllCheckBox",1
			End If
		End If
		If DlgItem$ = "AcckeyCheckBox" Or DlgItem$ = "EndSharCheckBox" Or DlgItem$ = "ShortCheckBox" Then
			cAccKey = DlgValue("AcckeyCheckBox")
			cEndChar = DlgValue("EndSharCheckBox")
			cAcceler = DlgValue("ShortCheckBox")
			If cAccKey = 1 Or cEndChar = 1 Or cAcceler = 1 Then
				DlgValue "AllCheckBox",0
			End If
			If cAccKey + cEndChar + cAcceler = 0 Then
				DlgValue "AllCheckBox",1
			End If
		End If
		If DlgItem$ = "TestButton" Then
			CheckID = DlgValue("SelSetBox")
			repStr = DlgValue("RepStrBox")
			ListDec = DlgText("TrnListBox")
			LineNum = CLng(DlgText("LineNumBox"))
			SpecifyText = DlgText("SpecifyTextBox")
			AllCont = DlgValue("AllCheckBox")
			AccKey = DlgValue("AcckeyCheckBox")
			EndChar = DlgValue("EndSharCheckBox")
			Acceler = DlgValue("ShortCheckBox")
			DlgText "InTextBox",Msg01
			inText = CheckStrings(CheckID,ListDec,LineNum,repStr,SpecifyText)
    		DlgText "InTextBox",inText
		End If
		If DlgItem$ = "ClearButton" Then
			DlgText "InTextBox",""
			DlgEnable "ClearButton",False
		End If
		If DlgItem$ = "HelpButton" Then
			Call CheckHelp("TestHelp")
		End If
		If DlgItem$ <> "CancelButton" Then
		    inText = DlgText("InTextBox")
    		If inText <> "" Then DlgEnable "ClearButton",True
    		If inText = "" Then DlgEnable "ClearButton",False
			CheckTestFunc = True '防止按下按钮关闭对话框窗口
		End If
	End Select
End Function


'处理符合条件的字串列表中的字串
Function CheckStrings(ID As Integer,ListDec As String,LineNum As Integer,rStr As Integer,sText As String) As String
	Dim i As Integer,j As Integer,k As Integer,srcString As String,trnString As String,tText As String
	Dim TrnList As PslTransList,TrnListDec As String,TranLang As String
	Dim CheckVer As String,CheckSet As String,CheckState As String,CheckDate As Date,TranDate As Date
	Dim LineNumErrCount As Integer,accKeyNumErrCount As Integer,ErrorCount As Integer,Massage As String
	Dim Find As Boolean,srcFindNum As Integer,trnFindNum As Integer

	If OSLanguage = "0404" Then
		Msg01 = "ゅ: "
		Msg02 = "亩ゅ: "
		Msg03 = "タ: "
		Msg04 = "癟: "
		Msg05 = "---------------------------------------"
		Msg06 = "т岿粇:"
		Msg07 = "⊿Τт岿粇"
		Msg08 = "﹚ず甧﹃い⊿Τт岿粇"
		Msg09 = "⊿Τт﹚ず甧﹃"
	Else
		Msg01 = "原文: "
		Msg02 = "译文: "
		Msg03 = "修正: "
		Msg04 = "消息: "
		Msg05 = "---------------------------------------"
		Msg06 = "找到下列错误:"
		Msg07 = "没有找到错误。"
		Msg08 = "包含指定内容的字串中没有找到错误。"
		Msg09 = "没有找到包含指定内容的字串！"
	End If

	'参数初始化
	CheckStrings = ""
	tText = ""
	k = 0
	LineNumErrCount = 0
	accKeyNumErrCount = 0
	Find = False

	'获取选定的翻译列表
	For i = 1 To trn.Project.TransLists.Count
		Set TrnList = trn.Project.TransLists(i)
		TrnListDec = TrnList.Title & " - " & PSL.GetLangCode(TrnList.Language.LangID,pslCodeText)
		If TrnListDec = ListDec Then Exit For
	Next i

	'获取目标语言
	trnLng = PSL.GetLangCode(TrnList.Language.LangID,pslCode639_1)
	If trnLng = "zh" Or trnLng = "ja" Or trnLng = "ko" Then TranLang = "Asia"

	If TrnList.SourceList.LastChange > TrnList.LastUpdate Then TrnList.Update
	If LineNum > TrnList.StringCount Then LineNum = TrnList.StringCount
	For i = 1 To TrnList.StringCount
		'参数初始化
		srcString = ""
		trnString = ""
		NewtrnString = ""
		LineMsg = ""
		AccKeyMsg = ""
		ReplaceMsg = ""

		'获取原文和翻译字串
		Set TransString = TrnList.String(i)
		If TransString.Text <> "" Then
			srcString = TransString.SourceText
			trnString = TransString.Text
			OldtrnString = trnString

			'开始处理字串
			If sText <> "" Then
				FindStrArr = Split(sText,";",-1)
				For j = LBound(FindStrArr) To UBound(FindStrArr)
					FindStr = FindStrArr(j)
					If Left(FindStr,1) <> "*" And Right(FindStr,1) <> "*" Then
						FindStr = "*" & FindStr & "*"
					End If
					srcFindNum = CheckKeyCode(srcString,FindStr)
					trnFindNum = CheckKeyCode(trnString,FindStr)
					If srcFindNum <> 0 Or trnFindNum <> 0 Then
						NewtrnString = CheckHanding(ID,srcString,trnString,TranLang)
						Find = True
					End If
					If NewtrnString <> "" Then Exit For
				Next j
			Else
				NewtrnString = CheckHanding(ID,srcString,trnString,TranLang)
			End If
			If rStr = 1 Then NewtrnString = ReplaceStr(ID,NewtrnString,0)

			'调用消息输出
			If NewtrnString <> "" And NewtrnString <> OldtrnString Then
				If srcLineNum <> trnLineNum Then
					'LineMsg = LineErrMassage(srcLineNum,trnLineNum,LineNumErrCount)
				End If
				If srcAccKeyNum <> trnAccKeyNum Then
					'AccKeyMsg = AccKeyErrMassage(srcAccKeyNum,trnAccKeyNum,accKeyNumErrCount)
				End If
				If NewtrnString <> OldtrnString Then
					ReplaceMsg = ReplaceMassage(OldtrnString,NewtrnString)
				End If
				srcString = Msg01 & srcString		'vbBlack
				OldtrnString = Msg02 & OldtrnString	'vbGreen
				NewtrnString = Msg03 & NewtrnString	'vbRed
				Massage = Msg04 & LineMsg & AccKeyMsg & ReplaceMsg
				tString = srcString & vbCrLf & OldtrnString & vbCrLf & NewtrnString
				tString = tString & vbCrLf & Massage & vbCrLf & Msg05 & Msg05 & Msg05
				If tText <> "" Then tText = tText & vbCrLf & tString
				If tText = "" Then tText = tString
				k = k + 1
			End If
		End If
		If k = LineNum Then Exit For
	Next i
	If tText <> "" Then CheckStrings = Msg06 & vbCrLf & Msg05 & Msg05 & Msg05 & vbCrLf & tText
	If tText = "" And sText = "" Then CheckStrings = Msg07
	If tText = "" And sText <> "" And Find = True Then CheckStrings = Msg08
	If tText = "" And sText <> "" And Find = False Then CheckStrings = Msg09
End Function


'除去字串前后指定的 PreStr 和 AppStr
Function RemoveBackslash(folderPath As String,PreStr As String,AppStr As String,Spase As Integer) As String
	Dim folderPathBak As String,Stemp As Boolean,i As Integer
	If folderPath = "" Then Exit Function
	folderPathBak = folderPath
	If Spase = 1 Then folderPathBak = Trim(folderPathBak)
	For i = 0 To 1
		Stemp = False
		If PreStr <> "" And Left(folderPathBak,Len(PreStr)) = PreStr Then
			folderPathBak = Mid(folderPathBak,Len(PreStr)+1)
			Stemp = True
		End If
		If AppStr <> "" And Right(folderPathBak,Len(AppStr)) = AppStr Then
			folderPathBak = Left(folderPathBak,Len(folderPathBak)-Len(AppStr))
			Stemp = True
		End If
		If Spase = 1 Then folderPathBak = Trim(folderPathBak)
		If Stemp = True Then i = 0
		If Stemp = False Then Exit For
	Next i
	RemoveBackslash = folderPathBak
End Function


'读取语言对
Function LangCodeList(DataName As String,OSLang As String,MinNum As Integer,MaxNum As Integer) As Variant
	Dim i As Integer,j As Integer,Code As String,LangName() As String,LangPair() As String
	ReDim LangName(MaxNum - MinNum),LangPair(MaxNum - MinNum)
	For i = MinNum To MaxNum
		j = i - MinNum
		If OSLang = "0404" Then
			If i = 0 Then LangName(j) = "笆盎代"
			If i = 1 Then LangName(j) = "玭獶颤孽粂"
			If i = 2 Then LangName(j) = "焊ぺェㄈ粂"
			If i = 3 Then LangName(j) = "﹊┰粂"
			If i = 4 Then LangName(j) = "┰粂"
			If i = 5 Then LangName(j) = "ㄈェㄈ粂"
			If i = 6 Then LangName(j) = "履﹊粂"
			If i = 7 Then LangName(j) = "峨忙粂"
			If i = 8 Then LangName(j) = "ぺぐ膀焊粂"
			If i = 9 Then LangName(j) = "ぺ吹粂"
			If i = 10 Then LangName(j) = "フ玐霉吹粂"
			If i = 11 Then LangName(j) = "﹕┰粂"
			If i = 12 Then LangName(j) = "猧﹁ェㄈ粂"
			If i = 13 Then LangName(j) = "ガ娥ェ粂"
			If i = 14 Then LangName(j) = "玂ㄈ粂"
			If i = 15 Then LangName(j) = "霉ェㄈ粂"
			If i = 16 Then LangName(j) = "虏砰いゅ"
			If i = 17 Then LangName(j) = "タ砰いゅ"
			If i = 18 Then LangName(j) = "﹁古粂"
			If i = 19 Then LangName(j) = "霉ㄈ粂"
			If i = 20 Then LangName(j) = "倍粂"
			If i = 21 Then LangName(j) = "う沉粂"
			If i = 22 Then LangName(j) = "颤孽粂"
			If i = 23 Then LangName(j) = "璣粂"
			If i = 24 Then LangName(j) = "稲‵ェㄈ粂"
			If i = 25 Then LangName(j) = "猭霉粂"
			If i = 26 Then LangName(j) = "猧吹粂"
			If i = 27 Then LangName(j) = "孽粂"
			If i = 28 Then LangName(j) = "猭粂"
			If i = 29 Then LangName(j) = "ケ柑﹁ㄈ粂"
			If i = 30 Then LangName(j) = "﹁ㄈ粂"
			If i = 31 Then LangName(j) = "緗ㄈ粂"
			If i = 32 Then LangName(j) = "紈粂"
			If i = 33 Then LangName(j) = "镁粂"
			If i = 34 Then LangName(j) = "钞孽粂"
			If i = 35 Then LangName(j) = "┰疭粂"
			If i = 36 Then LangName(j) = "花履粂"
			If i = 37 Then LangName(j) = "ㄓ粂"
			If i = 38 Then LangName(j) = "粂"
			If i = 39 Then LangName(j) = "粂"
			If i = 40 Then LangName(j) = "畄粂"
			If i = 41 Then LangName(j) = "ェ﹁ㄈ粂"
			If i = 42 Then LangName(j) = "疭粂"
			If i = 43 Then LangName(j) = "稲焊孽粂"
			If i = 44 Then LangName(j) = "痁瓜粂"
			If i = 45 Then LangName(j) = "緗粂"
			If i = 46 Then LangName(j) = "種粂"
			If i = 47 Then LangName(j) = "ら粂"
			If i = 48 Then LangName(j) = "笷粂"
			If i = 49 Then LangName(j) = "ぐμ焊粂"
			If i = 50 Then LangName(j) = "履粂"
			If i = 51 Then LangName(j) = "蔼粗粂"
			If i = 52 Then LangName(j) = "縞笷粂"
			If i = 53 Then LangName(j) = "ふェ粂"
			If i = 54 Then LangName(j) = "绰翧粂"
			If i = 55 Then LangName(j) = "焊吹粂"
			If i = 56 Then LangName(j) = "焊吹粂 (焊吹㈱)"
			If i = 57 Then LangName(j) = "ρ锯粂"
			If i = 58 Then LangName(j) = "┰叉蝴ㄈ粂"
			If i = 59 Then LangName(j) = "ミ吵﹞粂"
			If i = 60 Then LangName(j) = "縞此躇粂"
			If i = 61 Then LangName(j) = "皑ㄤ箉粂"
			If i = 62 Then LangName(j) = "皑ㄓ粂"
			If i = 63 Then LangName(j) = "皑┰懂┰﹊粂"
			If i = 64 Then LangName(j) = "皑φ粂"
			If i = 65 Then LangName(j) = "を粂"
			If i = 66 Then LangName(j) = "皑┰粂"
			If i = 67 Then LangName(j) = "籜粂"
			If i = 68 Then LangName(j) = "ェ獃焊粂"
			If i = 69 Then LangName(j) = "粂"
			If i = 70 Then LangName(j) = "粂 (痴皑焊ゅ)"
			If i = 71 Then LangName(j) = "粂 (ェ空吹ゅ)"
			If i = 72 Then LangName(j) = "而柑懂粂"
			If i = 73 Then LangName(j) = "炊ぐ瓜粂"
			If i = 74 Then LangName(j) = "猧孽粂"
			If i = 75 Then LangName(j) = "覆靛粂"
			If i = 76 Then LangName(j) = "綛炊粂"
			If i = 77 Then LangName(j) = "ㄈ粂"
			If i = 78 Then LangName(j) = "霉皑ェㄈ粂"
			If i = 79 Then LangName(j) = "玐粂"
			If i = 80 Then LangName(j) = "履μ粂"
			If i = 81 Then LangName(j) = "彪粂"
			If i = 82 Then LangName(j) = "峨焊蝴ㄈ粂"
			If i = 83 Then LangName(j) = "ぺΛ粂"
			If i = 84 Then LangName(j) = "ニ粂"
			If i = 85 Then LangName(j) = "獺紈粂"
			If i = 86 Then LangName(j) = "宫霉粂"
			If i = 87 Then LangName(j) = "吹ワ粂"
			If i = 88 Then LangName(j) = "吹ゅェㄈ粂"
			If i = 89 Then LangName(j) = "﹁痁粂"
			If i = 90 Then LangName(j) = "吹ニ柑粂"
			If i = 91 Then LangName(j) = "风ㄥ粂"
			If i = 92 Then LangName(j) = "痹ㄈ粂"
			If i = 93 Then LangName(j) = "娥粂"
			If i = 94 Then LangName(j) = "μ焊粂"
			If i = 95 Then LangName(j) = "哦晦粂"
			If i = 96 Then LangName(j) = "縞㏕粂"
			If i = 97 Then LangName(j) = "粂"
			If i = 98 Then LangName(j) = "旅粂"
			If i = 99 Then LangName(j) = "φㄤ粂"
			If i = 100 Then LangName(j) = "畐耙粂"
			If i = 101 Then LangName(j) = "蝴焊粂"
			If i = 102 Then LangName(j) = "疩孽粂"
			If i = 103 Then LangName(j) = "疩焊常粂"
			If i = 104 Then LangName(j) = "疩粂"
			If i = 105 Then LangName(j) = "禫玭粂"
			If i = 106 Then LangName(j) = "焊粂"
			If i = 107 Then LangName(j) = "║ひ粂"
		Else
			If i = 0 Then LangName(j) = "自动检测"
			If i = 1 Then LangName(j) = "南非荷兰语"
			If i = 2 Then LangName(j) = "阿尔巴尼亚语"
			If i = 3 Then LangName(j) = "阿姆哈拉语"
			If i = 4 Then LangName(j) = "阿拉伯语"
			If i = 5 Then LangName(j) = "亚美尼亚语"
			If i = 6 Then LangName(j) = "阿萨姆语"
			If i = 7 Then LangName(j) = "阿塞拜疆语"
			If i = 8 Then LangName(j) = "巴什基尔语"
			If i = 9 Then LangName(j) = "巴斯克语"
			If i = 10 Then LangName(j) = "白俄罗斯语"
			If i = 11 Then LangName(j) = "孟加拉语"
			If i = 12 Then LangName(j) = "波西尼亚语"
			If i = 13 Then LangName(j) = "布列塔尼语"
			If i = 14 Then LangName(j) = "保加利亚语"
			If i = 15 Then LangName(j) = "加泰罗尼亚语"
			If i = 16 Then LangName(j) = "简体中文"
			If i = 17 Then LangName(j) = "繁体中文"
			If i = 18 Then LangName(j) = "科西嘉语"
			If i = 19 Then LangName(j) = "克罗地亚语"
			If i = 20 Then LangName(j) = "捷克语"
			If i = 21 Then LangName(j) = "丹麦语"
			If i = 22 Then LangName(j) = "荷兰语"
			If i = 23 Then LangName(j) = "英语"
			If i = 24 Then LangName(j) = "爱沙尼亚语"
			If i = 25 Then LangName(j) = "法罗语"
			If i = 26 Then LangName(j) = "波斯语"
			If i = 27 Then LangName(j) = "芬兰语"
			If i = 28 Then LangName(j) = "法语"
			If i = 29 Then LangName(j) = "弗里西亚语"
			If i = 30 Then LangName(j) = "加利西亚语"
			If i = 31 Then LangName(j) = "格鲁吉亚语"
			If i = 32 Then LangName(j) = "德语"
			If i = 33 Then LangName(j) = "希腊语"
			If i = 34 Then LangName(j) = "格陵兰语"
			If i = 35 Then LangName(j) = "古吉拉特语"
			If i = 36 Then LangName(j) = "豪萨语"
			If i = 37 Then LangName(j) = "希伯来语"
			If i = 38 Then LangName(j) = "印地语"
			If i = 39 Then LangName(j) = "匈牙利语"
			If i = 40 Then LangName(j) = "冰岛语"
			If i = 41 Then LangName(j) = "印度尼西亚语"
			If i = 42 Then LangName(j) = "因纽特语"
			If i = 43 Then LangName(j) = "爱尔兰语"
			If i = 44 Then LangName(j) = "班图语"
			If i = 45 Then LangName(j) = "祖鲁语"
			If i = 46 Then LangName(j) = "意大利语"
			If i = 47 Then LangName(j) = "日语"
			If i = 48 Then LangName(j) = "卡纳达语"
			If i = 49 Then LangName(j) = "克什米尔语"
			If i = 50 Then LangName(j) = "哈萨克语"
			If i = 51 Then LangName(j) = "高棉语"
			If i = 52 Then LangName(j) = "卢旺达语"
			If i = 53 Then LangName(j) = "孔卡尼语"
			If i = 54 Then LangName(j) = "朝鲜语"
			If i = 55 Then LangName(j) = "吉尔吉斯语"
			If i = 56 Then LangName(j) = "吉尔吉斯语 (吉尔吉斯坦)"
			If i = 57 Then LangName(j) = "老挝语"
			If i = 58 Then LangName(j) = "拉脱维亚语"
			If i = 59 Then LangName(j) = "立陶宛语"
			If i = 60 Then LangName(j) = "卢森堡语"
			If i = 61 Then LangName(j) = "马其顿语"
			If i = 62 Then LangName(j) = "马来语"
			If i = 63 Then LangName(j) = "马拉雅拉姆语"
			If i = 64 Then LangName(j) = "马耳他语"
			If i = 65 Then LangName(j) = "毛利语"
			If i = 66 Then LangName(j) = "马拉地语"
			If i = 67 Then LangName(j) = "蒙古语"
			If i = 68 Then LangName(j) = "尼泊尔语"
			If i = 69 Then LangName(j) = "挪威语"
			If i = 70 Then LangName(j) = "挪威语 (博克马尔文)"
			If i = 71 Then LangName(j) = "挪威语 (尼诺斯克文)"
			If i = 72 Then LangName(j) = "奥里雅语"
			If i = 73 Then LangName(j) = "普什图语"
			If i = 74 Then LangName(j) = "波兰语"
			If i = 75 Then LangName(j) = "葡萄牙语"
			If i = 76 Then LangName(j) = "旁遮普语"
			If i = 77 Then LangName(j) = "克丘亚语"
			If i = 78 Then LangName(j) = "罗马尼亚语"
			If i = 79 Then LangName(j) = "俄语"
			If i = 80 Then LangName(j) = "萨米语"
			If i = 81 Then LangName(j) = "梵语"
			If i = 82 Then LangName(j) = "塞尔维亚语"
			If i = 83 Then LangName(j) = "巴索托语"
			If i = 84 Then LangName(j) = "茨瓦纳语"
			If i = 85 Then LangName(j) = "信德语"
			If i = 86 Then LangName(j) = "僧伽罗语"
			If i = 87 Then LangName(j) = "斯洛伐克语"
			If i = 88 Then LangName(j) = "斯洛文尼亚语"
			If i = 89 Then LangName(j) = "西班牙语"
			If i = 90 Then LangName(j) = "斯瓦希里语"
			If i = 91 Then LangName(j) = "瑞典语"
			If i = 92 Then LangName(j) = "叙利亚语"
			If i = 93 Then LangName(j) = "塔吉克语"
			If i = 94 Then LangName(j) = "泰米尔语"
			If i = 95 Then LangName(j) = "鞑靼语"
			If i = 96 Then LangName(j) = "泰卢固语"
			If i = 97 Then LangName(j) = "泰语"
			If i = 98 Then LangName(j) = "藏语"
			If i = 99 Then LangName(j) = "土耳其语"
			If i = 100 Then LangName(j) = "土库曼语"
			If i = 101 Then LangName(j) = "维吾尔语"
			If i = 102 Then LangName(j) = "乌克兰语"
			If i = 103 Then LangName(j) = "乌尔都语"
			If i = 104 Then LangName(j) = "乌兹别克语"
			If i = 105 Then LangName(j) = "越南语"
			If i = 106 Then LangName(j) = "威尔士语"
			If i = 107 Then LangName(j) = "沃洛夫语"
		End If
	Next i

	PslLangCode = "|af|sq|am|ar|hy|As|az|ba|eu|be|BN|bs|br|bg|ca|zh-CN|zh-TW|co|hr|cs|da|nl|" & _
				"en|et|fo|fa|fi|fr|fy|gl|ka|de|el|kl|gu|ha|he|hi|hu|Is|id|iu|ga|xh|zu|it|" & _
				"ja|kn|KS|kk|km|rw|kok|ko|kz|ky|lo|lv|lt|lb|mk|ms|ML|mt|mi|mr|mn|ne|no|nb|" & _
				"nn|Or|ps|pl|pt|pa|qu|ro|ru|se|sa|sr|st|tn|SD|si|sk|sl|es|sw|sv|sy|tg|ta|tt|" & _
				"te|th|bo|tr|tk|ug|uk|ur|uz|vi|cy|wo"

	en2zhCheck = "||||||||||||||||zh-CN|zh-TW||||||||||||||||||||||||||||||ja|||||||ko|||||||" & _
				"||||||||||||||||||||||||||||||||||||||||||||||"

	zh2enCheck = "|af|sq|am|ar|hy|As|az|ba|eu|be|BN|bs|br|bg|ca|||co|hr|cs|da|nl|" & _
				"en|et|fo|fa|fi|fr|fy|gl|ka|de|el|kl|gu|ha|he|hi|hu|Is|id|iu|ga|xh|zu|it|" & _
				"|kn|KS|kk|km|rw|kok||kz|ky|lo|lv|lt|lb|mk|ms|ML|mt|mi|mr|mn|ne|no|nb|" & _
				"nn|Or|ps|pl|pt|pa|qu|ro|ru|se|sa|sr|st|tn|SD|si|sk|sl|es|sw|sv|sy|tg|ta|tt|" & _
				"te|th|bo|tr|tk|ug|uk|ur|uz|vi|cy|wo"

	PslLangCodeList = Split(PslLangCode,LngJoinStr)
	en2zhCheckList = Split(en2zhCheck,LngJoinStr)
	zh2enCheckList = Split(zh2enCheck,LngJoinStr)

	For i = MinNum To MaxNum
		j = i - MinNum
		If DataName = DefaultCheckList(0) Then Code = en2zhCheckList(i)
		If DataName = DefaultCheckList(1) Then Code = zh2enCheckList(i)
		LangPair(j) = LangName(j) & LngJoinStr & PslLangCodeList(i) & LngJoinStr & Code
	Next i
	LangCodeList = LangPair
End Function


'检查帮助
Sub CheckHelp(HelpTip As String)
	If OSLanguage = "0404" Then
	AboutTitle = "闽"
	HelpTitle = "弧"
	HelpTipTitle = "獽倍龄沧ゎ才㎝硉竟浪琩エ栋"
	AboutWindows = " 闽 "
	MainWindows = " 跌怠 "
	SetWindows = " 砞﹚跌怠 "
	TestWindows = " 代刚跌怠 "
	Lines = "-----------------------"
	Sys = "硁砰セ" & Version & vbCrLf & _
			"続ノ╰参Windows XP/2000 ╰参" & vbCrLf & _
			"続ノセ┮Τや穿エ栋矪瞶 Passolo 6.0 のセ" & vbCrLf & _
			"ざ粂ē虏砰いゅ㎝タ砰いゅ (笆侩醚)" & vbCrLf & _
			"舦┮Τ簙て穝" & vbCrLf & _
			"甭舦Α禣硁砰" & vbCrLf & _
			"﹛よhttp://www.hanzify.org" & vbCrLf & _
			"玡秨祇簙て穝Θ gnatix (2007-2008)" & vbCrLf & _
			"秨祇簙て穝Θ wanfu (2009-2010)" & vbCrLf & vbCrLf & vbCrLf
	Ement = "「磅︽吏挂「" & vbCrLf & _
			"============" & vbCrLf & _
			"- や穿エ栋矪瞶 Passolo 6.0 のセゲ惠" & vbCrLf & _
			"- Windows Script Host (WSH) ン (VBS)ゲ惠" & vbCrLf & _
			"- Adodb.Stream ン (VBS)や穿笆穝┮惠" & vbCrLf & _
			"- Microsoft.XMLHTTP ンや穿笆穝┮惠" & vbCrLf & vbCrLf & vbCrLf
	Dec = "「硁砰虏ざ「" & vbCrLf & _
			"============" & vbCrLf & _
			"獽倍龄沧ゎ才㎝硉竟浪琩エ栋琌ノ Passolo 陆亩浪琩エ栋祘ΑウㄣΤ" & vbCrLf & _
			"- 浪琩陆亩い獽倍龄沧ゎ才硉竟㎝" & vbCrLf & _
			"- 浪琩タ浪琩陆亩い獽倍龄沧ゎ才硉竟㎝" & vbCrLf & _
			"- 埃陆亩い獽倍龄" & vbCrLf & _
			"- ず竚璹笆穝" & vbCrLf & vbCrLf & _
			"セ祘Α郎" & vbCrLf & _
			"- 笆エ栋PslAutoAccessKey.bas" & vbCrLf & _
			"  陆亩﹃笆タ岿粇陆亩ノ赣エ栋眤ぃゲ块獽倍龄沧ゎ才硉竟" & vbCrLf & _
			"  ╰参盢沮眤匡砞﹚笆腊眤穝糤㎝ゅ妓獽倍龄沧ゎ才硉竟陆亩沧ゎ才" & vbCrLf & _
			"  ノウ矗蔼陆亩硉搭ぶ陆亩岿粇" & vbCrLf & _
			"  》猔種パ Passolo 赣エ栋砞﹚惠硄筁磅︽浪琩エ栋ㄓ匡㎝砞﹚" & vbCrLf & vbCrLf & _
			"- 浪琩エ栋PSLCheckAccessKeys.bas" & vbCrLf & _
			"  硄筁㊣穝糤 Passolo 匡虫い赣エ栋眤浪琩㎝タ陆亩い獽倍龄沧ゎ才硉竟" & vbCrLf & _
			"  陆亩沧ゎ才ウ临矗ㄑ璹砞﹚㎝砞﹚盎代" & vbCrLf & vbCrLf & _
			"- 虏砰いゅ弧郎AccessKey.txt" & vbCrLf & vbCrLf & vbCrLf
	Setup = "「杆よ猭「" & vbCrLf & _
			"============" & vbCrLf & _
			"- 狦ㄏノ Wanfu  Passolo 簙て杆エ栋舱ン玥蠢传ㄓ郎玥" & vbCrLf & _
			"  (1) 盢秆溃郎狡籹 Passolo ╰参戈Жい﹚竡 Macros 戈Жい" & vbCrLf & _
			"  (2) 笆エ栋秨币 Passolo ㄣ -> エ栋癸杠よ遏盢ウ砞﹚╰参エ栋翴阑跌怠à╰参" & vbCrLf & _
			"    エ栋币ノ匡虫币ノウ" & vbCrLf & _
			"  (3) 浪琩エ栋 Passolo ㄣ -> 璹ㄣ匡虫い穝糤赣郎" & vbCrLf & _
			"- パ笆エ栋礚猭磅︽筁祘い秈︽砞﹚┮叫ㄏノ浪琩エ栋ㄓ璹砞﹚盡" & vbCrLf & _
			"- 叫浪琩叭ゲ硋兵も狡琩祘Α矪瞶岿粇" & vbCrLf & vbCrLf & vbCrLf
	MainUse = "「砞﹚匡「" & vbCrLf & _
			"============" & vbCrLf & _
			"祘Α矗ㄑ箇砞砞﹚赣砞﹚続ノ计薄猵眤 [砞﹚] 秙璹砞﹚" & vbCrLf & _
			"穝糤璹砞﹚眤砞﹚睲虫い匡稱ㄏノ砞﹚" & vbCrLf & _
			"Τ闽璹砞﹚叫秨币砞﹚癸杠よ遏翴阑 [弧] 秙把綷弧い弧" & vbCrLf & vbCrLf & _
			"- 笆エ栋砞﹚" & vbCrLf & _
			"  匡砞﹚盢ノ笆エ栋叫猔種纗ぃ礛盢ㄏノ匡玡砞﹚" & vbCrLf & vbCrLf & _
			"- 浪琩エ栋砞﹚" & vbCrLf & _
			"  匡砞﹚盢ノ浪琩エ栋璶Ωㄏノ匡砞﹚惠璶纗" & vbCrLf & vbCrLf & _
			"- 笆エ栋㎝浪琩エ栋" & vbCrLf & _
			"  匡赣匡兜盢笆ㄏ笆エ栋砞﹚籔浪琩エ栋砞﹚璓" & vbCrLf & vbCrLf & _
			"- 笆匡" & vbCrLf & _
			"  匡拒赣匡兜盢沮砞﹚い続ノ粂ē睲虫笆匡籔陆亩睲虫ヘ夹粂ē才砞﹚" & vbCrLf & _
			"  》猔種赣匡兜度癸獽倍龄沧ゎ才㎝硉竟浪琩エ栋Τ" & vbCrLf & _
			"  璶盢砞﹚籔ヘ玡陆亩睲虫ヘ夹粂ē才 [砞﹚] 秙﹃矪瞶続ノ粂ēい" & vbCrLf & _
			"  穝糤癸莱粂ē" & vbCrLf & _
			"  笆エ栋祘Α盢笆 - 匡 - 箇砞抖匡癸莱砞﹚" & vbCrLf & vbCrLf & _
			"「浪琩夹癘「" & vbCrLf & _
			"============" & vbCrLf & _
			"赣硄筁癘魁﹃浪琩癟沮赣癟癸岿粇﹃秈︽浪琩碩矗蔼浪琩硉" & vbCrLf & _
			"Τ 4 匡兜ㄑ匡" & vbCrLf & vbCrLf & _
			"- ┛菠セ" & vbCrLf & _
			"  盢ぃσ納エ栋祘Αセ度沮ㄤウ癘魁癸岿粇﹃秈︽浪琩" & vbCrLf & vbCrLf & _
			"- ┛菠砞﹚" & vbCrLf & _
			"  盢ぃσ納砞﹚琌度沮ㄤウ癘魁癸岿粇﹃秈︽浪琩" & vbCrLf & vbCrLf & _
			"- ┛菠ら戳" & vbCrLf & _
			"  盢ぃσ納浪琩ら戳㎝陆亩ら戳度沮ㄤウ癘魁癸岿粇﹃秈︽浪琩" & vbCrLf & vbCrLf & _
			"- 场┛菠" & vbCrLf & _
			"  盢ぃσ納ヴ浪琩癘魁τ癸┮Τ﹃秈︽浪琩" & vbCrLf & vbCrLf & _
			"》猔種浪琩夹癘代刚砞﹚礚瞷Τ框簗代刚挡狦" & vbCrLf & _
			"狦跑砞﹚ず甧τぃ跑砞﹚嘿杠叫匡拒场┛菠┪埃浪琩夹癘匡兜" & vbCrLf & vbCrLf & _
			"「砞﹚匡「" & vbCrLf & _
			"============" & vbCrLf & _
			"Τ 3 匡兜ㄑ匡" & vbCrLf & vbCrLf & _
			"- 度浪琩" & vbCrLf & _
			"  癸陆亩秈︽浪琩τぃタ岿粇陆亩" & vbCrLf & vbCrLf & _
			"- 浪琩タ" & vbCrLf & _
			"  癸陆亩秈︽浪琩笆タ岿粇陆亩" & vbCrLf & vbCrLf & _
			"- 埃獽倍龄" & vbCrLf & _
			"  埃陆亩い瞷Τ獽倍龄" & vbCrLf & vbCrLf & _
			"「﹃摸「" & vbCrLf & _
			"============" & vbCrLf & _
			"矗ㄑ场匡虫癸杠よ遏﹃硉竟セㄤ度匡拒单匡兜" & vbCrLf & vbCrLf & _
			"- 狦匡场玥ㄤ虫兜盢砆笆匡" & vbCrLf & _
			"- 狦匡虫兜玥场匡兜盢砆笆匡" & vbCrLf & _
			"- 虫兜匡ㄤい匡度匡拒ㄤА砆笆匡" & vbCrLf & vbCrLf & _
			"「﹃ず甧「" & vbCrLf & _
			"============" & vbCrLf & _
			"矗ㄑ场獽倍龄沧ゎ才硉竟 4 匡兜" & vbCrLf & vbCrLf & _
			"- 狦匡场玥ㄤ虫兜盢砆笆匡" & vbCrLf & _
			"- 狦匡虫兜玥场匡兜盢砆笆匡" & vbCrLf & _
			"- 虫兜匡" & vbCrLf & vbCrLf & _
			"「ㄤ匡兜「" & vbCrLf & _
			"============" & vbCrLf & _
			"- ぃ跑﹍陆亩篈" & vbCrLf & _
			"  匡拒赣兜盢浪琩浪琩タ埃獽倍龄ぃ跑﹃﹍陆亩篈玥" & vbCrLf & _
			"  盢跑礚岿粇┪礚跑﹃陆亩篈喷靡篈Τ岿粇┪跑﹃陆亩篈" & vbCrLf & _
			"  狡糵篈獽眤泊碞笵ㄇ﹃Τ岿粇┪砆跑" & vbCrLf & vbCrLf & _
			"- ぃミ┪埃浪琩夹癘" & vbCrLf & _
			"  匡拒赣兜盢ぃ盡い纗浪琩夹癘癟狦浪琩夹癘癟盢砆埃" & vbCrLf & _
			"  》猔種匡拒赣兜浪琩夹癘よ遏い场┛菠兜盢砆匡拒" & vbCrLf & vbCrLf & _
			"- 膥尿笆纗┮Τ匡" & vbCrLf & _
			"  匡拒赣兜盢 [膥尿] 秙笆纗┮Τ匡Ω磅︽盢弄纗匡" & vbCrLf & _
			"  》猔種狦笆エ栋砞﹚砆跑╰参盢笆匡拒赣匡兜ㄏ匡拒笆エ栋砞﹚ネ" & vbCrLf & vbCrLf & _
			"- 蠢传疭﹚じ" & vbCrLf & _
  			"  浪琩タ筁祘いㄏノ砞﹚い﹚竡璶笆蠢传じ蠢传﹃い疭﹚じ" & vbCrLf & vbCrLf & _
			"「ㄤ「" & vbCrLf & _
			"============" & vbCrLf & _
			"- 闽" & vbCrLf & _
			"  翴阑赣秙盢闽癸杠よ遏陪ボ祘Αざ残磅︽吏挂秨祇坝の舦单癟" & vbCrLf& vbCrLf & _
			"- 弧" & vbCrLf & _
			"  翴阑赣秙盢紆ヘ玡跌怠弧癟" & vbCrLf& vbCrLf & _
			"- 纗┮Τ匡" & vbCrLf & _
			"  赣秙跑砞﹚τぃ秈︽浪琩ㄏノ" & vbCrLf & _
			"  狦ヴ匡兜砆跑赣匡兜盢笆跑ノ篈玥盢笆跑ぃノ篈" & vbCrLf & vbCrLf & _
			"- 絋﹚" & vbCrLf & _
			"  翴阑赣秙盢闽超癸杠よ遏匡拒匡兜秈︽﹃陆亩" & vbCrLf& vbCrLf & _
			"- " & vbCrLf & _
			"  翴阑赣秙盢挡祘Α" & vbCrLf& vbCrLf & vbCrLf
	SetUse ="「砞﹚睲虫「" & vbCrLf & _
			"============" & vbCrLf & _
			"- 匡砞﹚" & vbCrLf & _
			"  璶匡砞﹚翴阑砞﹚睲虫" & vbCrLf & vbCrLf & _
			"- 砞﹚砞﹚纔" & vbCrLf & _
			"  砞﹚纔ノ膀砞﹚続ノ粂ē笆匡砞﹚" & vbCrLf & _
			"  》猔種Τ砞﹚続ノ粂ē惠璶砞﹚ㄤ纔" & vbCrLf & _
			"  続ノ粂ē砞﹚い玡砞﹚砆纔匡ㄏノ" & vbCrLf & _
			"  璶砞﹚砞﹚纔翴阑娩 [...] 秙" & vbCrLf & vbCrLf & _
			"- 穝糤砞﹚" & vbCrLf & _
			"  璶穝糤砞﹚翴阑 [穝糤] 秙紆癸杠よ遏い块嘿" & vbCrLf & vbCrLf & _
			"- 跑砞﹚" & vbCrLf & _
			"  璶跑砞﹚嘿叫匡砞﹚睲虫い璶э砞﹚礛翴阑 [跑] 秙" & vbCrLf & vbCrLf & _
			"- 埃砞﹚" & vbCrLf & _
			"  璶埃砞﹚叫匡砞﹚睲虫い璶埃砞﹚礛翴阑 [埃] 秙" & vbCrLf & vbCrLf & _
			"穝糤砞﹚盢睲虫い陪ボ穝砞﹚砞﹚ず甧盢陪ボ" & vbCrLf & _
			"跑砞﹚盢睲虫い陪ボэ砞﹚砞﹚ず甧い砞﹚ぃ跑" & vbCrLf & _
			"埃砞﹚盢睲虫い陪ボ箇砞砞﹚砞﹚ず甧盢陪ボ箇砞砞﹚" & vbCrLf & vbCrLf & _
			"「纗摸「" & vbCrLf & _
			"============" & vbCrLf & _
			"- 郎" & vbCrLf & _
			"  砞﹚盢郎Α纗エ栋┮戈Ж Data 戈Жい" & vbCrLf & vbCrLf & _
			"- 爹" & vbCrLf & _
			"  砞﹚盢砆纗爹い HKCU\Software\VB and VBA Program Settings\AccessKey 兜" & vbCrLf & vbCrLf & _
			"- 蹲砞﹚" & vbCrLf & _
			"  す砛眖ㄤ砞﹚郎い蹲砞﹚蹲侣砞﹚盢砆笆ど瞷Τ砞﹚睲虫いΤ砞﹚盢砆" & vbCrLf & _
			"  跑⊿Τ砞﹚盢砆穝糤" & vbCrLf & vbCrLf & _
			"- 蹲砞﹚" & vbCrLf & _
			"  す砛蹲┮Τ砞﹚ゅ郎獽ユ传┪锣簿砞﹚" & vbCrLf & vbCrLf & _
			"》猔種ち传纗摸盢笆埃Τ竚い砞﹚ず甧" & vbCrLf & vbCrLf & _
			"「砞﹚ず甧「" & vbCrLf & _
			"============" & vbCrLf & _
			"<獽倍龄>" & vbCrLf & _
			"  - 璶逼埃 & 才腹獶獽倍龄舱" & vbCrLf & _
			"    獽倍龄 & 篨夹才Τㄇ﹃瘤礛赣才腹ぃ琌獽倍龄惠璶逼埃ウ块硂ㄇ" & vbCrLf & _
			"    璶逼埃 & 才腹獶獽倍龄舱" & vbCrLf & vbCrLf & _
			"  - ﹃だ澄ノ篨夹才" & vbCrLf & _
			"    硂ㄇじノだ澄Τ獽倍龄沧ゎ才┪硉竟﹃獽浪琩﹃い┮Τ獽倍龄" & vbCrLf & _
			"    沧ゎ才┪硉竟玥矪瞶﹃程场獽倍龄沧ゎ才┪硉竟" & vbCrLf & vbCrLf & _
			"  - 璶浪琩獽倍龄玡珹腹" & vbCrLf & _
			"    箇砞獽倍龄玡珹腹 ()﹚獽倍龄玡珹腹常盢砆蠢传箇砞珹腹" & vbCrLf & vbCrLf & _
			"  - 璶玂痙獶獽倍龄玡Θ癸じ" & vbCrLf & _
			"    狦ゅ㎝陆亩い常硂ㄇじ盢砆玂痙玥盢砆粄琌獽倍龄盢穝糤獽倍龄簿" & vbCrLf & _
			"    ﹃程" & vbCrLf & vbCrLf & _
			"  - ゅ陪ボ盿珹腹獽倍龄 (硄盽ノㄈ瑆粂ē)" & vbCrLf & _
			"    硄盽ㄈ瑆粂ēいゅらゅ单硁砰いㄏノ (&X) Α獽倍龄盢ㄤ竚﹃挡Ю沧" & vbCrLf & _
			"    ゎ才┪硉竟玡" & vbCrLf & _
			"    匡拒赣匡兜盢浪琩Τ獽倍龄陆亩﹃い獽倍龄琌才篋ㄒ狦ぃ才盢砆笆" & vbCrLf & _
			"    跑竚" & vbCrLf & vbCrLf & _
			"<沧ゎ才>" & vbCrLf & _
			"  - 璶浪琩沧ゎ才" & vbCrLf & _
			"    陆亩い沧ゎ才㎝ゅぃ盢砆跑ゅい沧ゎ才琌才璶笆蠢传沧ゎ才" & vbCrLf & _
			"    癸い沧ゎ才埃" & vbCrLf & vbCrLf & _
			"    赣逆や穿窾ノじ琌ぃ琌家絢τ琌弘絋ㄒA*C ぃ才 XAYYCZ才 AXYYC" & vbCrLf & _
			"    璶才 XAYYCZ莱赣 *A*C* ┪ *A??C*" & vbCrLf & vbCrLf & _
			"    》猔種パ Sax Basic ま篮拜肈獶璣ゅダぇ丁 ? 窾ノじぃ砆や穿" & vbCrLf & _
			"    ㄒ秨币??郎ぃ才秨币ㄏノ郎" & vbCrLf & vbCrLf & _
			"  - 璶玂痙沧ゎ才舱" & vbCrLf & _
			"    ┮Τ砆赣舱い璶浪琩沧ゎ才盢砆玂痙碞琌硂ㄇ沧ゎ才ぃ砆粄琌沧ゎ才" & vbCrLf & vbCrLf & _
			"    赣逆や穿窾ノじ琌ぃ琌家絢τ琌弘絋ㄒA*C ぃ才 XAYYCZ才 AXYYC" & vbCrLf & _
			"    璶才 XAYYCZ莱赣 *A*C* ┪ *A??C*" & vbCrLf & vbCrLf & _
			"    》猔種パ Sax Basic ま篮拜肈獶璣ゅダぇ丁 ? 窾ノじぃ砆や穿" & vbCrLf & _
			"    ㄒ秨币??郎ぃ才秨币ㄏノ郎" & vbCrLf & vbCrLf & _
			"  - 璶笆蠢传沧ゎ才癸" & vbCrLf & _
			"    才沧ゎ才癸い玡じ沧ゎ才常盢砆蠢传Θ沧ゎ才癸いじ沧ゎ才" & vbCrLf & _
			"    ノ兜笆陆亩┪タㄇ沧ゎ才" & vbCrLf & vbCrLf & _
			"<硉竟>" & vbCrLf & _
			"  - 璶浪琩硉竟篨夹才" & vbCrLf & _
			"    硉竟硄盽 \t 篨夹才 (Τㄒ)狦﹃い硂ㄇじ盢砆粄硉竟" & vbCrLf & _
			"    惠璶沮璶浪琩硉竟じ秈˙耞" & vbCrLf & vbCrLf & _
			"  - 璶浪琩硉竟じ" & vbCrLf & _
			"    硉竟篨夹才﹃い狦篨夹才じ才赣逆じ盢砆侩醚硉竟" & vbCrLf & _
			"    じ舱τΘ硉竟す砛ㄤいぃ才" & vbCrLf & vbCrLf & _
			"    赣逆や穿窾ノじ琌ぃ琌家絢τ琌弘絋ㄒA*C ぃ才 XAYYCZ才 AXYYC" & vbCrLf & _
			"    璶才 XAYYCZ莱赣 *A*C* ┪  *A??C*" & vbCrLf & vbCrLf & _
			"    》猔種パ Sax Basic ま篮拜肈獶璣ゅダぇ丁 ? 窾ノじぃ砆や穿" & vbCrLf & _
			"    ㄒ秨币??郎ぃ才秨币ㄏノ郎" & vbCrLf & vbCrLf & _
			"  - 璶玂痙硉竟じ" & vbCrLf & _
			"    才硂ㄇじ硉竟盢砆玂痙玥盢砆蠢传ノ兜玂痙琘ㄇ硉竟陆亩" & vbCrLf & vbCrLf & _
			"<じ蠢传>" & vbCrLf & _
			"  ﹃い–蠢传じ癸い|玡じ盢砆蠢传Θ|じ" & vbCrLf & vbCrLf & _
			"  - 璶笆蠢传じ" & vbCrLf & _
			"    ﹚竡浪琩タ筁祘い璶砆蠢传じの蠢传じ" & vbCrLf & vbCrLf & _
			"  》猔種蠢传跋だ糶" & vbCrLf & _
			"  狦璶奔硂ㄇじ盢|じ竚" & vbCrLf & vbCrLf & _
			"<続ノ粂ē>" & vbCrLf & _
			"  硂柑続ノ粂ē琌陆亩睲虫ヘ夹粂ēウノ沮陆亩睲虫ヘ夹粂ē笆匡癸莱砞﹚" & vbCrLf & _
			"  笆匡" & vbCrLf & vbCrLf & _
			"  - 穝糤" & vbCrLf & _
			"    璶穝糤続ノ粂ē匡ノ粂ē睲虫い粂ē礛翴阑 [穝糤] 秙" & vbCrLf & _
			"    翴阑赣秙ノ粂ē睲虫い匡拒粂ē盢簿笆続ノ粂ē睲虫い" & vbCrLf & vbCrLf & _
			"  - 场穝糤" & vbCrLf & _
			"    翴阑赣秙ノ粂ē睲虫い┮Τ粂ē盢场簿笆続ノ粂ē睲虫い" & vbCrLf & vbCrLf & _
			"  - 埃" & vbCrLf & _
			"    璶埃続ノ粂ē匡続ノ粂ē睲虫い粂ē礛翴阑 [埃] 秙" & vbCrLf & _
			"    翴阑赣秙続ノ粂ē睲虫い匡拒粂ē盢簿笆ノ粂ē睲虫い" & vbCrLf & vbCrLf & _
			"  - 场埃" & vbCrLf & _
			"    翴阑赣秙続ノ粂ē睲虫い┮Τ粂ē盢场簿笆ノ粂ē睲虫い" & vbCrLf & vbCrLf & _
			"  - 糤ノ粂ē" & vbCrLf & _
			"    翴阑赣秙盢紆块粂ē嘿㎝絏癸杠よ遏絋﹚盢穝糤ノ粂ē睲虫い" & vbCrLf & vbCrLf & _
			"  - 絪胯ノ粂ē" & vbCrLf & _
			"    璶絪胯ノ粂ē匡ノ粂ē睲虫い粂ē礛翴阑 [絪胯ノ粂ē] 秙" & vbCrLf & _
			"    翴阑赣秙盢紆絪胯粂ē嘿㎝絏癸杠よ遏絋﹚盢эノ粂ē睲虫い匡拒粂ē" & vbCrLf & vbCrLf & _
			"  - 埃ノ粂ē" & vbCrLf & _
			"    璶埃ノ粂ē匡ノ粂ē睲虫い璶埃粂ē礛翴阑 [埃ノ粂ē] 秙" & vbCrLf & vbCrLf & _
			"  - 糤続ノ粂ē" & vbCrLf & _
			"    翴阑赣秙盢紆块粂ē嘿㎝絏癸杠よ遏絋﹚盢穝糤続ノ粂ē睲虫い" & vbCrLf & vbCrLf & _
			"  - 絪胯続ノ粂ē" & vbCrLf & _
			"    璶絪胯続ノ粂ē匡続ノ粂ē睲虫い粂ē礛翴阑 [絪胯続ノ粂ē] 秙" & vbCrLf & _
			"    翴阑赣秙盢紆絪胯粂ē嘿㎝絏癸杠よ遏絋﹚盢э続ノ粂ē睲虫い匡拒粂ē" & vbCrLf & vbCrLf & _
			"  - 埃続ノ粂ē" & vbCrLf & _
			"    璶埃続ノ粂ē匡続ノ粂ē睲虫い璶埃粂ē礛翴阑 [埃続ノ粂ē] 秙" & vbCrLf & vbCrLf & _
			"  》猔種穝糤絪胯粂ē度ノ Passolo ゼㄓセ穝糤や穿粂ē" & vbCrLf & _
			"  粂ē絏叫㎝ Passolo  ISO 396-1 絏玂璓珹糶" & vbCrLf & vbCrLf & _
			"「ㄤ「" & vbCrLf & _
			"============" & vbCrLf & _
			"- 弧" & vbCrLf & _
			"  翴阑赣秙盢紆ヘ玡跌怠弧癟" & vbCrLf & vbCrLf & _
			"- 弄" & vbCrLf & _
			"  翴阑赣秙盢沮匡拒砞﹚ぃ紆匡虫:" & vbCrLf & _
			"  (1) 箇砞" & vbCrLf & _
			"      弄箇砞砞﹚陪ボ砞﹚ず甧い" & vbCrLf & _
			"      》度讽匡拒砞﹚╰参箇砞砞﹚陪ボ赣匡虫" & vbCrLf & vbCrLf & _
			"  (2) " & vbCrLf & _
			"      弄匡拒砞﹚﹍陪ボ砞﹚ず甧い" & vbCrLf & _
			"      》度讽匡拒砞﹚﹍獶陪ボ赣匡虫" & vbCrLf & vbCrLf & _
			"  (3) 把酚" & vbCrLf & _
			"      弄匡拒把酚砞﹚陪ボ砞﹚ず甧い" & vbCrLf & _
			"      》赣匡虫陪ボ埃匡拒砞﹚┮Τ砞﹚睲虫" & vbCrLf & vbCrLf & _
			"- 睲" & vbCrLf & _
			"  翴阑赣秙盢睲瞷Τ砞﹚场よ獽穝块砞﹚" & vbCrLf & vbCrLf & _
			"- 代刚" & vbCrLf & _
			"  翴阑赣秙盢紆代刚癸杠よ遏獽浪琩砞﹚タ絋┦" & vbCrLf & vbCrLf & _
			"- 絋﹚" & vbCrLf & _
			"  翴阑赣秙盢纗砞﹚跌怠いヴ跑挡砞﹚跌怠跌怠" & vbCrLf & _
			"  祘Α盢ㄏノ跑砞﹚" & vbCrLf & vbCrLf & _
			"- " & vbCrLf & _
			"  翴阑赣秙ぃ纗砞﹚跌怠いヴ跑挡砞﹚跌怠跌怠" & vbCrLf & _
			"  祘Α盢ㄏノㄓ砞﹚" & vbCrLf & vbCrLf & vbCrLf
	TestUse ="「砞﹚嘿「" & vbCrLf & _
			"============" & vbCrLf & _
			"- 璶代刚砞﹚嘿璶匡砞﹚翴阑砞﹚睲虫" & vbCrLf & vbCrLf & _
			"「陆亩睲虫「" & vbCrLf & _
			"============" & vbCrLf & _
			"- 赣睲虫盢陪ボ盡い┮Τ陆亩睲虫叫匡籔眤璹砞﹚才陆亩睲虫秈︽代刚" & vbCrLf & vbCrLf & _
			"「笆蠢传じ「" & vbCrLf & _
			"============" & vbCrLf & _
			"- 浪琩笆蠢传﹃い才砞﹚い┮﹚竡蠢传じ" & vbCrLf & vbCrLf & _
			"「弄︽计「" & vbCrLf & _
			"============" & vbCrLf & _
			"- ボ璶陪ボ岿粇陆亩﹃计某ぃ璶块び﹃耕单丁筁" & vbCrLf & vbCrLf & _
			"「ず甧「" & vbCrLf & _
			"============" & vbCrLf & _
			"- ﹚浪琩ず甧﹃ノ赣兜Τ皐癸┦代刚е代刚丁" & vbCrLf & _
			"- 赣逆や穿家絢窾ノじㄒA*C 才 XAYYCZ" & vbCrLf & vbCrLf & _
			"》猔種パ Sax Basic ま篮拜肈獶璣ゅダぇ丁 ? 窾ノじぃ砆や穿" & vbCrLf & _
			"ㄒ秨币??郎ぃ才秨币ㄏノ郎" & vbCrLf & vbCrLf & _
			"「﹃ず甧「" & vbCrLf & _
			"============" & vbCrLf & _
			"矗ㄑ场獽倍龄沧ゎ才硉竟 4 匡兜" & vbCrLf & vbCrLf & _
			"- 狦匡场玥ㄤ虫兜盢砆笆匡" & vbCrLf & _
			"- 狦匡虫兜玥场匡兜盢砆笆匡" & vbCrLf & _
			"- 虫兜匡" & vbCrLf & vbCrLf & _
			"「ㄤ「" & vbCrLf & _
			"============" & vbCrLf & _
			"- 弧" & vbCrLf & _
			"  翴阑赣秙盢紆ヘ玡跌怠弧癟" & vbCrLf & vbCrLf & _
			"- 代刚" & vbCrLf & _
			"  翴阑赣秙盢酚匡拒兵ン秈︽代刚" & vbCrLf & vbCrLf & _
			"- 睲" & vbCrLf & _
			"  翴阑赣秙盢睲瞷Τ代刚挡狦" & vbCrLf & vbCrLf & _
			"- " & vbCrLf & _
			"  翴阑赣秙盢挡代刚祘Α砞﹚跌怠" & vbCrLf & vbCrLf & vbCrLf
	CopyRight = "「舦「" & vbCrLf & _
			"============" & vbCrLf & _
			"- 硁砰舦耴秨祇㎝э┮Τヴ禣ㄏノэ狡籹床セ硁砰" & vbCrLf & _
			"- э床セ硁砰ゲ斗繦セ弧郎爹硁砰﹍秨祇のэ" & vbCrLf & _
			"- ゼ竒秨祇㎝э種ヴ舱麓┪ぃ眔ノ坝穨硁砰坝穨┪琌ㄤウ犁┦笆" & vbCrLf & _
			"- 癸ㄏノセ硁砰﹍セのㄏノ竒э獶﹍セ┮硑Θ穕ア㎝穕甡秨祇ぃ" & vbCrLf & _
			"  ┯踞ヴ砫ヴ" & vbCrLf & _
			"- パ禣硁砰秨祇㎝э⊿Τ竡叭矗ㄑ硁砰м砃や穿礚竡叭э秈┪穝セ" & vbCrLf & _
			"- 舧タ岿粇矗э秈種ǎΤ岿粇┪某叫肚癳: z_shangyi@163.com" & vbCrLf & vbCrLf & vbCrLf
	Thank = "「璓谅「" & vbCrLf & _
			"============" & vbCrLf & _
			"- セ硁砰э筁祘い眔簙て穝穦代刚ボ癑み稰谅" & vbCrLf & _
			"- 稰谅芖 Heaven ネ矗タ砰ノ粂э種ǎ" & vbCrLf & vbCrLf & vbCrLf
	Contact = "「籔и羛么「" & vbCrLf & _
			"============" & vbCrLf & _
			"wanfuz_shangyi@163.com" & vbCrLf & vbCrLf & vbCrLf
	Logs = "稰谅や眤や琌и程笆舧ㄏノи籹硁砰" & vbCrLf & vbCrLf & _
			"**************************************************" & vbCrLf & _
			"惠璶穝簙て叫砐:" & vbCrLf & _
			"簙て穝 -- http://www.hanzify.org" & vbCrLf & _
			"簙て穝阶韭 -- http://bbs.hanzify.org" & vbCrLf & _
			"**************************************************" & vbCrLf
	Else
	AboutTitle = "关于"
	HelpTitle = "帮助"
	HelpTipTitle = "快捷键、终止符和加速器检查宏"
	AboutWindows = " 关于 "
	MainWindows = " 主窗口 "
	SetWindows = " 配置窗口 "
	TestWindows = " 测试窗口 "
	Lines = "-----------------------"
	Sys = "软件版本：" & Version & vbCrLf & _
			"适用系统：Windows XP/2000 以上系统" & vbCrLf & _
			"适用版本：所有支持宏处理的 Passolo 6.0 及以上版本" & vbCrLf & _
			"界面语言：简体中文和繁体中文 (自动识别)" & vbCrLf & _
			"版权所有：汉化新世纪" & vbCrLf & _
			"授权形式：免费软件" & vbCrLf & _
			"官方主页：http://www.hanzify.org" & vbCrLf & _
			"前开发者：汉化新世纪成员 gnatix (2007-2008)" & vbCrLf & _
			"后开发者：汉化新世纪成员 wanfu (2009-2010)" & vbCrLf & vbCrLf & vbCrLf
	Ement = "☆运行环境☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 支持宏处理的 Passolo 6.0 及以上版本，必需" & vbCrLf & _
			"- Windows Script Host (WSH) 对象 (VBS)，必需" & vbCrLf & _
			"- Adodb.Stream 对象 (VBS)，支持自动更新所需" & vbCrLf & _
			"- Microsoft.XMLHTTP 对象，支持自动更新所需" & vbCrLf & vbCrLf & vbCrLf
	Dec = "☆软件简介☆" & vbCrLf & _
			"============" & vbCrLf & _
			"快捷键、终止符和加速器检查宏是一个用于 Passolo 翻译检查的宏程序。它具有以下功能：" & vbCrLf & _
			"- 检查翻译中快捷键、终止符、加速器和空格" & vbCrLf & _
			"- 检查并修正检查翻译中快捷键、终止符、加速器和空格" & vbCrLf & _
			"- 删除翻译中的快捷键" & vbCrLf & _
			"- 内置可自定义的自动更新功能" & vbCrLf & vbCrLf & _
			"本程序包含下列文件：" & vbCrLf & _
			"- 自动宏：PslAutoAccessKey.bas" & vbCrLf & _
			"  在翻译字串时，自动更正错误的翻译。利用该宏，您可以不必输入快捷键、终止符、加速器，" & vbCrLf & _
			"  系统将根据您选择的配置自动帮您添加和原文一样的快捷键、终止符、加速器，并翻译终止符。" & vbCrLf & _
			"  利用它可以提高翻译速度，并减少翻译错误。" & vbCrLf & _
			"  ◎注意：由于 Passolo 的限制，该宏的配置需通过运行检查宏来选择和设置。" & vbCrLf & vbCrLf & _
			"- 检查宏：PSLCheckAccessKeys.bas" & vbCrLf & _
			"  通过调用添加到 Passolo 菜单中的该宏，帮助您检查和修正翻译中的快捷键、终止符、加速器，" & vbCrLf & _
			"  并翻译终止符。此外，它还提供自定义配置和配置检测功能。" & vbCrLf & vbCrLf & _
			"- 简体中文说明文件：AccessKey.txt" & vbCrLf & vbCrLf & vbCrLf
	Setup = "☆安装方法☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 如果使用了 Wanfu 的 Passolo 汉化版，并安装了附加宏组件，则替换原来的文件即可，否则：" & vbCrLf & _
			"  (1) 将解压后的文件复制到 Passolo 系统文件夹中定义的 Macros 文件夹中" & vbCrLf & _
			"  (2) 自动宏：打开 Passolo 的工具 -> 宏对话框，将它设置为系统宏并单击主窗口的右下角的系统" & vbCrLf & _
			"  　  宏激活菜单激活它" & vbCrLf & _
			"  (3) 检查宏：在 Passolo 的工具 -> 自定义工具菜单中添加该文件" & vbCrLf & _
			"- 由于自动宏无法在运行过程中进行配置，所以请使用检查宏来自定义配置方案。" & vbCrLf & _
			"- 请检查后务必再逐条手工复查，以免程序处理错误。" & vbCrLf & vbCrLf & vbCrLf
	MainUse = "☆配置选择☆" & vbCrLf & _
			"============" & vbCrLf & _
			"程序提供了默认的配置，该配置可以适用于大多数情况。您也可以按 [设置] 按钮自定义配置。" & vbCrLf & _
			"添加自定义配置后，您可以在配置列表中选择想使用的配置。" & vbCrLf & _
			"有关自定义配置，请打开配置对话框，单击 [帮助] 按钮，参阅帮助中的说明。" & vbCrLf & vbCrLf & _
			"- 自动宏配置" & vbCrLf & _
			"  选择的配置将用于自动宏。请注意保存，不然将使用选择前的配置。" & vbCrLf & vbCrLf & _
			"- 检查宏配置" & vbCrLf & _
			"  选择的配置将用于检查宏。要在下次使用选择的配置，需要保存。" & vbCrLf & vbCrLf & _
			"- 自动宏和检查宏相同" & vbCrLf & _
			"  选择该选项时，将自动使自动宏的配置与检查宏的配置一致。" & vbCrLf & vbCrLf & _
			"- 自动选择" & vbCrLf & _
			"  选定该选项时将根据配置中的适用语言列表自动选择与翻译列表目标语言匹配的配置。" & vbCrLf & _
			"  ◎注意：该选项仅对快捷键、终止符和加速器检查宏有效。" & vbCrLf & _
			"  　　　　要将配置与当前翻译列表的目标语言匹配，按 [设置] 按钮，在字串处理的适用语言中" & vbCrLf & _
			"  　　　　添加相应的语言。" & vbCrLf & _
			"  　　　　自动宏程序将按“自动 - 自选 - 默认”顺序选择相应的配置。" & vbCrLf & vbCrLf & _
			"☆检查标记☆" & vbCrLf & _
			"============" & vbCrLf & _
			"该功能通过记录字串检查信息，并根据该信息只对错误字串进行检查，可大幅提高再检查速度。" & vbCrLf & _
			"有以下 4 个选项可供选择：" & vbCrLf & vbCrLf & _
			"- 忽略版本" & vbCrLf & _
			"  将不考虑宏程序的版本，仅根据其它记录对错误字串进行检查。" & vbCrLf & vbCrLf & _
			"- 忽略配置" & vbCrLf & _
			"  将不考虑配置是否相同，仅根据其它记录对错误字串进行检查。" & vbCrLf & vbCrLf & _
			"- 忽略日期" & vbCrLf & _
			"  将不考虑检查日期和翻译日期，仅根据其它记录对错误字串进行检查。" & vbCrLf & vbCrLf & _
			"- 全部忽略" & vbCrLf & _
			"  将不考虑任何检查记录，而对所有字串进行检查。" & vbCrLf & vbCrLf & _
			"◎注意：检查标记功能在测试配置时无效，以免可能出现有遗漏的测试结果。" & vbCrLf & _
			"　　　　如果更改配置内容而不更改配置名称的话，请选定全部忽略或删除检查标记选项。" & vbCrLf & vbCrLf & _
			"☆配置选择☆" & vbCrLf & _
			"============" & vbCrLf & _
			"有以下 3 个选项可供选择：" & vbCrLf & vbCrLf & _
			"- 仅检查" & vbCrLf & _
			"  只对翻译进行检查，而不修正错误的翻译。" & vbCrLf & vbCrLf & _
			"- 检查并修正" & vbCrLf & _
			"  对翻译进行检查，并自动修正错误的翻译。" & vbCrLf & vbCrLf & _
			"- 删除快捷键" & vbCrLf & _
			"  删除翻译中现有的快捷键。" & vbCrLf & vbCrLf & _
			"☆字串类型☆" & vbCrLf & _
			"============" & vbCrLf & _
			"提供了全部、菜单、对话框、字符串表、加速器、版本、其他、仅选定等选项。" & vbCrLf & vbCrLf & _
			"- 如果选择全部，则其他单项将被自动取消选择。" & vbCrLf & _
			"- 如果选择单项，则全部选项将被自动取消选择。" & vbCrLf & _
			"- 单项可以多选。其中选择仅选定时，其他均被自动取消选择。" & vbCrLf & vbCrLf & _
			"☆字串内容☆" & vbCrLf & _
			"============" & vbCrLf & _
			"提供了全部、快捷键、终止符、加速器 4 个选项。" & vbCrLf & vbCrLf & _
			"- 如果选择全部，则其他单项将被自动取消选择。" & vbCrLf & _
			"- 如果选择单项，则全部选项将被自动取消选择。" & vbCrLf & _
			"- 单项可以多选。" & vbCrLf & vbCrLf & _
			"☆其他选项☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 不更改原始翻译状态" & vbCrLf & _
			"  选定该项时，将在检查、检查并修正、删除快捷键时不更改字串的原始翻译状态。否则，" & vbCrLf & _
			"  将更改无错误或无更改字串的翻译状态为已验证状态，有错误或已更改字串的翻译状态" & vbCrLf & _
			"  为待复审状态，以便您一眼就可以知道哪些字串有错误或已被更改。" & vbCrLf & vbCrLf & _
			"- 不创建或删除检查标记" & vbCrLf & _
			"  选定该项时，将不在方案中保存检查标记信息，如果已存在检查标记信息，将被删除。" & vbCrLf & _
			"  ◎注意：选定该项时，检查标记框中的全部忽略项将被选定。" & vbCrLf & vbCrLf & _
			"- 继续时自动保存所有选择" & vbCrLf & _
			"  选定该项时，将在按 [继续] 按钮时自动保存所有选择，下次运行时将读入保存的选择。" & vbCrLf & _
			"  ◎注意：如果自动宏配置被更改，系统将自动选定该选项，以使选定的自动宏配置生效。" & vbCrLf & vbCrLf & _
			"- 替换特定字符" & vbCrLf & _
  			"  在检查并修正过程中使用配置中定义的要自动替换的字符，替换字串中特定的字符。" & vbCrLf & vbCrLf & _
			"☆其他功能☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 关于" & vbCrLf & _
			"  单击该按钮，将关于对话框。并显示程序介绍、运行环境、开发商及版权等信息。" & vbCrLf& vbCrLf & _
			"- 帮助" & vbCrLf & _
			"  单击该按钮，将弹出当前窗口的帮助信息。" & vbCrLf& vbCrLf & _
			"- 保存所有选择" & vbCrLf & _
			"  该按钮可以在更改配置而不进行检查时使用。" & vbCrLf & _
			"  如果任一选项被更改，该选项将自动变为可用状态，否则将自动变为不可用状态。" & vbCrLf & vbCrLf & _
			"- 确定" & vbCrLf & _
			"  单击该按钮，将关闭主对话框，并按选定的选项进行字串翻译。" & vbCrLf& vbCrLf & _
			"- 取消" & vbCrLf & _
			"  单击该按钮，将退出程序。" & vbCrLf& vbCrLf & vbCrLf
	SetUse ="☆配置列表☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 选择配置" & vbCrLf & _
			"  要选择配置，单击配置列表。" & vbCrLf & vbCrLf & _
			"- 设置配置的优先级" & vbCrLf & _
			"  配置优先级用于基于配置的适用语言的自动选择配置功能。" & vbCrLf & _
			"  ◎注意：有多个配置包含了相同的适用语言时，需要设置其优先级。" & vbCrLf & _
			"  　　　　在相同适用语言的配置中，前面的配置被优先选择使用。" & vbCrLf & _
			"  要设置配置的优先级，单击右边的 [...] 按钮。" & vbCrLf & vbCrLf & _
			"- 添加配置" & vbCrLf & _
			"  要添加配置，单击 [添加] 按钮，在弹出的对话框中输入名称。" & vbCrLf & vbCrLf & _
			"- 更改配置" & vbCrLf & _
			"  要更改配置名称，请选择配置列表中要改名的配置，然后单击 [更改] 按钮。" & vbCrLf & vbCrLf & _
			"- 删除配置" & vbCrLf & _
			"  要删除配置，请选择配置列表中要删除的配置，然后单击 [删除] 按钮。" & vbCrLf & vbCrLf & _
			"添加配置后，将在列表中显示新的配置，配置内容将显示空值。" & vbCrLf & _
			"更改配置后，将在列表中显示改名的配置，配置内容中的配置值不变。" & vbCrLf & _
			"删除配置后，将在列表中显示默认配置，配置内容将显示默认配置值。" & vbCrLf & vbCrLf & _
			"☆保存类型☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 文件" & vbCrLf & _
			"  配置将以文件形式保存在宏所在文件夹下的 Data 文件夹中。" & vbCrLf & vbCrLf & _
			"- 注册表" & vbCrLf & _
			"  配置将被保存注册表中的 HKCU\Software\VB and VBA Program Settings\AccessKey 项下。" & vbCrLf & vbCrLf & _
			"- 导入配置" & vbCrLf & _
			"  允许从其他配置文件中导入配置。导入旧配置时将被自动升级，现有配置列表中已有的配置将被" & vbCrLf & _
			"  更改，没有的配置将被添加。" & vbCrLf & vbCrLf & _
			"- 导出配置" & vbCrLf & _
			"  允许导出所有配置到文本文件，以便可以交换或转移配置。" & vbCrLf & vbCrLf & _
			"◎注意：切换保存类型时，将自动删除原有位置中的配置内容。" & vbCrLf & vbCrLf & _
			"☆配置内容☆" & vbCrLf & _
			"============" & vbCrLf & _
			"<快捷键>" & vbCrLf & _
			"  - 要排除的含 & 符号的非快捷键组合" & vbCrLf & _
			"    快捷键以 & 为标志符，有些字串虽然包含该符号但不是快捷键，需要排除它。在此输入这些" & vbCrLf & _
			"    要排除包含 & 符号的非快捷键组合。" & vbCrLf & vbCrLf & _
			"  - 字串拆分用标志符" & vbCrLf & _
			"    这些字符用于拆分含有多个快捷键，终止符或加速器的字串，以便检查字串中所有的快捷键、" & vbCrLf & _
			"    终止符或加速器。否则只能处理字串最后部分的快捷键、终止符或加速器。" & vbCrLf & vbCrLf & _
			"  - 要检查的快捷键前后括号" & vbCrLf & _
			"    默认的快捷键前后括号为 ()，在此指定的快捷键前后括号，都将被替换为默认的括号。" & vbCrLf & vbCrLf & _
			"  - 要保留的非快捷键前后成对字符" & vbCrLf & _
			"    如果原文和翻译中都存在这些字符。将被保留，否则将被认为是快捷键，并将添加快捷键并移" & vbCrLf & _
			"    位到字串最后。" & vbCrLf & vbCrLf & _
			"  - 在文本后面显示带括号的快捷键 (通常用于亚洲语言)" & vbCrLf & _
			"    通常在亚洲语言如中文、日文等软件中使用 (&X) 形式的快捷键，并将其置于字串结尾（在终" & vbCrLf & _
			"    止符或加速器前）。" & vbCrLf & _
			"    选定该选项后，将检查含有快捷键的翻译字串中的快捷键是否符合惯例。如果不符合将被自动" & vbCrLf & _
			"    更改并置后。" & vbCrLf & vbCrLf & _
			"<终止符>" & vbCrLf & _
			"  - 要检查的终止符" & vbCrLf & _
			"    翻译中的终止符和原文不一致时，将被更改为原文中的终止符，但是符合要自动替换的终止符" & vbCrLf & _
			"    对中的终止符除外。" & vbCrLf & vbCrLf & _
			"    该字段支持通配符，但是不是模糊的而是精确的。例如：A*C 不匹配 XAYYCZ，只匹配 AXYYC。" & vbCrLf & _
			"    要匹配 XAYYCZ，应该为 *A*C* 或 *A??C*。" & vbCrLf & vbCrLf & _
			"    ◎注意：由于 Sax Basic 引擎的问题，二个非英文字母之间的 ? 通配符不被支持。" & vbCrLf & _
			"    　　　　例如：“打开??文件”不匹配“打开用户文件”。" & vbCrLf & vbCrLf & _
			"  - 要保留的终止符组合" & vbCrLf & _
			"    所有被包含在该组合中的要检查的终止符将被保留。也就是这些终止符不被认为是终止符。" & vbCrLf & vbCrLf & _
			"    该字段支持通配符，但是不是模糊的而是精确的。例如：A*C 不匹配 XAYYCZ，只匹配 AXYYC。" & vbCrLf & _
			"    要匹配 XAYYCZ，应该为 *A*C* 或 *A??C*。" & vbCrLf & vbCrLf & _
			"    ◎注意：由于 Sax Basic 引擎的问题，二个非英文字母之间的 ? 通配符不被支持。" & vbCrLf & _
			"    　　　　例如：“打开??文件”不匹配“打开用户文件”。" & vbCrLf & vbCrLf & _
			"  - 要自动替换的终止符对" & vbCrLf & _
			"    符合终止符对中前一个字符的终止符，都将被替换成终止符对中后一个字符的终止符。" & vbCrLf & _
			"    利用此项可以自动翻译或修正一些终止符。" & vbCrLf & vbCrLf & _
			"<加速器>" & vbCrLf & _
			"  - 要检查的加速器标志符" & vbCrLf & _
			"    加速器通常以 \t 为标志符 (也有例外的)，如果字串中包含这些字符，将被认为包含加速器，" & vbCrLf & _
			"    但需要根据要检查的加速器字符进一步判断。" & vbCrLf & vbCrLf & _
			"  - 要检查的加速器字符" & vbCrLf & _
			"    包含加速器标志符的字串中，如果标志符后面的字符符合该字段的字符，将被识别为加速器，" & vbCrLf & _
			"    二个以上字符组合而成的加速器允许其中一个不匹配。" & vbCrLf & vbCrLf & _
			"    该字段支持通配符，但是不是模糊的而是精确的。例如：A*C 不匹配 XAYYCZ，只匹配 AXYYC。" & vbCrLf & _
			"    要匹配 XAYYCZ，应该为 *A*C* 或  *A??C*。" & vbCrLf & vbCrLf & _
			"    ◎注意：由于 Sax Basic 引擎的问题，二个非英文字母之间的 ? 通配符不被支持。" & vbCrLf & _
			"    　　　　例如：“打开??文件”不匹配“打开用户文件”。" & vbCrLf & vbCrLf & _
			"  - 要保留的加速器字符" & vbCrLf & _
			"    符合这些字符的加速器将被保留，否则将被替换。利用此项可保留某些加速器的翻译。" & vbCrLf & vbCrLf & _
			"<字符替换>" & vbCrLf & _
			"  字串中包含每个替换字符对中的“|”前的字符时，将被替换成“|”后的字符。" & vbCrLf & vbCrLf & _
			"  - 要自动替换的字符" & vbCrLf & _
			"    定义在检查并修正过程中要被替换的字符以及替换后的字符。" & vbCrLf & vbCrLf & _
			"  ◎注意：替换时区分大小写。" & vbCrLf & _
			"  　　　　如果要去掉这些字符，可以将“|”后的字符置空。" & vbCrLf & vbCrLf & _
			"<适用语言>" & vbCrLf & _
			"  这里的适用语言是指翻译列表的目标语言，它用于根据翻译列表的目标语言自动选择相应配置的" & vbCrLf & _
			"  自动选择功能。" & vbCrLf & vbCrLf & _
			"  - 添加" & vbCrLf & _
			"    要添加适用语言，选择可用语言列表中的语言，然后单击 [添加] 按钮。" & vbCrLf & _
			"    单击该按钮后，可用语言列表中的选定语言将移动到适用语言列表中。" & vbCrLf & vbCrLf & _
			"  - 全部添加" & vbCrLf & _
			"    单击该按钮后，可用语言列表中的所有语言将全部移动到适用语言列表中。" & vbCrLf & vbCrLf & _
			"  - 删除" & vbCrLf & _
			"    要删除适用语言，选择适用语言列表中的语言，然后单击 [删除] 按钮。" & vbCrLf & _
			"    单击该按钮后，适用语言列表中的选定语言将移动到可用语言列表中。" & vbCrLf & vbCrLf & _
			"  - 全部删除" & vbCrLf & _
			"    单击该按钮后，适用语言列表中的所有语言将全部移动到可用语言列表中。" & vbCrLf & vbCrLf & _
			"  - 增加可用语言" & vbCrLf & _
			"    单击该按钮后，将弹出可输入语言名称和代码对话框，确定后将添加到可用语言列表中。" & vbCrLf & vbCrLf & _
			"  - 编辑可用语言" & vbCrLf & _
			"    要编辑可用语言，选择可用语言列表中的语言，然后单击 [编辑可用语言] 按钮。" & vbCrLf & _
			"    单击该按钮后，将弹出可编辑语言名称和代码对话框，确定后将修改可用语言列表中选定的语言。" & vbCrLf & vbCrLf & _
			"  - 删除可用语言" & vbCrLf & _
			"    要删除可用语言，选择可用语言列表中要删除的语言，然后单击 [删除可用语言] 按钮。" & vbCrLf & vbCrLf & _
			"  - 增加适用语言" & vbCrLf & _
			"    单击该按钮后，将弹出可输入语言名称和代码对话框，确定后将添加到适用语言列表中。" & vbCrLf & vbCrLf & _
			"  - 编辑适用语言" & vbCrLf & _
			"    要编辑适用语言，选择适用语言列表中的语言，然后单击 [编辑适用语言] 按钮。" & vbCrLf & _
			"    单击该按钮后，将弹出可编辑语言名称和代码对话框，确定后将修改适用语言列表中选定的语言。" & vbCrLf & vbCrLf & _
			"  - 删除适用语言" & vbCrLf & _
			"    要删除适用语言，选择适用语言列表中要删除的语言，然后单击 [删除适用语言] 按钮。" & vbCrLf & vbCrLf & _
			"  ◎注意：添加、编辑语言仅用于 Passolo 未来版本新增的支持语言。" & vbCrLf & _
			"  　　　　语言代码请和 Passolo 的 ISO 396-1 代码保持一致，包括大小写。" & vbCrLf & vbCrLf & _
			"☆其他功能☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 帮助" & vbCrLf & _
			"  单击该按钮，将弹出当前窗口的帮助信息。" & vbCrLf & vbCrLf & _
			"- 读取" & vbCrLf & _
			"  单击该按钮，将根据选定配置的不同弹出下列菜单:" & vbCrLf & _
			"  (1) 默认值" & vbCrLf & _
			"      读取默认配置值，并显示在配置内容中。" & vbCrLf & _
			"      ◎仅当选定的配置为系统默认的配置时，才显示该菜单。" & vbCrLf & vbCrLf & _
			"  (2) 原值" & vbCrLf & _
			"      读取选定配置的原始值，并显示在配置内容中。" & vbCrLf & _
			"      ◎仅当选定配置的原始值为非空时，才显示该菜单。" & vbCrLf & vbCrLf & _
			"  (3) 参照值" & vbCrLf & _
			"      读取选定的参照配置值，并显示在配置内容中。" & vbCrLf & _
			"      ◎该菜单显示除选定配置外的所有配置列表。" & vbCrLf & vbCrLf & _
			"- 清空" & vbCrLf & _
			"  单击该按钮，将清空现有配置的全部值，以方便重新输入配置值。" & vbCrLf & vbCrLf & _
			"- 测试" & vbCrLf & _
			"  单击该按钮，将弹出测试对话框，以便检查配置的正确性。" & vbCrLf & vbCrLf & _
			"- 确定" & vbCrLf & _
			"  单击该按钮，将保存配置窗口中的任何更改，退出配置窗口并返回主窗口。" & vbCrLf & _
			"  程序将使用更改后的配置值。" & vbCrLf & vbCrLf & _
			"- 取消" & vbCrLf & _
			"  单击该按钮，不保存配置窗口中的任何更改，退出配置窗口并返回主窗口。" & vbCrLf & _
			"  程序将使用原来的配置值。" & vbCrLf & vbCrLf & vbCrLf
	TestUse ="☆配置名称☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 要测试的配置名称。要选择配置，单击配置列表。" & vbCrLf & vbCrLf & _
			"☆翻译列表☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 该列表将显示方案中的所有翻译列表。请选择与您的自定义配置匹配的翻译列表进行测试。" & vbCrLf & vbCrLf & _
			"☆自动替换字符☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 在检查时自动替换字串中符合配置中所定义的替换字符。" & vbCrLf & vbCrLf & _
			"☆读入行数☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 表示要显示的错误翻译字串数。建议不要输入太大的值，以免在字串较多时等待时间过长。" & vbCrLf & vbCrLf & _
			"☆包含内容☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 指定只检查包含内容的字串。利用该项可以有针对性的测试，并且加快测试时间。" & vbCrLf & _
			"- 该字段支持模糊型通配符。例如：A*C 可以匹配 XAYYCZ。" & vbCrLf & vbCrLf & _
			"◎注意：由于 Sax Basic 引擎的问题，二个非英文字母之间的 ? 通配符不被支持。" & vbCrLf & _
			"　　　　例如：“打开??文件”不匹配“打开用户文件”。" & vbCrLf & vbCrLf & _
			"☆字串内容☆" & vbCrLf & _
			"============" & vbCrLf & _
			"提供了全部、快捷键、终止符、加速器 4 个选项。" & vbCrLf & vbCrLf & _
			"- 如果选择全部，则其他单项将被自动取消选择。" & vbCrLf & _
			"- 如果选择单项，则全部选项将被自动取消选择。" & vbCrLf & _
			"- 单项可以多选。" & vbCrLf & vbCrLf & _
			"☆其他功能☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 帮助" & vbCrLf & _
			"  单击该按钮，将弹出当前窗口的帮助信息。" & vbCrLf & vbCrLf & _
			"- 测试" & vbCrLf & _
			"  单击该按钮，将按照选定的条件进行测试。" & vbCrLf & vbCrLf & _
			"- 清空" & vbCrLf & _
			"  单击该按钮，将清空现有的测试结果。" & vbCrLf & vbCrLf & _
			"- 取消" & vbCrLf & _
			"  单击该按钮，将退出测试程序并返回配置窗口。" & vbCrLf & vbCrLf & vbCrLf
	CopyRight = "☆版权声明☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 此软件的版权归开发者和修改者所有，任何人可以免费使用、修改、复制、散布本软件。" & vbCrLf & _
			"- 修改、散布本软件必须随附本说明文件，并注明软件原始开发者以及修改者。" & vbCrLf & _
			"- 未经开发者和修改者同意，任何组织或个人，不得用于商业软件、商业或是其它营利性活动。" & vbCrLf & _
			"- 对使用本软件的原始版本，以及使用经他人修改的非原始版本所造成的损失和损害，开发者不" & vbCrLf & _
			"  承担任何责任。" & vbCrLf & _
			"- 由于为免费软件，开发者和修改者没有义务提供软件技术支持，也无义务改进或更新版本。" & vbCrLf & _
			"- 欢迎指出错误并提出改进意见。如有错误或建议，请发送到: z_shangyi@163.com。" & vbCrLf & vbCrLf & vbCrLf
	Thank = "☆致　　谢☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 本软件在修改过程中得到汉化新世纪会员的测试，在此表示衷心的感谢！" & vbCrLf & _
			"- 感谢台湾 Heaven 先生提出繁体用语修改意见！" & vbCrLf & vbCrLf & vbCrLf
	Contact = "☆与我联系☆" & vbCrLf & _
			"============" & vbCrLf & _
			"wanfu：z_shangyi@163.com" & vbCrLf & vbCrLf & vbCrLf
	Logs = "感谢支持！您的支持是我最大的动力！同时欢迎使用我们制作的软件！" & vbCrLf & vbCrLf & _
			"**************************************************" & vbCrLf & _
			"需要更多、更新、更好的汉化，请访问:" & vbCrLf & _
			"汉化新世纪 -- http://www.hanzify.org" & vbCrLf & _
			"汉化新世纪论坛 -- http://bbs.hanzify.org" & vbCrLf & _
			"**************************************************" & vbCrLf
	End If

	preLines = Lines & Lines & Lines
	appLines = Lines & Lines & Lines & vbCrLf & vbCrLf
	If HelpTip = "MainHelp" Then
		Title = HelpTitle
		HelpMsg = preLines & MainWindows & appLines & MainUse & Logs
	ElseIf HelpTip = "SetHelp" Then
		Title = HelpTitle
		HelpMsg = preLines & SetWindows & appLines & SetUse & Logs
	ElseIf HelpTip = "TestHelp" Then
		Title = HelpTitle
		HelpMsg = preLines & TestWindows & appLines & TestUse & Logs
	ElseIf HelpTip = "About" Then
		Title = AboutTitle
		HelpMsg = preLines & AboutWindows & appLines & Sys & Dec & Ement & Setup & CopyRight & Thank & Contact & Logs
	End If

	Begin Dialog UserDialog 760,413,Title ' %GRID:10,7,1,1
		Text 0,7,760,14,HelpTipTitle,.Text,2
		TextBox 0,28,760,350,.TextBox,1
		OKButton 330,385,100,21
	End Dialog
	Dim dlg As UserDialog
	dlg.TextBox = HelpMsg
	Dialog dlg
End Sub


'自动更新帮助
Sub UpdateHelp(HelpTip As String)
	If OSLanguage = "0404" Then
	HelpTitle = "弧"
	HelpTipTitle = "笆穝"
	SetWindows = " 砞﹚跌怠 "
	Lines = "-----------------------"
	SetUse ="「穝よΑ「" & vbCrLf & _
			"============" & vbCrLf & _
			"- 笆更穝杆" & vbCrLf & _
			"  匡拒赣匡兜祘Α盢沮穝繵瞯い砞﹚笆浪琩穝狦盎代Τ穝セノ" & vbCrLf & _
			"  盢ぃ紉―ㄏノ種ǎ薄猵笆更穝杆穝Ч拨盢紆癸杠よ遏硄ㄏノ" & vbCrLf & _
			"  [絋﹚] 秙祘Α挡" & vbCrLf & vbCrLf & _
			"- Τ穝硄иパи∕﹚更杆" & vbCrLf & _
			"  匡拒赣匡兜祘Α盢沮穝繵瞯い砞﹚笆浪琩穝狦盎代Τ穝セノ" & vbCrLf & _
			"  盢紆癸杠よ遏矗ボㄏノ狦ㄏノ∕﹚穝祘Α盢更穝杆穝Ч拨盢紆癸" & vbCrLf & _
			"  杠よ遏硄ㄏノ [絋﹚] 秙祘Α挡" & vbCrLf & vbCrLf & _
			"- 闽超笆穝" & vbCrLf & _
			"  匡拒赣匡兜祘Α盢ぃ浪琩穝" & vbCrLf & vbCrLf & _
			"》猔種礚阶贺穝よΑ穝Θ挡祘Α常惠璶ㄏノ穝币笆エ栋" & vbCrLf & vbCrLf & _
			"「穝繵瞯「" & vbCrLf & _
			"============" & vbCrLf & _
			"- 浪琩丁筳" & vbCrLf & _
			"  浪琩穝丁秅戳祘Α浪琩秅戳ず浪琩Ω" & vbCrLf & vbCrLf & _
			"- 程浪琩ら戳" & vbCrLf & _
			"  程浪琩ら戳祘Α浪琩穝笆癘魁浪琩ら戳ら戳ず浪琩Ω" & vbCrLf & vbCrLf & _
			"- 浪琩" & vbCrLf & _
			"  翴阑赣秙祘Α盢┛菠浪琩丁筳㎝浪琩ら戳秈︽穝浪琩穝浪琩ら戳" & vbCrLf & vbCrLf & _
			"  》狦盎代Τ穝セノ盢紆癸杠よ遏矗ボㄏノ狦ㄏノ∕﹚穝祘Α盢更穝" & vbCrLf & _
			"  杆穝Ч拨盢紆癸杠よ遏硄ㄏノ [絋﹚] 秙祘Α挡" & vbCrLf & _
			"  》狦⊿Τ盎代Τ穝セノ盢紆ヘ玡呼セ腹矗ボ琌穝更穝" & vbCrLf & vbCrLf & _
			"「穝呼睲虫「" & vbCrLf & _
			"============" & vbCrLf & _
			"矪穝呼ㄏノ﹚竡ㄏノ﹚竡呼盢砆纔ㄏノ" & vbCrLf & _
			"ㄏノ﹚竡呼礚猭盢ㄏノ祘Α秨祇﹚竡穝呼" & vbCrLf & vbCrLf & _
			"「RAR 秆溃祘Α「" & vbCrLf & _
			"================" & vbCrLf & _
			"パ祘Α蹦ノ RAR Α溃罽珿惠璶セ诀杆Τ癸莱秆溃祘Α" & vbCrLf & _
			"Ωㄏノ祘Α盢笆穓爹い爹 RAR 捌郎箇砞秆溃祘Α秈︽続讽砞﹚" & vbCrLf & _
			"–Ωㄏノ笆浪琩秆溃祘Α琌临狦ぃ盢穝穓砞﹚" & vbCrLf & vbCrLf & _
			"》猔種祘Α箇砞や穿秆溃祘ΑWinRARWinZIP7z狦诀竟い⊿Τ硂ㄇ秆溃祘Α" & vbCrLf & _
			"惠璶も笆砞﹚" & vbCrLf & vbCrLf & _
			"- 祘Α隔畖" & vbCrLf & _
			"  秆溃祘ΑЧ俱隔畖翴阑娩 [...] 秙も穝糤" & vbCrLf & vbCrLf & _
			"- 秆溃把计" & vbCrLf & _
			"  秆溃祘Α秆溃罽 RAR 郎㏑把计ㄤい" & vbCrLf & _
			"  %1 溃罽郎%2 璶眖溃罽い耝祘Α郎%3 秆溃郎隔畖" & vbCrLf & _
			"  硂ㄇ把计ゲ璶把计ぃぶぃノㄤ才腹蠢抖ㄌ酚秆溃祘Α" & vbCrLf & _
			"  ㏑砏玥" & vbCrLf & _
			"  翴阑娩 [>] 秙も穝糤硂ㄇゲ璶把计" & vbCrLf & vbCrLf & _
			"「ㄤ「" & vbCrLf & _
			"============" & vbCrLf & _
			"- 弧" & vbCrLf & _
			"  翴阑赣秙盢莉弧癟" & vbCrLf & vbCrLf & _
			"- 弄" & vbCrLf & _
			"  翴阑赣秙盢沮匡拒砞﹚ぃ紆匡虫:" & vbCrLf & _
			"  (1) 箇砞" & vbCrLf & _
			"      弄笆穝砞﹚箇砞陪ボ穝呼睲虫㎝ RAR 秆溃祘Αい" & vbCrLf & _
			"  (2) " & vbCrLf & _
			"      弄笆穝砞﹚﹍陪ボ穝呼睲虫㎝ RAR 秆溃祘Αい" & vbCrLf & vbCrLf & _
			"- 睲" & vbCrLf & _
			"  盢穝呼睲虫㎝RAR 秆溃祘Αいず甧场睲" & vbCrLf & vbCrLf & _
			"- 代刚" & vbCrLf & _
			"  翴阑赣秙盢代刚穝呼睲虫㎝ RAR 秆溃祘Α琌タ絋" & vbCrLf & vbCrLf & _
			"- 絋﹚" & vbCrLf & _
			"  翴阑赣秙盢纗砞﹚跌怠いヴ跑挡砞﹚跌怠跌怠" & vbCrLf & _
			"  祘Α盢ㄏノ跑砞﹚" & vbCrLf & vbCrLf & _
			"- " & vbCrLf & _
			"  翴阑赣秙ぃ纗砞﹚跌怠いヴ跑挡砞﹚跌怠跌怠" & vbCrLf & _
			"  祘Α盢ㄏノㄓ砞﹚" & vbCrLf & vbCrLf & vbCrLf
	Logs = "稰谅や眤や琌и程笆舧ㄏノи籹硁砰" & vbCrLf & vbCrLf & _
			"**************************************************" & vbCrLf & _
			"惠璶穝簙て叫砐:" & vbCrLf & _
			"簙て穝 -- http://www.hanzify.org" & vbCrLf & _
			"簙て穝阶韭 -- http://bbs.hanzify.org" & vbCrLf & _
			"**************************************************" & vbCrLf
	Else
	HelpTitle = "帮助"
	HelpTipTitle = "自动更新"
	SetWindows = " 配置窗口 "
	Lines = "-----------------------"
	SetUse ="☆更新方式☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 自动下载更新并安装" & vbCrLf & _
			"  选定该选项时，程序将根据更新频率中的设置自动检查更新，如果检测到有新的版本可用时，" & vbCrLf & _
			"  将在不征求用户意见的情况下自动下载更新并安装。更新完毕后将弹出对话框通知用户，按" & vbCrLf & _
			"  [确定] 按钮后程序即退出。" & vbCrLf & vbCrLf & _
			"- 有更新时通知我，由我决定下载并安装" & vbCrLf & _
			"  选定该选项时，程序将根据更新频率中的设置自动检查更新，如果检测到有新的版本可用时，" & vbCrLf & _
			"  将弹出对话框提示用户，如果用户决定更新，程序将下载更新并安装。更新完毕后将弹出对" & vbCrLf & _
			"  话框通知用户，按 [确定] 按钮后程序即退出。" & vbCrLf & vbCrLf & _
			"- 关闭自动更新" & vbCrLf & _
			"  选定该选项时，程序将不检查更新。" & vbCrLf & vbCrLf & _
			"◎注意：无论何种更新方式，更新成功并退出程序后都需要用户重新启动宏。" & vbCrLf & vbCrLf & _
			"☆更新频率☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 检查间隔" & vbCrLf & _
			"  检查更新的时间周期。程序在检查周期内只检查一次。" & vbCrLf & vbCrLf & _
			"- 最后检查日期" & vbCrLf & _
			"  最后检查时的日期。程序在检查更新时自动记录检查日期，并在同一日期内只检查一次。" & vbCrLf & vbCrLf & _
			"- 检查" & vbCrLf & _
			"  单击该按钮后，程序将忽略检查间隔和检查日期进行更新检查，并更新检查日期。" & vbCrLf & vbCrLf & _
			"  ◎如果检测到有新的版本可用，将弹出对话框提示用户，如果用户决定更新，程序将下载更新" & vbCrLf & _
			"  　并安装。更新完毕后将弹出对话框通知用户，按 [确定] 按钮后程序即退出。" & vbCrLf & _
			"  ◎如果没有检测到有新的版本可用，将弹出当前网上的版本号。并提示是否重新下载更新。" & vbCrLf & vbCrLf & _
			"☆更新网址列表☆" & vbCrLf & _
			"============" & vbCrLf & _
			"此处的更新网址用户可以自己定义。用户定义的网址将被优先使用。" & vbCrLf & _
			"用户定义的网址无法访问时，将使用程序开发者定义的更新网址。" & vbCrLf & vbCrLf & _
			"☆RAR 解压程序☆" & vbCrLf & _
			"================" & vbCrLf & _
			"由于程序包采用了 RAR 格式压缩，故需要在本机上安装有相应的解压程序。" & vbCrLf & _
			"初次使用时，程序将自动搜索注册表中注册的 RAR 扩展名默认解压程序，并进行适当的配置。" & vbCrLf & _
			"并在每次使用时自动检查解压程序是否还存在，如果不存在将重新搜索并配置。" & vbCrLf & vbCrLf & _
			"◎注意：程序默认支持的解压程序为：WinRAR、WinZIP、7z。如果机器中没有这些解压程序，" & vbCrLf & _
			"　　　　需要手动配置。" & vbCrLf & vbCrLf & _
			"- 程序路径" & vbCrLf & _
			"  解压程序的完整路径。可单击右边的 [...] 按钮手工添加。" & vbCrLf & vbCrLf & _
			"- 解压参数" & vbCrLf & _
			"  解压程序解压缩 RAR 文件时的命令行参数。其中：" & vbCrLf & _
			"  %1 为压缩文件，%2 为要从压缩包中提取的主程序文件，%3 为解压后的文件路径。" & vbCrLf & _
			"  这些参数为必要参数，不可缺少并且不可用其他符号代替。至于先后顺序，依照解压程序的" & vbCrLf & _
			"  命令行规则。" & vbCrLf & _
			"  可单击右边的 [>] 按钮手工添加这些必要参数。" & vbCrLf & vbCrLf & _
			"☆其他功能☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 帮助" & vbCrLf & _
			"  单击该按钮，将获取帮助信息。" & vbCrLf & vbCrLf & _
			"- 读取" & vbCrLf & _
			"  单击该按钮，将根据选定配置的不同弹出下列菜单:" & vbCrLf & _
			"  (1) 默认值" & vbCrLf & _
			"      读取自动更新配置的默认值，并显示在更新网址列表和 RAR 解压程序中。" & vbCrLf & _
			"  (2) 原值" & vbCrLf & _
			"      读取自动更新配置的原始值，并显示在更新网址列表和 RAR 解压程序中。" & vbCrLf & vbCrLf & _
			"- 清空" & vbCrLf & _
			"  将更新网址列表和RAR 解压程序中的内容全部清空。" & vbCrLf & vbCrLf & _
			"- 测试" & vbCrLf & _
			"  单击该按钮，将测试更新网址列表和 RAR 解压程序是否正确。" & vbCrLf & vbCrLf & _
			"- 确定" & vbCrLf & _
			"  单击该按钮，将保存配置窗口中的任何更改，退出配置窗口并返回主窗口。" & vbCrLf & _
			"  程序将使用更改后的配置值。" & vbCrLf & vbCrLf & _
			"- 取消" & vbCrLf & _
			"  单击该按钮，不保存配置窗口中的任何更改，退出配置窗口并返回主窗口。" & vbCrLf & _
			"  程序将使用原来的配置值。" & vbCrLf & vbCrLf & vbCrLf
	Logs = "感谢支持！您的支持是我最大的动力！同时欢迎使用我们制作的软件！" & vbCrLf & vbCrLf & _
			"**************************************************" & vbCrLf & _
			"需要更多、更新、更好的汉化，请访问:" & vbCrLf & _
			"汉化新世纪 -- http://www.hanzify.org" & vbCrLf & _
			"汉化新世纪论坛 -- http://bbs.hanzify.org" & vbCrLf & _
			"**************************************************" & vbCrLf
	End If

	preLines = Lines & Lines & Lines
	appLines = Lines & Lines & Lines & vbCrLf & vbCrLf
	If HelpTip = "SetHelp" Then
		Title = HelpTitle
		HelpMsg = preLines & SetWindows & appLines & SetUse & Logs
	End If

	Begin Dialog UserDialog 760,413,Title ' %GRID:10,7,1,1
		Text 0,7,760,14,HelpTipTitle,.Text,2
		TextBox 0,28,760,350,.TextBox,1
		OKButton 330,385,100,21
	End Dialog
	Dim dlg As UserDialog
	dlg.TextBox = HelpMsg
	Dialog dlg
End Sub
