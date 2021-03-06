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


'趼揹揭燴蘇�珈髲�
Function CheckSettings(DataName As String,OSLang As String) As String
	Dim ExCr As String,LnSp As String,ChkBkt As String,KpPair As String,ChkEnd As String
	Dim NoTrnEnd As String,TrnEnd As String,Short As String,Key As String,KpKey As String
	If DataName = DefaultCheckList(0) Then
		If OSLang = "0404" Then
			ExCr = "&&,& ,&#,&amp;,&nbsp;,&quot;,&lt;,&gt;,&apos;"
			LnSp = "\r,\n,\r\n,|,&"
			ChkBkt = "(),（）,[],〔〕,<>,＜＞,〈〉"
			KpPair = "(),（）,[],〔〕,{},｛｝,<>,＜＞,〈〉,《》,【】,「」,『』,'',『』,ˋˊ,‵‵,「」,〝〞,"""""
			AsiaKey = "1"

			ChkEnd = ". .. ... .... ..... ...... 。 : ： ; ； ! ！ ? ？ , ， 、 > >> -> ] } + -"
			TrnEnd = ",|， .|。 ;|； !|！ ?|？"
			NoTrnEnd = "*[*],*{*},*<*>,*(?,*[?,*],*{?,*},*<?,*>,*%?!?!,*%?!??!"

			Short = "\t,\b"
			Key = "BackSpace,Tab,Clear,Enter,Shift,Ctrl,Alt,Pause,Caps Lock,Escape,Esc,Space,Break," & _
					"Page Up,Page Down,PgUp,PgDn,End,Home,Left,Up,Right,Down,Left Arrow,Up Arrow," & _
					"Right Arrow,Down Arrow,Select,Print,Ins,Delete,Del,Help,Num Lock,F#,F1#,%?,&?,?"
			KpKey = "向上?,向下?,向左?,向右?,上箭?,下箭?,左箭?,右箭?," & _
					"向上鍵,向下鍵,向左鍵,向右鍵,上箭頭,下箭頭,左箭頭,右箭頭,↑,↓,←,→"
			PreStr = "\001|\n001,\u00A9|(C),\u00AE|(R)"
			AppStr = "&lt;|<,&gt;|>,&quot;|"",&apos;|',&#9;|\t,&#10;|\n,&#13;|\r,&amp;|&,&#39;|'"
		Else
			ExCr = "&&,& ,&#,&amp;,&nbsp;,&quot;,&lt;,&gt;,&apos;"
			LnSp = "\r,\n,\r\n,|,&"
			ChkBkt = "(),ㄗㄘ,[],�菇�,<>,ˉˇ,●△"
			KpPair = "(),ㄗㄘ,[],�菇�,{},����,<>,ˉˇ,●△,▲◎,▽▼,☆★,◇◆,'',＆＊,杗杓,沙沙,※§,����,"""""
			AsiaKey = "1"

			ChkEnd = ". .. ... .... ..... ...... ﹝ : ㄩ ; ˙ ! ㄐ ? ˋ , ㄛ ﹜ > >> -> ] } + -"
			TrnEnd = ",|ㄛ .|﹝ ;|˙ !|ㄐ ?|ˋ"
			NoTrnEnd = "*[*],*{*},*<*>,*(?,*[?,*],*{?,*},*<?,*>,*%?!?!,*%?!??!"

			Short = "\t,\b"
			Key = "BackSpace,Tab,Clear,Enter,Shift,Ctrl,Alt,Pause,Caps Lock,Escape,Esc,Space,Break," & _
					"Page Up,Page Down,PgUp,PgDn,End,Home,Left,Up,Right,Down,Left Arrow,Up Arrow," & _
					"Right Arrow,Down Arrow,Select,Print,Ins,Delete,Del,Help,Num Lock,F#,F1#,%?,&?,?"
			KpKey = "砃奻瑩,砃狟瑩,砃酘瑩,砃衵瑩,奻璋芛,狟璋芛,酘璋芛,衵璋芛," & _
					"砃奻潬,砃狟潬,砃酘潬,砃衵潬,奻璋螹,狟璋螹,酘璋螹,衵璋螹,∥,∣,↘,↙"
			PreStr = "\001|\n001,\u00A9|(C),\u00AE|(R)"
			AppStr = "&lt;|<,&gt;|>,&quot;|"",&apos;|',&#9;|\t,&#10;|\n,&#13;|\r,&amp;|&,&#39;|'"
		End If
	ElseIf DataName = DefaultCheckList(1) Then
		If OSLang = "0404" Then
			ExCr = "&&,& ,&#,&amp;,&nbsp;,&quot;,&lt;,&gt;,&apos;"
			LnSp = "\r,\n,\r\n,|,&"
			ChkBkt = "(),（）,[],〔〕,<>,＜＞,〈〉"
			KpPair = "(),（）,[],〔〕,{},｛｝,<>,＜＞,〈〉,《》,【】,「」,『』,'',『』,ˋˊ,‵‵,「」,〝〞,"""""
			AsiaKey = "0"

			ChkEnd = ". .. ... .... ..... ...... 。 : ： ; ； ! ！ ? ？ , ， 、 > >> -> ] } + -"
			TrnEnd = "，|, 。|. ；|; ！|! ？|?"
			NoTrnEnd = "*[*],*{*},*<*>,*(?,*[?,*],*{?,*},*<?,*>,*%?!?!,*%?!??!"

			Short = "\t,\b"
			Key = "BackSpace,Tab,Clear,Enter,Shift,Ctrl,Alt,Pause,Caps Lock,Escape,Esc,Space,Break," & _
					"Page Up,Page Down,PgUp,PgDn,End,Home,Left,Up,Right,Down,Left Arrow,Up Arrow," & _
					"Right Arrow,Down Arrow,Select,Print,Ins,Delete,Del,Help,Num Lock,F#,F1#,%?,&?,?," & _
					"向上?,向下?,向左?,向右?,上箭?,下箭?,左箭?,右箭?," & _
					"向上鍵,向下鍵,向左鍵,向右鍵,上箭頭,下箭頭,左箭頭,右箭頭,↑,↓,←,→"
			KpKey = "Up,Right,Down,Left Arrow,Up Arrow,Right Arrow,Down Arrow"
			PreStr = "\001|\n001,\u00A9|(C),\u00AE|(R)"
			AppStr = "&lt;|<,&gt;|>,&quot;|"",&apos;|',&#9;|\t,&#10;|\n,&#13;|\r,&amp;|&,&#39;|'"
		Else
			ExCr = "&&,& ,&#,&amp;,&nbsp;,&quot;,&lt;,&gt;,&apos;"
			LnSp = "\r,\n,\r\n,|,&"
			ChkBkt = "(),ㄗㄘ,[],�菇�,<>,ˉˇ,●△"
			KpPair = "(),ㄗㄘ,[],�菇�,{},����,<>,ˉˇ,●△,▲◎,▽▼,☆★,◇◆,'',＆＊,杗杓,沙沙,※§,����,"""""
			AsiaKey = "0"

			ChkEnd = ". .. ... .... ..... ...... ﹝ : ㄩ ; ˙ ! ㄐ ? ˋ , ㄛ ﹜ > >> -> ] } + -"
			TrnEnd = "ㄛ|, ﹝|. ˙|; ㄐ|! ˋ|?"
			NoTrnEnd = "*[*],*{*},*<*>,*(?,*[?,*],*{?,*},*<?,*>,*%?!?!,*%?!??!"

			Short = "\t,\b"
			Key = "BackSpace,Tab,Clear,Enter,Shift,Ctrl,Alt,Pause,Caps Lock,Escape,Esc,Space,Break," & _
					"Page Up,Page Down,PgUp,PgDn,End,Home,Left,Up,Right,Down,Left Arrow,Up Arrow," & _
					"Right Arrow,Down Arrow,Select,Print,Ins,Delete,Del,Help,Num Lock,F#,F1#,%?,&?,?," & _
					"砃奻瑩,砃狟瑩,砃酘瑩,砃衵瑩,奻璋芛,狟璋芛,酘璋芛,衵璋芛," & _
					"砃奻潬,砃狟潬,砃酘潬,砃衵潬,奻璋螹,狟璋螹,酘璋螹,衵璋螹,∥,∣,↘,↙""
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


'③昢斛脤艘勤趕遺堆翑翋枙眕賸賤載嗣陓洘﹝
Private Function MainDlgFunc%(DlgItem$, Action%, SuppValue&)
	Dim TypeMsg As Integer,CountMsg As Integer,HeaderID As Integer,Header As String
	If OSLanguage = "0404" Then
		Msg29 = "錯誤"
		Msg30 = "請選取要處理的字串類型！"
		Msg31 = "請選取要處理的字串內容！"
		Msg32 = "請選取要處理的字串類型和內容！"
		Msg36 = "無法儲存！請檢查是否有寫入下列位置的權限:" & vbCrLf & vbCrLf
	Else
		Msg29 = "渣昫"
		Msg30 = "③恁寁猁揭燴腔趼揹濬倰ㄐ"
		Msg31 = "③恁寁猁揭燴腔趼揹囀�搟�"
		Msg32 = "③恁寁猁揭燴腔趼揹濬倰睿囀�搟�"
		Msg36 = "拸楊悵湔ㄐ③潰脤岆瘁衄迡�輴臏倛閥繭饑使�:" & vbCrLf & vbCrLf
	End If

	'鳳�＿膨縎擿�
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
	Case 2 ' 杅硉載蜊麼氪偌狟賸偌聽
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
		'潰聆趼揹濬倰睿囀�楪√鯓Й鮹畎敗▲亞垓蕩芢頖�等砐
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
			'潰聆趼揹濬倰睿囀�楪√鯓Й鮽矽�
			TypeMsg = mAllType + mMenu + mDialog + mString + mAccTable + mVer + mOther + mSelOnly
			CountMsg = mAllCont + mAccKey + mEndChar + mAcceler
			If TypeMsg = 0 And CountMsg <> 0 Then
				MsgBox(Msg30,vbOkOnly+vbInformation,Msg29)
				MainDlgFunc% = True ' 滅砦偌狟偌聽壽敕勤趕遺敦諳
			ElseIf TypeMsg <> 0 And CountMsg = 0 Then
				MsgBox(Msg31,vbOkOnly+vbInformation,Msg29)
				MainDlgFunc% = True ' 滅砦偌狟偌聽壽敕勤趕遺敦諳
			ElseIf TypeMsg = 0 And CountMsg = 0 Then
				MsgBox(Msg32,vbOkOnly+vbInformation,Msg29)
				MainDlgFunc% = True ' 滅砦偌狟偌聽壽敕勤趕遺敦諳
			End If
			If Join(cSelected) = nSelected Then Exit Function
			cSelected = Split(nSelected,JoinStr)
			If CheckWrite(CheckDataList,cWriteLoc,"Main") = False Then
				If WriteLocation = CheckFilePath Then Msg36 = Msg36 & CheckFilePath
				If WriteLocation = CheckRegKey Then Msg36 = Msg36 & CheckRegKey
				If MsgBox(Msg36,vbYesNo+vbInformation,Msg29) = vbNo Then Exit Function
				MainDlgFunc% = True ' 滅砦偌狟偌聽壽敕勤趕遺敦諳
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
			MainDlgFunc% = True ' 滅砦偌狟偌聽壽敕勤趕遺敦諳
		End If
	End Select
End Function


' 翋最唗
Sub Main
	Dim i As Integer,j As Integer,CheckVer As String,CheckSet As String,CheckID As Integer
	Dim srcString As String,trnString As String,OldtrnString As String,NewtrnString As String
	Dim LineNumErrCount As Integer,accKeyNumErrCount As Integer,ErrorCount As Integer
	Dim TransListOpen As Boolean,CheckDate As Date,TranDate As Date,CheckState As String
	Dim TranLang As String

	On Error GoTo SysErrorMsg
	'潰聆炵苀逄晟
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
		Msg00 = "(c) 2007-2008 by gnatix, 2009-2010 by wanfu  版本: " & Version
		Msg01 = "便捷鍵、終止符和加速器檢查巨集"
		Msg02 = "本程式用於檢查、修改和刪除該翻譯清單中所有翻譯字串" & _
				"的便捷鍵、終止符和加速器。請在檢查後逐條進行人工複查。"
		Msg03 = "翻譯清單: "

		Msg04 = "設定選取"
		Msg05 = "自動巨集:"
		Msg06 = "檢查巨集:"
		Msg07 = "自動巨集和檢查巨集相同(&U)"
		Msg08 = "自動選取(&X)"

		Msg09 = "檢查標記"
		Msg10 = "忽略版本(&B)"
		Msg11 = "忽略設定(&G)"
		Msg12 = "忽略日期(&T)"
		Msg13 = "忽略狀態(&Z)"
		Msg14 = "全部忽略(&Q)"

		Msg15 = "操作選取"
		Msg16 = "僅檢查(&O)"
		Msg17 = "檢查並修正(&M)"
		Msg18 = "刪除便捷鍵(&R)"

		Msg19 = "字串類型"
		Msg20 = "全部(&A)"
		Msg21 = "選單(&M)"
		Msg22 = "對話方塊(&D)"
		Msg23 = "字串表(&S)"
		Msg24 = "加速器表(&A)"
		Msg25 = "版本(&V)"
		Msg26 = "其他(&O)"
		Msg27 = "僅選擇(&L)"

		Msg28 = "字串內容"
		Msg29 = "全部(&F)"
		Msg30 = "便捷鍵(&K)"
		Msg31 = "終止符(&E)"
		Msg32 = "加速器(&P)"

		Msg33 = "跳過字串"
		Msg34 = "供覆審(&K)"
		Msg35 = "已驗證(&E)"
		Msg36 = "未翻譯(&N)"

		Msg37 = "其他選項"
		Msg38 = "繼續時自動儲存所有選取(&V)"
		Msg39 = "不建立或刪除檢查標記(&K)"
		Msg40 = "不變更原始翻譯狀態(&Y)"
		Msg41 = "自動替換特定字元(&L)"

		Msg42 = "關於(&A)"
		Msg43 = "說明(&H)"
		Msg44 = "設定(&S)"
		Msg45 = "儲存選取(&L)"

		Msg50 = "確認"
		Msg51 = "訊息"
		Msg52 = "錯誤"
		Msg53 =	"您的 Passolo 版本太低，本巨集僅適用於 Passolo 6.0 及以上版本，請升級後再使用。"
		Msg54 = "請選取一個翻譯清單！"
		Msg55 = "正在建立和更新翻譯清單..."
		Msg56 = "無法建立和更新翻譯清單，請檢查您的專案設定。"
		Msg57 = "該清單未被開啟。此狀態下可以檢查錯誤但無法修改字串。" & vbCrLf & _
				"您需要讓系統自動開啟該翻譯清單嗎？" & vbCrLf & vbCrLf & _
				"為了安全，開啟檢查後，如果找到錯誤將不會自動儲存修改" & vbCrLf & _
				"並關閉翻譯清單。否則將自動關閉翻譯清單。"
		Msg58 = "正在開啟翻譯清單..."
		Msg59 = "無法開啟翻譯清單，請檢查您的專案設定。"
		Msg60 = "該清單已處於開啟狀態。此狀態下進行錯誤檢查和修改將使" & vbCrLf & _
				"您未儲存的翻譯無法還原。為了安全，系統將先儲存您的翻" & vbCrLf & _
				"譯，然後進行檢查和修改。" & vbCrLf & vbCrLf & _
				"您確定要讓系統自動儲存您的翻譯嗎？"
		Msg62 = "正在檢查，可能需要幾分鐘，請稍候..."
		Msg65 = "原譯文: "
		Msg66 = "合計用時: "
		Msg67 = "hh 小時 mm 分 ss 秒"
		Msg70 = "英文到中文"
		Msg71 = "中文到英文"
		Msg72 = "M:程式版本"
		Msg73 = "M:設定名稱"
		Msg74 = "M:檢查狀態"
		Msg75 = "M:檢查日期"
		Msg76 = "有錯誤"
		Msg77 = "無錯誤"
		Msg78 = "已修正"
		Msg79 = "yyyy年m月d日 hh:mm:ss"
	Else
		Msg00 = "(c) 2007-2008 by gnatix, 2009-2010 by wanfu  唳掛: " & Version
		Msg01 = "辦豎瑩﹜笝砦睫睿樓厒ん潰脤粽"
		Msg02 = "掛最唗蚚衾潰脤﹜党蜊睿刉壺蜆楹祒蹈桶笢垀衄楹祒趼揹" & _
				"腔辦豎瑩﹜笝砦睫睿樓厒ん﹝③婓潰脤綴紨沭輛俴�佴元散憿�"
		Msg03 = "楹祒蹈桶: "

		Msg04 = "饜离恁寁"
		Msg05 = "赻雄粽:"
		Msg06 = "潰脤粽:"
		Msg07 = "赻雄粽睿潰脤粽眈肮(&U)"
		Msg08 = "赻雄恁寁(&X)"

		Msg09 = "潰脤梓暮"
		Msg10 = "綺謹唳掛(&B)"
		Msg11 = "綺謹饜离(&G)"
		Msg12 = "綺謹�梪�(&T)"
		Msg13 = "綺謹袨怓(&Z)"
		Msg14 = "�垓蕩鷈�(&Q)"

		Msg15 = "紱釬恁寁"
		Msg16 = "躺潰脤(&O)"
		Msg17 = "潰脤甜党淏(&M)"
		Msg18 = "刉壺辦豎瑩(&R)"

		Msg19 = "趼揹濬倰"
		Msg20 = "�垓�(&A)"
		Msg21 = "粕等(&M)"
		Msg22 = "勤趕遺(&D)"
		Msg23 = "趼揹桶(&S)"
		Msg24 = "樓厒ん桶(&A)"
		Msg25 = "唳掛(&V)"
		Msg26 = "む坻(&O)"
		Msg27 = "躺恁隅(&L)"

		Msg28 = "趼揹囀��"
		Msg29 = "�垓�(&F)"
		Msg30 = "辦豎瑩(&K)"
		Msg31 = "笝砦睫(&E)"
		Msg32 = "樓厒ん(&P)"

		Msg33 = "泐徹趼揹"
		Msg34 = "鼎葩机(&K)"
		Msg35 = "眒桄痐(&E)"
		Msg36 = "帤楹祒(&N)"

		Msg37 = "む坻恁砐"
		Msg38 = "樟哿奀赻雄悵湔垀衄恁寁(&V)"
		Msg39 = "祥斐膘麼刉壺潰脤梓暮(&K)"
		Msg40 = "祥載蜊埻宎楹祒袨怓(&Y)"
		Msg41 = "赻雄杸遙杻隅趼睫(&L)"

		Msg42 = "壽衾(&A)"
		Msg43 = "堆翑(&H)"
		Msg44 = "扢离(&S)"
		Msg45 = "悵湔恁寁(&L)"

		Msg50 = "�溜�"
		Msg51 = "陓洘"
		Msg52 = "渣昫"
		Msg53 =	"蠟腔 Passolo 唳掛怮腴ㄛ掛粽躺巠蚚衾 Passolo 6.0 摯眕奻唳掛ㄛ③汔撰綴婬妏蚚﹝"
		Msg54 = "③恁寁珨跺楹祒蹈桶ㄐ"
		Msg55 = "淏婓斐膘睿載陔楹祒蹈桶..."
		Msg56 = "拸楊斐膘睿載陔楹祒蹈桶ㄛ③潰脤蠟腔源偶扢离﹝"
		Msg57 = "蜆蹈桶帤掩湖羲﹝森袨怓狟褫眕潰脤渣昫筍拸楊党蜊趼揹﹝" & vbCrLf & _
				"蠟剒猁�譁腴啻堈秩翾疙繩倡蹅訇簋艞�" & vbCrLf & vbCrLf & _
				"峈賸假�咯炭翾盲麮暻鞶畏蝜�梑善渣昫蔚祥頗赻雄悵湔党蜊" & vbCrLf & _
				"甜壽敕楹祒蹈桶﹝瘁寀蔚赻雄壽敕楹祒蹈桶﹝"
		Msg58 = "淏婓湖羲楹祒蹈桶..."
		Msg59 = "拸楊湖羲楹祒蹈桶ㄛ③潰脤蠟腔源偶扢离﹝"
		Msg60 = "蜆蹈桶眒揭衾湖羲袨怓﹝森袨怓狟輛俴渣昫潰脤睿党蜊蔚妏" & vbCrLf & _
				"蠟帤悵湔腔楹祒拸楊遜埻﹝峈賸假�咯疢腴魚峙�悵湔蠟腔楹" & vbCrLf & _
				"祒ㄛ�遣騣靇邾麮暻迖瑏纂�" & vbCrLf & vbCrLf & _
				"蠟�毓例糾譁腴啻堈秧ㄣ磑�腔楹祒鎘ˋ"
		Msg62 = "淏婓潰脤ㄛ褫夔剒猁撓煦笘ㄛ③尕緊..."
		Msg65 = "埻祒恅: "
		Msg66 = "磁數蚚奀: "
		Msg67 = "hh 苤奀 mm 煦 ss 鏃"
		Msg70 = "荎恅善笢恅"
		Msg71 = "笢恅善荎恅"
		Msg72 = "M:最唗唳掛"
		Msg73 = "M:饜离靡備"
		Msg74 = "M:潰脤袨怓"
		Msg75 = "M:潰脤�梪�"
		Msg76 = "衄渣昫"
		Msg77 = "拸渣昫"
		Msg78 = "眒党淏"
		Msg79 = "yyyy爛m堎d�� hh:mm:ss"
	End If

	If PSL.Version < 600 Then
		MsgBox Msg53,vbOkOnly+vbInformation,Msg51
		Exit Sub
	End If

	Set trn = PSL.ActiveTransList
	'潰聆楹祒蹈桶岆瘁掩恁寁
	If trn Is Nothing Then
		MsgBox Msg54,vbOkOnly+vbInformation,Msg51
		Exit Sub
	End If

	'場宎趙杅郪
	ReDim DefaultCheckList(1),CheckList(0),CheckDataList(0)
	DefaultCheckList(0) = Msg70
	DefaultCheckList(1) = Msg71

	'黍�＝硒捎池篽髲�
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

	'鳳�☆�陔杅擂甜潰脤陔唳掛
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

	'勤趕遺
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
		CancelButton 510,434,90,21,.CancelButton '6 �＋�
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

	'鳳�＝硒挫閛邰暻�
	If dlg.Menu = 1 Then StrTypes = "|Menu|"
	If dlg.Dialog = 1 Then StrTypes = StrTypes & "|Dialog|"
	If dlg.Strings = 1 Then StrTypes = StrTypes & "|StringTable|"
	If dlg.AccTable = 1 Then StrTypes = StrTypes & "|AcceleratorTable|"
	If dlg.Versions = 1 Then StrTypes = StrTypes & "|Version|"

	'枑尨湖羲壽敕腔楹祒蹈桶ㄛ�蝜�剒猁党蜊睿刉壺
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

	'枑尨悵湔湖羲腔楹祒蹈桶ㄛ眕轎揭燴綴杅擂祥褫閥葩
	If trn.IsOpen = True And iVo <> 0 Then
		Massage = MsgBox(Msg03 & vbCrLf & vbCrLf & Msg60,vbYesNoCancel,Msg50)
 		If Massage = vbYes Then trn.Save
		If Massage = vbCancel Then GoTo ExitSub
	End If

	'�蝜�楹祒蹈桶腔載蜊奀潔俀衾埻宎蹈桶ㄛ赻雄載陔
	If trn.SourceList.LastChange > trn.LastChange Then
		PSL.Output Msg55
		If trn.Update = False Then
			MsgBox Msg56,vbOkOnly+vbInformation,Msg51
			GoTo ExitSub
		End If
	End If

	'扢离潰脤粽蚳蚚腔蚚誧隅砱扽俶
	If Not trn Is Nothing Then
		If trn.Property(19980) <> "CheckAccessKeys" Then
			trn.Property(19980) = "CheckAccessKeys"
		End If
	End If

	'鳳�·SL腔醴梓逄晟測鎢
	trnLng = PSL.GetLangCode(trn.Language.LangID,pslCode639_1)
	If trnLng = "zh" Or trnLng = "ja" Or trnLng = "ko" Then TranLang = "Asia"

	'庋溫祥婬妏蚚腔雄怓杅郪垀妏蚚腔囀湔
	Erase CheckListBak,CheckDataListBak,tempCheckList,tempCheckDataList

	'跦擂岆瘁恁寁 "躺恁隅趼揹" 砐扢离猁潰脤腔趼揹杅
	If dlg.Seleted = 0 Then
		StringCount = trn.StringCount
	Else
		StringCount = trn.StringCount(pslSelection)
	End If

	'羲宎趼揹紱釬
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
		'跦擂岆瘁恁寁 "躺恁隅趼揹" 砐扢离猁楹祒腔趼揹
		If dlg.Seleted = 0 Then Set TransString = trn.String(j)
		If dlg.Seleted = 1 Then Set TransString = trn.String(j,pslSelection)

		'秏洘睿趼揹場宎趙甜鳳�√倀贍芛倡鄶硒�
		Massage = ""
		LineMsg = ""
		AcckeyMsg = ""
		ChangeMsg = ""
		srcString = TransString.SourceText
		trnString = TransString.Text
		OldtrnString = trnString

		'刉壺潰脤梓暮
		If dlg.NoCheckTag = 1 Then TransString.Properties.RemoveAll

		'趼揹濬倰揭燴
		If dlg.AllType = 0 And dlg.Seleted = 0 Then
			If InStr("Menu|Dialog|StringTable|AcceleratorTable|Version",TransString.ResType) Then
				If InStr(StrTypes,TransString.ResType) = 0 Then GoTo Skip
			Else
				If dlg.Other = 0 Then GoTo Skip
			End If
		End If

		'泐徹帤楹祒﹜眒坶隅睿硐黍腔趼揹
		If iVo <> 2 Then
			If srcString = trnString Then GoTo Skip
			If TransString.State(pslStateTranslated) = False Then GoTo Skip
		End If
		If TransString.State(pslStateLocked) = True Then GoTo Skip
		If TransString.State(pslStateReadOnly) = True Then GoTo Skip
		If Trim(srcString) = "" Then GoTo Skip

		'泐徹眒潰脤拸渣昫甜й潰脤�梪硰縳皕倡躽梪痤儷硒�
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

		'羲宎揭燴趼揹
		NewtrnString = CheckHanding(CheckID,srcString,trnString,TranLang)
		If dlg.AutoRepStr = 1 Then NewtrnString = ReplaceStr(CheckID,NewtrnString,0)

		'覃蚚秏洘怀堤
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

		'杸遙麼蜇樓﹜刉壺楹祒趼揹
		If NewtrnString <> OldtrnString Then
			If iVo = 1 Then TransString.OutputError(Msg65 & OldtrnString)
			If iVo <> 0 Then TransString.Text = NewtrnString
		End If

		'載蜊睿埻恅祥珨祡腔祒恅趼揹袨怓,眕源晞脤艘
		If dlg.NoChangeState = 0 Then
			If NewtrnString <> OldtrnString Or srcLineNum <> trnLineNum Or srcAccKeyNum <> trnAccKeyNum Then
				TransString.State(pslStateReview) = True
			Else
				TransString.State(pslStateReview) = False
			End If
		End If

		'扢离趼揹腔赻隅砱潰脤扽俶睿潰脤�梪�
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

	'渣昫數杅摯秏洘怀堤
	ErrorCount = ModifiedCount + AddedCount + WarningCount + LineNumErrCount + accKeyNumErrCount
	PSL.Output CountMassage(ErrorCount,LineNumErrCount,accKeyNumErrCount)
	If ErrorCount = 0 And TransListOpen = True Then trn.Close
	EndTimes = Timer
	PSL.Output Msg66 & Format(DateAdd("s",EndTimes - StartTimes,0),Msg67)

	'�＋�潰脤粽蚳蚚腔蚚誧隅砱扽俶腔扢离
	ExitSub:
	If Not trn Is Nothing Then
		If trn.Property(19980) = "CheckAccessKeys" Then
			trn.Property(19980) = ""
		End If
	End If
	On Error GoTo 0
	Exit Sub

	'珆尨最唗渣昫秏洘
	SysErrorMsg:
	Call sysErrorMassage(Err)
	GoTo ExitSub
End Sub


'潰聆甜狟婥陔唳掛
Function Download(Method As String,Url As String,Async As String,Mode As String) As Boolean
	Dim i As Integer,n As Integer,m As Integer,k As Integer,updateINI As String
	Dim TempPath As String,File As String,OpenFile As Boolean,Body As Variant
	Dim xmlHttp As Object,UrlList() As String,Stemp As Boolean
	If OSLanguage = "0404" Then
		Msg01 = "錯誤"
		Msg02 = "更新失敗！"
		Msg03 = "測試失敗！"
		Msg04 = "系統沒有安裝 RAR 解壓縮套用程式！無法解壓縮下載檔案。"
		Msg05 = "缺少必要的參數，請檢查設定中的自動更新參數設定。"
		Msg06 = "開啟更新網址失敗！請檢查網址是否正確或可存取。"
		Msg08 = "無法獲取訊息！請檢查網址是否正確或可存取。"
		Msg09 = "無法解壓縮檔案！請檢查解壓程式或解壓參數是否正確。"
		Msg10 = "程式名稱: "
		Msg11 = "剖析路徑: "
		Msg12 = "執行參數: "
		Msg13 = "RAR 解壓縮程式未找到！可能是程式路徑錯誤或已被移除。"
		Msg14 = "自動更新測試"
		Msg15 = "測試成功！更新網址和解壓程式參數正確。"
		Msg16 = "目前網上的版本為 %s。"
		Msg17 = "發現有可用的新版本 - %s，是否下載更新？"
		Msg18 = "訊息"
		Msg19 = "確認"
		Msg20 = "更新成功！程式將結束，結束後請重新啟動程式。"
		Msg21 = "獲取的訊息不包含版本訊息！請檢查請檢查網址是否正確。"
		Msg22 = "更新內容:"
		Msg23 = "是否需要重新下載並更新？"
		Msg24 = "正在測試更新網址和解壓程式參數，請稍候..."
		Msg25 = "正在檢查新版本，請稍候..."
		Msg26 = "正在下載新版本，請稍候..."
		Msg27 = "正在解壓縮，請稍候..."
		Msg28 = "正在確認下載的程式版本，請稍候..."
		Msg29 = "正在安裝新版本，請稍候..."
		Msg30 = "您的系統缺少 Microsoft.XMLHTTP 物件，無法更新！"
	Else
		Msg01 = "渣昫"
		Msg02 = "載陔囮啖ㄐ"
		Msg03 = "聆彸囮啖ㄐ"
		Msg04 = "炵苀羶衄假蚾 RAR 賤揤坫茼蚚最唗ㄐ拸楊賤揤坫狟婥恅璃﹝"
		Msg05 = "�捻棱寋玥觸恀�ㄛ③潰脤饜离笢腔赻雄載陔統杅扢离﹝"
		Msg06 = "湖羲載陔厙硊囮啖ㄐ③潰脤厙硊岆瘁淏�溶翾伢襞吽�"
		Msg08 = "拸楊鳳�－欐╯﹊趧麮樿鱣滔Й鵙��溶翾伢襞吽�"
		Msg09 = "拸楊賤揤坫恅璃ㄐ③潰脤賤揤最唗麼賤揤統杅岆瘁淏�楚�"
		Msg10 = "最唗靡備: "
		Msg11 = "賤昴繚噤: "
		Msg12 = "堍俴統杅: "
		Msg13 = "RAR 賤揤坫最唗帤梑善ㄐ褫夔岆最唗繚噤渣昫麼眒掩迠婥﹝"
		Msg14 = "赻雄載陔聆彸"
		Msg15 = "聆彸傖髡ㄐ載陔厙硊睿賤揤最唗統杅淏�楚�"
		Msg16 = "醴ヶ厙奻腔唳掛峈 %s﹝"
		Msg17 = "楷珋衄褫蚚腔陔唳掛 - %sㄛ岆瘁狟婥載陔ˋ"
		Msg18 = "秏洘"
		Msg19 = "�溜�"
		Msg20 = "載陔傖髡ㄐ最唗蔚豖堤ㄛ豖堤綴③笭陔ゐ雄最唗﹝"
		Msg21 = "鳳�△鹹欐３趕�漪唳掛陓洘ㄐ③潰脤③潰脤厙硊岆瘁淏�楚�"
		Msg22 = "載陔囀��:"
		Msg23 = "岆瘁剒猁笭陔狟婥甜載陔ˋ"
		Msg24 = "淏婓聆彸載陔厙硊睿賤揤最唗統杅ㄛ③尕緊..."
		Msg25 = "淏婓潰脤陔唳掛ㄛ③尕緊..."
		Msg26 = "淏婓狟婥陔唳掛ㄛ③尕緊..."
		Msg27 = "淏婓賤揤坫ㄛ③尕緊..."
		Msg28 = "淏婓�溜珫謹媯議昐繵瘙麾甭輶埏�..."
		Msg29 = "淏婓假蚾陔唳掛ㄛ③尕緊..."
		Msg30 = "蠟腔炵苀�捻� Microsoft.XMLHTTP 勤砓ㄛ拸楊載陔ㄐ"
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


'植蛁聊桶笢鳳�� RAR 孺桯靡腔蘇�炡昐�
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


'蛌遙媼輛秶杅擂峈硌隅晤鎢跡宒腔趼睫
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


'迡�賱�輛秶杅擂善恅璃
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


'潰脤趼揹岆瘁躺婦漪杅趼睿睫瘍
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


'杸遙杻隅趼睫
Function ReplaceStr(CheckID As Integer,trnStr As String,fType As Integer) As String
	Dim i As Integer,BaktrnStr As String
	'鳳�×▲乳馺繭觸恀�
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


'潰脤党淏辦豎瑩﹜笝砦睫睿樓厒ん
Function CheckHanding(CheckID As Integer,srcStr As String,trnStr As String,TranLang As String) As String
	Dim i As Integer,BaksrcStr As String,BaktrnStr As String,srcStrBak As String,trnStrBak As String
	Dim srcNum As Integer,trnNum As Integer,srcSplitNum As Integer,trnSplitNum As Integer
	Dim FindStrArr() As String,srcStrArr() As String,trnStrArr() As String,LineSplitArr() As String
	Dim posinSrc As Integer,posinTrn As Integer

	'鳳�×▲乳馺繭觸恀�
	TempArray = Split(CheckDataList(CheckID),JoinStr)
	SetsArray = Split(TempArray(1),SubJoinStr)
	ExcludeChar = SetsArray(0)
	LineSplitChar = SetsArray(1)
	KeepCharPair = SetsArray(3)

	'偌趼睫酗僅齬唗
	If LineSplitChar <> "" Then
		FindStrArr = Split(LineSplitChar,",",-1)
		FindStrArr = SortArray(FindStrArr,0,"Lenght",">")
		LineSplitChar = Join(FindStrArr,",")
	End If

	'統杅場宎趙
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

	'齬壺趼揹笢腔準辦豎瑩
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

	'徹薦祥岆辦豎瑩腔辦豎瑩
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

	'蚚杸遙楊莞煦趼揹
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

	'趼揹揭燴
	srcSplitNum = UBound(srcStrArr)
	trnSplitNum = UBound(trnStrArr)
	If srcSplitNum = 0 And trnSplitNum = 0 Then
		BaktrnStr = StringReplace(CheckID,BaksrcStr,BaktrnStr,TranLang)
	ElseIf srcSplitNum <> 0 Or trnSplitNum <> 0 Then
		LineSplitArr = MergeArray(srcStrArr,trnStrArr)
		BaktrnStr = ReplaceStrSplit(CheckID,BaktrnStr,LineSplitArr,TranLang)
	End If

	'數呾俴杅
	LineSplitChars = "\r\n,\r,\n"
	FindStrArr = Split(Convert(LineSplitChars),",",-1)
	For i = LBound(FindStrArr) To UBound(FindStrArr)
		FindStr = Trim(FindStrArr(i))
		If InStr(BaksrcStr,FindStr) Then srcLineNum = UBound(Split(BaksrcStr,FindStr,-1))
		If InStr(BaktrnStr,FindStr) Then trnLineNum = UBound(Split(BaktrnStr,FindStr,-1))
	Next i

	'數呾辦豎瑩杅
	If InStr(BaksrcStr,"&") Then srcAccKeyNum = UBound(Split(BaksrcStr,"&",-1))
	If InStr(BaktrnStr,"&") Then trnAccKeyNum = UBound(Split(BaktrnStr,"&",-1))

	'遜埻祥岆辦豎瑩腔辦豎瑩
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

	'遜埻趼揹笢掩齬壺腔準辦豎瑩
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


' 偌俴鳳�＝硒挾譫鷒囆硍帣十皛遠倡鄶硊�揹
Function StringReplace(CheckID As Integer,srcStr As String,trnStr As String,TranLang As String) As String
	Dim posinSrc As Integer,posinTrn As Integer,StringSrc As String,StringTrn As String
	Dim accesskeySrc As String,accesskeyTrn As String,Temp As String
	Dim ShortcutPosSrc As Integer,ShortcutPosTrn As Integer,PreTrn As String
	Dim EndStringPosSrc As Integer,EndStringPosTrn As Integer,AppTrn As String
	Dim preKeyTrn As String,appKeyTrn As String,Stemp As Boolean,FindStrArr() As String
	Dim i As Integer,j As Integer,x As Integer,y As Integer,m As Integer,n As Integer

	'鳳�×▲乳馺繭觸恀�
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

	'偌趼睫酗僅齬唗
	If CheckEndChar <> "" Then
		FindStrArr = Split(CheckEndChar,-1)
		FindStrArr = SortArray(FindStrArr,0,"Lenght",">")
		CheckEndChar = Join(FindStrArr)
	End If

	'統杅場宎趙
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

	'枑�＝硒效或窒楖�
	EndSpaceSrc = Space(Len(srcStr) - Len(RTrim(srcStr)))
	EndSpaceTrn = Space(Len(trnStr) - Len(RTrim(trnStr)))

	'鳳�□蚎棐�
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

	'鳳�＞欶僩�ㄛ狟蹈趼睫歙頗掩潰脤
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

	'鳳�√倀贍芛倡賮醴儠敯�弇离
	posinSrc = InStrRev(srcStr,"&")
	posinTrn = InStrRev(trnStr,"&")

	'鳳�√倀贍芛倡賮醴儠敯� (婦嬤辦豎瑩睫ヶ綴腔趼睫)
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

	'鳳�▼儠敯�綴醱腔準笝砦睫睿準樓厒ん腔趼睫ㄛ涴虳趼睫蔚掩痄雄善辦豎瑩ヶ
	If posinTrn <> 0 Then
		x = Len(EndStringTrn & ShortcutTrn & EndSpaceTrn)
		If InStr(ShortcutTrn,"&") Then x = Len(EndSpaceTrn)
		If InStr(EndStringTrn,"&") Then x = Len(ShortcutTrn & EndSpaceTrn)
		If Len(trnStr) > x Then
			Temp = Left(trnStr,Len(trnStr) - x)
			ExpStringTrn = Mid(Temp,posinTrn + Len(acckeyTrn))
		End If
	End If

	'�戊�辦豎瑩麼笝砦睫麼樓厒んヶ醱腔諾跡
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

	'鳳�◎倡鄵郈儠敯�ヶ腔笝砦睫
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

	'赻雄楹祒睫磁沭璃腔笝砦睫
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

	'猁掩悵隱腔笝砦睫郪磁
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

	'悵隱睫磁沭璃腔樓厒ん楹祒
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

	'趼揹囀�楪√騑池�
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

	'杅擂摩傖
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

	'PSL.Output "------------------------------ "      '覃彸蚚
	'PSL.Output "srcStr = " & srcStr                   '覃彸蚚
	'PSL.Output "trnStr = " & trnStr                   '覃彸蚚
	'PSL.Output "SpaceTrn = " & SpaceTrn               '覃彸蚚
	'PSL.Output "acckeySrc = " & acckeySrc             '覃彸蚚
	'PSL.Output "acckeyTrn = " & acckeyTrn             '覃彸蚚
	'PSL.Output "EndStringSrc = " & EndStringSrc       '覃彸蚚
	'PSL.Output "EndStringTrn = " & EndStringTrn       '覃彸蚚
	'PSL.Output "ShortcutSrc = " & ShortcutSrc         '覃彸蚚
	'PSL.Output "ShortcutTrn = " & ShortcutTrn         '覃彸蚚
	'PSL.Output "ExpStringTrn = " & ExpStringTrn       '覃彸蚚
	'PSL.Output "StringSrc = " & StringSrc             '覃彸蚚
	'PSL.Output "StringTrn = " & StringTrn             '覃彸蚚
	'PSL.Output "NewStringTrn = " & NewStringTrn       '覃彸蚚
	'PSL.Output "PreStringTrn = " & PreStringTrn       '覃彸蚚

	'趼揹杸遙
	Temp = trnStr
	If StringSrc <> StringTrn Then
		If StringTrn <> "" And StringTrn <> NewStringTrn Then
			x = InStrRev(Temp,StringTrn)
			If x <> 0 Then
				PreTrn = Left(Temp,x - 1)
				AppTrn = Mid(Temp,x)
				'PSL.Output "PreTrn = " & PreTrn       '覃彸蚚
				'PSL.Output "AppTrn = " & AppTrn       '覃彸蚚
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


' 党蜊秏洘怀堤
Function ReplaceMassage(OldtrnString As String,NewtrnString As String) As String
	Dim AcckeyMsg As String,EndStringMsg As String,ShortcutMsg As String,Tmsg1 As String
	Dim Tmsg2 As String,Fmsg As String,Smsg As String,Massage1 As String,Massage2 As String
	Dim Massage3 As String,Massage4 As String,n As Integer,x As Integer,y As Integer

	If OSLanguage = "0404" And iVo = 0 Then
		Msg00 = "便捷鍵與原文不同"
		Msg01 = "終止符與原文不同或者需要翻譯"
		Msg02 = "加速器與原文不同"
		Msg03 = "便捷鍵和終止符與原文不同"
		Msg04 = "終止符和加速器與原文不同"
		Msg05 = "便捷鍵和加速器與原文不同"
		Msg06 = "便捷鍵、終止符和加速器與原文不同"
		Msg07 = "便捷鍵的大小寫不同"
		Msg08 = "便捷鍵的大小寫和終止符與原文不同"
		Msg09 = "便捷鍵的大小寫和加速器與原文不同"
		Msg10 = "便捷鍵的大小寫、終止符和加速器與原文不同"

		Msg11 = "譯文中缺少便捷鍵"
		Msg12 = "譯文中缺少終止符"
		Msg13 = "譯文中缺少加速器"
		Msg14 = "譯文中缺少便捷鍵和終止符"
		Msg15 = "譯文中缺少終止符和加速器"
		Msg16 = "譯文中缺少便捷鍵和加速器"
		Msg17 = "譯文中缺少便捷鍵、終止符和加速器"

		Msg21 = "僅譯文中有便捷鍵"
		Msg22 = "僅譯文中有終止符"
		Msg23 = "僅譯文中有加速器"
		Msg24 = "僅譯文中有便捷鍵和終止符"
		Msg25 = "僅譯文中有終止符和加速器"
		Msg26 = "僅譯文中有便捷鍵和加速器"
		Msg27 = "僅譯文中有便捷鍵、終止符和加速器"

		Msg31 = "需移動便捷鍵到最後"
		Msg32 = "需移動便捷鍵到終止符前"
		Msg33 = "需移動便捷鍵到加速器前"

		Msg41 = "便捷鍵前有空格"
		Msg42 = "終止符前有空格"
		Msg43 = "加速器前有空格"
		Msg44 = "便捷鍵前有終止符"
		Msg45 = "便捷鍵前有空格和終止符"

		Msg46 = "終止符前的空格比原文少"
		Msg47 = "終止符後的空格比原文少"
		Msg48 = "終止符前後的空格比原文少"
		Msg49 = "字串後的空格比原文少"
		Msg50 = "終止符前和字串後的空格比原文少"
		Msg51 = "終止符後和字串後的空格比原文少"
		Msg52 = "終止符前後和字串後的空格比原文少"
		Msg53 = "終止符前有多餘的空格"
		Msg54 = "終止符後有多餘的空格"
		Msg55 = "終止符前後有多餘的空格"
		Msg56 = "字串後有多餘的空格"
		Msg57 = "終止符前和字串後有多餘的空格"
		Msg58 = "終止符後和字串後有多餘的空格"
		Msg59 = "終止符前後和字串後有多餘的空格"

		Msg61 = "，"
		Msg62 = "並"
		Msg63 = "。"
		Msg64 = "、"
		Msg65 = "已將 "
		Msg66 = " 替換為 "
		Msg67 = "已刪除 "
	ElseIf OSLanguage = "0404" And iVo = 1 Then
		Msg00 = "已修改了便捷鍵"
		Msg01 = "已修改了終止符"
		Msg02 = "已修改了加速器"
		Msg03 = "已修改了便捷鍵和終止符"
		Msg04 = "已修改了終止符和加速器"
		Msg05 = "已修改了便捷鍵和加速器"
		Msg06 = "已修改了便捷鍵、終止符和加速器"
		Msg07 = "已修改了便捷鍵的大小寫"
		Msg08 = "已修改了便捷鍵的大小寫和終止符"
		Msg09 = "已修改了便捷鍵的大小寫和加速器"
		Msg10 = "已修改了便捷鍵的大小寫、終止符和加速器"

		Msg11 = "已新增了便捷鍵"
		Msg12 = "已新增了終止符"
		Msg13 = "已新增了加速器"
		Msg14 = "已新增了便捷鍵和終止符"
		Msg15 = "已新增了終止符和加速器"
		Msg16 = "已新增了便捷鍵和加速器"
		Msg17 = "已新增了便捷鍵、終止符和加速器"

		Msg21 = "僅譯文中有便捷鍵"
		Msg22 = "僅譯文中有終止符"
		Msg23 = "僅譯文中有加速器"
		Msg24 = "僅譯文中有便捷鍵和終止符"
		Msg25 = "僅譯文中有終止符和加速器"
		Msg26 = "僅譯文中有便捷鍵和加速器"
		Msg27 = "僅譯文中有便捷鍵、終止符和加速器"

		Msg31 = "已移動便捷鍵到最後"
		Msg32 = "已移動便捷鍵到終止符前"
		Msg33 = "已移動便捷鍵到加速器前"

		Msg41 = "已去除了便捷鍵前的空格"
		Msg42 = "已去除了終止符前的空格"
		Msg43 = "已去除了加速器前的空格"
		Msg44 = "已去除了便捷鍵前的終止符"
		Msg45 = "已去除了便捷鍵前的空格和終止符"

		Msg46 = "已新增了終止符前缺少的空格"
		Msg47 = "已新增了終止符後缺少的空格"
		Msg48 = "已新增了終止符前後缺少的空格"
		Msg49 = "已新增了字串後缺少的空格"
		Msg50 = "已新增了終止符前和字串後缺少的空格"
		Msg51 = "已新增了終止符後和字串後缺少的空格"
		Msg52 = "已新增了終止符前後和字串後缺少的空格"
		Msg53 = "已去除了終止符前多餘的空格"
		Msg54 = "已去除了終止符後多餘的空格"
		Msg55 = "已去除了終止符前後多餘的空格"
		Msg56 = "已去除了字串後多餘的空格"
		Msg57 = "已去除了終止符前和字串後多餘的空格"
		Msg58 = "已去除了終止符後和字串後多餘的空格"
		Msg59 = "已去除了終止符前後和字串後多餘的空格"

		Msg61 = "，"
		Msg62 = "並"
		Msg63 = "。"
		Msg64 = "、"
		Msg65 = "已將 "
		Msg66 = " 替換為 "
		Msg67 = "已刪除 "
	ElseIf OSLanguage <> "0404" And iVo = 0 Then
		Msg00 = "辦豎瑩迵埻恅祥肮"
		Msg01 = "笝砦睫迵埻恅祥肮麼氪剒猁楹祒"
		Msg02 = "樓厒ん迵埻恅祥肮"
		Msg03 = "辦豎瑩睿笝砦睫迵埻恅祥肮"
		Msg04 = "笝砦睫睿樓厒ん迵埻恅祥肮"
		Msg05 = "辦豎瑩睿樓厒ん迵埻恅祥肮"
		Msg06 = "辦豎瑩﹜笝砦睫睿樓厒ん迵埻恅祥肮"
		Msg07 = "辦豎瑩腔湮苤迡祥肮"
		Msg08 = "辦豎瑩腔湮苤迡睿笝砦睫迵埻恅祥肮"
		Msg09 = "辦豎瑩腔湮苤迡睿樓厒ん迵埻恅祥肮"
		Msg10 = "辦豎瑩腔湮苤迡﹜笝砦睫睿樓厒ん迵埻恅祥肮"

		Msg11 = "祒恅笢�捻楰儠敯�"
		Msg12 = "祒恅笢�捻棩欶僩�"
		Msg13 = "祒恅笢�捻椇蚎棐�"
		Msg14 = "祒恅笢�捻楰儠敯�睿笝砦睫"
		Msg15 = "祒恅笢�捻棩欶僩�睿樓厒ん"
		Msg16 = "祒恅笢�捻楰儠敯�睿樓厒ん"
		Msg17 = "祒恅笢�捻楰儠敯�﹜笝砦睫睿樓厒ん"

		Msg21 = "躺祒恅笢衄辦豎瑩"
		Msg22 = "躺祒恅笢衄笝砦睫"
		Msg23 = "躺祒恅笢衄樓厒ん"
		Msg24 = "躺祒恅笢衄辦豎瑩睿笝砦睫"
		Msg25 = "躺祒恅笢衄笝砦睫睿樓厒ん"
		Msg26 = "躺祒恅笢衄辦豎瑩睿樓厒ん"
		Msg27 = "躺祒恅笢衄辦豎瑩﹜笝砦睫睿樓厒ん"

		Msg31 = "剒痄雄辦豎瑩善郔綴"
		Msg32 = "剒痄雄辦豎瑩善笝砦睫ヶ"
		Msg33 = "剒痄雄辦豎瑩善樓厒んヶ"

		Msg41 = "辦豎瑩ヶ衄諾跡"
		Msg42 = "笝砦睫ヶ衄諾跡"
		Msg43 = "樓厒んヶ衄諾跡"
		Msg44 = "辦豎瑩ヶ衄笝砦睫"
		Msg45 = "辦豎瑩ヶ衄諾跡睿笝砦睫"

		Msg46 = "笝砦睫ヶ腔諾跡掀埻恅屾"
		Msg47 = "笝砦睫綴腔諾跡掀埻恅屾"
		Msg48 = "笝砦睫ヶ綴腔諾跡掀埻恅屾"
		Msg49 = "趼揹綴腔諾跡掀埻恅屾"
		Msg50 = "笝砦睫ヶ睿趼揹綴腔諾跡掀埻恅屾"
		Msg51 = "笝砦睫綴睿趼揹綴腔諾跡掀埻恅屾"
		Msg52 = "笝砦睫ヶ綴睿趼揹綴腔諾跡掀埻恅屾"
		Msg53 = "笝砦睫ヶ衄嗣豻腔諾跡"
		Msg54 = "笝砦睫綴衄嗣豻腔諾跡"
		Msg55 = "笝砦睫ヶ綴衄嗣豻腔諾跡"
		Msg56 = "趼揹綴衄嗣豻腔諾跡"
		Msg57 = "笝砦睫ヶ睿趼揹綴衄嗣豻腔諾跡"
		Msg58 = "笝砦睫綴睿趼揹綴衄嗣豻腔諾跡"
		Msg59 = "笝砦睫ヶ綴睿趼揹綴衄嗣豻腔諾跡"

		Msg61 = "ㄛ"
		Msg62 = "甜"
		Msg63 = "﹝"
		Msg64 = "﹜"
		Msg65 = "眒蔚 "
		Msg66 = " 杸遙峈 "
		Msg67 = "眒刉壺 "
	ElseIf OSLanguage <> "0404" And iVo = 1 Then
		Msg00 = "眒党蜊賸辦豎瑩"
		Msg01 = "眒党蜊賸笝砦睫"
		Msg02 = "眒党蜊賸樓厒ん"
		Msg03 = "眒党蜊賸辦豎瑩睿笝砦睫"
		Msg04 = "眒党蜊賸笝砦睫睿樓厒ん"
		Msg05 = "眒党蜊賸辦豎瑩睿樓厒ん"
		Msg06 = "眒党蜊賸辦豎瑩﹜笝砦睫睿樓厒ん"
		Msg07 = "眒党蜊賸辦豎瑩腔湮苤迡"
		Msg08 = "眒党蜊賸辦豎瑩腔湮苤迡睿笝砦睫"
		Msg09 = "眒党蜊賸辦豎瑩腔湮苤迡睿樓厒ん"
		Msg10 = "眒党蜊賸辦豎瑩腔湮苤迡﹜笝砦睫睿樓厒ん"

		Msg11 = "眒氝樓賸辦豎瑩"
		Msg12 = "眒氝樓賸笝砦睫"
		Msg13 = "眒氝樓賸樓厒ん"
		Msg14 = "眒氝樓賸辦豎瑩睿笝砦睫"
		Msg15 = "眒氝樓賸笝砦睫睿樓厒ん"
		Msg16 = "眒氝樓賸辦豎瑩睿樓厒ん"
		Msg17 = "眒氝樓賸辦豎瑩﹜笝砦睫睿樓厒ん"

		Msg21 = "躺祒恅笢衄辦豎瑩"
		Msg22 = "躺祒恅笢衄笝砦睫"
		Msg23 = "躺祒恅笢衄樓厒ん"
		Msg24 = "躺祒恅笢衄辦豎瑩睿笝砦睫"
		Msg25 = "躺祒恅笢衄笝砦睫睿樓厒ん"
		Msg26 = "躺祒恅笢衄辦豎瑩睿樓厒ん"
		Msg27 = "躺祒恅笢衄辦豎瑩﹜笝砦睫睿樓厒ん"

		Msg31 = "眒痄雄辦豎瑩善郔綴"
		Msg32 = "眒痄雄辦豎瑩善笝砦睫ヶ"
		Msg33 = "眒痄雄辦豎瑩善樓厒んヶ"

		Msg41 = "眒�戊�賸辦豎瑩ヶ腔諾跡"
		Msg42 = "眒�戊�賸笝砦睫ヶ腔諾跡"
		Msg43 = "眒�戊�賸樓厒んヶ腔諾跡"
		Msg44 = "眒�戊�賸辦豎瑩ヶ腔笝砦睫"
		Msg45 = "眒�戊�賸辦豎瑩ヶ腔諾跡睿笝砦睫"

		Msg46 = "眒氝樓賸笝砦睫ヶ�捻棫醴楖�"
		Msg47 = "眒氝樓賸笝砦睫綴�捻棫醴楖�"
		Msg48 = "眒氝樓賸笝砦睫ヶ綴�捻棫醴楖�"
		Msg49 = "眒氝樓賸趼揹綴�捻棫醴楖�"
		Msg50 = "眒氝樓賸笝砦睫ヶ睿趼揹綴�捻棫醴楖�"
		Msg51 = "眒氝樓賸笝砦睫綴睿趼揹綴�捻棫醴楖�"
		Msg52 = "眒氝樓賸笝砦睫ヶ綴睿趼揹綴�捻棫醴楖�"
		Msg53 = "眒�戊�賸笝砦睫ヶ嗣豻腔諾跡"
		Msg54 = "眒�戊�賸笝砦睫綴嗣豻腔諾跡"
		Msg55 = "眒�戊�賸笝砦睫ヶ綴嗣豻腔諾跡"
		Msg56 = "眒�戊�賸趼揹綴嗣豻腔諾跡"
		Msg57 = "眒�戊�賸笝砦睫ヶ睿趼揹綴嗣豻腔諾跡"
		Msg58 = "眒�戊�賸笝砦睫綴睿趼揹綴嗣豻腔諾跡"
		Msg59 = "眒�戊�賸笝砦睫ヶ綴睿趼揹綴嗣豻腔諾跡"

		Msg61 = "ㄛ"
		Msg62 = "甜"
		Msg63 = "﹝"
		Msg64 = "﹜"
		Msg65 = "眒蔚 "
		Msg66 = " 杸遙峈 "
		Msg67 = "眒刉壺 "
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


' 刉壺辦豎瑩秏洘怀堤
Function DelAccKeyMassage(OldtrnString As String,NewtrnString As String) As String
	If OSLanguage = "0404" Then
		Msg01 = "已刪除和原文不同的便捷鍵 (%s)"
		Msg02 = "已刪除和原文相同的便捷鍵 (%s)"
		Msg03 = "已刪除僅在譯文中存在的便捷鍵 (%s)"
	Else
		Msg01 = "眒刉壺睿埻恅祥肮腔辦豎瑩 (%s)"
		Msg02 = "眒刉壺睿埻恅眈肮腔辦豎瑩 (%s)"
		Msg03 = "眒刉壺躺婓祒恅笢湔婓腔辦豎瑩 (%s)"
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


'怀堤俴杅渣昫秏洘
Function LineErrMassage(srcLineNum As Integer,trnLineNum As Integer,LineNumErrCount As Integer) As String
	If OSLanguage = "0404" Then
		Msg01 = "譯文的行數比原文少 %s 行。"
		Msg02 = "譯文的行數比原文多 %s 行。"
	Else
		Msg01 = "祒恅腔俴杅掀埻恅屾 %s 俴﹝"
		Msg02 = "祒恅腔俴杅掀埻恅嗣 %s 俴﹝"
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


'怀堤辦豎瑩杅渣昫秏洘
Function AccKeyErrMassage(srcAccKeyNum As Integer,trnAccKeyNum As Integer,accKeyNumErrCount As Integer) As String
	If OSLanguage = "0404" Then
		Msg01 = "譯文的便捷鍵數比原文少 %s 個。"
		Msg02 = "譯文的便捷鍵數比原文多 %s 個。"
	Else
		Msg01 = "祒恅腔辦豎瑩杅掀埻恅屾 %s 跺﹝"
		Msg02 = "祒恅腔辦豎瑩杅掀埻恅嗣 %s 跺﹝"
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


'渣昫數杅秏洘怀堤
Function CountMassage(ErrorCount As Integer,LineNumErrCount As Integer,accKeyNumErrCount As Integer) As String
	If OSLanguage = "0404" Then
		Msg01 = "沒有找到錯誤。"
		Msg02 = "找到 " & ErrorCount & " 個錯誤。其中: "
		Msg03 = "和原文不同 " & ModifiedCount & " 個，" & _
				"譯文中缺少 " & AddedCount & " 個，" & _
				"僅在譯文中存在 " & WarningCount & " 個，"
		Msg04 = "原文和譯文中行數不同 " & LineNumErrCount & " 個，" & _
				"原文和譯文中便捷鍵數不同 " & accKeyNumErrCount & " 個。"
		Msg06 = "已修改 " & ModifiedCount & " 個，已新增 " & AddedCount & " 個，" & _
				"要檢查 " & WarningCount & " 個，行數不同 " & LineNumErrCount & " 個，" & _
				"便捷鍵數不同 " & accKeyNumErrCount & " 個。"
		Msg07 = "沒有刪除任何便捷鍵。"
		Msg08 = "已刪除 " & ErrorCount & " 個便捷鍵。其中: "
		Msg09 = "和原文不同 " & ModifiedCount & " 個，" & _
				"和原文相同 " & AddedCount & " 個，" & _
				"僅在譯文中存在 " & WarningCount & " 個。"
	Else
		Msg01 = "羶衄梑善渣昫﹝"
		Msg02 = "梑善 " & ErrorCount & " 跺渣昫﹝む笢: "
		Msg03 = "睿埻恅祥肮 " & ModifiedCount & " 跺ㄛ" & _
				"祒恅笢�捻� " & AddedCount & " 跺ㄛ" & _
				"躺婓祒恅笢湔婓 " & WarningCount & " 跺ㄛ"
		Msg04 = "埻恅睿祒恅笢俴杅祥肮 " & LineNumErrCount & " 跺ㄛ" & _
				"埻恅睿祒恅笢辦豎瑩杅祥肮 " & accKeyNumErrCount & " 跺﹝"
		Msg06 = "眒党蜊 " & ModifiedCount & " 跺ㄛ眒氝樓 " & AddedCount & " 跺ㄛ" & _
				"猁潰脤 " & WarningCount & " 跺ㄛ俴杅祥肮 " & LineNumErrCount & " 跺ㄛ" & _
				"辦豎瑩杅祥肮 " & accKeyNumErrCount & " 跺﹝"
		Msg07 = "羶衄刉壺�庥怷儠敯�﹝"
		Msg08 = "眒刉壺 " & ErrorCount & " 跺辦豎瑩﹝む笢: "
		Msg09 = "睿埻恅祥肮 " & ModifiedCount & " 跺ㄛ" & _
				"睿埻恅眈肮 " & AddedCount & " 跺ㄛ" & _
				"躺婓祒恅笢湔婓 " & WarningCount & " 跺﹝"
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


'怀堤最唗渣昫秏洘
Sub sysErrorMassage(sysError As ErrObject)
	If OSLanguage = "0404" Then
		Msg01 = "錯誤"
		Msg02 = "發生程式設計上的錯誤。錯誤代碼 "
	Else
		Msg01 = "渣昫"
		Msg02 = "楷汜最唗扢數奻腔渣昫﹝渣昫測鎢 "
	End If
	MsgBox(msg02 & CStr(sysError.Number) & ", " & sysError.Description,vbOkOnly+vbInformation,Msg01)
End Sub


'婓辦豎瑩綴脣�輲媔侈硊�甜眕森莞煦趼揹
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
	'PSL.Output "bakString = " & bakString       '覃彸蚚
End Function


'輛俴杅郪磁甜
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


'趼揹都杅淏砃蛌遙
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


'蛌遙匐輛秶麼坋鞠輛秶蛌砱睫
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


' 蚚芴: 蔚匐輛秶蛌趙峈坋輛秶
' 怀��: OctStr(匐輛秶杅)
' 怀�輮�擂濬倰ㄩString
' 怀堤: OCTtoDEC(坋輛秶杅)
' 怀堤杅擂濬倰: Long
' 怀�鄶豱鯞�峈17777777777,怀堤郔湮杅峈2147483647
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


' 蚚芴: 蔚坋鞠輛秶蛌趙峈坋輛秶
' 怀��: HexStr(坋鞠輛秶杅)
' 怀�輮�擂濬倰: String
' 怀堤: HEXtoDEC(坋輛秶杅)
' 怀堤杅擂濬倰: Long
' 怀�鄶豱鯞�峈7FFFFFFF,怀堤郔湮杅峈2147483647
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


'蛌遙趼睫峈淕杅杅硉
Function StrToInteger(mStr As String) As Integer
	If mStr = "" Then mStr = "0"
	StrToInteger = CInt(mStr)
End Function


'黍�﹋�郪笢腔藩跺趼揹甜杸遙揭燴
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

		'揭燴婓ヶ綴俴笢婦漪腔笭葩趼睫
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

		'勤藩俴腔杅擂輛俴蟀諉ㄛ蚚衾秏洘怀堤
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

	'峈覃蚚秏洘怀堤ㄛ蚚埻衄曹講杸遙蟀諉綴腔杅擂
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


'赻隅砱統杅
Function Settings(CheckID As Integer) As Integer
	Dim AppLngList() As String,UseLngList() As String

	If OSLanguage = "0404" Then
		Msg01 = "設定"
		Msg02 = "請選取設定並在下列文字方塊中輸入值，測試無誤後再套用於實際操作。"
		Msg05 = "讀取(&R)"
		Msg06 = "新增(&A)"
		Msg07 = "變更(&M)"
		Msg08 = "刪除(&D)"
		Msg09 = "清空(&C)"
		Msg10 = "測試(&T)"

		Msg11 = "設定清單"
		Msg12 = "儲存類型"
		Msg13 = "設定內容"
		Msg14 = "檔案"
		Msg15 = "註冊表"
		Msg16 = "匯入設定"
		Msg17 = "匯出設定"
		Msg18 = "便捷鍵"
		Msg19 = "終止符"
		Msg20 = "加速器"
		Msg21 = "字元替換"

		Msg22 = "要排除的含 && 符號的非便捷鍵字元組合 (半形逗號分隔):"
		Msg23 = "字串分割用旗標符 (用於多個便捷鍵字串的處理) (半形逗號分隔):"
		Msg24 = "要檢查的便捷鍵前後括號，例如 [&&F] (半形逗號分隔):"
		Msg25 = "要保留的非便捷鍵符前後成對字元，例如 (&&) (半形逗號分隔):"
		Msg26 = "在文字後面顯示帶括號的便捷鍵 (通常用於亞洲語言)"

		Msg27 = "要檢查的終止符 (用 - 表示範圍，支援萬用字元) (空格分隔):"
		Msg28 = "要保留的終止符組合 (用 - 表示範圍，支援萬用字元) (半形逗號分隔):"
		Msg29 = "要自動替換的終止符對 (用 | 分隔替換前後的字元) (空格分隔):"

		Msg30 = "要檢查的加速器旗標符，例如 \t (半形逗號分隔):"
		Msg31 = "要檢查的加速器字元 (用 - 表示範圍，支援萬用字元) (半形逗號分隔):"
		Msg32 = "要保留的加速器字元 (用 - 表示範圍，支援萬用字元) (半形逗號分隔):"

		Msg33 = "要自動替換的字元 (用 | 分隔替換前後的字元) (半形逗號分隔):"
		Msg34 = "翻譯後要被替換的字元 (用 | 分隔替換前後的字元) (半形逗號分隔):"

		Msg35 = "說明(&H)"
		Msg37 = "字串處理"
		Msg54 = ">"
		Msg55 = "..."

		Msg71 = "適用語言"
		Msg74 = "新增  >"
		Msg75 = "全部新增 >>"
		Msg76 = "<  刪除"
		Msg77 = "<< 全部刪除"
		Msg78 = "新增可用語言"
		Msg79 = "編輯可用語言"
		Msg80 = "刪除可用語言"
		Msg81 = "新增適用語言"
		Msg82 = "編輯適用語言"
		Msg83 = "刪除適用語言"

		Msg84 = "自動更新"
		Msg85 = "更新方式"
		Msg86 = "自動下載更新並安裝(&A)"
		Msg87 = "有更新時通知我，由我決定下載並安裝(&M)"
		Msg88 = "關閉自動更新(&O)"
		Msg89 = "檢查頻率"
		Msg90 = "檢查間隔: "
		Msg91 = "天"
		Msg92 = "最後檢查日期:"
		Msg93 = "更新網址清單 (分行輸入，前者優先)"
		Msg94 = "RAR 解壓程式"
		Msg95 = "程式路徑 (支援環境變數):"
		Msg96 = "解壓參數 (%1 為壓縮檔案，%2 為要擷取的檔案，%3 為解壓路徑):"
		Msg97 = "檢查"
	Else
		Msg01 = "饜离"
		Msg02 = "③恁寁饜离甜婓狟蹈恅掛遺笢怀�鄵童炬獃婠痸騢鯥棑而譚硱絳妦椕驉�"
		Msg05 = "黍��(&R)"
		Msg06 = "氝樓(&A)"
		Msg07 = "載蜊(&M)"
		Msg08 = "刉壺(&D)"
		Msg09 = "ь諾(&C)"
		Msg10 = "聆彸(&T)"

		Msg11 = "饜离蹈桶"
		Msg12 = "悵湔濬倰"
		Msg13 = "饜离囀��"
		Msg14 = "恅璃"
		Msg15 = "蛁聊桶"
		Msg16 = "絳�蹁馺�"
		Msg17 = "絳堤饜离"
		Msg18 = "辦豎瑩"
		Msg19 = "笝砦睫"
		Msg20 = "樓厒ん"
		Msg21 = "趼睫杸遙"

		Msg22 = "猁齬壺腔漪 && 睫瘍腔準辦豎瑩趼睫郪磁 (圉褒飯瘍煦路):"
		Msg23 = "趼揹莞煦蚚梓祩睫 (蚚衾嗣跺辦豎瑩趼揹腔揭燴) (圉褒飯瘍煦路):"
		Msg24 = "猁潰脤腔辦豎瑩ヶ綴嬤瘍ㄛ瞰�� [&&F] (圉褒飯瘍煦路):"
		Msg25 = "猁悵隱腔準辦豎瑩睫ヶ綴傖勤趼睫ㄛ瞰�� (&&) (圉褒飯瘍煦路):"
		Msg26 = "婓恅掛綴醱珆尨湍嬤瘍腔辦豎瑩 (籵都蚚衾捚粔逄晟)"

		Msg27 = "猁潰脤腔笝砦睫 (蚚 - 桶尨毓峓ㄛ盓厥籵饜睫) (諾跡煦路):"
		Msg28 = "猁悵隱腔笝砦睫郪磁 (蚚 - 桶尨毓峓ㄛ盓厥籵饜睫) (圉褒飯瘍煦路):"
		Msg29 = "猁赻雄杸遙腔笝砦睫勤 (蚚 | 煦路杸遙ヶ綴腔趼睫) (諾跡煦路):"

		Msg30 = "猁潰脤腔樓厒ん梓祩睫ㄛ瞰�� \t (圉褒飯瘍煦路):"
		Msg31 = "猁潰脤腔樓厒ん趼睫 (蚚 - 桶尨毓峓ㄛ盓厥籵饜睫) (圉褒飯瘍煦路):"
		Msg32 = "猁悵隱腔樓厒ん趼睫 (蚚 - 桶尨毓峓ㄛ盓厥籵饜睫) (圉褒飯瘍煦路):"

		Msg33 = "猁赻雄杸遙腔趼睫 (蚚 | 煦路杸遙ヶ綴腔趼睫) (圉褒飯瘍煦路):"
		Msg34 = "楹祒綴猁掩杸遙腔趼睫 (蚚 | 煦路杸遙ヶ綴腔趼睫) (圉褒飯瘍煦路):"

		Msg35 = "堆翑(&H)"
		Msg37 = "趼揹揭燴"
		Msg54 = ">"
		Msg55 = "..."

		Msg71 = "巠蚚逄晟"
		Msg74 = "氝樓  >"
		Msg75 = "�垓覦篲� >>"
		Msg76 = "<  刉壺"
		Msg77 = "<< �垓褫噫�"
		Msg78 = "氝樓褫蚚逄晟"
		Msg79 = "晤憮褫蚚逄晟"
		Msg80 = "刉壺褫蚚逄晟"
		Msg81 = "氝樓巠蚚逄晟"
		Msg82 = "晤憮巠蚚逄晟"
		Msg83 = "刉壺巠蚚逄晟"

		Msg84 = "赻雄載陔"
		Msg85 = "載陔源宒"
		Msg86 = "赻雄狟婥載陔甜假蚾(&A)"
		Msg87 = "衄載陔奀籵眭扂ㄛ蚕扂樵隅狟婥甜假蚾(&M)"
		Msg88 = "壽敕赻雄載陔(&O)"
		Msg89 = "潰脤け薹"
		Msg90 = "潰脤潔路: "
		Msg91 = "毞"
		Msg92 = "郔綴潰脤�梪�:"
		Msg93 = "載陔厙硊蹈桶 (煦俴怀�諴甭啦葯欐�)"
		Msg94 = "RAR 賤揤最唗"
		Msg95 = "最唗繚噤 (盓厥遠噫曹講):"
		Msg96 = "賤揤統杅 (%1 峈揤坫恅璃ㄛ%2 峈猁枑�△鰓躁�ㄛ%3 峈賤揤繚噤):"
		Msg97 = "潰脤"
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


'③昢斛脤艘勤趕遺堆翑翋枙眕賸賤載嗣陓洘﹝
Private Function SetFunc%(DlgItem$, Action%, SuppValue&)
	Dim Header As String,HeaderID As Integer,NewData As String,Path As String,cStemp As Boolean
	Dim i As Integer,n As Integer,TempArray() As String,Temp As String,LngName As String
	Dim LngID As Integer,LangArray() As String,AppLngList() As String,UseLngList() As String

	If OSLanguage = "0404" Then
		Msg01 = "錯誤"
		Msg02 = "預設值"
		Msg03 = "原值"
		Msg04 = "參照值"
		Msg08 = "未知"
		Msg11 = "警告"
		Msg12 =	"如果某些參數為空，將使程式執行結果不正確。" & vbCrLf & _
				"您確實想要這樣做嗎？"
		Msg13 = "設定內容已經變更但是沒有儲存！是否需要儲存？"
		Msg14 = "儲存類型已經變更但是沒有儲存！是否需要儲存？"
		Msg18 = "目前設定中，至少有一項參數為空！" & vbCrLf
		Msg19 = "所有設定中，至少有一項參數為空！" & vbCrLf
		Msg21 = "確認"
		Msg22 = "確實要刪除設定「%s」嗎？"
		Msg24 = "確實要刪除語言「%s」嗎？"
		Msg30 = "訊息"
		Msg32 = "匯出設定成功！"
		Msg33 = "匯入設定成功！"
		Msg36 = "無法儲存！請檢查是否有寫入下列位置的權限:" & vbCrLf & vbCrLf
		Msg39 = "匯入失敗！請檢查是否有寫入下列位置的權限" & vbCrLf & _
				"或匯入檔案的格式是否正確:" & vbCrLf & vbCrLf
		Msg40 = "匯出失敗！請檢查是否有寫入下列位置的權限" & vbCrLf & _
				"或匯出檔案的格式是否正確:" & vbCrLf & vbCrLf
		Msg41 = "儲存失敗！請檢查是否有寫入下列位置的權限:" & vbCrLf & vbCrLf
		Msg42 = "選取要匯入的檔案"
		Msg43 = "選取要匯出的檔案"
		Msg44 = "設定檔案 (*.dat)|*.dat|所有檔案 (*.*)|*.*||"
		Msg54 = "可用語言:"
		Msg55 = "適用語言:"
		Msg60 = "選取解壓程式"
		Msg61 = "可執行檔案 (*.exe)|*.exe|所有檔案 (*.*)|*.*||"
		Msg62 = "沒有指定解壓程式！請重新輸入或選取。"
		Msg63 = "檔案參照參數(%1)"
		Msg64 = "要擷取的檔案參數(%2)"
		Msg65 = "解壓路徑參數(%3)"
	Else
		Msg01 = "渣昫"
		Msg02 = "蘇�珋�"
		Msg03 = "埻硉"
		Msg04 = "統桽硉"
		Msg08 = "帤眭"
		Msg11 = "劑豢"
		Msg12 =	"�蝜�議虳統杅峈諾ㄛ蔚妏最唗堍俴賦彆祥淏�楚�" & vbCrLf & _
				"蠟�滔舜遻肪瑵�酕鎘ˋ"
		Msg13 = "饜离囀�椹挩飛�蜊筍岆羶衄悵湔ㄐ岆瘁剒猁悵湔ˋ"
		Msg14 = "悵湔濬倰眒冪載蜊筍岆羶衄悵湔ㄐ岆瘁剒猁悵湔ˋ"
		Msg18 = "絞ヶ饜离笢ㄛ祫屾衄珨砐統杅峈諾ㄐ" & vbCrLf
		Msg19 = "垀衄饜离笢ㄛ祫屾衄珨砐統杅峈諾ㄐ" & vbCrLf
		Msg21 = "�溜�"
		Msg22 = "�滔菸罔噫�饜离※%s§鎘ˋ"
		Msg24 = "�滔菸罔噫�逄晟※%s§鎘ˋ"
		Msg30 = "陓洘"
		Msg32 = "絳堤饜离傖髡ㄐ"
		Msg33 = "絳�蹁馺籀伄忙�"
		Msg36 = "拸楊悵湔ㄐ③潰脤岆瘁衄迡�輴臏倛閥繭饑使�:" & vbCrLf & vbCrLf
		Msg39 = "絳�輮妍隀﹊趧麮樨Й鵛俴椅輴臏倛閥繭饑使�" & vbCrLf & _
				"麼絳�輷躁�腔跡宒岆瘁淏��:" & vbCrLf & vbCrLf
		Msg40 = "絳堤囮啖ㄐ③潰脤岆瘁衄迡�輴臏倛閥繭饑使�" & vbCrLf & _
				"麼絳堤恅璃腔跡宒岆瘁淏��:" & vbCrLf & vbCrLf
		Msg41 = "悵湔囮啖ㄐ③潰脤岆瘁衄迡�輴臏倛閥繭饑使�:" & vbCrLf & vbCrLf
		Msg42 = "恁寁猁絳�賮鰓躁�"
		Msg43 = "恁寁猁絳堤腔恅璃"
		Msg44 = "饜离恅璃 (*.dat)|*.dat|垀衄恅璃 (*.*)|*.*||"
		Msg54 = "褫蚚逄晟:"
		Msg55 = "巠蚚逄晟:"
		Msg60 = "恁寁賤揤最唗"
		Msg61 = "褫硒俴恅璃 (*.exe)|*.exe|垀衄恅璃 (*.*)|*.*||"
		Msg62 = "羶衄硌隅賤揤最唗ㄐ③笭陔怀�趥藘√鞢�"
		Msg63 = "恅璃竘蚚統杅(%1)"
		Msg64 = "猁枑�△鰓躁�統杅(%2)"
		Msg65 = "賤揤繚噤統杅(%3)"
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
	Case 1 ' 勤趕遺敦諳場宎趙
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
	Case 2 ' 杅硉載蜊麼氪偌狟賸偌聽
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
						SetFunc% = True '滅砦偌狟偌聽壽敕勤趕遺敦諳
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
						SetFunc% = True '滅砦偌狟偌聽壽敕勤趕遺敦諳
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
						SetFunc% = True '滅砦偌狟偌聽壽敕勤趕遺敦諳
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
					SetFunc% = True '滅砦偌狟偌聽壽敕勤趕遺敦諳
					Exit Function
				End If
			End If
			If DlgValue("cWriteType") = 0 Then cPath = CheckFilePath
			If DlgValue("cWriteType") = 1 Then cPath = CheckRegKey
			If CheckWrite(CheckDataList,cPath,"Sets") = False Then
				MsgBox(Msg36 & cPath,vbOkOnly+vbInformation,Msg01)
				SetFunc% = True '滅砦偌狟偌聽壽敕勤趕遺敦諳
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
			SetFunc% = True '滅砦偌狟偌聽壽敕勤趕遺敦諳
		End If
	Case 3 ' 恅掛遺麼氪郪磁遺恅掛掩載蜊
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


'氝樓扢离靡備
Function AddSet(DataArr() As String) As String
	Dim NewHeader As String
	If OSLanguage = "0404" Then
		Msg01 = "新增"
		Msg04 = "請輸入新設定的名稱:"
		Msg06 = "錯誤"
		Msg07 = "您沒有輸入任何內容！請重新輸入。"
		Msg08 = "該名稱已經存在！請輸入一個不同的名稱。"
	Else
		Msg01 = "陔膘"
		Msg04 = "③怀�遶藍馺繭鏽�備:"
		Msg06 = "渣昫"
		Msg07 = "蠟羶衄怀�躽庥恅硜搟﹊鄵寪薹靿諢�"
		Msg08 = "蜆靡備眒冪湔婓ㄐ③怀�遻遘鶷銓炸鏽�備﹝"
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


'晤憮扢离靡備
Function EditSet(DataArr() As String,Header As String) As String
	Dim tempHeader As String,NewHeader As String
	If OSLanguage = "0404" Then
		Msg01 = "變更"
		Msg04 = "舊名稱:"
		Msg05 = "新名稱:"
		Msg06 = "錯誤"
		Msg07 = "您沒有輸入任何內容！請重新輸入。"
		Msg08 = "該名稱已經存在！請輸入一個不同的名稱。"
	Else
		Msg01 = "載蜊"
		Msg04 = "導靡備:"
		Msg05 = "陔靡備:"
		Msg06 = "渣昫"
		Msg07 = "蠟羶衄怀�躽庥恅硜搟﹊鄵寪薹靿諢�"
		Msg08 = "蜆靡備眒冪湔婓ㄐ③怀�遻遘鶷銓炸鏽�備﹝"
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


'氝樓麼晤憮逄晟勤
Function SetLang(DataArr() As String,LangName As String,LangCode As String) As String
	Dim tempHeader As String,NewLangName As String,NewLangCode As String
	If OSLanguage = "0404" Then
		Msg01 = "新增"
		Msg02 = "編輯"
		Msg04 = "語言名稱:"
		Msg05 = "Passolo 語言代碼:"
		Msg10 = "錯誤"
		Msg11 = "您沒有輸入任何內容！請重新輸入。"
		Msg12 = "語言名稱和 Passolo 語言代碼中至少有一個項目為空！請檢查並輸入。"
		Msg13 = "該語言名稱已經存在！請重新輸入。"
		Msg14 = "該 Passolo 語言代碼已經存在！請重新輸入。"
	Else
		Msg01 = "氝樓"
		Msg02 = "晤憮"
		Msg04 = "逄晟靡備:"
		Msg05 = "Passolo 逄晟測鎢:"
		Msg10 = "渣昫"
		Msg11 = "蠟羶衄怀�躽庥恅硜搟﹊鄵寪薹靿諢�"
		Msg12 = "逄晟靡備睿 Passolo 逄晟測鎢笢祫屾衄珨跺砐醴峈諾ㄐ③潰脤甜怀�諢�"
		Msg13 = "蜆逄晟靡備眒冪湔婓ㄐ③笭陔怀�諢�"
		Msg14 = "蜆 Passolo 逄晟測鎢眒冪湔婓ㄐ③笭陔怀�諢�"
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


'載蜊饜离蚥珂撰
Function SetLevel(HeaderList() As String,DataList() As String) As Boolean
	SetLevel = False
	If OSLanguage = "0404" Then
		Msg01 = "設定優先級"
		Msg02 = "設定優先級用於基於設定的適用語言的自動選取設定功能。"
		Msg03 = "提示:" & vbCrLf & _
				"- 有多個設定包含了相同的適用語言時，需要設定其優先級。" & vbCrLf & _
				"- 在相同適用語言的設定中，前面的設定被優先選取使用。"
		Msg04 = "上移(&U)"
		Msg05 = "下移(&D)"
		Msg06 = "重設(&R)"
	Else
		Msg01 = "饜离蚥珂撰"
		Msg02 = "饜离蚥珂撰蚚衾價衾饜离腔巠蚚逄晟腔赻雄恁寁饜离髡夔﹝"
		Msg03 = "枑尨:" & vbCrLf & _
				"- 衄嗣跺饜离婦漪賸眈肮腔巠蚚逄晟奀ㄛ剒猁扢离む蚥珂撰﹝" & vbCrLf & _
				"- 婓眈肮巠蚚逄晟腔饜离笢ㄛヶ醱腔饜离掩蚥珂恁寁妏蚚﹝"
		Msg04 = "奻痄(&U)"
		Msg05 = "狟痄(&D)"
		Msg06 = "笭离(&R)"
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


'載蜊饜离蚥珂撰勤趕遺滲杅
Private Function SetLevelFunc%(DlgItem$, Action%, SuppValue&)
	Dim i As Integer,ID As Integer,Temp As String
	Select Case Action%
	Case 1 ' 勤趕遺敦諳場宎趙
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
	Case 2 ' 杅硉載蜊麼氪偌狟賸偌聽
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
			SetLevelFunc = True '滅砦偌狟偌聽壽敕勤趕遺敦諳
		End If
	End Select
End Function


'鳳�＝硒挺麮橑髲�
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
			'鳳�� Option 砐睿硉
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
			'鳳�� Update 砐睿硉
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
			'鳳�� Option 砐俋腔�垓諫蹎邳�
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
				'載陔導唳腔蘇�狣馺譆�
				If InStr(Join(DefaultCheckList,JoinStr),Header) Then
					If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 Then
						Data = CheckDataUpdate(Header,Data)
					End If
				End If
				'悵湔杅擂善杅郪笢
				CreateArray(Header,Data,HeaderList,DataList)
				CheckGet = True
			End If
			'杅擂場宎趙
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
	'悵湔載陔綴腔杅擂善恅璃
	If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 And Path = CheckFilePath Then
		If Dir(CheckFilePath) <> "" Then CheckWrite(DataList,CheckFilePath,"All")
	End If
	If cWriteLoc = "" Then cWriteLoc = CheckFilePath
	Exit Function

	GetFromRegistry:
	'鳳�� Option 砐睿硉
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
		'鳳�� Update 砐睿硉
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

	'鳳�� Option 俋腔砐睿硉
	HeaderIDs = GetSetting("AccessKey","Option","Headers","")
	If HeaderIDs <> "" Then
		HeaderIDArr = Split(HeaderIDs,";",-1)
		For i = LBound(HeaderIDArr) To UBound(HeaderIDArr)
			HeaderID = HeaderIDArr(i)
			If HeaderID <> "" Then
				'蛌湔導唳腔藩跺砐睿硉
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
						'載陔導唳腔蘇�狣馺譆�
						If InStr(Join(DefaultCheckList,JoinStr),Header) Then
							If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 Then
								Data = CheckDataUpdate(Header,Data)
							End If
						End If
						'悵湔杅擂善杅郪笢
						CreateArray(Header,Data,HeaderList,DataList)
						CheckGet = True
					End If
					'刉壺導唳饜离硉
					On Error Resume Next
					If Header = HeaderID Then DeleteSetting("AccessKey",Header)
					On Error GoTo 0
				End If
			End If
		Next i
	End If
	'悵湔載陔綴腔杅擂善蛁聊桶
	If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 Then
		If HeaderIDs <> "" Then CheckWrite(DataList,CheckRegKey,"Sets")
	End If
	If cWriteLoc = "" Then cWriteLoc = CheckRegKey
End Function


'迡�鄶硒挺麮橑髲�
Function CheckWrite(DataList() As String,Path As String,WriteType As String) As Boolean
	Dim i As Integer,Header As String,HeaderIDs As String,TempPath As String
	Dim SetsArray() As String,HeaderIDArr() As String
	CheckWrite = False
	KeepSet = cSelected(UBound(cSelected))

	'迡�輷躁�
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

	'迡�鄶３嵿�
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
			'刉壺埻饜离砐
			HeaderIDs = GetSetting("AccessKey","Option","Headers")
			If HeaderIDs <> "" Then
				HeaderIDArr = Split(HeaderIDs,";",-1)
				On Error Resume Next
				For i = LBound(HeaderIDArr) To UBound(HeaderIDArr)
					DeleteSetting("AccessKey",HeaderIDArr(i))
				Next i
				On Error GoTo 0
			End If
			'迡�遶藍馺譁�
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
	'刉壺垀衄悵湔腔扢离
	ElseIf Path = "" Then
		'刉壺恅璃饜离砐
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
		'刉壺蛁聊桶饜离砐
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
		'扢离迡�輷閥蟹髲襞矽�
		CheckWrite = True
		cWriteLoc = ""
	End If
	ExitFunction:
End Function


'杸遙絳�蹁馺襞躁�笢腔逄晟靡備峈絞ヶ炵苀腔逄晟靡備
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


'杸遙絳�趧麮橝馺襞躁�笢腔趼睫峈絞ヶ炵苀腔逄晟趼睫
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


'載陔潰脤導唳掛饜离硉
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


'崝樓麼載蜊杅郪砐醴
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


'刉壺杅郪砐醴
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


'莞煦杅郪
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


'斐膘媼跺誑硃腔靡備杅郪
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


'磁甜梓袧睿赻隅砱逄晟蹈桶
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


'汜傖�戊�杅擂蹈桶笢諾砐綴腔逄晟勤
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


'誑遙杅郪砐醴
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


'脤梑硌隅硉岆瘁婓杅郪笢
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


'潰脤杅郪笢岆瘁衄諾硉
'ftype = 0     潰脤杅郪砐囀岆瘁�屋矽欶�
'ftype = 1     潰脤杅郪砐囀岆瘁衄諾硉
'Header = ""   潰脤淕跺杅郪
'Header <> ""  潰脤硌隅杅郪砐
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


'杅郪齬唗
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


'ь燴杅郪笢笭葩腔杅擂
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


'籵饜睫脤梑硌隅硉
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
			'PSL.Output Key & " : " &  FindStr  '覃彸蚚
			KeyCode = UCase(Key) Like UCase(FindStr)
			If KeyCode = True Then CheckKeyCode = 1
			If KeyCode = True Then Exit For
		Next i
	ElseIf CheckKey = "" And Key <> "" Then
		CheckKeyCode = 1
	End If
End Function


'聆彸潰脤最唗
Sub CheckTest(CheckID As Integer,HeaderList() As String)
	Dim TrnList As PslTransList,i As Integer,TrnListDec As String,TrnListArray() As String
	If OSLanguage = "0404" Then
		Msg01 = "參數測試"
		Msg02 = "根據下列條件組合搜尋翻譯錯誤並給出修正。按 [測試] 按鈕輸出結果。"
		Msg03 = "設定名稱:"
		Msg05 = "翻譯清單:"
		Msg06 = "讀入行數:"
		Msg07 = "包含內容:"
		Msg08 = "支援萬用字元和半形分號分隔的多項"
		Msg09 = "字串內容:"
		Msg10 = "全部(&F)"
		Msg11 = "便捷鍵(&K)"
		Msg12 = "終止符(&E)"
		Msg13 = "加速器(&P)"
		Msg15 = "測試(&T)"
		Msg16 = "清空(&C)"
		Msg18 = "說明(&H)"
		Msg19 = "自動替換字元(&R)"
	Else
		Msg01 = "統杅聆彸"
		Msg02 = "跦擂狟蹈沭璃郪磁脤梑楹祒渣昫甜跤堤党淏﹝偌 [聆彸] 偌聽怀堤賦彆﹝"
		Msg03 = "饜离靡備:"
		Msg05 = "楹祒蹈桶:"
		Msg06 = "黍�遶倇�:"
		Msg07 = "婦漪囀��:"
		Msg08 = "盓厥籵饜睫睿圉褒煦瘍煦路腔嗣砐"
		Msg09 = "趼揹囀��:"
		Msg10 = "�垓�(&F)"
		Msg11 = "辦豎瑩(&K)"
		Msg12 = "笝砦睫(&E)"
		Msg13 = "樓厒ん(&P)"
		Msg15 = "聆彸(&T)"
		Msg16 = "ь諾(&C)"
		Msg18 = "堆翑(&H)"
		Msg19 = "赻雄杸遙趼睫(&R)"
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


'聆彸勤趕遺滲杅
Private Function CheckTestFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim ListDec As String,TempDec As String,LineNum As Integer,inText As String,repStr As Integer
	Dim cAllCont As Integer,cAccKey As Integer,cEndChar As Integer,cAcceler As Integer
	Dim SpecifyText As String,CheckID As Integer
	If OSLanguage = "0404" Then
		Msg01 = "正在搜尋錯誤並給出修正，可能需要幾分鐘，請稍候..."
	Else
		Msg01 = "淏婓脤梑渣昫甜跤堤党淏ㄛ褫夔剒猁撓煦笘ㄛ③尕緊..."
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
	Case 2 ' 杅硉載蜊麼氪偌狟賸偌聽
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
			CheckTestFunc = True '滅砦偌狟偌聽壽敕勤趕遺敦諳
		End If
	End Select
End Function


'揭燴睫磁沭璃腔趼揹蹈桶笢腔趼揹
Function CheckStrings(ID As Integer,ListDec As String,LineNum As Integer,rStr As Integer,sText As String) As String
	Dim i As Integer,j As Integer,k As Integer,srcString As String,trnString As String,tText As String
	Dim TrnList As PslTransList,TrnListDec As String,TranLang As String
	Dim CheckVer As String,CheckSet As String,CheckState As String,CheckDate As Date,TranDate As Date
	Dim LineNumErrCount As Integer,accKeyNumErrCount As Integer,ErrorCount As Integer,Massage As String
	Dim Find As Boolean,srcFindNum As Integer,trnFindNum As Integer

	If OSLanguage = "0404" Then
		Msg01 = "原文: "
		Msg02 = "譯文: "
		Msg03 = "修正: "
		Msg04 = "訊息: "
		Msg05 = "---------------------------------------"
		Msg06 = "找到下列錯誤:"
		Msg07 = "沒有找到錯誤。"
		Msg08 = "包含指定內容的字串中沒有找到錯誤。"
		Msg09 = "沒有找到包含指定內容的字串！"
	Else
		Msg01 = "埻恅: "
		Msg02 = "祒恅: "
		Msg03 = "党淏: "
		Msg04 = "秏洘: "
		Msg05 = "---------------------------------------"
		Msg06 = "梑善狟蹈渣昫:"
		Msg07 = "羶衄梑善渣昫﹝"
		Msg08 = "婦漪硌隅囀�搧儷硒核陏閨俷珛蓬簊鞳�"
		Msg09 = "羶衄梑善婦漪硌隅囀�搧儷硒恐�"
	End If

	'統杅場宎趙
	CheckStrings = ""
	tText = ""
	k = 0
	LineNumErrCount = 0
	accKeyNumErrCount = 0
	Find = False

	'鳳�×▲巡譟倡蹅訇�
	For i = 1 To trn.Project.TransLists.Count
		Set TrnList = trn.Project.TransLists(i)
		TrnListDec = TrnList.Title & " - " & PSL.GetLangCode(TrnList.Language.LangID,pslCodeText)
		If TrnListDec = ListDec Then Exit For
	Next i

	'鳳�＿膨縎擿�
	trnLng = PSL.GetLangCode(TrnList.Language.LangID,pslCode639_1)
	If trnLng = "zh" Or trnLng = "ja" Or trnLng = "ko" Then TranLang = "Asia"

	If TrnList.SourceList.LastChange > TrnList.LastUpdate Then TrnList.Update
	If LineNum > TrnList.StringCount Then LineNum = TrnList.StringCount
	For i = 1 To TrnList.StringCount
		'統杅場宎趙
		srcString = ""
		trnString = ""
		NewtrnString = ""
		LineMsg = ""
		AccKeyMsg = ""
		ReplaceMsg = ""

		'鳳�√倀贍芛倡鄶硒�
		Set TransString = TrnList.String(i)
		If TransString.Text <> "" Then
			srcString = TransString.SourceText
			trnString = TransString.Text
			OldtrnString = trnString

			'羲宎揭燴趼揹
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

			'覃蚚秏洘怀堤
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


'壺�皿硒旁偕鯡葆巡� PreStr 睿 AppStr
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


'黍�±擿堈�
Function LangCodeList(DataName As String,OSLang As String,MinNum As Integer,MaxNum As Integer) As Variant
	Dim i As Integer,j As Integer,Code As String,LangName() As String,LangPair() As String
	ReDim LangName(MaxNum - MinNum),LangPair(MaxNum - MinNum)
	For i = MinNum To MaxNum
		j = i - MinNum
		If OSLang = "0404" Then
			If i = 0 Then LangName(j) = "自動偵測"
			If i = 1 Then LangName(j) = "南非荷蘭語"
			If i = 2 Then LangName(j) = "阿爾巴尼亞語"
			If i = 3 Then LangName(j) = "阿姆哈拉語"
			If i = 4 Then LangName(j) = "阿拉伯語"
			If i = 5 Then LangName(j) = "亞美尼亞語"
			If i = 6 Then LangName(j) = "阿薩姆語"
			If i = 7 Then LangName(j) = "阿塞拜疆語"
			If i = 8 Then LangName(j) = "巴什基爾語"
			If i = 9 Then LangName(j) = "巴斯克語"
			If i = 10 Then LangName(j) = "白俄羅斯語"
			If i = 11 Then LangName(j) = "孟加拉語"
			If i = 12 Then LangName(j) = "波西尼亞語"
			If i = 13 Then LangName(j) = "布列塔尼語"
			If i = 14 Then LangName(j) = "保加利亞語"
			If i = 15 Then LangName(j) = "加泰羅尼亞語"
			If i = 16 Then LangName(j) = "簡體中文"
			If i = 17 Then LangName(j) = "正體中文"
			If i = 18 Then LangName(j) = "科西嘉語"
			If i = 19 Then LangName(j) = "克羅地亞語"
			If i = 20 Then LangName(j) = "捷克語"
			If i = 21 Then LangName(j) = "丹麥語"
			If i = 22 Then LangName(j) = "荷蘭語"
			If i = 23 Then LangName(j) = "英語"
			If i = 24 Then LangName(j) = "愛沙尼亞語"
			If i = 25 Then LangName(j) = "法羅語"
			If i = 26 Then LangName(j) = "波斯語"
			If i = 27 Then LangName(j) = "芬蘭語"
			If i = 28 Then LangName(j) = "法語"
			If i = 29 Then LangName(j) = "弗裡西亞語"
			If i = 30 Then LangName(j) = "加利西亞語"
			If i = 31 Then LangName(j) = "格魯吉亞語"
			If i = 32 Then LangName(j) = "德語"
			If i = 33 Then LangName(j) = "希臘語"
			If i = 34 Then LangName(j) = "格陵蘭語"
			If i = 35 Then LangName(j) = "古吉拉特語"
			If i = 36 Then LangName(j) = "豪薩語"
			If i = 37 Then LangName(j) = "希伯來語"
			If i = 38 Then LangName(j) = "印地語"
			If i = 39 Then LangName(j) = "匈牙利語"
			If i = 40 Then LangName(j) = "冰島語"
			If i = 41 Then LangName(j) = "印度尼西亞語"
			If i = 42 Then LangName(j) = "因紐特語"
			If i = 43 Then LangName(j) = "愛爾蘭語"
			If i = 44 Then LangName(j) = "班圖語"
			If i = 45 Then LangName(j) = "祖魯語"
			If i = 46 Then LangName(j) = "意大利語"
			If i = 47 Then LangName(j) = "日語"
			If i = 48 Then LangName(j) = "卡納達語"
			If i = 49 Then LangName(j) = "克什米爾語"
			If i = 50 Then LangName(j) = "哈薩克語"
			If i = 51 Then LangName(j) = "高棉語"
			If i = 52 Then LangName(j) = "盧旺達語"
			If i = 53 Then LangName(j) = "孔卡尼語"
			If i = 54 Then LangName(j) = "朝鮮語"
			If i = 55 Then LangName(j) = "吉爾吉斯語"
			If i = 56 Then LangName(j) = "吉爾吉斯語 (吉爾吉斯坦)"
			If i = 57 Then LangName(j) = "老撾語"
			If i = 58 Then LangName(j) = "拉脫維亞語"
			If i = 59 Then LangName(j) = "立陶宛語"
			If i = 60 Then LangName(j) = "盧森堡語"
			If i = 61 Then LangName(j) = "馬其頓語"
			If i = 62 Then LangName(j) = "馬來語"
			If i = 63 Then LangName(j) = "馬拉雅拉姆語"
			If i = 64 Then LangName(j) = "馬耳他語"
			If i = 65 Then LangName(j) = "毛利語"
			If i = 66 Then LangName(j) = "馬拉地語"
			If i = 67 Then LangName(j) = "蒙古語"
			If i = 68 Then LangName(j) = "尼泊爾語"
			If i = 69 Then LangName(j) = "挪威語"
			If i = 70 Then LangName(j) = "挪威語 (博克馬爾文)"
			If i = 71 Then LangName(j) = "挪威語 (尼諾斯克文)"
			If i = 72 Then LangName(j) = "奧裡雅語"
			If i = 73 Then LangName(j) = "普什圖語"
			If i = 74 Then LangName(j) = "波蘭語"
			If i = 75 Then LangName(j) = "葡萄牙語"
			If i = 76 Then LangName(j) = "旁遮普語"
			If i = 77 Then LangName(j) = "克丘亞語"
			If i = 78 Then LangName(j) = "羅馬尼亞語"
			If i = 79 Then LangName(j) = "俄語"
			If i = 80 Then LangName(j) = "薩米語"
			If i = 81 Then LangName(j) = "梵語"
			If i = 82 Then LangName(j) = "塞爾維亞語"
			If i = 83 Then LangName(j) = "巴索托語"
			If i = 84 Then LangName(j) = "茨瓦納語"
			If i = 85 Then LangName(j) = "信德語"
			If i = 86 Then LangName(j) = "僧伽羅語"
			If i = 87 Then LangName(j) = "斯洛伐克語"
			If i = 88 Then LangName(j) = "斯洛文尼亞語"
			If i = 89 Then LangName(j) = "西班牙語"
			If i = 90 Then LangName(j) = "斯瓦希裡語"
			If i = 91 Then LangName(j) = "瑞典語"
			If i = 92 Then LangName(j) = "敘利亞語"
			If i = 93 Then LangName(j) = "塔吉克語"
			If i = 94 Then LangName(j) = "泰米爾語"
			If i = 95 Then LangName(j) = "韃靼語"
			If i = 96 Then LangName(j) = "泰盧固語"
			If i = 97 Then LangName(j) = "泰語"
			If i = 98 Then LangName(j) = "藏語"
			If i = 99 Then LangName(j) = "土耳其語"
			If i = 100 Then LangName(j) = "土庫曼語"
			If i = 101 Then LangName(j) = "維吾爾語"
			If i = 102 Then LangName(j) = "烏克蘭語"
			If i = 103 Then LangName(j) = "烏爾都語"
			If i = 104 Then LangName(j) = "烏茲別克語"
			If i = 105 Then LangName(j) = "越南語"
			If i = 106 Then LangName(j) = "威爾士語"
			If i = 107 Then LangName(j) = "沃洛夫語"
		Else
			If i = 0 Then LangName(j) = "赻雄潰聆"
			If i = 1 Then LangName(j) = "鰍準盡擘逄"
			If i = 2 Then LangName(j) = "陝嫌匙攝捚逄"
			If i = 3 Then LangName(j) = "陝譟慇嶺逄"
			If i = 4 Then LangName(j) = "陝嶺皎逄"
			If i = 5 Then LangName(j) = "捚藝攝捚逄"
			If i = 6 Then LangName(j) = "陝�躟照�"
			If i = 7 Then LangName(j) = "陝��問蔭逄"
			If i = 8 Then LangName(j) = "匙妦價嫌逄"
			If i = 9 Then LangName(j) = "匙佴親逄"
			If i = 10 Then LangName(j) = "啞塘蹕佴逄"
			If i = 11 Then LangName(j) = "譁樓嶺逄"
			If i = 12 Then LangName(j) = "疏昹攝捚逄"
			If i = 13 Then LangName(j) = "票蹈坢攝逄"
			If i = 14 Then LangName(j) = "悵樓瞳捚逄"
			If i = 15 Then LangName(j) = "樓怍蹕攝捚逄"
			If i = 16 Then LangName(j) = "潠极笢恅"
			If i = 17 Then LangName(j) = "楛极笢恅"
			If i = 18 Then LangName(j) = "褪昹樁逄"
			If i = 19 Then LangName(j) = "親蹕華捚逄"
			If i = 20 Then LangName(j) = "豎親逄"
			If i = 21 Then LangName(j) = "竣闔逄"
			If i = 22 Then LangName(j) = "盡擘逄"
			If i = 23 Then LangName(j) = "荎逄"
			If i = 24 Then LangName(j) = "乾伈攝捚逄"
			If i = 25 Then LangName(j) = "楊蹕逄"
			If i = 26 Then LangName(j) = "疏佴逄"
			If i = 27 Then LangName(j) = "煉擘逄"
			If i = 28 Then LangName(j) = "楊逄"
			If i = 29 Then LangName(j) = "艇爵昹捚逄"
			If i = 30 Then LangName(j) = "樓瞳昹捚逄"
			If i = 31 Then LangName(j) = "跡糧憚捚逄"
			If i = 32 Then LangName(j) = "肅逄"
			If i = 33 Then LangName(j) = "洷幫逄"
			If i = 34 Then LangName(j) = "跡鍬擘逄"
			If i = 35 Then LangName(j) = "嘉憚嶺杻逄"
			If i = 36 Then LangName(j) = "瑰�驞�"
			If i = 37 Then LangName(j) = "洷皎懂逄"
			If i = 38 Then LangName(j) = "荂華逄"
			If i = 39 Then LangName(j) = "倧挴瞳逄"
			If i = 40 Then LangName(j) = "梨絢逄"
			If i = 41 Then LangName(j) = "荂僅攝昹捚逄"
			If i = 42 Then LangName(j) = "秪臟杻逄"
			If i = 43 Then LangName(j) = "乾嫌擘逄"
			If i = 44 Then LangName(j) = "啤芞逄"
			If i = 45 Then LangName(j) = "逌糧逄"
			If i = 46 Then LangName(j) = "砩湮瞳逄"
			If i = 47 Then LangName(j) = "�梌�"
			If i = 48 Then LangName(j) = "縐馨湛逄"
			If i = 49 Then LangName(j) = "親妦譙嫌逄"
			If i = 50 Then LangName(j) = "慇�蠵剆�"
			If i = 51 Then LangName(j) = "詢蹬逄"
			If i = 52 Then LangName(j) = "竅咺湛逄"
			If i = 53 Then LangName(j) = "謂縐攝逄"
			If i = 54 Then LangName(j) = "陳珅逄"
			If i = 55 Then LangName(j) = "憚嫌憚佴逄"
			If i = 56 Then LangName(j) = "憚嫌憚佴逄 (憚嫌憚佴拊)"
			If i = 57 Then LangName(j) = "橾恄逄"
			If i = 58 Then LangName(j) = "嶺迕峎捚逄"
			If i = 59 Then LangName(j) = "蕾枎剄逄"
			If i = 60 Then LangName(j) = "竅伬惜逄"
			If i = 61 Then LangName(j) = "鎮む嗨逄"
			If i = 62 Then LangName(j) = "鎮懂逄"
			If i = 63 Then LangName(j) = "鎮嶺捇嶺譟逄"
			If i = 64 Then LangName(j) = "鎮嫉坻逄"
			If i = 65 Then LangName(j) = "禱瞳逄"
			If i = 66 Then LangName(j) = "鎮嶺華逄"
			If i = 67 Then LangName(j) = "蟹嘉逄"
			If i = 68 Then LangName(j) = "攝眼嫌逄"
			If i = 69 Then LangName(j) = "鑑哏逄"
			If i = 70 Then LangName(j) = "鑑哏逄 (痔親鎮嫌恅)"
			If i = 71 Then LangName(j) = "鑑哏逄 (攝霾佴親恅)"
			If i = 72 Then LangName(j) = "兜爵捇逄"
			If i = 73 Then LangName(j) = "ぱ妦芞逄"
			If i = 74 Then LangName(j) = "疏擘逄"
			If i = 75 Then LangName(j) = "に曶挴逄"
			If i = 76 Then LangName(j) = "籥殑ぱ逄"
			If i = 77 Then LangName(j) = "親⑧捚逄"
			If i = 78 Then LangName(j) = "蹕鎮攝捚逄"
			If i = 79 Then LangName(j) = "塘逄"
			If i = 80 Then LangName(j) = "�躞赽�"
			If i = 81 Then LangName(j) = "鼐逄"
			If i = 82 Then LangName(j) = "��嫌峎捚逄"
			If i = 83 Then LangName(j) = "匙坰迖逄"
			If i = 84 Then LangName(j) = "棕俓馨逄"
			If i = 85 Then LangName(j) = "陓肅逄"
			If i = 86 Then LangName(j) = "仵暀蹕逄"
			If i = 87 Then LangName(j) = "佴醫極親逄"
			If i = 88 Then LangName(j) = "佴醫恅攝捚逄"
			If i = 89 Then LangName(j) = "昹啤挴逄"
			If i = 90 Then LangName(j) = "佴俓洷爵逄"
			If i = 91 Then LangName(j) = "�藒駃�"
			If i = 92 Then LangName(j) = "唦瞳捚逄"
			If i = 93 Then LangName(j) = "坢憚親逄"
			If i = 94 Then LangName(j) = "怍譙嫌逄"
			If i = 95 Then LangName(j) = "鰷鱁逄"
			If i = 96 Then LangName(j) = "怍竅嘐逄"
			If i = 97 Then LangName(j) = "怍逄"
			If i = 98 Then LangName(j) = "紲逄"
			If i = 99 Then LangName(j) = "芩嫉む逄"
			If i = 100 Then LangName(j) = "芩踱霤逄"
			If i = 101 Then LangName(j) = "峎挓嫌逄"
			If i = 102 Then LangName(j) = "拫親擘逄"
			If i = 103 Then LangName(j) = "拫嫌飲逄"
			If i = 104 Then LangName(j) = "拫觕梗親逄"
			If i = 105 Then LangName(j) = "埣鰍逄"
			If i = 106 Then LangName(j) = "哏嫌尪逄"
			If i = 107 Then LangName(j) = "挋醫痲逄"
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


'潰脤堆翑
Sub CheckHelp(HelpTip As String)
	If OSLanguage = "0404" Then
	AboutTitle = "關於"
	HelpTitle = "說明"
	HelpTipTitle = "便捷鍵、終止符和加速器檢查巨集"
	AboutWindows = " 關於 "
	MainWindows = " 主視窗 "
	SetWindows = " 設定視窗 "
	TestWindows = " 測試視窗 "
	Lines = "-----------------------"
	Sys = "軟體版本：" & Version & vbCrLf & _
			"適用系統：Windows XP/2000 以上系統" & vbCrLf & _
			"適用版本：所有支援巨集處理的 Passolo 6.0 及以上版本" & vbCrLf & _
			"介面語言：簡體中文和正體中文 (自動辨識)" & vbCrLf & _
			"版權所有：漢化新世紀" & vbCrLf & _
			"授權形式：免費軟體" & vbCrLf & _
			"官方首頁：http://www.hanzify.org" & vbCrLf & _
			"前開發者：漢化新世紀成員 gnatix (2007-2008)" & vbCrLf & _
			"後開發者：漢化新世紀成員 wanfu (2009-2010)" & vbCrLf & vbCrLf & vbCrLf
	Ement = "☆執行環境☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 支援巨集處理的 Passolo 6.0 及以上版本，必需" & vbCrLf & _
			"- Windows Script Host (WSH) 物件 (VBS)，必需" & vbCrLf & _
			"- Adodb.Stream 物件 (VBS)，支援自動更新所需" & vbCrLf & _
			"- Microsoft.XMLHTTP 物件，支援自動更新所需" & vbCrLf & vbCrLf & vbCrLf
	Dec = "☆軟體簡介☆" & vbCrLf & _
			"============" & vbCrLf & _
			"便捷鍵、終止符和加速器檢查巨集是一個用於 Passolo 翻譯檢查的巨集程式。它具有以下功能：" & vbCrLf & _
			"- 檢查翻譯中便捷鍵、終止符、加速器和空格" & vbCrLf & _
			"- 檢查並修正檢查翻譯中便捷鍵、終止符、加速器和空格" & vbCrLf & _
			"- 刪除翻譯中的便捷鍵" & vbCrLf & _
			"- 內置可自訂的自動更新功能" & vbCrLf & vbCrLf & _
			"本程式包含下列檔案：" & vbCrLf & _
			"- 自動巨集：PslAutoAccessKey.bas" & vbCrLf & _
			"  在翻譯字串時，自動更正錯誤的翻譯。利用該巨集，您可以不必輸入便捷鍵、終止符、加速器，" & vbCrLf & _
			"  系統將根據您選取的設定自動幫您新增和原文一樣的便捷鍵、終止符、加速器，並翻譯終止符。" & vbCrLf & _
			"  利用它可以提高翻譯速度，並減少翻譯錯誤。" & vbCrLf & _
			"  ◎注意：由於 Passolo 的限制，該巨集的設定需通過執行檢查巨集來選取和設定。" & vbCrLf & vbCrLf & _
			"- 檢查巨集：PSLCheckAccessKeys.bas" & vbCrLf & _
			"  通過呼叫新增到 Passolo 選單中的該巨集，協助您檢查和修正翻譯中的便捷鍵、終止符、加速器，" & vbCrLf & _
			"  並翻譯終止符。此外，它還提供自訂設定和設定偵測功能。" & vbCrLf & vbCrLf & _
			"- 簡體中文說明檔案：AccessKey.txt" & vbCrLf & vbCrLf & vbCrLf
	Setup = "☆安裝方法☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 如果使用了 Wanfu 的 Passolo 漢化版，並安裝了附加巨集組件，則替換原來的檔案即可，否則：" & vbCrLf & _
			"  (1) 將解壓後的檔案複製到 Passolo 系統資料夾中定義的 Macros 資料夾中" & vbCrLf & _
			"  (2) 自動巨集：開啟 Passolo 的工具 -> 巨集對話方塊，將它設定為系統巨集並點擊主視窗的右下角的系統" & vbCrLf & _
			"  　  巨集啟用選單啟用它" & vbCrLf & _
			"  (3) 檢查巨集：在 Passolo 的工具 -> 自訂工具選單中新增該檔案" & vbCrLf & _
			"- 由於自動巨集無法在執行過程中進行設定，所以請使用檢查巨集來自訂設定專案。" & vbCrLf & _
			"- 請檢查後務必再逐條手工複查，以免程式處理錯誤。" & vbCrLf & vbCrLf & vbCrLf
	MainUse = "☆設定選取☆" & vbCrLf & _
			"============" & vbCrLf & _
			"程式提供了預設的設定，該設定可以適用於大多數情況。您也可以按 [設定] 按鈕自訂設定。" & vbCrLf & _
			"新增自訂設定後，您可以在設定清單中選取想使用的設定。" & vbCrLf & _
			"有關自訂設定，請開啟設定對話方塊，點擊 [說明] 按鈕，參閱說明中的說明。" & vbCrLf & vbCrLf & _
			"- 自動巨集設定" & vbCrLf & _
			"  選取的設定將用於自動巨集。請注意儲存，不然將使用選取前的設定。" & vbCrLf & vbCrLf & _
			"- 檢查巨集設定" & vbCrLf & _
			"  選取的設定將用於檢查巨集。要在下次使用選取的設定，需要儲存。" & vbCrLf & vbCrLf & _
			"- 自動巨集和檢查巨集相同" & vbCrLf & _
			"  選取該選項時，將自動使自動巨集的設定與檢查巨集的設定一致。" & vbCrLf & vbCrLf & _
			"- 自動選取" & vbCrLf & _
			"  選擇該選項時將根據設定中的適用語言清單自動選取與翻譯清單目標語言符合的設定。" & vbCrLf & _
			"  ◎注意：該選項僅對便捷鍵、終止符和加速器檢查巨集有效。" & vbCrLf & _
			"  　　　　要將設定與目前翻譯清單的目標語言符合，按 [設定] 按鈕，在字串處理的適用語言中" & vbCrLf & _
			"  　　　　新增對應的語言。" & vbCrLf & _
			"  　　　　自動巨集程式將按「自動 - 自選 - 預設」順序選取對應的設定。" & vbCrLf & vbCrLf & _
			"☆檢查標記☆" & vbCrLf & _
			"============" & vbCrLf & _
			"該功能通過記錄字串檢查訊息，並根據該訊息只對錯誤字串進行檢查，可大幅提高再檢查速度。" & vbCrLf & _
			"有以下 4 個選項可供選取：" & vbCrLf & vbCrLf & _
			"- 忽略版本" & vbCrLf & _
			"  將不考慮巨集程式的版本，僅根據其它記錄對錯誤字串進行檢查。" & vbCrLf & vbCrLf & _
			"- 忽略設定" & vbCrLf & _
			"  將不考慮設定是否相同，僅根據其它記錄對錯誤字串進行檢查。" & vbCrLf & vbCrLf & _
			"- 忽略日期" & vbCrLf & _
			"  將不考慮檢查日期和翻譯日期，僅根據其它記錄對錯誤字串進行檢查。" & vbCrLf & vbCrLf & _
			"- 全部忽略" & vbCrLf & _
			"  將不考慮任何檢查記錄，而對所有字串進行檢查。" & vbCrLf & vbCrLf & _
			"◎注意：檢查標記功能在測試設定時無效，以免可能出現有遺漏的測試結果。" & vbCrLf & _
			"　　　　如果變更設定內容而不變更設定名稱的話，請選擇全部忽略或刪除檢查標記選項。" & vbCrLf & vbCrLf & _
			"☆設定選取☆" & vbCrLf & _
			"============" & vbCrLf & _
			"有以下 3 個選項可供選取：" & vbCrLf & vbCrLf & _
			"- 僅檢查" & vbCrLf & _
			"  只對翻譯進行檢查，而不修正錯誤的翻譯。" & vbCrLf & vbCrLf & _
			"- 檢查並修正" & vbCrLf & _
			"  對翻譯進行檢查，並自動修正錯誤的翻譯。" & vbCrLf & vbCrLf & _
			"- 刪除便捷鍵" & vbCrLf & _
			"  刪除翻譯中現有的便捷鍵。" & vbCrLf & vbCrLf & _
			"☆字串類型☆" & vbCrLf & _
			"============" & vbCrLf & _
			"提供了全部、選單、對話方塊、字串表、加速器、版本、其他、僅選擇等選項。" & vbCrLf & vbCrLf & _
			"- 如果選取全部，則其他單項將被自動取消選取。" & vbCrLf & _
			"- 如果選取單項，則全部選項將被自動取消選取。" & vbCrLf & _
			"- 單項可以多選。其中選取僅選擇時，其他均被自動取消選取。" & vbCrLf & vbCrLf & _
			"☆字串內容☆" & vbCrLf & _
			"============" & vbCrLf & _
			"提供了全部、便捷鍵、終止符、加速器 4 個選項。" & vbCrLf & vbCrLf & _
			"- 如果選取全部，則其他單項將被自動取消選取。" & vbCrLf & _
			"- 如果選取單項，則全部選項將被自動取消選取。" & vbCrLf & _
			"- 單項可以多選。" & vbCrLf & vbCrLf & _
			"☆其他選項☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 不變更原始翻譯狀態" & vbCrLf & _
			"  選擇該項時，將在檢查、檢查並修正、刪除便捷鍵時不變更字串的原始翻譯狀態。否則，" & vbCrLf & _
			"  將變更無錯誤或無變更字串的翻譯狀態為已驗證狀態，有錯誤或已變更字串的翻譯狀態" & vbCrLf & _
			"  為待複審狀態，以便您一眼就可以知道哪些字串有錯誤或已被變更。" & vbCrLf & vbCrLf & _
			"- 不建立或刪除檢查標記" & vbCrLf & _
			"  選擇該項時，將不在專案中儲存檢查標記訊息，如果已存在檢查標記訊息，將被刪除。" & vbCrLf & _
			"  ◎注意：選擇該項時，檢查標記方塊中的全部忽略項將被選擇。" & vbCrLf & vbCrLf & _
			"- 繼續時自動儲存所有選取" & vbCrLf & _
			"  選擇該項時，將在按 [繼續] 按鈕時自動儲存所有選取，下次執行時將讀入儲存的選取。" & vbCrLf & _
			"  ◎注意：如果自動巨集設定被變更，系統將自動選擇該選項，以使選擇的自動巨集設定生效。" & vbCrLf & vbCrLf & _
			"- 替換特定字元" & vbCrLf & _
  			"  在檢查並修正過程中使用設定中定義的要自動替換的字元，替換字串中特定的字元。" & vbCrLf & vbCrLf & _
			"☆其他功能☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 關於" & vbCrLf & _
			"  點擊該按鈕，將關於對話方塊。並顯示程式介紹、執行環境、開發商及版權等訊息。" & vbCrLf& vbCrLf & _
			"- 說明" & vbCrLf & _
			"  點擊該按鈕，將彈出目前視窗的說明訊息。" & vbCrLf& vbCrLf & _
			"- 儲存所有選取" & vbCrLf & _
			"  該按鈕可以在變更設定而不進行檢查時使用。" & vbCrLf & _
			"  如果任一選項被變更，該選項將自動變為可用狀態，否則將自動變為不可用狀態。" & vbCrLf & vbCrLf & _
			"- 確定" & vbCrLf & _
			"  點擊該按鈕，將關閉主對話方塊，並按選擇的選項進行字串翻譯。" & vbCrLf& vbCrLf & _
			"- 取消" & vbCrLf & _
			"  點擊該按鈕，將結束程式。" & vbCrLf& vbCrLf & vbCrLf
	SetUse ="☆設定清單☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 選取設定" & vbCrLf & _
			"  要選取設定，點擊設定清單。" & vbCrLf & vbCrLf & _
			"- 設定設定的優先級" & vbCrLf & _
			"  設定優先級用於基於設定的適用語言的自動選取設定功能。" & vbCrLf & _
			"  ◎注意：有多個設定包含了相同的適用語言時，需要設定其優先級。" & vbCrLf & _
			"  　　　　在相同適用語言的設定中，前面的設定被優先選取使用。" & vbCrLf & _
			"  要設定設定的優先級，點擊右邊的 [...] 按鈕。" & vbCrLf & vbCrLf & _
			"- 新增設定" & vbCrLf & _
			"  要新增設定，點擊 [新增] 按鈕，在彈出的對話方塊中輸入名稱。" & vbCrLf & vbCrLf & _
			"- 變更設定" & vbCrLf & _
			"  要變更設定名稱，請選取設定清單中要改名的設定，然後點擊 [變更] 按鈕。" & vbCrLf & vbCrLf & _
			"- 刪除設定" & vbCrLf & _
			"  要刪除設定，請選取設定清單中要刪除的設定，然後點擊 [刪除] 按鈕。" & vbCrLf & vbCrLf & _
			"新增設定後，將在清單中顯示新的設定，設定內容將顯示空值。" & vbCrLf & _
			"變更設定後，將在清單中顯示改名的設定，設定內容中的設定值不變。" & vbCrLf & _
			"刪除設定後，將在清單中顯示預設設定，設定內容將顯示預設設定值。" & vbCrLf & vbCrLf & _
			"☆儲存類型☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 檔案" & vbCrLf & _
			"  設定將以檔案形式儲存在巨集所在資料夾下的 Data 資料夾中。" & vbCrLf & vbCrLf & _
			"- 註冊表" & vbCrLf & _
			"  設定將被儲存註冊表中的 HKCU\Software\VB and VBA Program Settings\AccessKey 項下。" & vbCrLf & vbCrLf & _
			"- 匯入設定" & vbCrLf & _
			"  允許從其他設定檔案中匯入設定。匯入舊設定時將被自動升級，現有設定清單中已有的設定將被" & vbCrLf & _
			"  變更，沒有的設定將被新增。" & vbCrLf & vbCrLf & _
			"- 匯出設定" & vbCrLf & _
			"  允許匯出所有設定到文字檔案，以便可以交換或轉移設定。" & vbCrLf & vbCrLf & _
			"◎注意：切換儲存類型時，將自動刪除原有位置中的設定內容。" & vbCrLf & vbCrLf & _
			"☆設定內容☆" & vbCrLf & _
			"============" & vbCrLf & _
			"<便捷鍵>" & vbCrLf & _
			"  - 要排除的含 & 符號的非便捷鍵組合" & vbCrLf & _
			"    便捷鍵以 & 為旗標符，有些字串雖然包含該符號但不是便捷鍵，需要排除它。在此輸入這些" & vbCrLf & _
			"    要排除包含 & 符號的非便捷鍵組合。" & vbCrLf & vbCrLf & _
			"  - 字串分割用旗標符" & vbCrLf & _
			"    這些字元用於分割含有多個便捷鍵，終止符或加速器的字串，以便檢查字串中所有的便捷鍵、" & vbCrLf & _
			"    終止符或加速器。否則只能處理字串最後部份的便捷鍵、終止符或加速器。" & vbCrLf & vbCrLf & _
			"  - 要檢查的便捷鍵前後括號" & vbCrLf & _
			"    預設的便捷鍵前後括號為 ()，在此指定的便捷鍵前後括號，都將被替換為預設的括號。" & vbCrLf & vbCrLf & _
			"  - 要保留的非便捷鍵前後成對字元" & vbCrLf & _
			"    如果原文和翻譯中都存在這些字元。將被保留，否則將被認為是便捷鍵，並將新增便捷鍵並移" & vbCrLf & _
			"    位到字串最後。" & vbCrLf & vbCrLf & _
			"  - 在文字後面顯示帶括號的便捷鍵 (通常用於亞洲語言)" & vbCrLf & _
			"    通常在亞洲語言如中文、日文等軟體中使用 (&X) 形式的便捷鍵，並將其置於字串結尾（在終" & vbCrLf & _
			"    止符或加速器前）。" & vbCrLf & _
			"    選擇該選項後，將檢查含有便捷鍵的翻譯字串中的便捷鍵是否符合慣例。如果不符合將被自動" & vbCrLf & _
			"    變更並置後。" & vbCrLf & vbCrLf & _
			"<終止符>" & vbCrLf & _
			"  - 要檢查的終止符" & vbCrLf & _
			"    翻譯中的終止符和原文不相同時，將被變更為原文中的終止符，但是符合要自動替換的終止符" & vbCrLf & _
			"    對中的終止符除外。" & vbCrLf & vbCrLf & _
			"    該欄位支援萬用字元，但是不是模糊的而是精確的。例如：A*C 不符合 XAYYCZ，只符合 AXYYC。" & vbCrLf & _
			"    要符合 XAYYCZ，應該為 *A*C* 或 *A??C*。" & vbCrLf & vbCrLf & _
			"    ◎注意：由於 Sax Basic 引擎的問題，二個非英文字母之間的 ? 萬用字元不被支援。" & vbCrLf & _
			"    　　　　例如：「開啟??檔案」不符合「開啟使用者檔案」。" & vbCrLf & vbCrLf & _
			"  - 要保留的終止符組合" & vbCrLf & _
			"    所有被包含在該組合中的要檢查的終止符將被保留。也就是這些終止符不被認為是終止符。" & vbCrLf & vbCrLf & _
			"    該欄位支援萬用字元，但是不是模糊的而是精確的。例如：A*C 不符合 XAYYCZ，只符合 AXYYC。" & vbCrLf & _
			"    要符合 XAYYCZ，應該為 *A*C* 或 *A??C*。" & vbCrLf & vbCrLf & _
			"    ◎注意：由於 Sax Basic 引擎的問題，二個非英文字母之間的 ? 萬用字元不被支援。" & vbCrLf & _
			"    　　　　例如：「開啟??檔案」不符合「開啟使用者檔案」。" & vbCrLf & vbCrLf & _
			"  - 要自動替換的終止符對" & vbCrLf & _
			"    符合終止符對中前一個字元的終止符，都將被替換成終止符對中後一個字元的終止符。" & vbCrLf & _
			"    利用此項可以自動翻譯或修正一些終止符。" & vbCrLf & vbCrLf & _
			"<加速器>" & vbCrLf & _
			"  - 要檢查的加速器旗標符" & vbCrLf & _
			"    加速器通常以 \t 為旗標符 (也有例外的)，如果字串中包含這些字元，將被認為包含加速器，" & vbCrLf & _
			"    但需要根據要檢查的加速器字元進一步判斷。" & vbCrLf & vbCrLf & _
			"  - 要檢查的加速器字元" & vbCrLf & _
			"    包含加速器旗標符的字串中，如果旗標符後面的字元符合該欄位的字元，將被辨識為加速器，" & vbCrLf & _
			"    二個以上字元組合而成的加速器允許其中一個不符合。" & vbCrLf & vbCrLf & _
			"    該欄位支援萬用字元，但是不是模糊的而是精確的。例如：A*C 不符合 XAYYCZ，只符合 AXYYC。" & vbCrLf & _
			"    要符合 XAYYCZ，應該為 *A*C* 或  *A??C*。" & vbCrLf & vbCrLf & _
			"    ◎注意：由於 Sax Basic 引擎的問題，二個非英文字母之間的 ? 萬用字元不被支援。" & vbCrLf & _
			"    　　　　例如：「開啟??檔案」不符合「開啟使用者檔案」。" & vbCrLf & vbCrLf & _
			"  - 要保留的加速器字元" & vbCrLf & _
			"    符合這些字元的加速器將被保留，否則將被替換。利用此項可保留某些加速器的翻譯。" & vbCrLf & vbCrLf & _
			"<字元替換>" & vbCrLf & _
			"  字串中包含每個替換字元對中的「|」前的字元時，將被替換成「|」後的字元。" & vbCrLf & vbCrLf & _
			"  - 要自動替換的字元" & vbCrLf & _
			"    定義在檢查並修正過程中要被替換的字元以及替換後的字元。" & vbCrLf & vbCrLf & _
			"  ◎注意：替換時區分大小寫。" & vbCrLf & _
			"  　　　　如果要去掉這些字元，可以將「|」後的字元置空。" & vbCrLf & vbCrLf & _
			"<適用語言>" & vbCrLf & _
			"  這裡的適用語言是指翻譯清單的目標語言，它用於根據翻譯清單的目標語言自動選取對應設定的" & vbCrLf & _
			"  自動選取功能。" & vbCrLf & vbCrLf & _
			"  - 新增" & vbCrLf & _
			"    要新增適用語言，選取可用語言清單中的語言，然後點擊 [新增] 按鈕。" & vbCrLf & _
			"    點擊該按鈕後，可用語言清單中的選擇語言將移動到適用語言清單中。" & vbCrLf & vbCrLf & _
			"  - 全部新增" & vbCrLf & _
			"    點擊該按鈕後，可用語言清單中的所有語言將全部移動到適用語言清單中。" & vbCrLf & vbCrLf & _
			"  - 刪除" & vbCrLf & _
			"    要刪除適用語言，選取適用語言清單中的語言，然後點擊 [刪除] 按鈕。" & vbCrLf & _
			"    點擊該按鈕後，適用語言清單中的選擇語言將移動到可用語言清單中。" & vbCrLf & vbCrLf & _
			"  - 全部刪除" & vbCrLf & _
			"    點擊該按鈕後，適用語言清單中的所有語言將全部移動到可用語言清單中。" & vbCrLf & vbCrLf & _
			"  - 增加可用語言" & vbCrLf & _
			"    點擊該按鈕後，將彈出可輸入語言名稱和代碼對話方塊，確定後將新增到可用語言清單中。" & vbCrLf & vbCrLf & _
			"  - 編輯可用語言" & vbCrLf & _
			"    要編輯可用語言，選取可用語言清單中的語言，然後點擊 [編輯可用語言] 按鈕。" & vbCrLf & _
			"    點擊該按鈕後，將彈出可編輯語言名稱和代碼對話方塊，確定後將修改可用語言清單中選擇的語言。" & vbCrLf & vbCrLf & _
			"  - 刪除可用語言" & vbCrLf & _
			"    要刪除可用語言，選取可用語言清單中要刪除的語言，然後點擊 [刪除可用語言] 按鈕。" & vbCrLf & vbCrLf & _
			"  - 增加適用語言" & vbCrLf & _
			"    點擊該按鈕後，將彈出可輸入語言名稱和代碼對話方塊，確定後將新增到適用語言清單中。" & vbCrLf & vbCrLf & _
			"  - 編輯適用語言" & vbCrLf & _
			"    要編輯適用語言，選取適用語言清單中的語言，然後點擊 [編輯適用語言] 按鈕。" & vbCrLf & _
			"    點擊該按鈕後，將彈出可編輯語言名稱和代碼對話方塊，確定後將修改適用語言清單中選擇的語言。" & vbCrLf & vbCrLf & _
			"  - 刪除適用語言" & vbCrLf & _
			"    要刪除適用語言，選取適用語言清單中要刪除的語言，然後點擊 [刪除適用語言] 按鈕。" & vbCrLf & vbCrLf & _
			"  ◎注意：新增、編輯語言僅用於 Passolo 未來版本新增的支援語言。" & vbCrLf & _
			"  　　　　語言代碼請和 Passolo 的 ISO 396-1 代碼保持一致，包括大小寫。" & vbCrLf & vbCrLf & _
			"☆其他功能☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 說明" & vbCrLf & _
			"  點擊該按鈕，將彈出目前視窗的說明訊息。" & vbCrLf & vbCrLf & _
			"- 讀取" & vbCrLf & _
			"  點擊該按鈕，將根據選擇設定的不同彈出下列選單:" & vbCrLf & _
			"  (1) 預設值" & vbCrLf & _
			"      讀取預設設定值，並顯示在設定內容中。" & vbCrLf & _
			"      ◎僅當選擇的設定為系統預設的設定時，才顯示該選單。" & vbCrLf & vbCrLf & _
			"  (2) 原值" & vbCrLf & _
			"      讀取選擇設定的原始值，並顯示在設定內容中。" & vbCrLf & _
			"      ◎僅當選擇設定的原始值為非空時，才顯示該選單。" & vbCrLf & vbCrLf & _
			"  (3) 參照值" & vbCrLf & _
			"      讀取選擇的參照設定值，並顯示在設定內容中。" & vbCrLf & _
			"      ◎該選單顯示除選擇設定外的所有設定清單。" & vbCrLf & vbCrLf & _
			"- 清空" & vbCrLf & _
			"  點擊該按鈕，將清空現有設定的全部值，以方便重新輸入設定值。" & vbCrLf & vbCrLf & _
			"- 測試" & vbCrLf & _
			"  點擊該按鈕，將彈出測試對話方塊，以便檢查設定的正確性。" & vbCrLf & vbCrLf & _
			"- 確定" & vbCrLf & _
			"  點擊該按鈕，將儲存設定視窗中的任何變更，結束設定視窗並返回主視窗。" & vbCrLf & _
			"  程式將使用變更後的設定值。" & vbCrLf & vbCrLf & _
			"- 取消" & vbCrLf & _
			"  點擊該按鈕，不儲存設定視窗中的任何變更，結束設定視窗並返回主視窗。" & vbCrLf & _
			"  程式將使用原來的設定值。" & vbCrLf & vbCrLf & vbCrLf
	TestUse ="☆設定名稱☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 要測試的設定名稱。要選取設定，點擊設定清單。" & vbCrLf & vbCrLf & _
			"☆翻譯清單☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 該清單將顯示專案中的所有翻譯清單。請選取與您的自訂設定符合的翻譯清單進行測試。" & vbCrLf & vbCrLf & _
			"☆自動替換字元☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 在檢查時自動替換字串中符合設定中所定義的替換字元。" & vbCrLf & vbCrLf & _
			"☆讀入行數☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 表示要顯示的錯誤翻譯字串數。建議不要輸入太大的值，以免在字串較多時等待時間過長。" & vbCrLf & vbCrLf & _
			"☆包含內容☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 指定只檢查包含內容的字串。利用該項可以有針對性的測試，並且加快測試時間。" & vbCrLf & _
			"- 該欄位支援模糊型萬用字元。例如：A*C 可以符合 XAYYCZ。" & vbCrLf & vbCrLf & _
			"◎注意：由於 Sax Basic 引擎的問題，二個非英文字母之間的 ? 萬用字元不被支援。" & vbCrLf & _
			"　　　　例如：「開啟??檔案」不符合「開啟使用者檔案」。" & vbCrLf & vbCrLf & _
			"☆字串內容☆" & vbCrLf & _
			"============" & vbCrLf & _
			"提供了全部、便捷鍵、終止符、加速器 4 個選項。" & vbCrLf & vbCrLf & _
			"- 如果選取全部，則其他單項將被自動取消選取。" & vbCrLf & _
			"- 如果選取單項，則全部選項將被自動取消選取。" & vbCrLf & _
			"- 單項可以多選。" & vbCrLf & vbCrLf & _
			"☆其他功能☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 說明" & vbCrLf & _
			"  點擊該按鈕，將彈出目前視窗的說明訊息。" & vbCrLf & vbCrLf & _
			"- 測試" & vbCrLf & _
			"  點擊該按鈕，將按照選擇的條件進行測試。" & vbCrLf & vbCrLf & _
			"- 清空" & vbCrLf & _
			"  點擊該按鈕，將清空現有的測試結果。" & vbCrLf & vbCrLf & _
			"- 取消" & vbCrLf & _
			"  點擊該按鈕，將結束測試程式並返回設定視窗。" & vbCrLf & vbCrLf & vbCrLf
	CopyRight = "☆版權宣告☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 此軟體的版權歸開發者和修改者所有，任何人可以免費使用、修改、複製、散佈本軟體。" & vbCrLf & _
			"- 修改、散佈本軟體必須隨附本說明檔案，並註明軟體原始開發者以及修改者。" & vbCrLf & _
			"- 未經開發者和修改者同意，任何組織或個人，不得用於商業軟體、商業或是其它營利性活動。" & vbCrLf & _
			"- 對使用本軟體的原始版本，以及使用經他人修改的非原始版本所造成的損失和損害，開發者不" & vbCrLf & _
			"  承擔任何責任。" & vbCrLf & _
			"- 由於為免費軟體，開發者和修改者沒有義務提供軟體技術支援，也無義務改進或更新版本。" & vbCrLf & _
			"- 歡迎指正錯誤並提出改進意見。如有錯誤或建議，請傳送到: z_shangyi@163.com。" & vbCrLf & vbCrLf & vbCrLf
	Thank = "☆致　　謝☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 本軟體在修改過程中得到漢化新世紀會員的測試，在此表示衷心的感謝！" & vbCrLf & _
			"- 感謝台灣 Heaven 先生提出正體用語修改意見！" & vbCrLf & vbCrLf & vbCrLf
	Contact = "☆與我聯繫☆" & vbCrLf & _
			"============" & vbCrLf & _
			"wanfu：z_shangyi@163.com" & vbCrLf & vbCrLf & vbCrLf
	Logs = "感謝支持！您的支持是我最大的動力！同時歡迎使用我們製作的軟體！" & vbCrLf & vbCrLf & _
			"**************************************************" & vbCrLf & _
			"需要更多、更新、更好的漢化，請拜訪:" & vbCrLf & _
			"漢化新世紀 -- http://www.hanzify.org" & vbCrLf & _
			"漢化新世紀論壇 -- http://bbs.hanzify.org" & vbCrLf & _
			"**************************************************" & vbCrLf
	Else
	AboutTitle = "壽衾"
	HelpTitle = "堆翑"
	HelpTipTitle = "辦豎瑩﹜笝砦睫睿樓厒ん潰脤粽"
	AboutWindows = " 壽衾 "
	MainWindows = " 翋敦諳 "
	SetWindows = " 饜离敦諳 "
	TestWindows = " 聆彸敦諳 "
	Lines = "-----------------------"
	Sys = "�篲�唳掛ㄩ" & Version & vbCrLf & _
			"巠蚚炵苀ㄩWindows XP/2000 眕奻炵苀" & vbCrLf & _
			"巠蚚唳掛ㄩ垀衄盓厥粽揭燴腔 Passolo 6.0 摯眕奻唳掛" & vbCrLf & _
			"賜醱逄晟ㄩ潠极笢恅睿楛极笢恅 (赻雄妎梗)" & vbCrLf & _
			"唳�佯齾苺犖獄耘薹擘�" & vbCrLf & _
			"忨�佬恀膛疑漞捑篲�" & vbCrLf & _
			"夥源翋珜ㄩhttp://www.hanzify.org" & vbCrLf & _
			"ヶ羲楷氪ㄩ犖趙陔岍槨傖埜 gnatix (2007-2008)" & vbCrLf & _
			"綴羲楷氪ㄩ犖趙陔岍槨傖埜 wanfu (2009-2010)" & vbCrLf & vbCrLf & vbCrLf
	Ement = "∵堍俴遠噫∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 盓厥粽揭燴腔 Passolo 6.0 摯眕奻唳掛ㄛ斛剒" & vbCrLf & _
			"- Windows Script Host (WSH) 勤砓 (VBS)ㄛ斛剒" & vbCrLf & _
			"- Adodb.Stream 勤砓 (VBS)ㄛ盓厥赻雄載陔垀剒" & vbCrLf & _
			"- Microsoft.XMLHTTP 勤砓ㄛ盓厥赻雄載陔垀剒" & vbCrLf & vbCrLf & vbCrLf
	Dec = "∵�篲�潠賡∵" & vbCrLf & _
			"============" & vbCrLf & _
			"辦豎瑩﹜笝砦睫睿樓厒ん潰脤粽岆珨跺蚚衾 Passolo 楹祒潰脤腔粽最唗﹝坳撿衄眕狟髡夔ㄩ" & vbCrLf & _
			"- 潰脤楹祒笢辦豎瑩﹜笝砦睫﹜樓厒ん睿諾跡" & vbCrLf & _
			"- 潰脤甜党淏潰脤楹祒笢辦豎瑩﹜笝砦睫﹜樓厒ん睿諾跡" & vbCrLf & _
			"- 刉壺楹祒笢腔辦豎瑩" & vbCrLf & _
			"- 囀离褫赻隅砱腔赻雄載陔髡夔" & vbCrLf & vbCrLf & _
			"掛最唗婦漪狟蹈恅璃ㄩ" & vbCrLf & _
			"- 赻雄粽ㄩPslAutoAccessKey.bas" & vbCrLf & _
			"  婓楹祒趼揹奀ㄛ赻雄載淏渣昫腔楹祒﹝瞳蚚蜆粽ㄛ蠟褫眕祥斛怀�踸儠敯�﹜笝砦睫﹜樓厒んㄛ" & vbCrLf & _
			"  炵苀蔚跦擂蠟恁寁腔饜离赻雄堆蠟氝樓睿埻恅珨欴腔辦豎瑩﹜笝砦睫﹜樓厒んㄛ甜楹祒笝砦睫﹝" & vbCrLf & _
			"  瞳蚚坳褫眕枑詢楹祒厒僅ㄛ甜熬屾楹祒渣昫﹝" & vbCrLf & _
			"  ♁蛁砩ㄩ蚕衾 Passolo 腔癹秶ㄛ蜆粽腔饜离剒籵徹堍俴潰脤粽懂恁寁睿扢离﹝" & vbCrLf & vbCrLf & _
			"- 潰脤粽ㄩPSLCheckAccessKeys.bas" & vbCrLf & _
			"  籵徹覃蚚氝樓善 Passolo 粕等笢腔蜆粽ㄛ堆翑蠟潰脤睿党淏楹祒笢腔辦豎瑩﹜笝砦睫﹜樓厒んㄛ" & vbCrLf & _
			"  甜楹祒笝砦睫﹝森俋ㄛ坳遜枑鼎赻隅砱饜离睿饜离潰聆髡夔﹝" & vbCrLf & vbCrLf & _
			"- 潠极笢恅佽隴恅璃ㄩAccessKey.txt" & vbCrLf & vbCrLf & vbCrLf
	Setup = "∵假蚾源楊∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- �蝜�妏蚚賸 Wanfu 腔 Passolo 犖趙唳ㄛ甜假蚾賸蜇樓粽郪璃ㄛ寀杸遙埻懂腔恅璃撈褫ㄛ瘁寀ㄩ" & vbCrLf & _
			"  (1) 蔚賤揤綴腔恅璃葩秶善 Passolo 炵苀恅璃標笢隅砱腔 Macros 恅璃標笢" & vbCrLf & _
			"  (2) 赻雄粽ㄩ湖羲 Passolo 腔馱撿 -> 粽勤趕遺ㄛ蔚坳扢离峈炵苀粽甜等僻翋敦諳腔衵狟褒腔炵苀" & vbCrLf & _
			"  ﹛  粽慾魂粕等慾魂坳" & vbCrLf & _
			"  (3) 潰脤粽ㄩ婓 Passolo 腔馱撿 -> 赻隅砱馱撿粕等笢氝樓蜆恅璃" & vbCrLf & _
			"- 蚕衾赻雄粽拸楊婓堍俴徹最笢輛俴饜离ㄛ垀眕③妏蚚潰脤粽懂赻隅砱饜离源偶﹝" & vbCrLf & _
			"- ③潰脤綴昢斛婬紨沭忒馱葩脤ㄛ眕轎最唗揭燴渣昫﹝" & vbCrLf & vbCrLf & vbCrLf
	MainUse = "∵饜离恁寁∵" & vbCrLf & _
			"============" & vbCrLf & _
			"最唗枑鼎賸蘇�炵霰馺瓊爰藷馺藩奿埰弝譚痟騥鉌�①錶﹝蠟珩褫眕偌 [扢离] 偌聽赻隅砱饜离﹝" & vbCrLf & _
			"氝樓赻隅砱饜离綴ㄛ蠟褫眕婓饜离蹈桶笢恁寁砑妏蚚腔饜离﹝" & vbCrLf & _
			"衄壽赻隅砱饜离ㄛ③湖羲饜离勤趕遺ㄛ等僻 [堆翑] 偌聽ㄛ統堐堆翑笢腔佽隴﹝" & vbCrLf & vbCrLf & _
			"- 赻雄粽饜离" & vbCrLf & _
			"  恁寁腔饜离蔚蚚衾赻雄粽﹝③蛁砩悵湔ㄛ祥�遢封墓識√鯁做霰馺獺�" & vbCrLf & vbCrLf & _
			"- 潰脤粽饜离" & vbCrLf & _
			"  恁寁腔饜离蔚蚚衾潰脤粽﹝猁婓狟棒妏蚚恁寁腔饜离ㄛ剒猁悵湔﹝" & vbCrLf & vbCrLf & _
			"- 赻雄粽睿潰脤粽眈肮" & vbCrLf & _
			"  恁寁蜆恁砐奀ㄛ蔚赻雄妏赻雄粽腔饜离迵潰脤粽腔饜离珨祡﹝" & vbCrLf & vbCrLf & _
			"- 赻雄恁寁" & vbCrLf & _
			"  恁隅蜆恁砐奀蔚跦擂饜离笢腔巠蚚逄晟蹈桶赻雄恁寁迵楹祒蹈桶醴梓逄晟ぁ饜腔饜离﹝" & vbCrLf & _
			"  ♁蛁砩ㄩ蜆恁砐躺勤辦豎瑩﹜笝砦睫睿樓厒ん潰脤粽衄虴﹝" & vbCrLf & _
			"  ﹛﹛﹛﹛猁蔚饜离迵絞ヶ楹祒蹈桶腔醴梓逄晟ぁ饜ㄛ偌 [扢离] 偌聽ㄛ婓趼揹揭燴腔巠蚚逄晟笢" & vbCrLf & _
			"  ﹛﹛﹛﹛氝樓眈茼腔逄晟﹝" & vbCrLf & _
			"  ﹛﹛﹛﹛赻雄粽最唗蔚偌※赻雄 - 赻恁 - 蘇�洁捨創藘√鵜隑朴霰馺獺�" & vbCrLf & vbCrLf & _
			"∵潰脤梓暮∵" & vbCrLf & _
			"============" & vbCrLf & _
			"蜆髡夔籵徹暮翹趼揹潰脤陓洘ㄛ甜跦擂蜆陓洘硐勤渣昫趼揹輛俴潰脤ㄛ褫湮盟枑詢婬潰脤厒僅﹝" & vbCrLf & _
			"衄眕狟 4 跺恁砐褫鼎恁寁ㄩ" & vbCrLf & vbCrLf & _
			"- 綺謹唳掛" & vbCrLf & _
			"  蔚祥蕉藉粽最唗腔唳掛ㄛ躺跦擂む坳暮翹勤渣昫趼揹輛俴潰脤﹝" & vbCrLf & vbCrLf & _
			"- 綺謹饜离" & vbCrLf & _
			"  蔚祥蕉藉饜离岆瘁眈肮ㄛ躺跦擂む坳暮翹勤渣昫趼揹輛俴潰脤﹝" & vbCrLf & vbCrLf & _
			"- 綺謹�梪�" & vbCrLf & _
			"  蔚祥蕉藉潰脤�梪睆芛倡躽梪琭狠鷏躨暔頖�暮翹勤渣昫趼揹輛俴潰脤﹝" & vbCrLf & vbCrLf & _
			"- �垓蕩鷈�" & vbCrLf & _
			"  蔚祥蕉藉�庥弮麮曌Ъ慫炮禷堍齾倜硒捐靇邾麮憿�" & vbCrLf & vbCrLf & _
			"♁蛁砩ㄩ潰脤梓暮髡夔婓聆彸饜离奀拸虴ㄛ眕轎褫夔堤珋衄疻穢腔聆彸賦彆﹝" & vbCrLf & _
			"﹛﹛﹛﹛�蝜�載蜊饜离囀�搋瓛遘�蜊饜离靡備腔趕ㄛ③恁隅�垓蕩鷈堇藙噫�潰脤梓暮恁砐﹝" & vbCrLf & vbCrLf & _
			"∵饜离恁寁∵" & vbCrLf & _
			"============" & vbCrLf & _
			"衄眕狟 3 跺恁砐褫鼎恁寁ㄩ" & vbCrLf & vbCrLf & _
			"- 躺潰脤" & vbCrLf & _
			"  硐勤楹祒輛俴潰脤ㄛ奧祥党淏渣昫腔楹祒﹝" & vbCrLf & vbCrLf & _
			"- 潰脤甜党淏" & vbCrLf & _
			"  勤楹祒輛俴潰脤ㄛ甜赻雄党淏渣昫腔楹祒﹝" & vbCrLf & vbCrLf & _
			"- 刉壺辦豎瑩" & vbCrLf & _
			"  刉壺楹祒笢珋衄腔辦豎瑩﹝" & vbCrLf & vbCrLf & _
			"∵趼揹濬倰∵" & vbCrLf & _
			"============" & vbCrLf & _
			"枑鼎賸�垓縑３佽央７堇倏礡Ｉ硊�揹桶﹜樓厒ん﹜唳掛﹜む坻﹜躺恁隅脹恁砐﹝" & vbCrLf & vbCrLf & _
			"- �蝜�恁寁�垓縛皈藫頖�等砐蔚掩赻雄�＋�恁寁﹝" & vbCrLf & _
			"- �蝜�恁寁等砐ㄛ寀�垓謀＋蹐垮閤堈紙＋�恁寁﹝" & vbCrLf & _
			"- 等砐褫眕嗣恁﹝む笢恁寁躺恁隅奀ㄛむ坻歙掩赻雄�＋�恁寁﹝" & vbCrLf & vbCrLf & _
			"∵趼揹囀�搳�" & vbCrLf & _
			"============" & vbCrLf & _
			"枑鼎賸�垓縑Ⅶ儠敯�﹜笝砦睫﹜樓厒ん 4 跺恁砐﹝" & vbCrLf & vbCrLf & _
			"- �蝜�恁寁�垓縛皈藫頖�等砐蔚掩赻雄�＋�恁寁﹝" & vbCrLf & _
			"- �蝜�恁寁等砐ㄛ寀�垓謀＋蹐垮閤堈紙＋�恁寁﹝" & vbCrLf & _
			"- 等砐褫眕嗣恁﹝" & vbCrLf & vbCrLf & _
			"∵む坻恁砐∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 祥載蜊埻宎楹祒袨怓" & vbCrLf & _
			"  恁隅蜆砐奀ㄛ蔚婓潰脤﹜潰脤甜党淏﹜刉壺辦豎瑩奀祥載蜊趼揹腔埻宎楹祒袨怓﹝瘁寀ㄛ" & vbCrLf & _
			"  蔚載蜊拸渣昫麼拸載蜊趼揹腔楹祒袨怓峈眒桄痐袨怓ㄛ衄渣昫麼眒載蜊趼揹腔楹祒袨怓" & vbCrLf & _
			"  峈渾葩机袨怓ㄛ眕晞蠟珨桉憩褫眕眭耋闡虳趼揹衄渣昫麼眒掩載蜊﹝" & vbCrLf & vbCrLf & _
			"- 祥斐膘麼刉壺潰脤梓暮" & vbCrLf & _
			"  恁隅蜆砐奀ㄛ蔚祥婓源偶笢悵湔潰脤梓暮陓洘ㄛ�蝜�眒湔婓潰脤梓暮陓洘ㄛ蔚掩刉壺﹝" & vbCrLf & _
			"  ♁蛁砩ㄩ恁隅蜆砐奀ㄛ潰脤梓暮遺笢腔�垓蕩鷈婘蹐垮銑▲芋�" & vbCrLf & vbCrLf & _
			"- 樟哿奀赻雄悵湔垀衄恁寁" & vbCrLf & _
			"  恁隅蜆砐奀ㄛ蔚婓偌 [樟哿] 偌聽奀赻雄悵湔垀衄恁寁ㄛ狟棒堍俴奀蔚黍�貑ㄣ瘚麵√鞢�" & vbCrLf & _
			"  ♁蛁砩ㄩ�蝜�赻雄粽饜离掩載蜊ㄛ炵苀蔚赻雄恁隅蜆恁砐ㄛ眕妏恁隅腔赻雄粽饜离汜虴﹝" & vbCrLf & vbCrLf & _
			"- 杸遙杻隅趼睫" & vbCrLf & _
  			"  婓潰脤甜党淏徹最笢妏蚚饜离笢隅砱腔猁赻雄杸遙腔趼睫ㄛ杸遙趼揹笢杻隅腔趼睫﹝" & vbCrLf & vbCrLf & _
			"∵む坻髡夔∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 壽衾" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚壽衾勤趕遺﹝甜珆尨最唗賡庄﹜堍俴遠噫﹜羲楷妀摯唳�巡�陓洘﹝" & vbCrLf& vbCrLf & _
			"- 堆翑" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚粟堤絞ヶ敦諳腔堆翑陓洘﹝" & vbCrLf& vbCrLf & _
			"- 悵湔垀衄恁寁" & vbCrLf & _
			"  蜆偌聽褫眕婓載蜊饜离奧祥輛俴潰脤奀妏蚚﹝" & vbCrLf & _
			"  �蝜��扂銑＋豏遘�蜊ㄛ蜆恁砐蔚赻雄曹峈褫蚚袨怓ㄛ瘁寀蔚赻雄曹峈祥褫蚚袨怓﹝" & vbCrLf & vbCrLf & _
			"- �毓�" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚壽敕翋勤趕遺ㄛ甜偌恁隅腔恁砐輛俴趼揹楹祒﹝" & vbCrLf& vbCrLf & _
			"- �＋�" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚豖堤最唗﹝" & vbCrLf& vbCrLf & vbCrLf
	SetUse ="∵饜离蹈桶∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 恁寁饜离" & vbCrLf & _
			"  猁恁寁饜离ㄛ等僻饜离蹈桶﹝" & vbCrLf & vbCrLf & _
			"- 扢离饜离腔蚥珂撰" & vbCrLf & _
			"  饜离蚥珂撰蚚衾價衾饜离腔巠蚚逄晟腔赻雄恁寁饜离髡夔﹝" & vbCrLf & _
			"  ♁蛁砩ㄩ衄嗣跺饜离婦漪賸眈肮腔巠蚚逄晟奀ㄛ剒猁扢离む蚥珂撰﹝" & vbCrLf & _
			"  ﹛﹛﹛﹛婓眈肮巠蚚逄晟腔饜离笢ㄛヶ醱腔饜离掩蚥珂恁寁妏蚚﹝" & vbCrLf & _
			"  猁扢离饜离腔蚥珂撰ㄛ等僻衵晚腔 [...] 偌聽﹝" & vbCrLf & vbCrLf & _
			"- 氝樓饜离" & vbCrLf & _
			"  猁氝樓饜离ㄛ等僻 [氝樓] 偌聽ㄛ婓粟堤腔勤趕遺笢怀�踼�備﹝" & vbCrLf & vbCrLf & _
			"- 載蜊饜离" & vbCrLf & _
			"  猁載蜊饜离靡備ㄛ③恁寁饜离蹈桶笢猁蜊靡腔饜离ㄛ�遣騕本� [載蜊] 偌聽﹝" & vbCrLf & vbCrLf & _
			"- 刉壺饜离" & vbCrLf & _
			"  猁刉壺饜离ㄛ③恁寁饜离蹈桶笢猁刉壺腔饜离ㄛ�遣騕本� [刉壺] 偌聽﹝" & vbCrLf & vbCrLf & _
			"氝樓饜离綴ㄛ蔚婓蹈桶笢珆尨陔腔饜离ㄛ饜离囀�斒峙埰噶欶窗�" & vbCrLf & _
			"載蜊饜离綴ㄛ蔚婓蹈桶笢珆尨蜊靡腔饜离ㄛ饜离囀�楺迮霰馺譆結跼銦�" & vbCrLf & _
			"刉壺饜离綴ㄛ蔚婓蹈桶笢珆尨蘇�狣馺瓊玳馺藥硜斒峙埰奮畏狣馺譆窗�" & vbCrLf & vbCrLf & _
			"∵悵湔濬倰∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 恅璃" & vbCrLf & _
			"  饜离蔚眕恅璃倛宒悵湔婓粽垀婓恅璃標狟腔 Data 恅璃標笢﹝" & vbCrLf & vbCrLf & _
			"- 蛁聊桶" & vbCrLf & _
			"  饜离蔚掩悵湔蛁聊桶笢腔 HKCU\Software\VB and VBA Program Settings\AccessKey 砐狟﹝" & vbCrLf & vbCrLf & _
			"- 絳�蹁馺�" & vbCrLf & _
			"  埰勍植む坻饜离恅璃笢絳�蹁馺獺ㄤ暫踾吇馺蟾掃垮閤堈紛�撰ㄛ珋衄饜离蹈桶笢眒衄腔饜离蔚掩" & vbCrLf & _
			"  載蜊ㄛ羶衄腔饜离蔚掩氝樓﹝" & vbCrLf & vbCrLf & _
			"- 絳堤饜离" & vbCrLf & _
			"  埰勍絳堤垀衄饜离善恅掛恅璃ㄛ眕晞褫眕蝠遙麼蛌痄饜离﹝" & vbCrLf & vbCrLf & _
			"♁蛁砩ㄩз遙悵湔濬倰奀ㄛ蔚赻雄刉壺埻衄弇离笢腔饜离囀�搳�" & vbCrLf & vbCrLf & _
			"∵饜离囀�搳�" & vbCrLf & _
			"============" & vbCrLf & _
			"<辦豎瑩>" & vbCrLf & _
			"  - 猁齬壺腔漪 & 睫瘍腔準辦豎瑩郪磁" & vbCrLf & _
			"    辦豎瑩眕 & 峈梓祩睫ㄛ衄虳趼揹呥�趕�漪蜆睫瘍筍祥岆辦豎瑩ㄛ剒猁齬壺坳﹝婓森怀�鄳瑱�" & vbCrLf & _
			"    猁齬壺婦漪 & 睫瘍腔準辦豎瑩郪磁﹝" & vbCrLf & vbCrLf & _
			"  - 趼揹莞煦蚚梓祩睫" & vbCrLf & _
			"    涴虳趼睫蚚衾莞煦漪衄嗣跺辦豎瑩ㄛ笝砦睫麼樓厒ん腔趼揹ㄛ眕晞潰脤趼揹笢垀衄腔辦豎瑩﹜" & vbCrLf & _
			"    笝砦睫麼樓厒ん﹝瘁寀硐夔揭燴趼揹郔綴窒煦腔辦豎瑩﹜笝砦睫麼樓厒ん﹝" & vbCrLf & vbCrLf & _
			"  - 猁潰脤腔辦豎瑩ヶ綴嬤瘍" & vbCrLf & _
			"    蘇�炵醴儠敯�ヶ綴嬤瘍峈 ()ㄛ婓森硌隅腔辦豎瑩ヶ綴嬤瘍ㄛ飲蔚掩杸遙峈蘇�炵釋那禳�" & vbCrLf & vbCrLf & _
			"  - 猁悵隱腔準辦豎瑩ヶ綴傖勤趼睫" & vbCrLf & _
			"    �蝜�埻恅睿楹祒笢飲湔婓涴虳趼睫﹝蔚掩悵隱ㄛ瘁寀蔚掩�玴羌Ч儠敯�ㄛ甜蔚氝樓辦豎瑩甜痄" & vbCrLf & _
			"    弇善趼揹郔綴﹝" & vbCrLf & vbCrLf & _
			"  - 婓恅掛綴醱珆尨湍嬤瘍腔辦豎瑩 (籵都蚚衾捚粔逄晟)" & vbCrLf & _
			"    籵都婓捚粔逄晟�誸倛纂〦梉警��篲�笢妏蚚 (&X) 倛宒腔辦豎瑩ㄛ甜蔚む离衾趼揹賦帣ㄗ婓笝" & vbCrLf & _
			"    砦睫麼樓厒んヶㄘ﹝" & vbCrLf & _
			"    恁隅蜆恁砐綴ㄛ蔚潰脤漪衄辦豎瑩腔楹祒趼揹笢腔辦豎瑩岆瘁睫磁嫦瞰﹝�蝜�祥睫磁蔚掩赻雄" & vbCrLf & _
			"    載蜊甜离綴﹝" & vbCrLf & vbCrLf & _
			"<笝砦睫>" & vbCrLf & _
			"  - 猁潰脤腔笝砦睫" & vbCrLf & _
			"    楹祒笢腔笝砦睫睿埻恅祥珨祡奀ㄛ蔚掩載蜊峈埻恅笢腔笝砦睫ㄛ筍岆睫磁猁赻雄杸遙腔笝砦睫" & vbCrLf & _
			"    勤笢腔笝砦睫壺俋﹝" & vbCrLf & vbCrLf & _
			"    蜆趼僇盓厥籵饜睫ㄛ筍岆祥岆耀緇腔奧岆儕�殿纂�瞰�蝤態*C 祥ぁ饜 XAYYCZㄛ硐ぁ饜 AXYYC﹝" & vbCrLf & _
			"    猁ぁ饜 XAYYCZㄛ茼蜆峈 *A*C* 麼 *A??C*﹝" & vbCrLf & vbCrLf & _
			"    ♁蛁砩ㄩ蚕衾 Sax Basic 竘э腔恀枙ㄛ媼跺準荎恅趼譫眳潔腔 ? 籵饜睫祥掩盓厥﹝" & vbCrLf & _
			"    ﹛﹛﹛﹛瞰�蝤滿偌翾�??恅璃§祥ぁ饜※湖羲蚚誧恅璃§﹝" & vbCrLf & vbCrLf & _
			"  - 猁悵隱腔笝砦睫郪磁" & vbCrLf & _
			"    垀衄掩婦漪婓蜆郪磁笢腔猁潰脤腔笝砦睫蔚掩悵隱﹝珩憩岆涴虳笝砦睫祥掩�玴羌н欶僩�﹝" & vbCrLf & vbCrLf & _
			"    蜆趼僇盓厥籵饜睫ㄛ筍岆祥岆耀緇腔奧岆儕�殿纂�瞰�蝤態*C 祥ぁ饜 XAYYCZㄛ硐ぁ饜 AXYYC﹝" & vbCrLf & _
			"    猁ぁ饜 XAYYCZㄛ茼蜆峈 *A*C* 麼 *A??C*﹝" & vbCrLf & vbCrLf & _
			"    ♁蛁砩ㄩ蚕衾 Sax Basic 竘э腔恀枙ㄛ媼跺準荎恅趼譫眳潔腔 ? 籵饜睫祥掩盓厥﹝" & vbCrLf & _
			"    ﹛﹛﹛﹛瞰�蝤滿偌翾�??恅璃§祥ぁ饜※湖羲蚚誧恅璃§﹝" & vbCrLf & vbCrLf & _
			"  - 猁赻雄杸遙腔笝砦睫勤" & vbCrLf & _
			"    睫磁笝砦睫勤笢ヶ珨跺趼睫腔笝砦睫ㄛ飲蔚掩杸遙傖笝砦睫勤笢綴珨跺趼睫腔笝砦睫﹝" & vbCrLf & _
			"    瞳蚚森砐褫眕赻雄楹祒麼党淏珨虳笝砦睫﹝" & vbCrLf & vbCrLf & _
			"<樓厒ん>" & vbCrLf & _
			"  - 猁潰脤腔樓厒ん梓祩睫" & vbCrLf & _
			"    樓厒ん籵都眕 \t 峈梓祩睫 (珩衄瞰俋腔)ㄛ�蝜�趼揹笢婦漪涴虳趼睫ㄛ蔚掩�玴狐�漪樓厒んㄛ" & vbCrLf & _
			"    筍剒猁跦擂猁潰脤腔樓厒ん趼睫輛珨祭瓚剿﹝" & vbCrLf & vbCrLf & _
			"  - 猁潰脤腔樓厒ん趼睫" & vbCrLf & _
			"    婦漪樓厒ん梓祩睫腔趼揹笢ㄛ�蝜�梓祩睫綴醱腔趼睫睫磁蜆趼僇腔趼睫ㄛ蔚掩妎梗峈樓厒んㄛ" & vbCrLf & _
			"    媼跺眕奻趼睫郪磁奧傖腔樓厒ん埰勍む笢珨跺祥ぁ饜﹝" & vbCrLf & vbCrLf & _
			"    蜆趼僇盓厥籵饜睫ㄛ筍岆祥岆耀緇腔奧岆儕�殿纂�瞰�蝤態*C 祥ぁ饜 XAYYCZㄛ硐ぁ饜 AXYYC﹝" & vbCrLf & _
			"    猁ぁ饜 XAYYCZㄛ茼蜆峈 *A*C* 麼  *A??C*﹝" & vbCrLf & vbCrLf & _
			"    ♁蛁砩ㄩ蚕衾 Sax Basic 竘э腔恀枙ㄛ媼跺準荎恅趼譫眳潔腔 ? 籵饜睫祥掩盓厥﹝" & vbCrLf & _
			"    ﹛﹛﹛﹛瞰�蝤滿偌翾�??恅璃§祥ぁ饜※湖羲蚚誧恅璃§﹝" & vbCrLf & vbCrLf & _
			"  - 猁悵隱腔樓厒ん趼睫" & vbCrLf & _
			"    睫磁涴虳趼睫腔樓厒ん蔚掩悵隱ㄛ瘁寀蔚掩杸遙﹝瞳蚚森砐褫悵隱議虳樓厒ん腔楹祒﹝" & vbCrLf & vbCrLf & _
			"<趼睫杸遙>" & vbCrLf & _
			"  趼揹笢婦漪藩跺杸遙趼睫勤笢腔※|§ヶ腔趼睫奀ㄛ蔚掩杸遙傖※|§綴腔趼睫﹝" & vbCrLf & vbCrLf & _
			"  - 猁赻雄杸遙腔趼睫" & vbCrLf & _
			"    隅砱婓潰脤甜党淏徹最笢猁掩杸遙腔趼睫眕摯杸遙綴腔趼睫﹝" & vbCrLf & vbCrLf & _
			"  ♁蛁砩ㄩ杸遙奀⑹煦湮苤迡﹝" & vbCrLf & _
			"  ﹛﹛﹛﹛�蝜�猁�扔譭瑱抿硊�ㄛ褫眕蔚※|§綴腔趼睫离諾﹝" & vbCrLf & vbCrLf & _
			"<巠蚚逄晟>" & vbCrLf & _
			"  涴爵腔巠蚚逄晟岆硌楹祒蹈桶腔醴梓逄晟ㄛ坳蚚衾跦擂楹祒蹈桶腔醴梓逄晟赻雄恁寁眈茼饜离腔" & vbCrLf & _
			"  赻雄恁寁髡夔﹝" & vbCrLf & vbCrLf & _
			"  - 氝樓" & vbCrLf & _
			"    猁氝樓巠蚚逄晟ㄛ恁寁褫蚚逄晟蹈桶笢腔逄晟ㄛ�遣騕本� [氝樓] 偌聽﹝" & vbCrLf & _
			"    等僻蜆偌聽綴ㄛ褫蚚逄晟蹈桶笢腔恁隅逄晟蔚痄雄善巠蚚逄晟蹈桶笢﹝" & vbCrLf & vbCrLf & _
			"  - �垓覦篲�" & vbCrLf & _
			"    等僻蜆偌聽綴ㄛ褫蚚逄晟蹈桶笢腔垀衄逄晟蔚�垓諜げ秘褊弝譚擿埡訇縪苤�" & vbCrLf & vbCrLf & _
			"  - 刉壺" & vbCrLf & _
			"    猁刉壺巠蚚逄晟ㄛ恁寁巠蚚逄晟蹈桶笢腔逄晟ㄛ�遣騕本� [刉壺] 偌聽﹝" & vbCrLf & _
			"    等僻蜆偌聽綴ㄛ巠蚚逄晟蹈桶笢腔恁隅逄晟蔚痄雄善褫蚚逄晟蹈桶笢﹝" & vbCrLf & vbCrLf & _
			"  - �垓褫噫�" & vbCrLf & _
			"    等僻蜆偌聽綴ㄛ巠蚚逄晟蹈桶笢腔垀衄逄晟蔚�垓諜げ秘蝙孖譚擿埡訇縪苤�" & vbCrLf & vbCrLf & _
			"  - 崝樓褫蚚逄晟" & vbCrLf & _
			"    等僻蜆偌聽綴ㄛ蔚粟堤褫怀�邆擿埼�備睿測鎢勤趕遺ㄛ�毓那騣屏篲茧蝙孖譚擿埡訇縪苤�" & vbCrLf & vbCrLf & _
			"  - 晤憮褫蚚逄晟" & vbCrLf & _
			"    猁晤憮褫蚚逄晟ㄛ恁寁褫蚚逄晟蹈桶笢腔逄晟ㄛ�遣騕本� [晤憮褫蚚逄晟] 偌聽﹝" & vbCrLf & _
			"    等僻蜆偌聽綴ㄛ蔚粟堤褫晤憮逄晟靡備睿測鎢勤趕遺ㄛ�毓那騣峒瑏醴孖譚擿埡訇縪倳▲巡鼯擿唌�" & vbCrLf & vbCrLf & _
			"  - 刉壺褫蚚逄晟" & vbCrLf & _
			"    猁刉壺褫蚚逄晟ㄛ恁寁褫蚚逄晟蹈桶笢猁刉壺腔逄晟ㄛ�遣騕本� [刉壺褫蚚逄晟] 偌聽﹝" & vbCrLf & vbCrLf & _
			"  - 崝樓巠蚚逄晟" & vbCrLf & _
			"    等僻蜆偌聽綴ㄛ蔚粟堤褫怀�邆擿埼�備睿測鎢勤趕遺ㄛ�毓那騣屏篲茧褊弝譚擿埡訇縪苤�" & vbCrLf & vbCrLf & _
			"  - 晤憮巠蚚逄晟" & vbCrLf & _
			"    猁晤憮巠蚚逄晟ㄛ恁寁巠蚚逄晟蹈桶笢腔逄晟ㄛ�遣騕本� [晤憮巠蚚逄晟] 偌聽﹝" & vbCrLf & _
			"    等僻蜆偌聽綴ㄛ蔚粟堤褫晤憮逄晟靡備睿測鎢勤趕遺ㄛ�毓那騣峒瑏騫弝譚擿埡訇縪倳▲巡鼯擿唌�" & vbCrLf & vbCrLf & _
			"  - 刉壺巠蚚逄晟" & vbCrLf & _
			"    猁刉壺巠蚚逄晟ㄛ恁寁巠蚚逄晟蹈桶笢猁刉壺腔逄晟ㄛ�遣騕本� [刉壺巠蚚逄晟] 偌聽﹝" & vbCrLf & vbCrLf & _
			"  ♁蛁砩ㄩ氝樓﹜晤憮逄晟躺蚚衾 Passolo 帤懂唳掛陔崝腔盓厥逄晟﹝" & vbCrLf & _
			"  ﹛﹛﹛﹛逄晟測鎢③睿 Passolo 腔 ISO 396-1 測鎢悵厥珨祡ㄛ婦嬤湮苤迡﹝" & vbCrLf & vbCrLf & _
			"∵む坻髡夔∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 堆翑" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚粟堤絞ヶ敦諳腔堆翑陓洘﹝" & vbCrLf & vbCrLf & _
			"- 黍��" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚跦擂恁隅饜离腔祥肮粟堤狟蹈粕等:" & vbCrLf & _
			"  (1) 蘇�珋�" & vbCrLf & _
			"      黍�＿畏狣馺譆童炬Ａ埰戰硠馺藥硜楺苤�" & vbCrLf & _
			"      ♁躺絞恁隅腔饜离峈炵苀蘇�炵霰馺蟾悵炬欐埰噪簷佽央�" & vbCrLf & vbCrLf & _
			"  (2) 埻硉" & vbCrLf & _
			"      黍�×▲乳馺繭齟倚樂童炬Ａ埰戰硠馺藥硜楺苤�" & vbCrLf & _
			"      ♁躺絞恁隅饜离腔埻宎硉峈準諾奀ㄛ符珆尨蜆粕等﹝" & vbCrLf & vbCrLf & _
			"  (3) 統桽硉" & vbCrLf & _
			"      黍�×▲巡觸挍桲馺譆童炬Ａ埰戰硠馺藥硜楺苤�" & vbCrLf & _
			"      ♁蜆粕等珆尨壺恁隅饜离俋腔垀衄饜离蹈桶﹝" & vbCrLf & vbCrLf & _
			"- ь諾" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚ь諾珋衄饜离腔�垓謁童皆埸蔣蜣寪薹靿蹁馺譆窗�" & vbCrLf & vbCrLf & _
			"- 聆彸" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚粟堤聆彸勤趕遺ㄛ眕晞潰脤饜离腔淏�煩唌�" & vbCrLf & vbCrLf & _
			"- �毓�" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚悵湔饜离敦諳笢腔�庥庛�蜊ㄛ豖堤饜离敦諳甜殿隙翋敦諳﹝" & vbCrLf & _
			"  最唗蔚妏蚚載蜊綴腔饜离硉﹝" & vbCrLf & vbCrLf & _
			"- �＋�" & vbCrLf & _
			"  等僻蜆偌聽ㄛ祥悵湔饜离敦諳笢腔�庥庛�蜊ㄛ豖堤饜离敦諳甜殿隙翋敦諳﹝" & vbCrLf & _
			"  最唗蔚妏蚚埻懂腔饜离硉﹝" & vbCrLf & vbCrLf & vbCrLf
	TestUse ="∵饜离靡備∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 猁聆彸腔饜离靡備﹝猁恁寁饜离ㄛ等僻饜离蹈桶﹝" & vbCrLf & vbCrLf & _
			"∵楹祒蹈桶∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 蜆蹈桶蔚珆尨源偶笢腔垀衄楹祒蹈桶﹝③恁寁迵蠟腔赻隅砱饜离ぁ饜腔楹祒蹈桶輛俴聆彸﹝" & vbCrLf & vbCrLf & _
			"∵赻雄杸遙趼睫∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 婓潰脤奀赻雄杸遙趼揹笢睫磁饜离笢垀隅砱腔杸遙趼睫﹝" & vbCrLf & vbCrLf & _
			"∵黍�遶倇�∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 桶尨猁珆尨腔渣昫楹祒趼揹杅﹝膘祜祥猁怀�輲契騕齡童皆埼瑮稊硒捐炩鉌接�渾奀潔徹酗﹝" & vbCrLf & vbCrLf & _
			"∵婦漪囀�搳�" & vbCrLf & _
			"============" & vbCrLf & _
			"- 硌隅硐潰脤婦漪囀�搧儷硒恣�瞳蚚蜆砐褫眕衄渀勤俶腔聆彸ㄛ甜й樓辦聆彸奀潔﹝" & vbCrLf & _
			"- 蜆趼僇盓厥耀緇倰籵饜睫﹝瞰�蝤態*C 褫眕ぁ饜 XAYYCZ﹝" & vbCrLf & vbCrLf & _
			"♁蛁砩ㄩ蚕衾 Sax Basic 竘э腔恀枙ㄛ媼跺準荎恅趼譫眳潔腔 ? 籵饜睫祥掩盓厥﹝" & vbCrLf & _
			"﹛﹛﹛﹛瞰�蝤滿偌翾�??恅璃§祥ぁ饜※湖羲蚚誧恅璃§﹝" & vbCrLf & vbCrLf & _
			"∵趼揹囀�搳�" & vbCrLf & _
			"============" & vbCrLf & _
			"枑鼎賸�垓縑Ⅶ儠敯�﹜笝砦睫﹜樓厒ん 4 跺恁砐﹝" & vbCrLf & vbCrLf & _
			"- �蝜�恁寁�垓縛皈藫頖�等砐蔚掩赻雄�＋�恁寁﹝" & vbCrLf & _
			"- �蝜�恁寁等砐ㄛ寀�垓謀＋蹐垮閤堈紙＋�恁寁﹝" & vbCrLf & _
			"- 等砐褫眕嗣恁﹝" & vbCrLf & vbCrLf & _
			"∵む坻髡夔∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 堆翑" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚粟堤絞ヶ敦諳腔堆翑陓洘﹝" & vbCrLf & vbCrLf & _
			"- 聆彸" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚偌桽恁隅腔沭璃輛俴聆彸﹝" & vbCrLf & vbCrLf & _
			"- ь諾" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚ь諾珋衄腔聆彸賦彆﹝" & vbCrLf & vbCrLf & _
			"- �＋�" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚豖堤聆彸最唗甜殿隙饜离敦諳﹝" & vbCrLf & vbCrLf & vbCrLf
	CopyRight = "∵唳�享驨驉�" & vbCrLf & _
			"============" & vbCrLf & _
			"- 森�篲�腔唳�邦橦疝Ｇ葴迖瑏齣葂齾苺畏庥恛佪奿埼漞捗墓獺Ｂ瑏纂９棉ヾ〧３摹寰篲�﹝" & vbCrLf & _
			"- 党蜊﹜汃票掛�篲�斛剕呴蜇掛佽隴恅璃ㄛ甜蛁隴�篲�埻宎羲楷氪眕摯党蜊氪﹝" & vbCrLf & _
			"- 帤冪羲楷氪睿党蜊氪肮砩ㄛ�庥拵橠粉繨鶲芄炬輓譚譚硭朊脹篲�﹜妀珛麼岆む坳茠瞳俶魂雄﹝" & vbCrLf & _
			"- 勤妏蚚掛�篲�腔埻宎唳掛ㄛ眕摯妏蚚冪坻�刵瑏警譟л倚摯瘙憲齉麭伂騰蟣宋豖虩忙玷疝Ｇ葀�" & vbCrLf & _
			"  創童�庥拏蟭峞�" & vbCrLf & _
			"- 蚕衾峈轎煤�篲�ㄛ羲楷氪睿党蜊氪羶衄砱昢枑鼎�篲�撮扲盓厥ㄛ珩拸砱昢蜊輛麼載陔唳掛﹝" & vbCrLf & _
			"- 辣茩硌堤渣昫甜枑堤蜊輛砩獗﹝�觬迡簊騠羷例憌甭賰〩芚�: z_shangyi@163.com﹝" & vbCrLf & vbCrLf & vbCrLf
	Thank = "∵祡﹛﹛郅∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 掛�篲�婓党蜊徹最笢腕善犖趙陔岍槨頗埜腔聆彸ㄛ婓森桶尨笪陑腔覜郅ㄐ" & vbCrLf & _
			"- 覜郅怢俜 Heaven 珂汜枑堤楛极蚚逄党蜊砩獗ㄐ" & vbCrLf & vbCrLf & vbCrLf
	Contact = "∵迵扂薊炵∵" & vbCrLf & _
			"============" & vbCrLf & _
			"wanfuㄩz_shangyi@163.com" & vbCrLf & vbCrLf & vbCrLf
	Logs = "覜郅盓厥ㄐ蠟腔盓厥岆扂郔湮腔雄薯ㄐ肮奀辣茩妏蚚扂蠅秶釬腔�篲�ㄐ" & vbCrLf & vbCrLf & _
			"**************************************************" & vbCrLf & _
			"剒猁載嗣﹜載陔﹜載疑腔犖趙ㄛ③溼恀:" & vbCrLf & _
			"犖趙陔岍槨 -- http://www.hanzify.org" & vbCrLf & _
			"犖趙陔岍槨蹦抭 -- http://bbs.hanzify.org" & vbCrLf & _
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


'赻雄載陔堆翑
Sub UpdateHelp(HelpTip As String)
	If OSLanguage = "0404" Then
	HelpTitle = "說明"
	HelpTipTitle = "自動更新"
	SetWindows = " 設定視窗 "
	Lines = "-----------------------"
	SetUse ="☆更新方式☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 自動下載更新並安裝" & vbCrLf & _
			"  選擇該選項時，程式將根據更新頻率中的設定自動檢查更新，如果偵測到有新的版本可用時，" & vbCrLf & _
			"  將在不徵求使用者意見的情況下自動下載更新並安裝。更新完畢後將彈出對話方塊通知使用者，按" & vbCrLf & _
			"  [確定] 按鈕後程式即結束。" & vbCrLf & vbCrLf & _
			"- 有更新時通知我，由我決定下載並安裝" & vbCrLf & _
			"  選擇該選項時，程式將根據更新頻率中的設定自動檢查更新，如果偵測到有新的版本可用時，" & vbCrLf & _
			"  將彈出對話方塊提示使用者，如果使用者決定更新，程式將下載更新並安裝。更新完畢後將彈出對" & vbCrLf & _
			"  話方塊通知使用者，按 [確定] 按鈕後程式即結束。" & vbCrLf & vbCrLf & _
			"- 關閉自動更新" & vbCrLf & _
			"  選擇該選項時，程式將不檢查更新。" & vbCrLf & vbCrLf & _
			"◎注意：無論何種更新方式，更新成功並結束程式後都需要使用者重新啟動巨集。" & vbCrLf & vbCrLf & _
			"☆更新頻率☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 檢查間隔" & vbCrLf & _
			"  檢查更新的時間週期。程式在檢查週期內只檢查一次。" & vbCrLf & vbCrLf & _
			"- 最後檢查日期" & vbCrLf & _
			"  最後檢查時的日期。程式在檢查更新時自動記錄檢查日期，並在同一日期內只檢查一次。" & vbCrLf & vbCrLf & _
			"- 檢查" & vbCrLf & _
			"  點擊該按鈕後，程式將忽略檢查間隔和檢查日期進行更新檢查，並更新檢查日期。" & vbCrLf & vbCrLf & _
			"  ◎如果偵測到有新的版本可用，將彈出對話方塊提示使用者，如果使用者決定更新，程式將下載更新" & vbCrLf & _
			"  　並安裝。更新完畢後將彈出對話方塊通知使用者，按 [確定] 按鈕後程式即結束。" & vbCrLf & _
			"  ◎如果沒有偵測到有新的版本可用，將彈出目前網上的版本號。並提示是否重新下載更新。" & vbCrLf & vbCrLf & _
			"☆更新網址清單☆" & vbCrLf & _
			"============" & vbCrLf & _
			"此處的更新網址使用者可以自己定義。使用者定義的網址將被優先使用。" & vbCrLf & _
			"使用者定義的網址無法存取時，將使用程式開發者定義的更新網址。" & vbCrLf & vbCrLf & _
			"☆RAR 解壓程式☆" & vbCrLf & _
			"================" & vbCrLf & _
			"由於程式包採用了 RAR 格式壓縮，故需要在本機上安裝有對應的解壓程式。" & vbCrLf & _
			"初次使用時，程式將自動搜索註冊表中註冊的 RAR 副檔名預設解壓程式，並進行適當的設定。" & vbCrLf & _
			"並在每次使用時自動檢查解壓程式是否還存在，如果不存在將重新搜索並設定。" & vbCrLf & vbCrLf & _
			"◎注意：程式預設支援的解壓程式為：WinRAR、WinZIP、7z。如果機器中沒有這些解壓程式，" & vbCrLf & _
			"　　　　需要手動設定。" & vbCrLf & vbCrLf & _
			"- 程式路徑" & vbCrLf & _
			"  解壓程式的完整路徑。可點擊右邊的 [...] 按鈕手工新增。" & vbCrLf & vbCrLf & _
			"- 解壓參數" & vbCrLf & _
			"  解壓程式解壓縮 RAR 檔案時的命令列參數。其中：" & vbCrLf & _
			"  %1 為壓縮檔案，%2 為要從壓縮包中擷取的主程式檔案，%3 為解壓後的檔案路徑。" & vbCrLf & _
			"  這些參數為必要參數，不可缺少並且不可用其他符號代替。至於先後順序，依照解壓程式的" & vbCrLf & _
			"  命令列規則。" & vbCrLf & _
			"  可點擊右邊的 [>] 按鈕手工新增這些必要參數。" & vbCrLf & vbCrLf & _
			"☆其他功能☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 說明" & vbCrLf & _
			"  點擊該按鈕，將獲取說明訊息。" & vbCrLf & vbCrLf & _
			"- 讀取" & vbCrLf & _
			"  點擊該按鈕，將根據選擇設定的不同彈出下列選單:" & vbCrLf & _
			"  (1) 預設值" & vbCrLf & _
			"      讀取自動更新設定的預設值，並顯示在更新網址清單和 RAR 解壓程式中。" & vbCrLf & _
			"  (2) 原值" & vbCrLf & _
			"      讀取自動更新設定的原始值，並顯示在更新網址清單和 RAR 解壓程式中。" & vbCrLf & vbCrLf & _
			"- 清空" & vbCrLf & _
			"  將更新網址清單和RAR 解壓程式中的內容全部清空。" & vbCrLf & vbCrLf & _
			"- 測試" & vbCrLf & _
			"  點擊該按鈕，將測試更新網址清單和 RAR 解壓程式是否正確。" & vbCrLf & vbCrLf & _
			"- 確定" & vbCrLf & _
			"  點擊該按鈕，將儲存設定視窗中的任何變更，結束設定視窗並返回主視窗。" & vbCrLf & _
			"  程式將使用變更後的設定值。" & vbCrLf & vbCrLf & _
			"- 取消" & vbCrLf & _
			"  點擊該按鈕，不儲存設定視窗中的任何變更，結束設定視窗並返回主視窗。" & vbCrLf & _
			"  程式將使用原來的設定值。" & vbCrLf & vbCrLf & vbCrLf
	Logs = "感謝支持！您的支持是我最大的動力！同時歡迎使用我們製作的軟體！" & vbCrLf & vbCrLf & _
			"**************************************************" & vbCrLf & _
			"需要更多、更新、更好的漢化，請拜訪:" & vbCrLf & _
			"漢化新世紀 -- http://www.hanzify.org" & vbCrLf & _
			"漢化新世紀論壇 -- http://bbs.hanzify.org" & vbCrLf & _
			"**************************************************" & vbCrLf
	Else
	HelpTitle = "堆翑"
	HelpTipTitle = "赻雄載陔"
	SetWindows = " 饜离敦諳 "
	Lines = "-----------------------"
	SetUse ="∵載陔源宒∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 赻雄狟婥載陔甜假蚾" & vbCrLf & _
			"  恁隅蜆恁砐奀ㄛ最唗蔚跦擂載陔け薹笢腔扢离赻雄潰脤載陔ㄛ�蝜�潰聆善衄陔腔唳掛褫蚚奀ㄛ" & vbCrLf & _
			"  蔚婓祥涽⑴蚚誧砩獗腔①錶狟赻雄狟婥載陔甜假蚾﹝載陔俇救綴蔚粟堤勤趕遺籵眭蚚誧ㄛ偌" & vbCrLf & _
			"  [�毓沘 偌聽綴最唗撈豖堤﹝" & vbCrLf & vbCrLf & _
			"- 衄載陔奀籵眭扂ㄛ蚕扂樵隅狟婥甜假蚾" & vbCrLf & _
			"  恁隅蜆恁砐奀ㄛ最唗蔚跦擂載陔け薹笢腔扢离赻雄潰脤載陔ㄛ�蝜�潰聆善衄陔腔唳掛褫蚚奀ㄛ" & vbCrLf & _
			"  蔚粟堤勤趕遺枑尨蚚誧ㄛ�蝜�蚚誧樵隅載陔ㄛ最唗蔚狟婥載陔甜假蚾﹝載陔俇救綴蔚粟堤勤" & vbCrLf & _
			"  趕遺籵眭蚚誧ㄛ偌 [�毓沘 偌聽綴最唗撈豖堤﹝" & vbCrLf & vbCrLf & _
			"- 壽敕赻雄載陔" & vbCrLf & _
			"  恁隅蜆恁砐奀ㄛ最唗蔚祥潰脤載陔﹝" & vbCrLf & vbCrLf & _
			"♁蛁砩ㄩ拸蹦睡笱載陔源宒ㄛ載陔傖髡甜豖堤最唗綴飲剒猁蚚誧笭陔ゐ雄粽﹝" & vbCrLf & vbCrLf & _
			"∵載陔け薹∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 潰脤潔路" & vbCrLf & _
			"  潰脤載陔腔奀潔笚ぶ﹝最唗婓潰脤笚ぶ囀硐潰脤珨棒﹝" & vbCrLf & vbCrLf & _
			"- 郔綴潰脤�梪�" & vbCrLf & _
			"  郔綴潰脤奀腔�梪琚ㄢ昐藜睄麮曏�陔奀赻雄暮翹潰脤�梪琭炬Ｆ硰皆銀梪矬稂遞麮橏輕峞�" & vbCrLf & vbCrLf & _
			"- 潰脤" & vbCrLf & _
			"  等僻蜆偌聽綴ㄛ最唗蔚綺謹潰脤潔路睿潰脤�梪睍靇邽�陔潰脤ㄛ甜載陔潰脤�梪琚�" & vbCrLf & vbCrLf & _
			"  ♁�蝜�潰聆善衄陔腔唳掛褫蚚ㄛ蔚粟堤勤趕遺枑尨蚚誧ㄛ�蝜�蚚誧樵隅載陔ㄛ最唗蔚狟婥載陔" & vbCrLf & _
			"  ﹛甜假蚾﹝載陔俇救綴蔚粟堤勤趕遺籵眭蚚誧ㄛ偌 [�毓沘 偌聽綴最唗撈豖堤﹝" & vbCrLf & _
			"  ♁�蝜�羶衄潰聆善衄陔腔唳掛褫蚚ㄛ蔚粟堤絞ヶ厙奻腔唳掛瘍﹝甜枑尨岆瘁笭陔狟婥載陔﹝" & vbCrLf & vbCrLf & _
			"∵載陔厙硊蹈桶∵" & vbCrLf & _
			"============" & vbCrLf & _
			"森揭腔載陔厙硊蚚誧褫眕赻撩隅砱﹝蚚誧隅砱腔厙硊蔚掩蚥珂妏蚚﹝" & vbCrLf & _
			"蚚誧隅砱腔厙硊拸楊溼恀奀ㄛ蔚妏蚚最唗羲楷氪隅砱腔載陔厙硊﹝" & vbCrLf & vbCrLf & _
			"∵RAR 賤揤最唗∵" & vbCrLf & _
			"================" & vbCrLf & _
			"蚕衾最唗婦粒蚚賸 RAR 跡宒揤坫ㄛ嘟剒猁婓掛儂奻假蚾衄眈茼腔賤揤最唗﹝" & vbCrLf & _
			"場棒妏蚚奀ㄛ最唗蔚赻雄刲坰蛁聊桶笢蛁聊腔 RAR 孺桯靡蘇�牮瑵像昐礗炬Ⅴ靇倇妗接霰馺獺�" & vbCrLf & _
			"甜婓藩棒妏蚚奀赻雄潰脤賤揤最唗岆瘁遜湔婓ㄛ�蝜�祥湔婓蔚笭陔刲坰甜饜离﹝" & vbCrLf & vbCrLf & _
			"♁蛁砩ㄩ最唗蘇�珋妊硉躅瑵像昐藬炒斡inRAR﹜WinZIP﹜7z﹝�蝜�儂ん笢羶衄涴虳賤揤最唗ㄛ" & vbCrLf & _
			"﹛﹛﹛﹛剒猁忒雄饜离﹝" & vbCrLf & vbCrLf & _
			"- 最唗繚噤" & vbCrLf & _
			"  賤揤最唗腔俇淕繚噤﹝褫等僻衵晚腔 [...] 偌聽忒馱氝樓﹝" & vbCrLf & vbCrLf & _
			"- 賤揤統杅" & vbCrLf & _
			"  賤揤最唗賤揤坫 RAR 恅璃奀腔韜鍔俴統杅﹝む笢ㄩ" & vbCrLf & _
			"  %1 峈揤坫恅璃ㄛ%2 峈猁植揤坫婦笢枑�△齡鰴昐藬躁�ㄛ%3 峈賤揤綴腔恅璃繚噤﹝" & vbCrLf & _
			"  涴虳統杅峈斛猁統杅ㄛ祥褫�捻椏〥珩遛孖蟻頖�睫瘍測杸﹝祫衾珂綴佼唗ㄛ甡桽賤揤最唗腔" & vbCrLf & _
			"  韜鍔俴寞寀﹝" & vbCrLf & _
			"  褫等僻衵晚腔 [>] 偌聽忒馱氝樓涴虳斛猁統杅﹝" & vbCrLf & vbCrLf & _
			"∵む坻髡夔∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 堆翑" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚鳳�※攃�陓洘﹝" & vbCrLf & vbCrLf & _
			"- 黍��" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚跦擂恁隅饜离腔祥肮粟堤狟蹈粕等:" & vbCrLf & _
			"  (1) 蘇�珋�" & vbCrLf & _
			"      黍�＝堈站�陔饜离腔蘇�珋童炬Ａ埰戰皒�陔厙硊蹈桶睿 RAR 賤揤最唗笢﹝" & vbCrLf & _
			"  (2) 埻硉" & vbCrLf & _
			"      黍�＝堈站�陔饜离腔埻宎硉ㄛ甜珆尨婓載陔厙硊蹈桶睿 RAR 賤揤最唗笢﹝" & vbCrLf & vbCrLf & _
			"- ь諾" & vbCrLf & _
			"  蔚載陔厙硊蹈桶睿RAR 賤揤最唗笢腔囀�朠垓褲敹捸�" & vbCrLf & vbCrLf & _
			"- 聆彸" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚聆彸載陔厙硊蹈桶睿 RAR 賤揤最唗岆瘁淏�楚�" & vbCrLf & vbCrLf & _
			"- �毓�" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚悵湔饜离敦諳笢腔�庥庛�蜊ㄛ豖堤饜离敦諳甜殿隙翋敦諳﹝" & vbCrLf & _
			"  最唗蔚妏蚚載蜊綴腔饜离硉﹝" & vbCrLf & vbCrLf & _
			"- �＋�" & vbCrLf & _
			"  等僻蜆偌聽ㄛ祥悵湔饜离敦諳笢腔�庥庛�蜊ㄛ豖堤饜离敦諳甜殿隙翋敦諳﹝" & vbCrLf & _
			"  最唗蔚妏蚚埻懂腔饜离硉﹝" & vbCrLf & vbCrLf & vbCrLf
	Logs = "覜郅盓厥ㄐ蠟腔盓厥岆扂郔湮腔雄薯ㄐ肮奀辣茩妏蚚扂蠅秶釬腔�篲�ㄐ" & vbCrLf & vbCrLf & _
			"**************************************************" & vbCrLf & _
			"剒猁載嗣﹜載陔﹜載疑腔犖趙ㄛ③溼恀:" & vbCrLf & _
			"犖趙陔岍槨 -- http://www.hanzify.org" & vbCrLf & _
			"犖趙陔岍槨蹦抭 -- http://bbs.hanzify.org" & vbCrLf & _
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
