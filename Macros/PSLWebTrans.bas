''This macro is a online translations macro program for strings
''in Passolo translation list.
''It has the following features:
''- Use the online translation engine automatically translate strings
''  in the Passolo translation list
''- Integrated some of the well-known online translation engines, and
''  you can customize other online translation engines
''- You can choose the string type, skiping some of string, and processing
''  the strings before and after translation
''- Integrated shortcuts, terminators, Accelerator check macro, and you can
''  check and correct errors in translations after the strings has be translated
''Idea and implemented by wanfu 2010.05.12 (modified on 2010.11.11)

Public trn As PslTransList,TransString As PslTransString,OSLanguage As String

Public SpaceTrn As String,acckeyTrn As String,ExpStringTrn As String,EndStringTrn As String
Public acckeySrc As String,EndStringSrc As String,ShortcutSrc As String,ShortcutTrn As String
Public PreStringTrn As String,EndSpaceSrc As String,EndSpaceTrn As String
Public AllCont As Integer,AccKey As Integer,EndChar As Integer,Acceler As Integer
Public srcLineNum As Integer,trnLineNum As Integer,srcAccKeyNum As Integer,trnAccKeyNum As Integer

Public DefaultCheckList() As String,AppRepStr As String,PreRepStr As String
Public cWriteLoc As String,cSelected() As String,cUpdateSet() As String
Public CheckList() As String,CheckListBak() As String,CheckDataList() As String,CheckDataListBak() As String
Public tempCheckList() As String,tempCheckDataList() As String

Public DefaultEngineList() As String,WaitTimes As Long
Public tWriteLoc As String,tSelected() As String,tUpdateSet() As String,tUpdateSetBak() As String
Public EngineList() As String,EngineListBak() As String,EngineDataList() As String,EngineDataListBak() As String
Public DelLngNameList() As String,DelSrcLngList() As String,DelTranLngList() As String

Public FileNo As Integer,FileText As String,FindText As String,FindLine As String
Public AppNames() As String,AppPaths() As String,FileList() As String,CodeList() As String

Private Const Version = "2010.11.11"
Private Const ToUpdateEngineVersion = "2010.05.24"
Private Const ToUpdateCheckVersion = "2010.09.25"
Private Const EngineRegKey = "HKCU\Software\VB and VBA Program Settings\WebTranslate\"
Private Const EngineFilePath = MacroDir & "\Data\PSLWebTrans.dat"
Private Const CheckRegKey = "HKCU\Software\VB and VBA Program Settings\AccessKey\"
Private Const CheckFilePath = MacroDir & "\Data\PSLCheckAccessKeys.dat"
Private Const JoinStr = vbBack
Private Const SubJoinStr = Chr$(1)
Private Const rSubJoinStr = Chr$(19) & Chr$(20)
Private Const LngJoinStr = "|"
Private Const NullValue = "Null"

Private Const DefaultObject = "Microsoft.XMLHTTP"
Private Const updateAppName = "PSLWebTrans"
Private Const updateMainFile = "PSLWebTrans.bas"
Private Const updateINIFile = "PSLMacrosUpdates.ini"
Private Const updateMethod = "GET"
Private Const updateINIMainUrl = "ftp://hhdown:0011@czftp.hanzify.org/update/PSLMacrosUpdates.ini"
Private Const updateINIMinorUrl = "http://www.wanfutrade.com/software/hanhua/PSLMacrosUpdates.ini"
Private Const updateMainUrl = "ftp://hhdown:0011@czftp.hanzify.org/download/PSLWebTrans.rar"
Private Const updateMinorUrl = "ftp://hhdown:0011@222.76.212.240:121/downloads/PSLWebTrans.rar"
Private Const updateAsync = "False"


'楹祒竘э蘇�珈髲�
Function EngineSettings(DataName As String) As String
	Dim ObjectName As String,Template As String,AppId As String,openUrl As String,openMethod As String
	Dim openAsync As String,openUser As String,openPassword As String,sendBody As String
	Dim setRequestHeader As String,responseType As String,TranBeforeStr As String,TranAfterStr As String

	If DataName = DefaultEngineList(0) Then
		ObjectName = DefaultObject
		AppId = "fefed727-bbc1-4421-828d-fc828b24d59b"
		openUrl = "http://api.microsofttranslator.com/V2/Http.svc/Translate?"
		Template = "{Url}&appId={appId}&text={text}&from={from}&to={to}"
		openMethod = "GET"
		openAsync = "False"
		openUser = ""
		openPassword = ""
		sendBody = ""
		setRequestHeader = "Content-Type,application/xml; charset=utf-8"
		responseType = "responseText"
		TranBeforeStr = "Serialization/"">"
		TranAfterStr = "</string>"
	ElseIf DataName = DefaultEngineList(1) Then
		ObjectName = DefaultObject
		AppId = ""
		openUrl = "http://translate.google.com/translate_t?"
		Template = "{Url}&text={text}&langpair={from}|{to}"
		openMethod = "POST"
		openAsync = "False"
		openUser = ""
		openPassword = ""
		sendBody = ""
		setRequestHeader = "Content-Type,text/html; charset=utf-8"
		responseType = "responseText"
		TranBeforeStr = "onmouseout=""this.style.backgroundColor='#fff'"">"
		TranAfterStr = "</span>"
	ElseIf DataName = DefaultEngineList(2) Then
		ObjectName = DefaultObject
		AppId = ""
		openUrl = "http://fanyi.yahoo.com.cn/translate_txt?"
		Template = "{Url}&ei=UTF-8&fr=&lp={from}_{to}&trtext={Text}"
		openMethod = "POST"
		openAsync = "False"
		openUser = ""
		openPassword = ""
		sendBody = ""
		setRequestHeader = "Content-Type,text/html; charset=utf-8"
		responseType = "responseText"
		TranBeforeStr = "<div id=""pd"" class=""pd"">"
		TranAfterStr = "</div>"
	End If
	EngineSettings = ObjectName & SubJoinStr & AppId & SubJoinStr & openUrl & SubJoinStr & _
					Template & SubJoinStr & openMethod & SubJoinStr & openAsync & SubJoinStr & _
					openUser & SubJoinStr & openPassword & SubJoinStr & sendBody & SubJoinStr & _
					setRequestHeader & SubJoinStr & responseType & SubJoinStr & _
					TranBeforeStr & SubJoinStr & TranAfterStr
End Function


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
	Dim nSelected As String,EngineID As Integer,CheckID As Integer,TranLang As String,xmlHttp As Object

	If OSLanguage = "0404" Then
		Msg01 = "錯誤"
		Msg02 = "請選取要翻譯的字串！"
		Msg03 = "無法儲存！請檢查是否有寫入下列位置的權限:" & vbCrLf & vbCrLf
		Msg04 = "確認"
		Msg05 = "您的系統缺少 Microsoft.XMLHTTP 物件，無法繼續執行！"
		Msg06 = "翻譯引擎伺服器響應逾時！需要延長等待時間嗎？"
		Msg07 = "等待時間:"
		Msg08 = "請輸入等待時間"
		Msg09 = "無法與翻譯引擎伺服器通信！可能是無 Internet 連接，" & vbCrLf & _
				"或者翻譯引擎的設定錯誤，或者翻譯引擎禁止存取。"
		Msg10 = "翻譯引擎網址為空，無法繼續！"
	Else
		Msg01 = "渣昫"
		Msg02 = "③恁寁猁楹祒腔趼揹ㄐ"
		Msg03 = "拸楊悵湔ㄐ③潰脤岆瘁衄迡�輴臏倛閥繭饑使�:" & vbCrLf & vbCrLf
		Msg04 = "�溜�"
		Msg05 = "蠟腔炵苀�捻� Microsoft.XMLHTTP 勤砓ㄛ拸楊樟哿堍俴ㄐ"
		Msg06 = "楹祒竘э督昢ん砒茼閉奀ㄐ剒猁晊酗脹渾奀潔鎘ˋ"
		Msg07 = "脹渾奀潔:"
		Msg08 = "③怀�賮�渾奀潔"
		Msg09 = "拸楊迵楹祒竘э督昢ん籵陓ㄐ褫夔岆拸 Internet 蟀諉ㄛ" & vbCrLf & _
				"麼氪楹祒竘э腔扢离渣昫ㄛ麼氪楹祒竘э輦砦溼恀﹝"
		Msg10 = "楹祒竘э厙硊峈諾ㄛ拸楊樟哿ㄐ"
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
		If Join(tSelected) <> "" Then
			EngineSet = tSelected(0)
			CheckSet = tSelected(1)
			mAllType = StrToInteger(tSelected(2))
			mMenu = StrToInteger(tSelected(3))
			mDialog = StrToInteger(tSelected(4))
			mString = StrToInteger(tSelected(5))
			mAccTable = StrToInteger(tSelected(6))
			mVer = StrToInteger(tSelected(7))
			mOther = StrToInteger(tSelected(8))
			mSelOnly = StrToInteger(tSelected(9))
			mForReview = StrToInteger(tSelected(10))
			mValidated = StrToInteger(tSelected(11))
			mNotTran = StrToInteger(tSelected(12))
			mNumAndSymbol = StrToInteger(tSelected(13))
			mAllUCase = StrToInteger(tSelected(14))
			mAllLCase = StrToInteger(tSelected(15))
			mAutoSele = StrToInteger(tSelected(16))
			mAccKey = StrToInteger(tSelected(17))
			mAccelerator = StrToInteger(tSelected(18))
			mPreStrRep = StrToInteger(tSelected(19))
			mSplitTran = StrToInteger(tSelected(20))
			mCheckTrn = StrToInteger(tSelected(21))
			mAppStrRep = StrToInteger(tSelected(22))
			KeepSet = StrToInteger(tSelected(23))
			ShowMsg = StrToInteger(tSelected(24))
			TranComm = StrToInteger(tSelected(25))
		End If

		DlgText "EngineList",EngineSet
		DlgText "CheckList",CheckSet
		If DlgText("EngineList") = "" Then DlgValue "EngineList",0
		If DlgText("CheckList") = "" Then DlgValue "CheckList",0
		DlgValue "AllType",mAllType
		DlgValue "Menu",mMenu
		DlgValue "Dialog",mDialog
		DlgValue "Strings",mString
		DlgValue "AccTable",mAccTable
		DlgValue "Versions",mVer
		DlgValue "Other",mOther
		DlgValue "Seleted",mSelOnly
		DlgValue "ForReview",mForReview
		DlgValue "Validated",mValidated
		DlgValue "NotTran",mNotTran
		DlgValue "NumAndSymbol",mNumAndSymbol
		DlgValue "AllUCase",mAllUCase
		DlgValue "AllLCase",mAllLCase
		DlgValue "AutoSelection",mAutoSele
		DlgValue "AccessKey",mAccKey
		DlgValue "Accelerator",mAccelerator
		DlgValue "PreStrRep",mPreStrRep
		DlgValue "SplitTran",mSplitTran
		DlgValue "CheckTrn",mCheckTrn
		DlgValue "AppStrRep",mAppStrRep
		DlgValue "KeepSet",KeepSet
		DlgValue "ShowMsg",ShowMsg
		DlgValue "TranComment",TranComm
		If mAllType + mMenu + mDialog + mString + mAccTable + mVer + mOther + mSelOnly = 0 Then
			DlgValue "AllType",1
		End If
		If trn.IsOpen = False Then
			DlgEnable "Seleted",False
			DlgValue "Seleted",0
		End If
		If mAutoSele = 1 Then
			DlgValue "CheckList",getCheckID(CheckDataList,trnLng,TranLang)
			DlgEnable "CheckList",False
		End If
		EngineSet = DlgText("EngineList")
		CheckSet = DlgText("CheckList")
		If CheckNullData(EngineSet,EngineDataList,"1,5,6,7,8",1) = True Then DlgEnable "OKButton",False
		If CheckNullData(CheckSet,CheckDataList,"4",0) = True Then
			DlgValue "AccessKey",0
			DlgValue "Accelerator",0
			DlgValue "PreStrRep",0
			DlgValue "CheckTrn",0
			DlgValue "AppStrRep",0
			DlgEnable "AccessKey",False
			DlgEnable "Accelerator",False
			DlgEnable "PreStrRep",False
			DlgEnable "CheckTrn",False
			DlgEnable "AppStrRep",False
		End If
		DlgEnable "SaveButton",False
	Case 2 ' 杅硉載蜊麼氪偌狟賸偌聽
		EngineID = DlgValue("EngineList")
		EngineSet = DlgText("EngineList")
		CheckSet = DlgText("CheckList")
		mAllType = DlgValue("AllType")
		mMenu = DlgValue("Menu")
		mDialog = DlgValue("Dialog")
		mString =  DlgValue("Strings")
		mAccTable = DlgValue("AccTable")
		mVer = DlgValue("Versions")
		mOther = DlgValue("Other")
		mSelOnly = DlgValue("Seleted")
		mForReview = DlgValue("ForReview")
		mValidated = DlgValue("Validated")
		mNotTran = DlgValue("NotTran")
		mNumAndSymbol = DlgValue("NumAndSymbol")
		mAllUCase = DlgValue("AllUCase")
		mAllLCase = DlgValue("AllLCase")
		mAutoSele = DlgValue("AutoSelection")
		mAccKey = DlgValue("AccessKey")
		mAccelerator = DlgValue("Accelerator")
		mPreStrRep = DlgValue("PreStrRep")
		mSplitTran = DlgValue("SplitTran")
		mCheckTrn = DlgValue("CheckTrn")
		mAppStrRep = DlgValue("AppStrRep")
		KeepSet = DlgValue("KeepSet")
		ShowMsg = DlgValue("ShowMsg")
		TranComm = DlgValue("TranComment")
		nSelected = EngineSet & JoinStr & CheckSet & JoinStr & mAllType & JoinStr & mMenu & JoinStr & _
					mDialog & JoinStr & mString & JoinStr & mAccTable & JoinStr & mVer & JoinStr & _
					mOther & JoinStr & mSelOnly & JoinStr & mForReview & JoinStr & mValidated & _
					JoinStr & mNotTran & JoinStr & mNumAndSymbol & JoinStr & mAllUCase & JoinStr & _
					mAllLCase & JoinStr & mAutoSele & JoinStr & mAccKey & JoinStr & mAccelerator & _
					JoinStr & mPreStrRep & JoinStr & mSplitTran & JoinStr & mCheckTrn & JoinStr & _
					mAppStrRep & JoinStr & KeepSet & JoinStr & ShowMsg & JoinStr & TranComm

		'潰聆趼揹濬倰睿囀�楪√鯓Й鮹畎敗▲亞垓蕩芢頖�等砐
		If DlgItem$ = "Menu" Or DlgItem$ = "Dialog" Or DlgItem$ = "Strings" Or	DlgItem$ = "AccTable" Or _
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
		If DlgItem$ = "SetButton" Then
			EngineListBak = EngineList
			EngineDataListBak = EngineDataList
			CheckListBak = CheckList
			CheckDataListBak = CheckDataList
			tUpdateSetBak = tUpdateSet
			EngineID = DlgValue("EngineList")
			CheckID = DlgValue("CheckList")
			Call Settings(EngineID,CheckID)
			DlgListBoxArray "EngineList",EngineList()
			DlgListBoxArray "CheckList",CheckList()
			DlgValue "EngineList",EngineID
			If DlgValue("AutoSelection") = 1 Then
				DlgValue "CheckList",getCheckID(CheckDataList,trnLng,TranLang)
			Else
				DlgValue "CheckList",CheckID
			End If
			If DlgText("EngineList") = "" Then DlgValue "EngineList",0
			If DlgText("CheckList") = "" Then DlgValue "CheckList",0
		End If
		If DlgItem$ = "AutoSelection" Then
			If DlgValue("AutoSelection") = 1 Then
				DlgValue "CheckList",getCheckID(CheckDataList,trnLng,TranLang)
				DlgEnable "CheckList",False
			Else
				If tSelected(1) = "" Then CheckSet = CheckList(0) Else CheckSet = tSelected(1)
				DlgText "CheckList",CheckSet
				DlgEnable "CheckList",True
			End If
		End If
		If CheckNullData(EngineSet,EngineDataList,"1,5,6,7,8",1) = True Then
			DlgEnable "OKButton",False
		Else
			DlgEnable "OKButton",True
		End If
		If CheckNullData(CheckSet,CheckDataList,"4",0) = True Then
			DlgValue "AccessKey",0
			DlgValue "Accelerator",0
			DlgValue "PreStrRep",0
			DlgValue "CheckTrn",0
			DlgValue "AppStrRep",0
			DlgEnable "AccessKey",False
			DlgEnable "Accelerator",False
			DlgEnable "PreStrRep",False
			DlgEnable "CheckTrn",False
			DlgEnable "AppStrRep",False
		Else
			DlgEnable "AccessKey",True
			DlgEnable "Accelerator",True
			DlgEnable "PreStrRep",True
			DlgEnable "CheckTrn",True
			DlgEnable "AppStrRep",True
		End If
		If DlgItem$ = "OKButton" Then
			'潰聆 Microsoft.XMLHTTP 岆瘁湔婓
			Set xmlHttp = CreateObject(DefaultObject)
			If xmlHttp Is Nothing Then Set xmlHttp = CreateObject("Msxml2.XMLHTTP")
			If xmlHttp Is Nothing Then
				MsgBox(Msg05,vbOkOnly+vbInformation,Msg01)
				MainDlgFunc% = True ' 滅砦偌狟偌聽壽敕勤趕遺敦諳
				Exit Function
			End If
			'鳳�〃獃埸倡�
			trnString = getTranslate(EngineID,xmlHttp,"Test","",3)
			Set xmlHttp = Nothing
			'聆彸 Internet 蟀諉
			If trnString = "NotConnected" Then
				MsgBox(Msg09,vbOkOnly+vbInformation,Msg01)
				MainDlgFunc% = True ' 滅砦偌狟偌聽壽敕勤趕遺敦諳
				Exit Function
			End If
			'聆彸竘э厙硊岆瘁峈諾
			If trnString = "NullUrl" Then
				MsgBox(Msg10,vbOkOnly+vbInformation,Msg01)
				MainDlgFunc% = True ' 滅砦偌狟偌聽壽敕勤趕遺敦諳
				Exit Function
			End If
			'聆彸竘э竘э岆瘁閉奀
			If trnString = "Timeout" Then
				Massage = MsgBox(Msg06,vbYesNoCancel+vbInformation,Msg04)
				If Massage = vbYes Then WaitTimes = InputBox(Msg07,Msg08,WaitTimes)
				If Massage = vbCancel Then
					MainDlgFunc% = True ' 滅砦偌狟偌聽壽敕勤趕遺敦諳
					Exit Function
				End If
			End If
			'潰聆趼揹濬倰睿囀�楪√鯓Й鮽矽�
			If mAllType + mMenu + mDialog + mString + mAccTable + mVer + mOther + mSelOnly = 0 Then
				MsgBox(Msg02,vbOkOnly+vbInformation,Msg01)
				MainDlgFunc% = True ' 滅砦偌狟偌聽壽敕勤趕遺敦諳
				Exit Function
			End If
			If Join(tSelected,JoinStr) = nSelected Then Exit Function
			tSelected = Split(nSelected,JoinStr)
		 	If EngineWrite(EngineDataList,tWriteLoc,"All") = False Then
				If tWriteLoc = EngineFilePath Then Msg03 = Msg03 & EngineFilePath
				If tWriteLoc = EngineRegKey Then Msg03 = Msg03 & EngineRegKey
				If MsgBox(Msg03,vbYesNo+vbInformation,Msg01) = vbNo Then Exit Function
				MainDlgFunc% = True ' 滅砦偌狟偌聽壽敕勤趕遺敦諳
				Exit Function
			End If
		End If
		If DlgItem$ = "SaveButton" And Join(tSelected,JoinStr) <> nSelected Then
			tSelected = Split(nSelected,JoinStr,-1)
			tSelected(23) = "1"
			If EngineWrite(EngineDataList,tWriteLoc,"All") = False Then
				If tWriteLoc = EngineFilePath Then Msg03 = Msg03 & EngineFilePath
				If tWriteLoc = EngineRegKey Then Msg03 = Msg03 & EngineRegKey
				MsgBox(Msg03,vbOkOnly+vbInformation,Msg01)
			Else
				DlgEnable "SaveButton",False
			End If
			DlgValue "KeepSet",1
		Else
			If Join(tSelected,JoinStr) <> nSelected Then
				DlgEnable "SaveButton",True
			Else
				DlgEnable "SaveButton",False
			End If
		End If
		If DlgItem$ = "HelpButton" Then	Call EngineHelp("MainHelp")
		If DlgItem$ = "AboutButton" Then Call EngineHelp("About")
		If DlgItem$ <> "OKButton" And DlgItem$ <> "CancelButton" Then
			MainDlgFunc% = True ' 滅砦偌狟偌聽壽敕勤趕遺敦諳
		End If
	End Select
End Function


' 翋最唗
Sub Main
	Dim i As Integer,j As Integer,srcString As String,trnString As String,TranLang As String
	Dim src As PslTransList,TrnList As PslTransList,TransListOpen As Boolean
	Dim SrcLangList() As String,LangPairList() As String,xmlHttp As Object,objStream As Object
	Dim srcLngFind As Integer,trnLngFind As Integer,StringCount As Integer
	Dim CheckID As Integer,EngineID As Integer,srcLng As String,trnLng As String
	Dim TranedCount As Integer,SkipedCount As Integer,NotChangeCount As Integer
	Dim LineNumErrCount As Integer,accKeyNumErrCount As Integer,ErrorCount As Integer

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
		Msg00 = "版權所有(C) 2010 by wanfu  版本: " & Version
		Msg01 = "連線翻譯巨集"
		Msg02 = "本程式通過所選的連線翻譯引擎和其他選項，自動翻譯清單中的字串。" & _
				"您可以自訂翻譯引擎及其參數。"

		Msg03 = "翻譯清單: "
		Msg04 = "翻譯引擎"
		Msg05 = "翻譯原文"
		Msg06 = "翻譯字串"
		Msg07 = "全部(&A)"
		Msg08 = "選單(&M)"
		Msg09 = "對話方塊(&D)"
		Msg10 = "字串表(&S)"
		Msg11 = "加速器表(&A)"
		Msg12 = "版本(&V)"
		Msg13 = "其他(&O)"
		Msg14 = "僅選擇(&L)"

		Msg15 = "跳過字串"
		Msg16 = "供覆審(&K)"
		Msg17 = "已驗證(&E)"
		Msg18 = "未翻譯(&N)"
		Msg19 = "全為數字和符號(&M)"
		Msg20 = "全為大寫英文(&U)"
		Msg21 = "全為小寫英文(&L)"

		Msg22 = "字串處理"
		Msg23 = "設定:"
		Msg24 = "翻譯前:"
		Msg25 = "翻譯後:"
		Msg26 = "自動選取設定"
		Msg27 = "去除便捷鍵(&K)"
		Msg28 = "去除加速器(&J)"
		Msg29 = "替換特定字元並在翻譯後還原(&P)"
		Msg30 = "分行翻譯(&F)"
		Msg31 = "糾正便捷鍵、終止符、加速器(&K)"
		Msg32 = "替換特定字元(&R)"

		Msg33 = "繼續時自動儲存所有選取(&V)"
		Msg34 = "顯示輸出訊息(&O)"
		Msg35 = "新增翻譯註解(&M)"

		Msg36 = "關於(&A)"
		Msg37 = "說明(&H)"
		Msg38 = "設定(&S)"
		Msg39 = "儲存選取(&L)"

		Msg42 = "確認"
		Msg43 = "訊息"
		Msg44 = "錯誤"
		Msg45 =	"您的 Passolo 版本太低，本巨集僅適用於 Passolo 6.0 及以上版本，請升級後再使用。"
		Msg46 = "請選取一個翻譯清單！"
		Msg47 = "正在建立和更新翻譯清單..."
		Msg48 = "無法建立和更新翻譯清單，請檢查您的專案設定。"
		Msg49 = "引擎自動翻譯"
		Msg50 = "該清單未被開啟。此狀態下不可進行連線翻譯。" & vbCrLf & _
				"您需要讓系統自動開啟該翻譯清單嗎？"
		Msg51 = "正在開啟翻譯清單..."
		Msg52 = "無法開啟翻譯清單，請檢查您的專案設定。"
		Msg53 = "該清單已處於開啟狀態。此狀態下進行連線翻譯將使您" & vbCrLf & _
				"未儲存的翻譯無法還原。為了安全，系統將先儲存您的" & vbCrLf & _
				"翻譯，然後進行連線翻譯。" & vbCrLf & vbCrLf & _
				"您確定要讓系統自動儲存您的翻譯嗎？"
		Msg54 = "正在建立和更新翻譯來源清單..."
		Msg55 = "無法建立和更新翻譯來源清單，請檢查您的專案設定。"
		Msg56 = "該翻譯清單目標語言所對應的翻譯引擎語言代碼為空，程式將結束。"
		Msg57 = "正在翻譯和處理字串，可能需要幾分鐘，請稍侯..."
		Msg58 = "已跳過"
		Msg59 = "已翻譯"
		Msg60 = "未變更，翻譯結果和現有翻譯相同。"
		Msg62 = "字串已鎖定。"
		Msg63 = "字串唯讀。"
		Msg64 = "字串已翻譯供覆審。"
		Msg65 = "字串已翻譯並驗證。"
		Msg66 = "字串未翻譯。"
		Msg67 = "字串為空或全為空格。"
		Msg68 = "字串全為數字和符號。"
		Msg69 = "字串全為大寫英文或數字元號。"
		Msg70 = "字串全為小寫英文或數字元號。"
		Msg71 = "合計用時: "
		Msg72 = "hh 小時 mm 分 ss 秒"
		Msg73 = "英文到中文"
		Msg74 = "中文到英文"
		Msg75 = "，"
		Msg76 = "並"
		Msg77 = "。"
	Else
		Msg00 = "唳�佯齾�(C) 2010 by wanfu  唳掛: " & Version
		Msg01 = "婓盄楹祒粽"
		Msg02 = "掛最唗籵徹垀恁腔婓盄楹祒竘э睿む坻恁砐ㄛ赻雄楹祒蹈桶笢腔趼揹﹝" & _
				"蠟褫眕赻隅砱楹祒竘э摯む統杅﹝"

		Msg03 = "楹祒蹈桶: "
		Msg04 = "楹祒竘э"
		Msg05 = "楹祒埭恅"
		Msg06 = "楹祒趼揹"
		Msg07 = "�垓�(&A)"
		Msg08 = "粕等(&M)"
		Msg09 = "勤趕遺(&D)"
		Msg10 = "趼揹桶(&S)"
		Msg11 = "樓厒ん桶(&A)"
		Msg12 = "唳掛(&V)"
		Msg13 = "む坻(&O)"
		Msg14 = "躺恁隅(&L)"

		Msg15 = "泐徹趼揹"
		Msg16 = "鼎葩机(&K)"
		Msg17 = "眒桄痐(&E)"
		Msg18 = "帤楹祒(&N)"
		Msg19 = "�屋羌�趼睿睫瘍(&M)"
		Msg20 = "�屋玫鯗植卅�(&U)"
		Msg21 = "�屋肱－植卅�(&L)"

		Msg22 = "趼揹揭燴"
		Msg23 = "饜离:"
		Msg24 = "楹祒ヶ:"
		Msg25 = "楹祒綴:"
		Msg26 = "赻雄恁寁饜离"
		Msg27 = "�戊�辦豎瑩(&K)"
		Msg28 = "�戊�樓厒ん(&J)"
		Msg29 = "杸遙杻隅趼睫甜婓楹祒綴遜埻(&P)"
		Msg30 = "煦俴楹祒(&F)"
		Msg31 = "壁淏辦豎瑩﹜笝砦睫﹜樓厒ん(&K)"
		Msg32 = "杸遙杻隅趼睫(&R)"

		Msg33 = "樟哿奀赻雄悵湔垀衄恁寁(&V)"
		Msg34 = "珆尨怀堤秏洘(&O)"
		Msg35 = "氝樓楹祒蛁庋(&M)"

		Msg36 = "壽衾(&A)"
		Msg37 = "堆翑(&H)"
		Msg38 = "扢离(&S)"
		Msg39 = "悵湔恁寁(&L)"

		Msg42 = "�溜�"
		Msg43 = "陓洘"
		Msg44 = "渣昫"
		Msg45 =	"蠟腔 Passolo 唳掛怮腴ㄛ掛粽躺巠蚚衾 Passolo 6.0 摯眕奻唳掛ㄛ③汔撰綴婬妏蚚﹝"
		Msg46 = "③恁寁珨跺楹祒蹈桶ㄐ"
		Msg47 = "淏婓斐膘睿載陔楹祒蹈桶..."
		Msg48 = "拸楊斐膘睿載陔楹祒蹈桶ㄛ③潰脤蠟腔源偶扢离﹝"
		Msg49 = "竘э赻雄楹祒"
		Msg50 = "蜆蹈桶帤掩湖羲﹝森袨怓狟祥褫眕輛俴婓盄楹祒﹝" & vbCrLf & _
				"蠟剒猁�譁腴啻堈秩翾疙繩倡蹅訇簋艞�"
		Msg51 = "淏婓湖羲楹祒蹈桶..."
		Msg52 = "拸楊湖羲楹祒蹈桶ㄛ③潰脤蠟腔源偶扢离﹝"
		Msg53 = "蜆蹈桶眒揭衾湖羲袨怓﹝森袨怓狟輛俴婓盄楹祒蔚妏蠟" & vbCrLf & _
				"帤悵湔腔楹祒拸楊遜埻﹝峈賸假�咯疢腴魚峙�悵湔蠟腔" & vbCrLf & _
				"楹祒ㄛ�遣騣靇俶硨葽倡諢�" & vbCrLf & vbCrLf & _
				"蠟�毓例糾譁腴啻堈秧ㄣ磑�腔楹祒鎘ˋ"
		Msg54 = "淏婓斐膘睿載陔楹祒懂埭蹈桶..."
		Msg55 = "拸楊斐膘睿載陔楹祒懂埭蹈桶ㄛ③潰脤蠟腔源偶扢离﹝"
		Msg56 = "蜆楹祒蹈桶醴梓逄晟垀勤茼腔楹祒竘э逄晟測鎢峈諾ㄛ最唗蔚豖堤﹝"
		Msg57 = "淏婓楹祒睿揭燴趼揹ㄛ褫夔剒猁撓煦笘ㄛ③尕綜..."
		Msg58 = "眒泐徹"
		Msg59 = "眒楹祒"
		Msg60 = "帤載蜊ㄛ楹祒賦彆睿珋衄楹祒眈肮﹝"
		Msg62 = "趼揹眒坶隅﹝"
		Msg63 = "趼揹硐黍﹝"
		Msg64 = "趼揹眒楹祒鼎葩机﹝"
		Msg65 = "趼揹眒楹祒甜桄痐﹝"
		Msg66 = "趼揹帤楹祒﹝"
		Msg67 = "趼揹峈諾麼�屋矽楖鞢�"
		Msg68 = "趼揹�屋羌�趼睿睫瘍﹝"
		Msg69 = "趼揹�屋玫鯗植卅躉藡�趼睫瘍﹝"
		Msg70 = "趼揹�屋肱－植卅躉藡�趼睫瘍﹝"
		Msg71 = "磁數蚚奀: "
		Msg72 = "hh 苤奀 mm 煦 ss 鏃"
		Msg73 = "荎恅善笢恅"
		Msg74 = "笢恅善荎恅"
		Msg75 = "ㄛ"
		Msg76 = "甜"
		Msg77 = "﹝"
	End If

	If PSL.Version < 600 Then
		MsgBox Msg45,vbOkOnly+vbInformation,Msg43
		Exit Sub
	End If

	'潰聆 Adodb.Stream 岆瘁湔婓甜鳳�＝硊�晤鎢蹈桶
	Set objStream = CreateObject("Adodb.Stream")
	If objStream Is Nothing Then
		MsgBox(Msg78,vbOkOnly+vbInformation,Msg43)
		Exit Sub
	End If
	Set objStream = Nothing
	CodeList = CodePageList(0,49)

	Set trn = PSL.ActiveTransList
	'潰聆楹祒蹈桶岆瘁掩恁寁
	If trn Is Nothing Then
		MsgBox Msg46,vbOkOnly+vbInformation,Msg43
		Exit Sub
	End If

	'鳳�㊣椒植擿埡訇�
	ReDim SrcLangList(0)
	SrcLangList(0) = PSL.GetLangCode(trn.SourceList.LangID,pslCodeText)
	j = 1
	For i = 1 To trn.Project.Languages.Count
		Set TrnList = trn.Project.TransLists(i)
		If TrnList.ListID <> trn.ListID Then
			ReDim Preserve SrcLangList(j)
			SrcLangList(j) = PSL.GetLangCode(TrnList.Language.LangID,pslCodeText)
			j = j + 1
		End If
	Next i

	'場宎趙杅郪
	ReDim AppNames(3),AppPaths(3)
	ReDim DefaultEngineList(2),EngineList(0),EngineDataList(0)
	DefaultEngineList(0) = "Microsoft"
	DefaultEngineList(1) = "Google"
	DefaultEngineList(2) = "Yahoo"
	ReDim DefaultCheckList(1),CheckList(0),CheckDataList(0)
	DefaultCheckList(0) = Msg73
	DefaultCheckList(1) = Msg74

	'黍�◎倡遻�э扢离
	If EngineGet("",EngineList,EngineDataList,"") = False Then
		For i = LBound(DefaultEngineList) To UBound(DefaultEngineList)
			EngineName = DefaultEngineList(i)
			LangPair = Join(LangCodeList(EngineName,OSLanguage,0,107),SubJoinStr)
			Data = EngineName & JoinStr & EngineSettings(EngineName) & JoinStr & LangPair
			CreateArray(EngineName,Data,EngineList,EngineDataList)
		Next i
	Else
		For i = LBound(DefaultEngineList) To UBound(DefaultEngineList)
			EngineName = DefaultEngineList(i)
			If InStr(Join(DefaultEngineList,JoinStr),EngineName) = 0 Then
				LangPair = Join(LangCodeList(EngineName,OSLanguage,0,107),SubJoinStr)
				Data = EngineName & JoinStr & EngineSettings(EngineName) & JoinStr & LangPair
				CreateArray(EngineName,Data,EngineList,EngineDataList)
			End If
		Next i
	End If

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
	If Join(tUpdateSet) <> "" Then
		updateMode = tUpdateSet(0)
		updateUrl = tUpdateSet(1)
		CmdPath = tUpdateSet(2)
		CmdArg = tUpdateSet(3)
		updateCycle = tUpdateSet(4)
		updateDate = tUpdateSet(5)
		If updateMode = "" Then
			tUpdateSet(0) = "1"
			updateMode = "1"
		End If
		If updateUrl = "" Then tUpdateSet(1) = updateMainUrl & vbCrLf & updateMinorUrl
		If CmdPath = "" Or (CmdPath <> "" And Dir(CmdPath) = "") Then
			CmdPathArgList = Split(getCMDPath(".rar","",""),JoinStr)
			tUpdateSet(2) = CmdPathArgList(0)
			tUpdateSet(3) = CmdPathArgList(1)
		End If
	Else
		updateMode = "1"
		updateUrl = updateMainUrl & vbCrLf & updateMinorUrl
		updateCycle = "7"
		Temp = updateMode & JoinStr & updateUrl & JoinStr & getCMDPath(".rar","","") & _
				JoinStr & updateCycle & JoinStr & updateDate
		tUpdateSet = Split(Temp,JoinStr)
	End If
	If updateMode <> "" And updateMode <> "2" Then
		If updateDate <> "" Then
			i = CInt(DateDiff("d",CDate(updateDate),Date))
			m = StrComp(Format(Date,"yyyy-MM-dd"),updateDate)
			If updateCycle <> "" Then n = i - CInt(updateCycle)
		End If
		If updateDate = "" Or (m = 1 And n >= 0) Then
			If Download(updateMethod,updateUrl,updateAsync,updateMode) = True Then
				tUpdateSet(5) = Format(Date,"yyyy-MM-dd")
				EngineWrite(EngineDataList,tWriteLoc,"Update")
				GoTo ExitSub
			Else
				tUpdateSet(5) = Format(Date,"yyyy-MM-dd")
				EngineWrite(EngineDataList,tWriteLoc,"Update")
			End If
		End If
	End If

	'勤趕遺
	Msg03 = Msg03 & trn.Title & " - " & PSL.GetLangCode(trn.Language.LangID,pslCodeText)
	Begin Dialog UserDialog 620,462,Msg01,.MainDlgFunc ' %GRID:10,7,1,1
		Text 20,7,580,14,Msg00,.Text1,2
		Text 20,28,580,28,Msg02,.Text2
		Text 20,63,580,14,Msg03,.Text3,2

		GroupBox 20,84,280,56,Msg04,.Configuration
		DropListBox 40,105,240,21,EngineList(),.EngineList

		GroupBox 320,84,280,56,Msg05,.SrcLang
		DropListBox 340,105,240,21,SrcLangList(),.SrcLangList

		GroupBox 20,147,580,63,Msg06,.StrTypeSelection
		CheckBox 40,164,130,14,Msg07,.AllType
		CheckBox 180,164,130,14,Msg08,.Menu
		CheckBox 320,164,130,14,Msg09,.Dialog
		CheckBox 460,164,130,14,Msg10,.Strings
		CheckBox 40,185,130,14,Msg11,.AccTable
		CheckBox 180,185,130,14,Msg12,.Versions
		CheckBox 320,185,130,14,Msg13,.Other
		CheckBox 460,185,130,14,Msg14,.Seleted

		GroupBox 20,217,580,63,Msg15,.SkipSelection
		CheckBox 40,234,190,14,Msg16,.ForReview
		CheckBox 240,234,180,14,Msg17,.Validated
		CheckBox 430,234,160,14,Msg18,.NotTran
		CheckBox 40,255,190,14,Msg19,.NumAndSymbol
		CheckBox 240,255,180,14,Msg20,.AllUCase
		CheckBox 430,255,160,14,Msg21,.AllLCase

		GroupBox 20,287,580,112,Msg22,.PreProcessing
		Text 40,307,70,14,Msg23,.Text4
		Text 40,331,70,14,Msg24,.Text5
		Text 40,374,70,14,Msg25,.Text6
		DropListBox 120,304,270,21,CheckList(),.CheckList
		CheckBox 410,307,180,14,Msg26,.AutoSelection
		CheckBox 120,331,280,14,Msg27,.AccessKey
		CheckBox 410,331,180,14,Msg28,.Accelerator
		CheckBox 120,353,280,14,Msg29,.PreStrRep
		CheckBox 410,353,180,14,Msg30,.SplitTran
		CheckBox 120,374,280,14,Msg31,.CheckTrn
		CheckBox 410,374,180,14,Msg32,.AppStrRep

		CheckBox 20,406,240,14,Msg33,.KeepSet
		CheckBox 270,406,160,14,Msg34,.ShowMsg
		CheckBox 440,406,160,14,Msg35,.TranComment
		PushButton 20,434,90,21,Msg36,.AboutButton
		PushButton 110,434,90,21,Msg37,.HelpButton
		PushButton 200,434,90,21,Msg38,.SetButton
		PushButton 290,434,110,21,Msg39,.SaveButton
		OKButton 420,434,90,21,.OKButton '5 樟哿
		CancelButton 510,434,90,21,.CancelButton '6 �＋�
	End Dialog
	Dim dlg As UserDialog
	If Dialog(dlg) = 0 Then GoTo ExitSub
	AllCont = 1
	AccKey = 0
	EndChar = 0
	Acceler = 0
	EngineID = dlg.EngineList
	CheckID = dlg.CheckList

	'鳳�＝硒挫閛邰暻�
	If dlg.Menu = 1 Then StrTypes = "|Menu|"
	If dlg.Dialog = 1 Then StrTypes = StrTypes & "|Dialog|"
	If dlg.Strings = 1 Then StrTypes = StrTypes & "|StringTable|"
	If dlg.AccTable = 1 Then StrTypes = StrTypes & "|AcceleratorTable|"
	If dlg.Versions = 1 Then StrTypes = StrTypes & "|Version|"

	'枑尨湖羲壽敕腔楹祒蹈桶ㄛ眕晞褫眕婓盄楹祒
	TransListOpen = False
	If trn.IsOpen = False Then
		Massage = MsgBox(Msg03 & vbCrLf & vbCrLf & Msg50,vbOkCancel,Msg42)
		If Massage = vbOK Then
			PSL.Output Msg51
			If trn.Open = False Then
				MsgBox Msg52,vbOkOnly+vbInformation,Msg44
				GoTo ExitSub
			Else
				TransListOpen = True
			End If
		End If
		If Massage = vbCancel Then GoTo ExitSub
	End If

	'枑尨悵湔湖羲腔楹祒蹈桶ㄛ眕轎揭燴綴杅擂祥褫閥葩
	'If trn.IsOpen = True Then
		'Massage = MsgBox(Msg03 & vbCrLf & vbCrLf & Msg53,vbYesNoCancel,Msg42)
		'If Massage = vbYes Then trn.Save
		'If Massage = vbCancel Then Exit Sub
	'End If

	'�蝜�楹祒蹈桶腔載蜊奀潔俀衾埻宎蹈桶ㄛ赻雄載陔
	If trn.SourceList.LastChange > trn.LastChange Then
		PSL.Output Msg47
		If trn.Update = False Then
			MsgBox Msg48,vbOkOnly+vbInformation,Msg44
			GoTo ExitSub
		End If
	End If

	'恁寁楹祒懂埭蹈桶
	If dlg.SrcLangList <> 0 Then
		LangDec = SrcLangList(dlg.SrcLangList)
		For i = 1 To trn.Project.TransLists.Count
			Set TrnList = trn.Project.TransLists(i)
			If TrnList.Title = trn.Title Then
				If PSL.GetLangCode(TrnList.Language.LangID,pslCodeText) = LangDec Then
					Set src = trn.Project.TransLists(i)
					Exit For
				End If
			End If
		Next i
		'�蝜�楹祒懂埭蹈桶腔載蜊奀潔俀衾埻宎蹈桶ㄛ赻雄載陔
		If src.SourceList.LastChange > src.LastChange Then
			PSL.Output Msg54
			If src.Update = False Then
				MsgBox Msg55,vbOkOnly+vbInformation,Msg44
				GoTo ExitSub
			End If
		End If
	End If

	'扢离潰脤粽蚳蚚腔蚚誧隅砱扽俶
	If Not trn Is Nothing Then
		If trn.Property(19980) <> "CheckAccessKeys" Then
			trn.Property(19980) = "CheckAccessKeys"
		End If
	End If

	'鳳�·SL腔懂埭逄晟測鎢
	If dlg.SrcLangList = 0 Then
		srcLng = PSL.GetLangCode(trn.SourceList.LangID,pslCode639_1)
	Else
		srcLng = PSL.GetLangCode(src.Language.LangID,pslCode639_1)
	End If
	If srcLng = "zh" Then
		If dlg.SrcLangList = 0 Then
			srcLng = PSL.GetLangCode(trn.SourceList.LangID,pslCodeLangRgn)
		Else
			srcLng = PSL.GetLangCode(src.Language.LangID,pslCodeLangRgn)
		End If
		If srcLng = "zh-CHS" Or srcLng = "zh-SG" Then srcLng = "zh-CN"
		If srcLng = "zh-CHT" Or srcLng = "zh-HK" Or srcLng = "zh-MO" Then srcLng = "zh-TW"
	End If

	'鳳�·SL腔醴梓逄晟測鎢
	trnLng = PSL.GetLangCode(trn.Language.LangID,pslCode639_1)
	If trnLng = "zh" Or trnLng = "ja" Or trnLng = "ko" Then TranLang = "Asia"
	If trnLng = "zh" Then
		trnLng = PSL.GetLangCode(trn.Language.LangID,pslCodeLangRgn)
		If trnLng = "zh-CHS" Or trnLng = "zh-SG" Then trnLng = "zh-CN"
		If trnLng = "zh-CHT" Or trnLng = "zh-HK" Or trnLng = "zh-MO" Then trnLng = "zh-TW"
	End If

	'脤梑楹祒竘э笢勤茼腔逄晟測鎢
	srcLngFind = 0
	trnLngFind = 0
	TempArray = Split(EngineDataList(EngineID),JoinStr)
	LangArray = Split(TempArray(2),SubJoinStr)
	EngineName = TempArray(0)
	For i = 0 To UBound(LangArray)
		LangPairList = Split(LangArray(i),LngJoinStr)
		If LCase(srcLng) = LCase(LangPairList(1)) Then
			srcLng = LangPairList(2)
			srcLngFind = 1
		End If
		If LCase(trnLng) = LCase(LangPairList(1)) Then
			trnLng = LangPairList(2)
			trnLngFind = 1
		End If
		If srcLngFind + trnLngFind = 2 Then Exit For
	Next i
	If trnLng = "" Or trnLng = NullValue Then
		MsgBox Msg56,vbOkOnly+vbInformation,Msg44
		GoTo ExitSub
	End If
	LangPair = srcLng & LngJoinStr & trnLng

	'庋溫祥婬妏蚚腔雄怓杅郪垀妏蚚腔囀湔
	Erase TempArray,LangArray,SrcLangList,LangPairList
	Erase CheckListBak,CheckDataListBak,tempCheckList,tempCheckDataList
	Erase EngineListBak,EngineDataListBak
	Erase DelLngNameList,DelSrcLngList,DelTranLngList
	Erase AppNames,AppPaths,FileList

	'跦擂岆瘁恁寁 "躺恁隅趼揹" 砐扢离猁楹祒腔趼揹杅
	If dlg.Seleted = 0 Then StringCount = trn.StringCount
	If dlg.Seleted = 1 Then StringCount = trn.StringCount(pslSelection)

	'羲宎揭燴藩沭趼揹
	PSL.OutputWnd.Clear
	PSL.Output Msg57
	StartTimes = Timer
	Set xmlHttp = CreateObject(DefaultObject)
	If xmlHttp Is Nothing Then Set xmlHttp = CreateObject("Msxml2.XMLHTTP")
	For j = 1 To StringCount
		'跦擂岆瘁恁寁 "躺恁隅趼揹" 砐扢离猁楹祒腔趼揹
		If dlg.Seleted = 0 Then Set TransString = trn.String(j)
		If dlg.Seleted = 1 Then Set TransString = trn.String(j,pslSelection)

		'秏洘睿趼揹場宎趙甜鳳�◎倡蹅訇穔鰍笘釔椒景芛倡鄶硒�
		SkipMsg = ""
		Massage = ""
		LineMsg = ""
		AcckeyMsg = ""
		ChangeMsg = ""
		orjSrcString = TransString.SourceText
		orjtrnString = TransString.Text

		'趼揹濬倰揭燴
		If dlg.AllType = 0 And dlg.Seleted = 0 Then
			If InStr("Menu|Dialog|StringTable|AcceleratorTable|Version",TransString.ResType) Then
				If InStr(StrTypes,TransString.ResType) = 0 Then GoTo Skip
			Else
				If dlg.Other = 0 Then GoTo Skip
			End If
		End If

		'泐徹眒坶隅腔趼揹
		If TransString.State(pslStateLocked) = True Then
			SkipMsg = Msg58 & Msg75 & Msg62
			SkipedCount = SkipedCount + 1
			GoTo Skip
		End If
		'泐徹硐黍腔趼揹
		If TransString.State(pslStateReadOnly) = True Then
			SkipMsg = Msg58 & Msg75 & Msg63
			SkipedCount = SkipedCount + 1
			GoTo Skip
		End If
		'泐徹眒楹祒鼎葩机腔趼揹
		If dlg.ForReview = 1 And TransString.State(pslStateTranslated) = True Then
			If TransString.State(pslStateReview) = True Then
				SkipMsg = Msg58 & Msg75 & Msg64
				SkipedCount = SkipedCount + 1
				GoTo Skip
			End If
		End If
		'泐徹眒楹祒甜桄痐腔趼揹
		If dlg.Validated = 1 And TransString.State(pslStateTranslated) = True Then
			If TransString.State(pslStateReview) = False Then
				SkipMsg = Msg58 & Msg75 & Msg65
				SkipedCount = SkipedCount + 1
				GoTo Skip
			End If
		End If
		'泐徹帤楹祒腔趼揹
		If dlg.NotTran = 1 And TransString.State(pslStateTranslated) = False Then
			SkipMsg = Msg58 & Msg75 & Msg66
			SkipedCount = SkipedCount + 1
			GoTo Skip
		End If
		'泐徹峈諾麼�屋矽楖騊儷硒�
		If Trim(orjSrcString) = "" Then
			SkipMsg = Msg58 & Msg75 & Msg67
			SkipedCount = SkipedCount + 1
			GoTo Skip
		End If
		'泐徹�屋羌�趼睿睫瘍腔趼揹
		If dlg.NumAndSymbol = 1 Then
			If CheckStr(orjSrcString,"0-64,91-96,123-191") = True Then
				SkipMsg = Msg58 & Msg75 & Msg68
				SkipedCount = SkipedCount + 1
				GoTo Skip
			End If
		End If
		'泐徹�屋玫鯗植卅警儷硒�
		If dlg.AllUCase = 1 Then
			If CheckStr(orjSrcString,"65-90") = True Then
				SkipMsg = Msg58 & Msg75 & Msg69
				SkipedCount = SkipedCount + 1
				GoTo Skip
			End If
		End If
		'泐徹�屋肱－植卅警儷硒�
		If dlg.AllLCase = 1 Then
			If CheckStr(orjSrcString,"97-122") = True Then
				SkipMsg = Msg58 & Msg75 & Msg70
				SkipedCount = SkipedCount + 1
				GoTo Skip
			End If
		End If

		'鳳�◎倡郺棒儷硒�
		If dlg.SrcLangList = 0 Then
			srcString = TransString.SourceText
		Else
			If dlg.Seleted = 0 Then
				srcString = src.String(j).Text
			Else
				For i = 1 To src.StringCount
					If src.String(i).Number = TransString.Number Then
						srcString = src.String(i).Text
						Exit For
					End If
				Next i
			End If
		End If

		'羲宎啎揭燴甜楹祒趼揹
		If dlg.PreStrRep = 1 Then srcString = ReplaceStr(CheckID,srcString,1)
		If dlg.SplitTran = 0 Then
			If dlg.AccessKey = 1 Then srcString = AccessKeyHanding(CheckID,srcString)
			If dlg.Accelerator = 1 Then srcString = AcceleratorHanding(CheckID,srcString)
			trnString = getTranslate(EngineID,xmlHttp,srcString,LangPair,0)
		Else
			mAccKey = dlg.AccessKey
			mAccelerator = dlg.Accelerator
			Temp = EngineID & JoinStr & CheckID & JoinStr & mAccKey & JoinStr & mAccelerator
			trnString = SplitTran(xmlHttp,srcString,LangPair,Temp,0)
		End If

		'羲宎綴揭燴趼揹甜杸遙埻衄楹祒
		If trnString <> orjSrcString And trnString <> orjTrnString Then
			If dlg.CheckTrn = 1 Then
				NewtrnString = CheckHanding(CheckID,orjSrcString,trnString,TranLang)
			Else
				NewtrnString = trnString
			End If
			If dlg.PreStrRep = 1 Then NewtrnString = ReplaceStr(CheckID,NewtrnString,2)
			If dlg.AppStrRep = 1 Then NewtrnString = ReplaceStr(CheckID,NewtrnString,0)

			If NewtrnString <> orjTrnString Then
				TransString.Text = NewtrnString
				TransString.State(pslStateReview) = True
				If dlg.TranComment = 1 Then
					TransString.TransComment = EngineName & " " & Msg49
				Else
					TransString.TransComment = ""
				End If
				TranedCount = TranedCount + 1
			End If
		Else
			NotChangeCount = NotChangeCount + 1
		End If

		'郪眽秏洘甜怀堤
		If dlg.ShowMsg = 1 Then
			If trnString <> orjSrcString And trnString <> orjTrnString Then
				If NewtrnString <> orjTrnString Then
					If srcLineNum <> trnLineNum Then
						LineMsg = LineErrMassage(srcLineNum,trnLineNum,LineNumErrCount)
					End If
					If srcAccKeyNum <> trnAccKeyNum Then
						AcckeyMsg = AccKeyErrMassage(srcAccKeyNum,trnAccKeyNum,accKeyNumErrCount)
					End If
					ChangeMsg = ReplaceMassage(trnString,NewtrnString)
					If AcckeyMsg & LineMsg & ChangeMsg <> "" Then
						Massage = Msg59 & Msg75 & Msg76 & ChangeMsg & AcckeyMsg & LineMsg
					Else
						Massage = Msg59 & Msg77
					End If
				Else
					Massage = Msg60
				End If
			Else
				Massage = Msg60
			End If
			TransString.OutputError(Massage)
		End If
		Skip:
		If dlg.ShowMsg = 1 And SkipMsg <> "" Then TransString.OutputError(SkipMsg)
	Next j
	Set xmlHttp = Nothing

	'楹祒數杅摯秏洘怀堤
	ErrorCount = LineNumErrCount + accKeyNumErrCount
	PSL.Output TranMassage(TranedCount,SkipedCount,NotChangeCount,ErrorCount)
	If ErrorCount = 0 And TransListOpen = True Then trn.Close
	EndTimes = Timer
	PSL.Output Msg71 & Format(DateAdd("s",EndTimes - StartTimes,0),Msg72)

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
	If Join(tUpdateSet) <> "" Then
		If Mode = "" Then Mode = tUpdateSet(0)
		If Url = "" Then Url = tUpdateSet(1)
		ExePath = tUpdateSet(2)
		Argument = tUpdateSet(3)
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
				n = InStrRev(LCase(updateUrl),"/download")
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
		If Url <> Join(UrlList,vbCrLf) Then tUpdateSet(1) = Join(UrlList,vbCrLf)
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


'鳳�√硨葽倡�
Function getTranslate(ID As Integer,xmlHttp As Object,srcStr As String,LngPair As String,fType As Integer) As String
    Dim UrlData As String,trnStr As String,LangFrom As String,LangTo As String
	Dim Temp As String,pos As Integer,Code As String,srcStrBak As String

	TempArray = Split(EngineDataList(ID),JoinStr)
	SetsArray = Split(TempArray(1),SubJoinStr)
	AppId = SetsArray(1)
	Url = SetsArray(2)
	UrlTemplate = SetsArray(3)
	Method = SetsArray(4)
	Async = SetsArray(5)
	User = SetsArray(6)
	Password = SetsArray(7)
	Body = SetsArray(8)
	RequestHeader = SetsArray(9)
	responseType = SetsArray(10)
	TranBeforeStr = SetsArray(11)
	TranAfterStr = SetsArray(12)

	If Url = "" Then
		If fType = 3 Then getTranslate = "NullUrl"
		Exit Function
	End If

	If LngPair <> "" Then
		LangFrom = Left(LngPair,InStr(LngPair,LngJoinStr)-1)
		LangTo = Mid(LngPair,InStr(LngPair,LngJoinStr)+1)
	End If

	srcStrBak = srcStr
	If InStr(LCase(RequestHeader),"charset") Then
		Temp = Mid(RequestHeader,InStr(LCase(RequestHeader),"charset"))
		If InStr(Temp,"=") Then Code = ExtractStr(Temp,"=",";",1)
	Else
		On Error GoTo ErrorHandler
		xmlHttp.Open Method,Url,Async,User,Password
		xmlHttp.send()
		Temp = xmlHttp.getResponseHeader("Content-Type")
		If InStr(LCase(Temp),"charset") Then
			Temp = Mid(Temp,InStr(LCase(Temp),"charset"))
			If InStr(Temp,"=") Then Code = ExtractStr(Temp,"=",";",1)
		Else
			Temp = xmlHttp.responseText
			If InStr(LCase(Temp),"charset") Then
				Temp = Mid(Temp,InStr(LCase(Temp),"charset"))
				If InStr(Temp,"=") Then Code = ExtractStr(Temp,"=",">",1)
			ElseIf InStr(LCase(Temp),"lang") Then
				Temp = Mid(Temp,InStr(LCase(Temp),"lang"))
				If InStr(Temp,"=") Then Code = ExtractStr(Temp,"=",">",1)
			End If
		End If
		xmlHttp.Abort
		On Error GoTo 0
	End If
	If Code <> "" Then Code = RemoveBackslash(Code,"""","""",1)
	If LCase(Code) = "utf-8" Or LCase(Code) = "utf8" Then
		srcStrBak = Utf8Encode(srcStrBak)
	Else
		srcStrBak = ANSIEncode(srcStrBak)
	End If

	If UrlTemplate <> "" Then
		UrlData = LCase(UrlTemplate)
		If InStr(UrlData,"{url}") = 0 Then UrlData = Url & UrlData
		If InStr(UrlData,"{url}") Then UrlData = Replace(UrlData,"{url}",Url)
		If InStr(UrlData,"{appid}") Then UrlData = Replace(UrlData,"{appid}",AppId)
		If InStr(UrlData,"{text}") Then UrlData = Replace(UrlData,"{text}",srcStrBak)
		If InStr(UrlData,"{from}") Then UrlData = Replace(UrlData,"{from}",LangFrom)
		If InStr(UrlData,"{to}") Then UrlData = Replace(UrlData,"{to}",LangTo)
	End If

	If Body <> "" Then
		BodyData = LCase(Body)
		If InStr(BodyData,"{url}") Then BodyData = Replace(BodyData,"{url}",Url)
		If InStr(BodyData,"{appid}") Then BodyData = Replace(BodyData,"{appid}",AppId)
		If InStr(BodyData,"{text}") Then BodyData = Replace(BodyData,"{text}",srcStrBak)
		If InStr(BodyData,"{from}") Then BodyData = Replace(BodyData,"{from}",LangFrom)
		If InStr(BodyData,"{to}") Then BodyData = Replace(BodyData,"{to}",LangTo)
	End If

	If RequestHeader <> "" Then
		RequestData = LCase(RequestHeader)
		If InStr(RequestData,"{url}") Then RequestData = Replace(RequestData,"{url}",Url)
		If InStr(RequestData,"{appid}") Then RequestData = Replace(RequestData,"{appid}",AppId)
		If InStr(RequestData,"{text}") Then RequestData = Replace(RequestData,"{text}",srcStrBak)
		If InStr(RequestData,"{from}") Then RequestData = Replace(RequestData,"{from}",LangFrom)
		If InStr(RequestData,"{to}") Then RequestData = Replace(RequestData,"{to}",LangTo)
	End If

	On Error GoTo ErrorHandler
    xmlHttp.Open Method,UrlData,Async,User,Password
    If RequestHeader <> "" Then
		FindStrArr = Split(RequestData,vbCrLf)
		For i = LBound(FindStrArr) To UBound(FindStrArr)
    		FindStr = FindStrArr(i)
    		pos = InStr(FindStr,",")
			If pos <> 0 Then
				bstrHeader = Trim(Left(FindStr,pos-1))
				bstrValue = Trim(Mid(FindStr,pos+1))
				If LCase(bstrHeader) = LCase("Content-Length") Then
					xmlHttp.setRequestHeader bstrHeader,Len(bstrValue)
				Else
					xmlHttp.setRequestHeader bstrHeader,bstrValue
				End If
			End If
		Next i
	End If
   	xmlHttp.send(BodyData)
	If xmlHttp.readyState < 4 Then
		Wait WaitTimes
		If xmlHttp.readyState < 4 Then getTranslate = "Timeout"
	End If
	If xmlHttp.readyState = 4 Then
		If fType <> 2 Then
			If responseType = "responseText" Then trnStr = xmlHttp.responseText
			If responseType = "responseXML" Then
				If fType = 0 Then trnStr = ReadXML(xmlHttp.responseXML,TranBeforeStr,TranAfterStr)
				If fType = 1 Then trnStr = xmlHttp.responseText
			End If
			If responseType = "responseStream" Then
				If fType = 0 Then trnStr = BytesToBstr(xmlHttp.responseStream,Code)
				If fType = 1 Then trnStr = xmlHttp.responseText
			End If
			If responseType = "responseBody" Then trnStr = BytesToBstr(xmlHttp.responseBody,Code)
		End If
		If fType = 2 Then trnStr = xmlHttp.getAllResponseHeaders
		xmlHttp.Abort
		On Error GoTo 0

		'If responseType = "responseText" Then
		'	If LCase(Code) <> "utf-8" And LCase(Code) <> "utf8" Then
				'trnStr = ConvStr(trnStr,Code,"utf-8")
				'codepage = trn.Language.Option(pslOptionActualCodepage)
				'trnStr = PSL.ConvertASCII2Unicode(trnStr,codepage)
		'	End If
		'End If

		If fType = 0 Then
			If responseType = "responseXML" Then
				getTranslate = trnStr
			Else
				getTranslate = ExtractStr(trnStr,TranBeforeStr,TranAfterStr,0)
			End If
    	Else
    		getTranslate = trnStr
		End If
		Exit Function
	End If
	On Error GoTo 0

    ErrorHandler:
    If Err.Number <> 0 Then
    	If fType = 3 Then getTranslate = "NotConnected"
	End If
End Function


'Utf-8 晤鎢
Function Utf8Encode(textStr As String) As String
	Dim Wch As String,Uch As String,Szret As String,i As Integer,Nasc As Long
	If textStr = "" Then Exit Function
	For i = 1 To Len(textStr)
		Wch = Mid(textStr,i,1)
		Nasc = AscW(Wch)
		If Nasc < 0 Then Nasc = Nasc + 65536
		If (Nasc And &hff80) = 0 Then
			Szret = Szret & Wch
		Else
			If (Nasc And &hf000) = 0 Then
				Uch = "%" & Hex(((Nasc \2 ^ 6)) Or &hc0) & Hex(Nasc And &h3f Or &h80)
				Szret = Szret & Uch
			Else
				Uch = "%" & Hex((Nasc \ 2 ^ 12) Or &he0) & "%" & _
							Hex((Nasc \ 2 ^ 6) And &h3f Or &h80) & "%" & _
							Hex(Nasc And &h3f Or &h80)
				Szret = Szret & Uch
			End If
		End If
	Next i
	Utf8Encode = Szret
End Function


'ANSI 晤鎢
Public Function ANSIEncode(textStr As String) As String
    Dim i As Long,startIndex As Long,endIndex As Long,x() As Byte
    x = StrConv(textStr,vbFromUnicode)
    startIndex = LBound(x)
    endIndex = UBound(x)
    For i = startIndex To endIndex
        ANSIEncode = ANSIEncode & "%" & Hex(x(i))
    Next i
End Function


'蛌遙趼睫腔晤鎢跡宒
Function ConvStr(textStr As String,inCode As String,outCode As String) As String
	Dim objStream As Object
    Set objStream = CreateObject("Adodb.Stream")
    With objStream
    	.Type = 2
    	.Mode = 3
    	.CharSet = inCode
    	.Open
    	.WriteText textStr
    	.Position = 0
    	.CharSet = outCode
    	ConvStr = .ReadText
    	.Close
    End With
    Set objStream = Nothing
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


'賤昴 XML 跡宒勤砓甜枑�◎倡輷覺�
Function ReadXML(xmlObj As Object,IdNames As String,TagNames As String) As String
	Dim xmlDoc As Object,Node As Object,Item As Object,IdName As String,TagName As String
	Dim x As Integer,y As Integer,i As Integer,max As Integer
	If xmlObj Is Nothing Then Exit Function

	Set xmlDoc = CreateObject("Microsoft.XMLDOM")
	'xmlDoc.Async = False
	'xmlDoc.ValidateOnParse = False
	'xmlDoc.loadXML(xmlObj)	'樓婥趼揹
	xmlDoc.Load(xmlObj)		'樓婥勤砓
  	If xmlDoc.ReadyState > 2 Then
  		IdNameArray = Split(IdNames,"|")
		TagNameArray = Split(TagNames,"|")
		For x = 0 To UBound(IdNameArray)
			For y = 0 To UBound(TagNameArray)
				IdName = IdNameArray(x)
				TagName = TagNameArray(y)
				If IdName <> "" And TagName = "" Then
					On Error Resume Next
					Set Item = xmlDoc.getElementById(IdName)
					If Item Is Nothing Then
						Set Item = xmlDoc.getElementsByTagName(IdName)
					End If
					On Error GoTo 0
				ElseIf IdName <> "" And TagName <> "" Then
					On Error Resume Next
					Set Node = xmlDoc.getElementById(IdName)
					If Node Is Nothing Then
						Set Item = xmlDoc.getElementsByTagName(TagName)
					Else
						Set Item = Node.getElementsByTagName(TagName)
						If Item.Length = 0 Then Set Item = xmlDoc.getElementById(IdName)
					End If
					On Error GoTo 0
				ElseIf IdName = "" And TagName <> "" Then
					Set Item = xmlDoc.getElementsByTagName(TagName)
				End If
				max = Item.Length
				If max > 0 Then Exit For
			Next y
			If max > 0 Then Exit For
		Next x
		If max > 0 Then
			For i = 0 To max-1
				If ReadXML <> "" Then ReadXML = ReadXML & Item(i).Text
				If ReadXML = "" Then ReadXML = Item(i).Text    'firstChild.nodeValue
			Next i
		End If
	End If
	Set xmlDoc = Nothing
End Function


'枑�＞葆些偕鯚硊�眳潔腔硉
Function ExtractStr(textStr As String,BeforeStr As String,AfterStr As String,fType As Integer) As String
	Dim i As Integer,x As Integer,y As Integer,L1 As Long,L2 As Long
	Dim Temp As String,bStr As String,aStr As String,toFindText As String

	toFindText = textStr & vbCrLf
	BeforeStrArray = Split(BeforeStr,"|")
	AfterStrArray = Split(AfterStr,"|")
	For x = 0 To UBound(BeforeStrArray)
		For y = 0 To UBound(AfterStrArray)
			L1 = 1
			For i = 0 To 1
				bStr = BeforeStrArray(x)
				aStr = AfterStrArray(y)
				L1 = InStr(L1,toFindText,bStr)
				If L1 > 0 Then
					L1 = L1 + Len(bStr)
					L2 = InStr(L1,toFindText,aStr)
					If fType > 0 And L2 = 0 Then L2 = InStr(L1,toFindText,vbCrLf)
					If L2 <> 0 Then	Temp = Mid(toFindText,L1,L2-L1)
					If ExtractStr <> "" Then ExtractStr = ExtractStr & Temp
					If ExtractStr = "" Then ExtractStr = Temp
					If fType > 0 And i + 1 = fType Then Exit For
					i = 0
				Else
					Exit For
				End If
			Next i
			If ExtractStr <> "" Then Exit For
		Next y
		If ExtractStr <> "" Then Exit For
	Next x
End Function


'潰脤趼揹岆瘁躺婦漪杅趼睿睫瘍
Function CheckStr(textStr As String,AscRange As String) As Boolean
	Dim i As Integer,j As Integer,n As Integer,InpAsc As Long,Length As Long
	Dim pos As Integer,MinV As Long,MaxV As Long,Temp As String
	CheckStr = False
	If Len(Trim(textStr)) = 0 Then Exit Function
	n = 0
	Length = Len(textStr)
	For i = 1 To Length
		InpAsc = AscW(Mid(textStr,i,1))
		AscValue = Split(AscRange,",",-1)
		For j = 0 To UBound(AscValue)
			Temp = AscValue(j)
			pos = InStr(Temp,"-")
			If pos <> 0 Then
				MinV = CLng(Left(Temp,pos-1))
				MaxV = CLng(Mid(Temp,pos+1))
			Else
				MinV = CLng(Temp)
				MaxV = CLng(Temp)
			End If
			If InpAsc >= MinV And InpAsc <= MaxV Then
				n = n + 1
				Exit For
			End If
		Next j
	Next i
	If n = Length Then CheckStr = True
End Function


'煦俴楹祒揭燴
Function SplitTran(xmlHttp As Object,srcStr As String,LangPair As String,Arg As String,fType As Integer) As String
	Dim i As Integer,srcStrBak As String,SplitStr As String
	Dim EngineID As Integer,CheckID As Integer,mAccKey As Integer,mAccelerator As Integer

	TempArray = Split(Arg,JoinStr,-1)
	EngineID = CLng(TempArray(0))
	CheckID = CLng(TempArray(1))
	mAccKey = CLng(TempArray(2))
	mAccelerator = CLng(TempArray(3))

	'蚚杸遙楊莞煦趼揹
	srcStrBak = srcStr
	LineSplitChar = "\r\n,\r,\n"
	FindStrArr = Split(Convert(LineSplitChar),",",-1)
	For i = LBound(FindStrArr) To UBound(FindStrArr)
		FindStr = Trim(FindStrArr(i))
		If InStr(srcStrBak,FindStr) Then
			srcStrBak = Replace(srcStrBak,FindStr,"*c!N!g*")
			SplitStr = FindStr
		End If
	Next i
	srcStrArr = Split(srcStrBak,"*c!N!g*",-1)

	'鳳�￣諱迮譟倡�
	For i = LBound(srcStrArr) To UBound(srcStrArr)
		srcString = srcStrArr(i)
		If srcString <> "" Then
			If mAccKey = 1 Then srcString = AccessKeyHanding(CheckID,srcString)
			If mAccelerator = 1 Then srcString = AcceleratorHanding(CheckID,srcString)
			trnString = getTranslate(EngineID,xmlHttp,srcString,LangPair,fType)
		Else
			trnString = srcString
		End If
		If i > LBound(srcStrArr) Then SplitTran = SplitTran & SplitStr & trnString
		If i = LBound(srcStrArr) Then SplitTran = trnString
	Next i
End Function


'揭燴辦豎瑩趼睫
Function AccessKeyHanding(CheckID As Integer,srcStr As String) As String
	Dim i As Integer,j As Integer,n As Integer,posin As Integer
	Dim AccessKey As String,Stemp As Boolean

	srcStrBak = srcStr
	If InStr(srcStr,"&") = 0 Then
		AccessKeyHanding = srcStr
		Exit Function
	End If

	'鳳�×▲乳馺繭觸恀�
	TempArray = Split(CheckDataList(CheckID),JoinStr)
	SetsArray = Split(TempArray(1),SubJoinStr)
	ExcludeChar = SetsArray(0)
	CheckBracket = SetsArray(2)

	'齬壺趼揹笢腔準辦豎瑩
	If ExcludeChar <> "" Then
		FindStrArr = Split(Convert(ExcludeChar),",",-1)
		For i = LBound(FindStrArr) To UBound(FindStrArr)
			FindStr = LTrim(FindStrArr(i))
			If FindStr <> "" Then
				srcStrBak = Replace(srcStrBak,FindStr,"*a" & i & "!N!" & i & "d*")
			End If
		Next i
	End If

	'鳳�▼儠敯�甜�戊�
	If CheckBracket <> "" Then
		posin = InStrRev(srcStrBak,"&")
		If posin > 1 Then
			FindStrArr = Split(Convert(CheckBracket),",",-1)
			For i = LBound(FindStrArr) To UBound(FindStrArr)
				FindStr = Trim(FindStrArr(i))
				If FindStr <> "" Then
					LFindStr = Trim(Left(FindStr,1))
					RFindStr = Trim(Right(FindStr,1))
					For n = 0 To 1
						Stemp = False
						If InStrRev(srcStrBak,LFindStr & Mid(srcStrBak,posin,2) & RFindStr) Then
							AccessKey = Mid(srcStrBak,posin-1,4)
							srcStrBak = Replace(srcStrBak,AccessKey,"")
						Else
							PreStr = Left(srcStrBak,posin-1)
							AppStr = Mid(srcStrBak,posin)
							srcStrBak = PreStr & Replace(AppStr,"&","",,1)
						End If
						If posin > 1 Then
							posin = InStrRev(srcStrBak,"&",posin - 1)
							If posin > 1 Then n = 0
						End If
						If posin <= 1 Then Exit For
					Next n
				End If
				If posin <= 1 Then Exit For
			Next i
		ElseIf posin = 1 Then
			srcStrBak = Replace(srcStrBak,"&","")
		End If
	End If

	'遜埻趼揹笢掩齬壺腔準辦豎瑩
	If ExcludeChar <> "" Then
		FindStrArr = Split(Convert(ExcludeChar),",",-1)
		For i = LBound(FindStrArr) To UBound(FindStrArr)
			FindStr = LTrim(FindStrArr(i))
			If FindStr <> "" Then
				srcStrBak = Replace(srcStrBak,"*a" & i & "!N!" & i & "d*",FindStr)
			End If
		Next i
	End If
	AccessKeyHanding = srcStrBak
End Function


'揭燴樓厒ん趼睫
Function AcceleratorHanding(CheckID As Integer,srcStr As String) As String
	Dim i As Integer,j As Integer,x As Integer,y As Integer
	Dim Shortcut As String,ShortcutKey As String,FindStr As String

	'鳳�×▲乳馺繭觸恀�
	TempArray = Split(CheckDataList(CheckID),JoinStr)
	SetsArray = Split(TempArray(1),SubJoinStr)
	CheckShortChar = SetsArray(8)
	CheckShortKey = SetsArray(9)
	KeepShortKey = SetsArray(10)

	'鳳�□蚎棐�
	If CheckShortChar <> "" Then
		FindStrArr = Split(Convert(CheckShortChar),",",-1)
		For i = LBound(FindStrArr) To UBound(FindStrArr)
			FindStr = Trim(FindStrArr(i))
			If FindStr <> "" And Shortcut = "" Then
				x = UBound(Split(srcStr,FindStr,-1))
				y = 0
				If x = 1 Then
					If Left(LTrim(srcStr),Len(FindStr)) <> FindStr Then
						y = InStrRev(srcStr,FindStr)
					End If
				ElseIf x > 1 Then
					y = InStrRev(srcStr,FindStr)
				End If
				If y <> 0 Then
					ShortcutKey = Trim(Mid(srcStr,y+1))
					If Trim(ShortcutKey) = "+" Then
						Shortcut = Trim(Mid(srcStr,y))
					Else
						x = 0
						KeyArr = Split(ShortcutKey,"+",-1)
						For j = LBound(KeyArr) To UBound(KeyArr)
							x = x + CheckKeyCode(KeyArr(j),CheckShortKey)
						Next j
						If x <> 0 And x >= UBound(KeyArr) Then
							Shortcut = Trim(Mid(srcStr,y))
						End If
					End If
				End If
			End If
			If Shortcut <> "" Then Exit For
		Next i
	End If

	'�戊�樓厒ん
	If Shortcut <> "" Then
		x = InStrRev(srcStr,Shortcut)
		If x <> 0 Then AcceleratorHanding = Left(srcStr,x-1)
	Else
		AcceleratorHanding = srcStr
	End If
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
	If AsiaKey = "1" Then
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
	Else
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
	If AsiaKey <> "1" Then
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
	Dim Massage3 As String,Massage4 As String,n As Integer

	If OSLanguage = "0404" Then
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
	Else
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
		If EndStringSrc <> "" Then
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


'楹祒秏洘怀堤
Function TranMassage(tCount As Integer,sCount As Integer,nCount As Integer,eCount As Integer) As String
	If OSLanguage = "0404" Then
		Msg01 = "翻譯完成，沒有翻譯任何字串。"
		Msg02 = "翻譯完成，其中："
		Msg03 = "已翻譯 " & tCount & " 個，已跳過 " & sCount & " 個，" & _
				"未變更 " & nCount & " 個，有錯誤 " & eCount & " 個"
	Else
		Msg01 = "楹祒俇傖ㄛ羶衄楹祒�庥拵硒恣�"
		Msg02 = "楹祒俇傖ㄛむ笢ㄩ"
		Msg03 = "眒楹祒 " & tCount & " 跺ㄛ眒泐徹 " & sCount & " 跺ㄛ" & _
				"帤載蜊 " & nCount & " 跺ㄛ衄渣昫 " & eCount & " 跺"
	End If
	TranCount = tCount + sCount + nCount
	If TranCount = 0 Then TranMassage = Msg01
	If TranCount <> 0 Then TranMassage = Msg02 & Msg03
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
	MsgBox(Msg02 & CStr(sysError.Number) & ", " & sysError.Description,vbInformation,Msg01)
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
			ConvCode = Mid(ConvertB,i+2,4)
			Stemp = CheckStr(UCase(ConvCode),"48-57,65-70")
		ElseIf EscStr <> "" Then
			EscStr = "\"
			ConvCode = Mid(ConvertB,i+1,3)
			Stemp = CheckStr(ConvCode,"48-55")
		End If

		If Stemp = True Then
			If EscStr = "\x" Then ConvString = ChrW(Val("&H" & ConvCode))
			If LCase(EscStr) = "\u" Then ConvString = ChrW(Val("&H" & ConvCode))
			If EscStr = "\" Then ConvString = ChrW(Val("&O" & ConvCode))
			If ConvString <> "" Then
				ConvertB = Replace(ConvertB,EscStr & ConvCode,ConvString)
				i = 0
			End If
		End If

		i = InStr(i+1,ConvertB,"\")
		If i = 0 Then Exit Do
	Loop
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
Function Settings(EngineID As Integer,CheckID As Integer) As Integer
	Dim LngNameList() As String,SrcLngList() As String,TranLngList() As String
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

		Msg33 = "翻譯前要被替換的字元 (用 | 分隔替換前後的字元) (半形逗號分隔):"
		Msg34 = "翻譯後要被替換的字元 (用 | 分隔替換前後的字元) (半形逗號分隔):"

		Msg35 = "說明(&H)"
		Msg36 = "翻譯引擎"
		Msg37 = "字串處理"
		Msg38 = "引擎參數"
		Msg39 = "語言配對"
		Msg40 = "使用物件:"
		Msg41 = "引擎註冊 ID:"
		Msg42 = "引擎網址:"
		Msg43 = "傳送內容範本:"
		Msg44 = "資料傳送方式:"
		Msg45 = "同步方式:"
		Msg46 = "使用者名(可空):"
		Msg47 = "密碼(可空):"
		Msg48 = "指令集(可空):"
		Msg49 = "HTTP 頭和值:"
		Msg50 = "(分行輸入)"
		Msg51 = "返回結果格式:"
		Msg52 = "翻譯開始自:"
		Msg53 = "翻譯結束到:"
		Msg54 = ">"
		Msg55 = "..."

		Msg56 = "語言配對"
		Msg57 = "語言名稱"
		Msg58 = "Passolo 代碼"
		Msg59 = "翻譯引擎代碼"
		Msg60 = "新增(&A)"
		Msg61 = "刪除(&D)"
		Msg62 = "全部刪除"
		Msg63 = "編輯(&E)"
		Msg64 = "外部編輯"
		Msg65 = "置空(&N)"
		Msg66 = "重設(&R)"
		Msg67 = "顯示非空項"
		Msg68 = "顯示空項"
		Msg69 = "全部顯示"

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

		Tools00 = "內置程式(&E)"
		Tools01 = "記事本(&N)"
		Tools02 = "Excel(&E)"
		Tools03 = "自訂程式(&C)"
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

		Msg33 = "楹祒ヶ猁掩杸遙腔趼睫 (蚚 | 煦路杸遙ヶ綴腔趼睫) (圉褒飯瘍煦路):"
		Msg34 = "楹祒綴猁掩杸遙腔趼睫 (蚚 | 煦路杸遙ヶ綴腔趼睫) (圉褒飯瘍煦路):"

		Msg35 = "堆翑(&H)"
		Msg36 = "楹祒竘э"
		Msg37 = "趼揹揭燴"
		Msg38 = "竘э統杅"
		Msg39 = "逄晟饜勤"
		Msg40 = "妏蚚勤砓:"
		Msg41 = "竘э蛁聊 ID:"
		Msg42 = "竘э厙硊:"
		Msg43 = "楷冞囀�暊ㄟ�:"
		Msg44 = "杅擂換冞源宒:"
		Msg45 = "肮祭源宒:"
		Msg46 = "蚚誧靡(褫諾):"
		Msg47 = "躇鎢(褫諾):"
		Msg48 = "硌鍔摩(褫諾):"
		Msg49 = "HTTP 芛睿硉:"
		Msg50 = "(煦俴怀��)"
		Msg51 = "殿隙賦彆跡宒:"
		Msg52 = "楹祒羲宎赻:"
		Msg53 = "楹祒賦旰善:"
		Msg54 = ">"
		Msg55 = "..."

		Msg56 = "逄晟饜勤"
		Msg57 = "逄晟靡備"
		Msg58 = "Passolo 測鎢"
		Msg59 = "楹祒竘э測鎢"
		Msg60 = "氝樓(&A)"
		Msg61 = "刉壺(&D)"
		Msg62 = "�垓褫噫�"
		Msg63 = "晤憮(&E)"
		Msg64 = "俋窒晤憮"
		Msg65 = "离諾(&N)"
		Msg66 = "笭离(&R)"
		Msg67 = "珆尨準諾砐"
		Msg68 = "珆尨諾砐"
		Msg69 = "�垓諫埰�"

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

		Tools00 = "囀离最唗(&E)"
		Tools01 = "暮岈掛(&N)"
		Tools02 = "Excel(&E)"
		Tools03 = "赻隅砱最唗(&C)"
	End If

	If AppNames(3) = "" Then
		ReDim AppNames(3),AppPaths(3)
		AppNames(0) = Tools00
		AppNames(1) = Tools01
   		AppNames(2) = Tools02
		AppNames(3) = Tools03
	   	AppPaths(0) = ""
  		AppPaths(1) = "Notepad.exe"
		AppPaths(2) = "Excel.exe"
  		AppPaths(3) = ""
  	End If

	Begin Dialog UserDialog 620,462,Msg01,.SetFunc ' %GRID:10,7,1,1
		Text 20,7,580,14,Msg02
		OptionGroup .Options
			OptionButton 130,28,130,14,Msg36
			OptionButton 270,28,130,14,Msg37
			OptionButton 410,28,130,14,Msg84

		GroupBox 20,49,330,70,Msg11,.GroupBox1
		DropListBox 40,66,290,21,EngineList(),.EngineList
		DropListBox 40,66,260,21,CheckList(),.CheckList
		PushButton 300,66,30,21,Msg55,.LevelButton
		PushButton 40,91,90,21,Msg06,.AddButton
		PushButton 140,91,90,21,Msg07,.ChangButton
		PushButton 240,91,90,21,Msg08,.DelButton

		GroupBox 370,49,230,70,Msg12,.GroupBox2
		OptionGroup .tWriteType
			OptionButton 390,69,90,14,Msg14,.tWriteToFile
			OptionButton 490,69,90,14,Msg15,.tWriteToRegistry
		OptionGroup .cWriteType
			OptionButton 390,69,90,14,Msg14,.cWriteToFile
			OptionButton 490,69,90,14,Msg15,.cWriteToRegistry
		PushButton 390,91,90,21,Msg16,.ImportButton
		PushButton 490,91,90,21,Msg17,.ExportButton

		GroupBox 20,147,580,280,Msg13,.GroupBox4
		OptionGroup .Engine
			OptionButton 160,129,180,14,Msg38
			OptionButton 350,129,180,14,Msg39
		Text 40,168,120,14,Msg40,.ObjectNameText
		Text 40,189,120,14,Msg41,.AppIDText
		Text 40,210,120,14,Msg42,.UrlText
		Text 40,231,120,14,Msg43,.UrlTemplateText
		Text 40,252,120,14,Msg44,.bstrMethodText
		Text 320,252,100,14,Msg45,.varAsyncText
		Text 40,273,120,14,Msg46,.bstrUserText
		Text 320,273,100,14,Msg47,.bstrPasswordText
		Text 40,294,120,14,Msg48,.varBodyText
		Text 40,315,120,14,Msg49,.setRequestHeaderText
		Text 40,329,120,28,Msg50,.setRequestHeaderText2
		Text 40,357,120,14,Msg51,.responseTypeText
		Text 40,378,120,14,Msg52,.TranBeforeStrText
		Text 40,399,120,14,Msg53,.TranAfterStrText
		TextBox 170,165,410,21,.ObjectNameBox
		TextBox 170,186,410,21,.AppIDBox
		TextBox 170,207,410,21,.UrlBox
		TextBox 170,228,380,21,.UrlTemplateBox
		TextBox 170,250,110,21,.bstrMethodBox
		TextBox 430,250,120,21,.varAsyncBox
		TextBox 170,272,140,21,.bstrUserBox
		TextBox 430,272,150,21,.bstrPasswordBox,-1
		TextBox 170,293,380,21,.varBodyBox
		TextBox 170,314,380,39,.setRequestHeaderBox,1
		TextBox 170,354,380,21,.responseTypeBox
		TextBox 170,375,380,21,.TranBeforeStrBox
		TextBox 170,396,380,21,.TranAfterStrBox
		PushButton 550,228,30,21,Msg54,.UrlTemplateButton
		PushButton 280,250,30,21,Msg54,.bstrMethodButton
		PushButton 550,250,30,21,Msg54,.varAsyncButton
		PushButton 550,293,30,21,Msg54,.varBodyButton
		PushButton 550,314,30,21,Msg54,.RequestButton
		PushButton 550,354,30,21,Msg54,.responseTypeButton
		PushButton 550,375,30,21,Msg55,.TranBeforeStrButton
		PushButton 550,396,30,21,Msg55,.TranAfterStrButton

		Text 40,161,190,14,Msg57,.LngNameText
		Text 240,161,110,14,Msg58,.SrcLngText
		Text 360,161,110,14,Msg59,.TranLngText
		ListBox 40,175,190,245,LngNameList(),.LngNameList
		ListBox 240,175,110,245,SrcLngList(),.SrcLngList
		ListBox 360,175,110,245,TranLngList(),.TranLngList
		PushButton 480,175,100,21,Msg60,.AddLngButton
		PushButton 480,196,100,21,Msg61,.DelLngButton
		PushButton 480,217,100,21,Msg62,.DelAllButton
		PushButton 480,252,100,21,Msg63,.EditLngButton
		PushButton 480,273,100,21,Msg64,.ExtEditButton
		PushButton 480,294,100,21,Msg65,.NullLngButton
		PushButton 480,315,100,21,Msg66,.ResetLngButton
		PushButton 480,350,100,21,Msg67,.ShowNoNullLngButton
		PushButton 480,371,100,21,Msg68,.ShowNullLngButton
		PushButton 480,392,100,21,Msg69,.ShowAllLngButton

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
		TextBox 40,315,540,98,.AutoWebFlagBox,1

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
	dlg.EngineList = EngineID
	dlg.CheckList = CheckID
	If Dialog(dlg) = 0 Then Exit Function
	Settings = dlg.EngineList
End Function


'③昢斛脤艘勤趕遺堆翑翋枙眕賸賤載嗣陓洘﹝
Private Function SetFunc%(DlgItem$, Action%, SuppValue&)
	Dim Header As String,HeaderID As Integer,NewData As String,Path As String,cStemp As Boolean
	Dim i As Integer,n As Integer,TempArray() As String,Temp As String,tStemp As Boolean
	Dim LngName As String,LngID As Integer,SrcLngCode As String,TranLngCode As String
	Dim LngNameList() As String,SrcLngList() As String,TranLngList() As String
	Dim AppLngList() As String,UseLngList() As String,LangArray() As String

	If OSLanguage = "0404" Then
		Msg01 = "錯誤"
		Msg02 = "預設值"
		Msg03 = "原值"
		Msg04 = "參照值"
		Msg08 = "未知"
		Msg11 = "警告"
		Msg12 = "如果某些參數為空，將使程式執行結果不正確。" & vbCrLf & _
				"您確實想要這樣做嗎？"
		Msg13 = "設定內容已經變更但是沒有儲存！" & vbCrLf & "是否需要儲存？"
		Msg14 = "儲存類型已經變更但是沒有儲存！" & vbCrLf & "是否需要儲存？"
		Msg18 = "目前設定中，至少有一項參數為空！" & vbCrLf
		Msg19 = "所有設定中，至少有一項參數為空！" & vbCrLf
		Msg21 = "確認"
		Msg22 = "確實要刪除設定「%s」嗎？"
		Msg24 = "確實要刪除語言「%s」嗎？"
		Msg25 = "翻譯引擎的"
		Msg26 = "字串處理的"
		Msg27 = "翻譯引擎和字串處理的"
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
		Msg45 = "確實要刪除全部語言嗎？"
		Msg50 = "翻譯開始自:"
		Msg51 = "翻譯結束到:"
		Msg52 = "按 ID 搜索:"
		Msg53 = "按標籤名搜索:"
		Msg54 = "可用語言:"
		Msg55 = "適用語言:"
		Msg60 = "選取解壓程式"
		Msg61 = "可執行檔案 (*.exe)|*.exe|所有檔案 (*.*)|*.*||"
		Msg62 = "沒有指定解壓程式！請重新輸入或選取。"
		Msg63 = "檔案參照參數(%1)"
		Msg64 = "要擷取的檔案參數(%2)"
		Msg65 = "解壓路徑參數(%3)"
		Item0 = "翻譯引擎網址 {Url}"
		Item1 = "翻譯引擎註冊號 {AppId}"
		Item2 = "要翻譯的文字 {Text}"
		Item3 = "來源語言 {From}"
		Item4 = "目標語言 {To}"
		Item5 = "同步執行(預設) {True}"
		Item6 = "異步執行 {False}"
		Item7 = "無符號整數陣列 {responseBody}"
		Item8 = "ADO Stream 物件 {responseStream}"
		Item9 = "字串 {responseText}"
		Item10 = "XML 格式資料 {responseXML}"
	Else
		Msg01 = "渣昫"
		Msg02 = "蘇�珋�"
		Msg03 = "埻硉"
		Msg04 = "統桽硉"
		Msg08 = "帤眭"
		Msg11 = "劑豢"
		Msg12 = "�蝜�議虳統杅峈諾ㄛ蔚妏最唗堍俴賦彆祥淏�楚�" & vbCrLf & _
				"蠟�滔舜遻肪瑵�酕鎘ˋ"
		Msg13 = "饜离囀�椹挩飛�蜊筍岆羶衄悵湔ㄐ" & vbCrLf & "岆瘁剒猁悵湔ˋ"
		Msg14 = "悵湔濬倰眒冪載蜊筍岆羶衄悵湔ㄐ" & vbCrLf & "岆瘁剒猁悵湔ˋ"
		Msg18 = "絞ヶ饜离笢ㄛ祫屾衄珨砐統杅峈諾ㄐ" & vbCrLf
		Msg19 = "垀衄饜离笢ㄛ祫屾衄珨砐統杅峈諾ㄐ" & vbCrLf
		Msg21 = "�溜�"
		Msg22 = "�滔菸罔噫�饜离※%s§鎘ˋ"
		Msg24 = "�滔菸罔噫�逄晟※%s§鎘ˋ"
		Msg25 = "楹祒竘э腔"
		Msg26 = "趼揹揭燴腔"
		Msg27 = "楹祒竘э睿趼揹揭燴腔"
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
		Msg45 = "�滔菸罔噫��垓諧擿堎艞�"
		Msg50 = "楹祒羲宎赻:"
		Msg51 = "楹祒賦旰善:"
		Msg52 = "偌 ID 刲坰:"
		Msg53 = "偌梓ワ靡刲坰:"
		Msg54 = "褫蚚逄晟:"
		Msg55 = "巠蚚逄晟:"
		Msg60 = "恁寁賤揤最唗"
		Msg61 = "褫硒俴恅璃 (*.exe)|*.exe|垀衄恅璃 (*.*)|*.*||"
		Msg62 = "羶衄硌隅賤揤最唗ㄐ③笭陔怀�趥藘√鞢�"
		Msg63 = "恅璃竘蚚統杅(%1)"
		Msg64 = "猁枑�△鰓躁�統杅(%2)"
		Msg65 = "賤揤繚噤統杅(%3)"
		Item0 = "楹祒竘э厙硊 {Url}"
		Item1 = "楹祒竘э蛁聊瘍 {AppId}"
		Item2 = "猁楹祒腔恅掛 {Text}"
		Item3 = "懂埭逄晟 {From}"
		Item4 = "醴梓逄晟 {To}"
		Item5 = "肮祭硒俴(蘇��) {True}"
		Item6 = "祑祭硒俴 {False}"
		Item7 = "拸睫瘍淕杅杅郪 {responseBody}"
		Item8 = "ADO Stream 勤砓 {responseStream}"
		Item9 = "趼睫揹 {responseText}"
		Item10 = "XML 跡宒杅擂 {responseXML}"
	End If

	If DlgValue("Options") = 0 Then
		DlgVisible "GroupBox1",True
		DlgVisible "EngineList",True
		DlgVisible "Engine",True
		DlgVisible "AddButton",True
		DlgVisible "ChangButton",True
		DlgVisible "DelButton",True

		DlgVisible "GroupBox2",True
		DlgVisible "tWriteType",True
		DlgVisible "ImportButton",True
		DlgVisible "ExportButton",True

		DlgVisible "GroupBox4",True
		DlgEnable "ObjectNameBox",False
		If DlgValue("Engine") = 0 Then
			DlgVisible "ObjectNameText",True
			DlgVisible "AppIDText",True
			DlgVisible "UrlText",True
			DlgVisible "UrlTemplateText",True
			DlgVisible "bstrMethodText",True
			DlgVisible "varAsyncText",True
			DlgVisible "bstrUserText",True
			DlgVisible "bstrPasswordText",True
			DlgVisible "varBodyText",True
			DlgVisible "setRequestHeaderText",True
			DlgVisible "setRequestHeaderText2",True
			DlgVisible "responseTypeText",True
			DlgVisible "TranBeforeStrText",True
			DlgVisible "TranAfterStrText",True
			DlgVisible "ObjectNameBox",True
			DlgVisible "AppIDBox",True
			DlgVisible "UrlBox",True
			DlgVisible "UrlTemplateBox",True
			DlgVisible "bstrMethodBox",True
			DlgVisible "varAsyncBox",True
			DlgVisible "bstrUserBox",True
			DlgVisible "bstrPasswordBox",True
			DlgVisible "varBodyBox",True
			DlgVisible "setRequestHeaderBox",True
			DlgVisible "responseTypeBox",True
			DlgVisible "TranBeforeStrBox",True
			DlgVisible "TranAfterStrBox",True
			DlgVisible "UrlTemplateButton",True
			DlgVisible "bstrMethodButton",True
			DlgVisible "varAsyncButton",True
			DlgVisible "varBodyButton",True
			DlgVisible "RequestButton",True
			DlgVisible "responseTypeButton",True
			DlgVisible "TranBeforeStrButton",True
			DlgVisible "TranAfterStrButton",True

			DlgVisible "LngNameText",False
			DlgVisible "SrcLngText",False
			DlgVisible "TranLngText",False
			DlgVisible "LngNameList",False
			DlgVisible "SrcLngList",False
			DlgVisible "TranLngList",False
			DlgVisible "AddLngButton",False
			DlgVisible "DelLngButton",False
			DlgVisible "DelAllButton",False
			DlgVisible "EditLngButton",False
			DlgVisible "ExtEditButton",False
			DlgVisible "NullLngButton",False
			DlgVisible "ResetLngButton",False
			DlgVisible "ShowNoNullLngButton",False
			DlgVisible "ShowNullLngButton",False
			DlgVisible "ShowAllLngButton",False
		Else
			DlgVisible "ObjectNameText",False
			DlgVisible "AppIDText",False
			DlgVisible "UrlText",False
			DlgVisible "UrlTemplateText",False
			DlgVisible "bstrMethodText",False
			DlgVisible "varAsyncText",False
			DlgVisible "bstrUserText",False
			DlgVisible "bstrPasswordText",False
			DlgVisible "varBodyText",False
			DlgVisible "setRequestHeaderText",False
			DlgVisible "setRequestHeaderText2",False
			DlgVisible "responseTypeText",False
			DlgVisible "TranBeforeStrText",False
			DlgVisible "TranAfterStrText",False
			DlgVisible "ObjectNameBox",False
			DlgVisible "AppIDBox",False
			DlgVisible "UrlBox",False
			DlgVisible "UrlTemplateBox",False
			DlgVisible "bstrMethodBox",False
			DlgVisible "varAsyncBox",False
			DlgVisible "bstrUserBox",False
			DlgVisible "bstrPasswordBox",False
			DlgVisible "varBodyBox",False
			DlgVisible "setRequestHeaderBox",False
			DlgVisible "responseTypeBox",False
			DlgVisible "TranBeforeStrBox",False
			DlgVisible "TranAfterStrBox",False
			DlgVisible "UrlTemplateButton",False
			DlgVisible "bstrMethodButton",False
			DlgVisible "varAsyncButton",False
			DlgVisible "varBodyButton",False
			DlgVisible "RequestButton",False
			DlgVisible "responseTypeButton",False
			DlgVisible "TranBeforeStrButton",False
			DlgVisible "TranAfterStrButton",False

			DlgVisible "LngNameText",True
			DlgVisible "SrcLngText",True
			DlgVisible "TranLngText",True
			DlgVisible "LngNameList",True
			DlgVisible "SrcLngList",True
			DlgVisible "TranLngList",True
			DlgVisible "AddLngButton",True
			DlgVisible "DelLngButton",True
			DlgVisible "DelAllButton",True
			DlgVisible "EditLngButton",True
			DlgVisible "ExtEditButton",True
			DlgVisible "NullLngButton",True
			DlgVisible "ResetLngButton",True
			DlgVisible "ShowNoNullLngButton",True
			DlgVisible "ShowNullLngButton",True
			DlgVisible "ShowAllLngButton",True
		End If
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
		DlgVisible "GroupBox1",True
		DlgVisible "EngineList",False
		DlgVisible "Engine",False
		DlgVisible "AddButton",True
		DlgVisible "ChangButton",True
		DlgVisible "DelButton",True

		DlgVisible "GroupBox2",True
		DlgVisible "tWriteType",False
		DlgVisible "ImportButton",True
		DlgVisible "ExportButton",True

		DlgVisible "GroupBox4",False
		DlgVisible "ObjectNameText",False
		DlgVisible "AppIDText",False
		DlgVisible "UrlText",False
		DlgVisible "UrlTemplateText",False
		DlgVisible "bstrMethodText",False
		DlgVisible "varAsyncText",False
		DlgVisible "bstrUserText",False
		DlgVisible "bstrPasswordText",False
		DlgVisible "varBodyText",False
		DlgVisible "setRequestHeaderText",False
		DlgVisible "setRequestHeaderText2",False
		DlgVisible "responseTypeText",False
		DlgVisible "TranBeforeStrText",False
		DlgVisible "TranAfterStrText",False
		DlgVisible "ObjectNameBox",False
		DlgVisible "AppIDBox",False
		DlgVisible "UrlBox",False
		DlgVisible "UrlTemplateBox",False
		DlgVisible "bstrMethodBox",False
		DlgVisible "varAsyncBox",False
		DlgVisible "bstrUserBox",False
		DlgVisible "bstrPasswordBox",False
		DlgVisible "varBodyBox",False
		DlgVisible "setRequestHeaderBox",False
		DlgVisible "responseTypeBox",False
		DlgVisible "TranBeforeStrBox",False
		DlgVisible "TranAfterStrBox",False
		DlgVisible "UrlTemplateButton",False
		DlgVisible "bstrMethodButton",False
		DlgVisible "varAsyncButton",False
		DlgVisible "varBodyButton",False
		DlgVisible "RequestButton",False
		DlgVisible "responseTypeButton",False
		DlgVisible "TranBeforeStrButton",False
		DlgVisible "TranAfterStrButton",False

		DlgVisible "LngNameText",False
		DlgVisible "SrcLngText",False
		DlgVisible "TranLngText",False
		DlgVisible "LngNameList",False
		DlgVisible "SrcLngList",False
		DlgVisible "TranLngList",False
		DlgVisible "AddLngButton",False
		DlgVisible "DelLngButton",False
		DlgVisible "DelAllButton",False
		DlgVisible "EditLngButton",False
		DlgVisible "ExtEditButton",False
		DlgVisible "NullLngButton",False
		DlgVisible "ResetLngButton",False
		DlgVisible "ShowNoNullLngButton",False
		DlgVisible "ShowNullLngButton",False
		DlgVisible "ShowAllLngButton",False

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
			DlgVisible "AutoWebFlagBoxTxt",True
			DlgVisible "PreRepStrBox",True
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
	ElseIf DlgValue("Options") = 2 Then
		DlgVisible "GroupBox1",False
		DlgVisible "EngineList",False
		DlgVisible "Engine",False
		DlgVisible "AddButton",False
		DlgVisible "ChangButton",False
		DlgVisible "DelButton",False

		DlgVisible "GroupBox2",False
		DlgVisible "tWriteType",False
		DlgVisible "ImportButton",False
		DlgVisible "ExportButton",False

		DlgVisible "GroupBox4",False
		DlgEnable "ObjectNameBox",False
		DlgVisible "ObjectNameText",False
		DlgVisible "AppIDText",False
		DlgVisible "UrlText",False
		DlgVisible "UrlTemplateText",False
		DlgVisible "bstrMethodText",False
		DlgVisible "varAsyncText",False
		DlgVisible "bstrUserText",False
		DlgVisible "bstrPasswordText",False
		DlgVisible "varBodyText",False
		DlgVisible "setRequestHeaderText",False
		DlgVisible "setRequestHeaderText2",False
		DlgVisible "responseTypeText",False
		DlgVisible "TranBeforeStrText",False
		DlgVisible "TranAfterStrText",False

		DlgVisible "LngNameText",False
		DlgVisible "SrcLngText",False
		DlgVisible "TranLngText",False
		DlgVisible "LngNameList",False
		DlgVisible "SrcLngList",False
		DlgVisible "TranLngList",False
		DlgVisible "AddLngButton",False
		DlgVisible "DelLngButton",False
		DlgVisible "DelAllButton",False
		DlgVisible "EditLngButton",False
		DlgVisible "ExtEditButton",False
		DlgVisible "NullLngButton",False
		DlgVisible "ResetLngButton",False
		DlgVisible "ShowNoNullLngButton",False
		DlgVisible "ShowNullLngButton",False
		DlgVisible "ShowAllLngButton",False

		DlgVisible "ObjectNameBox",False
		DlgVisible "AppIDBox",False
		DlgVisible "UrlBox",False
		DlgVisible "UrlTemplateBox",False
		DlgVisible "bstrMethodBox",False
		DlgVisible "varAsyncBox",False
		DlgVisible "bstrUserBox",False
		DlgVisible "bstrPasswordBox",False
		DlgVisible "varBodyBox",False
		DlgVisible "setRequestHeaderBox",False
		DlgVisible "responseTypeBox",False
		DlgVisible "TranBeforeStrBox",False
		DlgVisible "TranAfterStrBox",False
		DlgVisible "UrlTemplateButton",False
		DlgVisible "bstrMethodButton",False
		DlgVisible "varAsyncButton",False
		DlgVisible "varBodyButton",False
		DlgVisible "RequestButton",False
		DlgVisible "responseTypeButton",False
		DlgVisible "TranBeforeStrButton",False
		DlgVisible "TranAfterStrButton",False

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
		If Join(EngineList) <> "" Then
			HeaderID = DlgValue("EngineList")
			TempArray = Split(EngineDataList(HeaderID),JoinStr)
			SetsArray = Split(TempArray(1),SubJoinStr)
			DlgText "ObjectNameBox",SetsArray(0)
			DlgText "AppIDBox",SetsArray(1)
			DlgText "UrlBox",SetsArray(2)
			DlgText "UrlTemplateBox",SetsArray(3)
			DlgText "bstrMethodBox",SetsArray(4)
			DlgText "varAsyncBox",SetsArray(5)
			DlgText "bstrUserBox",SetsArray(6)
			DlgText "bstrPasswordBox",SetsArray(7)
			DlgText "varBodyBox",SetsArray(8)
			DlgText "setRequestHeaderBox",SetsArray(9)
			DlgText "responseTypeBox",SetsArray(10)
			DlgText "TranBeforeStrBox",SetsArray(11)
			DlgText "TranAfterStrBox",SetsArray(12)
			If SetsArray(0) = "" Then DlgText "ObjectNameBox",DefaultObject
			SplitData(TempArray(2),LngNameList,SrcLngList,TranLngList)
			DlgListBoxArray "LngNameList",LngNameList()
			DlgListBoxArray "SrcLngList",SrcLngList()
			DlgListBoxArray "TranLngList",TranLngList()
			DlgValue "LngNameList",0
			DlgValue "SrcLngList",0
			DlgValue "TranLngList",0
		End If

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

		If tWriteLoc = EngineFilePath Then DlgValue "tWriteType",0
		If tWriteLoc = EngineRegKey Then DlgValue "tWriteType",1
		If tWriteLoc = "" Then DlgValue "tWriteType",0
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

		If Join(tUpdateSet) <> "" Then
			DlgValue "UpdateSet",StrToInteger(tUpdateSet(0))
			DlgText "WebSiteBox",tUpdateSet(1)
			DlgText "CmdPathBox",tUpdateSet(2)
			DlgText "ArgumentBox",tUpdateSet(3)
			DlgText "UpdateCycleBox",tUpdateSet(4)
			DlgText "UpdateDateBox",tUpdateSet(5)
		End If
		If DlgText("UpdateDateBox") = "" Then DlgText "UpdateDateBox",Msg08
		DlgEnable "UpdateDateBox",False

		If DlgValue("Options") = 0 Then
			Header = DlgText("EngineList")
			If InStr(Join(DefaultEngineList,JoinStr),Header) Then
				DlgEnable "ChangButton",False
				DlgEnable "DelButton",False
			Else
				If UBound(EngineList) = 0 Then DlgEnable "DelButton",False
			End If
			If DlgText("responseTypeBox") = "responseXML" Then
				DlgText "TranBeforeStrText",Msg52
				DlgText "TranAfterStrText",Msg53
			Else
				DlgText "TranBeforeStrText",Msg50
				DlgText "TranAfterStrText",Msg51
			End If
			DlgEnable "ShowAllLngButton",False
		ElseIf DlgValue("Options") = 1 Then
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
		End If
	Case 2 ' 杅硉載蜊麼氪偌狟賸偌聽
		If DlgValue("Options") = 0 Then
			HeaderID = DlgValue("EngineList")
			TempArray = Split(EngineDataList(HeaderID),JoinStr)
			SetsArray = Split(TempArray(1),SubJoinStr)
			SplitData(TempArray(2),LngNameList,SrcLngList,TranLngList)

			If DlgItem$ = "EngineList" Then
				DlgText "ObjectNameBox",SetsArray(0)
				DlgText "AppIDBox",SetsArray(1)
				DlgText "UrlBox",SetsArray(2)
				DlgText "UrlTemplateBox",SetsArray(3)
				DlgText "bstrMethodBox",SetsArray(4)
				DlgText "varAsyncBox",SetsArray(5)
				DlgText "bstrUserBox",SetsArray(6)
				DlgText "bstrPasswordBox",SetsArray(7)
				DlgText "varBodyBox",SetsArray(8)
				DlgText "setRequestHeaderBox",SetsArray(9)
				DlgText "responseTypeBox",SetsArray(10)
				DlgText "TranBeforeStrBox",SetsArray(11)
				DlgText "TranAfterStrBox",SetsArray(12)
				If SetsArray(0) = "" Then DlgText "ObjectNameBox",DefaultObject

				n = DlgValue("LngNameList")
				DlgListBoxArray "LngNameList",LngNameList()
				DlgListBoxArray "SrcLngList",SrcLngList()
				DlgListBoxArray "TranLngList",TranLngList()
				DlgValue "LngNameList",n
				DlgValue "SrcLngList",n
				DlgValue "TranLngList",n
				DlgEnable "ShowNoNullLngButton",True
				DlgEnable "ShowNullLngButton",True
				DlgEnable "ShowAllLngButton",False
			End If

			If DlgItem$ = "AddButton" Then
				NewData = AddSet(EngineList)
				If NewData <> "" Then
					For i = 0 To UBound(SetsArray)
						If i = 0 Then SetsArray(i) = DefaultObject
						If i <> 0 Then SetsArray(i) = ""
					Next i
					Data = Join(SetsArray,SubJoinStr)
					LangPairList = LangCodeList("",OSLanguage,0,107)
					Temp = NewData & JoinStr & Data & JoinStr & Join(LangPairList,SubJoinStr)
					CreateArray(NewData,Temp,EngineList,EngineDataList)
					DlgListBoxArray "EngineList",EngineList()
					DlgText "EngineList",NewData
					DlgText "ObjectNameBox",DefaultObject
					DlgText "AppIDBox",""
					DlgText "UrlBox",""
					DlgText "UrlTemplateBox",""
					DlgText "bstrMethodBox",""
					DlgText "varAsyncBox",""
					DlgText "bstrUserBox",""
					DlgText "bstrPasswordBox",""
					DlgText "varBodyBox",""
					DlgText "setRequestHeaderBox",""
					DlgText "responseTypeBox",""
					DlgText "TranBeforeStrBox",""
					DlgText "TranAfterStrBox",""
					HeaderID = DlgValue("EngineList")
					TempArray = Split(EngineDataList(HeaderID),JoinStr)
					SplitData(TempArray(2),LngNameList,SrcLngList,TranLngList)
					DlgListBoxArray "TranLngList",LngNameList()
					DlgListBoxArray "TranLngList",SrcLngList()
					DlgListBoxArray "TranLngList",TranLngList()
					DlgValue "LngNameList",0
					DlgValue "SrcLngList",0
					DlgValue "TranLngList",0
					DlgEnable "ResetLngButton",False
				End If
			End If

			If DlgItem$ = "ChangButton" Then
				Header = DlgText("EngineList")
				NewData = EditSet(EngineList,Header)
				If NewData <> "" Then
					EngineList(HeaderID) = NewData
					EngineDataList(HeaderID) = NewData & JoinStr & TempArray(1) & JoinStr & TempArray(2)
					DlgListBoxArray "EngineList",EngineList()
					DlgValue "EngineList",HeaderID
				End If
			End If

	    	If DlgItem$ = "DelButton" Then
				Header = DlgText("EngineList")
				Msg = Replace(Msg24,"%s",Header)
				If MsgBox(Msg,vbYesNo+vbInformation,Msg21) = vbYes Then
					If HeaderID > 0 And HeaderID = UBound(EngineList) Then HeaderID = HeaderID - 1
					EngineList = DelArray(Header,EngineList,"",0)
					EngineDataList = DelArray(Header,EngineDataList,JoinStr,0)
					DlgListBoxArray "EngineList",EngineList()
					DlgValue "EngineList",HeaderID

					TempArray = Split(EngineDataList(HeaderID),JoinStr)
					SetsArray = Split(TempArray(1),SubJoinStr)
					DlgText "ObjectNameBox",SetsArray(0)
					DlgText "AppIDBox",SetsArray(1)
					DlgText "UrlBox",SetsArray(2)
					DlgText "UrlTemplateBox",SetsArray(3)
					DlgText "bstrMethodBox",SetsArray(4)
					DlgText "varAsyncBox",SetsArray(5)
					DlgText "bstrUserBox",SetsArray(6)
					DlgText "bstrPasswordBox",SetsArray(7)
					DlgText "varBodyBox",SetsArray(8)
					DlgText "setRequestHeaderBox",SetsArray(9)
					DlgText "responseTypeBox",SetsArray(10)
					DlgText "TranBeforeStrBox",SetsArray(11)
					DlgText "TranAfterStrBox",SetsArray(12)
					If SetsArray(0) = "" Then DlgText "ObjectNameBox",DefaultObject

					n = DlgValue("LngNameList")
					SplitData(TempArray(2),LngNameList,SrcLngList,TranLngList)
					DlgListBoxArray "LngNameList",LngNameList()
					DlgListBoxArray "SrcLngList",SrcLngList()
					DlgListBoxArray "TranLngList",TranLngList()
					DlgValue "LngNameList",n
					DlgValue "SrcLngList",n
					DlgValue "TranLngList",n
					DlgEnable "ShowNoNullLngButton",True
					DlgEnable "ShowNullLngButton",True
					DlgEnable "ShowAllLngButton",False
				End If
			End If

			If DlgItem$ = "ResetButton" Then
				Header = DlgText("EngineList")
				ReDim TempArray(1)
				If InStr(Join(DefaultEngineList,JoinStr),Header) Then TempArray(0) = Msg02
				tStemp = CheckNullData(Header,EngineDataListBak,"0,1,5,6,7,8",0)
				If tStemp = False Then TempArray(1) = Msg03
				For i = LBound(EngineList) To UBound(EngineList)
					If i <> HeaderID Then
						ReDim Preserve TempArray(i+2)
						TempArray(i+2) = Msg04 & " - " & EngineList(i)
					End If
				Next i
				i = ShowPopupMenu(TempArray,vbPopupUseRightButton)
				If i = 0 Then
					SetsArray = Split(EngineSettings(Header),SubJoinStr)
					TempArray = LangCodeList(Header,OSLanguage,0,107)
					Temp = Join(TempArray,SubJoinStr)
				ElseIf i = 1 Then
					For n = LBound(EngineDataListBak) To UBound(EngineDataListBak)
						TempArray = Split(EngineDataListBak(n),JoinStr)
						If TempArray(0) = Header Then
							SetsArray = Split(TempArray(1),SubJoinStr)
							Temp = TempArray(2)
							Exit For
						End If
					Next n
				ElseIf i >= 2 Then
					For n = LBound(EngineList) To UBound(EngineList)
						Temp = Mid(TempArray(i),InStr(TempArray(i),Msg04 & " - ") + Len(Msg04 & " - "))
						If Temp = EngineList(n) Then
							HeaderID = n
							Exit For
						End If
					Next n
					TempArray = Split(EngineDataList(HeaderID),JoinStr)
					SetsArray = Split(TempArray(1),SubJoinStr)
					Temp = TempArray(2)
				End If
				If i >= 0 Then
					If DlgValue("Engine") = 0 Then
						DlgText "ObjectNameBox",SetsArray(0)
						DlgText "AppIDBox",SetsArray(1)
						DlgText "UrlBox",SetsArray(2)
						DlgText "UrlTemplateBox",SetsArray(3)
						DlgText "bstrMethodBox",SetsArray(4)
						DlgText "varAsyncBox",SetsArray(5)
						DlgText "bstrUserBox",SetsArray(6)
						DlgText "bstrPasswordBox",SetsArray(7)
						DlgText "varBodyBox",SetsArray(8)
						DlgText "setRequestHeaderBox",SetsArray(9)
						DlgText "responseTypeBox",SetsArray(10)
						DlgText "TranBeforeStrBox",SetsArray(11)
						DlgText "TranAfterStrBox",SetsArray(12)
						If SetsArray(0) = "" Then DlgText "ObjectNameBox",DefaultObject
					ElseIf DlgValue("Engine") <> 0 Then
						n = DlgValue("LngNameList")
						SplitData(Temp,LngNameList,SrcLngList,TranLngList)
						DlgListBoxArray "LngNameList",LngNameList()
						DlgListBoxArray "SrcLngList",SrcLngList()
						DlgListBoxArray "TranLngList",TranLngList()
						DlgValue "LngNameList",n
						DlgValue "SrcLngList",n
						DlgValue "TranLngList",n
						DlgEnable "ShowNoNullLngButton",True
						DlgEnable "ShowNullLngButton",True
						DlgEnable "ShowAllLngButton",False
					End If
				End If
			End If

	    	If DlgItem$ = "CleanButton" Then
				If DlgValue("Engine") = 0 Then
					'DlgText "ObjectNameBox",""
					DlgText "AppIDBox",""
					DlgText "UrlBox",""
					DlgText "UrlTemplateBox",""
					DlgText "bstrMethodBox",""
					DlgText "varAsyncBox",""
					DlgText "bstrUserBox",""
					DlgText "bstrPasswordBox",""
					DlgText "varBodyBox",""
					DlgText "setRequestHeaderBox",""
					DlgText "responseTypeBox",""
					DlgText "TranBeforeStrBox",""
					DlgText "TranAfterStrBox",""
				Else
					n = DlgValue("TranLngList")
					For i = 0 To UBound(LngNameList)
						TranLngList(i) = NullValue
					Next i
					DlgListBoxArray "TranLngList",TranLngList()
					DlgValue "TranLngList",n
				End If
			End If

			If DlgItem$ = "LngNameList" Or DlgItem$ = "SrcLngList" Or DlgItem$ = "TranLngList" Then
				If DlgItem$ = "LngNameList" Then
					n = DlgValue("LngNameList")
					DlgValue "SrcLngList",n
					DlgValue "TranLngList",n
				End If
				If DlgItem$ = "SrcLngList" Then
					n = DlgValue("SrcLngList")
					DlgValue "LngNameList",n
					DlgValue "TranLngList",n
				End If
				If DlgItem$ = "TranLngList" Then
					n = DlgValue("TranLngList")
					DlgValue "LngNameList",n
					DlgValue "SrcLngList",n
				End If
			End If

			If DlgItem$ = "AddLngButton" Or DlgItem$ = "DelLngButton" Or DlgItem$ = "DelAllButton" Then
				If DlgItem$ = "AddLngButton" Then
					LangArray = Split(TempArray(2),SubJoinStr)
					NewData = EditLang(LangArray,"","","")
					If NewData <> "" Then
						n = UBound(LngNameList) + 1
						ReDim Preserve LngNameList(n),SrcLngList(n),TranLngList(n)
						LangPairList = Split(NewData,LngJoinStr)
						LngNameList(n) = LangPairList(0)
						SrcLngList(n) = LangPairList(1)
						TranLngList(n) = LangPairList(2)
					End If
				End If
				If DlgItem$ = "DelLngButton" Then
					LngName = DlgText("LngNameList")
					If LngName <> "" Then
						Msg = Replace(Msg24,"%s",LngName)
						If MsgBox(Msg,vbYesNo+vbInformation,Msg21) = vbYes Then
							n = DlgValue("LngNameList")
							Header = DlgText("EngineList")
							NewData = DlgText("LngNameList")
							LangArray = Split(TempArray(2),SubJoinStr)
							If n > 0 And n = UBound(LangArray) Then n = n - 1
							LangArray = DelArray(NewData,LangArray,LngJoinStr,0)
							SplitData(Join(LangArray,SubJoinStr),LngNameList,SrcLngList,TranLngList)
						End If
					End If
				End If
				If DlgItem$ = "DelAllButton" Then
					If MsgBox(Msg45,vbYesNo+vbInformation,Msg21) = vbYes Then
						n = 0
						NewData = DlgText("LngNameList")
						ReDim LngNameList(0),SrcLngList(0),TranLngList(0)
						DlgEnable "ShowNoNullLngButton",False
						DlgEnable "ShowNullLngButton",False
						DlgEnable "ShowAllLngButton",False
					End If
				End If
				If NewData <> "" Then
					DlgListBoxArray "LngNameList",LngNameList()
					DlgListBoxArray "SrcLngList",SrcLngList()
					DlgListBoxArray "TranLngList",TranLngList()
					DlgValue "LngNameList",n
					DlgValue "SrcLngList",n
					DlgValue "TranLngList",n
					DlgEnable "ShowNoNullLngButton",True
					DlgEnable "ShowNullLngButton",True
					DlgEnable "ShowAllLngButton",False
				End If
			End If

			If DlgItem$ = "EditLngButton" Or DlgItem$ = "NullLngButton" Or DlgItem$ = "ResetLngButton" Then
				Header = DlgText("EngineList")
				LngName = DlgText("LngNameList")
				SrcLngCode = DlgText("SrcLngList")
				TranLngCode = DlgText("TranLngList")

				If DlgItem$ = "EditLngButton" Then
					LangArray = Split(TempArray(2),SubJoinStr)
					NewData = EditLang(LangArray,LngName,SrcLngCode,TranLngCode)
					If NewData <> "" Then
						LangPairList = Split(NewData,LngJoinStr)
						NewLngName = LangPairList(0)
						NewSrcLngCode = LangPairList(1)
						NewTranLngCode = LangPairList(2)
					End If
				End If
				If DlgItem$ = "NullLngButton" Then
					If TranLngCode <> NullValue Then
						NewData = NullValue
						NewLngName = LngName
						NewSrcLngCode = SrcLngCode
						NewTranLngCode = NullValue
					End If
				End If
				If DlgItem$ = "ResetLngButton" Then
					For i = LBound(EngineDataListBak) To UBound(EngineDataListBak)
						TempArray = Split(EngineDataListBak(i),JoinStr)
						If TempArray(0) = Header Then
							LangArray = Split(TempArray(2),SubJoinStr)
							For n = 0 To UBound(LangArray)
								LangPairList = Split(LangArray(n),LngJoinStr)
								If LangPairList(0) = LngName Then
									NewData = NullValue
									NewLngName = LangPairList(0)
									NewSrcLngCode = LangPairList(1)
									NewTranLngCode = LangPairList(2)
									Exit For
								End If
							Next n
							Exit For
						End If
					Next i
				End If

				If NewData <> "" Then
					If NewSrcLngCode = "" Then NewSrcLngCode = NullValue
					If NewTranLngCode = "" Then NewTranLngCode = NullValue
					For i = 0 To UBound(LngNameList)
						If LngNameList(i) = LngName Then
							LngNameList(i) = NewLngName
							SrcLngList(i) = NewSrcLngCode
							TranLngList(i) = NewTranLngCode
							Exit For
						End If
					Next i

					n = DlgValue("LngNameList")
					If DlgEnable("ShowAllLngButton") = False Then
						If n < UBound(LngNameList) Then n = n + 1
						DlgListBoxArray "LngNameList",LngNameList()
						DlgListBoxArray "SrcLngList",SrcLngList()
						DlgListBoxArray "TranLngList",TranLngList()
					Else
						For i = 0 To UBound(DelLngNameList)
							If DelLngNameList(i) = LngName Then
								DelLngNameList(i) = NewLngName
								DelSrcLngList(i) = NewSrcLngCode
								DelTranLngList(i) = NewTranLngCode
								Exit For
							End If
						Next i
						If n < UBound(DelLngNameList) Then n = n + 1
						DlgListBoxArray "LngNameList",DelLngNameList()
						DlgListBoxArray "SrcLngList",DelSrcLngList()
						DlgListBoxArray "TranLngList",DelTranLngList()
					End If
					DlgValue "LngNameList",n
					DlgValue "SrcLngList",n
					DlgValue "TranLngList",n
				End If
			End If

			If DlgItem$ = "ShowNoNullLngButton" Or DlgItem$ = "ShowNullLngButton" Then
				n = 0
				For i = 0 To UBound(LngNameList)
					If DlgItem$ = "ShowNoNullLngButton" Then
						If TranLngList(i) <> NullValue Then
							ReDim Preserve DelLngNameList(n),DelSrcLngList(n),DelTranLngList(n)
							DelLngNameList(n) = LngNameList(i)
							DelSrcLngList(n) = SrcLngList(i)
							DelTranLngList(n) = TranLngList(i)
							n = n + 1
							DlgEnable "ShowNoNullLngButton",False
							DlgEnable "ShowNullLngButton",True
							DlgEnable "ShowAllLngButton",True
						End If
					Else
						If TranLngList(i) = NullValue Then
							ReDim Preserve DelLngNameList(n),DelSrcLngList(n),DelTranLngList(n)
							DelLngNameList(n) = LngNameList(i)
							DelSrcLngList(n) = SrcLngList(i)
							DelTranLngList(n) = TranLngList(i)
							n = n + 1
							DlgEnable "ShowNoNullLngButton",True
							DlgEnable "ShowNullLngButton",False
							DlgEnable "ShowAllLngButton",True
						End If
					End If
				Next i
				DlgListBoxArray "LngNameList",DelLngNameList()
				DlgListBoxArray "SrcLngList",DelSrcLngList()
				DlgListBoxArray "TranLngList",DelTranLngList()
				DlgValue "LngNameList",0
				DlgValue "SrcLngList",0
				DlgValue "TranLngList",0
			End If

			If DlgItem$ = "ShowAllLngButton" Then
				DlgListBoxArray "LngNameList",LngNameList()
				DlgListBoxArray "SrcLngList",SrcLngList()
				DlgListBoxArray "TranLngList",TranLngList()
				DlgValue "LngNameList",0
				DlgValue "SrcLngList",0
				DlgValue "TranLngList",0
				DlgEnable "ShowNoNullLngButton",True
				DlgEnable "ShowNullLngButton",True
				DlgEnable "ShowAllLngButton",False
			End If

			If DlgItem$ = "ExtEditButton" Then
				n = ShowPopupMenu(AppNames,vbPopupUseRightButton)
				If n >= 0 Then
					prjFolder = trn.Project.Location
					FilePath = prjFolder & "\~temp.txt"
					Code = "ANSI"
					For i = 0 To UBound(LngNameList)
						LngName = LngNameList(i)
						SrcLngCode = SrcLngList(i)
						TranLngCode = TranLngList(i)
						ReDim Preserve LangArray(i)
						LangArray(i) = LngName & vbTab & SrcLngCode & vbTab & TranLngCode
					Next i
					LngPair = Join(LangArray,vbCrLf)
					If WriteToFile(FilePath,LngPair,Code) = True Then
						If Dir(FilePath) <> "" Then
							ReDim FileList(0)
							If OpenFile(FilePath,FileList,n,True) = True Then
								textStr = ReadFile(FilePath,Code)
							End If
							On Error Resume Next
							Kill FilePath
							On Error GoTo 0
						End If
					End If
					If textStr <> "" And LngPair <> textStr Then
						ReDim LngNameList(0),SrcLngList(0),TranLngList(0)
						FileLines = Split(textStr,vbCrLf)
						For i = LBound(FileLines) To UBound(FileLines)
							ReDim Preserve LngNameList(i),SrcLngList(i),TranLngList(i)
							TempArray = Split(FileLines(i),vbTab)
							If TempArray(1) = "" Then TempArray(1) = NullValue
							If TempArray(2) = "" Then TempArray(2) = NullValue
							LngNameList(i) = TempArray(0)
							SrcLngList(i) = TempArray(1)
							TranLngList(i) = TempArray(2)
						Next i
						DlgListBoxArray "LngNameList",LngNameList()
						DlgListBoxArray "SrcLngList",SrcLngList()
						DlgListBoxArray "TranLngList",TranLngList()
						DlgValue "LngNameList",0
						DlgValue "SrcLngList",0
						DlgValue "TranLngList",0
						DlgEnable "ShowNoNullLngButton",True
						DlgEnable "ShowNullLngButton",True
						DlgEnable "ShowAllLngButton",False
					End If
				End If
			End If

			If DlgItem$ = "UrlTemplateButton" Or DlgItem$ = "varBodyButton" Or DlgItem$ = "RequestButton" Then
				ReDim TempArray(4)
				TempArray(0) = Item0
				TempArray(1) = Item1
				TempArray(2) = Item2
				TempArray(3) = Item3
				TempArray(4) = Item4
	  			i = ShowPopupMenu(TempArray,vbPopupUseRightButton)
				If i >= 0 Then
					Temp = TempArray(i)
					Temp = Mid(Temp,InStr(Temp,"{"))
					If DlgItem$ = "UrlTemplateButton" Then
						DlgText "UrlTemplateBox",DlgText("UrlTemplateBox") & Temp
					ElseIf DlgItem$ = "varBodyButton" Then
						DlgText "varBodyBox",DlgText("varBodyBox") & Temp
					ElseIf DlgItem$ = "RequestButton" Then
						DlgText "setRequestHeaderBox",DlgText("setRequestHeaderBox") & Temp
					End If
				End If
			End If

			If DlgItem$ = "bstrMethodButton" Then
				ReDim TempArray(1)
				TempArray(0) = "GET"
				TempArray(1) = "POST"
	  			i = ShowPopupMenu(TempArray,vbPopupUseRightButton)
				If i >= 0 Then
					Temp = TempArray(i)
					DlgText "bstrMethodBox",Temp
				End If
			End If

			If DlgItem$ = "varAsyncButton" Then
				ReDim TempArray(1)
				TempArray(0) = Item5
				TempArray(1) = Item6
		  		i = ShowPopupMenu(TempArray,vbPopupUseRightButton)
				If i >= 0 Then
					Temp = TempArray(i)
					Temp = Mid(Temp,InStr(Temp,"{")+1,InStr(Temp,"}")-InStr(Temp,"{")-1)
					DlgText "varAsyncBox",Temp
				End If
			End If

			If DlgItem$ = "responseTypeButton" Then
				ReDim TempArray(3)
				TempArray(0) = Item7
				TempArray(1) = Item8
				TempArray(2) = Item9
				TempArray(3) = Item10
		  		i = ShowPopupMenu(TempArray,vbPopupUseRightButton)
				If i >= 0 Then
					Temp = TempArray(i)
					Temp = Mid(Temp,InStr(Temp,"{")+1,InStr(Temp,"}")-InStr(Temp,"{")-1)
					DlgText "responseTypeBox",Temp
				End If
			End If

			If DlgItem$ = "TranBeforeStrButton" Or DlgItem$ = "TranAfterStrButton" Then
				Call TranTest(DlgValue("EngineList"),EngineList,1)
			End If

			If DlgValue("Engine") = 1 Then
				Header = DlgText("EngineList")
				LngName = DlgText("LngNameList")
				TranLngCode = DlgText("TranLngList")
				If TranLngCode = NullValue Then TranLngCode = ""
				If TranLngCode = "" Then DlgEnable "NullLngButton",False
				If TranLngCode <> "" Then DlgEnable "NullLngButton",True
				DlgEnable "ResetLngButton",False
				For i = LBound(EngineDataListBak) To UBound(EngineDataListBak)
					TempArray = Split(EngineDataListBak(i),JoinStr)
					If TempArray(0) = Header Then
						LangArray = Split(TempArray(2),SubJoinStr)
						For n = 0 To UBound(LangArray)
							LangPairList = Split(LangArray(n),LngJoinStr)
							If LangPairList(0) = LngName Then
								If LCase(LangPairList(2)) <> LCase(TranLngCode) Then
									DlgEnable "ResetLngButton",True
								End If
								Exit For
							End If
						Next n
						Exit For
					End If
				Next i
			End If

			If DlgItem$ = "ImportButton" Then
				If PSL.SelectFile(Path,True,Msg25 & Msg44,Msg42) = True Then
					If EngineGet("",EngineList,EngineDataList,Path) = True Then
						'EngineGet("",EngineListBak,EngineDataListBak,Path)
						DlgListBoxArray "EngineList",EngineList()
						HeaderID = UBound(EngineList)
						DlgValue "EngineList",HeaderID
						TempArray = Split(EngineDataList(HeaderID),JoinStr)
						SetsArray = Split(TempArray(1),SubJoinStr)
						DlgText "ObjectNameBox",SetsArray(0)
						DlgText "AppIDBox",SetsArray(1)
						DlgText "UrlBox",SetsArray(2)
						DlgText "UrlTemplateBox",SetsArray(3)
						DlgText "bstrMethodBox",SetsArray(4)
						DlgText "varAsyncBox",SetsArray(5)
						DlgText "bstrUserBox",SetsArray(6)
						DlgText "bstrPasswordBox",SetsArray(7)
						DlgText "varBodyBox",SetsArray(8)
						DlgText "setRequestHeaderBox",SetsArray(9)
						DlgText "responseTypeBox",SetsArray(10)
						DlgText "TranBeforeStrBox",SetsArray(11)
						DlgText "TranAfterStrBox",SetsArray(12)
						If SetsArray(0) = "" Then DlgText "ObjectNameBox",DefaultObject

						n = DlgValue("LngNameList")
						SplitData(TempArray(2),LngNameList,SrcLngList,TranLngList)
						DlgListBoxArray "LngNameList",LngNameList()
						DlgListBoxArray "SrcLngList",SrcLngList()
						DlgListBoxArray "TranLngList",TranLngList()
						DlgValue "LngNameList",n
						DlgValue "SrcLngList",n
						DlgValue "TranLngList",n
						DlgEnable "ShowNoNullLngButton",True
						DlgEnable "ShowNullLngButton",True
						DlgEnable "ShowAllLngButton",False
						MsgBox(Msg33,vbOkOnly+vbInformation,Msg30)
					Else
						MsgBox(Msg39 & tWriteLoc,vbOkOnly+vbInformation,Msg01)
					End If
				End If
			End If

			If DlgItem$ = "ExportButton" Then
				If CheckNullData("",EngineDataList,"1,5,6,7,8",1) = True Then
					If MsgBox(Msg19 & Msg12,vbYesNo+vbInformation,Msg11) = vbNo Then
						SetFunc% = True '滅砦偌狟偌聽壽敕勤趕遺敦諳
						Exit Function
					End If
				End If
				If PSL.SelectFile(Path,False,Msg25 & Msg44,Msg43) = True Then
					If EngineWrite(EngineDataList,Path,"All") = False Then
						MsgBox(Msg40 & Path,vbOkOnly+vbInformation,Msg01)
					Else
						MsgBox(Msg32,vbOkOnly+vbInformation,Msg30)
					End If
				End If
			End If

			If DlgItem$ <> "CancelButton" And DlgItem$ <> "OKButton" And DlgItem$ <> "ExportButton" Then
				Header = DlgText("EngineList")
				ObjName = DlgText("ObjectNameBox")
				AppId = DlgText("AppIDBox")
				Url = DlgText("UrlBox")
				UrlTmp = DlgText("UrlTemplateBox")
				Method = DlgText("bstrMethodBox")
				Async = DlgText("varAsyncBox")
				User = DlgText("bstrUserBox")
				Pwd = DlgText("bstrPasswordBox")
				Body = DlgText("varBodyBox")
				rHead = DlgText("setRequestHeaderBox")
				rType = DlgText("responseTypeBox")
				bStr = DlgText("TranBeforeStrBox")
				aStr = DlgText("TranAfterStrBox")

				For i = 0 To UBound(LngNameList)
					LngName = LngNameList(i)
					SrcLngCode = SrcLngList(i)
					TranLngCode = TranLngList(i)
					If SrcLngCode = NullValue Then SrcLngCode = ""
					If TranLngCode = NullValue Then TranLngCode = ""
					ReDim Preserve LangArray(i)
					LangArray(i) = LngName & LngJoinStr & SrcLngCode & LngJoinStr & TranLngCode
				Next i
				LngPair = Join(LangArray,SubJoinStr)

				Temp = AppId & Url & UrlTmp & Method & Async & User & Pwd & Body & _
						rHead & rType & bStr & aStr & LngPair
				If Temp <> "" Then
					Data = Header & JoinStr & ObjName & SubJoinStr & AppId & SubJoinStr & Url & _
							SubJoinStr & UrlTmp & SubJoinStr & Method & SubJoinStr & Async & SubJoinStr & _
							User & SubJoinStr & Pwd & SubJoinStr & Body & SubJoinStr & rHead & _
							SubJoinStr & rType & SubJoinStr & bStr & SubJoinStr & aStr & JoinStr & LngPair
					CreateArray(Header,Data,EngineList,EngineDataList)
				End If

				If InStr(Join(DefaultEngineList,JoinStr),Header) Then
					DlgEnable "ChangButton",False
					DlgEnable "DelButton",False
				Else
					DlgEnable "ChangButton",True
					If UBound(EngineList) = 0 Then DlgEnable "DelButton",False
					If UBound(EngineList) <> 0 Then DlgEnable "DelButton",True
				End If
				If DlgText("responseTypeBox") = "responseXML" Then
					DlgText "TranBeforeStrText",Msg52
					DlgText "TranAfterStrText",Msg53
				Else
					DlgText "TranBeforeStrText",Msg50
					DlgText "TranAfterStrText",Msg51
				End If
			End If

			If DlgItem$ = "TestButton" Then
				If CheckNullData("",EngineDataList,"1,5,6,7,8",1) = True Then
					If MsgBox(Msg25 & Msg19 & Msg12,vbYesNo+vbInformation,Msg11) = vbNo Then
						SetFunc% = True '滅砦偌狟偌聽壽敕勤趕遺敦諳
						Exit Function
					End If
				End If
				Call TranTest(DlgValue("EngineList"),EngineList,0)
			End If

			If DlgItem$ = "HelpButton" Then Call EngineHelp("SetHelp")
		ElseIf DlgValue("Options") = 1 Then
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
				If CheckNullData("",CheckDataList,"4",1) = True Then
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
				If CheckNullData("",CheckDataList,"4",1) = True Then
					If MsgBox(Msg26 & Msg19 & Msg12,vbYesNo+vbInformation,Msg11) = vbNo Then
						SetFunc% = True '滅砦偌狟偌聽壽敕勤趕遺敦諳
						Exit Function
					End If
				End If
				Call CheckTest(DlgValue("CheckList"),CheckList)
			End If

			If DlgItem$ = "HelpButton" Then Call CheckHelp("SetHelp")
		ElseIf DlgValue("Options") = 2 Then
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
					DlgValue "UpdateSet",StrToInteger(tUpdateSetBak(0))
					DlgText "WebSiteBox",tUpdateSetBak(1)
					DlgText "CmdPathBox",tUpdateSetBak(2)
					DlgText "ArgumentBox",tUpdateSetBak(3)
					DlgText "UpdateCycleBox",tUpdateSetBak(4)
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
					tUpdateSet(5) = DlgText("UpdateDateBox")
					If DlgValue("tWriteType") = 0 Then tPath = EngineFilePath
					If DlgValue("tWriteType") = 1 Then tPath = EngineRegKey
					If EngineWrite(EngineDataList,tPath,"Update") = False Then
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
			tUpdateSet = Split(Data,JoinStr)
			If DlgItem$ = "TestButton" Then Download(updateMethod,UpdateUrl,updateAsync,"4")
			If DlgItem$ = "HelpButton" Then Call UpdateHelp("SetHelp")
		End If

		If DlgItem$ = "OKButton" Then
			tStemp = CheckNullData("",EngineDataList,"1,5,6,7,8",1)
			cStemp = CheckNullData("",CheckDataList,"4",1)
			If tStemp = True Or cStemp = True Then
				If tStemp = True And cStemp <> True Then Temp = Msg25 & Msg19 & Msg12
				If tStemp <> True And cStemp = True Then Temp = Msg26 & Msg19 & Msg12
				If tStemp = True And cStemp = True Then Temp = Msg27 & Msg19 & Msg12
				If MsgBox(Temp,vbYesNo+vbInformation,Msg11) = vbNo Then
					SetFunc% = True '滅砦偌狟偌聽壽敕勤趕遺敦諳
					Exit Function
				End If
			End If

			tStemp = False
			cStemp = False
			If DlgValue("tWriteType") = 0 Then tPath = EngineFilePath
			If DlgValue("tWriteType") = 1 Then tPath = EngineRegKey
			If DlgValue("cWriteType") = 0 Then cPath = CheckFilePath
			If DlgValue("cWriteType") = 1 Then cPath = CheckRegKey
			If EngineWrite(EngineDataList,tPath,"Sets") = False Then tStemp = True
			If CheckWrite(CheckDataList,cPath,"Sets") = False Then cStemp = True
			If tStemp = True Or cStemp = True Then
				If tStemp = True And cStemp <> True Then Temp = Msg25 & Msg36 & tPath
				If tStemp <> True And cStemp = True Then Temp = Msg26 & Msg36 & cPath
				If tStemp = True And cStemp = True Then Temp = Msg27 & Msg36 & tPath & vbCrLf & cPath
				MsgBox(Temp,vbOkOnly+vbInformation,Msg01)
				SetFunc% = True '滅砦偌狟偌聽壽敕勤趕遺敦諳
				Exit Function
			Else
				EngineListBak = EngineList
				EngineDataListBak = EngineDataList
				CheckListBak = CheckList
				CheckDataListBak = CheckDataList
				tUpdateSetBak = tUpdateSet
			End If
			Temp = trn.Project.Location & "\~temp.txt"
			On Error Resume Next
			If Dir(Temp) <> "" Then Kill Temp
			On Error GoTo 0
		End If

		If DlgItem$ = "CancelButton" Then
			EngineList = EngineListBak
			EngineDataList = EngineDataListBak
			CheckList = CheckListBak
			CheckDataList = CheckDataListBak
			tUpdateSet = tUpdateSetBak
			Temp = trn.Project.Location & "\~temp.txt"
			On Error Resume Next
			If Dir(Temp) <> "" Then Kill Temp
			On Error GoTo 0
		End If
		If DlgItem$ <> "OKButton" And DlgItem$ <> "CancelButton" Then
			SetFunc% = True '滅砦偌狟偌聽壽敕勤趕遺敦諳
		End If
	Case 3 ' 恅掛遺麼氪郪磁遺恅掛掩載蜊
		If DlgValue("Options") = 0 Then
			HeaderID = DlgValue("EngineList")
			Header = DlgText("EngineList")
			ObjName = DlgText("ObjectNameBox")
			AppId = DlgText("AppIDBox")
			Url = DlgText("UrlBox")
			UrlTmp = DlgText("UrlTemplateBox")
			Method = DlgText("bstrMethodBox")
			Async = DlgText("varAsyncBox")
			User = DlgText("bstrUserBox")
			Pwd = DlgText("bstrPasswordBox")
			Body = DlgText("varBodyBox")
			rHead = DlgText("setRequestHeaderBox")
			rType = DlgText("responseTypeBox")
			bStr = DlgText("TranBeforeStrBox")
			aStr = DlgText("TranAfterStrBox")
			TempArray = Split(EngineDataList(HeaderID),JoinStr)
			LngPair = TempArray(2)

			Temp = AppId & Url & UrlTmp & Method & Async & User & Pwd & Body & _
					rHead & rType & bStr & aStr & LngPair
			If Temp <> "" Then
				Data = Header & JoinStr & ObjName & SubJoinStr & AppId & SubJoinStr & Url & _
						SubJoinStr & UrlTmp & SubJoinStr & Method & SubJoinStr & Async & SubJoinStr & _
						User & SubJoinStr & Pwd & SubJoinStr & Body & SubJoinStr & rHead & _
						SubJoinStr & rType & SubJoinStr & bStr & SubJoinStr & aStr & JoinStr & LngPair
				CreateArray(Header,Data,EngineList,EngineDataList)
			End If
			If DlgText("responseTypeBox") = "responseXML" Then
				DlgText "TranBeforeStrText",Msg52
				DlgText "TranAfterStrText",Msg53
			Else
				DlgText "TranBeforeStrText",Msg50
				DlgText "TranAfterStrText",Msg51
			End If
		ElseIf DlgValue("Options") = 1 Then
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
		ElseIf DlgValue("Options") = 2 Then
			UpdateMode = DlgValue("UpdateSet")
			UpdateUrl = DlgText("WebSiteBox")
			CmdPath = DlgText("CmdPathBox")
			CmdArg = DlgText("ArgumentBox")
			UpdateCycle = DlgText("UpdateCycleBox")
			UpdateDate = DlgText("UpdateDateBox")
			If UpdateDate = Msg08 Then UpdateDate = ""
			Data = UpdateMode & JoinStr & updateUrl & JoinStr & CmdPath & JoinStr & CmdArg & _
					JoinStr & UpdateCycle & JoinStr & UpdateDate
			tUpdateSet = Split(Data,JoinStr)
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

	Begin Dialog UserDialog 310,126,Msg01 ' %GRID:10,7,1,1
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
Function EditLang(DataArr() As String,LangName As String,SrcCode As String,TarnCode As String) As String
	Dim tempHeader As String,NewLangName As String,NewSrcCode As String,NewTarnCode As String
	If OSLanguage = "0404" Then
		Msg01 = "新增"
		Msg02 = "編輯"
		Msg04 = "語言名稱:"
		Msg05 = "Passolo 語言代碼:"
		Msg06 = "翻譯引擎語言代碼:"
		Msg10 = "錯誤"
		Msg11 = "您沒有輸入任何內容！請重新輸入。"
		Msg12 = "語言名稱和 Passolo 語言代碼中至少有一個項目為空！請檢查並輸入。"
		Msg13 = "該語言名稱已經存在！請重新輸入。"
		Msg14 = "該 Passolo 語言代碼已經存在！請重新輸入。"
		Msg15 = "翻譯引擎語言代碼為空！是否要重新輸入？"
		Msg16 = "如果確實需要空值，系統將自動設定為「Null」值。"
	Else
		Msg01 = "氝樓"
		Msg02 = "晤憮"
		Msg04 = "逄晟靡備:"
		Msg05 = "Passolo 逄晟測鎢:"
		Msg06 = "楹祒竘э逄晟測鎢:"
		Msg10 = "渣昫"
		Msg11 = "蠟羶衄怀�躽庥恅硜搟﹊鄵寪薹靿諢�"
		Msg12 = "逄晟靡備睿 Passolo 逄晟測鎢笢祫屾衄珨跺砐醴峈諾ㄐ③潰脤甜怀�諢�"
		Msg13 = "蜆逄晟靡備眒冪湔婓ㄐ③笭陔怀�諢�"
		Msg14 = "蜆 Passolo 逄晟測鎢眒冪湔婓ㄐ③笭陔怀�諢�"
		Msg15 = "楹祒竘э逄晟測鎢峈諾ㄐ岆瘁猁笭陔怀�諴�"
		Msg16 = "�蝜��滔菩骳矽欶童疢腴魚度堈胱襞炕衹ull§硉﹝"
	End If
	If LangName = "" Then Msg = Msg01
	If LangName <> "" Then Msg = Msg02
	Begin Dialog UserDialog 390,189,Msg ' %GRID:10,7,1,1
		Text 10,7,370,14,Msg04
		TextBox 10,28,370,21,.LangName
		Text 10,56,370,14,Msg05
		TextBox 10,77,370,21,.SrcCode
		Text 10,105,370,14,Msg06
		TextBox 10,126,370,21,.TarnCode
		OKButton 90,161,80,21,.OKButton
		CancelButton 220,161,80,21,.CancelButton
	End Dialog
    Dim dlg As UserDialog
    If LangName <> "" Then dlg.LangName = LangName
    If SrcCode <> "" Then dlg.SrcCode = SrcCode
    If TarnCode <> "" Then dlg.TarnCode = TarnCode
    DataInPutDlg:
    If Dialog(dlg) = 0 Then Exit Function

	NewLangName = Trim(dlg.LangName)
	NewSrcCode = Trim(dlg.SrcCode)
	NewTarnCode = Trim(dlg.TarnCode)
	OldLangName = ""
	OldSrcCode = ""
	If NewLangName <> "" And NewSrcCode <> "" And NewTarnCode <> "" Then
		For i = LBound(DataArr) To UBound(DataArr)
			LangPairList = Split(DataArr(i),LngJoinStr)
			If NewLangName = LangPairList(0) Then OldLangName = LangPairList(0)
			If NewSrcCode = LangPairList(1) Then OldSrcCode = LangPairList(1)
			If OldLangName <> "" And OldSrcCode <> "" Then Exit For
		Next i
	End If

	If NewLangName & NewSrcCode & NewTarnCode = "" Then
		MsgBox Msg11,vbOkOnly+vbInformation,Msg10
		GoTo DataInPutDlg
	End If
	If NewLangName = "" Or NewSrcCode = "" Then
		MsgBox Msg12,vbOkOnly+vbInformation,Msg10
		GoTo DataInPutDlg
	End If
	If OldLangName <> "" And NewLangName <> LangName Then
		MsgBox Msg13,vbOkOnly+vbInformation,Msg10
		GoTo DataInPutDlg
	End If
	If OldSrcCode <> "" And NewSrcCode <> SrcCode Then
		MsgBox Msg14,vbOkOnly+vbInformation,Msg10
		GoTo DataInPutDlg
	End If
	If NewTarnCode = "" Then
		Massage = MsgBox(Msg15 & vbCrLf & Msg16,vbYesNo+vbInformation,Msg10)
		If Massage = vbYes Then GoTo DataInPutDlg
	End If
	If NewHeader = NullValue Then NewHeader = ""

	EditLang = NewLangName & LngJoinStr & NewSrcCode & LngJoinStr & NewTarnCode
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
	Begin Dialog UserDialog 390,168,Msg ' %GRID:10,7,1,1
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


'鳳�﹎髲�
Function EngineGet(SelSet As String,HeaderList() As String,DataList() As String,Path As String) As Boolean
	Dim i As Integer,Header As String,HeaderIDs As String,HeaderIDArr() As String,Temp As String
	Dim LangPairList() As String,TempArray() As String
	EngineGet = False
	NewVersion = ToUpdateEngineVersion
	LangPairList = LangCodeList("engine",OSLanguage,0,107)

	If Path = EngineRegKey Then GoTo GetFromRegistry
	If Path = "" Then Path = EngineFilePath
	If Dir(Path) = "" Then GoTo GetFromRegistry
	If Path = EngineFilePath Then On Error GoTo GetFromRegistry
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
				If setPreStr = "TranEngineSet" Then EngineSet = setAppStr
				If setPreStr = "CheckSet" Then CheckSet = setAppStr
				If setPreStr = "TranAllType" Then mAllType = setAppStr
				If setPreStr = "TranMenu" Then mMenu = setAppStr
				If setPreStr = "TranDialog" Then mDialog = setAppStr
				If setPreStr = "TranString" Then mString = setAppStr
				If setPreStr = "TranAcceleratorTable" Then mAccTable = setAppStr
				If setPreStr = "TranVersion" Then mVer = setAppStr
				If setPreStr = "TranOther" Then mOther = setAppStr
				If setPreStr = "TranSeletedOnly" Then mSelOnly = setAppStr
				If setPreStr = "SkipForReview" Then mForReview = setAppStr
				If setPreStr = "SkipValidated" Then mValidated = setAppStr
				If setPreStr = "SkipNotTran" Then mNotTran = setAppStr
				If setPreStr = "SkipAllNumAndSymbol" Then mNumAndSymbol = setAppStr
				If setPreStr = "SkipAllUCase" Then mAllUCase = setAppStr
				If setPreStr = "SkipAllLCase" Then mAllLCase = setAppStr
				If setPreStr = "AutoSelection" Then mAutoSele = setAppStr
				If setPreStr = "CheckAccessKey" Then mAccKey = setAppStr
				If setPreStr = "CheckAccelerator" Then mAccelerator = setAppStr
				If setPreStr = "CheckPreString" Then mPreStrRep = setAppStr
				If setPreStr = "CheckWebPageFlag" Then mWebFlag = setAppStr
				If setPreStr = "CheckTranslation" Then mCheckTrn = setAppStr
				If setPreStr = "AutoRepString" Then mAppStrRep = setAppStr
				If setPreStr = "KeepSetting" Then KeepSet = setAppStr
				If setPreStr = "ShowMassage" Then ShowMsg = setAppStr
				If setPreStr = "AddTranComment" Then TranComm = setAppStr
			End If
			'鳳�� Update 砐睿硉
			If Header = "Update" And SelSet = "" And setPreStr <> "" Then
				If InStr(setPreStr,"Site_") Then Site = setAppStr
				If setPreStr = "UpdateMode" Then UpdateMode = setAppStr
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
				If setPreStr = "ObjectName" Then ObjName = setAppStr
				If setPreStr = "AppId" Then AppId = setAppStr
				If setPreStr = "EngineUrl" Then Url = setAppStr
				If setPreStr = "UrlTemplate" Then UrlTmp = setAppStr
				If setPreStr = "Method" Then Method = setAppStr
				If setPreStr = "Async" Then Async = setAppStr
				If setPreStr = "User" Then User = setAppStr
				If setPreStr = "Password" Then Pwd = setAppStr
				If setPreStr = "SendBody" Then Body = setAppStr
				If setPreStr = "RequestHeader" Then rHead = setAppStr
				If setPreStr = "ResponseType" Then rType = setAppStr
				If setPreStr = "TranBeforeStr" Then bStr = setAppStr
				If setPreStr = "TranAfterStr" Then aStr = setAppStr
				If setPreStr = "LangCodePair" Then LngPair = setAppStr
			End If
		End If
		If (L$ = "" Or EOF(1)) And Header = "Option" Then
			Data = EngineSet & JoinStr & CheckSet & JoinStr & mAllType & JoinStr & mMenu & _
					JoinStr & mDialog & JoinStr & mString & JoinStr & mAccTable & JoinStr & _
					mVer & JoinStr & mOther & JoinStr & mSelOnly & JoinStr & mForReview & _
					JoinStr & mValidated & JoinStr & mNotTran & JoinStr & mNumAndSymbol & _
					JoinStr & mAllUCase & JoinStr & mAllLCase & JoinStr & mAutoSele & _
					JoinStr & mAccKey & JoinStr & mAccelerator & JoinStr & mPreStrRep & _
					JoinStr & mWebFlag & JoinStr & mCheckTrn & JoinStr & mAppStrRep & _
					JoinStr & KeepSet & JoinStr & ShowMsg & JoinStr & TranComm
			tSelected = Split(Data,JoinStr)
		End If
		If (L$ = "" Or EOF(1)) And Header = "Update" Then
			Data = UpdateMode & JoinStr & UpdateSite & JoinStr & CmdPath & JoinStr & CmdArg & _
					JoinStr & UpdateCycle & JoinStr & UpdateDate
			tUpdateSet = Split(Data,JoinStr)
		End If
		If (L$ = "" Or EOF(1)) And Header <> "" And Header <> "Option" And Header <> "Update" Then
			Temp = ObjName & UrlTmp & AppId & Url & Method & Async & User & Pwd & Body & _
					rHead & rType & bStr & aStr & LngPair
			If Temp <> "" Then
				If LngPair <> "" Then
					If Path <> EngineFilePath Then LngPair = LangNameUpdate(Header,LngPair)
					TempArray = Split(LngPair,SubJoinStr)
					LngPair = Join(MergeLngList(LangPairList,TempArray,"engine"),SubJoinStr)
				Else
					LngPair = Join(LangCodeList(Header,OSLanguage,0,107),SubJoinStr)
				End If
				Data = Header & JoinStr & ObjName & SubJoinStr & AppId & SubJoinStr & Url & _
						SubJoinStr & UrlTmp & SubJoinStr & Method & SubJoinStr & Async & SubJoinStr & _
						User & SubJoinStr & Pwd & SubJoinStr & Body & SubJoinStr & rHead & _
						SubJoinStr & rType & SubJoinStr & bStr & SubJoinStr & aStr & JoinStr & LngPair
				'載陔導唳腔蘇�狣馺譆�
				If InStr(Join(DefaultEngineList,JoinStr),Header) Then
					If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 Then
						Data = EngineDataUpdate(Header,Data)
					End If
				End If
				'悵湔杅擂善杅郪笢
				CreateArray(Header,Data,HeaderList,DataList)
				EngineGet = True
			End If
			'杅擂場宎趙
			ObjName = ""
			UrlTmp = ""
			AppId = ""
			Url = ""
			Method = ""
			Async = ""
			User = ""
			Pwd = ""
			Body = ""
			rHead = ""
			rType = ""
			bStr = ""
			aStr = ""
			LngPair = ""
		End If
	Wend
	Close #1
	If Path = EngineFilePath Then On Error GoTo 0
	'悵湔載陔睿絳�赮騕騫�擂善恅璃
	If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 And Path = EngineFilePath Then
		If Dir(EngineFilePath) <> "" Then EngineWrite(DataList,EngineFilePath,"All")
	End If
	If tWriteLoc = "" Then tWriteLoc = EngineFilePath
	Exit Function

	GetFromRegistry:
	'鳳�� Option 砐睿硉
	OldVersion = GetSetting("WebTranslate","Option","Version","")
	If SelSet = "" Then
		EngineSet = GetSetting("WebTranslate","Option","TranEngineSet","")
		CheckSet = GetSetting("WebTranslate","Option","CheckSet","")
		mAllType = GetSetting("WebTranslate","Option","TranAllType",0)
		mMenu = GetSetting("WebTranslate","Option","TranMenu",0)
		mDialog = GetSetting("WebTranslate","Option","TranDialog",0)
		mString = GetSetting("WebTranslate","Option","TranString",0)
		mAccTable = GetSetting("WebTranslate","Option","TranAcceleratorTable",0)
		mVer = GetSetting("WebTranslate","Option","TranVersion",0)
		mOther = GetSetting("WebTranslate","Option","TranOther",0)
		mSelOnly = GetSetting("WebTranslate","Option","TranSeletedOnly",1)
		mForReview = GetSetting("WebTranslate","Option","SkipForReview",1)
		mValidated = GetSetting("WebTranslate","Option","SkipValidated",1)
		mNotTran = GetSetting("WebTranslate","Option","SkipNotTran",0)
		mNumAndSymbol = GetSetting("WebTranslate","Option","SkipAllNumAndSymbol",1)
		mAllUCase = GetSetting("WebTranslate","Option","SkipAllUCase",1)
		mAllLCase = GetSetting("WebTranslate","Option","SkipAllLCase",0)
		mAutoSele = GetSetting("WebTranslate","Option","AutoSelection",1)
		mAccKey = GetSetting("WebTranslate","Option","CheckAccessKey",1)
		mAccelerator = GetSetting("WebTranslate","Option","CheckAccelerator",1)
		mPreStrRep = GetSetting("WebTranslate","Option","CheckPreString",1)
		mWebFlag = GetSetting("WebTranslate","Option","CheckWebPageFlag",1)
		mCheckTrn = GetSetting("WebTranslate","Option","CheckTranslation",1)
		mAppStrRep = GetSetting("WebTranslate","Option","AutoRepString",1)
		KeepSet = GetSetting("WebTranslate","Option","KeepSetting",1)
		ShowMsg = GetSetting("WebTranslate","Option","ShowMassage",1)
		TranComm = GetSetting("WebTranslate","Option","AddTranComment",0)
		Data = EngineSet & JoinStr & CheckSet & JoinStr & mAllType & JoinStr & mMenu & _
				JoinStr & mDialog & JoinStr & mString & JoinStr & mAccTable & JoinStr & _
				mVer & JoinStr & mOther & JoinStr & mSelOnly & JoinStr & mForReview & _
				JoinStr & mValidated & JoinStr & mNotTran & JoinStr & mNumAndSymbol & _
				JoinStr & mAllUCase & JoinStr & mAllLCase & JoinStr & mAutoSele & _
				JoinStr & mAccKey & JoinStr & mAccelerator & JoinStr & mPreStrRep & _
				JoinStr & mWebFlag & JoinStr & mCheckTrn & JoinStr & mAppStrRep & _
				JoinStr & KeepSet & JoinStr & ShowMsg & JoinStr & TranComm
		tSelected = Split(Data,JoinStr)
		'鳳�� Update 砐睿硉
		UpdateMode = GetSetting("WebTranslate","Update","UpdateMode",1)
		Count = GetSetting("WebTranslate","Update","Count",0)
		For i = 0 To Count
			Site = GetSetting("WebTranslate","Update",CStr(i),"")
			If Site <> "" Then
				If i = 0 Then UpdateSite = Site
				If i > 0 Then UpdateSite = UpdateSite & vbCrLf & Site
			End If
		Next i
		CmdPath = GetSetting("WebTranslate","Update","Path","")
		CmdArg = GetSetting("WebTranslate","Update","Argument","")
		UpdateCycle = GetSetting("WebTranslate","Update","UpdateCycle",7)
		UpdateDate = GetSetting("WebTranslate","Update","UpdateDate","")
		Data = UpdateMode & JoinStr & UpdateSite & JoinStr & CmdPath & JoinStr & CmdArg & _
				JoinStr & UpdateCycle & JoinStr & UpdateDate
		tUpdateSet = Split(Data,JoinStr)
	End If

	'鳳�� Option 俋腔砐睿硉
	HeaderIDs = GetSetting("WebTranslate","Option","Headers","")
	If HeaderIDs <> "" Then
		HeaderIDArr = Split(HeaderIDs,";",-1)
		For i = LBound(HeaderIDArr) To UBound(HeaderIDArr)
			HeaderID = HeaderIDArr(i)
			If HeaderID <> "" Then
				'蛌湔導唳腔藩跺砐睿硉
				Header = GetSetting("WebTranslate",HeaderID,"Name","")
				If Header = "" Then Header = HeaderID
				If Header <> "" Then
					ObjName = GetSetting("WebTranslate",HeaderID,"ObjectName","")
					AppId = GetSetting("WebTranslate",HeaderID,"AppId","")
					Url = GetSetting("WebTranslate",HeaderID,"EngineUrl","")
					UrlTmp = GetSetting("WebTranslate",HeaderID,"UrlTemplate","")
					Method = GetSetting("WebTranslate",HeaderID,"Method","")
					Async = GetSetting("WebTranslate",HeaderID,"Async","")
					User = GetSetting("WebTranslate",HeaderID,"User","")
					Pwd = GetSetting("WebTranslate",HeaderID,"Password","")
					Body = GetSetting("WebTranslate",HeaderID,"SendBody","")
					rHead = GetSetting("WebTranslate",HeaderID,"RequestHeader","")
					rType = GetSetting("WebTranslate",HeaderID,"ResponseType","")
					bStr = GetSetting("WebTranslate",HeaderID,"TranBeforeStr","")
					aStr = GetSetting("WebTranslate",HeaderID,"TranAfterStr","")
					LngPair = GetSetting("WebTranslate",HeaderID,"LangCodePair","")
					Temp = ObjName & UrlTmp & AppId & Url & Method & Async & User & Pwd & Body & _
							rHead & rType & bStr & aStr & LngPair
					If Temp <> "" Then
						If LngPair <> "" Then
							TempArray = Split(LngPair,SubJoinStr)
							LngPair = Join(MergeLngList(LangPairList,TempArray,"engine"),SubJoinStr)
						Else
							LngPair = Join(LangCodeList(Header,OSLanguage,0,107),SubJoinStr)
						End If
						Data = Header & JoinStr & ObjName & SubJoinStr & AppId & SubJoinStr & Url & _
								SubJoinStr & UrlTmp & SubJoinStr & Method & SubJoinStr & Async & SubJoinStr & _
								User & SubJoinStr & Pwd & SubJoinStr & Body & SubJoinStr & rHead & _
								SubJoinStr & rType & SubJoinStr & bStr & SubJoinStr & aStr & JoinStr & LngPair
						'載陔導唳腔蘇�狣馺譆�
						If InStr(Join(DefaultEngineList,JoinStr),Header) Then
							If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 Then
								Data = EngineDataUpdate(Header,Data)
							End If
						End If
						'悵湔杅擂善杅郪笢
						CreateArray(Header,Data,HeaderList,DataList)
						EngineGet = True
					End If
					'刉壺導唳饜离硉
					On Error Resume Next
					If Header = HeaderID Then DeleteSetting("WebTranslate",Header)
					On Error GoTo 0
				End If
			End If
		Next i
	End If
	'悵湔載陔綴腔杅擂善蛁聊桶
	If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 Then
		If HeaderIDs <> "" Then EngineWrite(DataList,EngineRegKey,"Sets")
	End If
	If tWriteLoc = "" Then tWriteLoc = EngineRegKey
End Function


'迡�賰倡遻�э扢离
Function EngineWrite(DataList() As String,Path As String,WriteType As String) As Boolean
	Dim i As Integer,HeaderID As String,HeaderIDs As String,TempPath As String
	Dim SetsArray() As String,HeaderIDArr() As String
	EngineWrite = False
	KeepSet = tSelected(23)

	'迡�輷躁�
	If Path <> "" And Path <> EngineRegKey Then
   		On Error Resume Next
   		TempPath = Left(Path,InStrRev(Path,"\"))
   		If Dir(TempPath & "*.*") = "" Then MkDir TempPath
		If Dir(Path) <> "" Then SetAttr Path,vbNormal
		On Error GoTo 0
		On Error GoTo ExitFunction
		Open Path For Output As #2
			Print #2,";------------------------------------------------------------"
			Print #2,";Settings for PSLWebTrans.bas"
			Print #2,";------------------------------------------------------------"
			Print #2,""
			Print #2,"[Option]"
			Print #2,"Version = " & Version
			If KeepSet = "1" Then
				Print #2,"TranEngineSet = " & tSelected(0)
				Print #2,"CheckSet = " & tSelected(1)
				Print #2,"TranAllType = " & tSelected(2)
				Print #2,"TranMenu = " & tSelected(3)
				Print #2,"TranDialog = " & tSelected(4)
				Print #2,"TranString = " & tSelected(5)
				Print #2,"TranAcceleratorTable = " & tSelected(6)
				Print #2,"TranVersion = " & tSelected(7)
				Print #2,"TranOther = " & tSelected(8)
				Print #2,"TranSeletedOnly = " & tSelected(9)
				Print #2,"SkipForReview = " & tSelected(10)
				Print #2,"SkipValidated = " & tSelected(11)
				Print #2,"SkipNotTran = " & tSelected(12)
				Print #2,"SkipAllNumAndSymbol = " & tSelected(13)
				Print #2,"SkipAllUCase = " & tSelected(14)
				Print #2,"SkipAllLCase = " & tSelected(15)
				Print #2,"AutoSelection = " & tSelected(16)
				Print #2,"CheckAccessKey = " & tSelected(17)
				Print #2,"CheckAccelerator = " & tSelected(18)
				Print #2,"CheckPreString = " & tSelected(19)
				Print #2,"CheckWebPageFlag = " & tSelected(20)
				Print #2,"CheckTranslation = " & tSelected(21)
				Print #2,"AutoRepString = " & tSelected(22)
				Print #2,"KeepSetting = " & tSelected(23)
				Print #2,"ShowMassage = " & tSelected(24)
				Print #2,"AddTranComment = " & tSelected(25)
			End If
			Print #2,""
			If Join(tUpdateSet) <> "" Then
				UpdateSiteList = Split(tUpdateSet(1),vbCrLf,-1)
				Print #2,"[Update]"
				Print #2,"UpdateMode = " & tUpdateSet(0)
				For i = LBound(UpdateSiteList) To UBound(UpdateSiteList)
					Print #2,"Site_" & CStr(i) & " = " & UpdateSiteList(i)
				Next i
				Print #2,"Path = " & tUpdateSet(2)
				Print #2,"Argument = " & tUpdateSet(3)
				Print #2,"UpdateCycle = " & tUpdateSet(4)
				Print #2,"UpdateDate = " & tUpdateSet(5)
				Print #2,""
			End If
			For i = LBound(DataList) To UBound(DataList)
				TempArray = Split(DataList(i),JoinStr)
				SetsArray = Split(TempArray(1),SubJoinStr)
				Print #2,"[" & TempArray(0) & "]"
				Print #2,"ObjectName = " & SetsArray(0)
				Print #2,"AppId = " & SetsArray(1)
				Print #2,"EngineUrl = " & SetsArray(2)
				Print #2,"UrlTemplate = " & SetsArray(3)
				Print #2,"Method = " & SetsArray(4)
				Print #2,"Async = " & SetsArray(5)
				Print #2,"User = " & SetsArray(6)
				Print #2,"Password = " & SetsArray(7)
				Print #2,"SendBody = " & SetsArray(8)
				Print #2,"RequestHeader = " & SetsArray(9)
				Print #2,"ResponseType = " & SetsArray(10)
				Print #2,"TranBeforeStr = " & SetsArray(11)
				Print #2,"TranAfterStr = " & SetsArray(12)
				Print #2,"LangCodePair = " & getLngPair(TempArray(2),"engine")
				If i <> UBound(DataList) Then Print #2,""
			Next i
		Close #2
		On Error GoTo 0
		EngineWrite = True
		If Path = EngineFilePath Then tWriteLoc = EngineFilePath
		If Path = EngineFilePath Then GoTo RemoveRegKey

	'迡�鄶３嵿�
	ElseIf Path = EngineRegKey Then
		On Error GoTo ExitFunction
		SaveSetting("WebTranslate","Option","Version",Version)
		If WriteType = "Main" Or WriteType = "All" Then
			If KeepSet = "1" Then
				SaveSetting("WebTranslate","Option","TranEngineSet",tSelected(0))
				SaveSetting("WebTranslate","Option","CheckSet",tSelected(1))
				SaveSetting("WebTranslate","Option","TranAllType",tSelected(2))
				SaveSetting("WebTranslate","Option","TranMenu",tSelected(3))
				SaveSetting("WebTranslate","Option","TranDialog",tSelected(4))
				SaveSetting("WebTranslate","Option","TranString",tSelected(5))
				SaveSetting("WebTranslate","Option","TranAcceleratorTable",tSelected(6))
				SaveSetting("WebTranslate","Option","TranVersion",tSelected(7))
				SaveSetting("WebTranslate","Option","TranOther",tSelected(8))
				SaveSetting("WebTranslate","Option","TranSeletedOnly",tSelected(9))
				SaveSetting("WebTranslate","Option","SkipForReview",tSelected(10))
				SaveSetting("WebTranslate","Option","SkipValidated",tSelected(11))
				SaveSetting("WebTranslate","Option","SkipNotTran",tSelected(12))
				SaveSetting("WebTranslate","Option","SkipAllNumAndSymbol",tSelected(13))
				SaveSetting("WebTranslate","Option","SkipAllUCase",tSelected(14))
				SaveSetting("WebTranslate","Option","SkipAllLCase",tSelected(15))
				SaveSetting("WebTranslate","Option","AutoSelection",tSelected(16))
				SaveSetting("WebTranslate","Option","CheckAccessKey",tSelected(17))
				SaveSetting("WebTranslate","Option","CheckAccelerator",tSelected(18))
				SaveSetting("WebTranslate","Option","CheckPreString",tSelected(19))
				SaveSetting("WebTranslate","Option","CheckWebPageFlag",tSelected(20))
				SaveSetting("WebTranslate","Option","CheckTranslation",tSelected(21))
				SaveSetting("WebTranslate","Option","AutoRepString",tSelected(22))
				SaveSetting("WebTranslate","Option","KeepSetting",tSelected(23))
				SaveSetting("WebTranslate","Option","ShowMassage",tSelected(24))
				SaveSetting("WebTranslate","Option","AddTranComment",tSelected(25))
			End If
		End If
		If WriteType = "Sets" Or WriteType = "All" Then
			'刉壺埻饜离砐
			HeaderIDs = GetSetting("WebTranslate","Option","Headers")
			If HeaderIDs <> "" Then
				HeaderIDArr = Split(HeaderIDs,";",-1)
				On Error Resume Next
				For i = LBound(HeaderIDArr) To UBound(HeaderIDArr)
					DeleteSetting("WebTranslate",HeaderIDArr(i))
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
				SaveSetting("WebTranslate",HeaderID,"Name",TempArray(0))
				SaveSetting("WebTranslate",HeaderID,"ObjectName",SetsArray(0))
				SaveSetting("WebTranslate",HeaderID,"AppId",SetsArray(1))
				SaveSetting("WebTranslate",HeaderID,"EngineUrl",SetsArray(2))
				SaveSetting("WebTranslate",HeaderID,"UrlTemplate",SetsArray(3))
				SaveSetting("WebTranslate",HeaderID,"Method",SetsArray(4))
				SaveSetting("WebTranslate",HeaderID,"Async",SetsArray(5))
				SaveSetting("WebTranslate",HeaderID,"User",SetsArray(6))
				SaveSetting("WebTranslate",HeaderID,"Password",SetsArray(7))
				SaveSetting("WebTranslate",HeaderID,"SendBody",SetsArray(8))
				SaveSetting("WebTranslate",HeaderID,"RequestHeader",SetsArray(9))
				SaveSetting("WebTranslate",HeaderID,"ResponseType",SetsArray(10))
				SaveSetting("WebTranslate",HeaderID,"TranBeforeStr",SetsArray(11))
				SaveSetting("WebTranslate",HeaderID,"TranAfterStr",SetsArray(12))
				SaveSetting("WebTranslate",HeaderID,"LangCodePair",getLngPair(TempArray(2),"engine"))
			Next i
			HeaderIDs = Join(HeaderIDArr,";")
			SaveSetting("WebTranslate","Option","Headers",HeaderIDs)
		End If
		If WriteType = "Update" Or WriteType = "Sets" Or WriteType = "All" Then
			If Join(tUpdateSet) <> "" Then
				On Error Resume Next
				DeleteSetting("WebTranslate","Update")
				On Error GoTo 0
				UpdateSiteList = Split(tUpdateSet(1),vbCrLf,-1)
				SaveSetting("WebTranslate","Update","UpdateMode",tUpdateSet(0))
				For i = LBound(UpdateSiteList) To UBound(UpdateSiteList)
					SaveSetting("WebTranslate","Update",CStr(i),UpdateSiteList(i))
				Next i
				SaveSetting("WebTranslate","Update","Count",UBound(UpdateSiteList))
				SaveSetting("WebTranslate","Update","Path",tUpdateSet(2))
				SaveSetting("WebTranslate","Update","Argument",tUpdateSet(3))
				SaveSetting("WebTranslate","Update","UpdateCycle",tUpdateSet(4))
				SaveSetting("WebTranslate","Update","UpdateDate",tUpdateSet(5))
			End If
		End If
		EngineWrite = True
		tWriteLoc = EngineRegKey
		GoTo RemoveFilePath
	'刉壺垀衄悵湔腔扢离
	ElseIf Path = "" Then
		'刉壺恅璃饜离砐
 		RemoveFilePath:
		On Error Resume Next
		If Dir(EngineFilePath) <> "" Then
			SetAttr EngineFilePath,vbNormal
			Kill EngineFilePath
		End If
		TempPath = Left(EngineFilePath,InStrRev(EngineFilePath,"\"))
		If Dir(TempPath & "*.*") = "" Then RmDir TempPath
		On Error GoTo 0
		If Path = EngineRegKey Then GoTo ExitFunction
		'刉壺蛁聊桶饜离砐
		RemoveRegKey:
		If GetSetting("WebTranslate","Option","Version") <> "" Then
			HeaderIDs = GetSetting("WebTranslate","Option","Headers")
			On Error Resume Next
			If HeaderIDs <> "" Then
				HeaderIDArr = Split(HeaderIDs,";",-1)
				For i = LBound(HeaderIDArr) To UBound(HeaderIDArr)
					DeleteSetting("WebTranslate",HeaderIDArr(i))
				Next i
			End If
			DeleteSetting("WebTranslate","Option")
			DeleteSetting("WebTranslate","Update")
			Dim WshShell As Object
			Set WshShell = CreateObject("WScript.Shell")
			WshShell.RegDelete EngineRegKey
			Set WshShell = Nothing
			On Error GoTo 0
		End If
		If Path = EngineFilePath Then GoTo ExitFunction
		'扢离迡�輷閥蟹髲襞矽�
		EngineWrite = True
		tWriteLoc = ""
	End If
	ExitFunction:
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
	'悵湔載陔睿絳�赮騕騫�擂善恅璃
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


'載陔竘э導唳掛饜离硉
Function EngineDataUpdate(Header As String,Data As String) As String
	Dim UpdatedData As String,uV As String,dV As String
	TempArray = Split(Data,JoinStr)
	uSetsArray = Split(TempArray(1),SubJoinStr)
	dSetsArray = Split(EngineSettings(Header),SubJoinStr)
	For i = LBound(uSetsArray) To UBound(uSetsArray)
		uV = uSetsArray(i)
		dV = dSetsArray(i)
		If uV = "" Then uV = dV
		If uV <> "" And uV <> dV Then uV = dV
		uSetsArray(i) = uV
	Next i
	UpdatedData = Join(uSetsArray,SubJoinStr)
	EngineDataUpdate = TempArray(0) & JoinStr & UpdatedData & JoinStr & TempArray(2)
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
	Dim KeyCode As Boolean,FindStr As String,Key As String,pos As Integer
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


'聆彸婓盄楹祒最唗
Sub TranTest(EngineID As Integer,HeaderList() As String,fType As Integer)
	Dim i As Integer,n As Integer,TrnList As PslTransList,TrnListDec As String
	Dim LngNameList() As String,TrnListArray() As String

	If OSLanguage = "0404" Then
		Msg01 = "翻譯引擎測試"
		Msg02 = "翻譯清單和來源語言會根據專案現有的語言自動確定。要增加項目請新增對應的語言。"
		Msg03 = "翻譯引擎:"
		Msg04 = "來源語言:"
		Msg05 = "目標語言:"
		Msg06 = "翻譯清單:"
		Msg07 = "讀入行數:"
		Msg08 = "翻譯內容(自動從選擇清單中讀入或手動輸入):"
		Msg09 = "翻譯結果(按測試按鈕後在此輸出結果):"
		Msg10 = "來源字串"
		Msg11 = "翻譯字串"
		Msg12 = "翻譯文字"
		Msg13 = "全部文字"
		Msg14 = "說明(&H)"
		Msg15 = "翻譯(&T)"
		Msg16 = "響應頭(&D)"
		Msg17 = "清空(&C)"
	Else
		Msg01 = "楹祒竘э聆彸"
		Msg02 = "楹祒蹈桶睿懂埭逄晟頗跦擂源偶珋衄腔逄晟赻雄�毓芋�猁崝樓砐醴③氝樓眈茼腔逄晟﹝"
		Msg03 = "楹祒竘э:"
		Msg04 = "懂埭逄晟:"
		Msg05 = "醴梓逄晟:"
		Msg06 = "楹祒蹈桶:"
		Msg07 = "黍�遶倇�:"
		Msg08 = "楹祒囀��(赻雄植恁隅蹈桶笢黍�趥藡硍缺靿�):"
		Msg09 = "楹祒賦彆(偌聆彸偌聽綴婓森怀堤賦彆):"
		Msg10 = "懂埭趼揹"
		Msg11 = "楹祒趼揹"
		Msg12 = "楹祒恅掛"
		Msg13 = "�垓諺覺�"
		Msg14 = "堆翑(&H)"
		Msg15 = "楹祒(&T)"
		Msg16 = "砒茼芛(&D)"
		Msg17 = "ь諾(&C)"
	End If

	ReDim TrnListArray(0)
	For i = 1 To trn.Project.TransLists.Count
		Set TrnList = trn.Project.TransLists(i)
		ReDim Preserve TrnListArray(i-1)
		TrnListDec = TrnList.Title & " - " & PSL.GetLangCode(TrnList.Language.LangID,pslCodeText)
		TrnListArray(i-1) = TrnListDec
	Next i

	Begin Dialog UserDialog 650,462,Msg01,.TranTestFunc ' %GRID:10,7,1,1
		GroupBox 10,21,630,70,"",.GroupBox
		Text 10,7,630,14,Msg02
		Text 30,38,80,14,Msg03
		DropListBox 120,35,200,21,HeaderList(),.SelSetBox
		Text 330,38,80,14,Msg05
		DropListBox 420,35,200,21,LngNameList(),.LngNameBox
		Text 30,66,80,14,Msg06
		DropListBox 120,63,330,21,TrnListArray(),.TrnListBox
		Text 470,66,90,14,Msg07
		TextBox 570,63,50,18,.LineNumBox
		Text 10,98,370,14,Msg08
		TextBox 10,119,630,126,.InTextBox,1
		OptionGroup .StrType
			OptionButton 390,98,120,14,Msg10,.SrcString
			OptionButton 520,98,120,14,Msg11,.TranString
		Text 10,252,370,14,Msg09
		TextBox 10,273,630,147,.OutTextBox,1
		OptionGroup .TranType
			OptionButton 390,252,120,14,Msg12,.TranOnly
			OptionButton 520,252,120,14,Msg13,.AllTran
		PushButton 10,434,100,21,Msg14,.HelpButton
		PushButton 120,434,100,21,Msg15,.TestButton
		PushButton 230,434,100,21,Msg16,.HeaderButton
		PushButton 340,434,100,21,Msg17,.ClearButton
		CancelButton 540,434,100,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	dlg.SelSetBox = EngineID
	dlg.TranType = fType
	If Dialog(dlg) = 0 Then Exit Sub
End Sub


'聆彸勤趕遺滲杅
Private Function TranTestFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim ListDec As String,TempDec As String,LineNum As Integer,EngineID As Integer
	Dim inText As String,outText As String,i As Integer,n As Integer,m As Integer
	Dim srcLngFind As Integer,trnLngFind As Integer,LngNameList() As String
	Dim TrnList As PslTransList,xmlHttp As Object

	If OSLanguage = "0404" Then
		Msg01 = "清空(&C)"
		Msg02 = "讀入(&R)"
		Msg05 = "正在翻譯，可能需要幾分鐘，請稍候..."
		Msg06 = "測試的語言代碼對為: "
		Msg07 = "翻譯失敗！請檢查 Internet 連接、翻譯引擎的設定或是否可存取後再試。"
		Msg08 = "返回的字串為空。可能是翻譯引擎不支援選擇的目標語言。"
		Msg09 = "========================================="
	Else
		Msg01 = "ь諾(&C)"
		Msg02 = "黍��(&R)"
		Msg05 = "淏婓楹祒ㄛ褫夔剒猁撓煦笘ㄛ③尕緊..."
		Msg06 = "聆彸腔逄晟測鎢勤峈: "
		Msg07 = "楹祒囮啖ㄐ③潰脤 Internet 蟀諉﹜楹祒竘э腔扢离麼岆瘁褫溼恀綴婬彸﹝"
		Msg08 = "殿隙腔趼揹峈諾﹝褫夔岆楹祒竘э祥盓厥恁隅腔醴梓逄晟﹝"
		Msg09 = "========================================="
	End If

	Select Case Action%
	Case 1
		TempDec = trn.Title & " - " & PSL.GetLangCode(trn.Language.LangID,pslCodeText)
		If trn.SourceList.LastChange > trn.LastUpdate Then trn.Update
		LineNum = 10
		If LineNum > trn.StringCount Then LineNum = trn.StringCount
		For i = 1 To LineNum
			Set TransString = trn.String(i)
			If TransString.Text <> "" Then
				If DlgValue("StrType") = 0 Then srcString = TransString.SourceText
				If DlgValue("StrType") = 1 Then srcString = TransString.Text
				If inText <> "" Then inText = inText & " " & srcString
				If inText = "" Then inText = srcString
			End If
		Next i
		trnLng = PSL.GetLangCode(trn.Language.LangID,pslCode639_1)
		If trnLng = "zh" Then
			trnLng = PSL.GetLangCode(trn.Language.LangID,pslCodeLangRgn)
			If trnLng = "zh-CHS" Or trnLng = "zh-SG" Then trnLng = "zh-CN"
			If trnLng = "zh-CHT" Or trnLng = "zh-HK" Or trnLng = "zh-MO" Then trnLng = "zh-TW"
		End If
		EngineID = DlgValue("SelSetBox")
		TempArray = Split(EngineDataList(EngineID),JoinStr)
		LangArray = Split(TempArray(2),SubJoinStr)
		n = 0
		m = 0
		ReDim LngNameList(n)
		For i = 0 To UBound(LangArray)
			LangPairList = Split(LangArray(i),LngJoinStr)
			If LangPairList(2) <> "" Then
				ReDim Preserve LngNameList(n)
				LngNameList(n) = LangPairList(0)
				If LangPairList(1) = trnLng Then m = n
				n = n + 1
			End If
		Next i
		DlgListBoxArray "LngNameBox",LngNameList()
		DlgValue "LngNameBox",m
		DlgText "TrnListBox",TempDec
		DlgText "LineNumBox",CStr(LineNum)
		DlgText "InTextBox",inText
		If DlgText("InTextBox") <> "" Then DlgText "ClearButton",Msg01
    	If DlgText("InTextBox") = "" Then DlgText "ClearButton",Msg02
    	If DlgText("InTextBox") <> "" Then DlgEnable "ClearButton",True
    	If DlgText("InTextBox") = "" Then DlgEnable "ClearButton",False
    	If DlgText("InTextBox") <> "" Then DlgEnable "HeaderButton",True
		If DlgText("InTextBox") = "" Then DlgEnable "HeaderButton",False
	Case 2 ' 杅硉載蜊麼氪偌狟賸偌聽
		If DlgItem$ = "SelSetBox" Then
			LangName = DlgText("LngNameBox")
			EngineID = DlgValue("SelSetBox")
			TempArray = Split(EngineDataList(EngineID),JoinStr)
			LangArray = Split(TempArray(2),SubJoinStr)
			n = 0
			ReDim LngNameList(n)
			For i = 0 To UBound(LangArray)
				LangPairList = Split(LangArray(i),LngJoinStr)
				If LangPairList(2) <> "" Then
					ReDim Preserve LngNameList(n)
					LngNameList(n) = LangPairList(0)
					n = n + 1
				End If
			Next i
			DlgListBoxArray "LngNameBox",LngNameList()
			DlgText "LngNameBox",LangName
			If DlgText("LngNameBox") = "" Then DlgValue "LngNameBox",0
		End If

		If DlgItem$ = "TrnListBox" Or DlgItem$ = "StrType" Or DlgItem$ = "ClearButton" Then
			If DlgItem$ <> "ClearButton" Or DlgText("ClearButton") = Msg02 Then
				ListDec = DlgText("TrnListBox")
				For i = 1 To trn.Project.TransLists.Count
					Set TrnList = trn.Project.TransLists(i)
					TrnListDec = TrnList.Title & " - " & PSL.GetLangCode(TrnList.Language.LangID,pslCodeText)
					If TrnListDec = ListDec Then Exit For
				Next i
				If TrnList.SourceList.LastChange > TrnList.LastUpdate Then TrnList.Update
				LineNum = CLng(DlgText("LineNumBox"))
				If LineNum > TrnList.StringCount Then LineNum = TrnList.StringCount
				n = 0
				For i = 1 To TrnList.StringCount
					Set TransString = TrnList.String(i)
					If TransString.Text <> "" Then
						If DlgValue("StrType") = 0 Then srcString = TransString.SourceText
						If DlgValue("StrType") = 1 Then srcString = TransString.Text
						If inText <> "" Then inText = inText & " " & srcString
						If inText = "" Then inText = srcString
						n = n + 1
					End If
					If n = LineNum Then Exit For
				Next i
				DlgText "InTextBox",inText
			End If
		End If

		If DlgItem$ = "TestButton" Or DlgItem$ = "StrType" Or DlgItem$ = "TranType" Or DlgItem$ = "HeaderButton" Then
			EngineID = DlgValue("SelSetBox")
			inText = DlgText("InTextBox")
			ListDec = DlgText("TrnListBox")
			DlgText "OutTextBox",""
			DlgText "OutTextBox",Msg05

			'趼揹蘇�珆仍池�
			inText = AccessKeyHanding(0,inText)
			inText = AcceleratorHanding(0,inText)

			'鳳�◎倡蹅訇穔釋椒景苃膨縎擿�
			For i = 1 To trn.Project.TransLists.Count
				Set TrnList = trn.Project.TransLists(i)
				TrnListDec = TrnList.Title & " - " & PSL.GetLangCode(TrnList.Language.LangID,pslCodeText)
				If TrnListDec = ListDec Then Exit For
			Next i
			If DlgValue("StrType") = 0 Then
				srcLng = PSL.GetLangCode(TrnList.SourceList.LangID,pslCode639_1)
			Else
				srcLng = PSL.GetLangCode(TrnList.Language.LangID,pslCode639_1)
			End If
			If srcLng = "zh" Then
				If DlgValue("StrType") = 0 Then
					srcLng = PSL.GetLangCode(TrnList.SourceList.LangID,pslCodeLangRgn)
				Else
					srcLng = PSL.GetLangCode(TrnList.Language.LangID,pslCodeLangRgn)
				End If
				If srcLng = "zh-CHS" Or srcLng = "zh-SG" Then srcLng = "zh-CN"
				If srcLng = "zh-CHT" Or srcLng = "zh-HK" Or srcLng = "zh-MO" Then srcLng = "zh-TW"
			End If
			trnLng = DlgText("LngNameBox")

			'鳳�◎倡遻�э腔懂埭睿醴梓逄晟
			srcLngFind = 0
			trnLngFind = 0
			TempArray = Split(EngineDataList(EngineID),JoinStr)
			LangArray = Split(TempArray(2),SubJoinStr)
			For i = 0 To UBound(LangArray)
				LangPairList = Split(LangArray(i),LngJoinStr)
				If LCase(srcLng) = LCase(LangPairList(1)) Then
					srcLng = LangPairList(2)
					srcLngFind = 1
				End If
				If LCase(trnLng) = LCase(LangPairList(0)) Then
					trnLng = LangPairList(2)
					trnLngFind = 1
				End If
				If srcLngFind + trnLngFind = 2 Then Exit For
			Next i
			LangPair = srcLng & LngJoinStr & trnLng

			'羲宎楹祒甜怀堤
			Msg = Msg05 & vbCrLf & Msg06 & srcLng & " > " & trnLng
			DlgText "OutTextBox",Msg
			Set xmlHttp = CreateObject(DefaultObject)
			If xmlHttp Is Nothing Then Set xmlHttp = CreateObject("Msxml2.XMLHTTP")
			If DlgItem$ = "HeaderButton" Then
				outText = getTranslate(EngineID,xmlHttp,inText,LangPair,2)
			Else
				outText = getTranslate(EngineID,xmlHttp,inText,LangPair,DlgValue("TranType"))
				'Temp = EngineID & JoinStr & 0 & JoinStr & 1 & JoinStr & 1
				'outText = SplitTran(xmlHttp,inText,LangPair,Temp,DlgValue("TranType"))
			End If
			Set xmlHttp = Nothing
			If outText = "" Then
				DlgText "OutTextBox",Msg & vbCrLf & Msg09 & vbCrLf & Msg07
			ElseIf Trim(Replace(outText,vbCrLf,"")) = "" Then
				DlgText "OutTextBox",Msg & vbCrLf & Msg09 & vbCrLf & Msg08
			Else
				DlgText "OutTextBox",Msg & vbCrLf & Msg09 & vbCrLf & outText
			End If
		End If

		If DlgItem$ = "HelpButton" Then
			Call EngineHelp("TestHelp")
		End If
		If DlgItem$ = "ClearButton" And DlgText("ClearButton") = Msg01 Then
			DlgText "InTextBox",""
			DlgText "OutTextBox",""
			If DlgText("InTextBox") <> "" Then DlgText "ClearButton",Msg01
    		If DlgText("InTextBox") = "" Then DlgText "ClearButton",Msg02
		End If

		If DlgItem$ <> "CancelButton" Then
			TranTestFunc = True '滅砦偌狟偌聽壽敕勤趕遺敦諳
		End If
		If DlgText("InTextBox") <> "" Then DlgText "ClearButton",Msg01
    	If DlgText("InTextBox") = "" Then DlgText "ClearButton",Msg02
    	If DlgText("InTextBox") <> "" Then DlgEnable "TestButton",True
		If DlgText("InTextBox") = "" Then DlgEnable "TestButton",False
    	If DlgText("InTextBox") <> "" Then DlgEnable "HeaderButton",True
		If DlgText("InTextBox") = "" Then DlgEnable "HeaderButton",False
	Case 3 ' 恅掛遺麼氪郪磁遺恅掛掩載蜊
		If DlgItem$ = "LineNumBox" Then
			DlgText "InTextBox",""
			ListDec = DlgText("TrnListBox")
			LineNum = CLng(DlgText("LineNumBox"))
			For i = 1 To trn.Project.TransLists.Count
				Set TrnList = trn.Project.TransLists(i)
				TrnListDec = TrnList.Title & " - " & PSL.GetLangCode(TrnList.Language.LangID,pslCodeText)
				If TrnListDec = ListDec Then Exit For
			Next i
			If LineNum > TrnList.StringCount Then LineNum = TrnList.StringCount
			n = 0
			For i = 1 To TrnList.StringCount
				Set TransString = TrnList.String(i)
				If TransString.Text <> "" Then
					srcString = TransString.SourceText
					If inText <> "" Then inText = inText & " " & srcString
					If inText = "" Then inText = srcString
					n = n + 1
				End If
				If n = LineNum Then Exit For
			Next i
			DlgText "InTextBox",inText
		End If
		If DlgText("InTextBox") <> "" Then DlgText "ClearButton",Msg01
    	If DlgText("InTextBox") = "" Then DlgText "ClearButton",Msg02
    	If DlgText("InTextBox") <> "" Then DlgEnable "TestButton",True
		If DlgText("InTextBox") = "" Then DlgEnable "TestButton",False
    	If DlgText("InTextBox") <> "" Then DlgEnable "HeaderButton",True
		If DlgText("InTextBox") = "" Then DlgEnable "HeaderButton",False
	End Select
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


'湖羲恅璃
Function OpenFile(FilePath As String,FileList() As String,x As Integer,Stemp As Boolean) As Boolean
	Dim ExePathStr As String,Argument As String
	If OSLanguage = "0404" Then
		Msg01 = "錯誤"
		Msg02 = "記事本未在系統中找到。請選取其他開啟方法。"
		Msg03 = "系統沒有安裝 Excel 套用程式。請選取其他開啟方法。"
		Msg04 = "程式未找到！可能是程式路徑剖析錯誤或不被支援，請選取其他開啟方法。"
		Msg05 = "無法開啟檔案！套用程式返回了錯誤代碼，請選取其他開啟方法。"
		Msg06 = "無法開啟檔案！套用程式返回了錯誤代碼，可能是該檔案不被支援或執行參數有問題。"
		Msg07 = "程式名稱: "
		Msg08 = "剖析路徑: "
		Msg09 = "執行參數: "
	Else
		Msg01 = "渣昫"
		Msg02 = "暮岈掛帤婓炵苀笢梑善﹝③恁寁む坻湖羲源楊﹝"
		Msg03 = "炵苀羶衄假蚾 Excel 茼蚚最唗﹝③恁寁む坻湖羲源楊﹝"
		Msg04 = "最唗帤梑善ㄐ褫夔岆最唗繚噤賤昴渣昫麼祥掩盓厥ㄛ③恁寁む坻湖羲源楊﹝"
		Msg05 = "拸楊湖羲恅璃ㄐ茼蚚最唗殿隙賸渣昫測鎢ㄛ③恁寁む坻湖羲源楊﹝"
		Msg06 = "拸楊湖羲恅璃ㄐ茼蚚最唗殿隙賸渣昫測鎢ㄛ褫夔岆蜆恅璃祥掩盓厥麼堍俴統杅衄恀枙﹝"
		Msg07 = "最唗靡備: "
		Msg08 = "賤昴繚噤: "
		Msg09 = "堍俴統杅: "
	End If

	OpenFile = False
	File = FilePath
	If x = 0 Then
		Call Edit(File,FileList)
		OpenFile = True
	ElseIf x = 1 Then
		Set WshShell = CreateObject("WScript.Shell")
		If Not WshShell Is Nothing Then
			If Dir(Environ("SystemRoot") & "\system32\notepad.exe") = "" Then
				If Dir(Environ("SystemRoot") & "\notepad.exe") = "" Then
					MsgBox Msg02,vbOkOnly+vbInformation,Msg01
				Else
					ExePath = "%SystemRoot%\notepad.exe"
				End If
			Else
				ExePath = "%SystemRoot%\system32\notepad.exe"
			End If
			If ExePath <> "" Then
				File = """" & File & """"
				Return = WshShell.Run("""" & ExePath & """ " & File,1,Stemp)
				If Return <> 0 Then
					MsgBox Msg05,vbOkOnly+vbInformation,Msg01
				Else
					OpenFile = True
				End If
			End If
			Set WshShell = Nothing
		End If
	ElseIf x = 2 Then
		Set WshShell = CreateObject("WScript.Shell")
		If Not WshShell Is Nothing Then
			ExtName = ".xls"
			On Error Resume Next
			strKeyPath = "HKCR\" & ExtName & "\"
			ExtCmdStr = WshShell.RegRead(strKeyPath)
			If ExtCmdStr <> "" Then
				ExePathStr = WshShell.RegRead("HKCR\" & ExtCmdStr & "\shell\edit\command\")
				If ExePathStr = "" Then
					ExePathStr = WshShell.RegRead("HKCR\" & ExtCmdStr & "\shell\open\command\")
				End If
				If ExePathStr = "" Then
					ExePathStr = WshShell.RegRead("HKCR\" & ExtCmdStr & "\shell\preview\command\")
				End If
			End If
			On Error GoTo 0
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
				ExePath = PreExePath & AppExePath
				Argument = Mid(ExePathStr,Len(ExePath)+1)
				ExePath = RemoveBackslash(ExePath,"""","""",1)

				If InStr(ExePath,"\") = 0 Then
					If Dir(Environ("SystemRoot") & "\system32\" & ExePath) <> "" Then
						ExePath = Environ("SystemRoot") & "\system32\" & ExePath
					ElseIf Dir(Environ("SystemRoot") & "\" & ExePath) <> "" Then
						ExePath = Environ("SystemRoot") & "\" & ExePath
					End If
				End If

				If Argument <> "" Then
					If InStr(Argument,"%1") Then
						ArgumentFile = Replace(Argument,"%1",File)
					ElseIf InStr(Argument,"%L") Then
						ArgumentFile = Replace(Argument,"%L",File)
					Else
						ArgumentFile = Argument & " " & """" & File & """"
					End If
				Else
					ArgumentFile = """" & File & """"
				End If
			End If
			If ExePath <> "" Then
				On Error Resume Next
				strKeyPath = "HKCU\Software\Microsoft\Windows\ShellNoRoam\MUICache\" & ExePath
				ExeName = WshShell.RegRead(strKeyPath)
				If ExeName = "" Then
					ExeName = Mid(ExePath,InStrRev(ExePath,"\")+1)
				End If
				On Error GoTo 0
			End If
			If ExePath <> "" And Dir(ExePath) <> "" Then
				Return = WshShell.Run("""" & ExePath & """ " & ArgumentFile,1,Stemp)
				If Return <> 0 Then
					MsgBox Msg05,vbOkOnly+vbInformation,Msg01
				Else
					'If LCase(ExeName) <> "notepad.exe" And InStr(Tools02,ExeName) = 0 Then
					'	Call AddArray(AppNames,AppPaths,ExeName,ExePath & JoinStr & Argument)
					'End If
					OpenFile = True
				End If
			End If
			If ExePath = "" Then MsgBox Msg03,vbOkOnly+vbInformation,Msg01
			If ExePath <> "" And Dir(ExePath) = "" Then
				Msg = Msg07 & ExeName & vbCrLf & Msg08 & ExePath & vbCrLf & Msg09 & _
				Argument & vbCrLf & vbCrLf & Msg04
				MsgBox Msg,vbOkOnly+vbInformation,Msg01
			End If
			Set WshShell = Nothing
		End If
  	ElseIf x = 3 Then
		Set WshShell = CreateObject("WScript.Shell")
		If Not WshShell Is Nothing Then
			Call CmdInput(ExePathStr,Argument)
			If ExePathStr <> "" Then
				If UBound(Split(ExePathStr,"%",-1)) >= 2 Then
					AppExePathStr = Mid(ExePathStr,InStr(ExePathStr,"%")+1)
					Env = Left(AppExePathStr,InStr(AppExePathStr,"%")-1)
					ExePathStr = Replace(ExePathStr,"%" & Env & "%",Environ(Env),,1)
				End If
				ExePath = RemoveBackslash(ExePathStr,"""","""",1)
			End If
			If Argument <> "" Then
				If InStr(Argument,"%1") Then
					ArgumentFile = Replace(Argument,"%1",File)
				ElseIf InStr(Argument,"%L") Then
					ArgumentFile = Replace(Argument,"%L",File)
				Else
					ArgumentFile = Argument & " " & """" & File & """"
				End If
			Else
				ArgumentFile = """" & File & """"
			End If
			If ExePath <> "" And Dir(ExePath) <> "" Then
				Return = WshShell.Run("""" & ExePath & """ " & ArgumentFile,1,Stemp)
				If Return <> 0 Then
					MsgBox Msg06,vbOkOnly+vbInformation,Msg01
				Else
					ExeName = Mid(ExePath,InStrRev(ExePath,"\")+1)
					If LCase(ExeName) <> "notepad.exe" And InStr(Tools02,ExeName) = 0 Then
						Call AddArray(AppNames,AppPaths,ExeName,ExePath & JoinStr & Argument)
					End If
					OpenFile = True
				End If
			End If
			If ExePath <> "" And Dir(ExePath) = "" Then
				MsgBox ExeName & Msg04,vbOkOnly+vbInformation,Msg01
			End If
			Set WshShell = Nothing
		End If
	ElseIf x > 3 Then
		Set WshShell = CreateObject("WScript.Shell")
		If Not WshShell Is Nothing Then
			ExeName = AppNames(x)
			ExePathArgument = AppPaths(x)
			ExePathArray = Split(ExePathArgument,JoinStr,-1)
			ExePath = ExePathArray(0)
			Argument = ExePathArray(1)
			If Argument <> "" Then
				If InStr(Argument,"%1") Then
					ArgumentFile = Replace(Argument,"%1",File)
				ElseIf InStr(Argument,"%L") Then
					ArgumentFile = Replace(Argument,"%L",File)
				Else
					ArgumentFile = Argument & " " & """" & File & """"
				End If
			Else
				ArgumentFile = """" & File & """"
			End If
			If ExePath <> "" And Dir(ExePath) <> "" Then
				Return = WshShell.Run("""" & ExePath & """ " & ArgumentFile,1,Stemp)
				If Return <> 0 Then
					MsgBox Msg05,vbOkOnly+vbInformation,Msg01
				Else
					OpenFile = True
				End If
			End If
			If ExePath = "" Then MsgBox Msg03,vbOkOnly+vbInformation,Msg01
			If ExePath <> "" And Dir(ExePath) = "" Then
				MsgBox ExeName & Msg04,vbOkOnly+vbInformation,Msg01
			End If
			Set WshShell = Nothing
		End If
	End If
End Function


'晤憮恅璃
Sub Edit(File As String,FileList() As String)
	If OSLanguage = "0404" Then
		Msg01 = "編輯"
		Msg02 = "檔案 - "
		Msg03 = "字元編碼:"
		Msg05 = "搜尋內容:"
		Msg06 = "搜尋"
		Msg10 = "讀入(&R)"
		Msg12 = "上一個(&P)"
		Msg13 = "下一個(&N)"
		Msg14 = "儲存(&S)"
		Msg16 = "結束搜尋模式"
	Else
		Msg01 = "晤憮"
		Msg02 = "恅璃 - "
		Msg03 = "趼睫晤鎢:"
		Msg05 = "脤梑囀��:"
		Msg06 = "脤梑"
		Msg10 = "黍��(&R)"
		Msg12 = "奻珨跺(&P)"
		Msg13 = "狟珨跺(&N)"
		Msg14 = "悵湔(&S)"
		Msg16 = "豖堤脤梑耀宒"
	End If

    'Dim objStream As Object
	'Set objStream = CreateObject("Adodb.Stream")
	'If objStream Is Nothing Then CodeList = CodePageList(0,0)
	'If Not objStream Is Nothing Then CodeList = CodePageList(0,49)
	'Set objStream = Nothing

	Dim CodeNameList() As String,i As Integer
	For i = LBound(CodeList) To UBound(CodeList)
		ReDim Preserve CodeNameList(i)
		TempArray = Split(CodeList(i),JoinStr)
		CodeNameList(i) = TempArray(0)
	Next i

	Begin Dialog UserDialog 820,504,Msg01,.EditFunc ' %GRID:10,7,1,1
		Text 10,7,800,14,File,.FileName,2
		Text 510,28,80,14,Msg03,.CodeText
		DropListBox 600,24,210,21,CodeNameList(),.CodeNameList
		TextBox 0,49,820,420,.InTextBox,1
		Text 10,28,80,14,Msg05,.FindText
		TextBox 100,25,310,19,.FindBox
		PushButton 420,24,80,21,Msg06,.FindButton
		PushButton 20,476,90,21,Msg10,.ReadButton
		PushButton 120,476,90,21,Msg12,.PreviousButton
		PushButton 220,476,90,21,Msg13,.NextButton
		PushButton 600,476,100,21,Msg14,.SaveButton
		PushButton 380,476,140,21,Msg16,.EditButton
		CancelButton 710,476,90,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	If Dialog(dlg) = 0 Then Exit Sub
End Sub


'晤憮勤趕遺滲杅
Private Function EditFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim File As String,inText As String,outText As String
	Dim Code As String,CodeID As Integer,m As Integer,n As Integer,j As Integer

	If OSLanguage = "0404" Then
		Msg01 = "訊息"
		Msg02 = "檔案內容已被變更，是否需要儲存？"
		Msg03 = "找到的內容已被變更，是否需要替換原來內容後顯示？"
		Msg04 = "行數有變化！在搜尋模式下，行號不可刪除，內容可刪除和修改，請修改後再試。"
		Msg05 = "未找到指定內容。"
		Msg06 = "檔案儲存成功！"
		Msg07 = "檔案儲存失敗！請檢查檔案是否正被開啟。"
		Msg08 = "檔案內容未被變更，不需要儲存！"
		Msg10 = "讀入(&R)"
		Msg11 = "清空(&C)"
		LineNo = "行"
	Else
		Msg01 = "陓洘"
		Msg02 = "恅璃囀�椹拲遘�蜊ㄛ岆瘁剒猁悵湔ˋ"
		Msg03 = "梑善腔囀�椹拲遘�蜊ㄛ岆瘁剒猁杸遙埻懂囀�搡鯔埰麾�"
		Msg04 = "俴杅衄曹趙ㄐ婓脤梑耀宒狟ㄛ俴瘍祥褫刉壺ㄛ囀�暆圪噫�睿党蜊ㄛ③党蜊綴婬彸﹝"
		Msg05 = "帤梑善硌隅囀�搳�"
		Msg06 = "恅璃悵湔傖髡ㄐ"
		Msg07 = "恅璃悵湔囮啖ㄐ③潰脤恅璃岆瘁淏掩湖羲﹝"
		Msg08 = "恅璃囀�楱敢遘�蜊ㄛ祥剒猁悵湔ㄐ"
		Msg10 = "黍��(&R)"
		Msg11 = "ь諾(&C)"
		LineNo = "俴"
	End If

	Select Case Action%
	Case 1
		File = DlgText("FileName")
		Code = CheckCode(File)
		For i = LBound(CodeList) To UBound(CodeList)
			TempArray = Split(CodeList(i),JoinStr)
			If TempArray(1) = Code Then
 				DlgValue "CodeNameList",i
 				Exit For
 			End If
		Next i
		FileText = ReadFile(File,Code)
		If FileText <> "" Then
			DlgText "InTextBox",FileText
			DlgText "ReadButton",Msg11
    		DlgEnable "FindButton",True
			DlgEnable "SaveButton",True
			DlgEnable "EditButton",False
    	Else
    		DlgText "ReadButton",Msg10
    		DlgEnable "FindButton",False
    		DlgEnable "SaveButton",False
    		DlgEnable "EditButton",False
    	End If
		For i = LBound(FileList) To UBound(FileList)
			If InStr(FileList(i),File) Then
				FileNo = i
				Exit For
			End If
		Next i
		If UBound(FileList) = 0 Then
			DlgEnable "PreviousButton",False
			DlgEnable "NextButton",False
		ElseIf FileNo = 0 Then
			DlgEnable "PreviousButton",False
			If FileList(FileNo+1) = "" Then DlgEnable "NextButton",False
			If FileList(FileNo+1) <> "" Then DlgEnable "NextButton",True
		ElseIf FileNo = UBound(FileList) Then
			If FileList(FileNo-1) = "" Then DlgEnable "PreviousButton",False
			If FileList(FileNo-1) <> "" Then DlgEnable "PreviousButton",True
			DlgEnable "NextButton",False
		Else
			If FileList(FileNo-1) = "" Then DlgEnable "PreviousButton",False
			If FileList(FileNo-1) <> "" Then DlgEnable "PreviousButton",True
			If FileList(FileNo+1) = "" Then DlgEnable "NextButton",False
			If FileList(FileNo+1) <> "" Then DlgEnable "NextButton",True
    	End If
	Case 2 ' 杅硉載蜊麼氪偌狟賸偌聽
		If DlgItem$ = "CodeNameList" Then
			File = DlgText("FileName")
			CodeID = DlgValue("CodeNameList")
			TempArray = Split(CodeList(CodeID),JoinStr)
			Code = TempArray(1)
			If Code = "_autodetect_all" Or Code = "_autodetect" Or Code = "_autodetect_kr" Then
				Code = CheckCode(File)
				For i = LBound(CodeList) To UBound(CodeList)
					TempArray = Split(CodeList(i),JoinStr)
					If TempArray(1) = Code Then
 						DlgValue "CodeNameList",i
 						Exit For
 					End If
				Next i
				FileText = ReadFile(File,Code)
			Else
				FileText = ReadFile(File,Code)
			End If
			If FileText <> "" Then
				DlgText "InTextBox",FileText
				FindText = ""
				FindLine = ""
			End If
		End If

		If DlgItem$ = "ReadButton" Then
			If DlgText("ReadButton") = Msg10 Then
				File = DlgText("FileName")
				CodeID = DlgValue("CodeNameList")
				TempArray = Split(CodeList(CodeID),JoinStr)
				Code = TempArray(1)
				FileText = ReadFile(File,Code)
				If FileText <> "" Then
					DlgText "InTextBox",FileText
					DlgText "ReadButton",Msg11
					FindText = ""
					FindLine = ""
				End If
			Else
				DlgText "InTextBox",""
				DlgText "ReadButton",Msg10
				FindText = ""
				FindLine = ""
			End If
		End If

		If DlgItem$ = "FindButton" And DlgText("FindBox") <> "" Then
			FindText = ""
			FindLine = ""
			n = 0
			toFindText = "*" & DlgText("FindBox") & "*"
			InTextArray = Split(FileText,vbCrLf,-1)
			For i = LBound(InTextArray) To UBound(InTextArray)
				tempText = InTextArray(i)
				If tempText Like toFindText Then
					Temp = "▽" & i+1 & LineNo & "▼" & tempText
					If FindText <> "" Then FindText = FindText & vbCrLf & Temp
					If FindText = "" Then FindText =  Temp
					Temp = CStr(i) & "," & CStr(n)
					If FindLine <> "" Then FindLine = FindLine & JoinStr & Temp
					If FindLine = "" Then FindLine = Temp
					n = n + 1
				End If
			Next i
			If FindText <> "" Then
				DlgText "InTextBox",FindText
			Else
				DlgText "InTextBox",Msg05
			End If
    	End If

		If DlgItem$ = "EditButton" And DlgText("InTextBox") <> "" Then
			inText = DlgText("InTextBox")
			If FindLine <> "" And inText <> FindText Then
				If MsgBox(Msg03,vbYesNo+vbInformation,Msg01) = vbYes Then
					InTextArray = Split(FileText,vbCrLf,-1)
					OutTextArray = Split(inText,vbCrLf,-1)
					FindLineArray = Split(FindLine,JoinStr,-1)
					If UBound(FindLineArray) = UBound(OutTextArray) Then
						For i = LBound(FindLineArray) To UBound(OutTextArray)
							LineArray = Split(FindLineArray(i),",",-1)
							OldLineNum = LineArray(0)
							NewLineNum = LineArray(1)
							OldString = InTextArray(OldLineNum)
							NewString = OutTextArray(NewLineNum)
							Temp = LineNo & "▼"
							LineNoStr = Left(NewString,InStr(NewString,Temp)+1)
							Temp = "▽" & "*" & LineNo & "▼"
							If LineNoStr Like Temp Then
								NewString = Mid(NewString,Len(LineNoStr)+1)
								InTextArray(OldLineNum) = NewString
							End If
						Next i
						inText = Join(InTextArray,vbCrLf)
						DlgText "InTextBox",inText
						FindText = ""
						FindLine = ""
					Else
						MsgBox(Msg04,vbOkOnly+vbInformation,Msg01)
					End If
				Else
				DlgText "InTextBox",FileText
				FindText = ""
				FindLine = ""
				End If
			Else
				DlgText "InTextBox",FileText
				FindText = ""
				FindLine = ""
			End If
		End If

		If DlgItem$ = "PreviousButton" Or DlgItem$ = "NextButton" Then
			j = FileNo
			If DlgItem$ = "PreviousButton" And FileNo <> 0 Then FileNo = FileNo - 1
			If DlgItem$ = "NextButton" And FileNo < UBound(FileList) Then FileNo = FileNo + 1
			If FileNo <> j Then
				File = FileList(FileNo)
				DlgText "FileName",File
				Code = CheckCode(File)
				FileText = ReadFile(File,Code)
				If FileText <> "" Then
					DlgText "InTextBox",FileText
					For i = LBound(CodeList) To UBound(CodeList)
						TempArray = Split(CodeList(i),JoinStr)
						If TempArray(1) = Code Then
 							DlgValue "CodeNameList",i
 							Exit For
 						End If
					Next i
					FindText = ""
					FindLine = ""
				End If
    		End If
    	End If

		If DlgItem$ = "SaveButton" And DlgText("InTextBox") <> "" Then
			File = DlgText("FileName")
			inText = DlgText("InTextBox")
			If inText <> FileText Then
				CodeID = DlgValue("CodeNameList")
				TempArray = Split(CodeList(CodeID),JoinStr)
				Code = TempArray(1)
				If Dir(File) <> "" Then SetAttr File,vbNormal
				If WriteToFile(File,inText,Code) = True Then
					MsgBox(Msg06,vbOkOnly+vbInformation,Msg01)
					FileText = inText
					FindText = ""
					FindLine = ""
				Else
					MsgBox(Msg07,vbOkOnly+vbInformation,Msg01)
				End If
			Else
				MsgBox(Msg06,vbOkOnly+vbInformation,Msg01)
			End If
    	End If

		If DlgItem$ = "CancelButton" And DlgText("InTextBox") <> "" Then
			File = DlgText("FileName")
			inText = DlgText("InTextBox")
			If inText <> FileText Then
				If MsgBox(Msg02,vbYesNo+vbInformation,Msg01) = vbYes Then
					CodeID = DlgValue("CodeNameList")
					TempArray = Split(CodeList(CodeID),JoinStr)
					Code = TempArray(1)
					If Dir(File) <> "" Then SetAttr File,vbNormal
					If WriteToFile(File,inText,Code) = True Then
						MsgBox(Msg06,vbOkOnly+vbInformation,Msg01)
						FileText = inText
						FindText = ""
						FindLine = ""
					Else
						MsgBox(Msg07,vbOkOnly+vbInformation,Msg01)
					End If
				End If
			End If
    	End If

		If UBound(FileList) = 0 Then
			DlgEnable "PreviousButton",False
			DlgEnable "NextButton",False
		ElseIf FileNo = 0 Then
			DlgEnable "PreviousButton",False
			If FileList(FileNo+1) = "" Then DlgEnable "NextButton",False
			If FileList(FileNo+1) <> "" Then DlgEnable "NextButton",True
		ElseIf FileNo = UBound(FileList) Then
			If FileList(FileNo-1) = "" Then DlgEnable "PreviousButton",False
			If FileList(FileNo-1) <> "" Then DlgEnable "PreviousButton",True
			DlgEnable "NextButton",False
		Else
			If FileList(FileNo-1) = "" Then DlgEnable "PreviousButton",False
			If FileList(FileNo-1) <> "" Then DlgEnable "PreviousButton",True
			If FileList(FileNo+1) = "" Then DlgEnable "NextButton",False
			If FileList(FileNo+1) <> "" Then DlgEnable "NextButton",True
    	End If

		If DlgItem$ <> "CancelButton" Then
		    inText = DlgText("InTextBox")
		    If inText <> "" Then
    			DlgText "ReadButton",Msg11
    			DlgEnable "FindButton",True
				If FindLine = "" Then
					DlgEnable "SaveButton",True
					DlgEnable "EditButton",False
					DlgEnable "CancelButton",True
				Else
					DlgEnable "SaveButton",False
					DlgEnable "EditButton",True
					DlgEnable "CancelButton",False
				End If
    		Else
    			DlgText "ReadButton",Msg10
    			DlgEnable "FindButton",False
    			DlgEnable "SaveButton",False
    			DlgEnable "EditButton",False
    			DlgEnable "CancelButton",True
    		End If
			EditFunc = True '滅砦偌狟偌聽壽敕勤趕遺敦諳
		End If
	Case 3 ' 恅掛遺麼氪郪磁遺恅掛掩載蜊
		If DlgItem$ = "InTextBox" Then
		    inText = DlgText("InTextBox")
		    If inText <> "" Then
    			DlgText "ReadButton",Msg11
    			DlgEnable "FindButton",True
				If FindLine = "" Then
					DlgEnable "SaveButton",True
					DlgEnable "EditButton",False
				Else
					DlgEnable "SaveButton",False
					DlgEnable "EditButton",True
				End If
    		Else
    			DlgText "ReadButton",Msg10
    			DlgEnable "FindButton",False
    			DlgEnable "SaveButton",False
    			DlgEnable "EditButton",False
    		End If
    	End If
	End Select
End Function


'氝樓祥眈肮腔杅郪啋匼
Sub AddArray(AppNames() As String,AppPaths() As String,CmdName As String,CmdPath As String)
	Dim i As Integer,n As Integer,Stemp As Boolean
	If CmdName = "" And CmdPath = "" Then Exit Sub
	n = UBound(AppNames)
	Stemp = False
	For i = LBound(AppNames) To UBound(AppNames)
		If InStr(LCase(AppNames(i)),LCase(CmdName)) Then
			Stemp = True
			Exit For
		End If
	Next i
	If Stemp = False Then
		ReDim Preserve AppNames(n+1)
		ReDim Preserve AppPaths(n+1)
		AppNames(n+1) = LCase(CmdName)
		AppPaths(n+1) = LCase(CmdPath)
	End If
End Sub


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


'怀�貑鉏韋昐�
Sub CmdInput(CmdPath As String,Argument As String)
	If OSLanguage = "0404" Then
		Msg01 = "自訂編輯程式"
		Msg02 = "請指定編輯程式及其執行參數 (檔案參照參數和其他參數)。" & vbCrLf & vbCrLf & _
				"注意: " & vbCrLf & _
				"- 如果執行參數中檔案參照參數需要在其他參數前面的話，請點擊右邊的按鈕輸入，" & _
				"  或直接輸入檔案參照符 %1，否則可以不輸入檔案參照參數。" & vbCrLf & _
				"- 檔案參照符 %1 欄位為系統參數，不可變更為其他符號。"
		Msg03 = "編輯程式 (支援環境變數，變數名前後請附加 % 符號):"
		Msg04 = "..."
		Msg05 = "執行參數 (如果程式支援並需要的話):"
		Msg06 = "清空(&K)"
		Msg09 = ">"
	Else
		Msg01 = "赻隅砱晤憮最唗"
		Msg02 = "③硌隅晤憮最唗摯む堍俴統杅 (恅璃竘蚚統杅睿む坻統杅)﹝" & vbCrLf & vbCrLf & _
				"蛁砩: " & vbCrLf & _
				"- �蝜�堍俴統杅笢恅璃竘蚚統杅剒猁婓む坻統杅ヶ醱腔趕ㄛ③等僻衵晚腔偌聽怀�諴�" & _
				"  麼眻諉怀�輷躁�竘蚚睫 %1ㄛ瘁寀褫眕祥怀�輷躁�竘蚚統杅﹝" & vbCrLf & _
				"- 恅璃竘蚚睫 %1 趼僇峈炵苀統杅ㄛ祥褫載蜊峈む坻睫瘍﹝"
		Msg03 = "晤憮最唗 (盓厥遠噫曹講ㄛ曹講靡ヶ綴③蜇樓 % 睫瘍):"
		Msg04 = "..."
		Msg05 = "堍俴統杅 (�蝜�最唗盓厥甜剒猁腔趕):"
		Msg06 = "ь諾(&K)"
		Msg09 = ">"
	End If
	Begin Dialog UserDialog 510,266,Msg01,.CmdInputFunc ' %GRID:10,7,1,1
		Text 10,7,490,98,Msg02
		Text 10,119,460,14,Msg03
		TextBox 10,140,460,21,.CmdPath
		PushButton 470,140,30,21,Msg04,.BrowseButton
		Text 10,175,460,14,Msg05
		TextBox 10,196,460,21,.Argument
		PushButton 470,196,30,21,Msg09,.FileArgButton
		PushButton 20,238,90,21,Msg06,.ClearButton
		OKButton 280,238,100,21,.OKButton
		CancelButton 390,238,100,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	dlg.CmdPath = CmdPath
	dlg.Argument = Argument
	If Dialog(dlg) = 0 Then Exit Sub
	CmdPath = dlg.CmdPath
	Argument = dlg.Argument
End Sub


'鳳�§鉏韋昐繲堇倏罊缺�
Private Function CmdInputFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim File As String,Items(0) As String,x As Integer
	If OSLanguage = "0404" Then
		Msg01 = "錯誤"
		Msg02 = "選取編輯程式"
		Msg03 = "可執行檔案 (*.exe)|*.exe|所有檔案 (*.*)|*.*||"
		Msg04 = "沒有指定編輯程式！請重新輸入或選取。"
		FileArg = "檔案參照參數(%1)"
	Else
		Msg01 = "渣昫"
		Msg02 = "恁寁晤憮最唗"
		Msg03 = "褫硒俴恅璃 (*.exe)|*.exe|垀衄恅璃 (*.*)|*.*||"
		Msg04 = "羶衄硌隅晤憮最唗ㄐ③笭陔怀�趥藘√鞢�"
		FileArg = "恅璃竘蚚統杅(%1)"
	End If
	Items(0) = FileArg
	Select Case Action%
	Case 1 ' 勤趕遺敦諳場宎趙
		DlgEnable "ClearButton",False
	Case 2 ' 杅硉載蜊麼氪偌狟賸偌聽
		If DlgItem$ = "BrowseButton" Then
			If PSL.SelectFile(File,True,Msg03,Msg02) = True Then
 				DlgText "CmdPath",File
 			Else
 				DlgText "CmdPath",""
 			End If
		End If
		If DlgItem$ = "ClearButton" Then
 			DlgText "CmdPath",""
 			DlgText "Argument",""
 			DlgEnable "ClearButton",False
		End If
		If DlgItem$ = "FileArgButton" Then
			x = ShowPopupMenu(Items)
			If x = 0 Then
				Argument = DlgText("Argument")
				DlgText "Argument",Argument & " " & """%1"""
			End If
		End If
		If DlgItem$ = "OKButton" Then
			If DlgText("CmdPath") = "" Then
 				MsgBox Msg04,vbOkOnly+vbInformation,Msg01
				CmdInputFunc = True ' 滅砦偌狟偌聽壽敕勤趕遺敦諳
			End If
		End If
		If DlgItem$ <> "OKButton" And DlgItem$ <> "CancelButton" Then
 			If DlgText("CmdPath") = "" And DlgText("Argument") = "" Then
 				DlgEnable "ClearButton",False
 			Else
 				DlgEnable "ClearButton",True
 			End If
 			CmdInputFunc = True ' 滅砦偌狟偌聽壽敕勤趕遺敦諳
		End If
	Case 3 ' 恅掛遺麼氪郪磁遺恅掛掩載蜊
 		If DlgText("CmdPath") = "" And DlgText("Argument") = "" Then
 			DlgEnable "ClearButton",False
 		Else
 			DlgEnable "ClearButton",True
 		End If
	End Select
End Function


' 恅璃揭燴渣昫秏洘
Sub ErrorMassage(MsgType As String)
	If OSLanguage = "0404" Then
		Msg01 = "錯誤"
		msg02 = "無法讀取檔案！" & vbCrLf & _
				"可能是非文字檔案或編碼不被支援。" & vbCrLf & _
				"請確認檔案類型或選取其它編碼 (預覽) 後再試。"
		Msg03 = "無法寫入檔案！" & vbCrLf & _
				"請檢查目標檔案是否可寫入或有寫入權限。"
	Else
		Msg01 = "渣昫"
		msg02 = "拸楊黍�﹡躁�ㄐ" & vbCrLf & _
				"褫夔岆準恅掛恅璃麼晤鎢祥掩盓厥﹝" & vbCrLf & _
				"③�溜玴躁�濬倰麼恁寁む坳晤鎢 (啎擬) 綴婬彸﹝"
		Msg03 = "拸楊迡�輷躁�ㄐ" & vbCrLf & _
				"③潰脤醴梓恅璃岆瘁褫迡麼衄迡�躽使煄�"
	End If
	If MsgType = "NotReadFile" Then MsgBox(msg02,vbOkOnly+vbInformation,Msg01)
	If MsgType = "NotWriteFile" Then MsgBox(Msg03,vbOkOnly+vbInformation,Msg01)
End Sub


' 潰脤恅璃晤鎢
' ----------------------------------------------------
' ANSI      拸跡宒隅砱
' EFBB BF   UTF-8
' FFFE      UTF-16LE/UCS-2, Little Endian with BOM
' FEFF      UTF-16BE/UCS-2, Big Endian with BOM
' XX00 XX00 UTF-16LE/UCS-2, Little Endian without BOM
' 00XX 00XX UTF-16BE/UCS-2, Big Endian without BOM
' FFFE 0000 UTF-32LE/UCS-4, Little Endian with BOM
' 0000 FEFF UTF-32BE/UCS-4, Big Endian with BOM
' XX00 0000 UTF-32LE/UCS-4, Little Endian without BOM
' 0000 00XX UTF-32BE/UCS-4, Big Endian without BOM
' 奻扴笢腔 XX 桶尨�扂獃挨虌鱣び硊�

Function CheckCode(FilePath As String) As String
	Dim objStream As Object,i As Integer,Code As String
	On Error Resume Next
	Set objStream = CreateObject("Adodb.Stream")
	If Not objStream Is Nothing Then
		objStream.Type = 1
		objStream.Mode = 3
		objStream.Open
		objStream.Position = 0
		objStream.LoadFromFile FilePath
		Bin = objStream.read(4)
		objStream.Close
		Bin1 = AscB(MidB(Bin,1,1))
		Bin2 = AscB(MidB(Bin,2,1))
		Bin3 = AscB(MidB(Bin,3,1))
		Bin4 = AscB(MidB(Bin,4,1))
		If Bin1 = &HEF And Bin2 = &HBB Then
			CheckCode = "utf-8EFBB"
		ElseIf Bin1 = &HFF And Bin2 = &HFE Then
			CheckCode = "unicodeFFFE"
		ElseIf Bin1 = &HFE And Bin2 = &HFF Then
			CheckCode = "unicodeFEFF"
		ElseIf Bin1 <> &H00 And Bin2 = &H00 And Bin3 <> &H00 And Bin4 = &H00 Then
			CheckCode = "utf-16LE"
		ElseIf Bin1 = &H00 And Bin2 <> &H00 And Bin3 = &H00 And Bin4 <> &H00 Then
			CheckCode = "utf-16BE"
		ElseIf Bin1 = &HFF And Bin2 = &HFE And Bin3 = &H00 And Bin4 = &H00 Then
			CheckCode = "unicode-32FFFE"
		ElseIf Bin1 = &H00 And Bin2 = &H00 And Bin3 = &HFE And Bin4 = &HFF Then
			CheckCode = "unicode-32FEFF"
		ElseIf Bin1 = &H00 And Bin2 = &H00 And Bin3 = &H00 And Bin4 <> &H00 Then
			CheckCode = "utf-32LE"
		ElseIf Bin1 <> &H00 And Bin2 = &H00 And Bin3 = &H00 And Bin4 = &H00 Then
			CheckCode = "utf-32BE"
		Else
			Code = "_autodetect_all"
			objStream.Type = 2
			objStream.Mode = 3
			objStream.Charset = Code
			objStream.Open
			objStream.LoadFromFile FilePath
			PreTextStr = objStream.ReadText
			objStream.Close
			For i = 38 To 2 Step -1
				If i <> 9 And i <> 13 Then
					TempArray = Split(CodeList(i),JoinStr)
					Code = TempArray(1)
					objStream.Type = 2
					objStream.Mode = 3
					objStream.Charset = Code
					objStream.Open
					objStream.LoadFromFile FilePath
					AppTextStr = objStream.ReadText
					objStream.Close
					If PreTextStr = AppTextStr Then
						CheckCode = Code
						Exit For
					End If
				End If
			Next i
			If CheckCode = "" Then
				For i = 41 To 39 Step -1
					If i <> 40 Then
						TempArray = Split(CodeList(i),JoinStr)
						Code = TempArray(1)
						objStream.Type = 2
						objStream.Mode = 3
						objStream.Charset = Code
						objStream.Open
						objStream.LoadFromFile FilePath
						AppTextStr = objStream.ReadText
						objStream.Close
						If PreTextStr = AppTextStr Then
							CheckCode = Code
							Exit For
						End If
					End If
				Next i
			End If
			If CheckCode = "" Then CheckCode = "ANSI"
		End If
		Set objStream = Nothing
	Else
		CheckCode = "ANSI"
	End If
End Function


' 黍�﹡躁�
Function ReadFile(FilePath As String,CharSet As String) As String
	Dim objStream As Object,Code As String
	Set objStream = CreateObject("Adodb.Stream")
	Code = CharSet
	If Not objStream Is Nothing Then
		On Error GoTo ErrorMsg
		If Code = "" Then Code = CheckCode(FilePath)
		If Code = "utf-8EFBB" Then Code = "utf-8"
		If Code <> "ANSI" Then
			objStream.Type = 2
			objStream.Mode = 3
			objStream.Charset = Code
			objStream.Open
			objStream.LoadFromFile FilePath
			ReadFile = objStream.ReadText
			objStream.Close
		End If
		On Error GoTo 0
	End If
	If objStream Is Nothing Or Code = "ANSI" Then
		On Error GoTo ErrorMsg
		Open FilePath For Input As #1
		While Not EOF(1)
			Line Input #1,l$
			If ReadFile <> "" Then ReadFile = ReadFile & vbCrLf & l$
			If ReadFile = "" Then ReadFile = l$
		Wend
		Close #1
		On Error GoTo 0
	End If
	If Not objStream Is Nothing Then Set objStream = Nothing
	Exit Function

	ErrorMsg:
	If Not objStream Is Nothing Then Set objStream = Nothing
	ErrorMassage("NotReadFile")
	ReadFile = ""
End Function


' 迡�輷躁�
Function WriteToFile(FilePath As String,textStr As String,CharSet As String) As Boolean
	Dim objStream As Object,Code As String
	On Error GoTo ErrorMsg
	Set objStream = CreateObject("Adodb.Stream")
	Code = CharSet
	If LCase(Code) = "_autodetect_all" Then Code = "ANSI"
	If Not objStream Is Nothing Then
		If Code = "" Then Code = CheckCode(FilePath)
		If Code = "utf-8EFBB" Then Code = "utf-8"
		If Code <> "ANSI" Then
			objStream.Type = 2
			objStream.Mode = 3
			objStream.CharSet = Code
			objStream.Open
			objStream.WriteText textStr
			objStream.SaveToFile FilePath,2
			objStream.Flush
			objStream.Close
			WriteToFile = True
		End If
	End If
	If objStream Is Nothing Or Code = "ANSI" Then
		Open FilePath For Output As #1
		Print #1,textStr;
		Close #1
		WriteToFile = True
	End If
	On Error GoTo 0
	If Not objStream Is Nothing Then Set objStream = Nothing
	Exit Function

	ErrorMsg:
	WriteToFile = False
	If Not objStream Is Nothing Then Set objStream = Nothing
	ErrorMassage("NotWriteFile")
End Function


'斐膘測鎢珜杅郪
Public Function CodePageList(MinNum As Integer,MaxNum As Integer) As Variant
	Dim CodePage() As String,i As Integer,j As Integer
	ReDim CodePage(MaxNum - MinNum)
	For i = MinNum To MaxNum
		j = i - MinNum
		If OSLanguage = "0404" Then
			If i = 0 Then CodePage(j) = "系統預設" & JoinStr & "ANSI"
			If i = 1 Then CodePage(j) = "自動選取" & JoinStr & "_autodetect_all"
			If i = 2 Then CodePage(j) = "簡體中文(GB2312)" & JoinStr & "gb2312"
			If i = 3 Then CodePage(j) = "簡體中文(HZ)" & JoinStr & "hz-gb-2312"
			If i = 4 Then CodePage(j) = "簡體中文(GB18030)" & JoinStr & "gb18030"
			If i = 5 Then CodePage(j) = "正體中文(Big5)" & JoinStr & "big5"
			If i = 6 Then CodePage(j) = "日文(EUC)" & JoinStr & "euc-jp"
			If i = 7 Then CodePage(j) = "日文(JIS)" & JoinStr & "iso-2022-jp"
			If i = 8 Then CodePage(j) = "日文(Shift-JIS)" & JoinStr & "shift_jis"
			If i = 9 Then CodePage(j) = "日文(自動選取)" & JoinStr & "_autodetect"
			If i = 10 Then CodePage(j) = "韓文" & JoinStr & "ks_c_5601-1987"
			If i = 11 Then CodePage(j) = "韓文(EUC)" & JoinStr & "euc-kr"
			If i = 12 Then CodePage(j) = "韓文(ISO)" & JoinStr & "iso-2022-kr"
			If i = 13 Then CodePage(j) = "韓文(自動選取)" & JoinStr & "_autodetect_kr"
			If i = 14 Then CodePage(j) = "泰文(Windows)" & JoinStr & "windows-874"
			If i = 15 Then CodePage(j) = "越南文(Windows)" & JoinStr & "windows-1258"
			If i = 16 Then CodePage(j) = "波羅的海文(ISO)" & JoinStr & "iso-8859-4"
			If i = 17 Then CodePage(j) = "波羅的海文(Windows)" & JoinStr & "windows-1257"
			If i = 18 Then CodePage(j) = "阿拉伯文(ASMO 708)" & JoinStr & "ASMO-708"
			If i = 19 Then CodePage(j) = "阿拉伯文(DOS)" & JoinStr & "DOS-720"
			If i = 20 Then CodePage(j) = "阿拉伯文(ISO)" & JoinStr & "iso-8859-6"
			If i = 21 Then CodePage(j) = "阿拉伯文(Windows)" & JoinStr & "windows-1256"
			If i = 22 Then CodePage(j) = "希伯來文(DOS)" & JoinStr & "DOS-862"
			If i = 23 Then CodePage(j) = "希伯來文(ISO-邏輯)" & JoinStr & "iso-8859-8-i"
			If i = 24 Then CodePage(j) = "希伯來文(ISO-視覺)" & JoinStr & "iso-8859-8"
			If i = 25 Then CodePage(j) = "希伯來文(Windows)" & JoinStr & "windows-1255"
			If i = 26 Then CodePage(j) = "土耳其文(Windows)" & JoinStr & "iso-8859-9"
			If i = 27 Then CodePage(j) = "希臘文(ISO)" & JoinStr & "iso-8859-7"
			If i = 28 Then CodePage(j) = "希臘文(Windows)" & JoinStr & "windows-1253"
			If i = 29 Then CodePage(j) = "西歐(Windows)" & JoinStr & "iso-8859-1"
			If i = 30 Then CodePage(j) = "西裡爾文(DOS)" & JoinStr & "cp866"
			If i = 31 Then CodePage(j) = "西裡爾文(ISO)" & JoinStr & "iso-8859-5"
			If i = 32 Then CodePage(j) = "西裡爾文(KOI8-R)" & JoinStr & "koi8-r"
			If i = 33 Then CodePage(j) = "西裡爾文(KOI8-U)" & JoinStr & "koi8-ru"
			If i = 34 Then CodePage(j) = "西裡爾文(Windows)" & JoinStr & "windows-1251"
			If i = 35 Then CodePage(j) = "中歐(DOS)" & JoinStr & "ibm852"
			If i = 36 Then CodePage(j) = "中歐(ISO)" & JoinStr & "iso-8859-2"
			If i = 37 Then CodePage(j) = "中歐(Windows)" & JoinStr & "windows-1250"
			If i = 38 Then CodePage(j) = "拉丁文 3 (ISO)" & JoinStr & "iso-8859-3"
			If i = 39 Then CodePage(j) = "Unicode (UTF-7)" & JoinStr & "utf-7"
			If i = 40 Then CodePage(j) = "Unicode (UTF-8 有 BOM)" & JoinStr & "utf-8EFBB"
			If i = 41 Then CodePage(j) = "Unicode (UTF-8 無 BOM)" & JoinStr & "utf-8"
			If i = 42 Then CodePage(j) = "Unicode (UTF-16LE 有 BOM)" & JoinStr & "unicodeFFFE"
			If i = 43 Then CodePage(j) = "Unicode (UTF-16BE 有 BOM)" & JoinStr & "unicodeFEFF"
			If i = 44 Then CodePage(j) = "Unicode (UTF-16LE 無 BOM)" & JoinStr & "utf-16LE"
			If i = 45 Then CodePage(j) = "Unicode (UTF-16BE 無 BOM)" & JoinStr & "utf-16BE"
			If i = 46 Then CodePage(j) = "Unicode (UTF-32LE 有 BOM)" & JoinStr & "unicode-32FFFE"
			If i = 47 Then CodePage(j) = "Unicode (UTF-32BE 有 BOM)" & JoinStr & "unicode-32FEFF"
			If i = 48 Then CodePage(j) = "Unicode (UTF-32LE 無 BOM)" & JoinStr & "utf-32LE"
			If i = 49 Then CodePage(j) = "Unicode (UTF-32BE 無 BOM)" & JoinStr & "utf-32BE"
		Else
			If i = 0 Then CodePage(j) = "炵苀蘇��" & JoinStr & "ANSI"
			If i = 1 Then CodePage(j) = "赻雄恁寁" & JoinStr & "_autodetect_all"
			If i = 2 Then CodePage(j) = "潠极笢恅(GB2312)" & JoinStr & "gb2312"
			If i = 3 Then CodePage(j) = "潠极笢恅(HZ)" & JoinStr & "hz-gb-2312"
			If i = 4 Then CodePage(j) = "潠极笢恅(GB18030)" & JoinStr & "gb18030"
			If i = 5 Then CodePage(j) = "楛极笢恅(Big5)" & JoinStr & "big5"
			If i = 6 Then CodePage(j) = "�梉�(EUC)" & JoinStr & "euc-jp"
			If i = 7 Then CodePage(j) = "�梉�(JIS)" & JoinStr & "iso-2022-jp"
			If i = 8 Then CodePage(j) = "�梉�(Shift-JIS)" & JoinStr & "shift_jis"
			If i = 9 Then CodePage(j) = "�梉�(赻雄恁寁)" & JoinStr & "_autodetect"
			If i = 10 Then CodePage(j) = "澈恅" & JoinStr & "ks_c_5601-1987"
			If i = 11 Then CodePage(j) = "澈恅(EUC)" & JoinStr & "euc-kr"
			If i = 12 Then CodePage(j) = "澈恅(ISO)" & JoinStr & "iso-2022-kr"
			If i = 13 Then CodePage(j) = "澈恅(赻雄恁寁)" & JoinStr & "_autodetect_kr"
			If i = 14 Then CodePage(j) = "怍恅(Windows)" & JoinStr & "windows-874"
			If i = 15 Then CodePage(j) = "埣鰍恅(Windows)" & JoinStr & "windows-1258"
			If i = 16 Then CodePage(j) = "疏蹕腔漆恅(ISO)" & JoinStr & "iso-8859-4"
			If i = 17 Then CodePage(j) = "疏蹕腔漆恅(Windows)" & JoinStr & "windows-1257"
			If i = 18 Then CodePage(j) = "陝嶺皎恅(ASMO 708)" & JoinStr & "ASMO-708"
			If i = 19 Then CodePage(j) = "陝嶺皎恅(DOS)" & JoinStr & "DOS-720"
			If i = 20 Then CodePage(j) = "陝嶺皎恅(ISO)" & JoinStr & "iso-8859-6"
			If i = 21 Then CodePage(j) = "陝嶺皎恅(Windows)" & JoinStr & "windows-1256"
			If i = 22 Then CodePage(j) = "洷皎懂恅(DOS)" & JoinStr & "DOS-862"
			If i = 23 Then CodePage(j) = "洷皎懂恅(ISO-軀憮)" & JoinStr & "iso-8859-8-i"
			If i = 24 Then CodePage(j) = "洷皎懂恅(ISO-弝橇)" & JoinStr & "iso-8859-8"
			If i = 25 Then CodePage(j) = "洷皎懂恅(Windows)" & JoinStr & "windows-1255"
			If i = 26 Then CodePage(j) = "芩嫉む恅(Windows)" & JoinStr & "iso-8859-9"
			If i = 27 Then CodePage(j) = "洷幫恅(ISO)" & JoinStr & "iso-8859-7"
			If i = 28 Then CodePage(j) = "洷幫恅(Windows)" & JoinStr & "windows-1253"
			If i = 29 Then CodePage(j) = "昹韁(Windows)" & JoinStr & "iso-8859-1"
			If i = 30 Then CodePage(j) = "昹爵嫌恅(DOS)" & JoinStr & "cp866"
			If i = 31 Then CodePage(j) = "昹爵嫌恅(ISO)" & JoinStr & "iso-8859-5"
			If i = 32 Then CodePage(j) = "昹爵嫌恅(KOI8-R)" & JoinStr & "koi8-r"
			If i = 33 Then CodePage(j) = "昹爵嫌恅(KOI8-U)" & JoinStr & "koi8-ru"
			If i = 34 Then CodePage(j) = "昹爵嫌恅(Windows)" & JoinStr & "windows-1251"
			If i = 35 Then CodePage(j) = "笢韁(DOS)" & JoinStr & "ibm852"
			If i = 36 Then CodePage(j) = "笢韁(ISO)" & JoinStr & "iso-8859-2"
			If i = 37 Then CodePage(j) = "笢韁(Windows)" & JoinStr & "windows-1250"
			If i = 38 Then CodePage(j) = "嶺間恅 3 (ISO)" & JoinStr & "iso-8859-3"
			If i = 39 Then CodePage(j) = "Unicode (UTF-7)" & JoinStr & "utf-7"
			If i = 40 Then CodePage(j) = "Unicode (UTF-8 衄 BOM)" & JoinStr & "utf-8EFBB"
			If i = 41 Then CodePage(j) = "Unicode (UTF-8 拸 BOM)" & JoinStr & "utf-8"
			If i = 42 Then CodePage(j) = "Unicode (UTF-16LE 衄 BOM)" & JoinStr & "unicodeFFFE"
			If i = 43 Then CodePage(j) = "Unicode (UTF-16BE 衄 BOM)" & JoinStr & "unicodeFEFF"
			If i = 44 Then CodePage(j) = "Unicode (UTF-16LE 拸 BOM)" & JoinStr & "utf-16LE"
			If i = 45 Then CodePage(j) = "Unicode (UTF-16BE 拸 BOM)" & JoinStr & "utf-16BE"
			If i = 46 Then CodePage(j) = "Unicode (UTF-32LE 衄 BOM)" & JoinStr & "unicode-32FFFE"
			If i = 47 Then CodePage(j) = "Unicode (UTF-32BE 衄 BOM)" & JoinStr & "unicode-32FEFF"
			If i = 48 Then CodePage(j) = "Unicode (UTF-32LE 拸 BOM)" & JoinStr & "utf-32LE"
			If i = 49 Then CodePage(j) = "Unicode (UTF-32BE 拸 BOM)" & JoinStr & "utf-32BE"
		End If
	Next i
	CodePageList = CodePage
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

	BingLangCode = "||||ar||||||||||bg||zh-CHS|zh-CHT|||cs|da|nl|en||||fi|fr||||de|el||||he||" & _
				"hu|||||||it|ja|||||||ko||||lv|lt||||||||||no|no|no|||pl|pt|||ro|ru||||||||sk|" & _
				"sl|es||sv||||||th||tr||||||||"

	GoogleLangCode = "auto|af|sq||ar||||||be||||bg|ca|zh-CN|zh-TW||hr|cs|da|nl|en|et||fa|fi|" & _
				"fr||gl||de|el||||iw|hi|hu|Is|id||ga|||it|ja|||||||ko||||lv|lt||mk|ms||mt|||||" & _
				"no|no|no|||pl|pt|||ro|ru|||sr|||||sk|sl|es|sw|sv||||||th||tr|||uk|||vi|cy|"

	YahooLangCode = "||||||||||||||||zh|zt|||||nl|en|||||fr||||de|el|||||||||||||it|ja|||||||" & _
				"ko|||||||||||||||||||||pt||||ru||||||||||es||||||||||||||||||"

	en2zhCheck = "||||||||||||||||zh-CN|zh-TW||||||||||||||||||||||||||||||ja|||||||ko|||||||" & _
				"||||||||||||||||||||||||||||||||||||||||||||||"

	zh2enCheck = "|af|sq|am|ar|hy|As|az|ba|eu|be|BN|bs|br|bg|ca|||co|hr|cs|da|nl|" & _
				"en|et|fo|fa|fi|fr|fy|gl|ka|de|el|kl|gu|ha|he|hi|hu|Is|id|iu|ga|xh|zu|it|" & _
				"|kn|KS|kk|km|rw|kok||kz|ky|lo|lv|lt|lb|mk|ms|ML|mt|mi|mr|mn|ne|no|nb|" & _
				"nn|Or|ps|pl|pt|pa|qu|ro|ru|se|sa|sr|st|tn|SD|si|sk|sl|es|sw|sv|sy|tg|ta|tt|" & _
				"te|th|bo|tr|tk|ug|uk|ur|uz|vi|cy|wo"

	PslLangCodeList = Split(PslLangCode,LngJoinStr)
	BingLangCodeList = Split(BingLangCode,LngJoinStr)
	GoogleLangCodeList = Split(GoogleLangCode,LngJoinStr)
	YahooLangCodeList = Split(YahooLangCode,LngJoinStr)
	en2zhCheckList = Split(en2zhCheck,LngJoinStr)
	zh2enCheckList = Split(zh2enCheck,LngJoinStr)

	For i = MinNum To MaxNum
		j = i - MinNum
		If DataName = DefaultEngineList(0) Then Code = BingLangCodeList(i)
		If DataName = DefaultEngineList(1) Then Code = GoogleLangCodeList(i)
		If DataName = DefaultEngineList(2) Then Code = YahooLangCodeList(i)
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
			"  - 翻譯前要替換的字元" & vbCrLf & _
			"    定義在翻譯前要被替換的字元以及替換後的字元。" & vbCrLf & vbCrLf & _
			"  - 翻譯後要替換的字元" & vbCrLf & _
			"    定義在翻譯後要被替換的字元以及替換後的字元。" & vbCrLf & vbCrLf & _
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
			"  - 楹祒ヶ猁杸遙腔趼睫" & vbCrLf & _
			"    隅砱婓楹祒ヶ猁掩杸遙腔趼睫眕摯杸遙綴腔趼睫﹝" & vbCrLf & vbCrLf & _
			"  - 楹祒綴猁杸遙腔趼睫" & vbCrLf & _
			"    隅砱婓楹祒綴猁掩杸遙腔趼睫眕摯杸遙綴腔趼睫﹝" & vbCrLf & vbCrLf & _
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


'竘э堆翑
Sub EngineHelp(HelpTip As String)
	If OSLanguage = "0404" Then
	AboutTitle = "關於"
	HelpTitle = "說明"
	HelpTipTitle = "Passolo 連線翻譯巨集"
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
			"開 發 者：漢化新世紀成員 wanfu (2010)" & vbCrLf & vbCrLf & vbCrLf
	Ement = "☆執行環境☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 支援巨集處理的 Passolo 6.0 及以上版本，必需" & vbCrLf & _
			"- Windows Script Host (WSH) 物件 (VBS)，必需" & vbCrLf & _
			"- Adodb.Stream 物件，支援 Utf-8、Unicode 必需" & vbCrLf & _
			"- Microsoft.XMLHTTP 物件，必需" & vbCrLf & _
			"- Microsoft.XMLDOM 物件，剖析 responseXML 返回格式所需" & vbCrLf & vbCrLf & vbCrLf
	Dec = "☆軟體簡介☆" & vbCrLf & _
			"============" & vbCrLf & _
			"Passolo 連線翻譯巨集是一個用於 Passolo 翻譯清單字串的連線翻譯巨集程式。它具有以下功能：" & vbCrLf & _
			"- 利用連線翻譯引擎自動翻譯 Passolo 翻譯清單中的字串" & vbCrLf & _
			"- 整合了一些著名的連線翻譯引擎，並可自訂其他連線翻譯引擎" & vbCrLf & _
			"- 可選取字串類型、跳過字串以及對翻譯前後的字串進行處理" & vbCrLf & _
			"- 整合便捷鍵、終止符、加速器檢查巨集、可在翻譯後檢查並糾正翻譯中的錯誤" & vbCrLf & _
			"- 內置可自訂的自動更新功能" & vbCrLf & vbCrLf & _
			"本程式包含下列檔案：" & vbCrLf & _
			"- PSLWebTrans.bas (Passolo 連線翻譯巨集檔案 - 對話方塊方式執行)" & vbCrLf & _
			"- PSLWebTrans_Silent.bas (Passolo 連線翻譯巨集檔案 - 靜默方式執行)" & vbCrLf & _
			"- PSLWebTrans.txt (簡體中文說明檔案)" & vbCrLf & vbCrLf & vbCrLf
	Setup = "☆安裝方法☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 將解壓後的檔案複製到 Passolo 系統資料夾中定義的 Macros 資料夾中" & vbCrLf & _
			"- 在 Passolo 的工具 -> 自訂工具選單中新增該巨集檔案並定義該選單名稱，" & vbCrLf & _
			"  此後就可以點擊該選單直接呼叫" & vbCrLf & _
			"- 靜默方式執行的連線翻譯巨集僅作用於選擇字串，跳過字串和字串處理按對話方塊方式執行" & vbCrLf & _
 			"  的連線翻譯巨集設定執行。要定義該巨集的執行參數，需使用對話方塊方式設定" & vbCrLf & vbCrLf & vbCrLf
	MainUse = "☆翻譯引擎☆" & vbCrLf & _
			"============" & vbCrLf & _
			"程式已經提供一些著名的連線翻譯引擎。您也可以按下面的 [設定] 按鈕自訂設定。" & vbCrLf & _
			"新增自訂設定後，您可以在翻譯引擎清單中選取想使用的翻譯引擎。" & vbCrLf & _
			"有關自訂翻譯引擎，請開啟設定對話方塊，點擊 [說明] 按鈕，參閱說明中的說明。" & vbCrLf & vbCrLf & _
			"☆翻譯原文☆" & vbCrLf & _
			"============" & vbCrLf & _
			"程式自動列出目前翻譯清單所屬檔案中除目前翻譯清單的目標語言外的所有目標語言。" & vbCrLf & _
			"通過選取不同的語言，可以選取目前翻譯清單中的來源字串，或者同一檔案中的其他翻譯清單中" & vbCrLf & _
			"的翻譯字串。" & vbCrLf & vbCrLf & _
			"- 選取與目前翻譯清單目標語言相同的語言時，將使用目前翻譯清單中的來源字串作為翻譯原文。" & vbCrLf & _
			"- 選取與目前翻譯清單目標語言不同的語言時，將使用選擇的其他翻譯清單中的翻譯字串作為翻譯" & vbCrLf & _
			"  原文。以方便將已有的翻譯轉換成其他語言。比如將簡體中文翻譯成正體中文。" & vbCrLf & vbCrLf & _
			"☆翻譯字串☆" & vbCrLf & _
			"============" & vbCrLf & _
			"提供了全部、選單、對話方塊、字串表、加速器、版本、其他、僅選擇等選項。" & vbCrLf & vbCrLf & _
			"- 如果選取全部，則其他單項將被自動取消選取。" & vbCrLf & _
			"- 如果選取單項，則全部選項將被自動取消選取。" & vbCrLf & _
			"- 單項可以多選。其中選取僅選擇時，其他均被自動取消選取。" & vbCrLf & vbCrLf & _
			"☆跳過字串☆" & vbCrLf & _
			"============" & vbCrLf & _
			"提供了供複審、已驗證、未翻譯、全為數字和符號、全為大寫英文、全為小寫英文等選項。" & vbCrLf & vbCrLf & _
			"- 所有項目均可以多選。" & vbCrLf & _
			"- 選取全為大寫英文、全為小寫英文選項時將自動選擇全為數字和符號選項，並且全為數字" & vbCrLf & _
			"  和符號選項將自動變為不可用狀態。" & vbCrLf & vbCrLf & _
			"☆字串處理☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 設定" & vbCrLf & _
			"  程式提供了和便捷鍵、終止符、加速器檢查巨集相同的預設設定，該設定可以適用於大多數情況。" & vbCrLf & _
			"  您也可以按下面的 [設定] 按鈕自訂設定。" & vbCrLf & _
			"  新增自訂設定後，您可以在設定清單中選取想使用的設定。" & vbCrLf & _
			"  有關自訂設定，請開啟設定對話方塊，點擊 [說明] 按鈕，參閱說明中的說明。" & vbCrLf & vbCrLf & _
			"- 自動選取設定" & vbCrLf & _
			"  選擇該選項時將根據設定中的可用語言清單自動選取與翻譯清單目標語言符合的設定。" & vbCrLf & _
			"  ◎注意：該選項僅對連線翻譯巨集有效。" & vbCrLf & _
			"  　　　　要將設定與目前翻譯清單的目標語言符合，按 [設定] 按鈕，在字串處理的適用語言中" & vbCrLf & _
			"  　　　　新增對應的語言。" & vbCrLf & _
			"  　　　　在靜默方式情況下，程式將按「自動 - 自選 - 預設」順序選取對應的設定。" & vbCrLf & vbCrLf & _
			"- 去除便捷鍵" & vbCrLf & _
			"  在翻譯前將字串中的便捷鍵刪除，以便翻譯引擎可以正確翻譯。" & vbCrLf & vbCrLf & _
			"- 去除加速器" & vbCrLf & _
			"  在翻譯前將字串中的加速器刪除，以便翻譯引擎可以正確翻譯。" & vbCrLf & vbCrLf & _
			"- 替換特定字元並在翻譯後還原" & vbCrLf & _
			"  在翻譯前使用選擇的設定，替換特定字元並在翻譯後還原。" & vbCrLf & vbCrLf & _
			"  要新增這些特定字元，請在設定對話方塊的字元替換的翻譯前要替換的字元中定義。" & vbCrLf & _
			"  ◎注意：由於巨集引擎的問題，有些字元是不可被輸入或者無法被辨識。" & vbCrLf & _
			"  　　　　如果翻譯引擎翻譯了這些替換後的字元，被替換的字元無法被還原。" & vbCrLf & vbCrLf & _
			"- 分行翻譯" & vbCrLf & _
			"  在翻譯前將多行的字串分割為單行字串進行翻譯，以便提高翻譯引擎翻譯的正確度。" & vbCrLf & vbCrLf & _
			"- 糾正便捷鍵、終止符、加速器" & vbCrLf & _
			"  在翻譯後使用選擇的設定，檢查並糾正翻譯中的錯誤。" & vbCrLf & _
			"  糾正方式和便捷鍵、終止符、加速器檢查巨集完全相同。" & vbCrLf & vbCrLf & _
			"- 替換特定字元" & vbCrLf & _
			"  在翻譯後使用選擇的設定，替換翻譯中特定的字元。" & vbCrLf & vbCrLf & _
			"  要新增這些特定字元，請在設定對話方塊的字元替換的翻譯後要替換的字元中定義。" & vbCrLf & _
			"  ◎注意：由於巨集引擎的問題，有些字元是不可被輸入或者無法被辨識。" & vbCrLf & _
			"  　　　　如果翻譯引擎翻譯了這些替換前的字元，將無法被替換。" & vbCrLf & vbCrLf & _
			"☆其他選項☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 繼續時自動儲存所有選取" & vbCrLf & _
			"  選擇該項時，將在按 [繼續] 按鈕時自動儲存所有選取，下次執行時將讀入儲存的選取。" & vbCrLf & vbCrLf & _
			"- 顯示輸出訊息" & vbCrLf & _
			"  選擇該項時，將在 Passolo 訊息輸出視窗中輸出翻譯、糾正和替換訊息。" & vbCrLf & _
			"  在不選擇該項情況下，將顯著提高程式的執行速度。" & vbCrLf & vbCrLf & _
			"- 新增翻譯註解" & vbCrLf & _
			"  選擇該項時，將在翻譯清單的字串註解中增加選擇翻譯引擎自動翻譯這樣的註解。" & vbCrLf & _
			"  利用該註解可以區分翻譯字串的翻譯出處。" & vbCrLf& vbCrLf & _
			"☆其他功能☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 關於" & vbCrLf & _
			"  點擊該按鈕，將關於對話方塊。並顯示程式介紹、執行環境、開發商及版權等訊息。" & vbCrLf & vbCrLf & _
			"- 說明" & vbCrLf & _
			"  點擊該按鈕，將彈出目前視窗的說明訊息。" & vbCrLf & vbCrLf & _
			"- 設定" & vbCrLf & _
			"  點擊該按鈕，將彈出設定對話方塊。可以在設定對話方塊中設定各種參數。" & vbCrLf & vbCrLf & _
			"- 確定" & vbCrLf & _
			"  點擊該按鈕，將關閉主對話方塊，並按選擇的選項進行字串翻譯。" & vbCrLf & _
			"  ◎注意：翻譯後的字串翻譯狀態將變更為待複審狀態，以示區別。" & vbCrLf& vbCrLf & _
			"- 取消" & vbCrLf & _
			"  點擊該按鈕，將結束程式。" & vbCrLf& vbCrLf & vbCrLf
	SetUse ="☆設定清單☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 選取設定" & vbCrLf & _
			"  要選取設定項，點擊設定清單。" & vbCrLf & vbCrLf & _
			"- 新增設定" & vbCrLf & _
			"  要新增設定項，點擊 [新增] 按鈕，在彈出的對話方塊中輸入名稱。" & vbCrLf & vbCrLf & _
			"- 變更設定" & vbCrLf & _
			"  要變更設定名稱，選取設定清單中要改名的設定，然後點擊 [變更] 按鈕。" & vbCrLf & vbCrLf & _
			"- 刪除設定" & vbCrLf & _
			"  要刪除設定項，選取設定清單中要刪除的設定，然後點擊 [刪除] 按鈕。" & vbCrLf & vbCrLf & _
			"新增設定後，將在清單中顯示新的設定，設定內容將顯示空值。" & vbCrLf & _
			"變更設定後，將在清單中顯示改名後的設定，設定內容中的設定值不變。" & vbCrLf & _
			"刪除設定後，將在清單中顯示上一個設定，設定內容將顯示上一個設定值。" & vbCrLf & vbCrLf & _
			"☆儲存類型☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 檔案" & vbCrLf & _
			"  設定將以檔案形式儲存在巨集所在資料夾下的 Data 資料夾中。" & vbCrLf & vbCrLf & _
			"- 註冊表" & vbCrLf & _
			"  設定將被儲存註冊表中的 HKCU\Software\VB and VBA Program Settings\WebTranslate 項下。" & vbCrLf & vbCrLf & _
			"- 匯入設定" & vbCrLf & _
			"  允許從其他設定檔案中匯入設定。匯入舊設定時將被自動升級，現有設定清單中已有的設定將" & vbCrLf & _
			"  被變更，沒有的設定將被新增。" & vbCrLf & vbCrLf & _
			"- 匯出設定" & vbCrLf & _
			"  允許匯出所有設定到文字檔案，以便可以交換或轉移設定。" & vbCrLf & vbCrLf & _
			"◎注意：切換儲存類型時，將自動刪除原有位置中的設定內容。" & vbCrLf & vbCrLf & _
			"☆設定內容☆" & vbCrLf & _
			"============" & vbCrLf & _
			"<引擎參數>" & vbCrLf & _
			"  - 使用物件" & vbCrLf & _
			"    該物件預設為「Microsoft.XMLHTTP」，不可變更。" & vbCrLf & _
			"    有關該物件的一些使用方法，請參閱相關文件。" & vbCrLf & vbCrLf & _
			"  - 引擎註冊 ID" & vbCrLf & _
			"    有些翻譯引擎需要註冊才能使用。相關問題請與翻譯引擎提供商聯繫。" & vbCrLf & _
			"    程式已經提供了 Microsoft 翻譯引擎的註冊 ID。但不保證可以永久使用。" & vbCrLf & vbCrLf & _
			"  - 引擎網址" & vbCrLf & _
			"    翻譯引擎的存取網址。" & vbCrLf & _
			"    ◎注意：許多連線翻譯提供商的網址並不是真正的翻譯引擎存取網址。需要分析網頁代碼或" & vbCrLf & _
			"    　　　　詢問真正的翻譯引擎提供商才能取得。" & vbCrLf & vbCrLf & _
			"  - 傳送內容範本" & vbCrLf & _
			"    該欄位實際上是 Microsoft.XMLHTTP 的 Open 方法中的 bstrUrl 參數。" & vbCrLf & _
			"    ◎注意：大括號 {} 中的字元為系統欄位，不可變更，否則系統無法辨識。" & vbCrLf & _
			"    　　　　要新增這些系統欄位，請點擊右邊的「>」按鈕，選取需要的系統欄位。" & vbCrLf & vbCrLf & _
			"  - 資料傳送方式" & vbCrLf & _
			"    該欄位實際上是 Microsoft.XMLHTTP 物件的 Open 方法中的 bstrMethod 參數。" & vbCrLf & _
			"    即 GET 或 POST。用 POST 方式傳送資料,可以大到 4MB，也可以為 GET，只能 256KB。" & vbCrLf & vbCrLf & _
			"  - 同步方式" & vbCrLf & _
			"    該欄位實際上是 Microsoft.XMLHTTP 物件的 Open 方法中的 varAsync 參數。可省略。" & vbCrLf & _
			"    預設為 True，即同步執行，但只能在 DOM 中實施同步執行。一般將其設定為 False，即異步執行。" & vbCrLf & vbCrLf & _
			"  - 使用者名" & vbCrLf & _
			"    該欄位實際上是 Microsoft.XMLHTTP 物件的 Open 方法中的 bstrUser 參數。可省略。" & vbCrLf & vbCrLf & _
			"  - 密碼" & vbCrLf & _
			"    該欄位實際上是 Microsoft.XMLHTTP 物件的 Open 方法中的 bstrPassword 參數。可省略。" & vbCrLf & vbCrLf & _
			"  - 指令集" & vbCrLf & _
			"    該欄位實際上是 Microsoft.XMLHTTP 物件的 Send 方法中的 varBody 參數。可省略。" & vbCrLf & _
			"    它可以是 XML 格式資料，也可以是字串、資料流，或者一個無符號整數陣列。讓指令通過" & vbCrLf & _
			"    Open 方法中的 URL 參數代入。" & vbCrLf & vbCrLf & _
			"    傳送資料的方式分為同步和異步兩種：" & vbCrLf & _
			"    (1) 異步方式：資料包一旦傳送完畢，就結束 Send 進程，客戶機執行其他的操作。" & vbCrLf & _
			"    (2) 同步方式：客戶機要等到伺服器返回確認訊息後才結束 Send 進程。" & vbCrLf & vbCrLf & _
			"  - HTTP 頭和值" & vbCrLf & _
			"    該欄位實際上是 Microsoft.XMLHTTP 物件中的 setRequestHeader 方法，GET 方式下無效。" & vbCrLf & _
			"    (1) 該欄位支援多個 setRequestHeader 項目，每個項目用分行分隔。" & vbCrLf & _
			"    (2) 每個項目的頭名稱 (bstrHeader) 和值 (bstrValue) 用半形逗號分隔。" & vbCrLf & _
			"    (3) 需要傳送 ContentT-Length 項目時，其值可以使用系統欄位，系統將自動計算其長度。" & vbCrLf & vbCrLf & _
			"  - 返回結果格式參數" & vbCrLf & _
			"    該欄位實際上是 Microsoft.XMLHTTP 物件中的一個返回結果格式屬性。有以下幾種：" & vbCrLf & _
			"    responseBody	Variant	型	結果返回為無符號整數陣列" & vbCrLf & _
			"    responseStream	Variant	型	結果返回為 ADO Stream 物件" & vbCrLf & _
			"    responseText	String	型	結果返回為字串" & vbCrLf & _
			"    responseXML	Object	型	結果返回為 XML 格式資料" & vbCrLf & vbCrLf & _
			"    ◎注意：當選取 responseXML 時，「翻譯開始自」和「翻譯結束到」會分別顯示為「按 ID 搜索」" & vbCrLf & _
			"    　　　　和「按標籤名搜索」。" & vbCrLf & vbCrLf & _
			"  - 翻譯開始自" & vbCrLf & _
			"    該欄位用於辨識返回結果中的翻譯字串前的字串。" & vbCrLf & _
			"    可點擊右邊的「...」按鈕檢視從伺服器返回的所有文字中應該輸入的字元。" & vbCrLf & _
			"    ◎注意：該欄位支援用「|」分隔的多個或項項目。" & vbCrLf & vbCrLf & _
			"  - 翻譯結束到" & vbCrLf & _
			"    該欄位用於辨識返回結果中的翻譯字串後的字串。" & vbCrLf & _
			"    可點擊右邊的「...」按鈕檢視從伺服器返回的所有文字中應該輸入的字元。" & vbCrLf & _
			"    ◎注意：該欄位支援用「|」分隔的多個或項項目。" & vbCrLf & vbCrLf & _
			"  - 按 ID 搜索" & vbCrLf & _
			"    該欄位使用 XML DOM 的 getElementById 方法搜尋 XML 檔案中具有指定 ID 的元素中的文字值。" & vbCrLf & _
			"    可點擊右邊的「...」按鈕檢視從伺服器返回的所有文字中應該輸入的 ID 號。" & vbCrLf & _
			"    ◎注意：該欄位支援用「|」分隔的多個或項項目。" & vbCrLf & vbCrLf & _
			"  - 按標籤名搜索" & vbCrLf & _
			"    該欄位使用 XML DOM 的 getElementsByTagName 方法搜尋 XML 檔案中所有具有指定標籤名稱的" & vbCrLf & _
			"    元素的文字值。" & vbCrLf & _
			"    可點擊右邊的「...」按鈕檢視從伺服器返回的所有文字中應該輸入的標籤名稱。" & vbCrLf & _
			"    ◎注意：該欄位支援用「|」分隔的多個或項項目。" & vbCrLf & vbCrLf & _
			"<語言配對>" & vbCrLf & _
			"  - 語言名稱" & vbCrLf & _
			"    預設的語言名稱清單是將 Passolo 的語言清單進行了精簡，去掉了國家/地區的語言項。" & vbCrLf & vbCrLf & _
			"  - Passolo 代碼" & vbCrLf & _
			"    預設的語言代碼取自 Passolo 語言清單中的 ISO 936-1 語言代碼。其中：" & vbCrLf & _
			"    簡體中文和正體中文的語言代碼取自 Passolo 語言清單中的國家/地區語言代碼。" & vbCrLf & vbCrLf & _
			"  - 翻譯引擎代碼" & vbCrLf & _
			"    翻譯引擎所規定的語言代碼。需要分析網頁代碼或詢問翻譯引擎提供商才能取得。" & vbCrLf & vbCrLf & _
			"  - 增加語言" & vbCrLf & _
			"    要增加語言配對清單項，點擊 [新增] 按鈕，將彈出新增對話方塊。" & vbCrLf & vbCrLf & _
			"  - 刪除語言" & vbCrLf & _
			"    要刪除語言配對清單項，選取清單中要刪除的語言，然後點擊 [刪除] 按鈕。" & vbCrLf & vbCrLf & _
			"  - 編輯語言" & vbCrLf & _
			"    要編輯語言配對清單項，選取清單中要編輯的語言，然後點擊 [編輯] 按鈕。" & vbCrLf & vbCrLf & _
			"  - 外部編輯語言" & vbCrLf & _
			"    呼叫內置或外部編輯程式編輯整個語言配對清單。" & vbCrLf & vbCrLf & _
			"  - 置空語言" & vbCrLf & _
			"    要將翻譯引擎代碼置空，選取清單中要置空的語言，然後點擊 [置空] 按鈕。" & vbCrLf & _
			"    [置空] 按鈕會根據翻譯引擎代碼是否為空 (Null) 而顯示不同的可用狀態。" & vbCrLf & vbCrLf & _
			"  - 重設語言" & vbCrLf & _
			"    要將翻譯引擎代碼重設為原始值，選取清單中要重設的語言，然後點擊 [重設] 按鈕。" & vbCrLf & _
			"    [重設] 按鈕只有在翻譯引擎代碼被變更後才能轉為可用狀態。" & vbCrLf & vbCrLf & _
			"  - 顯示非空項" & vbCrLf & _
			"    要僅顯示翻譯引擎代碼為非空的語言項，點擊 [顯示非空項] 按鈕。點擊後該按鈕" & vbCrLf & _
			"    將轉為不可用狀態，[全部顯示] 和 [顯示空項] 按鈕將轉為可用狀態。" & vbCrLf & vbCrLf & _
			"  - 顯示空項" & vbCrLf & _
			"    要僅顯示翻譯引擎代碼為空 (Null) 的語言項，點擊 [顯示空項] 按鈕。點擊後" & vbCrLf & _
			"    該按鈕將轉為不可用狀態，[全部顯示] 和 [顯示非空項] 按鈕將轉為可用狀態。" & vbCrLf & vbCrLf & _
			"  - 全部顯示" & vbCrLf & _
			"    要顯示全部語言項，點擊 [全部顯示] 按鈕。點擊後該按鈕將轉為不可用狀態，" & vbCrLf & _
			"    [顯示非空項] 和 [顯示非空項] 按鈕將轉為可用狀態。" & vbCrLf & vbCrLf & _
			"  新增語言後，將在清單中顯示並選擇新的語言及代碼對。" & vbCrLf & _
			"  刪除語言後，將在清單中選擇所刪除語言的前一項語言。" & vbCrLf & _
			"  編輯語言後，將在清單中顯示並保持選擇編輯後的語言及代碼對。" & vbCrLf & _
			"  置空語言後，將在清單中顯示並保持選擇置空後的語言及代碼對。" & vbCrLf & _
			"  重設語言後，將在清單中顯示並保持選擇重設後的語言及代碼對。" & vbCrLf & vbCrLf & _
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
			"- 測試" & vbCrLf & _
			"  點擊該按鈕，將彈出測試對話方塊，以便檢查設定的正確性。" & vbCrLf & vbCrLf & _
			"- 清空" & vbCrLf & _
			"  點擊該按鈕，將清空現有設定的全部值，以方便重新輸入設定值。" & vbCrLf & vbCrLf & _
			"- 確定" & vbCrLf & _
			"  點擊該按鈕，將儲存設定視窗中的任何變更，結束設定視窗並返回主視窗。" & vbCrLf & _
			"  程式將使用變更後的設定值。" & vbCrLf & vbCrLf & _
			"- 取消" & vbCrLf & _
			"  點擊該按鈕，不儲存設定視窗中的任何變更，結束設定視窗並返回主視窗。" & vbCrLf & _
			"  程式將使用原來的設定值。" & vbCrLf & vbCrLf & vbCrLf
	TestUse ="☆翻譯引擎☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 要測試的翻譯引擎名稱。要選取翻譯引擎，點擊翻譯引擎清單。" & vbCrLf & vbCrLf & _
			"☆目標語言☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 翻譯的目標語言。該清單僅顯示語言配對清單中翻譯引擎語言代碼不為空的語言。" & vbCrLf & vbCrLf & _
			"☆翻譯清單☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 翻譯原文的翻譯清單源。該清單包含了專案中的所有翻譯清單。" & vbCrLf & vbCrLf & _
			"☆讀入行數☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 指定從選擇翻譯清單中讀入的字串數。" & vbCrLf & vbCrLf & _
			"☆翻譯內容☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 要翻譯的文字。選取翻譯清單、來源字串和翻譯字串選項時自動從選擇的翻譯清單中讀入字串。" & vbCrLf & _
			"- 被讀入的字串自動用空格連接。" & vbCrLf & vbCrLf & _
			"☆來源字串☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 選取該選項時，自動從選擇的翻譯清單中讀入來源字串。" & vbCrLf & vbCrLf & _
			"☆翻譯字串☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 選取該選項時，自動從選擇的翻譯清單中讀入翻譯字串。" & vbCrLf & vbCrLf & _
			"☆翻譯結果☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 從翻譯引擎返回的結果。根據選擇的來源字串和翻譯字串選項，分別顯示翻譯文字或從翻譯引擎" & vbCrLf & _
			"  返回的全部文字。" & vbCrLf & vbCrLf & _
			"☆翻譯文字☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 從翻譯引擎返回的翻譯。該翻譯是從翻譯引擎返回的全部文字中按照引擎參數的「翻譯開始自」和" & vbCrLf & _
			"  「翻譯結束到」欄位的設定，取出介於二者之間的文字。" & vbCrLf & vbCrLf & _
			"☆全部文字☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 從翻譯引擎返回的全部文字。該文字可能還包括網頁代碼。利用該文字可以知道「翻譯開始自」和" & vbCrLf & _
			"  「翻譯結束到」欄位應該如何設定。" & vbCrLf & vbCrLf & _
			"☆其他功能☆" & vbCrLf & _
			"============" & vbCrLf & _
			"- 說明" & vbCrLf & _
			"  點擊該按鈕，將獲取說明訊息。" & vbCrLf & vbCrLf & _
			"- 翻譯" & vbCrLf & _
			"  點擊該按鈕，將按照選擇的條件進行翻譯。" & vbCrLf & vbCrLf & _
			"- 響應頭" & vbCrLf & _
			"  點擊該按鈕，將獲取 HTTP 響應頭，以瞭解文字編碼等訊息。" & vbCrLf & vbCrLf & _
			"- 清空" & vbCrLf & _
			"  將現有翻譯內容和翻譯結果全部清空。點擊後將該按鈕將變為「讀入」狀態。" & vbCrLf & vbCrLf & _
			"- 讀入" & vbCrLf & _
			"  再次從選擇的翻譯清單中讀入字串。點擊後將該按鈕將變為「清空」狀態。" & vbCrLf & vbCrLf & _
			"- 取消" & vbCrLf & _
			"  結束測試程式並返回設定視窗。" & vbCrLf & vbCrLf & vbCrLf
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
			"- 本軟體在修改過程中得到漢化新世紀會員的測試，在此表示衷心的感謝！" & vbCrLf & vbCrLf & vbCrLf
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
	HelpTipTitle = "Passolo 婓盄楹祒粽"
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
			"羲 楷 氪ㄩ犖趙陔岍槨傖埜 wanfu (2010)" & vbCrLf & vbCrLf & vbCrLf
	Ement = "∵堍俴遠噫∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 盓厥粽揭燴腔 Passolo 6.0 摯眕奻唳掛ㄛ斛剒" & vbCrLf & _
			"- Windows Script Host (WSH) 勤砓 (VBS)ㄛ斛剒" & vbCrLf & _
			"- Adodb.Stream 勤砓ㄛ盓厥 Utf-8﹜Unicode 斛剒" & vbCrLf & _
			"- Microsoft.XMLHTTP 勤砓ㄛ斛剒" & vbCrLf & _
			"- Microsoft.XMLDOM 勤砓ㄛ賤昴 responseXML 殿隙跡宒垀剒" & vbCrLf & vbCrLf & vbCrLf
	Dec = "∵�篲�潠賡∵" & vbCrLf & _
			"============" & vbCrLf & _
			"Passolo 婓盄楹祒粽岆珨跺蚚衾 Passolo 楹祒蹈桶趼揹腔婓盄楹祒粽最唗﹝坳撿衄眕狟髡夔ㄩ" & vbCrLf & _
			"- 瞳蚚婓盄楹祒竘э赻雄楹祒 Passolo 楹祒蹈桶笢腔趼揹" & vbCrLf & _
			"- 摩傖賸珨虳翍靡腔婓盄楹祒竘эㄛ甜褫赻隅砱む坻婓盄楹祒竘э" & vbCrLf & _
			"- 褫恁寁趼揹濬倰﹜泐徹趼揹眕摯勤楹祒ヶ綴腔趼揹輛俴揭燴" & vbCrLf & _
			"- 摩傖辦豎瑩﹜笝砦睫﹜樓厒ん潰脤粽﹜褫婓楹祒綴潰脤甜壁淏楹祒笢腔渣昫" & vbCrLf & _
			"- 囀离褫赻隅砱腔赻雄載陔髡夔" & vbCrLf & vbCrLf & _
			"掛最唗婦漪狟蹈恅璃ㄩ" & vbCrLf & _
			"- PSLWebTrans.bas (Passolo 婓盄楹祒粽恅璃 - 勤趕遺源宒堍俴)" & vbCrLf & _
			"- PSLWebTrans_Silent.bas (Passolo 婓盄楹祒粽恅璃 - 噙蘇源宒堍俴)" & vbCrLf & _
			"- PSLWebTrans.txt (潠极笢恅佽隴恅璃)" & vbCrLf & vbCrLf & vbCrLf
	Setup = "∵假蚾源楊∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 蔚賤揤綴腔恅璃葩秶善 Passolo 炵苀恅璃標笢隅砱腔 Macros 恅璃標笢" & vbCrLf & _
			"- 婓 Passolo 腔馱撿 -> 赻隅砱馱撿粕等笢氝樓蜆粽恅璃甜隅砱蜆粕等靡備ㄛ" & vbCrLf & _
			"  森綴憩褫眕等僻蜆粕等眻諉覃蚚" & vbCrLf & _
			"- 噙蘇源宒堍俴腔婓盄楹祒粽躺釬蚚衾恁隅趼揹ㄛ泐徹趼揹睿趼揹揭燴偌勤趕遺源宒堍俴" & vbCrLf & _
 			"  腔婓盄楹祒粽扢离堍俴﹝猁隅砱蜆粽腔堍俴統杅ㄛ剒妏蚚勤趕遺源宒扢离" & vbCrLf & vbCrLf & vbCrLf
	MainUse = "∵楹祒竘э∵" & vbCrLf & _
			"============" & vbCrLf & _
			"最唗眒冪枑鼎珨虳翍靡腔婓盄楹祒竘э﹝蠟珩褫眕偌狟醱腔 [扢离] 偌聽赻隅砱饜离﹝" & vbCrLf & _
			"氝樓赻隅砱饜离綴ㄛ蠟褫眕婓楹祒竘э蹈桶笢恁寁砑妏蚚腔楹祒竘э﹝" & vbCrLf & _
			"衄壽赻隅砱楹祒竘эㄛ③湖羲饜离勤趕遺ㄛ等僻 [堆翑] 偌聽ㄛ統堐堆翑笢腔佽隴﹝" & vbCrLf & vbCrLf & _
			"∵楹祒埭恅∵" & vbCrLf & _
			"============" & vbCrLf & _
			"最唗赻雄蹈堤絞ヶ楹祒蹈桶垀扽恅璃笢壺絞ヶ楹祒蹈桶腔醴梓逄晟俋腔垀衄醴梓逄晟﹝" & vbCrLf & _
			"籵徹恁寁祥肮腔逄晟ㄛ褫眕恁寁絞ヶ楹祒蹈桶笢腔懂埭趼揹ㄛ麼氪肮珨恅璃笢腔む坻楹祒蹈桶笢" & vbCrLf & _
			"腔楹祒趼揹﹝" & vbCrLf & vbCrLf & _
			"- 恁寁迵絞ヶ楹祒蹈桶醴梓逄晟眈肮腔逄晟奀ㄛ蔚妏蚚絞ヶ楹祒蹈桶笢腔懂埭趼揹釬峈楹祒埭恅﹝" & vbCrLf & _
			"- 恁寁迵絞ヶ楹祒蹈桶醴梓逄晟祥肮腔逄晟奀ㄛ蔚妏蚚恁隅腔む坻楹祒蹈桶笢腔楹祒趼揹釬峈楹祒" & vbCrLf & _
			"  埭恅﹝眕源晞蔚眒衄腔楹祒蛌遙傖む坻逄晟﹝掀�蝵姨藚樛倛譟倡貐伢捺樛倛纂�" & vbCrLf & vbCrLf & _
			"∵楹祒趼揹∵" & vbCrLf & _
			"============" & vbCrLf & _
			"枑鼎賸�垓縑３佽央７堇倏礡Ｉ硊�揹桶﹜樓厒ん﹜唳掛﹜む坻﹜躺恁隅脹恁砐﹝" & vbCrLf & vbCrLf & _
			"- �蝜�恁寁�垓縛皈藫頖�等砐蔚掩赻雄�＋�恁寁﹝" & vbCrLf & _
			"- �蝜�恁寁等砐ㄛ寀�垓謀＋蹐垮閤堈紙＋�恁寁﹝" & vbCrLf & _
			"- 等砐褫眕嗣恁﹝む笢恁寁躺恁隅奀ㄛむ坻歙掩赻雄�＋�恁寁﹝" & vbCrLf & vbCrLf & _
			"∵泐徹趼揹∵" & vbCrLf & _
			"============" & vbCrLf & _
			"枑鼎賸鼎葩机﹜眒桄痐﹜帤楹祒﹜�屋羌�趼睿睫瘍﹜�屋玫鯗植卅纂〦屋肱－植卅警�恁砐﹝" & vbCrLf & vbCrLf & _
			"- 垀衄砐醴歙褫眕嗣恁﹝" & vbCrLf & _
			"- 恁寁�屋玫鯗植卅纂〦屋肱－植卅麵＋醡掃度堈耕▲亞屋羌�趼睿睫瘍恁砐ㄛ甜й�屋羌�趼" & vbCrLf & _
			"  睿睫瘍恁砐蔚赻雄曹峈祥褫蚚袨怓﹝" & vbCrLf & vbCrLf & _
			"∵趼揹揭燴∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 饜离" & vbCrLf & _
			"  最唗枑鼎賸睿辦豎瑩﹜笝砦睫﹜樓厒ん潰脤粽眈肮腔蘇�狣馺瓊爰藷馺藩奿埰弝譚痟騥鉌�①錶﹝" & vbCrLf & _
			"  蠟珩褫眕偌狟醱腔 [扢离] 偌聽赻隅砱饜离﹝" & vbCrLf & _
			"  氝樓赻隅砱饜离綴ㄛ蠟褫眕婓饜离蹈桶笢恁寁砑妏蚚腔饜离﹝" & vbCrLf & _
			"  衄壽赻隅砱饜离ㄛ③湖羲扢离勤趕遺ㄛ等僻 [堆翑] 偌聽ㄛ統堐堆翑笢腔佽隴﹝" & vbCrLf & vbCrLf & _
			"- 赻雄恁寁饜离" & vbCrLf & _
			"  恁隅蜆恁砐奀蔚跦擂饜离笢腔褫蚚逄晟蹈桶赻雄恁寁迵楹祒蹈桶醴梓逄晟ぁ饜腔饜离﹝" & vbCrLf & _
			"  ♁蛁砩ㄩ蜆恁砐躺勤婓盄楹祒粽衄虴﹝" & vbCrLf & _
			"  ﹛﹛﹛﹛猁蔚饜离迵絞ヶ楹祒蹈桶腔醴梓逄晟ぁ饜ㄛ偌 [扢离] 偌聽ㄛ婓趼揹揭燴腔巠蚚逄晟笢" & vbCrLf & _
			"  ﹛﹛﹛﹛氝樓眈茼腔逄晟﹝" & vbCrLf & _
			"  ﹛﹛﹛﹛婓噙蘇源宒①錶狟ㄛ最唗蔚偌※赻雄 - 赻恁 - 蘇�洁捨創藘√鵜隑朴霰馺獺�" & vbCrLf & vbCrLf & _
			"- �戊�辦豎瑩" & vbCrLf & _
			"  婓楹祒ヶ蔚趼揹笢腔辦豎瑩刉壺ㄛ眕晞楹祒竘э褫眕淏�毽倡諢�" & vbCrLf & vbCrLf & _
			"- �戊�樓厒ん" & vbCrLf & _
			"  婓楹祒ヶ蔚趼揹笢腔樓厒ん刉壺ㄛ眕晞楹祒竘э褫眕淏�毽倡諢�" & vbCrLf & vbCrLf & _
			"- 杸遙杻隅趼睫甜婓楹祒綴遜埻" & vbCrLf & _
			"  婓楹祒ヶ妏蚚恁隅腔饜离ㄛ杸遙杻隅趼睫甜婓楹祒綴遜埻﹝" & vbCrLf & vbCrLf & _
			"  猁氝樓涴虳杻隅趼睫ㄛ③婓饜离勤趕遺腔趼睫杸遙腔楹祒ヶ猁杸遙腔趼睫笢隅砱﹝" & vbCrLf & _
			"  ♁蛁砩ㄩ蚕衾粽竘э腔恀枙ㄛ衄虳趼睫岆祥褫眕掩怀�趥藑葰瑀迂銘雇臐�" & vbCrLf & _
			"  ﹛﹛﹛﹛�蝜�楹祒竘э楹祒賸涴虳杸遙綴腔趼睫ㄛ掩杸遙腔趼睫拸楊掩遜埻﹝" & vbCrLf & vbCrLf & _
			"- 煦俴楹祒" & vbCrLf & _
			"  婓楹祒ヶ蔚嗣俴腔趼揹莞煦峈等俴趼揹輛俴楹祒ㄛ眕晞枑詢楹祒竘э楹祒腔淏�毓�﹝" & vbCrLf & vbCrLf & _
			"- 壁淏辦豎瑩﹜笝砦睫﹜樓厒ん" & vbCrLf & _
			"  婓楹祒綴妏蚚恁隅腔饜离ㄛ潰脤甜壁淏楹祒笢腔渣昫﹝" & vbCrLf & _
			"  壁淏源宒睿辦豎瑩﹜笝砦睫﹜樓厒ん潰脤粽俇�峙閟活�" & vbCrLf & vbCrLf & _
			"- 杸遙杻隅趼睫" & vbCrLf & _
			"  婓楹祒綴妏蚚恁隅腔饜离ㄛ杸遙楹祒笢杻隅腔趼睫﹝" & vbCrLf & vbCrLf & _
			"  猁氝樓涴虳杻隅趼睫ㄛ③婓饜离勤趕遺腔趼睫杸遙腔楹祒綴猁杸遙腔趼睫笢隅砱﹝" & vbCrLf & _
			"  ♁蛁砩ㄩ蚕衾粽竘э腔恀枙ㄛ衄虳趼睫岆祥褫眕掩怀�趥藑葰瑀迂銘雇臐�" & vbCrLf & _
			"  ﹛﹛﹛﹛�蝜�楹祒竘э楹祒賸涴虳杸遙ヶ腔趼睫ㄛ蔚拸楊掩杸遙﹝" & vbCrLf & vbCrLf & _
			"∵む坻恁砐∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 樟哿奀赻雄悵湔垀衄恁寁" & vbCrLf & _
			"  恁隅蜆砐奀ㄛ蔚婓偌 [樟哿] 偌聽奀赻雄悵湔垀衄恁寁ㄛ狟棒堍俴奀蔚黍�貑ㄣ瘚麵√鞢�" & vbCrLf & vbCrLf & _
			"- 珆尨怀堤秏洘" & vbCrLf & _
			"  恁隅蜆砐奀ㄛ蔚婓 Passolo 秏洘怀堤敦諳笢怀堤楹祒﹜壁淏睿杸遙陓洘﹝" & vbCrLf & _
			"  婓祥恁隅蜆砐①錶狟ㄛ蔚珆翍枑詢最唗腔堍俴厒僅﹝" & vbCrLf & vbCrLf & _
			"- 氝樓楹祒蛁庋" & vbCrLf & _
			"  恁隅蜆砐奀ㄛ蔚婓楹祒蹈桶腔趼揹蛁庋笢崝樓恁隅楹祒竘э赻雄楹祒涴欴腔蛁庋﹝" & vbCrLf & _
			"  瞳蚚蜆蛁庋褫眕⑹煦楹祒趼揹腔楹祒堤揭﹝" & vbCrLf& vbCrLf & _
			"∵む坻髡夔∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 壽衾" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚壽衾勤趕遺﹝甜珆尨最唗賡庄﹜堍俴遠噫﹜羲楷妀摯唳�巡�陓洘﹝" & vbCrLf & vbCrLf & _
			"- 堆翑" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚粟堤絞ヶ敦諳腔堆翑陓洘﹝" & vbCrLf & vbCrLf & _
			"- 扢离" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚粟堤饜离勤趕遺﹝褫眕婓饜离勤趕遺笢饜离跪笱統杅﹝" & vbCrLf & vbCrLf & _
			"- �毓�" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚壽敕翋勤趕遺ㄛ甜偌恁隅腔恁砐輛俴趼揹楹祒﹝" & vbCrLf & _
			"  ♁蛁砩ㄩ楹祒綴腔趼揹楹祒袨怓蔚載蜊峈渾葩机袨怓ㄛ眕尨⑹梗﹝" & vbCrLf& vbCrLf & _
			"- �＋�" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚豖堤最唗﹝" & vbCrLf& vbCrLf & vbCrLf
	SetUse ="∵饜离蹈桶∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 恁寁饜离" & vbCrLf & _
			"  猁恁寁饜离砐ㄛ等僻饜离蹈桶﹝" & vbCrLf & vbCrLf & _
			"- 氝樓饜离" & vbCrLf & _
			"  猁氝樓饜离砐ㄛ等僻 [氝樓] 偌聽ㄛ婓粟堤腔勤趕遺笢怀�踼�備﹝" & vbCrLf & vbCrLf & _
			"- 載蜊饜离" & vbCrLf & _
			"  猁載蜊饜离靡備ㄛ恁寁饜离蹈桶笢猁蜊靡腔饜离ㄛ�遣騕本� [載蜊] 偌聽﹝" & vbCrLf & vbCrLf & _
			"- 刉壺饜离" & vbCrLf & _
			"  猁刉壺饜离砐ㄛ恁寁饜离蹈桶笢猁刉壺腔饜离ㄛ�遣騕本� [刉壺] 偌聽﹝" & vbCrLf & vbCrLf & _
			"氝樓饜离綴ㄛ蔚婓蹈桶笢珆尨陔腔饜离ㄛ饜离囀�斒峙埰噶欶窗�" & vbCrLf & _
			"載蜊饜离綴ㄛ蔚婓蹈桶笢珆尨蜊靡綴腔饜离ㄛ饜离囀�楺迮霰馺譆結跼銦�" & vbCrLf & _
			"刉壺饜离綴ㄛ蔚婓蹈桶笢珆尨奻珨跺饜离ㄛ饜离囀�斒峙埰導玾遘鷌馺譆窗�" & vbCrLf & vbCrLf & _
			"∵悵湔濬倰∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 恅璃" & vbCrLf & _
			"  饜离蔚眕恅璃倛宒悵湔婓粽垀婓恅璃標狟腔 Data 恅璃標笢﹝" & vbCrLf & vbCrLf & _
			"- 蛁聊桶" & vbCrLf & _
			"  饜离蔚掩悵湔蛁聊桶笢腔 HKCU\Software\VB and VBA Program Settings\WebTranslate 砐狟﹝" & vbCrLf & vbCrLf & _
			"- 絳�蹁馺�" & vbCrLf & _
			"  埰勍植む坻饜离恅璃笢絳�蹁馺獺ㄤ暫踾吇馺蟾掃垮閤堈紛�撰ㄛ珋衄饜离蹈桶笢眒衄腔饜离蔚" & vbCrLf & _
			"  掩載蜊ㄛ羶衄腔饜离蔚掩氝樓﹝" & vbCrLf & vbCrLf & _
			"- 絳堤饜离" & vbCrLf & _
			"  埰勍絳堤垀衄饜离善恅掛恅璃ㄛ眕晞褫眕蝠遙麼蛌痄饜离﹝" & vbCrLf & vbCrLf & _
			"♁蛁砩ㄩз遙悵湔濬倰奀ㄛ蔚赻雄刉壺埻衄弇离笢腔饜离囀�搳�" & vbCrLf & vbCrLf & _
			"∵饜离囀�搳�" & vbCrLf & _
			"============" & vbCrLf & _
			"<竘э統杅>" & vbCrLf & _
			"  - 妏蚚勤砓" & vbCrLf & _
			"    蜆勤砓蘇�玴炕衽icrosoft.XMLHTTP§ㄛ祥褫載蜊﹝" & vbCrLf & _
			"    衄壽蜆勤砓腔珨虳妏蚚源楊ㄛ③統堐眈壽恅紫﹝" & vbCrLf & vbCrLf & _
			"  - 竘э蛁聊 ID" & vbCrLf & _
			"    衄虳楹祒竘э剒猁蛁聊符夔妏蚚﹝眈壽恀枙③迵楹祒竘э枑鼎妀薊炵﹝" & vbCrLf & _
			"    最唗眒冪枑鼎賸 Microsoft 楹祒竘э腔蛁聊 ID﹝筍祥悵痐褫眕蚗壅妏蚚﹝" & vbCrLf & vbCrLf & _
			"  - 竘э厙硊" & vbCrLf & _
			"    楹祒竘э腔溼恀厙硊﹝" & vbCrLf & _
			"    ♁蛁砩ㄩ勍嗣婓盄楹祒枑鼎妀腔厙硊甜祥岆淩淏腔楹祒竘э溼恀厙硊﹝剒猁煦昴厙珜測鎢麼" & vbCrLf & _
			"    ﹛﹛﹛﹛戙恀淩淏腔楹祒竘э枑鼎妀符夔�△獺�" & vbCrLf & vbCrLf & _
			"  - 楷冞囀�暊ㄟ�" & vbCrLf & _
			"    蜆趼僇妗暱奻岆 Microsoft.XMLHTTP 腔 Open 源楊笢腔 bstrUrl 統杅﹝" & vbCrLf & _
			"    ♁蛁砩ㄩ湮嬤瘍 {} 笢腔趼睫峈炵苀趼僇ㄛ祥褫載蜊ㄛ瘁寀炵苀拸楊妎梗﹝" & vbCrLf & _
			"    ﹛﹛﹛﹛猁氝樓涴虳炵苀趼僇ㄛ③等僻衵晚腔※>§偌聽ㄛ恁寁剒猁腔炵苀趼僇﹝" & vbCrLf & vbCrLf & _
			"  - 杅擂換冞源宒" & vbCrLf & _
			"    蜆趼僇妗暱奻岆 Microsoft.XMLHTTP 勤砓腔 Open 源楊笢腔 bstrMethod 統杅﹝" & vbCrLf & _
			"    撈 GET 麼 POST﹝蚚 POST 源宒楷冞杅擂,褫眕湮善 4MBㄛ珩褫眕峈 GETㄛ硐夔 256KB﹝" & vbCrLf & vbCrLf & _
			"  - 肮祭源宒" & vbCrLf & _
			"    蜆趼僇妗暱奻岆 Microsoft.XMLHTTP 勤砓腔 Open 源楊笢腔 varAsync 統杅﹝褫吽謹﹝" & vbCrLf & _
			"    �捩﹡� Trueㄛ撈肮祭硒俴ㄛ筍硐夔婓 DOM 笢妗囥肮祭硒俴﹝珨啜蔚む离峈 Falseㄛ撈祑祭硒俴﹝" & vbCrLf & vbCrLf & _
			"  - 蚚誧靡" & vbCrLf & _
			"    蜆趼僇妗暱奻岆 Microsoft.XMLHTTP 勤砓腔 Open 源楊笢腔 bstrUser 統杅﹝褫吽謹﹝" & vbCrLf & vbCrLf & _
			"  - 躇鎢" & vbCrLf & _
			"    蜆趼僇妗暱奻岆 Microsoft.XMLHTTP 勤砓腔 Open 源楊笢腔 bstrPassword 統杅﹝褫吽謹﹝" & vbCrLf & vbCrLf & _
			"  - 硌鍔摩" & vbCrLf & _
			"    蜆趼僇妗暱奻岆 Microsoft.XMLHTTP 勤砓腔 Send 源楊笢腔 varBody 統杅﹝褫吽謹﹝" & vbCrLf & _
			"    坳褫眕岆 XML 跡宒杅擂ㄛ珩褫眕岆趼睫揹﹜霜ㄛ麼氪珨跺拸睫瘍淕杅杅郪﹝�譆蜂鎡邦�" & vbCrLf & _
			"    Open 源楊笢腔 URL 統杅測�諢�" & vbCrLf & vbCrLf & _
			"    楷冞杅擂腔源宒煦峈肮祭睿祑祭謗笱ㄩ" & vbCrLf & _
			"    (1) 祑祭源宒ㄩ杅擂婦珨筒楷冞俇救ㄛ憩賦旰 Send 輛最ㄛ諦誧儂硒俴む坻腔紱釬﹝" & vbCrLf & _
			"    (2) 肮祭源宒ㄩ諦誧儂猁脹善督昢ん殿隙�溜珫�洘綴符賦旰 Send 輛最﹝" & vbCrLf & vbCrLf & _
			"  - HTTP 芛睿硉" & vbCrLf & _
			"    蜆趼僇妗暱奻岆 Microsoft.XMLHTTP 勤砓笢腔 setRequestHeader 源楊ㄛGET 源宒狟拸虴﹝" & vbCrLf & _
			"    (1) 蜆趼僇盓厥嗣跺 setRequestHeader 砐醴ㄛ藩跺砐醴蚚煦俴煦路﹝" & vbCrLf & _
			"    (2) 藩跺砐醴腔芛靡備 (bstrHeader) 睿硉 (bstrValue) 蚚圉褒飯瘍煦路﹝" & vbCrLf & _
			"    (3) 剒猁楷冞 ContentT-Length 砐醴奀ㄛむ硉褫眕妏蚚炵苀趼僇ㄛ炵苀蔚赻雄數呾む酗僅﹝" & vbCrLf & vbCrLf & _
			"  - 殿隙賦彆跡宒統杅" & vbCrLf & _
			"    蜆趼僇妗暱奻岆 Microsoft.XMLHTTP 勤砓笢腔珨跺殿隙賦彆跡宒扽俶﹝衄眕狟撓笱ㄩ" & vbCrLf & _
			"    responseBody	Variant	倰	賦彆殿隙峈拸睫瘍淕杅杅郪" & vbCrLf & _
			"    responseStream	Variant	倰	賦彆殿隙峈 ADO Stream 勤砓" & vbCrLf & _
			"    responseText	String	倰	賦彆殿隙峈趼睫揹" & vbCrLf & _
			"    responseXML	Object	倰	賦彆殿隙峈 XML 跡宒杅擂" & vbCrLf & vbCrLf & _
			"    ♁蛁砩ㄩ絞恁寁 responseXML 奀ㄛ※楹祒羲宎赻§睿※楹祒賦旰善§頗煦梗珆尨峈※偌 ID 刲坰§" & vbCrLf & _
			"    ﹛﹛﹛﹛睿※偌梓ワ靡刲坰§﹝" & vbCrLf & vbCrLf & _
			"  - 楹祒羲宎赻" & vbCrLf & _
			"    蜆趼僇蚚衾妎梗殿隙賦彆笢腔楹祒趼揹ヶ腔趼揹﹝" & vbCrLf & _
			"    褫等僻衵晚腔※...§偌聽脤艘植督昢ん殿隙腔垀衄恅掛笢茼蜆怀�賮儷硊�﹝" & vbCrLf & _
			"    ♁蛁砩ㄩ蜆趼僇盓厥蚚※|§煦路腔嗣跺麼砐砐醴﹝" & vbCrLf & vbCrLf & _
			"  - 楹祒賦旰善" & vbCrLf & _
			"    蜆趼僇蚚衾妎梗殿隙賦彆笢腔楹祒趼揹綴腔趼揹﹝" & vbCrLf & _
			"    褫等僻衵晚腔※...§偌聽脤艘植督昢ん殿隙腔垀衄恅掛笢茼蜆怀�賮儷硊�﹝" & vbCrLf & _
			"    ♁蛁砩ㄩ蜆趼僇盓厥蚚※|§煦路腔嗣跺麼砐砐醴﹝" & vbCrLf & vbCrLf & _
			"  - 偌 ID 刲坰" & vbCrLf & _
			"    蜆趼僇妏蚚 XML DOM 腔 getElementById 源楊脤梑 XML 恅璃笢撿衄硌隅 ID 腔啋匼笢腔恅掛硉﹝" & vbCrLf & _
			"    褫等僻衵晚腔※...§偌聽脤艘植督昢ん殿隙腔垀衄恅掛笢茼蜆怀�賮� ID 瘍﹝" & vbCrLf & _
			"    ♁蛁砩ㄩ蜆趼僇盓厥蚚※|§煦路腔嗣跺麼砐砐醴﹝" & vbCrLf & vbCrLf & _
			"  - 偌梓ワ靡刲坰" & vbCrLf & _
			"    蜆趼僇妏蚚 XML DOM 腔 getElementsByTagName 源楊脤梑 XML 恅璃笢垀衄撿衄硌隅梓ワ靡備腔" & vbCrLf & _
			"    啋匼腔恅掛硉﹝" & vbCrLf & _
			"    褫等僻衵晚腔※...§偌聽脤艘植督昢ん殿隙腔垀衄恅掛笢茼蜆怀�賮覺糒怔�備﹝" & vbCrLf & _
			"    ♁蛁砩ㄩ蜆趼僇盓厥蚚※|§煦路腔嗣跺麼砐砐醴﹝" & vbCrLf & vbCrLf & _
			"<逄晟饜勤>" & vbCrLf & _
			"  - 逄晟靡備" & vbCrLf & _
			"    蘇�炵鼯擿埼�備蹈桶岆蔚 Passolo 腔逄晟蹈桶輛俴賸儕潠ㄛ�扔蘅佴�模/華⑹腔逄晟砐﹝" & vbCrLf & vbCrLf & _
			"  - Passolo 測鎢" & vbCrLf & _
			"    蘇�炵鼯擿埭�鎢�＝� Passolo 逄晟蹈桶笢腔 ISO 936-1 逄晟測鎢﹝む笢ㄩ" & vbCrLf & _
			"    潠极笢恅睿楛极笢恅腔逄晟測鎢�＝� Passolo 逄晟蹈桶笢腔弊模/華⑹逄晟測鎢﹝" & vbCrLf & vbCrLf & _
			"  - 楹祒竘э測鎢" & vbCrLf & _
			"    楹祒竘э垀寞隅腔逄晟測鎢﹝剒猁煦昴厙珜測鎢麼戙恀楹祒竘э枑鼎妀符夔�△獺�" & vbCrLf & vbCrLf & _
			"  - 崝樓逄晟" & vbCrLf & _
			"    猁崝樓逄晟饜勤蹈桶砐ㄛ等僻 [氝樓] 偌聽ㄛ蔚粟堤氝樓勤趕遺﹝" & vbCrLf & vbCrLf & _
			"  - 刉壺逄晟" & vbCrLf & _
			"    猁刉壺逄晟饜勤蹈桶砐ㄛ恁寁蹈桶笢猁刉壺腔逄晟ㄛ�遣騕本� [刉壺] 偌聽﹝" & vbCrLf & vbCrLf & _
			"  - 晤憮逄晟" & vbCrLf & _
			"    猁晤憮逄晟饜勤蹈桶砐ㄛ恁寁蹈桶笢猁晤憮腔逄晟ㄛ�遣騕本� [晤憮] 偌聽﹝" & vbCrLf & vbCrLf & _
			"  - 俋窒晤憮逄晟" & vbCrLf & _
			"    覃蚚囀离麼俋窒晤憮最唗晤憮淕跺逄晟饜勤蹈桶﹝" & vbCrLf & vbCrLf & _
			"  - 离諾逄晟" & vbCrLf & _
			"    猁蔚楹祒竘э測鎢离諾ㄛ恁寁蹈桶笢猁离諾腔逄晟ㄛ�遣騕本� [离諾] 偌聽﹝" & vbCrLf & _
			"    [离諾] 偌聽頗跦擂楹祒竘э測鎢岆瘁峈諾 (Null) 奧珆尨祥肮腔褫蚚袨怓﹝" & vbCrLf & vbCrLf & _
			"  - 笭离逄晟" & vbCrLf & _
			"    猁蔚楹祒竘э測鎢笭离峈埻宎硉ㄛ恁寁蹈桶笢猁笭离腔逄晟ㄛ�遣騕本� [笭离] 偌聽﹝" & vbCrLf & _
			"    [笭离] 偌聽硐衄婓楹祒竘э測鎢掩載蜊綴符夔蛌峈褫蚚袨怓﹝" & vbCrLf & vbCrLf & _
			"  - 珆尨準諾砐" & vbCrLf & _
			"    猁躺珆尨楹祒竘э測鎢峈準諾腔逄晟砐ㄛ等僻 [珆尨準諾砐] 偌聽﹝等僻綴蜆偌聽" & vbCrLf & _
			"    蔚蛌峈祥褫蚚袨怓ㄛ[�垓諫埰霄 睿 [珆尨諾砐] 偌聽蔚蛌峈褫蚚袨怓﹝" & vbCrLf & vbCrLf & _
			"  - 珆尨諾砐" & vbCrLf & _
			"    猁躺珆尨楹祒竘э測鎢峈諾 (Null) 腔逄晟砐ㄛ等僻 [珆尨諾砐] 偌聽﹝等僻綴" & vbCrLf & _
			"    蜆偌聽蔚蛌峈祥褫蚚袨怓ㄛ[�垓諫埰霄 睿 [珆尨準諾砐] 偌聽蔚蛌峈褫蚚袨怓﹝" & vbCrLf & vbCrLf & _
			"  - �垓諫埰�" & vbCrLf & _
			"    猁珆尨�垓諧擿婘謑炸本� [�垓諫埰霄 偌聽﹝等僻綴蜆偌聽蔚蛌峈祥褫蚚袨怓ㄛ" & vbCrLf & _
			"    [珆尨準諾砐] 睿 [珆尨準諾砐] 偌聽蔚蛌峈褫蚚袨怓﹝" & vbCrLf & vbCrLf & _
			"  氝樓逄晟綴ㄛ蔚婓蹈桶笢珆尨甜恁隅陔腔逄晟摯測鎢勤﹝" & vbCrLf & _
			"  刉壺逄晟綴ㄛ蔚婓蹈桶笢恁隅垀刉壺逄晟腔ヶ珨砐逄晟﹝" & vbCrLf & _
			"  晤憮逄晟綴ㄛ蔚婓蹈桶笢珆尨甜悵厥恁隅晤憮綴腔逄晟摯測鎢勤﹝" & vbCrLf & _
			"  离諾逄晟綴ㄛ蔚婓蹈桶笢珆尨甜悵厥恁隅离諾綴腔逄晟摯測鎢勤﹝" & vbCrLf & _
			"  笭离逄晟綴ㄛ蔚婓蹈桶笢珆尨甜悵厥恁隅笭离綴腔逄晟摯測鎢勤﹝" & vbCrLf & vbCrLf & _
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
			"- 聆彸" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚粟堤聆彸勤趕遺ㄛ眕晞潰脤饜离腔淏�煩唌�" & vbCrLf & vbCrLf & _
			"- ь諾" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚ь諾珋衄饜离腔�垓謁童皆埸蔣蜣寪薹靿蹁馺譆窗�" & vbCrLf & vbCrLf & _
			"- �毓�" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚悵湔饜离敦諳笢腔�庥庛�蜊ㄛ豖堤饜离敦諳甜殿隙翋敦諳﹝" & vbCrLf & _
			"  最唗蔚妏蚚載蜊綴腔饜离硉﹝" & vbCrLf & vbCrLf & _
			"- �＋�" & vbCrLf & _
			"  等僻蜆偌聽ㄛ祥悵湔饜离敦諳笢腔�庥庛�蜊ㄛ豖堤饜离敦諳甜殿隙翋敦諳﹝" & vbCrLf & _
			"  最唗蔚妏蚚埻懂腔饜离硉﹝" & vbCrLf & vbCrLf & vbCrLf
	TestUse ="∵楹祒竘э∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 猁聆彸腔楹祒竘э靡備﹝猁恁寁楹祒竘эㄛ等僻楹祒竘э蹈桶﹝" & vbCrLf & vbCrLf & _
			"∵醴梓逄晟∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 楹祒腔醴梓逄晟﹝蜆蹈桶躺珆尨逄晟饜勤蹈桶笢楹祒竘э逄晟測鎢祥峈諾腔逄晟﹝" & vbCrLf & vbCrLf & _
			"∵楹祒蹈桶∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 楹祒埭恅腔楹祒蹈桶埭﹝蜆蹈桶婦漪賸源偶笢腔垀衄楹祒蹈桶﹝" & vbCrLf & vbCrLf & _
			"∵黍�遶倇�∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 硌隅植恁隅楹祒蹈桶笢黍�賮儷硒晉�﹝" & vbCrLf & vbCrLf & _
			"∵楹祒囀�搳�" & vbCrLf & _
			"============" & vbCrLf & _
			"- 猁楹祒腔恅掛﹝恁寁楹祒蹈桶﹜懂埭趼揹睿楹祒趼揹恁砐奀赻雄植恁隅腔楹祒蹈桶笢黍�鄶硒恣�" & vbCrLf & _
			"- 掩黍�賮儷硒案堈耗藩楖魌狠荂�" & vbCrLf & vbCrLf & _
			"∵懂埭趼揹∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 恁寁蜆恁砐奀ㄛ赻雄植恁隅腔楹祒蹈桶笢黍�蹀椒棚硒恣�" & vbCrLf & vbCrLf & _
			"∵楹祒趼揹∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 恁寁蜆恁砐奀ㄛ赻雄植恁隅腔楹祒蹈桶笢黍�賰倡鄶硒恣�" & vbCrLf & vbCrLf & _
			"∵楹祒賦彆∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 植楹祒竘э殿隙腔賦彆﹝跦擂恁隅腔懂埭趼揹睿楹祒趼揹恁砐ㄛ煦梗珆尨楹祒恅掛麼植楹祒竘э" & vbCrLf & _
			"  殿隙腔�垓諺覺鴃�" & vbCrLf & vbCrLf & _
			"∵楹祒恅掛∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 植楹祒竘э殿隙腔楹祒﹝蜆楹祒岆植楹祒竘э殿隙腔�垓諺覺擁訄椎梑�э統杅腔※楹祒羲宎赻§睿" & vbCrLf & _
			"  ※楹祒賦旰善§趼僇腔扢离ㄛ�○鶵橔痗�氪眳潔腔恅掛﹝" & vbCrLf & vbCrLf & _
			"∵�垓諺覺鴃�" & vbCrLf & _
			"============" & vbCrLf & _
			"- 植楹祒竘э殿隙腔�垓諺覺鴃ㄧ襞覺噶厊僈僭�嬤厙珜測鎢﹝瞳蚚蜆恅掛褫眕眭耋※楹祒羲宎赻§睿" & vbCrLf & _
			"  ※楹祒賦旰善§趼僇茼蜆�蝥恌髲獺�" & vbCrLf & vbCrLf & _
			"∵む坻髡夔∵" & vbCrLf & _
			"============" & vbCrLf & _
			"- 堆翑" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚鳳�※攃�陓洘﹝" & vbCrLf & vbCrLf & _
			"- 楹祒" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚偌桽恁隅腔沭璃輛俴楹祒﹝" & vbCrLf & vbCrLf & _
			"- 砒茼芛" & vbCrLf & _
			"  等僻蜆偌聽ㄛ蔚鳳�� HTTP 砒茼芛ㄛ眕賸賤恅掛晤鎢脹陓洘﹝" & vbCrLf & vbCrLf & _
			"- ь諾" & vbCrLf & _
			"  蔚珋衄楹祒囀�搡芛倡踳廜��垓褲敹捸ㄤ本鷚騣姜簸棠末垮餂炕偉賺諢斜棧活�" & vbCrLf & vbCrLf & _
			"- 黍��" & vbCrLf & _
			"  婬棒植恁隅腔楹祒蹈桶笢黍�鄶硒恣ㄤ本鷚騣姜簸棠末垮餂炕勒敹捸斜棧活�" & vbCrLf & vbCrLf & _
			"- �＋�" & vbCrLf & _
			"  豖堤聆彸最唗甜殿隙饜离敦諳﹝" & vbCrLf & vbCrLf & vbCrLf
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
			"- 掛�篲�婓党蜊徹最笢腕善犖趙陔岍槨頗埜腔聆彸ㄛ婓森桶尨笪陑腔覜郅ㄐ" & vbCrLf & vbCrLf & vbCrLf
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
