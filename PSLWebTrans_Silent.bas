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

Public srcLineNum As Integer,trnLineNum As Integer,srcAccKeyNum As Integer,trnAccKeyNum As Integer
Public DefaultCheckList() As String,AppRepStr As String,PreRepStr As String
Public CheckDataList() As String,RepString As Integer,WaitTimes As Long

Public DefaultEngineList() As String,EngineDataList() As String,tSelected() As String

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


'翻译引擎默认设置
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


' 主程序
Sub Main
	Dim i As Integer,j As Integer,srcString As String,trnString As String,TranLang As String
	Dim src As PslTransList,LangPairList() As String,xmlHttp As Object,objStream As Object
	Dim srcLngFind As Integer,trnLngFind As Integer,StringCount As Integer
	Dim srcLng As String,trnLng As String,TransListOpen As Boolean
	Dim TranedCount As Integer,SkipedCount As Integer,NotChangeCount As Integer
	Dim LineNumErrCount As Integer,accKeyNumErrCount As Integer,ErrorCount As Integer

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
		Msg03 = "陆亩睲虫: "
		Msg05 = "眤╰参ぶ Microsoft.XMLHTTP ン礚猭磅︽"
		Msg06 = "陆亩ま篮狝竟臫莱筄惠璶┑单丁盾"
		Msg07 = "单丁:"
		Msg08 = "叫块单丁"
		Msg09 = "礚猭籔陆亩ま篮狝竟硄獺琌礚 Internet 硈钡" & vbCrLf & _
				"┪陆亩ま篮砞﹚岿粇┪陆亩ま篮窽ゎ"
		Msg10 = "陆亩ま篮呼礚猭膥尿"
		Msg11 = "陆亩ま篮: "
		Msg12 = "﹃矪瞶: "
		Msg42 = "絋粄"
		Msg43 = "癟"
		Msg44 = "岿粇"
		Msg45 =	"眤 Passolo セびセエ栋度続ノ Passolo 6.0 のセ叫どㄏノ"
		Msg46 = "叫匡陆亩睲虫"
		Msg47 = "タミ㎝穝陆亩睲虫..."
		Msg48 = "礚猭ミ㎝穝陆亩睲虫叫浪琩眤盡砞﹚"
		Msg49 = "ま篮笆陆亩"
		Msg50 = "赣睲虫ゼ砆秨币篈ぃ秈︽硈絬陆亩" & vbCrLf & _
				"眤惠璶琵╰参笆秨币赣陆亩睲虫盾"
		Msg51 = "タ秨币陆亩睲虫..."
		Msg52 = "礚猭秨币陆亩睲虫叫浪琩眤盡砞﹚"
		Msg53 = "赣睲虫矪秨币篈篈秈︽硈絬陆亩盢ㄏ眤" & vbCrLf & _
				"ゼ纗陆亩礚猭临╰参盢纗眤" & vbCrLf & _
				"陆亩礛秈︽硈絬陆亩" & vbCrLf & vbCrLf & _
				"眤絋﹚璶琵╰参笆纗眤陆亩盾"
		Msg54 = "タミ㎝穝陆亩ㄓ方睲虫..."
		Msg55 = "礚猭ミ㎝穝陆亩ㄓ方睲虫叫浪琩眤盡砞﹚"
		Msg56 = "赣陆亩睲虫ヘ夹粂ē┮癸莱陆亩ま篮粂ē絏祘Α盢挡"
		Msg57 = "タ陆亩㎝矪瞶﹃惠璶碭だ牧叫祔獼..."
		Msg58 = "铬筁"
		Msg59 = "陆亩"
		Msg60 = "ゼ跑陆亩挡狦㎝瞷Τ陆亩"
		Msg62 = "﹃玛﹚"
		Msg63 = "﹃斑弄"
		Msg64 = "﹃陆亩ㄑ滦糵"
		Msg65 = "﹃陆亩喷靡"
		Msg66 = "﹃ゼ陆亩"
		Msg67 = "﹃┪"
		Msg68 = "﹃计㎝才腹"
		Msg69 = "﹃糶璣ゅ┪计じ腹"
		Msg70 = "﹃糶璣ゅ┪计じ腹"
		Msg71 = "璸ノ: "
		Msg72 = "hh  mm だ ss "
		Msg73 = "璣ゅいゅ"
		Msg74 = "いゅ璣ゅ"
		Msg75 = ""
		Msg76 = ""
		Msg77 = ""
		Msg78 = "眤╰参ぶ Adodb.Stream ン礚猭膥尿磅︽"
	Else
		Msg03 = "翻译列表: "
		Msg05 = "您的系统缺少 Microsoft.XMLHTTP 对象，无法运行！"
		Msg06 = "翻译引擎服务器响应超时！需要延长等待时间吗？"
		Msg07 = "等待时间:"
		Msg08 = "请输入等待时间"
		Msg09 = "无法与翻译引擎服务器通信！可能是无 Internet 连接，" & vbCrLf & _
				"或者翻译引擎的设置错误，或者翻译引擎禁止访问。"
		Msg10 = "翻译引擎网址为空，无法继续！"
		Msg11 = "翻译引擎: "
		Msg12 = "字串处理: "
		Msg42 = "确认"
		Msg43 = "信息"
		Msg44 = "错误"
		Msg45 =	"您的 Passolo 版本太低，本宏仅适用于 Passolo 6.0 及以上版本，请升级后再使用。"
		Msg46 = "请选择一个翻译列表！"
		Msg47 = "正在创建和更新翻译列表..."
		Msg48 = "无法创建和更新翻译列表，请检查您的方案设置。"
		Msg49 = "引擎自动翻译"
		Msg50 = "该列表未被打开。此状态下不可以进行在线翻译。" & vbCrLf & _
				"您需要让系统自动打开该翻译列表吗？"
		Msg51 = "正在打开翻译列表..."
		Msg52 = "无法打开翻译列表，请检查您的方案设置。"
		Msg53 = "该列表已处于打开状态。此状态下进行在线翻译将使您" & vbCrLf & _
				"未保存的翻译无法还原。为了安全，系统将先保存您的" & vbCrLf & _
				"翻译，然后进行在线翻译。" & vbCrLf & vbCrLf & _
				"您确定要让系统自动保存您的翻译吗？"
		Msg54 = "正在创建和更新翻译来源列表..."
		Msg55 = "无法创建和更新翻译来源列表，请检查您的方案设置。"
		Msg56 = "该翻译列表目标语言所对应的翻译引擎语言代码为空，程序将退出。"
		Msg57 = "正在翻译和处理字串，可能需要几分钟，请稍侯..."
		Msg58 = "已跳过"
		Msg59 = "已翻译"
		Msg60 = "未更改，翻译结果和现有翻译相同。"
		Msg62 = "字串已锁定。"
		Msg63 = "字串只读。"
		Msg64 = "字串已翻译供复审。"
		Msg65 = "字串已翻译并验证。"
		Msg66 = "字串未翻译。"
		Msg67 = "字串为空或全为空格。"
		Msg68 = "字串全为数字和符号。"
		Msg69 = "字串全为大写英文或数字符号。"
		Msg70 = "字串全为小写英文或数字符号。"
		Msg71 = "合计用时: "
		Msg72 = "hh 小时 mm 分 ss 秒"
		Msg73 = "英文到中文"
		Msg74 = "中文到英文"
		Msg75 = "，"
		Msg76 = "并"
		Msg77 = "。"
		Msg78 = "您的系统缺少 Adodb.Stream 对象，无法继续运行！"
	End If

	If PSL.Version < 600 Then
		MsgBox Msg45,vbOkOnly+vbInformation,Msg43
		Exit Sub
	End If

	'检测 Adodb.Stream 是否存在并获取字符编码列表
	Set objStream = CreateObject("Adodb.Stream")
	If objStream Is Nothing Then
		MsgBox(Msg78,vbOkOnly+vbInformation,Msg43)
		Exit Sub
	End If
	Set objStream = Nothing

	Set trn = PSL.ActiveTransList
	'检测翻译列表是否被选择
	If trn Is Nothing Then
		MsgBox Msg46,vbOkOnly+vbInformation,Msg43
		Exit Sub
	End If

	'初始化数组
	ReDim DefaultEngineList(2),DefaultCheckList(1)
	DefaultEngineList(0) = "Microsoft"
	DefaultEngineList(1) = "Google"
	DefaultEngineList(2) = "Yahoo"
	DefaultCheckList(0) = Msg73
	DefaultCheckList(1) = Msg74

	'读取翻译引擎设置
	If EngineGet("",EngineDataList,"") = False Then
		EngineName = DefaultEngineList(0)
		LangPair = Join(LangCodeList(EngineName,OSLanguage,0,107),SubJoinStr)
		Temp = EngineName & JoinStr & EngineSettings(EngineName) & JoinStr & LangPair
		EngineDataList = Split(Temp,JoinStr)
	End If

	If Join(tSelected) <> "" Then
		'EngineName = tSelected(0)
		CheckName = tSelected(1)
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
		mSplitTrn = StrToInteger(tSelected(20))
		mCheckTrn = StrToInteger(tSelected(21))
		mAppStrRep = StrToInteger(tSelected(22))
		'KeepSet = StrToInteger(tSelected(23))
		mShowMsg = StrToInteger(tSelected(24))
		mTranComm = StrToInteger(tSelected(25))
	End If

	'获取字串类型组合
	If mMenu = 1 Then StrTypes = "|Menu|"
	If mDialog = 1 Then StrTypes = StrTypes & "|Dialog|"
	If mString = 1 Then StrTypes = StrTypes & "|StringTable|"
	If mAccTable = 1 Then StrTypes = StrTypes & "|AcceleratorTable|"
	If mVer = 1 Then StrTypes = StrTypes & "|Version|"

	'检测 Microsoft.XMLHTTP 是否存在
	Set xmlHttp = CreateObject(DefaultObject)
	If xmlHttp Is Nothing Then Set xmlHttp = CreateObject("Msxml2.XMLHTTP")
	If xmlHttp Is Nothing Then
		MsgBox(Msg05,vbOkOnly+vbInformation,Msg44)
		GoTo ExitSub
	Else
		'获取测试翻译
		trnString = getTranslate(xmlHttp,"Test","",3)
		Set xmlHttp = Nothing
		'测试 Internet 连接
		If trnString = "NotConnected" Then
			MsgBox(Msg09,vbOkOnly+vbInformation,Msg44)
			GoTo ExitSub
		End If
		'测试引擎网址是否为空
		If trnString = "NullUrl" Then
			MsgBox(Msg10,vbOkOnly+vbInformation,Msg44)
			GoTo ExitSub
		End If
		'测试引擎引擎是否超时
		If trnString = "Timeout" Then
			Massage = MsgBox(Msg06,vbYesNoCancel+vbInformation,Msg42)
			If Massage = vbYes Then WaitTimes = InputBox(Msg07,Msg08,WaitTimes)
		End If
	End If

	'提示打开关闭的翻译列表，以便可以在线翻译
	TransListOpen = False
	If trn.IsOpen = False Then
		Msg03 = Msg03 & trn.Title & " - " & PSL.GetLangCode(trn.Language.LangID,pslCodeText)
		Massage = MsgBox(Msg03 & vbCrLf & vbCrLf & Msg50,vbOkCancel,Msg42)
		If Massage = vbOK Then
			PSL.Output Msg51
			If trn.Open = False Then
				MsgBox Msg51,vbOkOnly+vbInformation,Msg44
				GoTo ExitSub
			Else
				TransListOpen = True
			End If
		End If
		If Massage = vbCancel Then GoTo ExitSub
	End If

	'提示保存打开的翻译列表，以免处理后数据不可恢复
	'If trn.IsOpen = True Then
		'Massage = MsgBox(Msg03 & vbCrLf & vbCrLf & Msg53,vbYesNoCancel,Msg42)
		'If Massage = vbYes Then trn.Save
		'If Massage = vbCancel Then Goto ExitSub
	'End If

	'如果翻译列表的更改时间晚于原始列表，自动更新
	If trn.SourceList.LastChange > trn.LastChange Then
		PSL.Output Msg47
		If trn.Update = False Then
			MsgBox Msg48,vbOkOnly+vbInformation,Msg44
			GoTo ExitSub
		End If
	End If

	'设置检查宏专用的用户定义属性
	If Not trn Is Nothing Then
		If trn.Property(19980) <> "CheckAccessKeys" Then
			trn.Property(19980) = "CheckAccessKeys"
		End If
	End If

	'获取PSL的来源语言代码
	srcLng = PSL.GetLangCode(trn.SourceList.LangID,pslCode639_1)
	If srcLng = "zh" Then
		srcLng = PSL.GetLangCode(trn.SourceList.LangID,pslCodeLangRgn)
		If srcLng = "zh-CHS" Or srcLng = "zh-SG" Then srcLng = "zh-CN"
		If srcLng = "zh-CHT" Or srcLng = "zh-HK" Or srcLng = "zh-MO" Then srcLng = "zh-TW"
	End If

	'获取PSL的目标语言代码
	trnLng = PSL.GetLangCode(trn.Language.LangID,pslCode639_1)
	If trnLng = "zh" Or trnLng = "ja" Or trnLng = "ko" Then TranLang = "Asia"
	If trnLng = "zh" Then
		trnLng = PSL.GetLangCode(trn.Language.LangID,pslCodeLangRgn)
		If trnLng = "zh-CHS" Or trnLng = "zh-SG" Then trnLng = "zh-CN"
		If trnLng = "zh-CHT" Or trnLng = "zh-HK" Or trnLng = "zh-MO" Then trnLng = "zh-TW"
	End If

	'读取字串处理设置
	If mAutoSele = 1 Then
		If CheckGet(trnLng,CheckDataList,"") = False Then
			If CheckGet(CheckName,CheckDataList,"") = False Then
				If TranLang = "Asia" Then CheckName = DefaultCheckList(0)
				If TranLang <> "Asia" Then CheckName = DefaultCheckList(1)
				LangPair = Join(LangCodeList(CheckName,OSLanguage,1,107),SubJoinStr)
				Temp = CheckName & JoinStr & CheckSettings(CheckName,OSLanguage) & JoinStr & LangPair
				CheckDataList = Split(Temp,JoinStr)
			End If
		End If
	Else
		If CheckGet(CheckName,CheckDataList,"") = False Then
			If TranLang = "Asia" Then CheckName = DefaultCheckList(0)
			If TranLang <> "Asia" Then CheckName = DefaultCheckList(1)
			LangPair = Join(LangCodeList(CheckName,OSLanguage,1,107),SubJoinStr)
			Temp = CheckName & JoinStr & CheckSettings(CheckName,OSLanguage) & JoinStr & LangPair
			CheckDataList = Split(Temp,JoinStr)
		End If
	End If

	'查找翻译引擎中对应的语言代码
	srcLngFind = 0
	trnLngFind = 0
	LangArray = Split(EngineDataList(2),SubJoinStr)
	EngineName = EngineDataList(0)
	CheckName = CheckDataList(0)
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

	'释放不再使用的动态数组所使用的内存
	Erase LangArray,LangPairList

	'根据翻译列表是否已打开设置要翻译的字串数
	If mSelOnly = 0 And TransListOpen = True Then
		StringCount = trn.StringCount
	Else
		StringCount = trn.StringCount(pslSelection)
	End If

	'开始处理每条字串
	PSL.OutputWnd.Clear
	PSL.Output Msg57 & vbCrLf & Msg11 & EngineName & Msg75 & Msg12 & CheckName
	StartTimes = Timer
	Set xmlHttp = CreateObject(DefaultObject)
	If xmlHttp Is Nothing Then Set xmlHttp = CreateObject("Msxml2.XMLHTTP")
	For j = 1 To StringCount
		'选择要翻译的字串
		If mSelOnly = 0 And TransListOpen = True Then
			Set TransString = trn.String(j)
		Else
			Set TransString = trn.String(j,pslSelection)
		End If

		'消息和字串初始化并获取翻译列表的现有来源和翻译字串
		SkipMsg = ""
		Massage = ""
		LineMsg = ""
		AcckeyMsg = ""
		ChangeMsg = ""
		orjSrcString = TransString.SourceText
		orjtrnString = TransString.Text

		'字串类型处理
		'If mAllType = 0 And mSelOnly = 0 Then
		'	If InStr("Menu|Dialog|StringTable|AcceleratorTable|Version",TransString.ResType) Then
		'		If InStr(StrTypes,TransString.ResType) = 0 Then GoTo Skip
		'	Else
		'		If mOther = 0 Then GoTo Skip
		'	End If
		'End If

		'跳过已锁定的字串
		If TransString.State(pslStateLocked) = True Then
			SkipMsg = Msg58 & Msg75 & Msg62
			SkipedCount = SkipedCount + 1
			GoTo Skip
		End If
		'跳过只读的字串
		If TransString.State(pslStateReadOnly) = True Then
			SkipMsg = Msg58 & Msg75 & Msg63
			SkipedCount = SkipedCount + 1
			GoTo Skip
		End If
		'跳过已翻译供复审的字串
		If mForReview = 1 And TransString.State(pslStateTranslated) = True Then
			If TransString.State(pslStateReview) = True Then
				SkipMsg = Msg58 & Msg75 & Msg64
				SkipedCount = SkipedCount + 1
				GoTo Skip
			End If
		End If
		'跳过已翻译并验证的字串
		If mValidated = 1 And TransString.State(pslStateTranslated) = True Then
			If TransString.State(pslStateReview) = False Then
				SkipMsg = Msg58 & Msg75 & Msg65
				SkipedCount = SkipedCount + 1
				GoTo Skip
			End If
		End If
		'跳过未翻译的字串
		If mNotTran = 1 And TransString.State(pslStateTranslated) = False Then
			SkipMsg = Msg58 & Msg75 & Msg66
			SkipedCount = SkipedCount + 1
			GoTo Skip
		End If
		'跳过为空或全为空格的字串
		If Trim(orjSrcString) = "" Then
			SkipMsg = Msg58 & Msg75 & Msg67
			SkipedCount = SkipedCount + 1
			GoTo Skip
		End If
		'跳过全为数字和符号的字串
		If mNumAndSymbol = 1 Then
			If CheckStr(orjSrcString,"0-64,91-96,123-191") = True Then
				SkipMsg = Msg58 & Msg75 & Msg68
				SkipedCount = SkipedCount + 1
				GoTo Skip
			End If
		End If
		'跳过全为大写英文的字串
		If mAllUCase = 1 Then
			If CheckStr(orjSrcString,"65-90") = True Then
				SkipMsg = Msg58 & Msg75 & Msg69
				SkipedCount = SkipedCount + 1
				GoTo Skip
			End If
		End If
		'跳过全为小写英文的字串
		If mAllLCase = 1 Then
			If CheckStr(orjSrcString,"97-122") = True Then
				SkipMsg = Msg58 & Msg75 & Msg70
				SkipedCount = SkipedCount + 1
				GoTo Skip
			End If
		End If

		'获取翻译源文字串
		srcString = TransString.SourceText

		'开始预处理并翻译字串
		If mPreStrRep = 1 Then srcString = ReplaceStr(srcString,1)
		If mSplitTrn = 0 Then
			If mAccKey = 1 Then srcString = AccessKeyHanding(srcString)
			If mAccelerator = 1 Then srcString = AcceleratorHanding(srcString)
			trnString = getTranslate(xmlHttp,srcString,LangPair,0)
		Else
			Temp = mAccKey & JoinStr & mAccelerator
			trnString = SplitTran(xmlHttp,srcString,LangPair,Temp,0)
		End If

		'开始后处理字串并替换原有翻译
		If trnString <> orjSrcString And trnString <> orjTrnString Then
			If mCheckTrn = 1 Then
				NewtrnString = CheckHanding(orjSrcString,trnString,TranLang)
			Else
				NewtrnString = trnString
			End If
			If mPreStrRep = 1 Then NewtrnString = ReplaceStr(NewtrnString,2)
			If mAppStrRep = 1 Then NewtrnString = ReplaceStr(NewtrnString,0)

			If NewtrnString <> orjTrnString Then
				TransString.Text = NewtrnString
				TransString.State(pslStateReview) = True
				If mTranComm = 1 Then
					TransString.TransComment = EngineName & " " & Msg49
				Else
					TransString.TransComment = ""
				End If
				TranedCount = TranedCount + 1
			End If
		Else
			NotChangeCount = NotChangeCount + 1
		End If

		'组织消息并输出
		If mShowMsg = 1 Then
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
		If mShowMsg = 1 And SkipMsg <> "" Then TransString.OutputError(SkipMsg)
	Next j
	Set xmlHttp = Nothing

	'翻译计数及消息输出
	ErrorCount = LineNumErrCount + accKeyNumErrCount
	PSL.Output TranMassage(TranedCount,SkipedCount,NotChangeCount,ErrorCount)
	If ErrorCount = 0 And TransListOpen = True Then trn.Close
	EndTimes = Timer
	PSL.Output Msg71 & Format(DateAdd("s",EndTimes - StartTimes,0),Msg72)

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


'获取在线翻译
Function getTranslate(xmlHttp As Object,srcStr As String,LngPair As String,fType As Integer) As String
    Dim UrlData As String,trnStr As String,LangFrom As String,LangTo As String
	Dim Temp As String,Pos As Integer,Code As String,srcStrBak As String

	SetsArray = Split(EngineDataList(1),SubJoinStr)
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
    		Pos = InStr(FindStr,",")
			If Pos <> 0 Then
				bstrHeader = Trim(Left(FindStr,Pos-1))
				bstrValue = Trim(Mid(FindStr,Pos+1))
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


'Utf-8 编码
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
	Next
	Utf8Encode = Szret
End Function


'ANSI 编码
Public Function ANSIEncode(textStr As String) As String
    Dim i As Long,startIndex As Long,endIndex As Long,x() As Byte
    x = StrConv(textStr,vbFromUnicode)
    startIndex = LBound(x)
    endIndex = UBound(x)
    For i = startIndex To endIndex
        ANSIEncode = ANSIEncode & "%" & Hex(x(i))
    Next i
End Function


'转换字符的编码格式
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


'解析 XML 格式对象并提取翻译文本
Function ReadXML(xmlObj As Object,IdNames As String,TagNames As String) As String
	Dim xmlDoc As Object,Node As Object,Item As Object,IdName As String,TagName As String
	Dim x As Integer,y As Integer,i As Integer,max As Integer
	If xmlObj Is Nothing Then Exit Function

	Set xmlDoc = CreateObject("Microsoft.XMLDOM")
	'xmlDoc.Async = False
	'xmlDoc.ValidateOnParse = False
	'xmlDoc.loadXML(xmlObj)	'加载字串
	xmlDoc.Load(xmlObj)		'加载对象
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


'提取指定前后字符之间的值
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


'检查字串是否仅包含数字和符号
Function CheckStr(textStr As String,AscRange As String) As Boolean
	Dim i As Integer,j As Integer,n As Integer,InpAsc As Long
	Dim Pos As Integer,Min As Long,max As Long
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
				max = CLng(Mid(AscValue(j),Pos+1))
			Else
				Min = CLng(AscValue(j))
				max = CLng(AscValue(j))
			End If
			If InpAsc >= Min And InpAsc <= max Then n = n + 1
		Next j
	Next i
	If n = Len(textStr) Then CheckStr = True
End Function


'分行翻译处理
Function SplitTran(xmlHttp As Object,srcStr As String,LangPair As String,Arg As String,fType As Integer) As String
	Dim i As Integer,srcStrBak As String,SplitStr As String
	Dim mAccKey As Integer,mAccelerator As Integer

	TempArray = Split(Arg,JoinStr,-1)
	mAccKey = CInt(TempArray(0))
	mAccelerator = CInt(TempArray(1))

	'用替换法拆分字串
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

	'获取每行的翻译
	For i = LBound(srcStrArr) To UBound(srcStrArr)
		srcString = srcStrArr(i)
		If srcString <> "" Then
			If mAccKey = 1 Then srcString = AccessKeyHanding(srcString)
			If mAccelerator = 1 Then srcString = AcceleratorHanding(srcString)
			trnString = getTranslate(xmlHttp,srcString,LangPair,fType)
		Else
			trnString = srcString
		End If
		If i > LBound(srcStrArr) Then SplitTran = SplitTran & SplitStr & trnString
		If i = LBound(srcStrArr) Then SplitTran = trnString
	Next i
End Function


'处理快捷键字符
Function AccessKeyHanding(srcStr As String) As String
	Dim i As Integer,j As Integer,n As Integer,posin As Integer
	Dim AccessKey As String,Stemp As Boolean

	srcStrBak = srcStr
	If InStr(srcStr,"&") = 0 Then
		AccessKeyHanding = srcStr
		Exit Function
	End If

	'获取选定配置的参数
	SetsArray = Split(CheckDataList(1),SubJoinStr)
	ExcludeChar = SetsArray(0)
	CheckBracket = SetsArray(2)

	'排除字串中的非快捷键
	If ExcludeChar <> "" Then
		FindStrArr = Split(Convert(ExcludeChar),",",-1)
		For i = LBound(FindStrArr) To UBound(FindStrArr)
			FindStr = LTrim(FindStrArr(i))
			If FindStr <> "" Then
				srcStrBak = Replace(srcStrBak,FindStr,"*a" & i & "!N!" & i & "d*")
			End If
		Next i
	End If

	'获取快捷键并去除
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

	'还原字串中被排除的非快捷键
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


'处理加速器字符
Function AcceleratorHanding(srcStr As String) As String
	Dim i As Integer,j As Integer,x As Integer,y As Integer
	Dim Shortcut As String,ShortcutKey As String,FindStr As String

	'获取选定配置的参数
	SetsArray = Split(CheckDataList(1),SubJoinStr)
	CheckShortChar = SetsArray(8)
	CheckShortKey = SetsArray(9)
	KeepShortKey = SetsArray(10)

	'获取加速器
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

	'去除加速器
	If Shortcut <> "" Then
		x = InStrRev(srcStr,Shortcut)
		If x <> 0 Then AcceleratorHanding = Left(srcStr,x-1)
	Else
		AcceleratorHanding = srcStr
	End If
End Function


'替换特定字符
Function ReplaceStr(trnStr As String,fType As Integer) As String
	Dim i As Integer,BaktrnStr As String
	'获取选定配置的参数
	SetsArray = Split(CheckDataList(1),SubJoinStr)
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
Function CheckHanding(srcStr As String,trnStr As String,TranLang As String) As String
	Dim i As Integer,BaksrcStr As String,BaktrnStr As String,srcStrBak As String,trnStrBak As String
	Dim srcNum As Integer,trnNum As Integer,srcSplitNum As Integer,trnSplitNum As Integer
	Dim FindStrArr() As String,srcStrArr() As String,trnStrArr() As String,LineSplitArr() As String
	Dim posinSrc As Integer,posinTrn As Integer

	'获取选定配置的参数
	SetsArray = Split(CheckDataList(1),SubJoinStr)
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
		BaktrnStr = StringReplace(BaksrcStr,BaktrnStr,TranLang)
	ElseIf srcSplitNum <> 0 Or trnSplitNum <> 0 Then
		LineSplitArr = MergeArray(srcStrArr,trnStrArr)
		BaktrnStr = ReplaceStrSplit(BaktrnStr,LineSplitArr,TranLang)
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
Function StringReplace(srcStr As String,trnStr As String,TranLang As String) As String
	Dim posinSrc As Integer,posinTrn As Integer,StringSrc As String,StringTrn As String
	Dim accesskeySrc As String,accesskeyTrn As String,Temp As String
	Dim ShortcutPosSrc As Integer,ShortcutPosTrn As Integer,PreTrn As String
	Dim EndStringPosSrc As Integer,EndStringPosTrn As Integer,AppTrn As String
	Dim preKeyTrn As String,appKeyTrn As String,Stemp As Boolean,FindStrArr() As String
	Dim i As Integer,j As Integer,x As Integer,y As Integer,m As Integer,n As Integer

	'获取选定配置的参数
	SetsArray = Split(CheckDataList(1),SubJoinStr)
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

	'数据集成
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
				If (acckeyTrn <> "" And acckeyTrn <> keySrc) Then
					Temp = Replace(Temp,"&","")
					Temp = Replace(Temp,keyTrn,keySrc,,1)
				Else
					ShortcutTrn = ShortcutSrc
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
	Dim Massage3 As String,Massage4 As String,n As Integer

	If OSLanguage = "0404" Then
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
	Else
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


'翻译消息输出
Function TranMassage(tCount As Integer,sCount As Integer,nCount As Integer,eCount As Integer) As String
	If OSLanguage = "0404" Then
		Msg01 = "陆亩ЧΘ⊿Τ陆亩ヴ﹃"
		Msg02 = "陆亩ЧΘㄤい"
		Msg03 = "陆亩 " & tCount & " 铬筁 " & sCount & " " & _
				"ゼ跑 " & nCount & " Τ岿粇 " & eCount & " "
	Else
		Msg01 = "翻译完成，没有翻译任何字串。"
		Msg02 = "翻译完成，其中："
		Msg03 = "已翻译 " & tCount & " 个，已跳过 " & sCount & " 个，" & _
				"未更改 " & nCount & " 个，有错误 " & eCount & " 个"
	End If
	TranCount = tCount + sCount + nCount
	If TranCount = 0 Then TranMassage = Msg01
	If TranCount <> 0 Then TranMassage = Msg02 & Msg03
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
	MsgBox(Msg02 & CStr(sysError.Number) & ", " & sysError.Description,vbInformation,Msg01)
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
Function ReplaceStrSplit(trnStr As String,StrSplitArr() As String,TranLang As String) As String
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
		trnStrSplitNew = StringReplace(srcStrSplit,trnStrSplit,TranLang)

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


'获取设置
Function EngineGet(SelSet As String,DataList() As String,Path As String) As Boolean
	Dim i As Integer,Header As String,HeaderIDs As String,HeaderIDArr() As String,Temp As String
	EngineGet = False
	NewVersion = ToUpdateEngineVersion
	NewSet = SelSet

	If Path = EngineRegKey Then GoTo GetFromRegistry
	If Path = "" Then Path = EngineFilePath
	If Dir(Path) = "" Then GoTo GetFromRegistry
	On Error GoTo GetFromRegistry
	Open Path For Input As #1
	While Not EOF(1)
		Line Input #1,L$
		If Left(Trim(L$),1) = "[" And Right(Trim(L$),1) = "]" Then
			Header = Mid(Trim(L$),2,Len(Trim(L$))-2)
		End If
		If L$ <> "" And Header <> "" Then
			setPreStr = ""
			setAppStr = ""
			If InStr(L$,"=") Then setPreStr = Trim(Left(L$,InStr(L$,"=")-1))
			If InStr(L$,"=") Then setAppStr = LTrim(Mid(L$,InStr(L$,"=")+1))
			'获取 Option 项和值
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
			'获取 Option 项外的全部项和值
			If Header <> "Option" And setPreStr <> "" Then
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
		If (L$ = "" Or EOF(1)) And (Header = NewSet Or (NewSet = "" And Header = EngineSet)) Then
			Temp = ObjName & UrlTmp & AppId & Url & Method & Async & User & Pwd & Body & _
					rHead & rType & bStr & aStr & LngPair
			If Temp <> "" Then
				Data = Header & JoinStr & ObjName & SubJoinStr & AppId & SubJoinStr & Url & _
						SubJoinStr & UrlTmp & SubJoinStr & Method & SubJoinStr & Async & SubJoinStr & _
						User & SubJoinStr & Pwd & SubJoinStr & Body & SubJoinStr & rHead & _
						SubJoinStr & rType & SubJoinStr & bStr & SubJoinStr & aStr & JoinStr & LngPair
				'更新旧版的默认配置值
				If InStr(Join(DefaultEngineList,JoinStr),Header) Then
					If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 Then
						Data = EngineDataUpdate(Header,Data)
					End If
				End If
				DataList = Split(Data,JoinStr)
				EngineGet = True
			End If
		End If
	Wend
	Close #1
	On Error GoTo 0
	Exit Function

	GetFromRegistry:
	'获取 Option 项和值
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
		KeepSet = GetSetting("WebTranslate","Option","KeepSetting",0)
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
	End If
	'获取 Option 外的项和值
	HeaderIDs = GetSetting("WebTranslate","Option","Headers","")
	If HeaderIDs <> "" Then
		HeaderIDArr = Split(HeaderIDs,";",-1)
		For i = LBound(HeaderIDArr) To UBound(HeaderIDArr)
			HeaderID = HeaderIDArr(i)
			If HeaderID <> "" Then
				'转存旧版的每个项和值
				Header = GetSetting("WebTranslate",HeaderID,"Name","")
				If Header = "" Then Header = HeaderID
				If Header <> "" And (Header = NewSet Or (NewSet = "" And Header = EngineSet)) Then
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
						Data = Header & JoinStr & ObjName & SubJoinStr & AppId & SubJoinStr & Url & _
								SubJoinStr & UrlTmp & SubJoinStr & Method & SubJoinStr & Async & SubJoinStr & _
								User & SubJoinStr & Pwd & SubJoinStr & Body & SubJoinStr & rHead & _
								SubJoinStr & rType & SubJoinStr & bStr & SubJoinStr & aStr & JoinStr & LngPair
						'更新旧版的默认配置值
						If InStr(Join(DefaultEngineList,JoinStr),Header) Then
							If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 Then
								Data = EngineDataUpdate(Header,Data)
							End If
						End If
						DataList = Split(Data,JoinStr)
						EngineGet = True
					End If
				End If
			End If
		Next i
	End If
End Function


'获取字串检查设置
Function CheckGet(SelSet As String,DataList() As String,Path As String) As Boolean
	Dim i As Integer,Header As String,HeaderIDs As String,HeaderIDArr() As String,Temp As String
	CheckGet = False
	NewVersion = ToUpdateCheckVersion
	NewSet = SelSet

	If Path = CheckRegKey Then GoTo GetFromRegistry
	If Path = "" Then Path = CheckFilePath
	If Dir(Path) = "" Then GoTo GetFromRegistry
	On Error GoTo GetFromRegistry
	Open Path For Input As #1
	While Not EOF(1)
		Line Input #1,L$
		If Left(Trim(L$),1) = "[" And Right(Trim(L$),1) = "]" Then
			Header = Mid(Trim(L$),2,Len(Trim(L$))-2)
		End If
		If L$ <> "" And Header <> "" Then
			setPreStr = ""
			setAppStr = ""
			If InStr(L$,"=") Then setPreStr = Trim(Left(L$,InStr(L$,"=")-1))
			If InStr(L$,"=") Then setAppStr = LTrim(Mid(L$,InStr(L$,"=")+1))
			'获取 Option 项和值
			If Header = "Option" And setPreStr <> "" Then
				If setPreStr = "Version" Then OldVersion = setAppStr
				If setPreStr = "AutoSelection" Then AutoSele = setAppStr
				If setPreStr = "AutoRepString" Then RepString = setAppStr
				If SelSet = "" Then
					If setPreStr = "AutoMacroSet" Then AutoMacroSet = setAppStr
					If AutoMacroSet = "Default" Then AutoMacroSet = DefaultCheckList(0)
				End If
			End If
			'获取 Option 项外的全部项和值
			If Header <> "Option" And setPreStr <> "" Then
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
		If (L$ = "" Or EOF(1)) And NewSet <> "" And AutoSele = "1" Then
			If getCheckID(LngPair,NewSet) = True Then NewSet = Header
		End If
		If (L$ = "" Or EOF(1)) And (Header = NewSet Or (NewSet = "" And Header = AutoMacroSet)) Then
			Temp = ExCr & LnSp & ChkBkt & KpPair & ChkEnd & NoTrnEnd & TrnEnd & _
					Short & Key & KpKey & PreStr & AppStr & LngPair
			If Temp <> "" Then
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
				DataList = Split(Data,JoinStr)
				CheckGet = True
			End If
		End If
	Wend
	Close #1
	On Error GoTo 0
	Exit Function

	GetFromRegistry:
	'获取 Option 项和值
	OldVersion = GetSetting("AccessKey","Option","Version","")
	AutoSele = GetSetting("AccessKey","Option","AutoSelection",0)
	RepString = GetSetting("AccessKey","Option","AutoRepString",0)
	If SelSet = "" Then
		AutoMacroSet = GetSetting("AccessKey","Option","AutoMacroSet","")
		If AutoMacroSet = "Default" Then AutoMacroSet = DefaultCheckList(0)
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
				If NewSet <> "" And AutoSele = "1" Then
					LngPair = GetSetting("AccessKey",HeaderID,"ApplyLangList","")
					If getCheckID(LngPair,NewSet) = True Then NewSet = Header
				End If
				If Header <> "" And (Header = NewSet Or (NewSet = "" And Header = AutoMacroSet)) Then
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
						DataList = Split(Data,JoinStr)
						CheckGet = True
					End If
				End If
			End If
		Next i
	End If
End Function


'更新引擎旧版本配置值
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
	EngineDataUpdate = TempArray(0) & JoinStr & UpdatedData & JoinStr &  TempArray(2)
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


'查找指定值是否在数组中
Function getCheckID(Data As String,LngCode As String) As Boolean
	Dim i As Integer,Stemp As Boolean
	getCheckID = False
	If LngCode = "" Or Data = "" Then Exit Function
	LangArray = Split(Data,SubJoinStr)
	For i = LBound(LangArray) To UBound(LangArray)
		LangPairList = Split(LangArray(i),LngJoinStr)
		If LCase(LangPairList(1)) = LCase(LngCode) Then
			getCheckID = True
			Exit For
		End If
	Next i
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
