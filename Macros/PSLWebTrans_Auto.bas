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
''Idea and implemented by wanfu 2010.05.12 (Last modified on 2012.04.10)

Public Type STRING_INFO
	PreSpace			As String	'字符串前置空格
	EndSpace			As String	'字符串后置空格
	Spaces				As String	'字符串快捷键前空格
	AccKey				As String	'字符串快捷键
	AccKeyIFR			As String	'字符串快捷键标志符
	AccKeyKey			As String	'字符串快捷键字符
	EndString			As String	'字符串终止符
	Shortcut			As String	'字符串加速器
	PreString			As String	'字符串快捷键前的字符（不含快捷键前的空格和终止符）
	ExpString			As String	'字符串快捷键至加速器前的字符
	AccKeyPos			As Integer	'字符串快捷键位置
	AccKeyNum			As Integer	'字符串快捷键数
	Length				As Integer	'字符串长度
	LineNum				As Integer	'字符串行数
End Type

Public PslLangDataList() As String
Public UIFileList() As String,UIDataList() As String,UILangList() As String,LangFile As String

Public StringSrc As STRING_INFO,StringTrn As STRING_INFO,MoveAcckey As String
Public AllCont As Long,AccKey As Long,EndChar As Long,Acceler As Long

Public DefaultCheckList() As String,AppRepStr As String,PreRepStr As String
Public CheckDataList() As String,CheckDataListBak() As String
Public DefaultProjectList() As String,ProjectDataList() As String
Public DefaultEngineList() As String,EngineDataList() As String,tSelected() As String,WaitTimes As Long

Private Const Version = "2012.02.19"
Private Const ToUpdateEngineVersion = "2012.02.13"
Private Const ToUpdateCheckVersion = "2011.11.14"
Private Const EngineRegKey = "HKCU\Software\VB and VBA Program Settings\WebTranslate\"
Private Const EngineFilePath = MacroDir & "\Data\PSLWebTrans.dat"
Private Const CheckRegKey = "HKCU\Software\VB and VBA Program Settings\AccessKey\"
Private Const CheckFilePath = MacroDir & "\Data\PSLCheckAccessKeys.dat"
Private Const JoinStr = vbFormFeed  'vbBack
Private Const SubJoinStr = vbVerticalTab  'Chr$(1)
Private Const LngJoinStr = "|"
Private Const SubLngJoinStr = Chr$(1)
Private Const NullValue = "Null"
Private Const DefaultObject = "Microsoft.XMLHTTP;Msxml2.XMLHTTP"
Private Const updateAppName = "PSLWebTrans"
Private Const DefaultWaitTimes = 10    '10秒


'翻译引擎默认设置
Function EngineSettings(DataName As String) As String
	Dim StesArray(19) As String
	If DataName = DefaultEngineList(0) Then
		StesArray(0) = DefaultObject
		StesArray(1) = "fefed727-bbc1-4421-828d-fc828b24d59b"
		StesArray(2) = "http://api.microsofttranslator.com/V2/Http.svc/Translate?"
		StesArray(3) = "{Url}&appId={appId}&text={text}&from={from}&to={to}"
		StesArray(4) = "GET"
		StesArray(5) = "False"
		StesArray(6) = ""
		StesArray(7) = ""
		StesArray(8) = ""
		StesArray(9) = "Content-Type,application/xml; charset=utf-8"
		StesArray(10) = "responseBody"
		StesArray(11) = "Serialization/"">"
		StesArray(12) = "</string>"
		StesArray(13) = "Serialization/"">"
		StesArray(14) = "</string>"
		StesArray(15) = "Serialization/"">"
		StesArray(16) = "</string>"
		StesArray(17) = "string"
		StesArray(18) = "string"
		StesArray(19) = "1"
	ElseIf DataName = DefaultEngineList(1) Then
		StesArray(0) = DefaultObject
		StesArray(1) = ""
		StesArray(2) = "http://translate.google.com/translate_t?"
		StesArray(3) = "{Url}&text={text}&langpair={from}|{to}"
		StesArray(4) = "POST"
		StesArray(5) = "False"
		StesArray(6) = ""
		StesArray(7) = ""
		StesArray(8) = ""
		StesArray(9) = "Content-Type,text/html; charset=utf-8"
		StesArray(10) = "responseBody"
		StesArray(11) = "onmouseout=""this.style.backgroundColor='#fff'"">"
		StesArray(12) = "</span>"
		StesArray(13) = "onmouseout=""this.style.backgroundColor='#fff'"">"
		StesArray(14) = "</span>"
		StesArray(15) = "onmouseout=""this.style.backgroundColor='#fff'"">"
		StesArray(16) = "</span>"
		StesArray(17) = ""
		StesArray(18) = ""
		StesArray(19) = "1"
	ElseIf DataName = DefaultEngineList(2) Then
		StesArray(0) = DefaultObject
		StesArray(1) = ""
		StesArray(2) = "http://fanyi.yahoo.com.cn/translate_txt?"
		StesArray(3) = "{Url}&ei=UTF-8&fr=&lp={from}_{to}&trtext={Text}"
		StesArray(4) = "POST"
		StesArray(5) = "False"
		StesArray(6) = ""
		StesArray(7) = ""
		StesArray(8) = ""
		StesArray(9) = "Content-Type,text/html; charset=utf-8"
		StesArray(10) = "responseBody"
		StesArray(11) = "<div id=""pd"" class=""pd"">"
		StesArray(12) = "</div>"
		StesArray(13) = "<div id=""pd"" class=""pd"">"
		StesArray(14) = "</div>"
		StesArray(15) = "<div id=""pd"" class=""pd"">"
		StesArray(16) = "</div>"
		StesArray(17) = ""
		StesArray(18) = ""
		StesArray(19) = "0"
	End If
	EngineSettings = Join(StesArray,SubJoinStr)
End Function


'字串处理默认设置
Function CheckSettings(DataName As String,DataType As Long) As String
	Dim i As Long,n As Long,j As Long,Max As Long,CheckName As String
	Dim TempList() As String,readByte() As Byte,DefaultCheckDataList() As String

	If DataType = 0 Then ReDim TempList(17) As String
	If DataType = 1 Then ReDim TempList(20) As String

	If DataName <> "" And DataType = 0 Then
		If DataName = DefaultCheckList(0) Then CheckName = "en2zh"
		If DataName = DefaultCheckList(1) Then CheckName = "zh2en"
	ElseIf DataName <> "" And DataType = 1 Then
		If DataName = DefaultProjectList(0) Then CheckName = "CheckOnly"
		If DataName = DefaultProjectList(1) Then CheckName = "CheckAndCorrect"
		If DataName = DefaultProjectList(2) Then CheckName = "DelAccessKey"
		If DataName = DefaultProjectList(3) Then CheckName = "DelAccelerator"
		If DataName = DefaultProjectList(4) Then CheckName = "DelAccessKeyAndAccelerator"
	End If
	If CheckName = "" Then GoTo ExitFunction

	ConfigFile = MacroDir & "\Data\PSLCheckAccessKeys.ini"
	On Error GoTo ErrMassage
	If Dir(ConfigFile) = "" Then Err.Raise(1,"NotExitFile",LangFile)

	On Error GoTo NotReadFile
	i = FileLen(ConfigFile)
	ReDim readByte(i) As Byte
	FN = FreeFile
	Open ConfigFile For Binary As #FN
	Get #FN,,readByte
	Close #FN
	DefaultCheckDataList = Split(readByte,vbCrLf)
	Erase readByte
	If Join(DefaultCheckDataList,"") = "" Then GoTo ExitFunction

	On Error GoTo ErrMassage
	n = 0
	Max = UBound(DefaultCheckDataList)
	For i = 0 To Max
		l$ = DefaultCheckDataList(i)
		If Trim(l$) <> "" Then
			If Left(Trim(l$),1) = "[" And Right(Trim(l$),1) = "]" Then
				Header$ = Trim(Mid(Trim(l$),2,Len(Trim(l$))-2))
			End If
			If Header$ <> "" And HeaderBak$ = "" Then HeaderBak$ = Header$
			If Header$ <> "" And Header$ = HeaderBak$ Then
				setPreStr$ = ""
				setAppStr$ = ""
				j = InStr(l$,"=")
				If j > 0 Then
					setPreStr$ = Trim(Left(l$,j - 1))
					setAppStr$ = LTrim(Mid(l$,j + 1))
				End If
				If Header$ = "Option" And setPreStr$ <> "" Then
					If setPreStr$ = "Version" Then
						UpdateVersion = setAppStr$
						If UpdateVersion < ToUpdateCheckVersion Or UpdateVersion > Version Then
							CheckName = ""
							Err.Raise(1,"NotVersion",ConfigFile & JoinStr & UpdateVersion & _
										JoinStr & ToUpdateCheckVersion)
							Exit For
						End If
					End If
				End If
				If DataType = 0 And Header$ = CheckName And setPreStr$ <> "" Then
					If setPreStr$ = "ExcludeChar" Then TempList(0) = setAppStr$
					If setPreStr$ = "LineSplitChar" Then TempList(1) = setAppStr$
					If setPreStr$ = "CheckBracket" Then TempList(2) = setAppStr$
					If setPreStr$ = "KeepCharPair" Then TempList(3) = setAppStr$
					If setPreStr$ = "ShowAsiaKey" Then TempList(4) = setAppStr$
					If setPreStr$ = "CheckEndChar" Then TempList(5) = setAppStr$
					If setPreStr$ = "NoTrnEndChar" Then TempList(6) = setAppStr$
					If setPreStr$ = "AutoTrnEndChar" Then TempList(7) = setAppStr$
					If setPreStr$ = "CheckShortChar" Then TempList(8) = setAppStr$
					If setPreStr$ = "CheckShortKey" Then TempList(9) = setAppStr$
					If setPreStr$ = "KeepShortKey" Then TempList(10) = setAppStr$
					If setPreStr$ = "PreRepString" Then TempList(11) = setAppStr$
					If setPreStr$ = "AutoRepString" Then TempList(12) = setAppStr$
					If setPreStr$ = "AccessKeyChar" Then TempList(13) = setAppStr$
					If setPreStr$ = "AddAccessKeyWithFirstChar" Then TempList(14) = setAppStr$
					If setPreStr$ = "LineSplitMode" Then TempList(15) = setAppStr$
					If setPreStr$ = "AppInsertSplitChar" Then TempList(16) = setAppStr$
					If setPreStr$ = "ReplaceSplitChar" Then TempList(17) = setAppStr$
				ElseIf DataType = 1 And Header$ = "Projects" And setPreStr$ <> "" Then
					If setPreStr$ = CheckName Then TempList = Split(setAppStr$,LngJoinStr)
				End If
			End If
		End If
		If Header$ <> "" And (i = Max Or Header$ <> HeaderBak$) Then
			If Join(TempList,"") <> "" Then
				If DataType = 0 And HeaderBak$ = CheckName Then
					CheckSettings = Join(TempList,SubJoinStr)
				ElseIf DataType = 1 And HeaderBak$ = "Projects" Then
					CheckSettings = Join(TempList,LngJoinStr)
				End If
				n = n + 1
				Exit For
			End If
			HeaderBak$ = Header$
		End If
	Next i

	If n = 0 And CheckName <> "" Then
		If DataType = 0 Then Temp = "NotSection"
		If DataType = 1 Then Temp = "NotValue"
		Err.Raise(1,Temp,ConfigFile & JoinStr & CheckName)
	End If
	Exit Function

	NotReadFile:
	Err.Source = "NotReadFile"
	Err.Description = Err.Description & JoinStr & ConfigFile

	ErrMassage:
	Call sysErrorMassage(Err,1)

	ExitFunction:
	If n = 0 Then
		If DataType = 0 Then CheckSettings = Join(TempList,SubJoinStr)
		If DataType = 1 Then CheckSettings = Join(TempList,LngJoinStr)
	End If
End Function


' 主程序
Public Sub PSL_OnAutoTranslate(Translations As PslTranslations,ByVal MinMatch As Long,ByVal MaxCount As Long)
	Dim TransString As PslTransString,CheckID As Long,EngineID As Long
	Dim i As Long,j As Long,n As Long,srcString As String,trnString As String,TranLang As String
	Dim LangPairList() As String,xmlHttp As Object,objStream As Object
	Dim LangPair As String,Temp As String,TempList() As String,TempArray() As String
	Dim srcLng As String,trnLng As String,srcLngFind As Long,trnLngFind As Long
	Dim strKeyPath As String,WshShell As Object,MsgList() As String,Stemp As Boolean
	Dim mCheckSrc As Long,iVoSrc As Long,mCheckTrn As Long,iVoTrn As Long
	Dim ShowOriginalTran As Long,ApplyCheckResult As Long,k As Long

	'字串初始化并获取翻译列表的现有来源和翻译字串
	srcString = Translations.SourceString

	'跳过为空或全为空格的字串
	If Trim(srcString) = "" Then GoTo Skip

	'检测系统语言
	On Error Resume Next
	Set WshShell = CreateObject("WScript.Shell")
	If WshShell Is Nothing Then
		PSL.Output(Err.Description & " - " & "WScript.Shell")
		Exit Sub
	End If
	strKeyPath = "HKLM\SYSTEM\CurrentControlSet\Control\Nls\Language\Default"
	OSLanguage = WshShell.RegRead(strKeyPath)
	If OSLanguage = "" Then
		strKeyPath = "HKLM\SYSTEM\CurrentControlSet\Control\Nls\Language\InstallLanguage"
		OSLanguage = WshShell.RegRead(strKeyPath)
		If Err.Source = "WshShell.RegRead" Then
			PSL.Output(Err.Description)
			Exit Sub
		End If
	End If
	Set WshShell = Nothing

	'检测 Adodb.Stream 是否存在
	Set objStream = CreateObject("Adodb.Stream")
	If objStream Is Nothing Then
		PSL.Output(Err.Description & " - " & "Adodb.Stream")
		Exit Sub
	End If
	Set objStream = Nothing
	On Error GoTo SysErrorMsg

	'初始化数组
	ReDim UIFileList(0),UIDataList(0),UILangList(0)
	ReDim DefaultEngineList(2),EngineDataList(0),tSelected(27),TempArray(0)
	ReDim DefaultCheckList(1),CheckDataList(0)
	ReDim DefaultProjectList(4),ProjectDataList(0)
	UIFileList(0) = "Auto"
	UIDataList(0) = "Auto" & JoinStr & "0" & JoinStr

	'读取翻译引擎设置
	DefaultEngineList(0) = "Microsoft"
	DefaultEngineList(1) = "Google"
	DefaultEngineList(2) = "Yahoo"
	If EngineGet("",TempArray,EngineDataList,"") <> 4 Then
		For i = LBound(DefaultEngineList) To UBound(DefaultEngineList)
			EngineName = DefaultEngineList(i)
			Stemp = False
			For j = LBound(TempArray) To UBound(TempArray)
				If TempArray(j) = EngineName Then
					Stemp = True
					Exit For
				End If
			Next j
			If Stemp = False Then
				LangPairs = Join(LangCodeList(EngineName,0,108),SubLngJoinStr)
				Temp = EngineName & JoinStr & EngineSettings(EngineName) & JoinStr & LangPairs
				CreateArray(EngineName,Temp,TempArray,EngineDataList)
			End If
		Next i
	End If

	'转换引擎设置
	If Join(tSelected,"") <> "" Then
		'EngineName = tSelected(0)
		CheckName = tSelected(1)
		'mAllType = StrToLong(tSelected(2))
		'mMenu = StrToLong(tSelected(3))
		'mDialog = StrToLong(tSelected(4))
		'mString = StrToLong(tSelected(5))
		'mAccTable = StrToLong(tSelected(6))
		'mVer = StrToLong(tSelected(7))
		'mOther = StrToLong(tSelected(8))
		mSelOnly = StrToLong(tSelected(9))
		'mForReview = StrToLong(tSelected(10))
		'mValidated = StrToLong(tSelected(11))
		'mNotTran = StrToLong(tSelected(12))
		mNumAndSymbol = StrToLong(tSelected(13))
		mAllUCase = StrToLong(tSelected(14))
		mAllLCase = StrToLong(tSelected(15))
		mAutoSele = StrToLong(tSelected(16))
		iVoSrc = StrToLong(tSelected(17))
		mCheckSrc = StrToLong(tSelected(18))
		mPreStrRep = StrToLong(tSelected(19))
		mSplitTrn = StrToLong(tSelected(20))
		iVoTrn = StrToLong(tSelected(21))
		mCheckTrn = StrToLong(tSelected(22))
		mAppStrRep = StrToLong(tSelected(23))
		'KeepSet = StrToLong(tSelected(24))
		'mShowMsg = StrToLong(tSelected(25))
		mTranComm = StrToLong(tSelected(26))
	End If

	'获取字串类型组合
	'If mMenu = 1 Then StrTypes = "|Menu|"
	'If mDialog = 1 Then StrTypes = StrTypes & "|Dialog|"
	'If mString = 1 Then StrTypes = StrTypes & "|StringTable|"
	'If mAccTable = 1 Then StrTypes = StrTypes & "|AcceleratorTable|"
	'If mVer = 1 Then StrTypes = StrTypes & "|Version|"

	'字串类型处理
	'If mAllType = 0 And mSelOnly = 0 Then
	'	If InStr("Menu|Dialog|StringTable|AcceleratorTable|Version",TransString.ResType) Then
	'		If InStr(StrTypes,TransString.ResType) = 0 Then GoTo Skip
	'	Else
	'		If mOther = 0 Then GoTo Skip
	'	End If
	'End If

	'跳过全为数字和符号的字串
	If mNumAndSymbol = 1 Then
		If LCase(srcString) = UCase(srcString) Then
			If CheckStr(srcString,"0-64,91-96,123-191",1) = True Then GoTo Skip
		End If
	End If
	'跳过全为大写英文的字串
	If mAllUCase = 1 Then
		If UCase(srcString) = srcString Then
			If CheckStr(srcString,"0-96,123-191",1) = True Then GoTo Skip
		End If
	End If
	'跳过全为小写英文的字串
	If mAllLCase = 1 Then
		If LCase(srcString) = srcString Then
			If CheckStr(srcString,"0-64,91-191",1) = True Then GoTo Skip
		End If
	End If

	'读取界面语言字串
	If GetUIList(UIFileList,UIDataList) = True Then
		If Join(tSelected,"") <> "" Then UILangID = LCase(tSelected(27))
		If UILangID = "" Or UILangID = "0" Then UILangID = LCase(OSLanguage)
		TempList = Split(UILangID,";")
		For i = 1 To UBound(UIDataList)
			TempArray = Split(UIDataList(i),JoinStr)
			Temp = LCase(TempArray(1))
			File = TempArray(2)
			If Temp = UILangID Then
				LangFile = MacroDir & "\Data\" & File
				Exit For
			End If
			TempArray = Split(Temp,";")
			For j = 0 To UBound(TempList)
				For n = 0 To UBound(TempArray)
					If TempList(j) = TempArray(n) Then
						LangFile = MacroDir & "\Data\" & File
						Exit For
					End If
				Next n
				If LangFile <> "" Then Exit For
			Next j
			If LangFile <> "" Then Exit For
		Next i
	End If
	If LangFile = "" Then LangFile = MacroDir & "\Data\" & updateAppName & "_" & OSLanguage & ".lng"
	If Dir(LangFile) = "" Then
		LangFile = MacroDir & "\Data\" & updateAppName & "_" & OSLanguage & ".lng"
		If Dir(LangFile) = "" Then
			LangFile = ""
			For i = 0 To 2
				If i = 0 Then Temp = MacroDir & "\Data\" & updateAppName & "_0804.lng"
				If i = 1 Then Temp = MacroDir & "\Data\" & updateAppName & "_0404.lng"
				If i = 2 And UBound(UIDataList) > 0 Then
					TempArray = Split(UIDataList(1),JoinStr)
					Temp = MacroDir & "\Data\" & TempArray(2)
				End If
				If Dir(Temp) <> "" Then
					LangFile = Temp
					Exit For
				End If
			Next i
			If LangFile = "" Then Err.Raise(1,"NotExitFile",MacroDir & "\Data\" & updateAppName & "_*.lng")
		End If
	End If
	If getUILangList(LangFile,UILangList) = False Then Exit Sub
	If getMsgList(UILangList,MsgList,"Main",0) = False Then Exit Sub

	'检测 PSL 版本
	If PSL.Version < 600 Then
		PSL.Output(MsgList(41) & ":" & MsgList(43))
		Exit Sub
	End If

	'获取PSL的来源语言代码
	srcLng = PSL.GetLangCode(Translations.SourceLangID,pslCode639_1)
	If srcLng = "" Then srcLng = PSL.GetLangCode(Translations.SourceLangID,pslCodeLangRgn)
	If srcLng = "zh" Then
		srcLng = PSL.GetLangCode(Translations.SourceLangID,pslCodeLangRgn)
		If srcLng = "zh-CHS" Or srcLng = "zh-SG" Then srcLng = "zh-CN"
		If srcLng = "zh-CHT" Or srcLng = "zh-HK" Or srcLng = "zh-MO" Then srcLng = "zh-TW"
	End If

	'获取PSL的目标语言代码
	trnLng = PSL.GetLangCode(Translations.TargetLangID,pslCode639_1)
	If trnLng = "" Then trnLng = PSL.GetLangCode(Translations.TargetLangID,pslCodeLangRgn)
	If trnLng = "zh" Or trnLng = "ja" Or trnLng = "ko" Then TranLang = "Asia"
	If trnLng = "zh" Then
		trnLng = PSL.GetLangCode(Translations.TargetLangID,pslCodeLangRgn)
		If trnLng = "zh-CHS" Or trnLng = "zh-SG" Then trnLng = "zh-CN"
		If trnLng = "zh-CHT" Or trnLng = "zh-HK" Or trnLng = "zh-MO" Then trnLng = "zh-TW"
	End If

	'读取字串处理设置
	DefaultCheckList(0) = MsgList(70)
	DefaultCheckList(1) = MsgList(71)
	DefaultProjectList(0) = MsgList(89)
	DefaultProjectList(1) = MsgList(90)
	DefaultProjectList(2) = MsgList(91)
	DefaultProjectList(3) = MsgList(92)
	DefaultProjectList(4) = MsgList(93)
	If mCheckSrc = 1 Or mCheckTrn = 1 Or mPreStrRep = 1 Or mAppStrRep = 1 Then
		j = 0
		If mAutoSele = 1 Then j = CheckGet("",CheckDataList,"",trnLng)
		If j <> 4 Then j = CheckGet("",CheckDataList,"","")
		If j <> 4 Then
			If TranLang = "Asia" Then CheckName = DefaultCheckList(0)
			If TranLang <> "Asia" Then CheckName = DefaultCheckList(1)
			Temp = CheckSettings(CheckName,0)
			If Trim(Replace(Temp,SubJoinStr,"")) <> "" Then
				CheckDataList(0) = CheckName & JoinStr & Temp
			Else
				CheckName = MsgList(88)
				CheckDataList(0) = CheckName & JoinStr & Temp
				mAutoSele = 0
				iVoSrc = 0
				mCheckSrc = 0
				mPreStrRep = 0
				mSplitTrn = 0
				iVoTrn = 0
				mCheckTrn = 0
				mAppStrRep = 0
			End If
		End	If
		If Join(ProjectDataList,"") = "" Then
			TempArray = ProjectDataList
			For i = LBound(DefaultProjectList) To UBound(DefaultProjectList)
				ProjectName = DefaultProjectList(i)
				Temp = CheckSettings(ProjectName,1)
				If Trim(Replace(Temp,LngJoinStr,"")) <> "" Then
					Temp = ProjectName & JoinStr & Temp
					CreateArray(ProjectName,Temp,TempArray,ProjectDataList)
				End If
			Next i
			If Join(ProjectDataList,"") = "" Then
				iVoSrc = 0
				mCheckSrc = 0
				iVoTrn = 0
				mCheckTrn = 0
			End If
		End If
	End If
	AllCont = 1
	AccKey = 0
	EndChar = 0
	Acceler = 0
	CheckID = 0

	'更改字串检查配置名称
	If tSelected(1) = "en2zh" Then tSelected(1) = DefaultCheckList(0)
	If tSelected(1) = "zh2en" Then tSelected(1) = DefaultCheckList(1)

	'排序检查配置值并转换配置中的转义符
	If Join(CheckDataList,"") <> "" Then
		CheckDataListBak = CheckDataList
		TempArray = Split(CheckDataList(CheckID),JoinStr)
		TempDataList = Split(TempArray(1),SubJoinStr)
		For i = 0 To UBound(TempDataList)
			If i <> 4 And i <> 14 And i <> 15 And i < 18 Then
				If TempDataList(i) <> "" Then
					If i = 1 Or i = 5 Or i = 13 Or i = 16 Or i = 17 Then
						If i = 5 Or i = 7 Then Temp = " " Else Temp = ","
						TempList = SortArray(Split(TempDataList(i),Temp,-1),0,"Lenght","<")
						TempDataList(i) = Convert(Join(TempList,Temp))
					Else
						TempDataList(i) = Convert(TempDataList(i))
					End If
				End If
			End If
		Next i
		TempArray(1) = Join(TempDataList,SubJoinStr)
		CheckDataList(CheckID) = Join(TempArray,JoinStr)
		CheckName = TempArray(0)
	Else
		CheckName = MsgList(88)
	End If

	'获取检查方案设置
	If mCheckSrc = 1 Then
		TempArray = Split(ProjectDataList(iVoSrc),JoinStr)
		SrcProjectName = TempArray(0)
	Else
		SrcProjectName = MsgList(86)
	End If
	If mCheckTrn = 1 Then
		TempArray = Split(ProjectDataList(iVoTrn),JoinStr)
		TempDataList = Split(TempArray(1),LngJoinStr)
		ShowOriginalTran = StrToLong(TempDataList(19))
		ApplyCheckResult = StrToLong(TempDataList(20))
		TrnProjectName = TempArray(0)
	Else
		TrnProjectName = MsgList(86)
	End If

	'释放不再使用的动态数组所使用的内存
	Erase TempArray,TempList,TempDataList
	Erase UIFileList,UIDataList

	'开始预处理字串
	OldSrcString = srcString
	If mPreStrRep = 1 Then srcString = ReplaceStr(CheckID,srcString,0,0)
	If mSplitTrn = 0 Then
		If mCheckSrc = 1 Then srcString = CheckHanding(CheckID,OldSrcString,srcString,iVoSrc)
		If InStr(srcString,"&") Then srcString = Replace(srcString,"&","")
	End If

	'分别用启用的翻译引擎翻译字串并进行后处理
	k = 0
	MaxCount = MaxCount + UBound(EngineDataList) + 1
	For j = 0 To UBound(EngineDataList)
		'查找翻译引擎中对应的语言代码
		TempArray = Split(EngineDataList(j),JoinStr)
		TempDataList = Split(TempArray(1),SubJoinStr)
		LangArray = Split(TempArray(2),SubLngJoinStr)
		EngineName = TempArray(0)
		EngineID = j
		srcLngFind = 0
		trnLngFind = 0
		LangPair = ""
		trnString = ""
		If TempDataList(19) <> "1" Then GoTo NextNum

		'检测 Microsoft.XMLHTTP 是否存在
		On Error Resume Next
		TempList = Split(TempDataList(0),IIf(InStr(TempDataList(0),";"),";",","))
		For i = 0 To UBound(TempList)
			Temp = Trim(TempList(i))
			If Temp <> "" Then
				Set xmlHttp = CreateObject(Temp)
				If Not xmlHttp Is Nothing Then Exit For
			End If
		Next i
		On Error GoTo SysErrorMsg
		If xmlHttp Is Nothing Then GoTo NextNum

		'查找翻译引擎中对应的语言代码
		For i = 0 To UBound(LangArray)
			LangPairList = Split(LangArray(i),LngJoinStr)
			If srcLngFind = 0 Then
				If LCase(srcLng) = LCase(LangPairList(1)) Then
					srcLngCode = LangPairList(2)
					srcLngFind = 1
				End If
			End If
			If trnLngFind = 0 Then
				If LCase(trnLng) = LCase(LangPairList(1)) Then
					trnLngCode = LangPairList(2)
					trnLngFind = 1
				End If
			End If
			If srcLngFind + trnLngFind = 2 Then Exit For
		Next i
		If srcLngFind + trnLngFind < 2 Then GoTo NextNum
		LangPair = srcLngCode & LngJoinStr & trnLngCode

		'转换翻译引擎的配置中的转义符
		For i = 0 To UBound(TempDataList)
			If i > 10 And i < 19 Then
				If TempDataList(i) <> "" Then
					TempDataList(i) = Convert(TempDataList(i))
				End If
			End If
		Next i
		TempArray(1) = Join(TempDataList,SubJoinStr)
		EngineDataList(EngineID) = Join(TempArray,JoinStr)

		'释放不再使用的动态数组所使用的内存
		Erase TempArray,TempDataList,LangArray,LangPairList

		'获取测试翻译
		'Temp = getTranslate(EngineID,xmlHttp,"Testing at " & Time,LangPair,3)
		'测试 Internet 连接
		'If Temp = "NotConnected" Then Exit Sub
		'测试引擎网址是否为空
		'If Temp = "NullUrl" Then GoTo NextNum
		'测试引擎引擎是否超时
		'If Temp = "Timeout" Then GoTo NextNum
		'测试引擎结果是否为空
		'If Trim(Temp) = "" Then GoTo NextNum

		'开始翻译字串
		If mSplitTrn = 0 Then
			trnString = getTranslate(EngineID,xmlHttp,srcString,LangPair,0)
		Else
			Temp = EngineID & JoinStr & CheckID & JoinStr & iVoSrc & JoinStr & mCheckSrc & JoinStr & k
			trnString = SplitTran(xmlHttp,srcString,LangPair,Temp,0)
		End If

		'开始后处理字串并替换原有翻译
		If Trim(trnString) <> "" And trnString <> OldSrcString Then
			If mCheckTrn = 1 Then
				CheckTrnString = CheckHanding(CheckID,OldSrcString,trnString,iVoTrn)
				If ApplyCheckResult = 1 Then trnString = CheckTrnString
			End If
			If mAppStrRep = 1 Then trnString = ReplaceStr(CheckID,trnString,2,1)
			Translations.Add(trnString,OldSrcString,100,EngineName & " " & MsgList(4))
		End If
		NextNum:
		If Translations.Count = MaxCount Then Exit For
		k = 1
	Next j
	Skip:
	Set xmlHttp = Nothing
	Exit Sub

	'显示程序错误消息
	SysErrorMsg:
	If Err.Source <> "ExitSub" Then Call sysErrorMassage(Err,0)
End Sub


'获取在线翻译
Function getTranslate(ID As Long,xmlHttp As Object,srcStr As String,LngPair As String,fType As Long) As String
	Dim trnStr As String,srcStrBak As String,LangFrom As String,LangTo As String,StatusValue As Long
	Dim i As Long,Pos As Long,Code As String,TempList() As String,Temp As String

	If Trim(srcStr) = "" Or LngPair = "" Then Exit Function
	TempArray = Split(EngineDataList(ID),JoinStr)
	SetsArray = Split(TempArray(1),SubJoinStr)
	AppId = SetsArray(1)
	Url = SetsArray(2)
	UrlTemplate = SetsArray(3)
	Method = SetsArray(4)
	Async = SetsArray(5)
	User = SetsArray(6)
	Password = SetsArray(7)
	BodyData = SetsArray(8)
	RequestHeader = SetsArray(9)
	responseType = LCase(SetsArray(10))
	If responseType = "responsetext" Then
		TranBeforeStr = SetsArray(11)
		TranAfterStr = SetsArray(12)
	ElseIf responseType = "responsebody" Then
		TranBeforeStr = SetsArray(13)
		TranAfterStr = SetsArray(14)
	ElseIf responseType = "responsestream" Then
		TranBeforeStr = SetsArray(15)
		TranAfterStr = SetsArray(16)
	ElseIf responseType = "responsexml" Then
		TranBeforeStr = SetsArray(17)
		TranAfterStr = SetsArray(18)
	End If

	If Url = "" Then
		If fType = 3 Then getTranslate = "NullUrl"
		Exit Function
	End If

	On Error GoTo ErrorHandler
	If fType = 2 Then
		On Error GoTo ErrorHandler
		xmlHttp.Open Method,Url,Async,User,Password
		'xmlHttp.setRequestHeader("If-Modified-Since","0")
		xmlHttp.send()
		If OnReadyStateChange(xmlHttp,4,WaitTimes) = 4 Then
			getTranslate = xmlHttp.getAllResponseHeaders
		End If
		xmlHttp.Abort
		Exit Function
	End If

	If LngPair <> "" Then
		TempList = Split(LngPair,LngJoinStr)
		LangFrom = TempList(0)
		LangTo = TempList(1)
	End If

	srcStrBak = srcStr
	Pos = InStr(LCase(RequestHeader),"charset")
	If Pos > 0 Then
		Temp = Mid(RequestHeader,Pos)
		TempList = Split(Temp,vbCrLf)
		For i = 0 To UBound(TempList)
			Temp = TempList(i)
			If InStr(Temp,"=") Then
				Code = ExtractStr(Temp,"=",";|" & vbCrLf,1)
				If Code <> "" Then Exit For
			End If
		Next i
	Else
		xmlHttp.Open Method,Url,Async,User,Password
		'xmlHttp.setRequestHeader("If-Modified-Since","0")
		xmlHttp.send()
		If OnReadyStateChange(xmlHttp,4,WaitTimes) = 4 Then
			Temp = xmlHttp.getResponseHeader("Content-Type")
			Pos = InStr(LCase(Temp),"charset")
			If Pos > 0 Then
				Temp = Mid(Temp,Pos)
				If InStr(Temp,"=") Then Code = ExtractStr(Temp,"=",";|" & vbCrLf,1)
			Else
				Temp = xmlHttp.responseText
				Pos = InStr(LCase(Temp),"charset")
				If Pos = 0 Then Pos = InStr(LCase(Temp),"lang")
				If Pos > 0 Then
					Temp = Mid(Temp,Pos)
					If InStr(Temp,"=") Then Code = ExtractStr(Temp,"=",">",1)
				End If
			End If
		End If
		xmlHttp.Abort
	End If
	If Code <> "" Then Code = RemoveBackslash(Code,"""","""",1)
	If LCase(Code) = "utf-8" Or LCase(Code) = "utf8" Then
		srcStrBak = Utf8Encode(srcStrBak)
	Else
		srcStrBak = ANSIEncode(srcStrBak)
	End If

	If UrlTemplate <> "" Then
		If InStr(LCase(UrlTemplate),"{url}") = 0 Then UrlTemplate = Url & UrlTemplate
		UrlTemplate = strReplace(UrlTemplate,"{url}",Url)
		UrlTemplate = strReplace(UrlTemplate,"{appid}",AppId)
		UrlTemplate = strReplace(UrlTemplate,"{text}",srcStrBak)
		UrlTemplate = strReplace(UrlTemplate,"{from}",LangFrom)
		UrlTemplate = strReplace(UrlTemplate,"{to}",LangTo)
	Else
		UrlTemplate = Url
	End If

	If BodyData <> "" Then
		BodyData = strReplace(BodyData,"{url}",Url)
		BodyData = strReplace(BodyData,"{appid}",AppId)
		BodyData = strReplace(BodyData,"{text}",srcStrBak)
		BodyData = strReplace(BodyData,"{from}",LangFrom)
		BodyData = strReplace(BodyData,"{to}",LangTo)
	End If

    xmlHttp.Open Method,UrlTemplate,Async,User,Password
    If fType = 3 Then xmlHttp.setRequestHeader("If-Modified-Since","0")
    If RequestHeader <> "" And UCase(Method) <> "GET" Then
		TempList = Split(RequestHeader,vbCrLf)
		For i = 0 To UBound(TempList)
			Temp = TempList(i)
    		Pos = InStr(Temp,",")
    		If Pos = 0 Then Pos = InStr(Temp,":")
			If Pos > 0 Then
				bstrHeader = Trim(Left(Temp,Pos-1))
				bstrValue = Trim(Mid(Temp,Pos+1))
				If bstrValue <> "" Then
					bstrValue = strReplace(bstrValue,"{url}",Url)
					bstrValue = strReplace(bstrValue,"{appid}",AppId)
					bstrValue = strReplace(bstrValue,"{text}",srcStrBak)
					bstrValue = strReplace(bstrValue,"{from}",LangFrom)
					bstrValue = strReplace(bstrValue,"{to}",LangTo)
					If LCase(bstrHeader) = "content-length" Then
						xmlHttp.setRequestHeader bstrHeader,LenB(bstrValue)
					Else
						xmlHttp.setRequestHeader bstrHeader,bstrValue
					End If
				End If
			End If
		Next i
	End If
   	xmlHttp.send(BodyData)
   	StatusValue = OnReadyStateChange(xmlHttp,4,WaitTimes)
   	If StatusValue < 4 Then GoTo ErrorHandler
	If xmlHttp.Status = 200 Or xmlHttp.Status = 206 Then
		If fType = 1 Then
			getTranslate = BytesToBstr(xmlHttp.responseBody,Code)
		Else
			If responseType = "responsetext" Then
				trnStr = xmlHttp.responseText
			ElseIf responseType = "responsexml" Then
				getTranslate = ReadXML(xmlHttp.responseXML,TranBeforeStr,TranAfterStr)
			ElseIf responseType = "responsestream" Then
				trnStr = BytesToBstr(xmlHttp.responseStream,Code)
			ElseIf responseType = "responsebody" Then
				trnStr = BytesToBstr(xmlHttp.responseBody,Code)
			End If

			If responseType <> "responsexml" Then
				getTranslate = ExtractStr(trnStr,TranBeforeStr,TranAfterStr,0)
			End If
		End If
	End If
	xmlHttp.Abort
	On Error GoTo 0
	Exit Function

    ErrorHandler:
    If Err.Number <> 0 Then
    	If fType = 3 Then getTranslate = "NotConnected"
   	ElseIf fType = 3 Then
   		getTranslate = IIf(StatusValue <= 1,"NotConnected","Timeout")
	End If
	xmlHttp.Abort
End Function


'在 wTimes 等待时间内轮询服务器的状态
'tValue 为目标值，当 wTimes = 0 时为默认等待时间
Function OnReadyStateChange(xmlHttp As Object,tValue As Long,wTimes As Long) As Long
	Dim StartTime As Long
	StartTime = Timer
	If wTimes = 0 Then wTimes = DefaultWaitTimes
	OnReadyStateChange = xmlHttp.readyState
	Do While OnReadyStateChange < tValue
		OnReadyStateChange = xmlHttp.readyState
		If (Timer - StartTime) > wTimes Then Exit Do
	Loop
End Function


'不区分大小写的字符替换 (保留未替换字符的大小写)
Function strReplace(s As String,find As String,repwith As String) As String
	Dim i As Long,fL As Long,Ls As String,Lf As String
	strReplace = s
	If s = "" Or find = "" Then Exit Function
	Ls = LCase(s)
	Lf = LCase(find)
	i = InStr(Ls,Lf)
	If i = 0 Then Exit Function
	fL = Len(find)
	Do While i > 0
		strReplace = Replace(strReplace,Mid(strReplace,i,fL),repwith)
		i = InStr(i + fL,Ls,Lf)
	Loop
End Function


'Utf-8 编码
Function Utf8Encode(textStr As String) As String
	Dim Wch As String,Uch As String,Szret As String,i As Long,Nasc As Long
	Utf8Encode = textStr
	If Trim(textStr) = "" Then Exit Function
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


'ANSI 编码
Public Function ANSIEncode(textStr As String) As String
    Dim i As Long,startIndex As Long,endIndex As Long,x() As Byte,Szret As String
    ANSIEncode = textStr
    If Trim(textStr) = "" Then Exit Function
    x = StrConv(textStr,vbFromUnicode)
    startIndex = LBound(x)
    endIndex = UBound(x)
    For i = startIndex To endIndex
        Szret = Szret & "%" & Hex(x(i))
    Next i
    ANSIEncode = Szret
End Function


'转换字符的编码格式
Function ConvStr(textStr As String,inCode As String,outCode As String) As String
	Dim objStream As Object
    ConvStr = textStr
    If Trim(textStr) = "" Or inCode = "" Or outCode = "" Then Exit Function
    On Error GoTo ErrorMsg
    Set objStream = CreateObject("Adodb.Stream")
    If Not objStream Is Nothing Then
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
	End If
    Exit Function
    ErrorMsg:
    Err.Source = "Adodb.Stream"
    Call sysErrorMassage(Err,1)
End Function


'转换二进制数据为指定编码格式的字符
Function BytesToBstr(strBody As Variant,outCode As String) As String
    Dim objStream As Object
    If LenB(strBody) = 0 Or outCode = "" Then Exit Function
    On Error GoTo ErrorMsg
    Set objStream = CreateObject("Adodb.Stream")
    If Not objStream Is Nothing Then
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
    End If
    Exit Function
    ErrorMsg:
    Err.Source = "Adodb.Stream"
    Call sysErrorMassage(Err,1)
End Function


'写入二进制数据到文件
Function BytesToFile(strBody As Variant,File As String) As Boolean
	Dim objStream As Object
	BytesToFile = False
	If LenB(strBody) = 0 Or File = "" Then Exit Function
	On Error GoTo ErrorMsg
    Set objStream = CreateObject("Adodb.Stream")
    If Not objStream Is Nothing Then
	    With objStream
			.Type = 1
			.Mode = 3
			.Open
			.Write(strBody)
			.Position = 0
			.SaveToFile File,2
			.Flush
			.Close
		End With
		Set objStream = Nothing
		BytesToFile = True
	End If
	Exit Function
    ErrorMsg:
    Err.Source = "Adodb.Stream"
    Call sysErrorMassage(Err,1)
End Function


'解析 XML 格式对象并提取翻译文本
Function ReadXML(xmlObj As Object,IdNames As String,TagNames As String) As String
	Dim xmlDoc As Object,Node As Object,Item As Object,IdName As String,TagName As String
	Dim i As Long,j As Long,k As Long,Max As Long
	If xmlObj Is Nothing Then Exit Function
	If IdNames = "" And TagNames = "" Then Exit Function

	On Error GoTo ErrorMsg
	Set xmlDoc = CreateObject("Microsoft.XMLDOM")
	If Not xmlDoc Is Nothing Then
		'xmlDoc.Async = False
		'xmlDoc.ValidateOnParse = False
		'xmlDoc.loadXML(xmlObj)	'加载字串
		xmlDoc.Load(xmlObj)		'加载对象
  		If xmlDoc.ReadyState > 2 Then
  			IdNameArray = Split(IdNames,"|")
			TagNameArray = Split(TagNames,"|")
			k = UBound(TagNameArray)
			On Error Resume Next
			For i = 0 To UBound(IdNameArray)
				For j = 0 To k
					IdName = IdNameArray(i)
					TagName = TagNameArray(j)
					If IdName <> "" And TagName = "" Then
						Set Item = xmlDoc.getElementById(IdName)
						If Item Is Nothing Then
							Set Item = xmlDoc.getElementsByTagName(IdName)
						End If
					ElseIf IdName <> "" And TagName <> "" Then
						Set Node = xmlDoc.getElementById(IdName)
						If Node Is Nothing Then
							Set Item = xmlDoc.getElementsByTagName(TagName)
						Else
							Set Item = Node.getElementsByTagName(TagName)
							If Item.Length = 0 Then Set Item = xmlDoc.getElementById(IdName)
						End If
					ElseIf IdName = "" And TagName <> "" Then
						Set Item = xmlDoc.getElementsByTagName(TagName)
					End If
					Max = Item.Length
					If Max > 0 Then Exit For
				Next j
				If Max > 0 Then Exit For
			Next i
			On Error GoTo 0
			If Max > 0 Then
				For i = 0 To Max-1
					If ReadXML <> "" Then ReadXML = ReadXML & Item(i).Text
					If ReadXML = "" Then ReadXML = Item(i).Text    'firstChild.nodeValue
				Next i
			End If
		End If
		Set xmlDoc = Nothing
	End If
	Exit Function

	ErrorMsg:
	Err.Source = "Microsoft.XMLDOM"
	Call sysErrorMassage(Err,1)
End Function


'提取指定前后字符之间的值
Function ExtractStr(textStr As String,BeforeStr As String,AfterStr As String,fType As Long) As String
	Dim i As Long,j As Long,k As Long,L1 As Long,L2 As Long
	Dim Temp As String,bStr As String,aStr As String,toFindText As String

	If Trim(textStr) = "" Or (BeforeStr = "" And AfterStr = "") Then Exit Function
	toFindText = textStr & vbCrLf
	BeforeStrArray = Split(BeforeStr,"|")
	AfterStrArray = Split(AfterStr,"|")
	k = UBound(AfterStrArray)
	For i = 0 To UBound(BeforeStrArray)
		For j = 0 To k
			bStr = BeforeStrArray(i)
			aStr = AfterStrArray(j)
			L1 = InStr(toFindText,bStr)
			Do While L1 > 0
				L1 = L1 + Len(bStr)
				L2 = InStr(L1,toFindText,aStr)
				If fType > 0 And L2 = 0 Then L2 = InStr(L1,toFindText,vbCrLf)
				If L2 > 0 Then
					Temp = Mid(toFindText,L1,L2-L1)
					If ExtractStr <> "" Then ExtractStr = ExtractStr & Temp
					If ExtractStr = "" Then ExtractStr = Temp
					If fType > 0 Then Exit Do
				End If
				L1 = InStr(L1,toFindText,bStr)
			Loop
			If ExtractStr <> "" Then Exit For
		Next j
		If ExtractStr <> "" Then Exit For
	Next i
End Function


'fType = 0 检查字串是否包含指定字符，并找出指定字符的位置
'fType = 1 检查字串是否只包含指定字符
'fType = 2 检查字串是否包含指定字符
'fType = 3 检查字串是否只包含大小混写的指定字符
'AscRange  定义字串检查范围 (字符代码，可用 Min - Max 表示范围)
Function CheckStr(textStr As String,AscRange As String,fType As Long) As Boolean
	Dim i As Long,j As Long,n As Long,k As Long,l As Long,m As Long
	Dim InpAsc As Long,Pos As Long,Length As Long,Temp As String
	Dim TempStr As String,Stemp As Boolean,FindStemp As Boolean
	CheckStr = False
	TempStr = Trim(textStr)
	If TempStr = "" Or AscRange = "" Then Exit Function
	If fType = 3 Then
		If InStr(TempStr," ") Then Exit Function
		If Len(TempStr) > 2 Then TempStr = Mid(TempStr,2)
	End If
	k = 0
	l = 0
	m = 0
	Stemp = False
	Length = Len(TempStr)
	AscValue = Split(AscRange,",",-1)
	n = UBound(AscValue)
	For i = 1 To Length
		InpAsc = AscW(Mid(TempStr,i,1))
		FindStemp = False
		For j = 0 To n
			Temp = AscValue(j)
			Pos = InStr(Temp,"-")
			If Pos <> 0 Then
				Min = Left(Temp,Pos-1)
				Max = Mid(Temp,Pos+1)
			Else
				Min = Temp
				Max = Temp
			End If
			If Min <> "" And Max <> "" Then
				If InpAsc >= CLng(Min) And InpAsc <= CLng(Max) Then FindStemp = True
			ElseIf Min <> "" And Max = "" Then
				If InpAsc >= CLng(Min) Then FindStemp = True
			ElseIf Min = "" And Max <> "" Then
				If InpAsc <= CLng(Max) Then FindStemp = True
			End If
			If FindStemp = True Then
				If fType = 0 Then
					k = i
				ElseIf fType = 3 Then
					If j = 0 Then l = l + 1 Else If j = 1 Then m = m + 1
				End If
				Exit For
			End If
		Next j
		If fType = 1 Or fType = 3 Then
			If FindStemp = False Then Stemp = True
		Else
			If FindStemp = True Then Stemp = True
		End If
		If Stemp = True Then Exit For
	Next i
	If fType = 0 Or fType = 2 Then
		If Stemp = True Then
			CheckStr = True
			If fType = 0 Then fType = k
		End If
	ElseIf fType = 1 Then
		If Stemp = False Then CheckStr = True
	ElseIf fType = 3 Then
		If Stemp = False And l > 0 And m > 0 Then CheckStr = True
	End If
End Function


'分行翻译处理
Function SplitTran(xmlHttp As Object,srcStr As String,LangPair As String,Arg As String,fType As Long) As String
	Dim i As Long,srcStrBak As String,srcStringBak As String,Temp As String,Stemp As Boolean
	Dim EngineID As Long,CheckID As Long,iVoSrc As Long,mCheckSrc As Long,mHanding As Long

	If Trim(srcStr) = "" Or LangPair = "" Or Arg = "" Then Exit Function
	TempArray = Split(Arg,JoinStr,-1)
	EngineID = StrToLong(TempArray(0))
	CheckID = StrToLong(TempArray(1))
	iVoSrc =  StrToLong(TempArray(2))
	mCheckSrc = StrToLong(TempArray(3))
	mHanding = StrToLong(TempArray(4))

	'用替换法拆分字串
	srcStrBak = srcStr
	LineSplitChar = "\r\n,\r,\n"
	FindStrArr = Split(Convert(LineSplitChar),",",-1)
	For i = 0 To UBound(FindStrArr)
		FindStr = Trim(FindStrArr(i))
		If InStr(srcStrBak,FindStr) Then
			srcStrBak = Replace(srcStrBak,FindStr,"*c!N!g*")
		End If
	Next i
	srcStrArr = Split(srcStrBak,"*c!N!g*",-1)

	'获取每行的翻译
	Temp = srcStr
	Stemp = False
	For i = 0 To UBound(srcStrArr)
		srcString = srcStrArr(i)
		srcStringBak = srcString
		If srcString <> "" Then
			If mHanding = 0 Then
				If mCheckSrc = 1 Then srcString = CheckHanding(CheckID,srcStringBak,srcString,iVoSrc)
				If InStr(srcString,"&") Then srcString = Replace(srcString,"&","")
				If srcString <> "" And srcString <> srcStringBak Then
					srcStr = Replace(srcStr,srcStringBak,srcString,,1)
				End If
			End If
			trnString = getTranslate(EngineID,xmlHttp,srcString,LangPair,fType)
			If trnString <> "" And trnString <> srcStringBak Then
				Temp = Replace(Temp,srcStringBak,trnString,,1)
				Stemp = True
			End If
		End If
	Next i
	If Stemp = True Then SplitTran = Temp
End Function


'替换特定字符
'fType = 0 正向替换，使用第一个替换字符配置
'fType = 1 还原替换，使用第一个替换字符配置
'fType = 2 正向替换，使用第二个替换字符配置
'fType = 3 还原替换，使用第二个替换字符配置
'Record = 0 不记录替换字符
'Record = 1 记录替换字符
Function ReplaceStr(CheckID As Long,trnStr As String,fType As Long,Record As Long) As String
	Dim i As Long,PreStr As String,AppStr As String
	ReplaceStr = trnStr
	PreRepStr = ""
	AppRepStr = ""
	If Trim(trnStr) = "" Then Exit Function
	'获取选定配置的参数
	TempArray = Split(CheckDataListBak(CheckID),JoinStr)
	SetsArray = Split(TempArray(1),SubJoinStr)
	If fType < 2 Then AutoRepChar = SetsArray(11) Else AutoRepChar = SetsArray(12)
	If AutoRepChar <> "" Then
		FindStrArr = Split(AutoRepChar,",",-1)
		For i = 0 To UBound(FindStrArr)
			FindStr = FindStrArr(i)
			PreStr = ""
			AppStr = ""
			If InStr(FindStr,"|") Then
				TempArray = Split(FindStr,"|")
				If fType = 0 Or fType = 2 Then
					PreStr = TempArray(0)
					AppStr = TempArray(1)
				Else
					PreStr = TempArray(1)
					AppStr = TempArray(0)
				End If
				cPreStr = Convert(PreStr)
				cAppStr = Convert(AppStr)
			End If
			If PreStr <> "" And InStr(ReplaceStr,cPreStr) Then
				ReplaceStr = Replace(ReplaceStr,cPreStr,cAppStr)
				If Record = 1 Then
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
End Function


'检查修正快捷键、终止符和加速器
Function CheckHanding(CheckID As Long,srcStr As String,trnStr As String,iVo As Long) As String
	Dim i As Long,srcStrBak As String,trnStrBak As String,LineSplitMode As Long
	Dim srcNum As Long,trnNum As Long,srcSplitNum As Long,trnSplitNum As Long,Stemp As Boolean
	Dim FindStr As String,srcStrArr() As String,trnStrArr() As String,TempArray() As String
	Dim k As Long,l As Long,m As Long

	'参数初始化
	srcNum = 0
	trnNum = 0
	srcSplitNum = 0
	trnSplitNum = 0
	srcStrBak = srcStr
	trnStrBak = trnStr
	CheckHanding = trnStr
	If Trim(srcStr) = "" Or Trim(trnStr) = "" Then Exit Function

	'获取选定配置的参数
	TempArray = Split(CheckDataList(CheckID),JoinStr)
	SetsArray = Split(TempArray(1),SubJoinStr)
	ExcludeChar = SetsArray(0)
	PreInsertSplitChar = SetsArray(1)
	KeepCharPair = SetsArray(3)
	AccessKeyChar = SetsArray(13)
	LineSplitMode = StrToLong(SetsArray(15))
	AppInsertSplitChar = SetsArray(16)
	ReplaceSplitChar = SetsArray(17)

	TempArray = Split(ProjectDataList(iVo),JoinStr)
	SetsArray = Split(TempArray(1),LngJoinStr)
	EnableStringSplit = StrToLong(SetsArray(16))

	'配置参数数组化
	If ExcludeChar <> "" Then ExcludeCharArr = Split(ExcludeChar,",",-1)
	If KeepCharPair <> "" Then KeepCharPairArr = Split(KeepCharPair,",",-1)
	If AccessKeyChar <> "" Then AccessKeyCharArr = Split(AccessKeyChar,",",-1)

	If EnableStringSplit = 1 Then
		LineSplitChar = PreInsertSplitChar & AppInsertSplitChar & ReplaceSplitChar
		Temp = PreInsertSplitChar & "," & AppInsertSplitChar & "," & ReplaceSplitChar
		If LineSplitChar <> "" Then LineSplitCharArr = Split(Temp,",",-1)
		If PreInsertSplitChar <> "" Then k = UBound(Split(PreInsertSplitChar,",",-1)) + 1
		If AppInsertSplitChar <> "" Then l = UBound(Split(AppInsertSplitChar,",",-1)) + 1
		If ReplaceSplitChar <> "" Then m = UBound(Split(ReplaceSplitChar,",",-1)) + 1
	End If

	'排除字串中的非快捷键
	If ExcludeChar <> "" Then
		For i = 0 To UBound(ExcludeCharArr)
			FindStr = LTrim(ExcludeCharArr(i))
			If FindStr <> "" Then
				srcStrBak = Replace(srcStrBak,FindStr,"*a" & i & "!N!" & i & "d*")
				trnStrBak = Replace(trnStrBak,FindStr,"*a" & i & "!N!" & i & "d*")
			End If
		Next i
	End If

	'过滤不是快捷键的快捷键
	If KeepCharPair <> "" Then
		For i = 0 To UBound(KeepCharPairArr)
			FindStr = Trim(KeepCharPairArr(i))
			If FindStr <> "" Then
				LFindStr = Trim(Left(FindStr,1))
				RFindStr = Trim(Right(FindStr,1))
				ToRepStr = LFindStr & "&" & RFindStr
				BeRepStr = LFindStr & "*!N!" & i & "!M!" & i & "!N!*" & RFindStr
				srcStrBak = Replace(srcStrBak,ToRepStr,BeRepStr)
				trnStrBak = Replace(trnStrBak,ToRepStr,BeRepStr)
			End If
		Next i
	End If

	'用替换法拆分字串
	If EnableStringSplit = 1 Then
		BaksrcStr = srcStrBak
		BaktrnStr = trnStrBak
		If LineSplitChar <> "" Then
			For i = 0 To UBound(LineSplitCharArr)
				FindStr = Trim(LineSplitCharArr(i))
				If FindStr <> "" Then
					Stemp = False
					If LineSplitMode = 1 Then
						srcNum = UBound(Split(BaksrcStr,FindStr,-1))
						trnNum = UBound(Split(BaktrnStr,FindStr,-1))
						If srcNum = trnNum And srcNum > 0 And trnNum > 0 Then Stemp = True
					End If
					If LineSplitMode = 0 Or Stemp = True Then
						If InStr(LCase(AccessKeyChar),LCase(FindStr)) Then
							BaksrcStr = Insert(BaksrcStr,FindStr,"*c!N!g*",1)
							BaktrnStr = Insert(BaktrnStr,FindStr,"*c!N!g*",1)
						ElseIf i < k And k <> 0 Then
							BaksrcStr = Replace(BaksrcStr,FindStr,"*c!N!g*" & FindStr)
							BaktrnStr = Replace(BaktrnStr,FindStr,"*c!N!g*" & FindStr)
						ElseIf i >= k And i < k + l + 1 And l <> 0 Then
							BaksrcStr = Replace(BaksrcStr,FindStr,FindStr & "*c!N!g*")
							BaktrnStr = Replace(BaktrnStr,FindStr,FindStr & "*c!N!g*")
						ElseIf i >= k + l And i < k + l + m + 2 And m <> 0 Then
							BaksrcStr = Replace(BaksrcStr,FindStr,"*c!N!g*")
							BaktrnStr = Replace(BaktrnStr,FindStr,"*c!N!g*")
						End If
					End If
				End If
			Next i
		End If
		srcStrArr = Split(BaksrcStr,"*c!N!g*",-1)
		trnStrArr = Split(BaktrnStr,"*c!N!g*",-1)

		'字串处理
		Stemp = False
		srcNum = UBound(srcStrArr)
		trnNum = UBound(trnStrArr)
		If srcNum > 0 And trnNum > 0 Then
			If LineSplitMode = 0 Then Stemp = True
			If LineSplitMode = 1 And srcNum = trnNum Then Stemp = True
		End If
		If Stemp = True Then
			TempArray = MergeArray(srcStrArr,trnStrArr)
			trnStrBak = ReplaceStrSplit(CheckID,trnStrBak,TempArray,iVo)
		Else
			trnStrBak = StringReplace(CheckID,srcStrBak,trnStrBak,iVo)
		End If
	Else
		trnStrBak = StringReplace(CheckID,srcStrBak,trnStrBak,iVo)
	End If

	'计算快捷键数
	BaksrcStr = srcStrBak
	BaktrnStr = trnStrBak
	toRepStr = Trim(AccessKeyCharArr(0))
	If AccessKeyChar <> "" Then
		For i = 0 To UBound(AccessKeyCharArr)
			FindStr = Trim(AccessKeyCharArr(i))
			If FindStr <> "" And FindStr <> toRepStr Then
				BaksrcStr = Replace(BaksrcStr,FindStr,toRepStr)
				BaktrnStr = Replace(BaktrnStr,FindStr,toRepStr)
			End If
		Next i
	End If
	StringSrc.AccKeyNum = UBound(Split(BaksrcStr,toRepStr,-1))
	StringTrn.AccKeyNum = UBound(Split(BaktrnStr,toRepStr,-1))

	'还原不是快捷键的快捷键
	If KeepCharPair <> "" Then
		For i = 0 To UBound(KeepCharPairArr)
			FindStr = Trim(KeepCharPairArr(i))
			If FindStr <> "" Then
				LFindStr = Trim(Left(FindStr,1))
				RFindStr = Trim(Right(FindStr,1))
				ToRepStr = LFindStr & "*!N!" & i & "!M!" & i & "!N!*" & RFindStr
				BeRepStr = LFindStr & "&" & RFindStr
				srcStrBak = Replace(srcStrBak,ToRepStr,BeRepStr)
				trnStrBak = Replace(trnStrBak,ToRepStr,BeRepStr)
			End If
		Next i
	End If

	'还原字串中被排除的非快捷键
	If ExcludeChar <> "" Then
		For i = 0 To UBound(ExcludeCharArr)
			FindStr = LTrim(ExcludeCharArr(i))
			If FindStr <> "" Then
				srcStrBak = Replace(srcStrBak,"*a" & i & "!N!" & i & "d*",FindStr)
				trnStrBak = Replace(trnStrBak,"*a" & i & "!N!" & i & "d*",FindStr)
			End If
		Next i
	End If
	CheckHanding = trnStrBak
End Function


'在快捷键后插入特定字符并以此拆分字串
Function Insert(SplitString As String,SplitStr As String,InsStr As String,Leng As Long) As String
	Dim i As Long,j As Long,accesskeyStr As String
	Insert = SplitString
	If UBound(Split(SplitString,SplitStr)) < 2 Then Exit Function
	i = InStr(Insert,SplitStr)
	Do While i > Leng
		j = InStr(i + 1,Insert,SplitStr)
		If j > i Then
			accesskeyStr = Mid(Insert,i - Leng,j - i)
			If accesskeyStr <> "" Then
				Insert = Replace(Insert,accesskeyStr,accesskeyStr & InsStr)
			End If
		End If
		i = InStr(i + 1,Insert,SplitStr)
	Loop
	'PSL.Output "Insert = " & Insert       '调试用
End Function


'读取数组中的每个字串并替换处理
Function ReplaceStrSplit(CheckID As Long,trnStr As String,StrSplitArr() As String,iVo As Long) As String
	Dim srcStrSplit As String,trnStrSplit As String,trnStrSplitNew As String,i As Long,j As Long

	j = 1
	ReplaceStrSplit = trnStr
	For i = 0 To UBound(StrSplitArr) Step 2
		srcStrSplit = StrSplitArr(i)
		trnStrSplit = StrSplitArr(i+1)
		trnStrSplitNew = StringReplace(CheckID,srcStrSplit,trnStrSplit,iVo)

		'处理在前后行中包含的重复字符
		If trnStrSplit <> trnStrSplitNew Then
			Temp = Replace(ReplaceStrSplit,trnStrSplit,trnStrSplitNew,j,1)
			If j = 1 Then
				ReplaceStrSplit = Temp
			Else
				ReplaceStrSplit = Left(ReplaceStrSplit,j - 1) & Temp
			End If
		End If
		j = j + Len(trnStrSplitNew)

		'对每行的数据进行连接，用于消息输出
		TPreSpaceSrc = TPreSpaceSrc & StringSrc.PreSpace
		TPreSpaceTrn = TPreSpaceTrn & StringTrn.PreSpace
		TacckeySrc = TacckeySrc & StringSrc.AccKey
		TacckeyTrn = TacckeyTrn & StringTrn.AccKey
		TEndStringSrc = TEndStringSrc & StringSrc.EndString
		TEndStringTrn = TEndStringTrn & StringTrn.EndString
		TShortcutSrc = TShortcutSrc & StringSrc.Shortcut
		TShortcutTrn = TShortcutTrn & StringTrn.Shortcut
		TEndSpaceSrc = TEndSpaceSrc & StringSrc.EndSpace
		TEndSpaceTrn = TEndSpaceTrn & StringTrn.EndSpace
		TSpaceTrn = TSpaceTrn & StringTrn.Spaces
		TExpStringTrn = TExpStringTrn & StringTrn.ExpString
		TPreStringTrn = TPreStringTrn & StringTrn.PreString
		TMoveAcckey = TMoveAcckey & MoveAcckey
	Next i

	'为调用消息输出，用原有变量替换连接后的数据
	StringSrc.PreSpace = TPreSpaceSrc
	StringTrn.PreSpace = TPreSpaceTrn
	StringSrc.AccKey = TacckeySrc
	StringTrn.AccKey = TacckeyTrn
	StringSrc.EndString = TEndStringSrc
	StringTrn.EndString = TEndStringTrn
	StringSrc.Shortcut = TShortcutSrc
	StringTrn.Shortcut = TShortcutTrn
	StringSrc.EndSpace = TEndSpaceSrc
	StringTrn.EndSpace = TEndSpaceTrn
	StringTrn.Spaces = TSpaceTrn
	StringTrn.ExpString = TExpStringTrn
	StringTrn.PreString = TPreStringTrn
	MoveAcckey = TMoveAcckey
End Function


'按行获取字串的各个字段并替换翻译字符串
Function StringReplace(CheckID As Long,srcStr As String,trnStr As String,iVo As Long) As String
	Dim i As Long,j As Long,x As Long,y As Long,m As Long,n As Long
	Dim AsiaKey As Long,AddAccessKeyWithFirstChar As Long,LeadingSpaceInSource As Long
	Dim LeadingSpaceInTarget As Long,LeadingSpaceInBoth As Long,TrailingSpaceInSource As Long
	Dim TrailingSpaceInTarget As Long,TrailingSpaceInBoth As Long,AccessKeyInSource As Long
	Dim AccessKeyInTarget As Long,AccessKeyInBoth As Long,EndCharInSource As Long,EndCharInTarget As Long
	Dim EndCharInBoth As Long,ShortcutInSource As Long,ShortcutInTarget As Long,ShortcutInBoth As Long
	Dim DeleteExtraSpace As Long,TranslateEndChar As Long,AccKeyInShort As Long,TempArray() As String
	Dim FindStr As String,LastStringTrn As String,Temp As String,TempBak As String,Stemp As Boolean

	'参数初始化
	StringSrc.Length = 0
	StringTrn.Length = 0
	StringSrc.PreSpace = ""
	StringTrn.PreSpace = ""
	StringSrc.EndSpace = ""
	StringTrn.EndSpace = ""
	StringSrc.AccKey = ""
	StringTrn.AccKey = ""
	StringSrc.AccKeyIFR = ""
	StringTrn.AccKeyIFR = ""
	StringSrc.AccKeyKey = ""
	StringTrn.AccKeyKey = ""
	StringSrc.AccKeyPos = 0
	StringTrn.AccKeyPos = 0
	StringSrc.EndString = ""
	StringTrn.EndString = ""
	StringSrc.Shortcut = ""
	StringTrn.Shortcut = ""
	StringTrn.Spaces = ""
	StringTrn.ExpString = ""
	StringTrn.PreString = ""
	LastStringTrn = ""
	MoveAcckey = ""
	StringReplace = trnStr
	If Trim(srcStr) = "" Or Trim(trnStr) = "" Then Exit Function

	'获取选定配置的参数
	TempArray = Split(CheckDataList(CheckID),JoinStr)
	SetsArray = Split(TempArray(1),SubJoinStr)
	CheckBracket = SetsArray(2)
	AsiaKey = StrToLong(SetsArray(4))
	CheckEndChar = SetsArray(5)
	NoTrnEndChar = SetsArray(6)
	AutoTrnEndChar = SetsArray(7)
	CheckShortChar = SetsArray(8)
	CheckShortKey = SetsArray(9)
	KeepShortKey = SetsArray(10)
	AccessKeyChar = SetsArray(13)
	AddAccessKeyWithFirstChar = StrToLong(SetsArray(14))

	TempArray = Split(ProjectDataList(iVo),JoinStr)
	SetsArray = Split(TempArray(1),LngJoinStr)
	LeadingSpaceInSource = StrToLong(SetsArray(0))
	LeadingSpaceInTarget = StrToLong(SetsArray(1))
	LeadingSpaceInBoth = StrToLong(SetsArray(2))
	TrailingSpaceInSource = StrToLong(SetsArray(3))
	TrailingSpaceInTarget = StrToLong(SetsArray(4))
	TrailingSpaceInBoth = StrToLong(SetsArray(5))
	AccessKeyInSource = StrToLong(SetsArray(6))
	AccessKeyInTarget = StrToLong(SetsArray(7))
	AccessKeyInBoth = StrToLong(SetsArray(8))
	EndCharInSource = StrToLong(SetsArray(9))
	EndCharInTarget = StrToLong(SetsArray(10))
	EndCharInBoth = StrToLong(SetsArray(11))
	ShortcutInSource = StrToLong(SetsArray(12))
	ShortcutInTarget = StrToLong(SetsArray(13))
	ShortcutInBoth = StrToLong(SetsArray(14))
	DeleteExtraSpace = StrToLong(SetsArray(15))
	TranslateEndChar = StrToLong(SetsArray(17))
	AccKeyInShort = StrToLong(SetsArray(18))

	'配置参数数组化
	If CheckBracket <> "" Then CheckBracketArr = Split(CheckBracket,",",-1)
	If CheckEndChar <> "" Then CheckEndCharArr = Split(CheckEndChar," ",-1)
	If AutoTrnEndChar <> "" Then AutoTrnEndCharArr = Split(AutoTrnEndChar," ",-1)
	If CheckShortChar <> "" Then CheckShortCharArr = Split(CheckShortChar,",",-1)
	If AccessKeyChar <> "" Then AccessKeyCharArr = Split(AccessKeyChar,",",-1)

	'获取来源和翻译的长度
	StringSrc.Length = Len(srcStr)
	StringTrn.Length = Len(trnStr)

	'获取来源和翻译的前置空格
	StringSrc.PreSpace = Space(StringSrc.Length - Len(LTrim(srcStr)))
	StringTrn.PreSpace = Space(StringTrn.Length - Len(LTrim(trnStr)))

	'获取来源和翻译的尾随空格
	StringSrc.EndSpace = Space(StringSrc.Length - Len(RTrim(srcStr)))
	StringTrn.EndSpace = Space(StringTrn.Length - Len(RTrim(trnStr)))

	'获取来源和翻译的加速器
	If CheckShortChar <> "" Then
		CheckShortKey = CheckShortKey & "," & KeepShortKey
		If AccessKeyChar <> "" Then m = UBound(AccessKeyCharArr)
		For i = 0 To UBound(CheckShortCharArr)
			FindStr = Trim(CheckShortCharArr(i))
			If FindStr <> "" Then
				For n = 0 To 1
					If n = 0 Then Temp = RTrim(srcStr)
					If n = 1 Then Temp = RTrim(trnStr)
					Shortcut = ""
					ShortcutKey = ""
					If InStrRev(LTrim(Temp),FindStr) > 1 Then
						y = InStrRev(Temp,FindStr)
						ShortcutKey = Trim(Mid(Temp,y + 1))
					End If
					If ShortcutKey <> "" Then
						If AccessKeyChar <> "" Then
							For j = 0 To m
								ShortcutKey =Replace(ShortcutKey,AccessKeyCharArr(j),"")
							Next j
						End If
						If ShortcutKey = "+" Then
							If CheckKeyCode(ShortcutKey,CheckShortKey) <> 0 Then
								Shortcut = Mid(Temp,y)
							End If
						ElseIf InStr(ShortcutKey,"+") Then
							x = 0
							TempArray = Split(ShortcutKey,"+",-1)
							For j = 0 To UBound(TempArray)
								x = x + CheckKeyCode(TempArray(j),CheckShortKey)
							Next j
							If x > 0 And x >= UBound(TempArray) Then
								Shortcut = Mid(Temp,y)
							End If
						Else
							If CheckKeyCode(ShortcutKey,CheckShortKey) <> 0 Then
								Shortcut = Mid(Temp,y)
							End If
						End If
						If Shortcut <> "" Then
							If n = 0 And StringSrc.Shortcut = "" Then
								StringSrc.Shortcut = Shortcut
								ShortcutKeySrc = ShortcutKey
							ElseIf n = 1 And StringTrn.Shortcut = "" Then
								StringTrn.Shortcut = Shortcut
								ShortcutKeyTrn = ShortcutKey
							End If
						End If
					End If
				Next n
			End If
			If StringSrc.Shortcut <> "" And StringTrn.Shortcut <> "" Then Exit For
		Next i
	End If

	'获取来源和翻译的终止符及其前后空格
	If CheckEndChar <> "" Then
		xTemp = Left(srcStr,StringSrc.Length - Len(StringSrc.Shortcut & StringSrc.EndSpace))
		yTemp = Left(trnStr,StringTrn.Length - Len(StringTrn.Shortcut & StringTrn.EndSpace))
		If NoTrnEndChar <> "" Then
			If CheckKeyCode(xTemp,NoTrnEndChar) = 1 Then xTemp = ""
			If CheckKeyCode(yTemp,NoTrnEndChar) = 1 Then yTemp = ""
		End If
	End If
	If xTemp <> "" And yTemp <> "" Then
		If AccessKeyChar <> "" Then m = UBound(AccessKeyCharArr)
		For i = 0 To UBound(CheckEndCharArr)
			FindStr = Trim(CheckEndCharArr(i))
			If FindStr <> "" Then
				PreFindStr = Left(FindStr,1)
				AppFindStr = Right(FindStr,1)
				For j = 0 To 1
					If j = 0 Then
						Temp = xTemp
						EndSpace = StringSrc.EndSpace
						Shortcut = StringSrc.Shortcut
					Else
						Temp = yTemp
						EndSpace = StringTrn.EndSpace
						Shortcut = StringTrn.Shortcut
					End If
					n = 0
					y = InStrRev(Temp,FindStr)
					If y > 0 Then
						If Trim(Mid(Temp,y)) Like FindStr Then
							n = y
						ElseIf Right(Trim(Temp),1) = AppFindStr Then
							y = InStr(Temp,PreFindStr)
							Do While y > 0
								TempBak = Trim(Mid(Temp,y))
								If AccessKeyChar <> "" Then
									For x = 0 To m
										TempBak = Replace(TempBak,AccessKeyCharArr(x),"")
									Next x
								End If
								If TempBak Like FindStr Then
									n = y
									Exit Do
								End If
								y = InStr(y + 1,Temp,PreFindStr)
							Loop
						End If
					End If
					If n <> 0 Then
						PreStr = Left(Temp,n - 1)
						AppStr = Mid(Temp,n)
						x = Len(PreStr) - Len(RTrim(PreStr))
						If j = 0 And StringSrc.EndString = "" Then
							StringSrc.EndString = Space(x) & AppStr
						ElseIf j = 1 And StringTrn.EndString = "" Then
							StringTrn.EndString = Space(x) & AppStr
						End If
					End If
				Next j
			End If
			If StringSrc.EndString <> "" And StringTrn.EndString <> "" Then Exit For
		Next i
	End If

	'获取来源和翻译的快捷键位置及其字符
	If AccessKeyChar <> "" Then
		For i = 0 To UBound(AccessKeyCharArr)
			FindStr = Trim(AccessKeyCharArr(i))
			If FindStr <> "" Then
				For j = 0 To 1
					If j = 0 Then Temp = srcStr
					If j = 1 Then Temp = trnStr
					n = InStrRev(Temp,FindStr)
					If n > 0 Then
						If j = 0 And n > StringSrc.AccKeyPos Then
							StringSrc.AccKeyPos = n
							StringSrc.AccKeyIFR = FindStr
							StringSrc.AccKeyKey = Mid(Temp,n + Len(FindStr),1)
						ElseIf j = 1 And n > StringTrn.AccKeyPos Then
							StringTrn.AccKeyPos = n
							StringTrn.AccKeyIFR = FindStr
							StringTrn.AccKeyKey = Mid(Temp,n + Len(FindStr),1)
						End If
					End If
				Next j
			End If
		Next i
	End If
	If StringSrc.AccKeyIFR = "" Then StringSrc.AccKeyIFR = "&"
	If StringTrn.AccKeyIFR = "" Then StringTrn.AccKeyIFR = "&"

	'获取来源和翻译的快捷键 (包括快捷键前后的括号字符)
	If (StringSrc.AccKeyPos > 1 Or StringTrn.AccKeyPos > 1) And CheckBracket <> "" Then
		For i = 0 To UBound(CheckBracketArr)
			FindStr = Trim(CheckBracketArr(i))
			If FindStr <> "" Then
				PreFindStr = Trim(Left(FindStr,1))
				AppFindStr = Trim(Right(FindStr,1))
				For n = 0 To 1
					If n = 0 Then
						Temp = srcStr
						j = StringSrc.AccKeyPos
						xTemp = StringSrc.AccKeyIFR
						yTemp = StringSrc.AccKeyKey
					ElseIf n = 1 Then
						Temp = trnStr
						j = StringTrn.AccKeyPos
						xTemp = StringTrn.AccKeyIFR
						yTemp = StringTrn.AccKeyKey
					End If
					AccessKey = ""
					If j > 1 Then
						x = InStrRev(Temp,PreFindStr,j)
						y = InStr(j,Temp,AppFindStr)
						If x > 0 And y > x Then
							TempBak = Mid(Temp,x + 1,y - x - 1)
							If Trim(TempBak) = xTemp & yTemp Then
								AccessKey = Mid(Temp,x,y - x + 1)
								j = x
							End If
						End If
					ElseIf j = 1 Then
						AccessKey = xTemp
					End If
					If AccessKey <> "" Then
						If n = 0 And StringSrc.AccKey = "" Then
							StringSrc.AccKeyPos = j
							StringSrc.AccKey = AccessKey
						ElseIf n = 1 And StringTrn.AccKey = "" Then
							StringTrn.AccKeyPos = j
							StringTrn.AccKey = AccessKey
						End If
					End If
				Next n
			End If
			If StringSrc.AccKey <> "" And StringTrn.AccKey <> "" Then Exit For
		Next i
	End If
	If StringSrc.AccKey = "" And StringSrc.AccKeyPos > 0 Then StringSrc.AccKey = StringSrc.AccKeyIFR
	If StringTrn.AccKey = "" And StringTrn.AccKeyPos > 0 Then StringTrn.AccKey = StringTrn.AccKeyIFR

	'获取翻译的快捷键后面的非终止符和非加速器的字符(包括空格)
	If StringTrn.AccKeyPos > 0 Then
		x = Len(StringTrn.EndString & StringTrn.Shortcut & StringTrn.EndSpace)
		'If InStr(StringTrn.Shortcut,StringTrn.AcckeyIFR) Then x = Len(StringTrn.EndSpace)
		'If InStr(StringTrn.EndString,StringTrn.AcckeyIFR) Then x = Len(StringTrn.Shortcut & StringTrn.EndSpace)
		If StringTrn.Length > x Then
			Temp = Left(trnStr,StringTrn.Length - x)
			StringTrn.ExpString = Mid(Temp,StringTrn.AccKeyPos + Len(StringTrn.AccKey))
		End If
	End If

	'获取翻译的快捷键或终止符或加速器前面的空格
	Temp = StringTrn.AccKey & StringTrn.ExpString & StringTrn.EndString & StringTrn.Shortcut & StringTrn.EndSpace
	If Temp <> "" Then
		x = Len(Temp)
		If InStr(StringTrn.EndString & StringTrn.Shortcut,StringTrn.AccKeyIFR) Then
			x = Len(StringTrn.EndString & StringTrn.Shortcut & StringTrn.EndSpace)
		End If
		If StringTrn.Length > x Then
			Temp = Left(trnStr,StringTrn.Length - x)
			y = Len(Temp) - Len(RTrim(Temp))
			If y > 0 Then StringTrn.Spaces = Space(y)
		End If
	End If

	'获取翻译的快捷键前的终止符及其终止符前的空格
	If StringTrn.AccKey <> "" And CheckEndChar <> "" Then
		x = Len(StringTrn.Spaces & StringTrn.AccKey & StringTrn.ExpString & StringTrn.EndString & StringTrn.Shortcut & StringTrn.EndSpace)
		If InStr(StringTrn.EndString & StringTrn.Shortcut,StringTrn.AccKeyIFR) Then
			x = Len(StringTrn.Spaces & StringTrn.EndString & StringTrn.Shortcut & StringTrn.EndSpace)
		End If
		Temp = Left(trnStr,StringTrn.Length - x)
		If Trim(Temp) <> "" Then
			If NoTrnEndChar <> "" Then
				If CheckKeyCode(Temp,NoTrnEndChar) = 1 Then Temp = ""
			End If
			If Temp <> "" Then
				For i = 0 To UBound(CheckEndCharArr)
					FindStr = Trim(CheckEndCharArr(i))
					PreFindStr = Left(FindStr,1)
					AppFindStr = Right(FindStr,1)
					n = 0
					y = InStrRev(Temp,FindStr)
					If y > 0 Then
						If Trim(Mid(Temp,y)) Like FindStr Then
							n = y
						ElseIf Right(Trim(Temp),1) = AppFindStr Then
							y = InStr(Temp,PreFindStr)
							Do While y > 0
								TempBak = Trim(Mid(Temp,y))
								If TempBak Like FindStr Then
									n = y
									Exit Do
								End If
								y = InStr(y + 1,Temp,PreFindStr)
							Loop
						End If
					End If
					If n > 0 Then
						PreStr = Left(Temp,n - 1)
						AppStr = Mid(Temp,n)
						x = Len(PreStr) - Len(RTrim(PreStr))
						StringTrn.PreString = Space(x) & AppStr
					End If
					If StringTrn.PreString <> "" Then Exit For
				Next i
			End If
		End If
	End If

	'获取翻译中除已提取字符外的其他所有字符
	Temp = StringTrn.PreString & StringTrn.Spaces & StringTrn.AccKey & StringTrn.ExpString & StringTrn.EndString & StringTrn.Shortcut & StringTrn.EndSpace
	If Temp <> "" Then
		x = Len(Temp)
		If InStr(StringTrn.EndString & StringTrn.Shortcut,StringTrn.AccKeyIFR) Then
			x = Len(StringTrn.PreString & StringTrn.Spaces & StringTrn.EndString & StringTrn.Shortcut & StringTrn.EndSpace)
		End If
		Temp = LTrim(trnStr)
		y = Len(Temp)
		If y > x Then LastStringTrn = Left(Temp,y - x)
	Else
		LastStringTrn = Trim(trnStr)
	End If

	'保留符合条件的加速器翻译
	If StringSrc.Shortcut <> "" And StringTrn.Shortcut <> "" And KeepShortKey <> "" Then
		SrcKeyArr = Split(ShortcutKeySrc,"+",-1)
		x = UBound(SrcKeyArr)
		TrnKeyArr = Split(ShortcutKeyTrn,"+",-1)
		y = UBound(TrnKeyArr)
		If x = y Then
			For i = 0 To x
				Temp = Trim(SrcKeyArr(i))
				TempBak = Trim(TrnKeyArr(i))
				If Temp <> "" And TempBak <> "" Then
					If CheckKeyCode(TempBak,KeepShortKey) <> 0 Then
						StringSrc.Shortcut = Replace(StringSrc.Shortcut,Temp,TempBak)
					End If
				End If
			Next i
		End If
	End If

	'PSL.Output "------------------------------ "   		'调试用
	'PSL.Output "srcStr = " & srcStr                		'调试用
	'PSL.Output "trnStr = " & trnStr               			'调试用
	'PSL.Output "SpaceTrn = " & StringTrn.Spaces            '调试用
	'PSL.Output "KeySrc = " & StringTrnKeySrc               '调试用
	'PSL.Output "acckeySrc = " & StringSrc.AccKey          	'调试用
	'PSL.Output "acckeyTrn = " & StringTrn.AccKey         	'调试用
	'PSL.Output "EndStringSrc = " & StringSrc.EndString   	'调试用
	'PSL.Output "EndStringTrn = " & StringTrn.EndString   	'调试用
	'PSL.Output "ShortcutSrc = " & StringSrc.Shortcut     	'调试用
	'PSL.Output "ShortcutTrn = " & StringTrn.Shortcut     	'调试用
	'PSL.Output "ExpStringTrn = " & StringTrn.ExpString   	'调试用
	'PSL.Output "PreStringTrn = " & StringTrn.PreString   	'调试用
	'PSL.Output "LastStringTrn = " & LastStringTrn  		'调试用
	'PSL.Output "MoveAcckey = " & MoveAcckey                '调试用

	'备份参数值
	SpaceTrnBak = StringTrn.Spaces
	ExpStringTrnBak = StringTrn.ExpString
	PreStringTrnBak = StringTrn.PreString
	ShortcutSrcBak = StringSrc.Shortcut
	EndStringSrcBak = StringTrn.EndString

	'字串内容选择处理
	If AllCont <> 1 Then
		If Acceler <> 1 Then
			ShortcutInSource = 0
			ShortcutInTarget = 0
			ShortcutInBoth = 0
		End If
		If EndChar <> 1 Then
			EndCharInSource = 0
			EndCharInTarget = 0
			EndCharInBoth = 0
			TranslateEndChar = 0
		End If
		If AccKey <> 1 Then
			AccessKeyInSource = 0
			AccessKeyInTarget = 0
			AccessKeyInBoth = 0
			If AsiaKey = 1 Then
				If StringTrn.AccKey = StringTrn.AccKeyIFR Then AsiaKey = 0
			Else
				If StringTrn.AccKey <> StringTrn.AccKeyIFR Then AsiaKey = 1
			End If
		End If
	End If

	'数据集成
	If iVo >= 0 Then
		'执行检查规则
		If StringSrc.PreSpace <> "" And StringTrn.PreSpace = "" Then
			If LeadingSpaceInSource = 0 Then StringSrc.PreSpace = StringTrn.PreSpace
		ElseIf StringSrc.PreSpace = "" And StringTrn.PreSpace <> "" Then
			If LeadingSpaceInTarget = 0 Then StringSrc.PreSpace = StringTrn.PreSpace
		ElseIf StringSrc.PreSpace <> "" And StringTrn.PreSpace <> "" Then
			If LeadingSpaceInBoth = 0 Then StringSrc.PreSpace = StringTrn.PreSpace
			If LeadingSpaceInBoth = 2 Then StringSrc.PreSpace = ""
		End If
		If StringSrc.EndSpace <> "" And StringTrn.EndSpace = "" Then
			If LeadingSpaceInSource = 0 Then StringSrc.EndSpace = StringTrn.EndSpace
		ElseIf StringSrc.EndSpace = "" And StringTrn.EndSpace <> "" Then
			If TrailingSpaceInTarget = 0 Then StringSrc.EndSpace = StringTrn.EndSpace
		ElseIf StringSrc.EndSpace <> "" And StringTrn.EndSpace <> "" Then
			If TrailingSpaceInBoth = 0 Then StringSrc.EndSpace = StringTrn.EndSpace
			If TrailingSpaceInBoth = 2 Then StringSrc.EndSpace = ""
		End If
		If StringSrc.AccKey <> "" And StringTrn.AccKey = "" Then
			If AccessKeyInSource = 0 Then
				StringSrc.AccKey = StringTrn.AccKey
				StringSrc.AccKeyIFR = StringTrn.AccKeyIFR
				StringSrc.AccKeyKey = StringTrn.AccKeyKey
			End If
		ElseIf StringSrc.AccKey = "" And StringTrn.AccKey <> "" Then
			If AccessKeyInTarget = 0 Then
				StringSrc.AccKey = StringTrn.AccKey
				StringSrc.AccKeyIFR = StringTrn.AccKeyIFR
				StringSrc.AccKeyKey = StringTrn.AccKeyKey
			End If
		ElseIf StringSrc.AccKey <> "" And StringTrn.AccKey <> "" Then
			If AccessKeyInBoth = 0 Then
				StringSrc.AccKey = StringTrn.AccKey
				StringSrc.AccKeyIFR = StringTrn.AccKeyIFR
				StringSrc.AccKeyKey = StringTrn.AccKeyKey
			ElseIf AccessKeyInBoth = 2 Then
				StringSrc.AccKey = ""
			End If
		End If
		If StringSrc.EndString <> "" And StringTrn.EndString = "" Then
			If EndCharInSource = 0 Then StringSrc.EndString = StringTrn.EndString
		ElseIf StringSrc.EndString = "" And StringTrn.EndString <> "" Then
			If EndCharInTarget = 0 Then StringSrc.EndString = StringTrn.EndString
		ElseIf StringSrc.EndString <> "" And StringTrn.EndString <> "" Then
			If EndCharInBoth = 0 Then StringSrc.EndString = StringTrn.EndString
			If EndCharInBoth = 2 Then StringSrc.EndString = ""
		End If
		If StringSrc.Shortcut <> "" And StringTrn.Shortcut = "" Then
			If ShortcutInSource = 0 Then StringSrc.Shortcut = StringTrn.Shortcut
		ElseIf StringSrc.Shortcut = "" And StringTrn.Shortcut <> "" Then
			If ShortcutInTarget = 0 Then StringSrc.Shortcut = StringTrn.Shortcut
		ElseIf StringSrc.Shortcut <> "" And StringTrn.Shortcut <> "" Then
			If ShortcutInBoth = 0 Then StringSrc.Shortcut = StringTrn.Shortcut
			If ShortcutInBoth = 2 Then StringSrc.Shortcut = ""
		End If

		'设置快捷键方式
		If InStr(StringTrn.EndString & StringTrn.Shortcut,StringTrn.AccKeyIFR) Then
			StringTrn.ExpString = ""
			If AsiaKey = 0 Then StringTrn.PreString = ""
		End If
		If StringTrn.AccKey = StringTrn.AccKeyIFR Then StringTrn.AccKey = StringTrn.AccKeyIFR & StringTrn.AccKeyKey
		If StringSrc.AccKey <> "" Then
			If AsiaKey = 0 Then
				StringSrc.AccKey = StringSrc.AccKeyIFR & StringSrc.AccKeyKey
				KeySrc = StringSrc.AccKeyIFR
			Else
				StringSrc.AccKey = "(" & StringSrc.AccKeyIFR & UCase(StringSrc.AccKeyKey) & ")"
				KeySrc = StringSrc.AccKey
			End If
		End If

		'确定快捷键是否被移动
		Stemp = False
		If StringSrc.AccKey <> "" Then
			i = InStr(StringSrc.Shortcut,StringSrc.AccKeyIFR)
			j = InStr(StringTrn.Shortcut,StringTrn.AccKeyIFR)
			x = InStr(StringSrc.EndString,StringSrc.AcckeyIFR)
			y = InStr(StringTrn.EndString,StringTrn.AcckeyIFR)
			If LCase(StringSrc.AccKey) = LCase(StringTrn.AccKey) Then
				If AccKeyInShort = 1 Then
					If i > 0 And j = 0 Then MoveAcckey = "ShortcutSrc"
					If x > 0 And y = 0 Then MoveAcckey = "EndStringSrc"
					If i = 0 And j > 0 Then MoveAcckey = "ShortcutTrn"
					If x = 0 And y > 0 Then MoveAcckey = "EndStringTrn"
				Else
					If j > 0 Then MoveAcckey = "ShortcutTrn"
					If y > 0 Then MoveAcckey = "EndStringTrn"
				End If
			Else
				i = InStr(ShortcutSrcBak,StringSrc.AccKeyIFR)
				x = InStr(EndStringSrcBak,StringSrc.AccKeyIFR)
				If i = 0 And j > 0 Then Stemp = True
				If x = 0 And y > 0 Then Stemp = True
			End If
		End If

		'移动或删除快捷键前的终止符
		If StringTrn.PreString <> "" And AsiaKey = 1 Then
			If StringSrc.EndString & StringTrn.EndString = "" Then StringSrc.EndString = StringTrn.PreString
			StringTrn.PreString = ""
		End If

		'删除所有多余空格
		If DeleteExtraSpace = 1 Then
			If StringTrn.Spaces <> "" Then
				If AsiaKey = 0 Then
					If Len(StringTrn.Spaces) > 1 Then StringTrn.Spaces = Space(1)
				Else
					StringTrn.Spaces = ""
				End If
			End If
			If StringTrn.PreString <> "" Then StringTrn.PreString = Trim(StringTrn.PreString)
			If StringTrn.ExpString <> "" Then StringTrn.ExpString = Trim(StringTrn.ExpString)
			If StringSrc.Shortcut <> "" Then StringSrc.Shortcut = Trim(StringSrc.Shortcut)
			If StringSrc.EndString <> "" Then
				If StringSrc.EndString = Space(1) & LTrim(StringSrc.EndString) Then
					StringSrc.EndString = RTrim(StringSrc.EndString)
				Else
					StringSrc.EndString = Trim(StringSrc.EndString)
				End If
			End If
		End If

		'确定快捷键的方式
		If StringSrc.AccKey <> "" And AccKeyInShort = 1 Then
			If InStr(StringSrc.EndString & StringSrc.Shortcut,StringSrc.AccKeyIFR) Then
				If Stemp = False Then
					StringSrc.AccKey = StringSrc.AccKeyIFR & StringSrc.AccKeyKey
					KeySrc = ""
				End If
			End If
		End If
		If StringSrc.AccKey = "" Or KeySrc <> "" Then
			If InStr(StringSrc.Shortcut,StringSrc.AccKeyIFR) Then StringSrc.Shortcut = Replace(StringSrc.Shortcut,StringSrc.AccKeyIFR,"")
			If InStr(StringSrc.EndString,StringSrc.AccKeyIFR) Then StringSrc.EndString = Replace(StringSrc.EndString,StringSrc.AccKeyIFR,"")
		End If

		'自动翻译符合条件的终止符
		If StringSrc.EndString <> "" And TranslateEndChar = 1 And AutoTrnEndChar <> "" Then
			Temp = Replace(Trim(StringSrc.EndString),StringSrc.AccKeyIFR,"")
			If Trim(Temp) <> "" Then
				For i = 0 To UBound(AutoTrnEndCharArr)
					FindStr = Trim(AutoTrnEndCharArr(i))
					If InStr(FindStr,"|") Then
						TempArray = Split(FindStr,"|")
						If Temp = TempArray(0) Then
							StringSrc.EndString = Replace(StringSrc.EndString,Temp,TempArray(1))
							Exit For
						End If
					End If
				Next i
			End If
		End If

		'查找快捷键字符并设置快捷键
		If StringSrc.AccKey <> "" And KeySrc <> "" And AsiaKey = 0 Then
			If LCase(StringSrc.AccKey) <> LCase(StringTrn.AccKey) Or MoveAcckey <> "" Then
				For i = 0 To 3
					Temp = ""
					If AccKeyInShort = 0 Then
						If i = 0 Then Temp = LastStringTrn
						If i = 1 Then Temp = StringTrn.ExpString
					Else
						If i = 0 Then Temp = LastStringTrn
						If i = 1 Then Temp = StringTrn.ExpString
						If i = 2 Then Temp = StringSrc.Shortcut
						If i = 3 Then Temp = StringSrc.EndString
					End If
					If Trim(Temp) <> "" Then
						StringTrn.AccKeyPos = InStr(Temp,StringSrc.AccKeyKey)
						If StringTrn.AccKeyPos = 0 Then StringTrn.AccKeyPos = InStr(LCase(Temp),LCase(StringSrc.AccKeyKey))
						If StringTrn.AccKeyPos > 0 Then
							StringTrn.AccKeyKey = Mid(Temp,StringTrn.AccKeyPos,1)
							Temp = Replace(Temp,StringTrn.AccKeyKey,StringSrc.AccKeyIFR & StringTrn.AccKeyKey,,1)
							StringSrc.AccKey = StringSrc.AccKeyIFR & StringTrn.AccKeyKey
							KeySrc = ""
						End If
					End If
					If AccKeyInShort = 0 Then
						If i = 0 Then LastStringTrn = Temp
						If i = 1 Then StringTrn.ExpString = Temp
					Else
						If i = 0 Then LastStringTrn = Temp
						If i = 1 Then StringTrn.ExpString = Temp
						If i = 2 Then StringSrc.Shortcut = Temp
						If i = 3 Then StringSrc.EndString = Temp
					End If
					If KeySrc = "" Then Exit For
				Next i
				If KeySrc <> "" Then
					If AddAccessKeyWithFirstChar = 1 Then
						i = 0
						If Trim(LastStringTrn) <> "" Then
							If CheckStr(LastStringTrn,"-1,48-57,65-90,97-122,128-",i) = True Then
								PreTrn = Left(LastStringTrn,i - 1)
								AppTrn = Mid(LastStringTrn,i)
								StringTrn.AccKeyKey = Mid(LastStringTrn,i,1)
								LastStringTrn = PreTrn & StringSrc.AccKeyIFR & AppTrn
								StringSrc.AccKey = StringSrc.AccKeyIFR & StringTrn.AccKeyKey
								MoveAcckey = ""
								KeySrc = ""
							End If
						ElseIf Trim(StringTrn.ExpString) <> "" Then
							If CheckStr(StringTrn.ExpString,"-1,48-57,65-90,97-122,128-",i) = True Then
								PreTrn = Left(StringTrn.ExpString,i - 1)
								AppTrn = Mid(StringTrn.ExpString,i)
								StringTrn.AccKeyKey = Mid(StringTrn.ExpString,i,1)
								StringTrn.ExpString = PreTrn & StringSrc.AccKeyIFR & AppTrn
								StringSrc.AccKey = StringSrc.AccKeyIFR & StringTrn.AccKeyKey
								MoveAcckey = ""
								KeySrc = ""
							End If
						End If
					Else
						MoveAcckey = ""
						StringSrc.AccKey = ""
						KeySrc = ""
					End If
				End If
			Else
				StringTrn.AccKey = StringSrc.AccKey
			End If
		End If

		'组织替换字符
		If AsiaKey = 0 Then
			NewStringTrn = StringSrc.PreSpace & LastStringTrn & StringTrn.PreString & StringTrn.Spaces & KeySrc & _
							StringTrn.ExpString & StringSrc.EndString & StringSrc.Shortcut & StringSrc.EndSpace
		Else
			NewStringTrn = StringSrc.PreSpace & LastStringTrn & StringTrn.PreString & StringTrn.Spaces & _
							StringTrn.ExpString & KeySrc & StringSrc.EndString & StringSrc.Shortcut & StringSrc.EndSpace
		End If

		'字串替换
		If StringReplace <> NewStringTrn Then StringReplace = NewStringTrn
	End If

	'还原参数
	StringTrn.Spaces = SpaceTrnBak
	StringTrn.ExpString = ExpStringTrnBak
	StringTrn.PreString = PreStringTrnBak

	'删除终止符和加速器中的快捷键，以便可以正确比较终止符和加速器
	If InStr(StringSrc.Shortcut,StringSrc.AccKeyIFR) Then StringSrc.Shortcut = Replace(StringSrc.Shortcut,StringSrc.AccKeyIFR,"")
	If InStr(StringSrc.Shortcut,StringTrn.AccKeyIFR) Then StringSrc.Shortcut = Replace(StringSrc.Shortcut,StringTrn.AccKeyIFR,"")
	If InStr(StringTrn.Shortcut,StringSrc.AccKeyIFR) Then StringTrn.Shortcut = Replace(StringTrn.Shortcut,StringSrc.AcckeyIFR,"")
	If InStr(StringTrn.Shortcut,StringTrn.AccKeyIFR) Then StringTrn.Shortcut = Replace(StringTrn.Shortcut,StringTrn.AccKeyIFR,"")
	If InStr(StringSrc.EndString,StringSrc.AccKeyIFR) Then StringSrc.EndString = Replace(StringSrc.EndString,StringSrc.AccKeyIFR,"")
	If InStr(StringSrc.EndString,StringTrn.AccKeyIFR) Then StringSrc.EndString = Replace(StringSrc.EndString,StringTrn.AccKeyIFR,"")
	If InStr(StringTrn.EndString,acckeyIFRSrc) Then StringTrn.EndString = Replace(StringTrn.EndString,StringSrc.AccKeyIFR,"")
	If InStr(StringTrn.EndString,acckeyIFRTrn) Then StringTrn.EndString = Replace(StringTrn.EndString,StringTrn.AccKeyIFR,"")
	If InStr(MoveAcckey,StringSrc.AccKeyIFR) Then MoveAcckey = Replace(MoveAcckey,StringSrc.AccKeyIFR,"")
	If InStr(MoveAcckey,StringTrn.AccKeyIFR) Then MoveAcckey = Replace(MoveAcckey,StringTrn.AccKeyIFR,"")
End Function


'输出程序错误消息
Sub sysErrorMassage(sysError As ErrObject,fType As Long)
	Dim TempArray() As String,MsgList() As String
	Dim ErrorNumber As Long,ErrorSource As String,ErrorDescription As String
	Dim TitleMsg As String,ContinueMsg As String,Msg As String

	ErrorNumber = sysError.Number
	ErrorSource = sysError.Source
	ErrorDescription = sysError.Description

	TitleMsg = "Error"
	If fType = 0 Then
		ContinueMsg = vbCrLf & vbCrLf & "The program cannot continue and will exit."
	ElseIf fType = 1 Then
		ContinueMsg = vbCrLf & vbCrLf & "Do you want to continue?"
	ElseIf fType = 2 Then
		ContinueMsg = vbCrLf & vbCrLf & "The program will continue to run."
	End If

	If Join(UILangList,"") <> "" Then
		ItemList$ = "sysErrorMassage"
		If getMsgList(UILangList,MsgList,ItemList$,3) = False Then
			If getMsgList(UILangList,MsgList,"Main",3) = False Then
				Msg = "The following file is missing [Main|sysErrorMassage] section." & vbCrLf & "%s"
				Msg = Replace(Msg,"%s",LangFile)
			Else
				TitleMsg = MsgList(42)
				If fType <> 0 Then ContinueMsg = MsgList(94) Else ContinueMsg = MsgList(95)
				Msg = Replace(Replace(MsgList(75),"%s",ItemList$),"%d",LangFile)
			End If
		Else
			TitleMsg = MsgList(0)
			If fType = 0 Then ContinueMsg = MsgList(10)
			If fType = 1 Then ContinueMsg = MsgList(11)
			If fType = 2 Then ContinueMsg = MsgList(12)

			If ErrorSource = "" Then
				Msg = Replace(Replace(MsgList(1),"%d",CStr(ErrorNumber)),"%v",ErrorDescription)
			ElseIf ErrorSource = "NotSection" Then
				TempArray = Split(ErrorDescription,JoinStr,-1)
				Msg = Replace(Replace(MsgList(3),"%s",TempArray(1)),"%d",TempArray(0))
			ElseIf ErrorSource = "NotValue" Then
				TempArray = Split(ErrorDescription,JoinStr,-1)
				Msg = Replace(Replace(MsgList(4),"%s",TempArray(1)),"%d",TempArray(0))
			ElseIf ErrorSource = "NotReadFile" Then
				TempArray = Split(ErrorDescription,JoinStr,-1)
				Msg = Replace(MsgList(5),"%s",TempArray(1))
			ElseIf ErrorSource = "NotWriteFile" Then
				TempArray = Split(ErrorDescription,JoinStr,-1)
				Msg = Replace(MsgList(6),"%s",TempArray(1))
			ElseIf ErrorSource = "NotINIFile" Then
				Msg = Replace(MsgList(7),"%s",ErrorDescription)
			ElseIf ErrorSource = "NotExitFile" Then
				Msg = Replace(MsgList(8),"%s",ErrorDescription)
			ElseIf ErrorSource = "NotVersion" Then
				TempArray = Split(ErrorDescription,JoinStr,-1)
				Msg = Replace(MsgList(9),"%s",TempArray(0))
				Msg = Replace(Replace(Msg,"%d",TempArray(1)),"%v",TempArray(2))
			Else
				Msg = Replace(MsgList(2),"%s",ErrorSource)
				Msg = Replace(Replace(Msg,"%d",CStr(ErrorNumber)),"%v",ErrorDescription)
			End If
		End If
	Else
		If ErrorSource = "" Then
			Msg = "An Error occurred in the program design." & vbCrLf & "Error Code: %d, Content: %v"
			Msg = Replace(Replace(Msg,"%s",CStr(ErrorNumber)),"%v",ErrorDescription)
		ElseIf ErrorSource = "NotSection" Then
			TempArray = Split(ErrorDescription,JoinStr,-1)
			Msg = "The following file is missing [%s] section." & vbCrLf & "%d"
			Msg = Replace(Replace(Msg,"%s",TempArray(1)),"%d",TempArray(0))
		ElseIf ErrorSource = "NotValue" Then
			TempArray = Split(ErrorDescription,JoinStr,-1)
			Msg = "The following file is missing [%s] Value." & vbCrLf & "%d"
			Msg = Replace(Replace(Msg,"%s",TempArray(1)),"%d",TempArray(0))
		ElseIf ErrorSource = "NotReadFile" Then
			Msg = Replace(ErrorDescription,JoinStr,vbCrLf)
		ElseIf ErrorSource = "NotWriteFile" Then
			Msg = Replace(ErrorDescription,JoinStr,vbCrLf)
		ElseIf ErrorSource = "NotINIFile" Then
			Msg = "The following contents of the file is not correct." & vbCrLf & "%s"
			Msg = Replace(Msg,"%s",ErrorDescription)
		ElseIf ErrorSource = "NotExitFile" Then
			Msg = "The following file does not exist! Please check and try again." & vbCrLf & "%s"
			Msg = Replace(Msg,"%s",ErrorDescription)
		ElseIf ErrorSource = "NotVersion" Then
			TempArray = Split(ErrorDescription,JoinStr,-1)
			Msg = "The following file version is %d, requires version at least %v." & vbCrLf & "%s"
			Msg = Replace(Msg,"%s",TempArray(0))
			Msg = Replace(Replace(Msg,"%d",TempArray(1)),"%v",TempArray(2))
		Else
			Msg = "Your system is missing %s server." & vbCrLf & "Error Code: %d, Content: %v"
			Msg = Replace(Msg,"%s",ErrorSource)
			Msg = Replace(Replace(Msg,"%d",CStr(ErrorNumber)),"%v",ErrorDescription)
		End If
	End If

	If Msg <> "" Then
		'Msg = Msg & ContinueMsg
		If fType = 0 Then
			PSL.Output(TitleMsg & ": "  & Msg)
			Exit All
		ElseIf fType = 1 Then
			PSL.Output(TitleMsg & ": "  & Msg)
			Exit All 'Err.Raise(1,"ExitSub")
		Else
			PSL.Output(TitleMsg & ": "  & Msg)
		End If
	End If
End Sub


'进行数组合并
Function MergeArray(srcStrArr() As String,trnStrArr() As String) As Variant
	Dim i As Long,srcNum As Long,trnNum As Long
	Dim srcPassNum As Long,trnPassNum As Long
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
	If InStr(Convert,"\\") Then Convert = Replace(Convert,"\\","*a!N!d*")
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
	If InStr(Convert,"*a!N!d*") Then Convert = Replace(Convert,"*a!N!d*","\")
End Function


'转换八进制或十六进制转义符
Function ConvertB(ConverString As String) As String
	Dim i As Long,EscStr As String,ConvCode As String
	Dim ConvString As String,Stemp As Boolean
	ConvertB = ConverString
	If ConvertB = "" Then Exit Function
	i = InStr(ConvertB,"\")
	Do While i > 0
		EscStr = Mid(ConvertB,i,2)
		Stemp = False
		If EscStr = "\x" Then
			ConvCode = Mid(ConvertB,i+2,2)
			Stemp = CheckStr(UCase(ConvCode),"48-57,65-70",1)
		ElseIf EscStr = "\u" Then
			ConvCode = Mid(ConvertB,i+2,4)
			Stemp = CheckStr(UCase(ConvCode),"48-57,65-70",1)
		ElseIf EscStr = "\U" Then
			ConvCode = Mid(ConvertB,i+2,4)
			Stemp = CheckStr(UCase(ConvCode),"48-57,65-70",1)
		ElseIf EscStr <> "" Then
			EscStr = "\"
			ConvCode = Mid(ConvertB,i+1,3)
			Stemp = CheckStr(ConvCode,"48-55",1)
		End If
		If Stemp = True Then
			ConvString = ""
			If EscStr = "\x" Then ConvString = ChrW(Val("&H" & ConvCode))
			If LCase(EscStr) = "\u" Then ConvString = ChrW(Val("&H" & ConvCode))
			If EscStr = "\" Then ConvString = ChrW(Val("&O" & ConvCode))
			If ConvString <> "" Then
				ConvertB = Replace(ConvertB,EscStr & ConvCode,ConvString)
			End If
		End If
		i = InStr(i + 1,ConvertB,"\")
	Loop
End Function


'转换字符为整数数值
Function StrToLong(mStr As String) As Long
	If Trim(mStr) = "" Then mStr = "0"
	StrToLong = CLng(mStr)
End Function


'获取设置
Function EngineGet(SelSet As String,List() As String,DataList() As String,Path As String) As Long
	Dim i As Long,n As Long,j As Long,k As Long,m As Long,x As Long
	Dim Header As String,HeaderIDArr() As String,SetsArray() As String,Temp As String
	Dim LangPairList() As String,TempArray() As String,LineArray() As String

	EngineGet = 0
	NewVersion = ToUpdateEngineVersion
	ReDim SetsArray(19)

	If Path = EngineRegKey Then GoTo GetFromRegistry
	If Path = "" Then Path = EngineFilePath
	If Dir(Path) = "" Then GoTo GetFromRegistry
	On Error GoTo GetFromRegistry
	LineArray = Split(ReadFile(Path,"_autodetect_all"),vbCrLf)
	n = UBound(LineArray)
	For i = 0 To n
		l$ = LineArray(i)
		If Trim(l$) <> "" Then
			If Left(Trim(L$),1) = "[" And Right(Trim(L$),1) = "]" Then
				Header = Trim(Mid(Trim(L$),2,Len(Trim(L$))-2))
			End If
			If Header <> "" And HeaderBak = "" Then HeaderBak = Header
			If Header <> "" And Header = HeaderBak Then
				setPreStr = ""
				setAppStr = ""
				j = InStr(L$,"=")
				If j > 0 Then
					setPreStr = Trim(Left(l$,j - 1))
					setAppStr = LTrim(Mid(L$,j + 1))
				End If
				If setPreStr <> "" Then
					'获取 Option 项和值
					If setPreStr = "Version" Then OldVersion = setAppStr
					If Header = "Option" Then
						If SelSet = "" Or SelSet = "Option" Then
							If setPreStr = "TranEngineSet" Then tSelected(0) = setAppStr
							If setPreStr = "CheckSet" Then tSelected(1) = setAppStr
							If setPreStr = "TranAllType" Then tSelected(2) = setAppStr
							If setPreStr = "TranMenu" Then tSelected(3) = setAppStr
							If setPreStr = "TranDialog" Then tSelected(4) = setAppStr
							If setPreStr = "TranString" Then tSelected(5) = setAppStr
							If setPreStr = "TranAcceleratorTable" Then tSelected(6) = setAppStr
							If setPreStr = "TranVersion" Then tSelected(7) = setAppStr
							If setPreStr = "TranOther" Then tSelected(8) = setAppStr
							If setPreStr = "TranSeletedOnly" Then tSelected(9) = setAppStr
							If setPreStr = "SkipForReview" Then tSelected(10) = setAppStr
							If setPreStr = "SkipValidated" Then tSelected(11) = setAppStr
							If setPreStr = "SkipNotTran" Then tSelected(12) = setAppStr
							If setPreStr = "SkipAllNumAndSymbol" Then tSelected(13) = setAppStr
							If setPreStr = "SkipAllUCase" Then tSelected(14) = setAppStr
							If setPreStr = "SkipAllLCase" Then tSelected(15) = setAppStr
							If setPreStr = "AutoSelection" Then tSelected(16) = setAppStr
							If setPreStr = "CheckSrcProject" Then tSelected(17) = setAppStr
							If setPreStr = "CheckSrcString" Then tSelected(18) = setAppStr
							If setPreStr = "ReplaceSrcString" Then tSelected(19) = setAppStr
							If setPreStr = "SplitTranslate" Then tSelected(20) = setAppStr
							If setPreStr = "CheckTrnProject" Then tSelected(21) = setAppStr
							If setPreStr = "CheckTrnString" Then tSelected(22) = setAppStr
							If setPreStr = "ReplaceTrnString" Then tSelected(23) = setAppStr
							If setPreStr = "KeepSetting" Then tSelected(24) = setAppStr
							If setPreStr = "ShowMassage" Then tSelected(25) = setAppStr
							If setPreStr = "AddTranComment" Then tSelected(26) = setAppStr
							If setPreStr = "UILanguageID" Then tSelected(27) = setAppStr
						End If
					'获取 Option 项外的全部项和值
					ElseIf Header <> "Update" Then
						If SelSet = "" Or SelSet = "Sets" Or SelSet = Header Then
							If setPreStr = "ObjectName" Then SetsArray(0) = setAppStr
							If setPreStr = "AppId" Then SetsArray(1) = setAppStr
							If setPreStr = "EngineUrl" Then SetsArray(2) = setAppStr
							If setPreStr = "UrlTemplate" Then SetsArray(3) = setAppStr
							If setPreStr = "Method" Then SetsArray(4) = setAppStr
							If setPreStr = "Async" Then SetsArray(5) = setAppStr
							If setPreStr = "User" Then SetsArray(6) = setAppStr
							If setPreStr = "Password" Then SetsArray(7) = setAppStr
							If setPreStr = "SendBody" Then SetsArray(8) = setAppStr
							If setPreStr = "RequestHeader" Then SetsArray(9) = Convert(setAppStr)
							If setPreStr = "ResponseType" Then SetsArray(10) = setAppStr
							If setPreStr = "TranBeforeStrByText" Then SetsArray(11) = setAppStr
							If setPreStr = "TranAfterStrByText" Then SetsArray(12) = setAppStr
							If setPreStr = "TranBeforeStrByBody" Then SetsArray(13) = setAppStr
							If setPreStr = "TranAfterStrByBody" Then SetsArray(14) = setAppStr
							If setPreStr = "TranBeforeStrByStream" Then SetsArray(15) = setAppStr
							If setPreStr = "TranAfterStrByStream" Then SetsArray(16) = setAppStr
							If setPreStr = "TranXMLIdName" Then SetsArray(17) = setAppStr
							If setPreStr = "TranXMLTagName" Then SetsArray(18) = setAppStr
							If setPreStr = "Enable" Then SetsArray(19) = setAppStr
							If setPreStr = "LangCodePair" Then LngPair = setAppStr
							If setPreStr = "TranBeforeStr" Then bStr = setAppStr
							If setPreStr = "TranAfterStr" Then aStr = setAppStr
						End If
					End If
				End If
			End If
		End If
		If Header <> "" And (i = n Or Header <> HeaderBak) Then
			If SelSet = "Option" And HeaderBak = "Option" Then
				If Join(tSelected,"") <> "" Then EngineGet = 1
				Exit For
			ElseIf HeaderBak <> "Option" And HeaderBak <> "Update" Then
				If SelSet = "" Or SelSet = "Sets" Or SelSet = HeaderBak Then
					If SetsArray(10) = "responseXML" Then
						If SetsArray(17) = "" Then SetsArray(17) = bStr
						If SetsArray(18) = "" Then SetsArray(18) = aStr
					Else
						If SetsArray(11) = "" Then SetsArray(11) = bStr
						If SetsArray(12) = "" Then SetsArray(12) = aStr
						If SetsArray(13) = "" Then SetsArray(13) = bStr
						If SetsArray(14) = "" Then SetsArray(14) = aStr
						If SetsArray(15) = "" Then SetsArray(15) = bStr
						If SetsArray(16) = "" Then SetsArray(16) = aStr
					End If
					If LngPair <> "" Then
						If CheckNullData("",SetsArray,"1,6-9,15-19",6) = False Then
							Data = HeaderBak & JoinStr & Join(SetsArray,SubJoinStr) & JoinStr & LngPair
							'更新旧版的默认配置值
							If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 Then
								Data = EngineDataUpdate(HeaderBak,Data)
							End If
							'保存数据到数组中
							CreateArray(HeaderBak,Data,List,DataList)
							x = x + 1
						End If
					End If
					'数据初始化
					ReDim SetsArray(19)
					LngPair = ""
					bStr = ""
					aStr = ""
					m = m + 1
					If i = n And x = m Then EngineGet = 4
				End If
			End If
			HeaderBak = Header
		End If
	Next i
	On Error GoTo 0
	If EngineGet = 0 Then GoTo GetFromRegistry
	Exit Function

	GetFromRegistry:
	If tWriteLoc = "" Then tWriteLoc = EngineRegKey
	ReDim SetsArray(19)
	'获取 Option 项和值
	OldVersion = GetSetting("WebTranslate","Option","Version","")
	If SelSet = "" Or SelSet = "Option" Then
		tSelected(0) = GetSetting("WebTranslate","Option","TranEngineSet","")
		tSelected(1) = GetSetting("WebTranslate","Option","CheckSet","")
		tSelected(2) = GetSetting("WebTranslate","Option","TranAllType",0)
		tSelected(3) = GetSetting("WebTranslate","Option","TranMenu",0)
		tSelected(4) = GetSetting("WebTranslate","Option","TranDialog",0)
		tSelected(5) = GetSetting("WebTranslate","Option","TranString",0)
		tSelected(6) = GetSetting("WebTranslate","Option","TranAcceleratorTable",0)
		tSelected(7) = GetSetting("WebTranslate","Option","TranVersion",0)
		tSelected(8) = GetSetting("WebTranslate","Option","TranOther",0)
		tSelected(9) = GetSetting("WebTranslate","Option","TranSeletedOnly",1)
		tSelected(10) = GetSetting("WebTranslate","Option","SkipForReview",1)
		tSelected(11) = GetSetting("WebTranslate","Option","SkipValidated",1)
		tSelected(12) = GetSetting("WebTranslate","Option","SkipNotTran",0)
		tSelected(13) = GetSetting("WebTranslate","Option","SkipAllNumAndSymbol",1)
		tSelected(14) = GetSetting("WebTranslate","Option","SkipAllUCase",1)
		tSelected(15) = GetSetting("WebTranslate","Option","SkipAllLCase",0)
		tSelected(16) = GetSetting("WebTranslate","Option","AutoSelection",1)
		tSelected(17) = GetSetting("WebTranslate","Option","CheckSrcProject",4)
		tSelected(18) = GetSetting("WebTranslate","Option","CheckSrcString",1)
		tSelected(19) = GetSetting("WebTranslate","Option","ReplaceSrcString",1)
		tSelected(20) = GetSetting("WebTranslate","Option","SplitTranslate",1)
		tSelected(21) = GetSetting("WebTranslate","Option","CheckTrnProject",1)
		tSelected(22) = GetSetting("WebTranslate","Option","CheckTrnString",1)
		tSelected(23) = GetSetting("WebTranslate","Option","ReplaceTrnString",1)
		tSelected(24) = GetSetting("WebTranslate","Option","KeepSetting",1)
		tSelected(25) = GetSetting("WebTranslate","Option","ShowMassage",1)
		tSelected(26) = GetSetting("WebTranslate","Option","AddTranComment",0)
		tSelected(27) = GetSetting("WebTranslate","Option","UILanguageID",0)
		If SelSet = "Option" Then
			If Join(tSelected,"") <> "" Then EngineGet = 1
			Exit Function
		End If
	End If
	'获取 Option 外的项和值
	If SelSet <> "Option" And SelSet <> "Update" Then
		m = 0
		x = 0
		Header = GetSetting("WebTranslate","Option","Headers","")
		If Header <> "" Then
			HeaderIDArr = Split(Header,";",-1)
			n = UBound(HeaderIDArr)
			For i = 0 To n
				HeaderID = HeaderIDArr(i)
				If HeaderID <> "" Then
					'转存旧版的每个项和值
					Header = GetSetting("WebTranslate",HeaderID,"Name","")
				End If
				If Header = "" Then Header = HeaderID
				If SelSet = "" Or SelSet = "Sets" Or SelSet = Header Then
					SetsArray(0) = GetSetting("WebTranslate",HeaderID,"ObjectName","")
					SetsArray(1) = GetSetting("WebTranslate",HeaderID,"AppId","")
					SetsArray(2) = GetSetting("WebTranslate",HeaderID,"EngineUrl","")
					SetsArray(3) = GetSetting("WebTranslate",HeaderID,"UrlTemplate","")
					SetsArray(4) = GetSetting("WebTranslate",HeaderID,"Method","")
					SetsArray(5) = GetSetting("WebTranslate",HeaderID,"Async","")
					SetsArray(6) = GetSetting("WebTranslate",HeaderID,"User","")
					SetsArray(7) = GetSetting("WebTranslate",HeaderID,"Password","")
					SetsArray(8) = GetSetting("WebTranslate",HeaderID,"SendBody","")
					SetsArray(9) = Convert(GetSetting("WebTranslate",HeaderID,"RequestHeader",""))
					SetsArray(10) = GetSetting("WebTranslate",HeaderID,"ResponseType","")
					SetsArray(11) = GetSetting("WebTranslate",HeaderID,"TranBeforeStrByText","")
					SetsArray(12) = GetSetting("WebTranslate",HeaderID,"TranAfterStrByText","")
					SetsArray(13) = GetSetting("WebTranslate",HeaderID,"TranBeforeStrByBody","")
					SetsArray(14) = GetSetting("WebTranslate",HeaderID,"TranAfterStrByBody","")
					SetsArray(15) = GetSetting("WebTranslate",HeaderID,"TranBeforeStrByStream","")
					SetsArray(16) = GetSetting("WebTranslate",HeaderID,"TranAfterStrByStream","")
					SetsArray(17) = GetSetting("WebTranslate",HeaderID,"TranXMLIdName","")
					SetsArray(18) = GetSetting("WebTranslate",HeaderID,"TranXMLTagName","")
					SetsArray(19) = GetSetting("WebTranslate",HeaderID,"Enable","1")
					LngPair = GetSetting("WebTranslate",HeaderID,"LangCodePair","")
					bStr = GetSetting("WebTranslate",HeaderID,"TranBeforeStr","")
					aStr = GetSetting("WebTranslate",HeaderID,"TranAfterStr","")
					If SetsArray(10) = "responseXML" Then
						If SetsArray(17) = "" Then SetsArray(17) = bStr
						If SetsArray(18) = "" Then SetsArray(18) = aStr
					Else
						If SetsArray(11) = "" Then SetsArray(11) = bStr
						If SetsArray(12) = "" Then SetsArray(12) = aStr
						If SetsArray(13) = "" Then SetsArray(13) = bStr
						If SetsArray(14) = "" Then SetsArray(14) = aStr
						If SetsArray(15) = "" Then SetsArray(15) = bStr
						If SetsArray(16) = "" Then SetsArray(16) = aStr
					End If
					If LngPair <> "" Then
						If CheckNullData("",SetsArray,"1,6-9,15-19",6) = False Then
							Data = Header & JoinStr & Join(SetsArray,SubJoinStr) & JoinStr & LngPair
							'更新旧版的默认配置值
							If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 Then
								Data = EngineDataUpdate(Header,Data)
							End If
							'保存数据到数组中
							CreateArray(Header,Data,List,DataList)
							x = x + 1
						End If
					End If
					'数据初始化
					ReDim SetsArray(19)
					LngPair = ""
					bStr = ""
					aStr = ""
					m = m + 1
					If i = n And x = m Then EngineGet = 4
				End If
			Next i
		End If
	End If
End Function


'获取字串检查设置
Function CheckGet(SelSet As String,DataList() As String,Path As String,Lang As String) As Long
	Dim i As Long,n As Long,j As Long,k As Long,Header As String,HeaderIDArr() As String
	Dim TempArray() As String,LineArray() As String,SetsArray() As String,Temp As String

	CheckGet = 0
	NewVersion = ToUpdateCheckVersion
	ReDim SetsArray(17)
	If SelSet = DefaultCheckList(0) Then SelSet = "en2zh"
	If SelSet = DefaultCheckList(1) Then SelSet = "zh2en"

	If Path = CheckRegKey Then GoTo GetFromRegistry
	If Path = "" Then Path = CheckFilePath
	If Dir(Path) = "" Then GoTo GetFromRegistry
	On Error GoTo GetFromRegistry
	LineArray = Split(ReadFile(Path,"_autodetect_all"),vbCrLf)
	n = UBound(LineArray)
	For i = 0 To n
		l$ = LineArray(i)
		If Trim(l$) <> "" Then
			If Left(Trim(L$),1) = "[" And Right(Trim(L$),1) = "]" Then
				Header = Trim(Mid(Trim(l$),2,Len(Trim(l$))-2))
			End If
			If Header <> "" And HeaderBak = "" Then HeaderBak = Header
			If Header <> "" And Header = HeaderBak Then
				setPreStr = ""
				setAppStr = ""
				j = InStr(L$,"=")
				If j > 0 Then
					setPreStr = Trim(Left(L$,j - 1))
					setAppStr = LTrim(Mid(L$,j + 1))
				End If
				If setPreStr <> "" Then
					'获取 Option 项和值
					If setPreStr = "Version" Then CheckVersion = setAppStr
					'获取 Project 项和值
					If Header = "Projects" Then
						If SelSet = "" Or SelSet = "Project" Then
							If setAppStr <> "" Then
								If setPreStr = "CheckOnly" Then setPreStr = DefaultProjectList(0)
								If setPreStr = "CheckAndCorrect" Then setPreStr = DefaultProjectList(1)
								If setPreStr = "DelAccessKey" Then setPreStr = DefaultProjectList(2)
								If setPreStr = "DelAccelerator" Then setPreStr = DefaultProjectList(3)
								If setPreStr = "DelAccessKeyAndAccelerator" Then setPreStr = DefaultProjectList(4)
								ReDim Preserve ProjectDataList(k)
								ProjectDataList(k) = setPreStr & JoinStr & setAppStr
								k = k + 1
							End If
						End If
					'获取 Option 项外的全部项和值
					ElseIf Header <> "Option" And Header <> "Update" Then
						If SelSet = "" Or SelSet = "Sets" Or SelSet = Header Then
							If setPreStr = "ExcludeChar" Then SetsArray(0) = setAppStr
							If setPreStr = "LineSplitChar" Then SetsArray(1) = setAppStr
							If setPreStr = "CheckBracket" Then SetsArray(2) = setAppStr
							If setPreStr = "KeepCharPair" Then SetsArray(3) = setAppStr
							If setPreStr = "ShowAsiaKey" Then SetsArray(4) = setAppStr
							If setPreStr = "CheckEndChar" Then SetsArray(5) = setAppStr
							If setPreStr = "NoTrnEndChar" Then SetsArray(6) = setAppStr
							If setPreStr = "AutoTrnEndChar" Then SetsArray(7) = setAppStr
							If setPreStr = "CheckShortChar" Then SetsArray(8) = setAppStr
							If setPreStr = "CheckShortKey" Then SetsArray(9) = setAppStr
							If setPreStr = "KeepShortKey" Then SetsArray(10) = setAppStr
							If setPreStr = "PreRepString" Then SetsArray(11) = setAppStr
							If setPreStr = "AutoRepString" Then SetsArray(12) = setAppStr
							If setPreStr = "AccessKeyChar" Then SetsArray(13) = setAppStr
							If setPreStr = "AddAccessKeyWithFirstChar" Then SetsArray(14) = setAppStr
							If setPreStr = "LineSplitMode" Then SetsArray(15) = setAppStr
							If setPreStr = "AppInsertSplitChar" Then SetsArray(16) = setAppStr
							If setPreStr = "ReplaceSplitChar" Then SetsArray(17) = setAppStr
							If setPreStr = "ApplyLangList" Then LngPair = setAppStr
						End If
					End If
				End If
			End If
		End If
		If Header <> "" And (i = n Or Header <> HeaderBak) Then
			If SelSet = "Option" And HeaderBak = "Option" Then
				If Join(cSelected,"") <> "" Then CheckGet = 1
				Exit For
			ElseIf SelSet = "Project" And HeaderBak = "Projects" Then
				If k > 0 Then CheckGet = 3
				Exit For
			ElseIf HeaderBak <> "Option" And HeaderBak <> "Update" And HeaderBak <> "Projects" Then
				If SelSet = "" And tSelected(1) = "" And tSelected(16) = "" Then tSelected(16) = "1"
				If Lang <> "" And tSelected(16) = "1" Then
					If getCheckID(LngPair,Lang) = True Then SelSet = HeaderBak
				End If
				If HeaderBak = SelSet Or (Lang = "" And HeaderBak = tSelected(1)) Then
					Temp = Join(SetsArray,"")
					If Temp <> "" And Temp <> "0" And Temp <> "1" Then
						If HeaderBak = "en2zh" Then HeaderBak = DefaultCheckList(0)
						If HeaderBak = "zh2en" Then HeaderBak = DefaultCheckList(1)
						Data = HeaderBak & JoinStr & Join(SetsArray,SubJoinStr) & JoinStr & LngPair
						'更新旧版的默认配置值
						If CheckVersion <> "" And StrComp(NewVersion,CheckVersion) = 1 Then
							Data = CheckDataUpdate(HeaderBak,Data)
						End If
						DataList(0) = Data
						CheckGet = 4
						Exit For
					End If
					'数据初始化
					ReDim SetsArray(17)
					LngPair = ""
				End If
			End If
			HeaderBak = Header
		End If
	Next i
	On Error GoTo 0
	If CheckGet = 0 Then GoTo GetFromRegistry
	Exit Function

	GetFromRegistry:
	ReDim SetsArray(17)
	'获取 Option 项和值
	CheckVersion = GetSetting("AccessKey","Option","Version","")
	'获取 Project 项和值
	If SelSet = "" Or SelSet = "Project" Then
		k = 0
		On Error GoTo NextItem
		TempArray = GetAllSettings("AccessKey","Projects")
		For i = LBound(TempArray) To UBound(TempArray)
			setPreStr = TempArray(i,0)
			setAppStr = TempArray(i,1)
			If setPreStr <> "" And setAppStr <> "" Then
				If setPreStr = "CheckOnly" Then setPreStr = DefaultProjectList(0)
				If setPreStr = "CheckAndCorrect" Then setPreStr = DefaultProjectList(1)
				If setPreStr = "DelAccessKey" Then setPreStr = DefaultProjectList(2)
				If setPreStr = "DelAccelerator" Then setPreStr = DefaultProjectList(3)
				If setPreStr = "DelAccessKeyAndAccelerator" Then setPreStr = DefaultProjectList(4)
				ReDim Preserve ProjectDataList(k)
				ProjectDataList(k) = setPreStr & JoinStr & setAppStr
				k = k + 1
			End If
		Next i
		On Error GoTo 0
		NextItem:
		If SelSet = "Project" Then
			If k > 0 Then CheckGet = 3
			Exit Function
		End If
	End If
	'获取 Option 外的项和值
	If SelSet <> "Option" And SelSet <> "Update" And SelSet <> "Projects" Then
		If SelSet = "" And tSelected(1) = "" And tSelected(16) = "" Then tSelected(16) = "1"
		Header = GetSetting("AccessKey","Option","Headers","")
		If Header <> "" Then
			HeaderIDArr = Split(Header,";",-1)
			n = UBound(HeaderIDArr)
			For i = 0 To n
				HeaderID = HeaderIDArr(i)
				If HeaderID <> "" Then
					'转存旧版的每个项和值
					Header = GetSetting("AccessKey",HeaderID,"Name","")
				End If
				If Header = "" Then Header = HeaderID
				If Lang <> "" And tSelected(16) = "1" Then
					LngPair = GetSetting("AccessKey",HeaderID,"ApplyLangList","")
					If getCheckID(LngPair,Lang) = True Then SelSet = Header
				End If
				If Header = SelSet Or (Lang = "" And Header = tSelected(1)) Then
					SetsArray(0) = GetSetting("AccessKey",HeaderID,"ExcludeChar","")
					SetsArray(1) = GetSetting("AccessKey",HeaderID,"LineSplitChar","")
					SetsArray(2) = GetSetting("AccessKey",HeaderID,"CheckBracket","")
					SetsArray(3) = GetSetting("AccessKey",HeaderID,"KeepCharPair","")
					SetsArray(4) = GetSetting("AccessKey",HeaderID,"ShowAsiaKey","")
					SetsArray(5) = GetSetting("AccessKey",HeaderID,"CheckEndChar","")
					SetsArray(6) = GetSetting("AccessKey",HeaderID,"NoTrnEndChar","")
					SetsArray(7) = GetSetting("AccessKey",HeaderID,"AutoTrnEndChar","")
					SetsArray(8) = GetSetting("AccessKey",HeaderID,"CheckShortChar","")
					SetsArray(9) = GetSetting("AccessKey",HeaderID,"CheckShortKey","")
					SetsArray(10) = GetSetting("AccessKey",HeaderID,"KeepShortKey","")
					SetsArray(11) = GetSetting("AccessKey",HeaderID,"PreRepString","")
					SetsArray(12) = GetSetting("AccessKey",HeaderID,"AutoRepString","")
					SetsArray(13) = GetSetting("AccessKey",HeaderID,"AccessKeyChar","")
					SetsArray(14) = GetSetting("AccessKey",HeaderID,"AddAccessKeyWithFirstChar","")
					SetsArray(15) = GetSetting("AccessKey",HeaderID,"LineSplitMode","")
					SetsArray(16) = GetSetting("AccessKey",HeaderID,"AppInsertSplitChar","")
					SetsArray(17) = GetSetting("AccessKey",HeaderID,"ReplaceSplitChar","")
					LngPair = GetSetting("AccessKey",HeaderID,"ApplyLangList","")
					Temp = Join(SetsArray,"")
					If Temp <> "" And Temp <> "0" And Temp <> "1" Then
						If Header = "en2zh" Then Header = DefaultCheckList(0)
						If Header = "zh2en" Then Header = DefaultCheckList(1)
						Data = Header & JoinStr & Join(SetsArray,SubJoinStr) & JoinStr & LngPair
						'更新旧版的默认配置值
						If CheckVersion <> "" And StrComp(NewVersion,CheckVersion) = 1 Then
							Data = CheckDataUpdate(Header,Data)
						End If
						DataList(0) = Data
						CheckGet = 4
						Exit For
					End If
					'数据初始化
					ReDim SetsArray(17)
					LngPair = ""
				End If
			Next i
		End If
	End If
End Function


'更新引擎旧版本配置值
Function EngineDataUpdate(Header As String,Data As String) As String
	Dim i As Long,UpdatedData As String,uV As String,dV As String,Stemp As Boolean
	EngineDataUpdate = Data
	Stemp = False
	For i = LBound(DefaultEngineList) To UBound(DefaultEngineList)
		If DefaultEngineList(i) = Header Then
			Stemp = True
			Exit For
		End If
	Next i
	If Stemp = False Then Exit Function
	TempArray = Split(Data,JoinStr)
	uSetsArray = Split(TempArray(1),SubJoinStr)
	dSetsArray = Split(EngineSettings(Header),SubJoinStr)
	For i = 0 To UBound(uSetsArray)
		uV = uSetsArray(i)
		dV = dSetsArray(i)
		If Trim(uV) = "" Then uV = dV
		If uV <> "" And uV <> dV Then uV = dV
		uSetsArray(i) = uV
	Next i
	TempArray(1) = Join(uSetsArray,SubJoinStr)
	EngineDataUpdate = Join(TempArray,JoinStr)
End Function


'更新检查旧版本配置值
Function CheckDataUpdate(Header As String,Data As String) As String
	Dim UpdatedData As String,uV As String,dV As String,spStr As String,Stemp As Boolean
	Dim i As Long,j As Long,m As Long,uDataList() As String,dDataList() As String
	CheckDataUpdate = Data
	Stemp = False
	For i = LBound(DefaultCheckList) To UBound(DefaultCheckList)
		If DefaultCheckList(i) = Header Then
			Stemp = True
			Exit For
		End If
	Next i
	If Stemp = False Then Exit Function
	dData = CheckSettings(Header,0)
	If CheckDataUpdate = dData Then Exit Function
	dSetsArray = Split(dData,SubJoinStr)
	If Join(dSetsArray,"") = "" Then Exit Function
	TempArray = Split(Data,JoinStr)
	uSetsArray = Split(TempArray(1),SubJoinStr)
	For i = 0 To UBound(uSetsArray)
		uV = uSetsArray(i)
		dV = dSetsArray(i)
		If Trim(uV) = "" Or i = 1 Or i = 6 Then uV = dV
		If i <> 4 And i <> 14 And i <> 15 And i < 18 And uV <> "" And uV <> dV Then
			If i = 5 Or i = 7 Then spStr = " " Else spStr = ","
			If i = 7 And InStr(uV,"|") = 0 Then
				uDataList = Split(uV,spStr)
				For m = 0 To UBound(uDataList)
					uV = uDataList(m)
					uDataList(m) = Left(Trim(uV),1) & "|" & Right(Trim(uV),1)
				Next m
				uV = Join(uDataList,spStr)
			End If
			uV = Join(ClearArray(Split(uV & spStr & dV,spStr,-1)),spStr)
		End If
		uSetsArray(i) = uV
	Next i
	TempArray(1) = Join(uSetsArray,SubJoinStr)
	CheckDataUpdate = Join(TempArray,JoinStr)
End Function


'增加或更改数组项目
Function CreateArray(Header As String,Data As String,HeaderList() As String,DataList() As String) As Boolean
	Dim i As Long,n As Long,FindHeader As String,Stemp As Boolean
	If HeaderList(0) = "" Then
		HeaderList(0) = Header
		DataList(0) = Data
	Else
		Stemp = False
		n = UBound(HeaderList)
		FindHeader = LCase(Header)
		For i = LBound(HeaderList) To n
			If LCase(HeaderList(i)) = FindHeader Then
				If DataList(i) <> Data Then DataList(i) = Data
				Stemp = True
				Exit For
			End If
		Next i
		If Stemp = False Then
			n = n + 1
			ReDim Preserve HeaderList(n),DataList(n)
			HeaderList(n) = Header
			If DataList(n) <> Data Then DataList(n) = Data
		End If
	End If
End Function


'查找指定值是否在数组中
Function getCheckID(Data As String,LngCode As String) As Boolean
	Dim i As Long,FindCode As String,Stemp As Boolean
	getCheckID = False
	If LngCode = "" Or Data = "" Then Exit Function
	FindCode = LCase(LngCode)
	LangArray = Split(Data,SubLngJoinStr)
	For i = 0 To UBound(LangArray)
		LangPairList = Split(LangArray(i),LngJoinStr)
		If LCase(LangPairList(1)) = FindCode Then
			getCheckID = True
			Exit For
		End If
	Next i
End Function


'检查数组中是否有空值
'ftype = 0     检查多项数组项内是否全为空值
'ftype = 1     检查多项数组项内是否有空值
'ftype = 2     仅检查多项数组的参数项内是否全为空值
'ftype = 3     检查多项数组的参数项内是否有空值
'ftype = 6     检查单项数组项内是否有空值
'Header = ""   检查整个数组
'Header <> ""  检查指定数组项
Function CheckNullData(Header As String,DataList() As String,SkipNum As String,fType As Long) As Boolean
	Dim i As Long,j As Long,x As Long,m As Long,n As Long,Stemp As Boolean,hStemp As Boolean
	CheckNullData = False
	SkipNumArray = Split(SkipNum,",")
	If InStr(SkipNum,"-") Then
		For i = 0 To UBound(SkipNumArray)
			If InStr(SkipNumArray(i),"-") Then
				Temp = ""
				TempArray = Split(SkipNumArray(i),"-")
				For j = CLng(TempArray(0)) To CLng(TempArray(1))
					If Temp <> "" Then Temp = Temp & "," & j
					If Temp = "" Then Temp = j
				Next j
				If Temp <> "" Then SkipNumArray(i) = Temp
			End If
		Next i
		SkipNum = Join(SkipNumArray,",")
		SkipNumArray = Split(SkipNum,",")
	End If
	m = 0
	hStemp = False
	dMax = UBound(DataList)
	nMax = UBound(SkipNumArray)
	For i = LBound(DataList) To dMax
		If fType = 6 Then
			Stemp = False
			For x = 0 To nMax
				If CStr(i) = SkipNumArray(x) Then
					Stemp = True
					Exit For
				End If
			Next x
			If Stemp = False Then
				If Trim(DataList(i)) = "" Then
					CheckNullData = True
					Exit For
				End If
			End If
		Else
			n = 0
			TempArray = Split(DataList(i),JoinStr)
			SetsArray = Split(TempArray(1),SubJoinStr)
			sMax = UBound(SetsArray)
			If Header <> "" And TempArray(0) = Header Then hStemp = True
			If Header = "" Then hStemp = True
			If hStemp = True Then
				If fType < 4 Then
					For j = 0 To sMax
						Stemp = False
						For x = 0 To nMax
							If CStr(j) = SkipNumArray(x) Then
								Stemp = True
								Exit For
							End If
						Next x
						If Trim(SetsArray(j)) = "" And Stemp = False Then
							If fType = 0 Or fType = 2 Then n = n + 1
							If fType = 1 Or fType = 3 Then
								CheckNullData = True
								Exit For
							End If
						End If
					Next j
				End If
			End If
			If fType = 0 Then
				If Header <> "" Then
					If n = sMax - nMax + 1 Then CheckNullData = True
				Else
					If n = sMax - nMax + 1 Then m = m + 1
					If m = dMax + 1 Then CheckNullData = True
				End If
			ElseIf fType = 2 Then
				If Header <> "" Then
					If n = sMax - nMax Then CheckNullData = True
				Else
					If n = sMax - nMax Then m = m + 1
					If m = dMax + 1 Then CheckNullData = True
				End If
			ElseIf fType = 4 Then
				If Header <> "" Then
					If n = 1 Then CheckNullData = True
				Else
					If n = dMax + 1 Then CheckNullData = True
				End If
			End If
			If CheckNullData = True Then Exit For
		End If
	Next i
	If fType <> 6 And Header <> "" And hStemp = False Then CheckNullData = True
End Function


'数组排序
Function SortArray(xArray() As String,Comp As Long,CompType As String,Operator As String) As Variant
	Dim rMin As Long,rMax As Long,MaxLng As Long,Lng As Long,yArray() As String
	Dim fLng As Long,sLng As Long,MyComp As Long,x As Long,y As Long
	SortArray = xArray
	rMin = LBound(xArray)
	rMax = UBound(xArray)
	If rMax = 0 Or CompType = "" Or Operator = "" Then Exit Function
	yArray = xArray
    MaxLng = 1
    For x = rMin To rMax
        Lng = Len(Trim(yArray(x)))
        If Lng > MaxLng Then MaxLng = Lng
    Next
	For x = rMax To rMin Step -1
		For y = rMin To rMax - 1
			fLng = Len(Trim(yArray(y)))
			sLng = Len(Trim(yArray(y+1)))
			If CompType = "Size" Then
				fValue = String(MaxLng - fLng,"0") & yArray(y)
				sValue = String(MaxLng - sLng,"0") & yArray(y+1)
				MyComp = StrComp(fValue,sValue,Comp)
			End If
			If CompType = "Lenght" Then
				If fLng < sLng Then MyComp = -1
				If fLng = sLng Then MyComp = 0
				If fLng > sLng Then MyComp = 1
			End If
			If Operator = ">" Then
				If MyComp > 0 Then
					Mx = yArray(y + 1)
					yArray(y+1) = yArray(y)
					yArray(y) = Mx
				End If
			ElseIf Operator = "<" Then
				If MyComp < 0 Then
					Mx = yArray(y + 1)
					yArray(y+1) = yArray(y)
					yArray(y) = Mx
				End If
			ElseIf Operator = "=" Then
				If MyComp = 0 Then
					Mx = yArray(y + 1)
					yArray(y+1) = yArray(y)
					yArray(y) = Mx
				End If
			End If
		Next y
	Next x
	SortArray = yArray
End Function


'清理数组中重复的数据
Function ClearArray(xArray() As String) As Variant
	Dim i As Long,j As Long,y As Long,k As Long,l As Long
	Dim yArray() As String,Stemp As Boolean
	ClearArray = xArray
	k = LBound(xArray)
	l = UBound(xArray)
	If l = 0 Then Exit Function
	y = 0
	ReDim yArray(0)
	For i = k To l
		Stemp = False
		For j = i + 1 To l
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
Function CheckKeyCode(FindKey As String,CheckKey As String) As Long
	Dim KeyCode As Boolean,FindStr As String,Key As String,Pos As Long
	Key = Trim(FindKey)
	CheckKeyCode = 0
	If InStr(Key,"%") Then Key = Replace(Key,"%","x")
	If CheckKey <> "" And Key <> "" Then
		FindStrArr = Split(Convert(CheckKey),",",-1)
		For i = 0 To UBound(FindStrArr)
			FindStr = Trim(FindStrArr(i))
			If InStr(FindStr,"%") Then FindStr = Replace(FindStr,"%","x")
			If InStr(FindStr,"-") Then
				If Left(FindStr,1) <> "[" And Right(FindStr,1) <> "]" Then
					FindStr = "[" & FindStr & "]"
				End If
			End If
			Pos = InStr(FindStr,"[")
			If Pos > 0 Then
				If Left(FindStr,Pos-1) <> "[" And Right(FindStr,Pos+1) <> "]" Then
					FindStr = Replace(FindStr,"[","[[]")
				End If
			End If
			KeyCode = False
			CheckKeyCode = 0
			'PSL.Output Key & " : " &  FindStr  '调试用
			KeyCode = UCase(Key) Like UCase(FindStr)
			If KeyCode = True Then
				CheckKeyCode = 1
				Exit For
			End If
		Next i
	ElseIf CheckKey = "" And Key <> "" Then
		CheckKeyCode = 1
	End If
End Function


'除去字串前后指定的 PreStr 和 AppStr
'fType = 0 去除字串前后的空格和所有指定的 PreStr 和 AppStr，但不去除字串内前后空格
'fType = 1 去除字串前后的空格和所有指定的 PreStr 和 AppStr，并去除字串内前后空格
'fType = 2 去除字串前后的空格和指定的 PreStr 和 AppStr 1 次，但不去除字串内前后空格
'fType > 2 去除字串前后的空格和指定的 PreStr 和 AppStr 1 次，并去除字串内前后空格
Function RemoveBackslash(Path As String,PreStr As String,AppStr As String,fType As Long) As String
	Dim i As Long,Stemp As Boolean
	RemoveBackslash = Path
	If Path = "" Then Exit Function
	RemoveBackslash = Trim(RemoveBackslash)
	Do
		Stemp = False
		If PreStr <> "" And Left(RemoveBackslash,Len(PreStr)) = PreStr Then
			RemoveBackslash = Mid(RemoveBackslash,Len(PreStr)+1)
			Stemp = True
		End If
		If AppStr <> "" And Right(RemoveBackslash,Len(AppStr)) = AppStr Then
			RemoveBackslash = Left(RemoveBackslash,Len(RemoveBackslash)-Len(AppStr))
			Stemp = True
		End If
		If fType = 1 Or fType > 2 Then RemoveBackslash = Trim(RemoveBackslash)
		If Stemp = True Then
			If fType < 2 Then i = 0 Else i = 1
		Else
			i = 1
		End If
	Loop Until i = 1
End Function


' 读取文件
Function ReadFile(FilePath As String,CharSet As String) As String
	Dim objStream As Object,Code As String
	If Dir(FilePath) = "" Then Exit Function
	On Error Resume Next
	Set objStream = CreateObject("Adodb.Stream")
	On Error GoTo ErrorMsg
	Code = CharSet
	If Not objStream Is Nothing Then
		'If Code = "" Then Code = CheckCode(FilePath)
		If Code = "" Then Code = "_autodetect_all"
		If Code = "utf-8EFBB" Then Code = "utf-8"
		If Code <> "ANSI" Then
			With objStream
				.Type = 2
				.Mode = 3
				.Charset = Code
				.Open
				.LoadFromFile FilePath
				ReadFile = .ReadText
				.Close
			End With
		End If
	End If
	If objStream Is Nothing Or Code = "ANSI" Then
		FN = FreeFile
		Open FilePath For Input As #FN
		While Not EOF(FN)
			Line Input #FN,l$
			If ReadFile <> "" Then ReadFile = ReadFile & vbCrLf & l$
			If ReadFile = "" Then ReadFile = l$
		Wend
		Close #FN
	End If
	If CharSet = "" Then CharSet = Code
	Set objStream = Nothing
	Exit Function

	ErrorMsg:
	ReadFile = ""
	Set objStream = Nothing
	Err.Source = "NotReadFile"
	Err.Description = Err.Description & JoinStr & FilePath
	Call sysErrorMassage(Err,1)
End Function


'读取语言对
Function LangCodeList(DataName As String,MinNum As Long,MaxNum As Long) As Variant
	Dim i As Long,LangMaxNum As Long,LangCode As String,LangName As String,Stemp As Boolean
	Dim LangPairs() As String,TempList() As String

	Stemp = False
	If Join(PslLangDataList,"") = "" Then
		PslLangCode = "|af|sq|am|ar|hy|as|az|ba|eu|be|BN|bs|br|bg|ca|zh-CN|zh-TW|co|hr|cs|da|nl|" & _
				"en|et|fo|fa|fil-PH|fi|fr|fy|gl|ka|de|el|kl|gu|ha|he|hi|hu|is|id|iu|ga|xh|zu|it|" & _
				"ja|kn|KS|kk|km|rw|kok|ko|kz|ky|lo|lv|lt|lb|mk|ms|ML|mt|mi|mr|mn|ne|no|nb|" & _
				"nn|or|ps|pl|pt|pa|qu|ro|ru|se|sa|sr|st|tn|SD|si|sk|sl|es|sw|sv|sy|tg|ta|tt|" & _
				"te|th|bo|tr|tk|ug|uk|ur|uz|vi|cy|wo"
		PslLangDataList = Split(PslLangCode,LngJoinStr)
		Stemp = True
	End If

	BingLangCode = "||||ar||||||||||bg||zh-CHS|zh-CHT|||cs|da|nl|en|||||fi|fr||||de|el||||he||" & _
				"hu|||||||it|ja|||||||ko||||lv|lt||||||||||no|no|no|||pl|pt|||ro|ru||||||||sk|" & _
				"sl|es||sv||||||th||tr||||||||"

	GoogleLangCode = "auto|af|sq||ar||||||be||||bg|ca|zh-CN|zh-TW||hr|cs|da|nl|en|et||fa|tl|fi|" & _
				"fr||gl||de|el||||iw|hi|hu|is|id||ga|||it|ja|||||||ko||||lv|lt||mk|ms||mt|||||" & _
				"no|no|no|||pl|pt|||ro|ru|||sr|||||sk|sl|es|sw|sv||||||th||tr|||uk|||vi|cy|"

	YahooLangCode = "||||||||||||||||zh|zt|||||nl|en||||||fr||||de|el|||||||||||||it|ja|||||||" & _
				"ko|||||||||||||||||||||pt||||ru||||||||||es||||||||||||||||||"

	en2zhCheck = "||||||||||||||||zh-CN|zh-TW|||||||||||||||||||||||||||||||ja|||||||ko|||||||" & _
				"||||||||||||||||||||||||||||||||||||||||||||||"

	zh2enCheck = "|af|sq|am|ar|hy|as|az|ba|eu|be|BN|bs|br|bg|ca|||co|hr|cs|da|nl|" & _
				"en|et|fo|fa|fil-PH|fi|fr|fy|gl|ka|de|el|kl|gu|ha|he|hi|hu|is|id|iu|ga|xh|zu|it|" & _
				"|kn|KS|kk|km|rw|kok||kz|ky|lo|lv|lt|lb|mk|ms|ML|mt|mi|mr|mn|ne|no|nb|" & _
				"nn|or|ps|pl|pt|pa|qu|ro|ru|se|sa|sr|st|tn|SD|si|sk|sl|es|sw|sv|sy|tg|ta|tt|" & _
				"te|th|bo|tr|tk|ug|uk|ur|uz|vi|cy|wo"

	LangMaxNum = UBound(PslLangDataList)
	If MaxNum > LangMaxNum Then MaxNum = LangMaxNum
	ReDim TempList(LangMaxNum),LangPairs(MaxNum - MinNum)

	If DataName = DefaultEngineList(0) Then TempList = Split(BingLangCode,LngJoinStr)
	If DataName = DefaultEngineList(1) Then TempList = Split(GoogleLangCode,LngJoinStr)
	If DataName = DefaultEngineList(2) Then TempList = Split(YahooLangCode,LngJoinStr)
	If DataName = DefaultCheckList(0) Then TempList = Split(en2zhCheck,LngJoinStr)
	If DataName = DefaultCheckList(1) Then TempList = Split(zh2enCheck,LngJoinStr)

	For i = 0 To LangMaxNum
		If Stemp = True Then
			LangCode = PslLangDataList(i)
			'If LangCode = "zh-CN" Or LangCode = "zh-TW" Or LangCode = "fil-PH" Then
			'	LangName = PSL.GetLangCode(PSL.GetLangID(LangCode,pslCodeLangRgn),pslCodeText)
			'Else
			'	LangName = PSL.GetLangCode(PSL.GetLangID(LangCode,pslCode639_1),pslCodeText)
			'	If LangName = "3FF3F" Then
			'		LangName = PSL.GetLangCode(PSL.GetLangID(LangCode,pslCodeLangRgn),pslCodeText)
			'	End If
			'End If
			PslLangDataList(i) = LangName & LngJoinStr & LangCode
		End If
		If i >= MinNum And i <= MaxNum Then
			LangPairs(i - MinNum) = PslLangDataList(i) & LngJoinStr & TempList(i)
		End If
	Next i
	LangCodeList = LangPairs
End Function


'从来源列表从获取指定项目的目标列表
Function getMsgList(SourceList() As String,TargetList() As String,Items As String,fType As Long) As Boolean
	Dim i As Long,j As Long,n As Long,ItemMax As Long,ItemList() As String,TempList() As String
	getMsgList = False
	ItemList = Split(Items,"|")
	ItemMax = UBound(ItemList)
	n = 0
	For i = 0 To UBound(SourceList)
		TempArray = Split(SourceList(i),JoinStr)
		Header$ = TempArray(0)
		For j = 0 To ItemMax
			If Header$ = ItemList(j) Then
				ReDim Preserve TempList(n)
				TempList(n) = TempArray(2)
				ItemList(j) = ""
				n = n + 1
				Exit For
			End If
		Next j
		If n = ItemMax + 1 Then
			TargetList = Split(Join(TempList,SubJoinStr),SubJoinStr)
			Items = Join(ItemList,"|")
			getMsgList = True
			Exit For
		End If
	Next i
	If getMsgList = False Then
		If fType = 0 Then
			Err.Raise(1,"NotSection",LangFile & JoinStr & Items)
		ElseIf fType < 3 Then
			On Error GoTo ErrorMassage
			Err.Raise(1,"NotSection",LangFile & JoinStr & Items)
			ErrorMassage:
			Call sysErrorMassage(Err,fType)
		End If
	End If
End Function


'读取 INI 文件
Function getUILangList(UIFile As String,TargetList() As String) As Boolean
	Dim i As Long,m As Long,n As Long,j As Long,Max As Long
	Dim ItemList() As String,ValueList() As String,LineArray() As String
	On Error GoTo ErrorMsg
	i = FileLen(UIFile)
	ReDim readByte(i) As Byte,ItemList(0) As String,ValueList(0) As String
	FN = FreeFile
	Open UIFile For Binary As #FN
	Get #FN,,readByte
	Close #FN
	LineArray = Split(readByte,vbCrLf)
	Erase readByte
	If Join(LineArray,"") = "" Then Exit Function
	Max = UBound(LineArray)
	n = 0
	For i = 0 To Max
		L$ = Trim(LineArray(i))
		If L$ <> "" Then
			If Left(L$,1) = "[" And Right(L$,1) = "]" Then
				Header$ = Trim(Mid(L$,2,Len(L$)-2))
			End If
			If Header$ <> "" And HeaderBak$ = "" Then HeaderBak$ = Header$
			If Header$ <> "" And Header$ = HeaderBak$ Then
				setPreStr$ = ""
				setAppStr$ = ""
				j = InStr(L$,"=")
				If j > 0 Then
					setPreStr$ = Trim(Left(L$,j - 1))
					setAppStr$ = Trim(Mid(L$,j + 1))
				End If
				If setPreStr$ <> "" Then
					ReDim Preserve ItemList(n),ValueList(n)
					ItemList(n) = setPreStr$
					ValueList(n) = Convert(RemoveBackslash(setAppStr$,"""","""",2))
					n = n + 1
				End If
			End If
		End If
		If Header$ <> "" And (i = Max Or Header$ <> HeaderBak$) Then
			If n > 0 Then
				ReDim Preserve TargetList(m)
				Items = Join(ItemList,SubJoinStr)
				Values = Join(ValueList,SubJoinStr)
				TargetList(m) = HeaderBak$ & JoinStr & Items & JoinStr & Values
				m = m + 1
				n = 0
				getUILangList = True
			End If
			HeaderBak$ = Header$
		End If
	Next i
	Exit Function

	ErrorMsg:
	Err.Source = "NotINIFile"
	Err.Description = Err.Description & JoinStr & UIFile
	Call sysErrorMassage(Err,0)
End Function


'获取语言文件列表
Function GetUIList(List() As String,DataList() As String) As Boolean
	Dim i As Long,j As Long,Max As Long,readByte() As Byte,Header As String,setAppStr As String
	GetUIList = False
	File = Dir$(MacroDir & "\Data\" & updateAppName & "_*.lng")
	Do While File <> ""
		sExtName = Mid(File,InStrRev(File,".")+1)
		If LCase(sExtName) = "lng" Then
			i = FileLen(MacroDir & "\Data\" & File)
			ReDim readByte(i) As Byte
			On Error Resume Next
			FN = FreeFile
			Open MacroDir & "\Data\" & File For Binary As #FN
			Get #FN,,readByte
			Close #FN
			On Error GoTo 0
			TempArray = Split(readByte,vbCrLf)
			Max = UBound(TempArray)
			For i = 0 To Max
				L$ = Trim(TempArray(i))
				If L$ <> "" Then
					If Left(L$,1) = "[" And Right(L$,1) = "]" Then
						Header = Trim(Mid(L$,2,Len(L$)-2))
					End If
					If Header = "Option" Then
						setPreStr = ""
						setAppStr = ""
						j = InStr(L$,"=")
						If j > 0 Then
							setPreStr = Trim(Left(L$,j - 1))
							setAppStr = Trim(Mid(L$,j + 1))
						End If
						If setAppStr <> "" Then setAppStr = RemoveBackslash(setAppStr,"""","""",0)
						If setPreStr = "AppName" Then AppName = setAppStr
						If setPreStr = "Version" Then OldVersion = setAppStr
						If setPreStr = "LanguageName" Then LangName = setAppStr
						If setPreStr = "LanguageID" Then LangID = setAppStr
						If setPreStr = "Encoding" Then Encoding = setAppStr
					End If
				End If
				If Header <> "" And (i = Max Or Header <> "Option") Then
					Header = ""
					Exit For
				End If
			Next i
			If LCase(AppName) = LCase(updateAppName) And OldVersion = Version And LangName <> "" Then
				Data = LangName & JoinStr & LangID & JoinStr & File
				CreateArray(LangName,Data,List,DataList)
				OldVersion = ""
				AppName = ""
				LangName = ""
				LangID = ""
				Encoding = ""
				GetUIList = True
			End If
		End If
		File = Dir$()
	Loop
End Function
