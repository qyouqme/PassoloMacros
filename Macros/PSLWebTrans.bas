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


'��������Ĭ������
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


'�ִ�����Ĭ������
Function CheckSettings(DataName As String,OSLang As String) As String
	Dim ExCr As String,LnSp As String,ChkBkt As String,KpPair As String,ChkEnd As String
	Dim NoTrnEnd As String,TrnEnd As String,Short As String,Key As String,KpKey As String
	If DataName = DefaultCheckList(0) Then
		If OSLang = "0404" Then
			ExCr = "&&,& ,&#,&amp;,&nbsp;,&quot;,&lt;,&gt;,&apos;"
			LnSp = "\r,\n,\r\n,|,&"
			ChkBkt = "(),�]�^,[],�e�f,<>,�ա�,�q�r"
			KpPair = "(),�]�^,[],�e�f,{},�a�b,<>,�ա�,�q�r,�m�n,�i�j,�u�v,�y�z,'',�y�z,����,����,�u�v,����,"""""
			AsiaKey = "1"

			ChkEnd = ". .. ... .... ..... ...... �C : �G ; �F ! �I ? �H , �A �B > >> -> ] } + -"
			TrnEnd = ",|�A .|�C ;|�F !|�I ?|�H"
			NoTrnEnd = "*[*],*{*},*<*>,*(?,*[?,*],*{?,*},*<?,*>,*%?!?!,*%?!??!"

			Short = "\t,\b"
			Key = "BackSpace,Tab,Clear,Enter,Shift,Ctrl,Alt,Pause,Caps Lock,Escape,Esc,Space,Break," & _
					"Page Up,Page Down,PgUp,PgDn,End,Home,Left,Up,Right,Down,Left Arrow,Up Arrow," & _
					"Right Arrow,Down Arrow,Select,Print,Ins,Delete,Del,Help,Num Lock,F#,F1#,%?,&?,?"
			KpKey = "�V�W?,�V�U?,�V��?,�V�k?,�W�b?,�U�b?,���b?,�k�b?," & _
					"�V�W��,�V�U��,�V����,�V�k��,�W�b�Y,�U�b�Y,���b�Y,�k�b�Y,��,��,��,��"
			PreStr = "\001|\n001,\u00A9|(C),\u00AE|(R)"
			AppStr = "&lt;|<,&gt;|>,&quot;|"",&apos;|',&#9;|\t,&#10;|\n,&#13;|\r,&amp;|&,&#39;|'"
		Else
			ExCr = "&&,& ,&#,&amp;,&nbsp;,&quot;,&lt;,&gt;,&apos;"
			LnSp = "\r,\n,\r\n,|,&"
			ChkBkt = "(),����,[],�ۣ�,<>,����,����"
			KpPair = "(),����,[],�ۣ�,{},����,<>,����,����,����,����,����,����,'',����,�A�@,�F�F,����,����,"""""
			AsiaKey = "1"

			ChkEnd = ". .. ... .... ..... ...... �� : �� ; �� ! �� ? �� , �� �� > >> -> ] } + -"
			TrnEnd = ",|�� .|�� ;|�� !|�� ?|��"
			NoTrnEnd = "*[*],*{*},*<*>,*(?,*[?,*],*{?,*},*<?,*>,*%?!?!,*%?!??!"

			Short = "\t,\b"
			Key = "BackSpace,Tab,Clear,Enter,Shift,Ctrl,Alt,Pause,Caps Lock,Escape,Esc,Space,Break," & _
					"Page Up,Page Down,PgUp,PgDn,End,Home,Left,Up,Right,Down,Left Arrow,Up Arrow," & _
					"Right Arrow,Down Arrow,Select,Print,Ins,Delete,Del,Help,Num Lock,F#,F1#,%?,&?,?"
			KpKey = "���ϼ�,���¼�,�����,���Ҽ�,�ϼ�ͷ,�¼�ͷ,���ͷ,�Ҽ�ͷ," & _
					"�����I,�����I,�����I,�����I,�ϼ��^,�¼��^,����^,�Ҽ��^,��,��,��,��"
			PreStr = "\001|\n001,\u00A9|(C),\u00AE|(R)"
			AppStr = "&lt;|<,&gt;|>,&quot;|"",&apos;|',&#9;|\t,&#10;|\n,&#13;|\r,&amp;|&,&#39;|'"
		End If
	ElseIf DataName = DefaultCheckList(1) Then
		If OSLang = "0404" Then
			ExCr = "&&,& ,&#,&amp;,&nbsp;,&quot;,&lt;,&gt;,&apos;"
			LnSp = "\r,\n,\r\n,|,&"
			ChkBkt = "(),�]�^,[],�e�f,<>,�ա�,�q�r"
			KpPair = "(),�]�^,[],�e�f,{},�a�b,<>,�ա�,�q�r,�m�n,�i�j,�u�v,�y�z,'',�y�z,����,����,�u�v,����,"""""
			AsiaKey = "0"

			ChkEnd = ". .. ... .... ..... ...... �C : �G ; �F ! �I ? �H , �A �B > >> -> ] } + -"
			TrnEnd = "�A|, �C|. �F|; �I|! �H|?"
			NoTrnEnd = "*[*],*{*},*<*>,*(?,*[?,*],*{?,*},*<?,*>,*%?!?!,*%?!??!"

			Short = "\t,\b"
			Key = "BackSpace,Tab,Clear,Enter,Shift,Ctrl,Alt,Pause,Caps Lock,Escape,Esc,Space,Break," & _
					"Page Up,Page Down,PgUp,PgDn,End,Home,Left,Up,Right,Down,Left Arrow,Up Arrow," & _
					"Right Arrow,Down Arrow,Select,Print,Ins,Delete,Del,Help,Num Lock,F#,F1#,%?,&?,?," & _
					"�V�W?,�V�U?,�V��?,�V�k?,�W�b?,�U�b?,���b?,�k�b?," & _
					"�V�W��,�V�U��,�V����,�V�k��,�W�b�Y,�U�b�Y,���b�Y,�k�b�Y,��,��,��,��"
			KpKey = "Up,Right,Down,Left Arrow,Up Arrow,Right Arrow,Down Arrow"
			PreStr = "\001|\n001,\u00A9|(C),\u00AE|(R)"
			AppStr = "&lt;|<,&gt;|>,&quot;|"",&apos;|',&#9;|\t,&#10;|\n,&#13;|\r,&amp;|&,&#39;|'"
		Else
			ExCr = "&&,& ,&#,&amp;,&nbsp;,&quot;,&lt;,&gt;,&apos;"
			LnSp = "\r,\n,\r\n,|,&"
			ChkBkt = "(),����,[],�ۣ�,<>,����,����"
			KpPair = "(),����,[],�ۣ�,{},����,<>,����,����,����,����,����,����,'',����,�A�@,�F�F,����,����,"""""
			AsiaKey = "0"

			ChkEnd = ". .. ... .... ..... ...... �� : �� ; �� ! �� ? �� , �� �� > >> -> ] } + -"
			TrnEnd = "��|, ��|. ��|; ��|! ��|?"
			NoTrnEnd = "*[*],*{*},*<*>,*(?,*[?,*],*{?,*},*<?,*>,*%?!?!,*%?!??!"

			Short = "\t,\b"
			Key = "BackSpace,Tab,Clear,Enter,Shift,Ctrl,Alt,Pause,Caps Lock,Escape,Esc,Space,Break," & _
					"Page Up,Page Down,PgUp,PgDn,End,Home,Left,Up,Right,Down,Left Arrow,Up Arrow," & _
					"Right Arrow,Down Arrow,Select,Print,Ins,Delete,Del,Help,Num Lock,F#,F1#,%?,&?,?," & _
					"���ϼ�,���¼�,�����,���Ҽ�,�ϼ�ͷ,�¼�ͷ,���ͷ,�Ҽ�ͷ," & _
					"�����I,�����I,�����I,�����I,�ϼ��^,�¼��^,����^,�Ҽ��^,��,��,��,��""
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


'����ز鿴�Ի�������������˽������Ϣ��
Private Function MainDlgFunc%(DlgItem$, Action%, SuppValue&)
	Dim nSelected As String,EngineID As Integer,CheckID As Integer,TranLang As String,xmlHttp As Object

	If OSLanguage = "0404" Then
		Msg01 = "���~"
		Msg02 = "�п���n½Ķ���r��I"
		Msg03 = "�L�k�x�s�I���ˬd�O�_���g�J�U�C��m���v��:" & vbCrLf & vbCrLf
		Msg04 = "�T�{"
		Msg05 = "�z���t�ίʤ� Microsoft.XMLHTTP ����A�L�k�~�����I"
		Msg06 = "½Ķ�������A���T���O�ɡI�ݭn�������ݮɶ��ܡH"
		Msg07 = "���ݮɶ�:"
		Msg08 = "�п�J���ݮɶ�"
		Msg09 = "�L�k�P½Ķ�������A���q�H�I�i��O�L Internet �s���A" & vbCrLf & _
				"�Ϊ�½Ķ�������]�w���~�A�Ϊ�½Ķ�����T��s���C"
		Msg10 = "½Ķ�������}���šA�L�k�~��I"
	Else
		Msg01 = "����"
		Msg02 = "��ѡ��Ҫ������ִ���"
		Msg03 = "�޷����棡�����Ƿ���д������λ�õ�Ȩ��:" & vbCrLf & vbCrLf
		Msg04 = "ȷ��"
		Msg05 = "����ϵͳȱ�� Microsoft.XMLHTTP �����޷��������У�"
		Msg06 = "���������������Ӧ��ʱ����Ҫ�ӳ��ȴ�ʱ����"
		Msg07 = "�ȴ�ʱ��:"
		Msg08 = "������ȴ�ʱ��"
		Msg09 = "�޷��뷭�����������ͨ�ţ��������� Internet ���ӣ�" & vbCrLf & _
				"���߷�����������ô��󣬻��߷��������ֹ���ʡ�"
		Msg10 = "����������ַΪ�գ��޷�������"
	End If

	'��ȡĿ������
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
	Case 2 ' ��ֵ���Ļ��߰����˰�ť
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

		'����ִ����ͺ�����ѡ���Ƿ�ͬʱѡ��ȫ������������
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
			'��� Microsoft.XMLHTTP �Ƿ����
			Set xmlHttp = CreateObject(DefaultObject)
			If xmlHttp Is Nothing Then Set xmlHttp = CreateObject("Msxml2.XMLHTTP")
			If xmlHttp Is Nothing Then
				MsgBox(Msg05,vbOkOnly+vbInformation,Msg01)
				MainDlgFunc% = True ' ��ֹ���°�ť�رնԻ��򴰿�
				Exit Function
			End If
			'��ȡ���Է���
			trnString = getTranslate(EngineID,xmlHttp,"Test","",3)
			Set xmlHttp = Nothing
			'���� Internet ����
			If trnString = "NotConnected" Then
				MsgBox(Msg09,vbOkOnly+vbInformation,Msg01)
				MainDlgFunc% = True ' ��ֹ���°�ť�رնԻ��򴰿�
				Exit Function
			End If
			'����������ַ�Ƿ�Ϊ��
			If trnString = "NullUrl" Then
				MsgBox(Msg10,vbOkOnly+vbInformation,Msg01)
				MainDlgFunc% = True ' ��ֹ���°�ť�رնԻ��򴰿�
				Exit Function
			End If
			'�������������Ƿ�ʱ
			If trnString = "Timeout" Then
				Massage = MsgBox(Msg06,vbYesNoCancel+vbInformation,Msg04)
				If Massage = vbYes Then WaitTimes = InputBox(Msg07,Msg08,WaitTimes)
				If Massage = vbCancel Then
					MainDlgFunc% = True ' ��ֹ���°�ť�رնԻ��򴰿�
					Exit Function
				End If
			End If
			'����ִ����ͺ�����ѡ���Ƿ�Ϊ��
			If mAllType + mMenu + mDialog + mString + mAccTable + mVer + mOther + mSelOnly = 0 Then
				MsgBox(Msg02,vbOkOnly+vbInformation,Msg01)
				MainDlgFunc% = True ' ��ֹ���°�ť�رնԻ��򴰿�
				Exit Function
			End If
			If Join(tSelected,JoinStr) = nSelected Then Exit Function
			tSelected = Split(nSelected,JoinStr)
		 	If EngineWrite(EngineDataList,tWriteLoc,"All") = False Then
				If tWriteLoc = EngineFilePath Then Msg03 = Msg03 & EngineFilePath
				If tWriteLoc = EngineRegKey Then Msg03 = Msg03 & EngineRegKey
				If MsgBox(Msg03,vbYesNo+vbInformation,Msg01) = vbNo Then Exit Function
				MainDlgFunc% = True ' ��ֹ���°�ť�رնԻ��򴰿�
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
			MainDlgFunc% = True ' ��ֹ���°�ť�رնԻ��򴰿�
		End If
	End Select
End Function


' ������
Sub Main
	Dim i As Integer,j As Integer,srcString As String,trnString As String,TranLang As String
	Dim src As PslTransList,TrnList As PslTransList,TransListOpen As Boolean
	Dim SrcLangList() As String,LangPairList() As String,xmlHttp As Object,objStream As Object
	Dim srcLngFind As Integer,trnLngFind As Integer,StringCount As Integer
	Dim CheckID As Integer,EngineID As Integer,srcLng As String,trnLng As String
	Dim TranedCount As Integer,SkipedCount As Integer,NotChangeCount As Integer
	Dim LineNumErrCount As Integer,accKeyNumErrCount As Integer,ErrorCount As Integer

	On Error GoTo SysErrorMsg
	'���ϵͳ����
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
		Msg00 = "���v�Ҧ�(C) 2010 by wanfu  ����: " & Version
		Msg01 = "�s�u½Ķ����"
		Msg02 = "���{���q�L�ҿ諸�s�u½Ķ�����M��L�ﶵ�A�۰�½Ķ�M�椤���r��C" & _
				"�z�i�H�ۭq½Ķ�����Ψ�ѼơC"

		Msg03 = "½Ķ�M��: "
		Msg04 = "½Ķ����"
		Msg05 = "½Ķ���"
		Msg06 = "½Ķ�r��"
		Msg07 = "����(&A)"
		Msg08 = "���(&M)"
		Msg09 = "��ܤ��(&D)"
		Msg10 = "�r���(&S)"
		Msg11 = "�[�t����(&A)"
		Msg12 = "����(&V)"
		Msg13 = "��L(&O)"
		Msg14 = "�ȿ��(&L)"

		Msg15 = "���L�r��"
		Msg16 = "���мf(&K)"
		Msg17 = "�w����(&E)"
		Msg18 = "��½Ķ(&N)"
		Msg19 = "�����Ʀr�M�Ÿ�(&M)"
		Msg20 = "�����j�g�^��(&U)"
		Msg21 = "�����p�g�^��(&L)"

		Msg22 = "�r��B�z"
		Msg23 = "�]�w:"
		Msg24 = "½Ķ�e:"
		Msg25 = "½Ķ��:"
		Msg26 = "�۰ʿ���]�w"
		Msg27 = "�h���K����(&K)"
		Msg28 = "�h���[�t��(&J)"
		Msg29 = "�����S�w�r���æb½Ķ���٭�(&P)"
		Msg30 = "����½Ķ(&F)"
		Msg31 = "�ȥ��K����B�פ�šB�[�t��(&K)"
		Msg32 = "�����S�w�r��(&R)"

		Msg33 = "�~��ɦ۰��x�s�Ҧ����(&V)"
		Msg34 = "��ܿ�X�T��(&O)"
		Msg35 = "�s�W½Ķ����(&M)"

		Msg36 = "����(&A)"
		Msg37 = "����(&H)"
		Msg38 = "�]�w(&S)"
		Msg39 = "�x�s���(&L)"

		Msg42 = "�T�{"
		Msg43 = "�T��"
		Msg44 = "���~"
		Msg45 =	"�z�� Passolo �����ӧC�A�������ȾA�Ω� Passolo 6.0 �ΥH�W�����A�Фɯū�A�ϥΡC"
		Msg46 = "�п���@��½Ķ�M��I"
		Msg47 = "���b�إߩM��s½Ķ�M��..."
		Msg48 = "�L�k�إߩM��s½Ķ�M��A���ˬd�z���M�׳]�w�C"
		Msg49 = "�����۰�½Ķ"
		Msg50 = "�ӲM�楼�Q�}�ҡC�����A�U���i�i��s�u½Ķ�C" & vbCrLf & _
				"�z�ݭn���t�Φ۰ʶ}�Ҹ�½Ķ�M��ܡH"
		Msg51 = "���b�}��½Ķ�M��..."
		Msg52 = "�L�k�}��½Ķ�M��A���ˬd�z���M�׳]�w�C"
		Msg53 = "�ӲM��w�B��}�Ҫ��A�C�����A�U�i��s�u½Ķ�N�ϱz" & vbCrLf & _
				"���x�s��½Ķ�L�k�٭�C���F�w���A�t�αN���x�s�z��" & vbCrLf & _
				"½Ķ�A�M��i��s�u½Ķ�C" & vbCrLf & vbCrLf & _
				"�z�T�w�n���t�Φ۰��x�s�z��½Ķ�ܡH"
		Msg54 = "���b�إߩM��s½Ķ�ӷ��M��..."
		Msg55 = "�L�k�إߩM��s½Ķ�ӷ��M��A���ˬd�z���M�׳]�w�C"
		Msg56 = "��½Ķ�M��ؼлy���ҹ�����½Ķ�����y���N�X���šA�{���N�����C"
		Msg57 = "���b½Ķ�M�B�z�r��A�i��ݭn�X�����A�еy�J..."
		Msg58 = "�w���L"
		Msg59 = "�w½Ķ"
		Msg60 = "���ܧ�A½Ķ���G�M�{��½Ķ�ۦP�C"
		Msg62 = "�r��w��w�C"
		Msg63 = "�r���Ū�C"
		Msg64 = "�r��w½Ķ���мf�C"
		Msg65 = "�r��w½Ķ�����ҡC"
		Msg66 = "�r�ꥼ½Ķ�C"
		Msg67 = "�r�ꬰ�ũΥ����Ů�C"
		Msg68 = "�r������Ʀr�M�Ÿ��C"
		Msg69 = "�r������j�g�^��μƦr�����C"
		Msg70 = "�r������p�g�^��μƦr�����C"
		Msg71 = "�X�p�ή�: "
		Msg72 = "hh �p�� mm �� ss ��"
		Msg73 = "�^��줤��"
		Msg74 = "�����^��"
		Msg75 = "�A"
		Msg76 = "��"
		Msg77 = "�C"
	Else
		Msg00 = "��Ȩ����(C) 2010 by wanfu  �汾: " & Version
		Msg01 = "���߷����"
		Msg02 = "������ͨ����ѡ�����߷������������ѡ��Զ������б��е��ִ���" & _
				"�������Զ��巭�����漰�������"

		Msg03 = "�����б�: "
		Msg04 = "��������"
		Msg05 = "����Դ��"
		Msg06 = "�����ִ�"
		Msg07 = "ȫ��(&A)"
		Msg08 = "�˵�(&M)"
		Msg09 = "�Ի���(&D)"
		Msg10 = "�ִ���(&S)"
		Msg11 = "��������(&A)"
		Msg12 = "�汾(&V)"
		Msg13 = "����(&O)"
		Msg14 = "��ѡ��(&L)"

		Msg15 = "�����ִ�"
		Msg16 = "������(&K)"
		Msg17 = "����֤(&E)"
		Msg18 = "δ����(&N)"
		Msg19 = "ȫΪ���ֺͷ���(&M)"
		Msg20 = "ȫΪ��дӢ��(&U)"
		Msg21 = "ȫΪСдӢ��(&L)"

		Msg22 = "�ִ�����"
		Msg23 = "����:"
		Msg24 = "����ǰ:"
		Msg25 = "�����:"
		Msg26 = "�Զ�ѡ������"
		Msg27 = "ȥ����ݼ�(&K)"
		Msg28 = "ȥ��������(&J)"
		Msg29 = "�滻�ض��ַ����ڷ����ԭ(&P)"
		Msg30 = "���з���(&F)"
		Msg31 = "������ݼ�����ֹ����������(&K)"
		Msg32 = "�滻�ض��ַ�(&R)"

		Msg33 = "����ʱ�Զ���������ѡ��(&V)"
		Msg34 = "��ʾ�����Ϣ(&O)"
		Msg35 = "��ӷ���ע��(&M)"

		Msg36 = "����(&A)"
		Msg37 = "����(&H)"
		Msg38 = "����(&S)"
		Msg39 = "����ѡ��(&L)"

		Msg42 = "ȷ��"
		Msg43 = "��Ϣ"
		Msg44 = "����"
		Msg45 =	"���� Passolo �汾̫�ͣ������������ Passolo 6.0 �����ϰ汾������������ʹ�á�"
		Msg46 = "��ѡ��һ�������б�"
		Msg47 = "���ڴ����͸��·����б�..."
		Msg48 = "�޷������͸��·����б��������ķ������á�"
		Msg49 = "�����Զ�����"
		Msg50 = "���б�δ���򿪡���״̬�²����Խ������߷��롣" & vbCrLf & _
				"����Ҫ��ϵͳ�Զ��򿪸÷����б���"
		Msg51 = "���ڴ򿪷����б�..."
		Msg52 = "�޷��򿪷����б��������ķ������á�"
		Msg53 = "���б��Ѵ��ڴ�״̬����״̬�½������߷��뽫ʹ��" & vbCrLf & _
				"δ����ķ����޷���ԭ��Ϊ�˰�ȫ��ϵͳ���ȱ�������" & vbCrLf & _
				"���룬Ȼ��������߷��롣" & vbCrLf & vbCrLf & _
				"��ȷ��Ҫ��ϵͳ�Զ��������ķ�����"
		Msg54 = "���ڴ����͸��·�����Դ�б�..."
		Msg55 = "�޷������͸��·�����Դ�б��������ķ������á�"
		Msg56 = "�÷����б�Ŀ����������Ӧ�ķ����������Դ���Ϊ�գ������˳���"
		Msg57 = "���ڷ���ʹ����ִ���������Ҫ�����ӣ����Ժ�..."
		Msg58 = "������"
		Msg59 = "�ѷ���"
		Msg60 = "δ���ģ������������з�����ͬ��"
		Msg62 = "�ִ���������"
		Msg63 = "�ִ�ֻ����"
		Msg64 = "�ִ��ѷ��빩����"
		Msg65 = "�ִ��ѷ��벢��֤��"
		Msg66 = "�ִ�δ���롣"
		Msg67 = "�ִ�Ϊ�ջ�ȫΪ�ո�"
		Msg68 = "�ִ�ȫΪ���ֺͷ��š�"
		Msg69 = "�ִ�ȫΪ��дӢ�Ļ����ַ��š�"
		Msg70 = "�ִ�ȫΪСдӢ�Ļ����ַ��š�"
		Msg71 = "�ϼ���ʱ: "
		Msg72 = "hh Сʱ mm �� ss ��"
		Msg73 = "Ӣ�ĵ�����"
		Msg74 = "���ĵ�Ӣ��"
		Msg75 = "��"
		Msg76 = "��"
		Msg77 = "��"
	End If

	If PSL.Version < 600 Then
		MsgBox Msg45,vbOkOnly+vbInformation,Msg43
		Exit Sub
	End If

	'��� Adodb.Stream �Ƿ���ڲ���ȡ�ַ������б�
	Set objStream = CreateObject("Adodb.Stream")
	If objStream Is Nothing Then
		MsgBox(Msg78,vbOkOnly+vbInformation,Msg43)
		Exit Sub
	End If
	Set objStream = Nothing
	CodeList = CodePageList(0,49)

	Set trn = PSL.ActiveTransList
	'��ⷭ���б��Ƿ�ѡ��
	If trn Is Nothing Then
		MsgBox Msg46,vbOkOnly+vbInformation,Msg43
		Exit Sub
	End If

	'��ȡ��Դ�����б�
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

	'��ʼ������
	ReDim AppNames(3),AppPaths(3)
	ReDim DefaultEngineList(2),EngineList(0),EngineDataList(0)
	DefaultEngineList(0) = "Microsoft"
	DefaultEngineList(1) = "Google"
	DefaultEngineList(2) = "Yahoo"
	ReDim DefaultCheckList(1),CheckList(0),CheckDataList(0)
	DefaultCheckList(0) = Msg73
	DefaultCheckList(1) = Msg74

	'��ȡ������������
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

	'��ȡ�ִ���������
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

	'��ȡ�������ݲ�����°汾
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

	'�Ի���
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
		OKButton 420,434,90,21,.OKButton '5 ����
		CancelButton 510,434,90,21,.CancelButton '6 ȡ��
	End Dialog
	Dim dlg As UserDialog
	If Dialog(dlg) = 0 Then GoTo ExitSub
	AllCont = 1
	AccKey = 0
	EndChar = 0
	Acceler = 0
	EngineID = dlg.EngineList
	CheckID = dlg.CheckList

	'��ȡ�ִ��������
	If dlg.Menu = 1 Then StrTypes = "|Menu|"
	If dlg.Dialog = 1 Then StrTypes = StrTypes & "|Dialog|"
	If dlg.Strings = 1 Then StrTypes = StrTypes & "|StringTable|"
	If dlg.AccTable = 1 Then StrTypes = StrTypes & "|AcceleratorTable|"
	If dlg.Versions = 1 Then StrTypes = StrTypes & "|Version|"

	'��ʾ�򿪹رյķ����б��Ա�������߷���
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

	'��ʾ����򿪵ķ����б����⴦������ݲ��ɻָ�
	'If trn.IsOpen = True Then
		'Massage = MsgBox(Msg03 & vbCrLf & vbCrLf & Msg53,vbYesNoCancel,Msg42)
		'If Massage = vbYes Then trn.Save
		'If Massage = vbCancel Then Exit Sub
	'End If

	'��������б�ĸ���ʱ������ԭʼ�б��Զ�����
	If trn.SourceList.LastChange > trn.LastChange Then
		PSL.Output Msg47
		If trn.Update = False Then
			MsgBox Msg48,vbOkOnly+vbInformation,Msg44
			GoTo ExitSub
		End If
	End If

	'ѡ������Դ�б�
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
		'���������Դ�б�ĸ���ʱ������ԭʼ�б��Զ�����
		If src.SourceList.LastChange > src.LastChange Then
			PSL.Output Msg54
			If src.Update = False Then
				MsgBox Msg55,vbOkOnly+vbInformation,Msg44
				GoTo ExitSub
			End If
		End If
	End If

	'���ü���ר�õ��û���������
	If Not trn Is Nothing Then
		If trn.Property(19980) <> "CheckAccessKeys" Then
			trn.Property(19980) = "CheckAccessKeys"
		End If
	End If

	'��ȡPSL����Դ���Դ���
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

	'��ȡPSL��Ŀ�����Դ���
	trnLng = PSL.GetLangCode(trn.Language.LangID,pslCode639_1)
	If trnLng = "zh" Or trnLng = "ja" Or trnLng = "ko" Then TranLang = "Asia"
	If trnLng = "zh" Then
		trnLng = PSL.GetLangCode(trn.Language.LangID,pslCodeLangRgn)
		If trnLng = "zh-CHS" Or trnLng = "zh-SG" Then trnLng = "zh-CN"
		If trnLng = "zh-CHT" Or trnLng = "zh-HK" Or trnLng = "zh-MO" Then trnLng = "zh-TW"
	End If

	'���ҷ��������ж�Ӧ�����Դ���
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

	'�ͷŲ���ʹ�õĶ�̬������ʹ�õ��ڴ�
	Erase TempArray,LangArray,SrcLangList,LangPairList
	Erase CheckListBak,CheckDataListBak,tempCheckList,tempCheckDataList
	Erase EngineListBak,EngineDataListBak
	Erase DelLngNameList,DelSrcLngList,DelTranLngList
	Erase AppNames,AppPaths,FileList

	'�����Ƿ�ѡ�� "��ѡ���ִ�" ������Ҫ������ִ���
	If dlg.Seleted = 0 Then StringCount = trn.StringCount
	If dlg.Seleted = 1 Then StringCount = trn.StringCount(pslSelection)

	'��ʼ����ÿ���ִ�
	PSL.OutputWnd.Clear
	PSL.Output Msg57
	StartTimes = Timer
	Set xmlHttp = CreateObject(DefaultObject)
	If xmlHttp Is Nothing Then Set xmlHttp = CreateObject("Msxml2.XMLHTTP")
	For j = 1 To StringCount
		'�����Ƿ�ѡ�� "��ѡ���ִ�" ������Ҫ������ִ�
		If dlg.Seleted = 0 Then Set TransString = trn.String(j)
		If dlg.Seleted = 1 Then Set TransString = trn.String(j,pslSelection)

		'��Ϣ���ִ���ʼ������ȡ�����б��������Դ�ͷ����ִ�
		SkipMsg = ""
		Massage = ""
		LineMsg = ""
		AcckeyMsg = ""
		ChangeMsg = ""
		orjSrcString = TransString.SourceText
		orjtrnString = TransString.Text

		'�ִ����ʹ���
		If dlg.AllType = 0 And dlg.Seleted = 0 Then
			If InStr("Menu|Dialog|StringTable|AcceleratorTable|Version",TransString.ResType) Then
				If InStr(StrTypes,TransString.ResType) = 0 Then GoTo Skip
			Else
				If dlg.Other = 0 Then GoTo Skip
			End If
		End If

		'�������������ִ�
		If TransString.State(pslStateLocked) = True Then
			SkipMsg = Msg58 & Msg75 & Msg62
			SkipedCount = SkipedCount + 1
			GoTo Skip
		End If
		'����ֻ�����ִ�
		If TransString.State(pslStateReadOnly) = True Then
			SkipMsg = Msg58 & Msg75 & Msg63
			SkipedCount = SkipedCount + 1
			GoTo Skip
		End If
		'�����ѷ��빩������ִ�
		If dlg.ForReview = 1 And TransString.State(pslStateTranslated) = True Then
			If TransString.State(pslStateReview) = True Then
				SkipMsg = Msg58 & Msg75 & Msg64
				SkipedCount = SkipedCount + 1
				GoTo Skip
			End If
		End If
		'�����ѷ��벢��֤���ִ�
		If dlg.Validated = 1 And TransString.State(pslStateTranslated) = True Then
			If TransString.State(pslStateReview) = False Then
				SkipMsg = Msg58 & Msg75 & Msg65
				SkipedCount = SkipedCount + 1
				GoTo Skip
			End If
		End If
		'����δ������ִ�
		If dlg.NotTran = 1 And TransString.State(pslStateTranslated) = False Then
			SkipMsg = Msg58 & Msg75 & Msg66
			SkipedCount = SkipedCount + 1
			GoTo Skip
		End If
		'����Ϊ�ջ�ȫΪ�ո���ִ�
		If Trim(orjSrcString) = "" Then
			SkipMsg = Msg58 & Msg75 & Msg67
			SkipedCount = SkipedCount + 1
			GoTo Skip
		End If
		'����ȫΪ���ֺͷ��ŵ��ִ�
		If dlg.NumAndSymbol = 1 Then
			If CheckStr(orjSrcString,"0-64,91-96,123-191") = True Then
				SkipMsg = Msg58 & Msg75 & Msg68
				SkipedCount = SkipedCount + 1
				GoTo Skip
			End If
		End If
		'����ȫΪ��дӢ�ĵ��ִ�
		If dlg.AllUCase = 1 Then
			If CheckStr(orjSrcString,"65-90") = True Then
				SkipMsg = Msg58 & Msg75 & Msg69
				SkipedCount = SkipedCount + 1
				GoTo Skip
			End If
		End If
		'����ȫΪСдӢ�ĵ��ִ�
		If dlg.AllLCase = 1 Then
			If CheckStr(orjSrcString,"97-122") = True Then
				SkipMsg = Msg58 & Msg75 & Msg70
				SkipedCount = SkipedCount + 1
				GoTo Skip
			End If
		End If

		'��ȡ����Դ���ִ�
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

		'��ʼԤ���������ִ�
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

		'��ʼ�����ִ����滻ԭ�з���
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

		'��֯��Ϣ�����
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

	'�����������Ϣ���
	ErrorCount = LineNumErrCount + accKeyNumErrCount
	PSL.Output TranMassage(TranedCount,SkipedCount,NotChangeCount,ErrorCount)
	If ErrorCount = 0 And TransListOpen = True Then trn.Close
	EndTimes = Timer
	PSL.Output Msg71 & Format(DateAdd("s",EndTimes - StartTimes,0),Msg72)

	'ȡ������ר�õ��û��������Ե�����
	ExitSub:
	If Not trn Is Nothing Then
		If trn.Property(19980) = "CheckAccessKeys" Then
			trn.Property(19980) = ""
		End If
	End If
	On Error GoTo 0
	Exit Sub

	'��ʾ���������Ϣ
	SysErrorMsg:
	Call sysErrorMassage(Err)
	GoTo ExitSub
End Sub


'��Ⲣ�����°汾
Function Download(Method As String,Url As String,Async As String,Mode As String) As Boolean
	Dim i As Integer,n As Integer,m As Integer,k As Integer,updateINI As String
	Dim TempPath As String,File As String,OpenFile As Boolean,Body As Variant
	Dim xmlHttp As Object,UrlList() As String,Stemp As Boolean
	If OSLanguage = "0404" Then
		Msg01 = "���~"
		Msg02 = "��s���ѡI"
		Msg03 = "���ե��ѡI"
		Msg04 = "�t�ΨS���w�� RAR �����Y�M�ε{���I�L�k�����Y�U���ɮסC"
		Msg05 = "�ʤ֥��n���ѼơA���ˬd�]�w�����۰ʧ�s�ѼƳ]�w�C"
		Msg06 = "�}�ҧ�s���}���ѡI���ˬd���}�O�_���T�Υi�s���C"
		Msg08 = "�L�k����T���I���ˬd���}�O�_���T�Υi�s���C"
		Msg09 = "�L�k�����Y�ɮסI���ˬd�����{���θ����ѼƬO�_���T�C"
		Msg10 = "�{���W��: "
		Msg11 = "��R���|: "
		Msg12 = "����Ѽ�: "
		Msg13 = "RAR �����Y�{�������I�i��O�{�����|���~�Τw�Q�����C"
		Msg14 = "�۰ʧ�s����"
		Msg15 = "���զ��\�I��s���}�M�����{���Ѽƥ��T�C"
		Msg16 = "�ثe���W�������� %s�C"
		Msg17 = "�o�{���i�Ϊ��s���� - %s�A�O�_�U����s�H"
		Msg18 = "�T��"
		Msg19 = "�T�{"
		Msg20 = "��s���\�I�{���N�����A������Э��s�Ұʵ{���C"
		Msg21 = "������T�����]�t�����T���I���ˬd���ˬd���}�O�_���T�C"
		Msg22 = "��s���e:"
		Msg23 = "�O�_�ݭn���s�U���ç�s�H"
		Msg24 = "���b���է�s���}�M�����{���ѼơA�еy��..."
		Msg25 = "���b�ˬd�s�����A�еy��..."
		Msg26 = "���b�U���s�����A�еy��..."
		Msg27 = "���b�����Y�A�еy��..."
		Msg28 = "���b�T�{�U�����{�������A�еy��..."
		Msg29 = "���b�w�˷s�����A�еy��..."
		Msg30 = "�z���t�ίʤ� Microsoft.XMLHTTP ����A�L�k��s�I"
	Else
		Msg01 = "����"
		Msg02 = "����ʧ�ܣ�"
		Msg03 = "����ʧ�ܣ�"
		Msg04 = "ϵͳû�а�װ RAR ��ѹ��Ӧ�ó����޷���ѹ�������ļ���"
		Msg05 = "ȱ�ٱ�Ҫ�Ĳ��������������е��Զ����²������á�"
		Msg06 = "�򿪸�����ַʧ�ܣ�������ַ�Ƿ���ȷ��ɷ��ʡ�"
		Msg08 = "�޷���ȡ��Ϣ��������ַ�Ƿ���ȷ��ɷ��ʡ�"
		Msg09 = "�޷���ѹ���ļ��������ѹ������ѹ�����Ƿ���ȷ��"
		Msg10 = "��������: "
		Msg11 = "����·��: "
		Msg12 = "���в���: "
		Msg13 = "RAR ��ѹ������δ�ҵ��������ǳ���·��������ѱ�ж�ء�"
		Msg14 = "�Զ����²���"
		Msg15 = "���Գɹ���������ַ�ͽ�ѹ���������ȷ��"
		Msg16 = "Ŀǰ���ϵİ汾Ϊ %s��"
		Msg17 = "�����п��õ��°汾 - %s���Ƿ����ظ��£�"
		Msg18 = "��Ϣ"
		Msg19 = "ȷ��"
		Msg20 = "���³ɹ��������˳����˳�����������������"
		Msg21 = "��ȡ����Ϣ�������汾��Ϣ������������ַ�Ƿ���ȷ��"
		Msg22 = "��������:"
		Msg23 = "�Ƿ���Ҫ�������ز����£�"
		Msg24 = "���ڲ��Ը�����ַ�ͽ�ѹ������������Ժ�..."
		Msg25 = "���ڼ���°汾�����Ժ�..."
		Msg26 = "���������°汾�����Ժ�..."
		Msg27 = "���ڽ�ѹ�������Ժ�..."
		Msg28 = "����ȷ�����صĳ���汾�����Ժ�..."
		Msg29 = "���ڰ�װ�°汾�����Ժ�..."
		Msg30 = "����ϵͳȱ�� Microsoft.XMLHTTP �����޷����£�"
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


'��ע����л�ȡ RAR ��չ����Ĭ�ϳ���
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


'��ȡ���߷���
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


'Utf-8 ����
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


'ANSI ����
Public Function ANSIEncode(textStr As String) As String
    Dim i As Long,startIndex As Long,endIndex As Long,x() As Byte
    x = StrConv(textStr,vbFromUnicode)
    startIndex = LBound(x)
    endIndex = UBound(x)
    For i = startIndex To endIndex
        ANSIEncode = ANSIEncode & "%" & Hex(x(i))
    Next i
End Function


'ת���ַ��ı����ʽ
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


'ת������������Ϊָ�������ʽ���ַ�
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


'д����������ݵ��ļ�
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


'���� XML ��ʽ������ȡ�����ı�
Function ReadXML(xmlObj As Object,IdNames As String,TagNames As String) As String
	Dim xmlDoc As Object,Node As Object,Item As Object,IdName As String,TagName As String
	Dim x As Integer,y As Integer,i As Integer,max As Integer
	If xmlObj Is Nothing Then Exit Function

	Set xmlDoc = CreateObject("Microsoft.XMLDOM")
	'xmlDoc.Async = False
	'xmlDoc.ValidateOnParse = False
	'xmlDoc.loadXML(xmlObj)	'�����ִ�
	xmlDoc.Load(xmlObj)		'���ض���
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


'��ȡָ��ǰ���ַ�֮���ֵ
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


'����ִ��Ƿ���������ֺͷ���
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


'���з��봦��
Function SplitTran(xmlHttp As Object,srcStr As String,LangPair As String,Arg As String,fType As Integer) As String
	Dim i As Integer,srcStrBak As String,SplitStr As String
	Dim EngineID As Integer,CheckID As Integer,mAccKey As Integer,mAccelerator As Integer

	TempArray = Split(Arg,JoinStr,-1)
	EngineID = CLng(TempArray(0))
	CheckID = CLng(TempArray(1))
	mAccKey = CLng(TempArray(2))
	mAccelerator = CLng(TempArray(3))

	'���滻������ִ�
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

	'��ȡÿ�еķ���
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


'�����ݼ��ַ�
Function AccessKeyHanding(CheckID As Integer,srcStr As String) As String
	Dim i As Integer,j As Integer,n As Integer,posin As Integer
	Dim AccessKey As String,Stemp As Boolean

	srcStrBak = srcStr
	If InStr(srcStr,"&") = 0 Then
		AccessKeyHanding = srcStr
		Exit Function
	End If

	'��ȡѡ�����õĲ���
	TempArray = Split(CheckDataList(CheckID),JoinStr)
	SetsArray = Split(TempArray(1),SubJoinStr)
	ExcludeChar = SetsArray(0)
	CheckBracket = SetsArray(2)

	'�ų��ִ��еķǿ�ݼ�
	If ExcludeChar <> "" Then
		FindStrArr = Split(Convert(ExcludeChar),",",-1)
		For i = LBound(FindStrArr) To UBound(FindStrArr)
			FindStr = LTrim(FindStrArr(i))
			If FindStr <> "" Then
				srcStrBak = Replace(srcStrBak,FindStr,"*a" & i & "!N!" & i & "d*")
			End If
		Next i
	End If

	'��ȡ��ݼ���ȥ��
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

	'��ԭ�ִ��б��ų��ķǿ�ݼ�
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


'����������ַ�
Function AcceleratorHanding(CheckID As Integer,srcStr As String) As String
	Dim i As Integer,j As Integer,x As Integer,y As Integer
	Dim Shortcut As String,ShortcutKey As String,FindStr As String

	'��ȡѡ�����õĲ���
	TempArray = Split(CheckDataList(CheckID),JoinStr)
	SetsArray = Split(TempArray(1),SubJoinStr)
	CheckShortChar = SetsArray(8)
	CheckShortKey = SetsArray(9)
	KeepShortKey = SetsArray(10)

	'��ȡ������
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

	'ȥ��������
	If Shortcut <> "" Then
		x = InStrRev(srcStr,Shortcut)
		If x <> 0 Then AcceleratorHanding = Left(srcStr,x-1)
	Else
		AcceleratorHanding = srcStr
	End If
End Function


'�滻�ض��ַ�
Function ReplaceStr(CheckID As Integer,trnStr As String,fType As Integer) As String
	Dim i As Integer,BaktrnStr As String
	'��ȡѡ�����õĲ���
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


'���������ݼ�����ֹ���ͼ�����
Function CheckHanding(CheckID As Integer,srcStr As String,trnStr As String,TranLang As String) As String
	Dim i As Integer,BaksrcStr As String,BaktrnStr As String,srcStrBak As String,trnStrBak As String
	Dim srcNum As Integer,trnNum As Integer,srcSplitNum As Integer,trnSplitNum As Integer
	Dim FindStrArr() As String,srcStrArr() As String,trnStrArr() As String,LineSplitArr() As String
	Dim posinSrc As Integer,posinTrn As Integer

	'��ȡѡ�����õĲ���
	TempArray = Split(CheckDataList(CheckID),JoinStr)
	SetsArray = Split(TempArray(1),SubJoinStr)
	ExcludeChar = SetsArray(0)
	LineSplitChar = SetsArray(1)
	KeepCharPair = SetsArray(3)

	'���ַ���������
	If LineSplitChar <> "" Then
		FindStrArr = Split(LineSplitChar,",",-1)
		FindStrArr = SortArray(FindStrArr,0,"Lenght",">")
		LineSplitChar = Join(FindStrArr,",")
	End If

	'������ʼ��
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

	'�ų��ִ��еķǿ�ݼ�
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

	'���˲��ǿ�ݼ��Ŀ�ݼ�
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

	'���滻������ִ�
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

	'�ִ�����
	srcSplitNum = UBound(srcStrArr)
	trnSplitNum = UBound(trnStrArr)
	If srcSplitNum = 0 And trnSplitNum = 0 Then
		BaktrnStr = StringReplace(CheckID,BaksrcStr,BaktrnStr,TranLang)
	ElseIf srcSplitNum <> 0 Or trnSplitNum <> 0 Then
		LineSplitArr = MergeArray(srcStrArr,trnStrArr)
		BaktrnStr = ReplaceStrSplit(CheckID,BaktrnStr,LineSplitArr,TranLang)
	End If

	'��������
	LineSplitChars = "\r\n,\r,\n"
	FindStrArr = Split(Convert(LineSplitChars),",",-1)
	For i = LBound(FindStrArr) To UBound(FindStrArr)
		FindStr = Trim(FindStrArr(i))
		If InStr(BaksrcStr,FindStr) Then srcLineNum = UBound(Split(BaksrcStr,FindStr,-1))
		If InStr(BaktrnStr,FindStr) Then trnLineNum = UBound(Split(BaktrnStr,FindStr,-1))
	Next i

	'�����ݼ���
	If InStr(BaksrcStr,"&") Then srcAccKeyNum = UBound(Split(BaksrcStr,"&",-1))
	If InStr(BaktrnStr,"&") Then trnAccKeyNum = UBound(Split(BaktrnStr,"&",-1))

	'��ԭ���ǿ�ݼ��Ŀ�ݼ�
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

	'��ԭ�ִ��б��ų��ķǿ�ݼ�
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


' ���л�ȡ�ִ��ĸ����ֶβ��滻�����ַ���
Function StringReplace(CheckID As Integer,srcStr As String,trnStr As String,TranLang As String) As String
	Dim posinSrc As Integer,posinTrn As Integer,StringSrc As String,StringTrn As String
	Dim accesskeySrc As String,accesskeyTrn As String,Temp As String
	Dim ShortcutPosSrc As Integer,ShortcutPosTrn As Integer,PreTrn As String
	Dim EndStringPosSrc As Integer,EndStringPosTrn As Integer,AppTrn As String
	Dim preKeyTrn As String,appKeyTrn As String,Stemp As Boolean,FindStrArr() As String
	Dim i As Integer,j As Integer,x As Integer,y As Integer,m As Integer,n As Integer

	'��ȡѡ�����õĲ���
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

	'���ַ���������
	If CheckEndChar <> "" Then
		FindStrArr = Split(CheckEndChar,-1)
		FindStrArr = SortArray(FindStrArr,0,"Lenght",">")
		CheckEndChar = Join(FindStrArr)
	End If

	'������ʼ��
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

	'��ȡ�ִ�ĩβ�ո�
	EndSpaceSrc = Space(Len(srcStr) - Len(RTrim(srcStr)))
	EndSpaceTrn = Space(Len(trnStr) - Len(RTrim(trnStr)))

	'��ȡ������
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

	'��ȡ��ֹ���������ַ����ᱻ���
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

	'��ȡԭ�ĺͷ���Ŀ�ݼ�λ��
	posinSrc = InStrRev(srcStr,"&")
	posinTrn = InStrRev(trnStr,"&")

	'��ȡԭ�ĺͷ���Ŀ�ݼ� (������ݼ���ǰ����ַ�)
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

	'��ȡ��ݼ�����ķ���ֹ���ͷǼ��������ַ�����Щ�ַ������ƶ�����ݼ�ǰ
	If posinTrn <> 0 Then
		x = Len(EndStringTrn & ShortcutTrn & EndSpaceTrn)
		If InStr(ShortcutTrn,"&") Then x = Len(EndSpaceTrn)
		If InStr(EndStringTrn,"&") Then x = Len(ShortcutTrn & EndSpaceTrn)
		If Len(trnStr) > x Then
			Temp = Left(trnStr,Len(trnStr) - x)
			ExpStringTrn = Mid(Temp,posinTrn + Len(acckeyTrn))
		End If
	End If

	'ȥ����ݼ�����ֹ���������ǰ��Ŀո�
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

	'��ȡ�����п�ݼ�ǰ����ֹ��
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

	'�Զ����������������ֹ��
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

	'Ҫ����������ֹ�����
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

	'�������������ļ���������
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

	'�ִ�����ѡ����
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

	'���ݼ���
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

	'PSL.Output "------------------------------ "      '������
	'PSL.Output "srcStr = " & srcStr                   '������
	'PSL.Output "trnStr = " & trnStr                   '������
	'PSL.Output "SpaceTrn = " & SpaceTrn               '������
	'PSL.Output "acckeySrc = " & acckeySrc             '������
	'PSL.Output "acckeyTrn = " & acckeyTrn             '������
	'PSL.Output "EndStringSrc = " & EndStringSrc       '������
	'PSL.Output "EndStringTrn = " & EndStringTrn       '������
	'PSL.Output "ShortcutSrc = " & ShortcutSrc         '������
	'PSL.Output "ShortcutTrn = " & ShortcutTrn         '������
	'PSL.Output "ExpStringTrn = " & ExpStringTrn       '������
	'PSL.Output "StringSrc = " & StringSrc             '������
	'PSL.Output "StringTrn = " & StringTrn             '������
	'PSL.Output "NewStringTrn = " & NewStringTrn       '������
	'PSL.Output "PreStringTrn = " & PreStringTrn       '������

	'�ִ��滻
	Temp = trnStr
	If StringSrc <> StringTrn Then
		If StringTrn <> "" And StringTrn <> NewStringTrn Then
			x = InStrRev(Temp,StringTrn)
			If x <> 0 Then
				PreTrn = Left(Temp,x - 1)
				AppTrn = Mid(Temp,x)
				'PSL.Output "PreTrn = " & PreTrn       '������
				'PSL.Output "AppTrn = " & AppTrn       '������
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


' �޸���Ϣ���
Function ReplaceMassage(OldtrnString As String,NewtrnString As String) As String
	Dim AcckeyMsg As String,EndStringMsg As String,ShortcutMsg As String,Tmsg1 As String
	Dim Tmsg2 As String,Fmsg As String,Smsg As String,Massage1 As String,Massage2 As String
	Dim Massage3 As String,Massage4 As String,n As Integer

	If OSLanguage = "0404" Then
		Msg00 = "�w�ק�F�K����"
		Msg01 = "�w�ק�F�פ��"
		Msg02 = "�w�ק�F�[�t��"
		Msg03 = "�w�ק�F�K����M�פ��"
		Msg04 = "�w�ק�F�פ�ũM�[�t��"
		Msg05 = "�w�ק�F�K����M�[�t��"
		Msg06 = "�w�ק�F�K����B�פ�ũM�[�t��"
		Msg07 = "�w�ק�F�K���䪺�j�p�g"
		Msg08 = "�w�ק�F�K���䪺�j�p�g�M�פ��"
		Msg09 = "�w�ק�F�K���䪺�j�p�g�M�[�t��"
		Msg10 = "�w�ק�F�K���䪺�j�p�g�B�פ�ũM�[�t��"

		Msg11 = "�w�s�W�F�K����"
		Msg12 = "�w�s�W�F�פ��"
		Msg13 = "�w�s�W�F�[�t��"
		Msg14 = "�w�s�W�F�K����M�פ��"
		Msg15 = "�w�s�W�F�פ�ũM�[�t��"
		Msg16 = "�w�s�W�F�K����M�[�t��"
		Msg17 = "�w�s�W�F�K����B�פ�ũM�[�t��"

		Msg21 = "��Ķ�夤���K����"
		Msg22 = "��Ķ�夤���פ��"
		Msg23 = "��Ķ�夤���[�t��"
		Msg24 = "��Ķ�夤���K����M�פ��"
		Msg25 = "��Ķ�夤���פ�ũM�[�t��"
		Msg26 = "��Ķ�夤���K����M�[�t��"
		Msg27 = "��Ķ�夤���K����B�פ�ũM�[�t��"

		Msg31 = "�w���ʫK�����̫�"
		Msg32 = "�w���ʫK�����פ�ūe"
		Msg33 = "�w���ʫK�����[�t���e"

		Msg41 = "�w�h���F�K����e���Ů�"
		Msg42 = "�w�h���F�פ�ūe���Ů�"
		Msg43 = "�w�h���F�[�t���e���Ů�"
		Msg44 = "�w�h���F�K����e���פ��"
		Msg45 = "�w�h���F�K����e���Ů�M�פ��"

		Msg46 = "�w�s�W�F�פ�ūe�ʤ֪��Ů�"
		Msg47 = "�w�s�W�F�פ�ū�ʤ֪��Ů�"
		Msg48 = "�w�s�W�F�פ�ūe��ʤ֪��Ů�"
		Msg49 = "�w�s�W�F�r���ʤ֪��Ů�"
		Msg50 = "�w�s�W�F�פ�ūe�M�r���ʤ֪��Ů�"
		Msg51 = "�w�s�W�F�פ�ū�M�r���ʤ֪��Ů�"
		Msg52 = "�w�s�W�F�פ�ūe��M�r���ʤ֪��Ů�"
		Msg53 = "�w�h���F�פ�ūe�h�l���Ů�"
		Msg54 = "�w�h���F�פ�ū�h�l���Ů�"
		Msg55 = "�w�h���F�פ�ūe��h�l���Ů�"
		Msg56 = "�w�h���F�r���h�l���Ů�"
		Msg57 = "�w�h���F�פ�ūe�M�r���h�l���Ů�"
		Msg58 = "�w�h���F�פ�ū�M�r���h�l���Ů�"
		Msg59 = "�w�h���F�פ�ūe��M�r���h�l���Ů�"

		Msg61 = "�A"
		Msg62 = "��"
		Msg63 = "�C"
		Msg64 = "�B"
		Msg65 = "�w�N "
		Msg66 = " ������ "
		Msg67 = "�w�R�� "
	Else
		Msg00 = "���޸��˿�ݼ�"
		Msg01 = "���޸�����ֹ��"
		Msg02 = "���޸��˼�����"
		Msg03 = "���޸��˿�ݼ�����ֹ��"
		Msg04 = "���޸�����ֹ���ͼ�����"
		Msg05 = "���޸��˿�ݼ��ͼ�����"
		Msg06 = "���޸��˿�ݼ�����ֹ���ͼ�����"
		Msg07 = "���޸��˿�ݼ��Ĵ�Сд"
		Msg08 = "���޸��˿�ݼ��Ĵ�Сд����ֹ��"
		Msg09 = "���޸��˿�ݼ��Ĵ�Сд�ͼ�����"
		Msg10 = "���޸��˿�ݼ��Ĵ�Сд����ֹ���ͼ�����"

		Msg11 = "������˿�ݼ�"
		Msg12 = "���������ֹ��"
		Msg13 = "������˼�����"
		Msg14 = "������˿�ݼ�����ֹ��"
		Msg15 = "���������ֹ���ͼ�����"
		Msg16 = "������˿�ݼ��ͼ�����"
		Msg17 = "������˿�ݼ�����ֹ���ͼ�����"

		Msg21 = "���������п�ݼ�"
		Msg22 = "������������ֹ��"
		Msg23 = "���������м�����"
		Msg24 = "���������п�ݼ�����ֹ��"
		Msg25 = "������������ֹ���ͼ�����"
		Msg26 = "���������п�ݼ��ͼ�����"
		Msg27 = "���������п�ݼ�����ֹ���ͼ�����"

		Msg31 = "���ƶ���ݼ������"
		Msg32 = "���ƶ���ݼ�����ֹ��ǰ"
		Msg33 = "���ƶ���ݼ���������ǰ"

		Msg41 = "��ȥ���˿�ݼ�ǰ�Ŀո�"
		Msg42 = "��ȥ������ֹ��ǰ�Ŀո�"
		Msg43 = "��ȥ���˼�����ǰ�Ŀո�"
		Msg44 = "��ȥ���˿�ݼ�ǰ����ֹ��"
		Msg45 = "��ȥ���˿�ݼ�ǰ�Ŀո����ֹ��"

		Msg46 = "���������ֹ��ǰȱ�ٵĿո�"
		Msg47 = "���������ֹ����ȱ�ٵĿո�"
		Msg48 = "���������ֹ��ǰ��ȱ�ٵĿո�"
		Msg49 = "��������ִ���ȱ�ٵĿո�"
		Msg50 = "���������ֹ��ǰ���ִ���ȱ�ٵĿո�"
		Msg51 = "���������ֹ������ִ���ȱ�ٵĿո�"
		Msg52 = "���������ֹ��ǰ����ִ���ȱ�ٵĿո�"
		Msg53 = "��ȥ������ֹ��ǰ����Ŀո�"
		Msg54 = "��ȥ������ֹ�������Ŀո�"
		Msg55 = "��ȥ������ֹ��ǰ�����Ŀո�"
		Msg56 = "��ȥ�����ִ������Ŀո�"
		Msg57 = "��ȥ������ֹ��ǰ���ִ������Ŀո�"
		Msg58 = "��ȥ������ֹ������ִ������Ŀո�"
		Msg59 = "��ȥ������ֹ��ǰ����ִ������Ŀո�"

		Msg61 = "��"
		Msg62 = "��"
		Msg63 = "��"
		Msg64 = "��"
		Msg65 = "�ѽ� "
		Msg66 = " �滻Ϊ "
		Msg67 = "��ɾ�� "
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


'�������������Ϣ
Function LineErrMassage(srcLineNum As Integer,trnLineNum As Integer,LineNumErrCount As Integer) As String
	If OSLanguage = "0404" Then
		Msg01 = "Ķ�媺��Ƥ���� %s ��C"
		Msg02 = "Ķ�媺��Ƥ���h %s ��C"
	Else
		Msg01 = "���ĵ�������ԭ���� %s �С�"
		Msg02 = "���ĵ�������ԭ�Ķ� %s �С�"
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


'�����ݼ���������Ϣ
Function AccKeyErrMassage(srcAccKeyNum As Integer,trnAccKeyNum As Integer,accKeyNumErrCount As Integer) As String
	If OSLanguage = "0404" Then
		Msg01 = "Ķ�媺�K����Ƥ���� %s �ӡC"
		Msg02 = "Ķ�媺�K����Ƥ���h %s �ӡC"
	Else
		Msg01 = "���ĵĿ�ݼ�����ԭ���� %s ����"
		Msg02 = "���ĵĿ�ݼ�����ԭ�Ķ� %s ����"
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


'������Ϣ���
Function TranMassage(tCount As Integer,sCount As Integer,nCount As Integer,eCount As Integer) As String
	If OSLanguage = "0404" Then
		Msg01 = "½Ķ�����A�S��½Ķ����r��C"
		Msg02 = "½Ķ�����A�䤤�G"
		Msg03 = "�w½Ķ " & tCount & " �ӡA�w���L " & sCount & " �ӡA" & _
				"���ܧ� " & nCount & " �ӡA�����~ " & eCount & " ��"
	Else
		Msg01 = "������ɣ�û�з����κ��ִ���"
		Msg02 = "������ɣ����У�"
		Msg03 = "�ѷ��� " & tCount & " ���������� " & sCount & " ����" & _
				"δ���� " & nCount & " �����д��� " & eCount & " ��"
	End If
	TranCount = tCount + sCount + nCount
	If TranCount = 0 Then TranMassage = Msg01
	If TranCount <> 0 Then TranMassage = Msg02 & Msg03
End Function


'������������Ϣ
Sub sysErrorMassage(sysError As ErrObject)
	If OSLanguage = "0404" Then
		Msg01 = "���~"
		Msg02 = "�o�͵{���]�p�W�����~�C���~�N�X "
	Else
		Msg01 = "����"
		Msg02 = "������������ϵĴ��󡣴������ "
	End If
	MsgBox(Msg02 & CStr(sysError.Number) & ", " & sysError.Description,vbInformation,Msg01)
End Sub


'�ڿ�ݼ�������ض��ַ����Դ˲���ִ�
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
	'PSL.Output "bakString = " & bakString       '������
End Function


'��������ϲ�
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


'�ִ���������ת��
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


'ת���˽��ƻ�ʮ������ת���
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


'ת���ַ�Ϊ������ֵ
Function StrToInteger(mStr As String) As Integer
	If mStr = "" Then mStr = "0"
	StrToInteger = CInt(mStr)
End Function


'��ȡ�����е�ÿ���ִ����滻����
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

		'������ǰ�����а������ظ��ַ�
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

		'��ÿ�е����ݽ������ӣ�������Ϣ���
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

	'Ϊ������Ϣ�������ԭ�б����滻���Ӻ������
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


'�Զ������
Function Settings(EngineID As Integer,CheckID As Integer) As Integer
	Dim LngNameList() As String,SrcLngList() As String,TranLngList() As String
	Dim AppLngList() As String,UseLngList() As String

	If OSLanguage = "0404" Then
		Msg01 = "�]�w"
		Msg02 = "�п���]�w�æb�U�C��r�������J�ȡA���յL�~��A�M�Ω��ھާ@�C"
		Msg05 = "Ū��(&R)"
		Msg06 = "�s�W(&A)"
		Msg07 = "�ܧ�(&M)"
		Msg08 = "�R��(&D)"
		Msg09 = "�M��(&C)"
		Msg10 = "����(&T)"

		Msg11 = "�]�w�M��"
		Msg12 = "�x�s����"
		Msg13 = "�]�w���e"
		Msg14 = "�ɮ�"
		Msg15 = "���U��"
		Msg16 = "�פJ�]�w"
		Msg17 = "�ץX�]�w"
		Msg18 = "�K����"
		Msg19 = "�פ��"
		Msg20 = "�[�t��"
		Msg21 = "�r������"

		Msg22 = "�n�ư����t && �Ÿ����D�K����r���զX (�b�γr�����j):"
		Msg23 = "�r����ΥκX�в� (�Ω�h�ӫK����r�ꪺ�B�z) (�b�γr�����j):"
		Msg24 = "�n�ˬd���K����e��A���A�Ҧp [&&F] (�b�γr�����j):"
		Msg25 = "�n�O�d���D�K����ūe�ᦨ��r���A�Ҧp (&&) (�b�γr�����j):"
		Msg26 = "�b��r�᭱��ܱa�A�����K���� (�q�`�Ω�Ȭw�y��)"

		Msg27 = "�n�ˬd���פ�� (�� - ��ܽd��A�䴩�U�Φr��) (�Ů���j):"
		Msg28 = "�n�O�d���פ�ŲզX (�� - ��ܽd��A�䴩�U�Φr��) (�b�γr�����j):"
		Msg29 = "�n�۰ʴ������פ�Ź� (�� | ���j�����e�᪺�r��) (�Ů���j):"

		Msg30 = "�n�ˬd���[�t���X�вšA�Ҧp \t (�b�γr�����j):"
		Msg31 = "�n�ˬd���[�t���r�� (�� - ��ܽd��A�䴩�U�Φr��) (�b�γr�����j):"
		Msg32 = "�n�O�d���[�t���r�� (�� - ��ܽd��A�䴩�U�Φr��) (�b�γr�����j):"

		Msg33 = "½Ķ�e�n�Q�������r�� (�� | ���j�����e�᪺�r��) (�b�γr�����j):"
		Msg34 = "½Ķ��n�Q�������r�� (�� | ���j�����e�᪺�r��) (�b�γr�����j):"

		Msg35 = "����(&H)"
		Msg36 = "½Ķ����"
		Msg37 = "�r��B�z"
		Msg38 = "�����Ѽ�"
		Msg39 = "�y���t��"
		Msg40 = "�ϥΪ���:"
		Msg41 = "�������U ID:"
		Msg42 = "�������}:"
		Msg43 = "�ǰe���e�d��:"
		Msg44 = "��ƶǰe�覡:"
		Msg45 = "�P�B�覡:"
		Msg46 = "�ϥΪ̦W(�i��):"
		Msg47 = "�K�X(�i��):"
		Msg48 = "���O��(�i��):"
		Msg49 = "HTTP �Y�M��:"
		Msg50 = "(�����J)"
		Msg51 = "��^���G�榡:"
		Msg52 = "½Ķ�}�l��:"
		Msg53 = "½Ķ������:"
		Msg54 = ">"
		Msg55 = "..."

		Msg56 = "�y���t��"
		Msg57 = "�y���W��"
		Msg58 = "Passolo �N�X"
		Msg59 = "½Ķ�����N�X"
		Msg60 = "�s�W(&A)"
		Msg61 = "�R��(&D)"
		Msg62 = "�����R��"
		Msg63 = "�s��(&E)"
		Msg64 = "�~���s��"
		Msg65 = "�m��(&N)"
		Msg66 = "���](&R)"
		Msg67 = "��ܫD�Ŷ�"
		Msg68 = "��ܪŶ�"
		Msg69 = "�������"

		Msg71 = "�A�λy��"
		Msg74 = "�s�W  >"
		Msg75 = "�����s�W >>"
		Msg76 = "<  �R��"
		Msg77 = "<< �����R��"
		Msg78 = "�s�W�i�λy��"
		Msg79 = "�s��i�λy��"
		Msg80 = "�R���i�λy��"
		Msg81 = "�s�W�A�λy��"
		Msg82 = "�s��A�λy��"
		Msg83 = "�R���A�λy��"

		Msg84 = "�۰ʧ�s"
		Msg85 = "��s�覡"
		Msg86 = "�۰ʤU����s�æw��(&A)"
		Msg87 = "����s�ɳq���ڡA�ѧڨM�w�U���æw��(&M)"
		Msg88 = "�����۰ʧ�s(&O)"
		Msg89 = "�ˬd�W�v"
		Msg90 = "�ˬd���j: "
		Msg91 = "��"
		Msg92 = "�̫��ˬd���:"
		Msg93 = "��s���}�M�� (�����J�A�e���u��)"
		Msg94 = "RAR �����{��"
		Msg95 = "�{�����| (�䴩�����ܼ�):"
		Msg96 = "�����Ѽ� (%1 �����Y�ɮסA%2 ���n�^�����ɮסA%3 ���������|):"
		Msg97 = "�ˬd"

		Tools00 = "���m�{��(&E)"
		Tools01 = "�O�ƥ�(&N)"
		Tools02 = "Excel(&E)"
		Tools03 = "�ۭq�{��(&C)"
	Else
		Msg01 = "����"
		Msg02 = "��ѡ�����ò��������ı���������ֵ�������������Ӧ����ʵ�ʲ�����"
		Msg05 = "��ȡ(&R)"
		Msg06 = "���(&A)"
		Msg07 = "����(&M)"
		Msg08 = "ɾ��(&D)"
		Msg09 = "���(&C)"
		Msg10 = "����(&T)"

		Msg11 = "�����б�"
		Msg12 = "��������"
		Msg13 = "��������"
		Msg14 = "�ļ�"
		Msg15 = "ע���"
		Msg16 = "��������"
		Msg17 = "��������"
		Msg18 = "��ݼ�"
		Msg19 = "��ֹ��"
		Msg20 = "������"
		Msg21 = "�ַ��滻"

		Msg22 = "Ҫ�ų��ĺ� && ���ŵķǿ�ݼ��ַ���� (��Ƕ��ŷָ�):"
		Msg23 = "�ִ�����ñ�־�� (���ڶ����ݼ��ִ��Ĵ���) (��Ƕ��ŷָ�):"
		Msg24 = "Ҫ���Ŀ�ݼ�ǰ�����ţ����� [&&F] (��Ƕ��ŷָ�):"
		Msg25 = "Ҫ�����ķǿ�ݼ���ǰ��ɶ��ַ������� (&&) (��Ƕ��ŷָ�):"
		Msg26 = "���ı�������ʾ�����ŵĿ�ݼ� (ͨ��������������)"

		Msg27 = "Ҫ������ֹ�� (�� - ��ʾ��Χ��֧��ͨ���) (�ո�ָ�):"
		Msg28 = "Ҫ��������ֹ����� (�� - ��ʾ��Χ��֧��ͨ���) (��Ƕ��ŷָ�):"
		Msg29 = "Ҫ�Զ��滻����ֹ���� (�� | �ָ��滻ǰ����ַ�) (�ո�ָ�):"

		Msg30 = "Ҫ���ļ�������־�������� \t (��Ƕ��ŷָ�):"
		Msg31 = "Ҫ���ļ������ַ� (�� - ��ʾ��Χ��֧��ͨ���) (��Ƕ��ŷָ�):"
		Msg32 = "Ҫ�����ļ������ַ� (�� - ��ʾ��Χ��֧��ͨ���) (��Ƕ��ŷָ�):"

		Msg33 = "����ǰҪ���滻���ַ� (�� | �ָ��滻ǰ����ַ�) (��Ƕ��ŷָ�):"
		Msg34 = "�����Ҫ���滻���ַ� (�� | �ָ��滻ǰ����ַ�) (��Ƕ��ŷָ�):"

		Msg35 = "����(&H)"
		Msg36 = "��������"
		Msg37 = "�ִ�����"
		Msg38 = "�������"
		Msg39 = "�������"
		Msg40 = "ʹ�ö���:"
		Msg41 = "����ע�� ID:"
		Msg42 = "������ַ:"
		Msg43 = "��������ģ��:"
		Msg44 = "���ݴ��ͷ�ʽ:"
		Msg45 = "ͬ����ʽ:"
		Msg46 = "�û���(�ɿ�):"
		Msg47 = "����(�ɿ�):"
		Msg48 = "ָ�(�ɿ�):"
		Msg49 = "HTTP ͷ��ֵ:"
		Msg50 = "(��������)"
		Msg51 = "���ؽ����ʽ:"
		Msg52 = "���뿪ʼ��:"
		Msg53 = "���������:"
		Msg54 = ">"
		Msg55 = "..."

		Msg56 = "�������"
		Msg57 = "��������"
		Msg58 = "Passolo ����"
		Msg59 = "�����������"
		Msg60 = "���(&A)"
		Msg61 = "ɾ��(&D)"
		Msg62 = "ȫ��ɾ��"
		Msg63 = "�༭(&E)"
		Msg64 = "�ⲿ�༭"
		Msg65 = "�ÿ�(&N)"
		Msg66 = "����(&R)"
		Msg67 = "��ʾ�ǿ���"
		Msg68 = "��ʾ����"
		Msg69 = "ȫ����ʾ"

		Msg71 = "��������"
		Msg74 = "���  >"
		Msg75 = "ȫ����� >>"
		Msg76 = "<  ɾ��"
		Msg77 = "<< ȫ��ɾ��"
		Msg78 = "��ӿ�������"
		Msg79 = "�༭��������"
		Msg80 = "ɾ����������"
		Msg81 = "�����������"
		Msg82 = "�༭��������"
		Msg83 = "ɾ����������"

		Msg84 = "�Զ�����"
		Msg85 = "���·�ʽ"
		Msg86 = "�Զ����ظ��²���װ(&A)"
		Msg87 = "�и���ʱ֪ͨ�ң����Ҿ������ز���װ(&M)"
		Msg88 = "�ر��Զ�����(&O)"
		Msg89 = "���Ƶ��"
		Msg90 = "�����: "
		Msg91 = "��"
		Msg92 = "���������:"
		Msg93 = "������ַ�б� (�������룬ǰ������)"
		Msg94 = "RAR ��ѹ����"
		Msg95 = "����·�� (֧�ֻ�������):"
		Msg96 = "��ѹ���� (%1 Ϊѹ���ļ���%2 ΪҪ��ȡ���ļ���%3 Ϊ��ѹ·��):"
		Msg97 = "���"

		Tools00 = "���ó���(&E)"
		Tools01 = "���±�(&N)"
		Tools02 = "Excel(&E)"
		Tools03 = "�Զ������(&C)"
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


'����ز鿴�Ի�������������˽������Ϣ��
Private Function SetFunc%(DlgItem$, Action%, SuppValue&)
	Dim Header As String,HeaderID As Integer,NewData As String,Path As String,cStemp As Boolean
	Dim i As Integer,n As Integer,TempArray() As String,Temp As String,tStemp As Boolean
	Dim LngName As String,LngID As Integer,SrcLngCode As String,TranLngCode As String
	Dim LngNameList() As String,SrcLngList() As String,TranLngList() As String
	Dim AppLngList() As String,UseLngList() As String,LangArray() As String

	If OSLanguage = "0404" Then
		Msg01 = "���~"
		Msg02 = "�w�]��"
		Msg03 = "���"
		Msg04 = "�ѷӭ�"
		Msg08 = "����"
		Msg11 = "ĵ�i"
		Msg12 = "�p�G�Y�ǰѼƬ��šA�N�ϵ{�����浲�G�����T�C" & vbCrLf & _
				"�z�T��Q�n�o�˰��ܡH"
		Msg13 = "�]�w���e�w�g�ܧ���O�S���x�s�I" & vbCrLf & "�O�_�ݭn�x�s�H"
		Msg14 = "�x�s�����w�g�ܧ���O�S���x�s�I" & vbCrLf & "�O�_�ݭn�x�s�H"
		Msg18 = "�ثe�]�w���A�ܤ֦��@���ѼƬ��šI" & vbCrLf
		Msg19 = "�Ҧ��]�w���A�ܤ֦��@���ѼƬ��šI" & vbCrLf
		Msg21 = "�T�{"
		Msg22 = "�T��n�R���]�w�u%s�v�ܡH"
		Msg24 = "�T��n�R���y���u%s�v�ܡH"
		Msg25 = "½Ķ������"
		Msg26 = "�r��B�z��"
		Msg27 = "½Ķ�����M�r��B�z��"
		Msg30 = "�T��"
		Msg32 = "�ץX�]�w���\�I"
		Msg33 = "�פJ�]�w���\�I"
		Msg36 = "�L�k�x�s�I���ˬd�O�_���g�J�U�C��m���v��:" & vbCrLf & vbCrLf
		Msg39 = "�פJ���ѡI���ˬd�O�_���g�J�U�C��m���v��" & vbCrLf & _
				"�ζפJ�ɮת��榡�O�_���T:" & vbCrLf & vbCrLf
		Msg40 = "�ץX���ѡI���ˬd�O�_���g�J�U�C��m���v��" & vbCrLf & _
				"�ζץX�ɮת��榡�O�_���T:" & vbCrLf & vbCrLf
		Msg41 = "�x�s���ѡI���ˬd�O�_���g�J�U�C��m���v��:" & vbCrLf & vbCrLf
		Msg42 = "����n�פJ���ɮ�"
		Msg43 = "����n�ץX���ɮ�"
		Msg44 = "�]�w�ɮ� (*.dat)|*.dat|�Ҧ��ɮ� (*.*)|*.*||"
		Msg45 = "�T��n�R�������y���ܡH"
		Msg50 = "½Ķ�}�l��:"
		Msg51 = "½Ķ������:"
		Msg52 = "�� ID �j��:"
		Msg53 = "�����ҦW�j��:"
		Msg54 = "�i�λy��:"
		Msg55 = "�A�λy��:"
		Msg60 = "��������{��"
		Msg61 = "�i�����ɮ� (*.exe)|*.exe|�Ҧ��ɮ� (*.*)|*.*||"
		Msg62 = "�S�����w�����{���I�Э��s��J�ο���C"
		Msg63 = "�ɮװѷӰѼ�(%1)"
		Msg64 = "�n�^�����ɮװѼ�(%2)"
		Msg65 = "�������|�Ѽ�(%3)"
		Item0 = "½Ķ�������} {Url}"
		Item1 = "½Ķ�������U�� {AppId}"
		Item2 = "�n½Ķ����r {Text}"
		Item3 = "�ӷ��y�� {From}"
		Item4 = "�ؼлy�� {To}"
		Item5 = "�P�B����(�w�]) {True}"
		Item6 = "���B���� {False}"
		Item7 = "�L�Ÿ���ư}�C {responseBody}"
		Item8 = "ADO Stream ���� {responseStream}"
		Item9 = "�r�� {responseText}"
		Item10 = "XML �榡��� {responseXML}"
	Else
		Msg01 = "����"
		Msg02 = "Ĭ��ֵ"
		Msg03 = "ԭֵ"
		Msg04 = "����ֵ"
		Msg08 = "δ֪"
		Msg11 = "����"
		Msg12 = "���ĳЩ����Ϊ�գ���ʹ�������н������ȷ��" & vbCrLf & _
				"��ȷʵ��Ҫ��������"
		Msg13 = "���������Ѿ����ĵ���û�б��棡" & vbCrLf & "�Ƿ���Ҫ���棿"
		Msg14 = "���������Ѿ����ĵ���û�б��棡" & vbCrLf & "�Ƿ���Ҫ���棿"
		Msg18 = "��ǰ�����У�������һ�����Ϊ�գ�" & vbCrLf
		Msg19 = "���������У�������һ�����Ϊ�գ�" & vbCrLf
		Msg21 = "ȷ��"
		Msg22 = "ȷʵҪɾ�����á�%s����"
		Msg24 = "ȷʵҪɾ�����ԡ�%s����"
		Msg25 = "���������"
		Msg26 = "�ִ������"
		Msg27 = "����������ִ������"
		Msg30 = "��Ϣ"
		Msg32 = "�������óɹ���"
		Msg33 = "�������óɹ���"
		Msg36 = "�޷����棡�����Ƿ���д������λ�õ�Ȩ��:" & vbCrLf & vbCrLf
		Msg39 = "����ʧ�ܣ������Ƿ���д������λ�õ�Ȩ��" & vbCrLf & _
				"�����ļ��ĸ�ʽ�Ƿ���ȷ:" & vbCrLf & vbCrLf
		Msg40 = "����ʧ�ܣ������Ƿ���д������λ�õ�Ȩ��" & vbCrLf & _
				"�򵼳��ļ��ĸ�ʽ�Ƿ���ȷ:" & vbCrLf & vbCrLf
		Msg41 = "����ʧ�ܣ������Ƿ���д������λ�õ�Ȩ��:" & vbCrLf & vbCrLf
		Msg42 = "ѡ��Ҫ������ļ�"
		Msg43 = "ѡ��Ҫ�������ļ�"
		Msg44 = "�����ļ� (*.dat)|*.dat|�����ļ� (*.*)|*.*||"
		Msg45 = "ȷʵҪɾ��ȫ��������"
		Msg50 = "���뿪ʼ��:"
		Msg51 = "���������:"
		Msg52 = "�� ID ����:"
		Msg53 = "����ǩ������:"
		Msg54 = "��������:"
		Msg55 = "��������:"
		Msg60 = "ѡ���ѹ����"
		Msg61 = "��ִ���ļ� (*.exe)|*.exe|�����ļ� (*.*)|*.*||"
		Msg62 = "û��ָ����ѹ���������������ѡ��"
		Msg63 = "�ļ����ò���(%1)"
		Msg64 = "Ҫ��ȡ���ļ�����(%2)"
		Msg65 = "��ѹ·������(%3)"
		Item0 = "����������ַ {Url}"
		Item1 = "��������ע��� {AppId}"
		Item2 = "Ҫ������ı� {Text}"
		Item3 = "��Դ���� {From}"
		Item4 = "Ŀ������ {To}"
		Item5 = "ͬ��ִ��(Ĭ��) {True}"
		Item6 = "�첽ִ�� {False}"
		Item7 = "�޷����������� {responseBody}"
		Item8 = "ADO Stream ���� {responseStream}"
		Item9 = "�ַ��� {responseText}"
		Item10 = "XML ��ʽ���� {responseXML}"
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
	Case 1 ' �Ի��򴰿ڳ�ʼ��
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
	Case 2 ' ��ֵ���Ļ��߰����˰�ť
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
						SetFunc% = True '��ֹ���°�ť�رնԻ��򴰿�
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
						SetFunc% = True '��ֹ���°�ť�رնԻ��򴰿�
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
						SetFunc% = True '��ֹ���°�ť�رնԻ��򴰿�
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
						SetFunc% = True '��ֹ���°�ť�رնԻ��򴰿�
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
						SetFunc% = True '��ֹ���°�ť�رնԻ��򴰿�
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
					SetFunc% = True '��ֹ���°�ť�رնԻ��򴰿�
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
				SetFunc% = True '��ֹ���°�ť�رնԻ��򴰿�
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
			SetFunc% = True '��ֹ���°�ť�رնԻ��򴰿�
		End If
	Case 3 ' �ı��������Ͽ��ı�������
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


'�����������
Function AddSet(DataArr() As String) As String
	Dim NewHeader As String
	If OSLanguage = "0404" Then
		Msg01 = "�s�W"
		Msg04 = "�п�J�s�]�w���W��:"
		Msg06 = "���~"
		Msg07 = "�z�S����J���󤺮e�I�Э��s��J�C"
		Msg08 = "�ӦW�٤w�g�s�b�I�п�J�@�Ӥ��P���W�١C"
	Else
		Msg01 = "�½�"
		Msg04 = "�����������õ�����:"
		Msg06 = "����"
		Msg07 = "��û�������κ����ݣ����������롣"
		Msg08 = "�������Ѿ����ڣ�������һ����ͬ�����ơ�"
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


'�༭��������
Function EditSet(DataArr() As String,Header As String) As String
	Dim tempHeader As String,NewHeader As String
	If OSLanguage = "0404" Then
		Msg01 = "�ܧ�"
		Msg04 = "�¦W��:"
		Msg05 = "�s�W��:"
		Msg06 = "���~"
		Msg07 = "�z�S����J���󤺮e�I�Э��s��J�C"
		Msg08 = "�ӦW�٤w�g�s�b�I�п�J�@�Ӥ��P���W�١C"
	Else
		Msg01 = "����"
		Msg04 = "������:"
		Msg05 = "������:"
		Msg06 = "����"
		Msg07 = "��û�������κ����ݣ����������롣"
		Msg08 = "�������Ѿ����ڣ�������һ����ͬ�����ơ�"
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


'��ӻ�༭���Զ�
Function EditLang(DataArr() As String,LangName As String,SrcCode As String,TarnCode As String) As String
	Dim tempHeader As String,NewLangName As String,NewSrcCode As String,NewTarnCode As String
	If OSLanguage = "0404" Then
		Msg01 = "�s�W"
		Msg02 = "�s��"
		Msg04 = "�y���W��:"
		Msg05 = "Passolo �y���N�X:"
		Msg06 = "½Ķ�����y���N�X:"
		Msg10 = "���~"
		Msg11 = "�z�S����J���󤺮e�I�Э��s��J�C"
		Msg12 = "�y���W�٩M Passolo �y���N�X���ܤ֦��@�Ӷ��ج��šI���ˬd�ÿ�J�C"
		Msg13 = "�ӻy���W�٤w�g�s�b�I�Э��s��J�C"
		Msg14 = "�� Passolo �y���N�X�w�g�s�b�I�Э��s��J�C"
		Msg15 = "½Ķ�����y���N�X���šI�O�_�n���s��J�H"
		Msg16 = "�p�G�T��ݭn�ŭȡA�t�αN�۰ʳ]�w���uNull�v�ȡC"
	Else
		Msg01 = "���"
		Msg02 = "�༭"
		Msg04 = "��������:"
		Msg05 = "Passolo ���Դ���:"
		Msg06 = "�����������Դ���:"
		Msg10 = "����"
		Msg11 = "��û�������κ����ݣ����������롣"
		Msg12 = "�������ƺ� Passolo ���Դ�����������һ����ĿΪ�գ����鲢���롣"
		Msg13 = "�����������Ѿ����ڣ����������롣"
		Msg14 = "�� Passolo ���Դ����Ѿ����ڣ����������롣"
		Msg15 = "�����������Դ���Ϊ�գ��Ƿ�Ҫ�������룿"
		Msg16 = "���ȷʵ��Ҫ��ֵ��ϵͳ���Զ���Ϊ��Null��ֵ��"
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


'��ӻ�༭���Զ�
Function SetLang(DataArr() As String,LangName As String,LangCode As String) As String
	Dim tempHeader As String,NewLangName As String,NewLangCode As String
	If OSLanguage = "0404" Then
		Msg01 = "�s�W"
		Msg02 = "�s��"
		Msg04 = "�y���W��:"
		Msg05 = "Passolo �y���N�X:"
		Msg10 = "���~"
		Msg11 = "�z�S����J���󤺮e�I�Э��s��J�C"
		Msg12 = "�y���W�٩M Passolo �y���N�X���ܤ֦��@�Ӷ��ج��šI���ˬd�ÿ�J�C"
		Msg13 = "�ӻy���W�٤w�g�s�b�I�Э��s��J�C"
		Msg14 = "�� Passolo �y���N�X�w�g�s�b�I�Э��s��J�C"
	Else
		Msg01 = "���"
		Msg02 = "�༭"
		Msg04 = "��������:"
		Msg05 = "Passolo ���Դ���:"
		Msg10 = "����"
		Msg11 = "��û�������κ����ݣ����������롣"
		Msg12 = "�������ƺ� Passolo ���Դ�����������һ����ĿΪ�գ����鲢���롣"
		Msg13 = "�����������Ѿ����ڣ����������롣"
		Msg14 = "�� Passolo ���Դ����Ѿ����ڣ����������롣"
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


'�����������ȼ�
Function SetLevel(HeaderList() As String,DataList() As String) As Boolean
	SetLevel = False
	If OSLanguage = "0404" Then
		Msg01 = "�]�w�u����"
		Msg02 = "�]�w�u���ťΩ���]�w���A�λy�����۰ʿ���]�w�\��C"
		Msg03 = "����:" & vbCrLf & _
				"- ���h�ӳ]�w�]�t�F�ۦP���A�λy���ɡA�ݭn�]�w���u���šC" & vbCrLf & _
				"- �b�ۦP�A�λy�����]�w���A�e�����]�w�Q�u������ϥΡC"
		Msg04 = "�W��(&U)"
		Msg05 = "�U��(&D)"
		Msg06 = "���](&R)"
	Else
		Msg01 = "�������ȼ�"
		Msg02 = "�������ȼ����ڻ������õ��������Ե��Զ�ѡ�����ù��ܡ�"
		Msg03 = "��ʾ:" & vbCrLf & _
				"- �ж�����ð�������ͬ����������ʱ����Ҫ���������ȼ���" & vbCrLf & _
				"- ����ͬ�������Ե������У�ǰ������ñ�����ѡ��ʹ�á�"
		Msg04 = "����(&U)"
		Msg05 = "����(&D)"
		Msg06 = "����(&R)"
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


'�����������ȼ��Ի�����
Private Function SetLevelFunc%(DlgItem$, Action%, SuppValue&)
	Dim i As Integer,ID As Integer,Temp As String
	Select Case Action%
	Case 1 ' �Ի��򴰿ڳ�ʼ��
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
	Case 2 ' ��ֵ���Ļ��߰����˰�ť
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
			SetLevelFunc = True '��ֹ���°�ť�رնԻ��򴰿�
		End If
	End Select
End Function


'��ȡ����
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
			'��ȡ Option ���ֵ
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
			'��ȡ Update ���ֵ
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
			'��ȡ Option �����ȫ�����ֵ
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
				'���¾ɰ��Ĭ������ֵ
				If InStr(Join(DefaultEngineList,JoinStr),Header) Then
					If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 Then
						Data = EngineDataUpdate(Header,Data)
					End If
				End If
				'�������ݵ�������
				CreateArray(Header,Data,HeaderList,DataList)
				EngineGet = True
			End If
			'���ݳ�ʼ��
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
	'������º͵��������ݵ��ļ�
	If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 And Path = EngineFilePath Then
		If Dir(EngineFilePath) <> "" Then EngineWrite(DataList,EngineFilePath,"All")
	End If
	If tWriteLoc = "" Then tWriteLoc = EngineFilePath
	Exit Function

	GetFromRegistry:
	'��ȡ Option ���ֵ
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
		'��ȡ Update ���ֵ
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

	'��ȡ Option ������ֵ
	HeaderIDs = GetSetting("WebTranslate","Option","Headers","")
	If HeaderIDs <> "" Then
		HeaderIDArr = Split(HeaderIDs,";",-1)
		For i = LBound(HeaderIDArr) To UBound(HeaderIDArr)
			HeaderID = HeaderIDArr(i)
			If HeaderID <> "" Then
				'ת��ɰ��ÿ�����ֵ
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
						'���¾ɰ��Ĭ������ֵ
						If InStr(Join(DefaultEngineList,JoinStr),Header) Then
							If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 Then
								Data = EngineDataUpdate(Header,Data)
							End If
						End If
						'�������ݵ�������
						CreateArray(Header,Data,HeaderList,DataList)
						EngineGet = True
					End If
					'ɾ���ɰ�����ֵ
					On Error Resume Next
					If Header = HeaderID Then DeleteSetting("WebTranslate",Header)
					On Error GoTo 0
				End If
			End If
		Next i
	End If
	'������º�����ݵ�ע���
	If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 Then
		If HeaderIDs <> "" Then EngineWrite(DataList,EngineRegKey,"Sets")
	End If
	If tWriteLoc = "" Then tWriteLoc = EngineRegKey
End Function


'д�뷭����������
Function EngineWrite(DataList() As String,Path As String,WriteType As String) As Boolean
	Dim i As Integer,HeaderID As String,HeaderIDs As String,TempPath As String
	Dim SetsArray() As String,HeaderIDArr() As String
	EngineWrite = False
	KeepSet = tSelected(23)

	'д���ļ�
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

	'д��ע���
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
			'ɾ��ԭ������
			HeaderIDs = GetSetting("WebTranslate","Option","Headers")
			If HeaderIDs <> "" Then
				HeaderIDArr = Split(HeaderIDs,";",-1)
				On Error Resume Next
				For i = LBound(HeaderIDArr) To UBound(HeaderIDArr)
					DeleteSetting("WebTranslate",HeaderIDArr(i))
				Next i
				On Error GoTo 0
			End If
			'д����������
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
	'ɾ�����б��������
	ElseIf Path = "" Then
		'ɾ���ļ�������
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
		'ɾ��ע���������
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
		'����д��λ������Ϊ��
		EngineWrite = True
		tWriteLoc = ""
	End If
	ExitFunction:
End Function


'��ȡ�ִ��������
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
			'��ȡ Option ���ֵ
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
			'��ȡ Update ���ֵ
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
			'��ȡ Option �����ȫ�����ֵ
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
				'���¾ɰ��Ĭ������ֵ
				If InStr(Join(DefaultCheckList,JoinStr),Header) Then
					If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 Then
						Data = CheckDataUpdate(Header,Data)
					End If
				End If
				'�������ݵ�������
				CreateArray(Header,Data,HeaderList,DataList)
				CheckGet = True
			End If
			'���ݳ�ʼ��
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
	'������º͵��������ݵ��ļ�
	If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 And Path = CheckFilePath Then
		If Dir(CheckFilePath) <> "" Then CheckWrite(DataList,CheckFilePath,"All")
	End If
	If cWriteLoc = "" Then cWriteLoc = CheckFilePath
	Exit Function

	GetFromRegistry:
	'��ȡ Option ���ֵ
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
		'��ȡ Update ���ֵ
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
	'��ȡ Option ������ֵ
	HeaderIDs = GetSetting("AccessKey","Option","Headers","")
	If HeaderIDs <> "" Then
		HeaderIDArr = Split(HeaderIDs,";",-1)
		For i = LBound(HeaderIDArr) To UBound(HeaderIDArr)
			HeaderID = HeaderIDArr(i)
			If HeaderID <> "" Then
				'ת��ɰ��ÿ�����ֵ
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
						'���¾ɰ��Ĭ������ֵ
						If InStr(Join(DefaultCheckList,JoinStr),Header) Then
							If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 Then
								Data = CheckDataUpdate(Header,Data)
							End If
						End If
						'�������ݵ�������
						CreateArray(Header,Data,HeaderList,DataList)
						CheckGet = True
					End If
					'ɾ���ɰ�����ֵ
					On Error Resume Next
					If Header = HeaderID Then DeleteSetting("AccessKey",Header)
					On Error GoTo 0
				End If
			End If
		Next i
	End If
	'������º�����ݵ�ע���
	If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 Then
		If HeaderIDs <> "" Then CheckWrite(DataList,CheckRegKey,"Sets")
	End If
	If cWriteLoc = "" Then cWriteLoc = CheckRegKey
End Function


'д���ִ��������
Function CheckWrite(DataList() As String,Path As String,WriteType As String) As Boolean
	Dim i As Integer,Header As String,HeaderIDs As String,TempPath As String
	Dim SetsArray() As String,HeaderIDArr() As String
	CheckWrite = False
	KeepSet = cSelected(UBound(cSelected))

	'д���ļ�
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

	'д��ע���
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
			'ɾ��ԭ������
			HeaderIDs = GetSetting("AccessKey","Option","Headers")
			If HeaderIDs <> "" Then
				HeaderIDArr = Split(HeaderIDs,";",-1)
				On Error Resume Next
				For i = LBound(HeaderIDArr) To UBound(HeaderIDArr)
					DeleteSetting("AccessKey",HeaderIDArr(i))
				Next i
				On Error GoTo 0
			End If
			'д����������
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
	'ɾ�����б��������
	ElseIf Path = "" Then
		'ɾ���ļ�������
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
		'ɾ��ע���������
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
		'����д��λ������Ϊ��
		CheckWrite = True
		cWriteLoc = ""
	End If
	ExitFunction:
End Function


'�滻���������ļ��е���������Ϊ��ǰϵͳ����������
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


'�滻�����������ļ��е��ַ�Ϊ��ǰϵͳ�������ַ�
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


'��������ɰ汾����ֵ
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


'���¼��ɰ汾����ֵ
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


'���ӻ����������Ŀ
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


'ɾ��������Ŀ
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


'�������
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


'����������������������
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


'�ϲ���׼���Զ��������б�
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


'����ȥ�������б��п��������Զ�
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


'����������Ŀ
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


'����ָ��ֵ�Ƿ���������
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


'����������Ƿ��п�ֵ
'ftype = 0     ������������Ƿ�ȫΪ��ֵ
'ftype = 1     ������������Ƿ��п�ֵ
'Header = ""   �����������
'Header <> ""  ���ָ��������
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


'��������
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


'�����������ظ�������
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


'ͨ�������ָ��ֵ
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
			'PSL.Output Key & " : " &  FindStr  '������
			KeyCode = UCase(Key) Like UCase(FindStr)
			If KeyCode = True Then CheckKeyCode = 1
			If KeyCode = True Then Exit For
		Next i
	ElseIf CheckKey = "" And Key <> "" Then
		CheckKeyCode = 1
	End If
End Function


'�������߷������
Sub TranTest(EngineID As Integer,HeaderList() As String,fType As Integer)
	Dim i As Integer,n As Integer,TrnList As PslTransList,TrnListDec As String
	Dim LngNameList() As String,TrnListArray() As String

	If OSLanguage = "0404" Then
		Msg01 = "½Ķ��������"
		Msg02 = "½Ķ�M��M�ӷ��y���|�ھڱM�ײ{�����y���۰ʽT�w�C�n�W�[���ؽзs�W�������y���C"
		Msg03 = "½Ķ����:"
		Msg04 = "�ӷ��y��:"
		Msg05 = "�ؼлy��:"
		Msg06 = "½Ķ�M��:"
		Msg07 = "Ū�J���:"
		Msg08 = "½Ķ���e(�۰ʱq��ܲM�椤Ū�J�Τ�ʿ�J):"
		Msg09 = "½Ķ���G(�����ի��s��b����X���G):"
		Msg10 = "�ӷ��r��"
		Msg11 = "½Ķ�r��"
		Msg12 = "½Ķ��r"
		Msg13 = "������r"
		Msg14 = "����(&H)"
		Msg15 = "½Ķ(&T)"
		Msg16 = "�T���Y(&D)"
		Msg17 = "�M��(&C)"
	Else
		Msg01 = "�����������"
		Msg02 = "�����б����Դ���Ի���ݷ������е������Զ�ȷ����Ҫ������Ŀ�������Ӧ�����ԡ�"
		Msg03 = "��������:"
		Msg04 = "��Դ����:"
		Msg05 = "Ŀ������:"
		Msg06 = "�����б�:"
		Msg07 = "��������:"
		Msg08 = "��������(�Զ���ѡ���б��ж�����ֶ�����):"
		Msg09 = "������(�����԰�ť���ڴ�������):"
		Msg10 = "��Դ�ִ�"
		Msg11 = "�����ִ�"
		Msg12 = "�����ı�"
		Msg13 = "ȫ���ı�"
		Msg14 = "����(&H)"
		Msg15 = "����(&T)"
		Msg16 = "��Ӧͷ(&D)"
		Msg17 = "���(&C)"
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


'���ԶԻ�����
Private Function TranTestFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim ListDec As String,TempDec As String,LineNum As Integer,EngineID As Integer
	Dim inText As String,outText As String,i As Integer,n As Integer,m As Integer
	Dim srcLngFind As Integer,trnLngFind As Integer,LngNameList() As String
	Dim TrnList As PslTransList,xmlHttp As Object

	If OSLanguage = "0404" Then
		Msg01 = "�M��(&C)"
		Msg02 = "Ū�J(&R)"
		Msg05 = "���b½Ķ�A�i��ݭn�X�����A�еy��..."
		Msg06 = "���ժ��y���N�X�אּ: "
		Msg07 = "½Ķ���ѡI���ˬd Internet �s���B½Ķ�������]�w�άO�_�i�s����A�աC"
		Msg08 = "��^���r�ꬰ�šC�i��O½Ķ�������䴩��ܪ��ؼлy���C"
		Msg09 = "========================================="
	Else
		Msg01 = "���(&C)"
		Msg02 = "����(&R)"
		Msg05 = "���ڷ��룬������Ҫ�����ӣ����Ժ�..."
		Msg06 = "���Ե����Դ����Ϊ: "
		Msg07 = "����ʧ�ܣ����� Internet ���ӡ�������������û��Ƿ�ɷ��ʺ����ԡ�"
		Msg08 = "���ص��ִ�Ϊ�ա������Ƿ������治֧��ѡ����Ŀ�����ԡ�"
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
	Case 2 ' ��ֵ���Ļ��߰����˰�ť
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

			'�ִ�Ĭ��Ԥ����
			inText = AccessKeyHanding(0,inText)
			inText = AcceleratorHanding(0,inText)

			'��ȡ�����б����Դ��Ŀ������
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

			'��ȡ�����������Դ��Ŀ������
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

			'��ʼ���벢���
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
			TranTestFunc = True '��ֹ���°�ť�رնԻ��򴰿�
		End If
		If DlgText("InTextBox") <> "" Then DlgText "ClearButton",Msg01
    	If DlgText("InTextBox") = "" Then DlgText "ClearButton",Msg02
    	If DlgText("InTextBox") <> "" Then DlgEnable "TestButton",True
		If DlgText("InTextBox") = "" Then DlgEnable "TestButton",False
    	If DlgText("InTextBox") <> "" Then DlgEnable "HeaderButton",True
		If DlgText("InTextBox") = "" Then DlgEnable "HeaderButton",False
	Case 3 ' �ı��������Ͽ��ı�������
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


'���Լ�����
Sub CheckTest(CheckID As Integer,HeaderList() As String)
	Dim TrnList As PslTransList,i As Integer,TrnListDec As String,TrnListArray() As String
	If OSLanguage = "0404" Then
		Msg01 = "�Ѽƴ���"
		Msg02 = "�ھڤU�C����զX�j�M½Ķ���~�õ��X�ץ��C�� [����] ���s��X���G�C"
		Msg03 = "�]�w�W��:"
		Msg05 = "½Ķ�M��:"
		Msg06 = "Ū�J���:"
		Msg07 = "�]�t���e:"
		Msg08 = "�䴩�U�Φr���M�b�Τ������j���h��"
		Msg09 = "�r�ꤺ�e:"
		Msg10 = "����(&F)"
		Msg11 = "�K����(&K)"
		Msg12 = "�פ��(&E)"
		Msg13 = "�[�t��(&P)"
		Msg15 = "����(&T)"
		Msg16 = "�M��(&C)"
		Msg18 = "����(&H)"
		Msg19 = "�۰ʴ����r��(&R)"
	Else
		Msg01 = "��������"
		Msg02 = "��������������ϲ��ҷ�����󲢸����������� [����] ��ť��������"
		Msg03 = "��������:"
		Msg05 = "�����б�:"
		Msg06 = "��������:"
		Msg07 = "��������:"
		Msg08 = "֧��ͨ����Ͱ�Ƿֺŷָ��Ķ���"
		Msg09 = "�ִ�����:"
		Msg10 = "ȫ��(&F)"
		Msg11 = "��ݼ�(&K)"
		Msg12 = "��ֹ��(&E)"
		Msg13 = "������(&P)"
		Msg15 = "����(&T)"
		Msg16 = "���(&C)"
		Msg18 = "����(&H)"
		Msg19 = "�Զ��滻�ַ�(&R)"
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


'���ԶԻ�����
Private Function CheckTestFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim ListDec As String,TempDec As String,LineNum As Integer,inText As String,repStr As Integer
	Dim cAllCont As Integer,cAccKey As Integer,cEndChar As Integer,cAcceler As Integer
	Dim SpecifyText As String,CheckID As Integer
	If OSLanguage = "0404" Then
		Msg01 = "���b�j�M���~�õ��X�ץ��A�i��ݭn�X�����A�еy��..."
	Else
		Msg01 = "���ڲ��Ҵ��󲢸���������������Ҫ�����ӣ����Ժ�..."
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
	Case 2 ' ��ֵ���Ļ��߰����˰�ť
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
			CheckTestFunc = True '��ֹ���°�ť�رնԻ��򴰿�
		End If
	End Select
End Function


'��������������ִ��б��е��ִ�
Function CheckStrings(ID As Integer,ListDec As String,LineNum As Integer,rStr As Integer,sText As String) As String
	Dim i As Integer,j As Integer,k As Integer,srcString As String,trnString As String,tText As String
	Dim TrnList As PslTransList,TrnListDec As String,TranLang As String
	Dim CheckVer As String,CheckSet As String,CheckState As String,CheckDate As Date,TranDate As Date
	Dim LineNumErrCount As Integer,accKeyNumErrCount As Integer,ErrorCount As Integer,Massage As String
	Dim Find As Boolean,srcFindNum As Integer,trnFindNum As Integer

	If OSLanguage = "0404" Then
		Msg01 = "���: "
		Msg02 = "Ķ��: "
		Msg03 = "�ץ�: "
		Msg04 = "�T��: "
		Msg05 = "---------------------------------------"
		Msg06 = "���U�C���~:"
		Msg07 = "�S�������~�C"
		Msg08 = "�]�t���w���e���r�ꤤ�S�������~�C"
		Msg09 = "�S�����]�t���w���e���r��I"
	Else
		Msg01 = "ԭ��: "
		Msg02 = "����: "
		Msg03 = "����: "
		Msg04 = "��Ϣ: "
		Msg05 = "---------------------------------------"
		Msg06 = "�ҵ����д���:"
		Msg07 = "û���ҵ�����"
		Msg08 = "����ָ�����ݵ��ִ���û���ҵ�����"
		Msg09 = "û���ҵ�����ָ�����ݵ��ִ���"
	End If

	'������ʼ��
	CheckStrings = ""
	tText = ""
	k = 0
	LineNumErrCount = 0
	accKeyNumErrCount = 0
	Find = False

	'��ȡѡ���ķ����б�
	For i = 1 To trn.Project.TransLists.Count
		Set TrnList = trn.Project.TransLists(i)
		TrnListDec = TrnList.Title & " - " & PSL.GetLangCode(TrnList.Language.LangID,pslCodeText)
		If TrnListDec = ListDec Then Exit For
	Next i

	'��ȡĿ������
	trnLng = PSL.GetLangCode(TrnList.Language.LangID,pslCode639_1)
	If trnLng = "zh" Or trnLng = "ja" Or trnLng = "ko" Then TranLang = "Asia"

	If TrnList.SourceList.LastChange > TrnList.LastUpdate Then TrnList.Update
	If LineNum > TrnList.StringCount Then LineNum = TrnList.StringCount
	For i = 1 To TrnList.StringCount
		'������ʼ��
		srcString = ""
		trnString = ""
		NewtrnString = ""
		LineMsg = ""
		AccKeyMsg = ""
		ReplaceMsg = ""

		'��ȡԭ�ĺͷ����ִ�
		Set TransString = TrnList.String(i)
		If TransString.Text <> "" Then
			srcString = TransString.SourceText
			trnString = TransString.Text
			OldtrnString = trnString

			'��ʼ�����ִ�
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

			'������Ϣ���
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


'���ļ�
Function OpenFile(FilePath As String,FileList() As String,x As Integer,Stemp As Boolean) As Boolean
	Dim ExePathStr As String,Argument As String
	If OSLanguage = "0404" Then
		Msg01 = "���~"
		Msg02 = "�O�ƥ����b�t�Τ����C�п����L�}�Ҥ�k�C"
		Msg03 = "�t�ΨS���w�� Excel �M�ε{���C�п����L�}�Ҥ�k�C"
		Msg04 = "�{�������I�i��O�{�����|��R���~�Τ��Q�䴩�A�п����L�}�Ҥ�k�C"
		Msg05 = "�L�k�}���ɮסI�M�ε{����^�F���~�N�X�A�п����L�}�Ҥ�k�C"
		Msg06 = "�L�k�}���ɮסI�M�ε{����^�F���~�N�X�A�i��O���ɮפ��Q�䴩�ΰ���ѼƦ����D�C"
		Msg07 = "�{���W��: "
		Msg08 = "��R���|: "
		Msg09 = "����Ѽ�: "
	Else
		Msg01 = "����"
		Msg02 = "���±�δ��ϵͳ���ҵ�����ѡ�������򿪷�����"
		Msg03 = "ϵͳû�а�װ Excel Ӧ�ó�����ѡ�������򿪷�����"
		Msg04 = "����δ�ҵ��������ǳ���·����������򲻱�֧�֣���ѡ�������򿪷�����"
		Msg05 = "�޷����ļ���Ӧ�ó��򷵻��˴�����룬��ѡ�������򿪷�����"
		Msg06 = "�޷����ļ���Ӧ�ó��򷵻��˴�����룬�����Ǹ��ļ�����֧�ֻ����в��������⡣"
		Msg07 = "��������: "
		Msg08 = "����·��: "
		Msg09 = "���в���: "
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


'�༭�ļ�
Sub Edit(File As String,FileList() As String)
	If OSLanguage = "0404" Then
		Msg01 = "�s��"
		Msg02 = "�ɮ� - "
		Msg03 = "�r���s�X:"
		Msg05 = "�j�M���e:"
		Msg06 = "�j�M"
		Msg10 = "Ū�J(&R)"
		Msg12 = "�W�@��(&P)"
		Msg13 = "�U�@��(&N)"
		Msg14 = "�x�s(&S)"
		Msg16 = "�����j�M�Ҧ�"
	Else
		Msg01 = "�༭"
		Msg02 = "�ļ� - "
		Msg03 = "�ַ�����:"
		Msg05 = "��������:"
		Msg06 = "����"
		Msg10 = "����(&R)"
		Msg12 = "��һ��(&P)"
		Msg13 = "��һ��(&N)"
		Msg14 = "����(&S)"
		Msg16 = "�˳�����ģʽ"
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


'�༭�Ի�����
Private Function EditFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim File As String,inText As String,outText As String
	Dim Code As String,CodeID As Integer,m As Integer,n As Integer,j As Integer

	If OSLanguage = "0404" Then
		Msg01 = "�T��"
		Msg02 = "�ɮפ��e�w�Q�ܧ�A�O�_�ݭn�x�s�H"
		Msg03 = "��쪺���e�w�Q�ܧ�A�O�_�ݭn������Ӥ��e����ܡH"
		Msg04 = "��Ʀ��ܤơI�b�j�M�Ҧ��U�A�渹���i�R���A���e�i�R���M�ק�A�Эק��A�աC"
		Msg05 = "�������w���e�C"
		Msg06 = "�ɮ��x�s���\�I"
		Msg07 = "�ɮ��x�s���ѡI���ˬd�ɮ׬O�_���Q�}�ҡC"
		Msg08 = "�ɮפ��e���Q�ܧ�A���ݭn�x�s�I"
		Msg10 = "Ū�J(&R)"
		Msg11 = "�M��(&C)"
		LineNo = "��"
	Else
		Msg01 = "��Ϣ"
		Msg02 = "�ļ������ѱ����ģ��Ƿ���Ҫ���棿"
		Msg03 = "�ҵ��������ѱ����ģ��Ƿ���Ҫ�滻ԭ�����ݺ���ʾ��"
		Msg04 = "�����б仯���ڲ���ģʽ�£��кŲ���ɾ�������ݿ�ɾ�����޸ģ����޸ĺ����ԡ�"
		Msg05 = "δ�ҵ�ָ�����ݡ�"
		Msg06 = "�ļ�����ɹ���"
		Msg07 = "�ļ�����ʧ�ܣ������ļ��Ƿ������򿪡�"
		Msg08 = "�ļ�����δ�����ģ�����Ҫ���棡"
		Msg10 = "����(&R)"
		Msg11 = "���(&C)"
		LineNo = "��"
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
	Case 2 ' ��ֵ���Ļ��߰����˰�ť
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
					Temp = "��" & i+1 & LineNo & "��" & tempText
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
							Temp = LineNo & "��"
							LineNoStr = Left(NewString,InStr(NewString,Temp)+1)
							Temp = "��" & "*" & LineNo & "��"
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
			EditFunc = True '��ֹ���°�ť�رնԻ��򴰿�
		End If
	Case 3 ' �ı��������Ͽ��ı�������
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


'��Ӳ���ͬ������Ԫ��
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


'��ȥ�ִ�ǰ��ָ���� PreStr �� AppStr
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


'����༭����
Sub CmdInput(CmdPath As String,Argument As String)
	If OSLanguage = "0404" Then
		Msg01 = "�ۭq�s��{��"
		Msg02 = "�Ы��w�s��{���Ψ����Ѽ� (�ɮװѷӰѼƩM��L�Ѽ�)�C" & vbCrLf & vbCrLf & _
				"�`�N: " & vbCrLf & _
				"- �p�G����ѼƤ��ɮװѷӰѼƻݭn�b��L�Ѽƫe�����ܡA���I���k�䪺���s��J�A" & _
				"  �Ϊ�����J�ɮװѷӲ� %1�A�_�h�i�H����J�ɮװѷӰѼơC" & vbCrLf & _
				"- �ɮװѷӲ� %1 ��쬰�t�ΰѼơA���i�ܧ󬰨�L�Ÿ��C"
		Msg03 = "�s��{�� (�䴩�����ܼơA�ܼƦW�e��Ъ��[ % �Ÿ�):"
		Msg04 = "..."
		Msg05 = "����Ѽ� (�p�G�{���䴩�ûݭn����):"
		Msg06 = "�M��(&K)"
		Msg09 = ">"
	Else
		Msg01 = "�Զ���༭����"
		Msg02 = "��ָ���༭���������в��� (�ļ����ò�������������)��" & vbCrLf & vbCrLf & _
				"ע��: " & vbCrLf & _
				"- ������в������ļ����ò�����Ҫ����������ǰ��Ļ����뵥���ұߵİ�ť���룬" & _
				"  ��ֱ�������ļ����÷� %1��������Բ������ļ����ò�����" & vbCrLf & _
				"- �ļ����÷� %1 �ֶ�Ϊϵͳ���������ɸ���Ϊ�������š�"
		Msg03 = "�༭���� (֧�ֻ���������������ǰ���븽�� % ����):"
		Msg04 = "..."
		Msg05 = "���в��� (�������֧�ֲ���Ҫ�Ļ�):"
		Msg06 = "���(&K)"
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


'��ȡ�༭����Ի�����
Private Function CmdInputFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim File As String,Items(0) As String,x As Integer
	If OSLanguage = "0404" Then
		Msg01 = "���~"
		Msg02 = "����s��{��"
		Msg03 = "�i�����ɮ� (*.exe)|*.exe|�Ҧ��ɮ� (*.*)|*.*||"
		Msg04 = "�S�����w�s��{���I�Э��s��J�ο���C"
		FileArg = "�ɮװѷӰѼ�(%1)"
	Else
		Msg01 = "����"
		Msg02 = "ѡ��༭����"
		Msg03 = "��ִ���ļ� (*.exe)|*.exe|�����ļ� (*.*)|*.*||"
		Msg04 = "û��ָ���༭���������������ѡ��"
		FileArg = "�ļ����ò���(%1)"
	End If
	Items(0) = FileArg
	Select Case Action%
	Case 1 ' �Ի��򴰿ڳ�ʼ��
		DlgEnable "ClearButton",False
	Case 2 ' ��ֵ���Ļ��߰����˰�ť
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
				CmdInputFunc = True ' ��ֹ���°�ť�رնԻ��򴰿�
			End If
		End If
		If DlgItem$ <> "OKButton" And DlgItem$ <> "CancelButton" Then
 			If DlgText("CmdPath") = "" And DlgText("Argument") = "" Then
 				DlgEnable "ClearButton",False
 			Else
 				DlgEnable "ClearButton",True
 			End If
 			CmdInputFunc = True ' ��ֹ���°�ť�رնԻ��򴰿�
		End If
	Case 3 ' �ı��������Ͽ��ı�������
 		If DlgText("CmdPath") = "" And DlgText("Argument") = "" Then
 			DlgEnable "ClearButton",False
 		Else
 			DlgEnable "ClearButton",True
 		End If
	End Select
End Function


' �ļ����������Ϣ
Sub ErrorMassage(MsgType As String)
	If OSLanguage = "0404" Then
		Msg01 = "���~"
		msg02 = "�L�kŪ���ɮסI" & vbCrLf & _
				"�i��O�D��r�ɮשνs�X���Q�䴩�C" & vbCrLf & _
				"�нT�{�ɮ������ο���䥦�s�X (�w��) ��A�աC"
		Msg03 = "�L�k�g�J�ɮסI" & vbCrLf & _
				"���ˬd�ؼ��ɮ׬O�_�i�g�J�Φ��g�J�v���C"
	Else
		Msg01 = "����"
		msg02 = "�޷���ȡ�ļ���" & vbCrLf & _
				"�����Ƿ��ı��ļ�����벻��֧�֡�" & vbCrLf & _
				"��ȷ���ļ����ͻ�ѡ���������� (Ԥ��) �����ԡ�"
		Msg03 = "�޷�д���ļ���" & vbCrLf & _
				"����Ŀ���ļ��Ƿ��д����д��Ȩ�ޡ�"
	End If
	If MsgType = "NotReadFile" Then MsgBox(msg02,vbOkOnly+vbInformation,Msg01)
	If MsgType = "NotWriteFile" Then MsgBox(Msg03,vbOkOnly+vbInformation,Msg01)
End Sub


' ����ļ�����
' ----------------------------------------------------
' ANSI      �޸�ʽ����
' EFBB BF   UTF-8
' FFFE      UTF-16LE/UCS-2, Little Endian with BOM
' FEFF      UTF-16BE/UCS-2, Big Endian with BOM
' XX00 XX00 UTF-16LE/UCS-2, Little Endian without BOM
' 00XX 00XX UTF-16BE/UCS-2, Big Endian without BOM
' FFFE 0000 UTF-32LE/UCS-4, Little Endian with BOM
' 0000 FEFF UTF-32BE/UCS-4, Big Endian with BOM
' XX00 0000 UTF-32LE/UCS-4, Little Endian without BOM
' 0000 00XX UTF-32BE/UCS-4, Big Endian without BOM
' �����е� XX ��ʾ����ʮ�������ַ�

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


' ��ȡ�ļ�
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


' д���ļ�
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


'��������ҳ����
Public Function CodePageList(MinNum As Integer,MaxNum As Integer) As Variant
	Dim CodePage() As String,i As Integer,j As Integer
	ReDim CodePage(MaxNum - MinNum)
	For i = MinNum To MaxNum
		j = i - MinNum
		If OSLanguage = "0404" Then
			If i = 0 Then CodePage(j) = "�t�ιw�]" & JoinStr & "ANSI"
			If i = 1 Then CodePage(j) = "�۰ʿ��" & JoinStr & "_autodetect_all"
			If i = 2 Then CodePage(j) = "²�餤��(GB2312)" & JoinStr & "gb2312"
			If i = 3 Then CodePage(j) = "²�餤��(HZ)" & JoinStr & "hz-gb-2312"
			If i = 4 Then CodePage(j) = "²�餤��(GB18030)" & JoinStr & "gb18030"
			If i = 5 Then CodePage(j) = "���餤��(Big5)" & JoinStr & "big5"
			If i = 6 Then CodePage(j) = "���(EUC)" & JoinStr & "euc-jp"
			If i = 7 Then CodePage(j) = "���(JIS)" & JoinStr & "iso-2022-jp"
			If i = 8 Then CodePage(j) = "���(Shift-JIS)" & JoinStr & "shift_jis"
			If i = 9 Then CodePage(j) = "���(�۰ʿ��)" & JoinStr & "_autodetect"
			If i = 10 Then CodePage(j) = "����" & JoinStr & "ks_c_5601-1987"
			If i = 11 Then CodePage(j) = "����(EUC)" & JoinStr & "euc-kr"
			If i = 12 Then CodePage(j) = "����(ISO)" & JoinStr & "iso-2022-kr"
			If i = 13 Then CodePage(j) = "����(�۰ʿ��)" & JoinStr & "_autodetect_kr"
			If i = 14 Then CodePage(j) = "����(Windows)" & JoinStr & "windows-874"
			If i = 15 Then CodePage(j) = "�V�n��(Windows)" & JoinStr & "windows-1258"
			If i = 16 Then CodePage(j) = "�iù������(ISO)" & JoinStr & "iso-8859-4"
			If i = 17 Then CodePage(j) = "�iù������(Windows)" & JoinStr & "windows-1257"
			If i = 18 Then CodePage(j) = "���ԧB��(ASMO 708)" & JoinStr & "ASMO-708"
			If i = 19 Then CodePage(j) = "���ԧB��(DOS)" & JoinStr & "DOS-720"
			If i = 20 Then CodePage(j) = "���ԧB��(ISO)" & JoinStr & "iso-8859-6"
			If i = 21 Then CodePage(j) = "���ԧB��(Windows)" & JoinStr & "windows-1256"
			If i = 22 Then CodePage(j) = "�ƧB�Ӥ�(DOS)" & JoinStr & "DOS-862"
			If i = 23 Then CodePage(j) = "�ƧB�Ӥ�(ISO-�޿�)" & JoinStr & "iso-8859-8-i"
			If i = 24 Then CodePage(j) = "�ƧB�Ӥ�(ISO-��ı)" & JoinStr & "iso-8859-8"
			If i = 25 Then CodePage(j) = "�ƧB�Ӥ�(Windows)" & JoinStr & "windows-1255"
			If i = 26 Then CodePage(j) = "�g�ը��(Windows)" & JoinStr & "iso-8859-9"
			If i = 27 Then CodePage(j) = "��þ��(ISO)" & JoinStr & "iso-8859-7"
			If i = 28 Then CodePage(j) = "��þ��(Windows)" & JoinStr & "windows-1253"
			If i = 29 Then CodePage(j) = "���(Windows)" & JoinStr & "iso-8859-1"
			If i = 30 Then CodePage(j) = "��̺���(DOS)" & JoinStr & "cp866"
			If i = 31 Then CodePage(j) = "��̺���(ISO)" & JoinStr & "iso-8859-5"
			If i = 32 Then CodePage(j) = "��̺���(KOI8-R)" & JoinStr & "koi8-r"
			If i = 33 Then CodePage(j) = "��̺���(KOI8-U)" & JoinStr & "koi8-ru"
			If i = 34 Then CodePage(j) = "��̺���(Windows)" & JoinStr & "windows-1251"
			If i = 35 Then CodePage(j) = "����(DOS)" & JoinStr & "ibm852"
			If i = 36 Then CodePage(j) = "����(ISO)" & JoinStr & "iso-8859-2"
			If i = 37 Then CodePage(j) = "����(Windows)" & JoinStr & "windows-1250"
			If i = 38 Then CodePage(j) = "�ԤB�� 3 (ISO)" & JoinStr & "iso-8859-3"
			If i = 39 Then CodePage(j) = "Unicode (UTF-7)" & JoinStr & "utf-7"
			If i = 40 Then CodePage(j) = "Unicode (UTF-8 �� BOM)" & JoinStr & "utf-8EFBB"
			If i = 41 Then CodePage(j) = "Unicode (UTF-8 �L BOM)" & JoinStr & "utf-8"
			If i = 42 Then CodePage(j) = "Unicode (UTF-16LE �� BOM)" & JoinStr & "unicodeFFFE"
			If i = 43 Then CodePage(j) = "Unicode (UTF-16BE �� BOM)" & JoinStr & "unicodeFEFF"
			If i = 44 Then CodePage(j) = "Unicode (UTF-16LE �L BOM)" & JoinStr & "utf-16LE"
			If i = 45 Then CodePage(j) = "Unicode (UTF-16BE �L BOM)" & JoinStr & "utf-16BE"
			If i = 46 Then CodePage(j) = "Unicode (UTF-32LE �� BOM)" & JoinStr & "unicode-32FFFE"
			If i = 47 Then CodePage(j) = "Unicode (UTF-32BE �� BOM)" & JoinStr & "unicode-32FEFF"
			If i = 48 Then CodePage(j) = "Unicode (UTF-32LE �L BOM)" & JoinStr & "utf-32LE"
			If i = 49 Then CodePage(j) = "Unicode (UTF-32BE �L BOM)" & JoinStr & "utf-32BE"
		Else
			If i = 0 Then CodePage(j) = "ϵͳĬ��" & JoinStr & "ANSI"
			If i = 1 Then CodePage(j) = "�Զ�ѡ��" & JoinStr & "_autodetect_all"
			If i = 2 Then CodePage(j) = "��������(GB2312)" & JoinStr & "gb2312"
			If i = 3 Then CodePage(j) = "��������(HZ)" & JoinStr & "hz-gb-2312"
			If i = 4 Then CodePage(j) = "��������(GB18030)" & JoinStr & "gb18030"
			If i = 5 Then CodePage(j) = "��������(Big5)" & JoinStr & "big5"
			If i = 6 Then CodePage(j) = "����(EUC)" & JoinStr & "euc-jp"
			If i = 7 Then CodePage(j) = "����(JIS)" & JoinStr & "iso-2022-jp"
			If i = 8 Then CodePage(j) = "����(Shift-JIS)" & JoinStr & "shift_jis"
			If i = 9 Then CodePage(j) = "����(�Զ�ѡ��)" & JoinStr & "_autodetect"
			If i = 10 Then CodePage(j) = "����" & JoinStr & "ks_c_5601-1987"
			If i = 11 Then CodePage(j) = "����(EUC)" & JoinStr & "euc-kr"
			If i = 12 Then CodePage(j) = "����(ISO)" & JoinStr & "iso-2022-kr"
			If i = 13 Then CodePage(j) = "����(�Զ�ѡ��)" & JoinStr & "_autodetect_kr"
			If i = 14 Then CodePage(j) = "̩��(Windows)" & JoinStr & "windows-874"
			If i = 15 Then CodePage(j) = "Խ����(Windows)" & JoinStr & "windows-1258"
			If i = 16 Then CodePage(j) = "���޵ĺ���(ISO)" & JoinStr & "iso-8859-4"
			If i = 17 Then CodePage(j) = "���޵ĺ���(Windows)" & JoinStr & "windows-1257"
			If i = 18 Then CodePage(j) = "��������(ASMO 708)" & JoinStr & "ASMO-708"
			If i = 19 Then CodePage(j) = "��������(DOS)" & JoinStr & "DOS-720"
			If i = 20 Then CodePage(j) = "��������(ISO)" & JoinStr & "iso-8859-6"
			If i = 21 Then CodePage(j) = "��������(Windows)" & JoinStr & "windows-1256"
			If i = 22 Then CodePage(j) = "ϣ������(DOS)" & JoinStr & "DOS-862"
			If i = 23 Then CodePage(j) = "ϣ������(ISO-�߼�)" & JoinStr & "iso-8859-8-i"
			If i = 24 Then CodePage(j) = "ϣ������(ISO-�Ӿ�)" & JoinStr & "iso-8859-8"
			If i = 25 Then CodePage(j) = "ϣ������(Windows)" & JoinStr & "windows-1255"
			If i = 26 Then CodePage(j) = "��������(Windows)" & JoinStr & "iso-8859-9"
			If i = 27 Then CodePage(j) = "ϣ����(ISO)" & JoinStr & "iso-8859-7"
			If i = 28 Then CodePage(j) = "ϣ����(Windows)" & JoinStr & "windows-1253"
			If i = 29 Then CodePage(j) = "��ŷ(Windows)" & JoinStr & "iso-8859-1"
			If i = 30 Then CodePage(j) = "�������(DOS)" & JoinStr & "cp866"
			If i = 31 Then CodePage(j) = "�������(ISO)" & JoinStr & "iso-8859-5"
			If i = 32 Then CodePage(j) = "�������(KOI8-R)" & JoinStr & "koi8-r"
			If i = 33 Then CodePage(j) = "�������(KOI8-U)" & JoinStr & "koi8-ru"
			If i = 34 Then CodePage(j) = "�������(Windows)" & JoinStr & "windows-1251"
			If i = 35 Then CodePage(j) = "��ŷ(DOS)" & JoinStr & "ibm852"
			If i = 36 Then CodePage(j) = "��ŷ(ISO)" & JoinStr & "iso-8859-2"
			If i = 37 Then CodePage(j) = "��ŷ(Windows)" & JoinStr & "windows-1250"
			If i = 38 Then CodePage(j) = "������ 3 (ISO)" & JoinStr & "iso-8859-3"
			If i = 39 Then CodePage(j) = "Unicode (UTF-7)" & JoinStr & "utf-7"
			If i = 40 Then CodePage(j) = "Unicode (UTF-8 �� BOM)" & JoinStr & "utf-8EFBB"
			If i = 41 Then CodePage(j) = "Unicode (UTF-8 �� BOM)" & JoinStr & "utf-8"
			If i = 42 Then CodePage(j) = "Unicode (UTF-16LE �� BOM)" & JoinStr & "unicodeFFFE"
			If i = 43 Then CodePage(j) = "Unicode (UTF-16BE �� BOM)" & JoinStr & "unicodeFEFF"
			If i = 44 Then CodePage(j) = "Unicode (UTF-16LE �� BOM)" & JoinStr & "utf-16LE"
			If i = 45 Then CodePage(j) = "Unicode (UTF-16BE �� BOM)" & JoinStr & "utf-16BE"
			If i = 46 Then CodePage(j) = "Unicode (UTF-32LE �� BOM)" & JoinStr & "unicode-32FFFE"
			If i = 47 Then CodePage(j) = "Unicode (UTF-32BE �� BOM)" & JoinStr & "unicode-32FEFF"
			If i = 48 Then CodePage(j) = "Unicode (UTF-32LE �� BOM)" & JoinStr & "utf-32LE"
			If i = 49 Then CodePage(j) = "Unicode (UTF-32BE �� BOM)" & JoinStr & "utf-32BE"
		End If
	Next i
	CodePageList = CodePage
End Function


'��ȡ���Զ�
Function LangCodeList(DataName As String,OSLang As String,MinNum As Integer,MaxNum As Integer) As Variant
	Dim i As Integer,j As Integer,Code As String,LangName() As String,LangPair() As String
	ReDim LangName(MaxNum - MinNum),LangPair(MaxNum - MinNum)
	For i = MinNum To MaxNum
		j = i - MinNum
		If OSLang = "0404" Then
			If i = 0 Then LangName(j) = "�۰ʰ���"
			If i = 1 Then LangName(j) = "�n�D�����y"
			If i = 2 Then LangName(j) = "�����ڥ��Ȼy"
			If i = 3 Then LangName(j) = "���i���Իy"
			If i = 4 Then LangName(j) = "���ԧB�y"
			If i = 5 Then LangName(j) = "�Ȭ����Ȼy"
			If i = 6 Then LangName(j) = "���ĩi�y"
			If i = 7 Then LangName(j) = "�����æ�y"
			If i = 8 Then LangName(j) = "�ڤ��򺸻y"
			If i = 9 Then LangName(j) = "�ڴ��J�y"
			If i = 10 Then LangName(j) = "�իXù���y"
			If i = 11 Then LangName(j) = "�s�[�Իy"
			If i = 12 Then LangName(j) = "�i�襧�Ȼy"
			If i = 13 Then LangName(j) = "���C�𥧻y"
			If i = 14 Then LangName(j) = "�O�[�Q�Ȼy"
			If i = 15 Then LangName(j) = "�[��ù���Ȼy"
			If i = 16 Then LangName(j) = "²�餤��"
			If i = 17 Then LangName(j) = "���餤��"
			If i = 18 Then LangName(j) = "���Ży"
			If i = 19 Then LangName(j) = "�Jù�a�Ȼy"
			If i = 20 Then LangName(j) = "���J�y"
			If i = 21 Then LangName(j) = "�����y"
			If i = 22 Then LangName(j) = "�����y"
			If i = 23 Then LangName(j) = "�^�y"
			If i = 24 Then LangName(j) = "�R�F���Ȼy"
			If i = 25 Then LangName(j) = "�kù�y"
			If i = 26 Then LangName(j) = "�i���y"
			If i = 27 Then LangName(j) = "�����y"
			If i = 28 Then LangName(j) = "�k�y"
			If i = 29 Then LangName(j) = "���̦�Ȼy"
			If i = 30 Then LangName(j) = "�[�Q��Ȼy"
			If i = 31 Then LangName(j) = "��|�N�Ȼy"
			If i = 32 Then LangName(j) = "�w�y"
			If i = 33 Then LangName(j) = "��þ�y"
			If i = 34 Then LangName(j) = "�泮���y"
			If i = 35 Then LangName(j) = "�j�N�ԯS�y"
			If i = 36 Then LangName(j) = "���Ļy"
			If i = 37 Then LangName(j) = "�ƧB�ӻy"
			If i = 38 Then LangName(j) = "�L�a�y"
			If i = 39 Then LangName(j) = "�I���Q�y"
			If i = 40 Then LangName(j) = "�B�q�y"
			If i = 41 Then LangName(j) = "�L�ץ���Ȼy"
			If i = 42 Then LangName(j) = "�]�ïS�y"
			If i = 43 Then LangName(j) = "�R�����y"
			If i = 44 Then LangName(j) = "�Z�ϻy"
			If i = 45 Then LangName(j) = "���|�y"
			If i = 46 Then LangName(j) = "�N�j�Q�y"
			If i = 47 Then LangName(j) = "��y"
			If i = 48 Then LangName(j) = "�d�ǹF�y"
			If i = 49 Then LangName(j) = "�J���̺��y"
			If i = 50 Then LangName(j) = "���ħJ�y"
			If i = 51 Then LangName(j) = "���ֻy"
			If i = 52 Then LangName(j) = "�c���F�y"
			If i = 53 Then LangName(j) = "�եd���y"
			If i = 54 Then LangName(j) = "���A�y"
			If i = 55 Then LangName(j) = "�N���N���y"
			If i = 56 Then LangName(j) = "�N���N���y (�N���N���Z)"
			If i = 57 Then LangName(j) = "�Ѿ�y"
			If i = 58 Then LangName(j) = "�Բ���Ȼy"
			If i = 59 Then LangName(j) = "�߳��{�y"
			If i = 60 Then LangName(j) = "�c�˳��y"
			If i = 61 Then LangName(j) = "����y�y"
			If i = 62 Then LangName(j) = "���ӻy"
			If i = 63 Then LangName(j) = "���Զ��ԩi�y"
			If i = 64 Then LangName(j) = "���եL�y"
			If i = 65 Then LangName(j) = "��Q�y"
			If i = 66 Then LangName(j) = "���Ԧa�y"
			If i = 67 Then LangName(j) = "�X�j�y"
			If i = 68 Then LangName(j) = "���y���y"
			If i = 69 Then LangName(j) = "���»y"
			If i = 70 Then LangName(j) = "���»y (�էJ������)"
			If i = 71 Then LangName(j) = "���»y (���մ��J��)"
			If i = 72 Then LangName(j) = "���̶��y"
			If i = 73 Then LangName(j) = "�����ϻy"
			If i = 74 Then LangName(j) = "�i���y"
			If i = 75 Then LangName(j) = "������y"
			If i = 76 Then LangName(j) = "�ǾB���y"
			If i = 77 Then LangName(j) = "�J�C�Ȼy"
			If i = 78 Then LangName(j) = "ù�����Ȼy"
			If i = 79 Then LangName(j) = "�X�y"
			If i = 80 Then LangName(j) = "�Ħ̻y"
			If i = 81 Then LangName(j) = "��y"
			If i = 82 Then LangName(j) = "�뺸���Ȼy"
			If i = 83 Then LangName(j) = "�گ����y"
			If i = 84 Then LangName(j) = "���˯ǻy"
			If i = 85 Then LangName(j) = "�H�w�y"
			If i = 86 Then LangName(j) = "����ù�y"
			If i = 87 Then LangName(j) = "������J�y"
			If i = 88 Then LangName(j) = "�����奧�Ȼy"
			If i = 89 Then LangName(j) = "��Z���y"
			If i = 90 Then LangName(j) = "���˧Ƹ̻y"
			If i = 91 Then LangName(j) = "���y"
			If i = 92 Then LangName(j) = "�ԧQ�Ȼy"
			If i = 93 Then LangName(j) = "��N�J�y"
			If i = 94 Then LangName(j) = "���̺��y"
			If i = 95 Then LangName(j) = "Ŷ�޻y"
			If i = 96 Then LangName(j) = "���c�T�y"
			If i = 97 Then LangName(j) = "���y"
			If i = 98 Then LangName(j) = "�ûy"
			If i = 99 Then LangName(j) = "�g�ը�y"
			If i = 100 Then LangName(j) = "�g�w�һy"
			If i = 101 Then LangName(j) = "���^���y"
			If i = 102 Then LangName(j) = "�Q�J���y"
			If i = 103 Then LangName(j) = "�Q�����y"
			If i = 104 Then LangName(j) = "�Q���O�J�y"
			If i = 105 Then LangName(j) = "�V�n�y"
			If i = 106 Then LangName(j) = "�º��h�y"
			If i = 107 Then LangName(j) = "�U���һy"
		Else
			If i = 0 Then LangName(j) = "�Զ����"
			If i = 1 Then LangName(j) = "�ϷǺ�����"
			If i = 2 Then LangName(j) = "������������"
			If i = 3 Then LangName(j) = "��ķ������"
			If i = 4 Then LangName(j) = "��������"
			If i = 5 Then LangName(j) = "����������"
			If i = 6 Then LangName(j) = "����ķ��"
			If i = 7 Then LangName(j) = "�����ݽ���"
			If i = 8 Then LangName(j) = "��ʲ������"
			If i = 9 Then LangName(j) = "��˹����"
			If i = 10 Then LangName(j) = "�׶���˹��"
			If i = 11 Then LangName(j) = "�ϼ�����"
			If i = 12 Then LangName(j) = "����������"
			If i = 13 Then LangName(j) = "����������"
			If i = 14 Then LangName(j) = "����������"
			If i = 15 Then LangName(j) = "��̩��������"
			If i = 16 Then LangName(j) = "��������"
			If i = 17 Then LangName(j) = "��������"
			If i = 18 Then LangName(j) = "��������"
			If i = 19 Then LangName(j) = "���޵�����"
			If i = 20 Then LangName(j) = "�ݿ���"
			If i = 21 Then LangName(j) = "������"
			If i = 22 Then LangName(j) = "������"
			If i = 23 Then LangName(j) = "Ӣ��"
			If i = 24 Then LangName(j) = "��ɳ������"
			If i = 25 Then LangName(j) = "������"
			If i = 26 Then LangName(j) = "��˹��"
			If i = 27 Then LangName(j) = "������"
			If i = 28 Then LangName(j) = "����"
			If i = 29 Then LangName(j) = "����������"
			If i = 30 Then LangName(j) = "����������"
			If i = 31 Then LangName(j) = "��³������"
			If i = 32 Then LangName(j) = "����"
			If i = 33 Then LangName(j) = "ϣ����"
			If i = 34 Then LangName(j) = "��������"
			If i = 35 Then LangName(j) = "�ż�������"
			If i = 36 Then LangName(j) = "������"
			If i = 37 Then LangName(j) = "ϣ������"
			If i = 38 Then LangName(j) = "ӡ����"
			If i = 39 Then LangName(j) = "��������"
			If i = 40 Then LangName(j) = "������"
			If i = 41 Then LangName(j) = "ӡ����������"
			If i = 42 Then LangName(j) = "��Ŧ����"
			If i = 43 Then LangName(j) = "��������"
			If i = 44 Then LangName(j) = "��ͼ��"
			If i = 45 Then LangName(j) = "��³��"
			If i = 46 Then LangName(j) = "�������"
			If i = 47 Then LangName(j) = "����"
			If i = 48 Then LangName(j) = "���ɴ���"
			If i = 49 Then LangName(j) = "��ʲ�׶���"
			If i = 50 Then LangName(j) = "��������"
			If i = 51 Then LangName(j) = "������"
			If i = 52 Then LangName(j) = "¬������"
			If i = 53 Then LangName(j) = "�׿�����"
			If i = 54 Then LangName(j) = "������"
			If i = 55 Then LangName(j) = "������˹��"
			If i = 56 Then LangName(j) = "������˹�� (������˹̹)"
			If i = 57 Then LangName(j) = "������"
			If i = 58 Then LangName(j) = "����ά����"
			If i = 59 Then LangName(j) = "��������"
			If i = 60 Then LangName(j) = "¬ɭ����"
			If i = 61 Then LangName(j) = "�������"
			If i = 62 Then LangName(j) = "������"
			If i = 63 Then LangName(j) = "��������ķ��"
			If i = 64 Then LangName(j) = "�������"
			If i = 65 Then LangName(j) = "ë����"
			If i = 66 Then LangName(j) = "��������"
			If i = 67 Then LangName(j) = "�ɹ���"
			If i = 68 Then LangName(j) = "�Ჴ����"
			If i = 69 Then LangName(j) = "Ų����"
			If i = 70 Then LangName(j) = "Ų���� (���������)"
			If i = 71 Then LangName(j) = "Ų���� (��ŵ˹����)"
			If i = 72 Then LangName(j) = "��������"
			If i = 73 Then LangName(j) = "��ʲͼ��"
			If i = 74 Then LangName(j) = "������"
			If i = 75 Then LangName(j) = "��������"
			If i = 76 Then LangName(j) = "��������"
			If i = 77 Then LangName(j) = "��������"
			If i = 78 Then LangName(j) = "����������"
			If i = 79 Then LangName(j) = "����"
			If i = 80 Then LangName(j) = "������"
			If i = 81 Then LangName(j) = "����"
			If i = 82 Then LangName(j) = "����ά����"
			If i = 83 Then LangName(j) = "��������"
			If i = 84 Then LangName(j) = "��������"
			If i = 85 Then LangName(j) = "�ŵ���"
			If i = 86 Then LangName(j) = "ɮ٤����"
			If i = 87 Then LangName(j) = "˹�工����"
			If i = 88 Then LangName(j) = "˹����������"
			If i = 89 Then LangName(j) = "��������"
			If i = 90 Then LangName(j) = "˹��ϣ����"
			If i = 91 Then LangName(j) = "�����"
			If i = 92 Then LangName(j) = "��������"
			If i = 93 Then LangName(j) = "��������"
			If i = 94 Then LangName(j) = "̩�׶���"
			If i = 95 Then LangName(j) = "������"
			If i = 96 Then LangName(j) = "̩¬����"
			If i = 97 Then LangName(j) = "̩��"
			If i = 98 Then LangName(j) = "����"
			If i = 99 Then LangName(j) = "��������"
			If i = 100 Then LangName(j) = "��������"
			If i = 101 Then LangName(j) = "ά�����"
			If i = 102 Then LangName(j) = "�ڿ�����"
			If i = 103 Then LangName(j) = "�ڶ�����"
			If i = 104 Then LangName(j) = "���ȱ����"
			If i = 105 Then LangName(j) = "Խ����"
			If i = 106 Then LangName(j) = "����ʿ��"
			If i = 107 Then LangName(j) = "�������"
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


'������
Sub CheckHelp(HelpTip As String)
	If OSLanguage = "0404" Then
	AboutTitle = "����"
	HelpTitle = "����"
	HelpTipTitle = "�K����B�פ�ũM�[�t���ˬd����"
	AboutWindows = " ���� "
	MainWindows = " �D���� "
	SetWindows = " �]�w���� "
	TestWindows = " ���յ��� "
	Lines = "-----------------------"
	Sys = "�n�骩���G" & Version & vbCrLf & _
			"�A�Ψt�ΡGWindows XP/2000 �H�W�t��" & vbCrLf & _
			"�A�Ϊ����G�Ҧ��䴩�����B�z�� Passolo 6.0 �ΥH�W����" & vbCrLf & _
			"�����y���G²�餤��M���餤�� (�۰ʿ���)" & vbCrLf & _
			"���v�Ҧ��G�~�Ʒs�@��" & vbCrLf & _
			"���v�Φ��G�K�O�n��" & vbCrLf & _
			"�x�譺���Ghttp://www.hanzify.org" & vbCrLf & _
			"�e�}�o�̡G�~�Ʒs�@������ gnatix (2007-2008)" & vbCrLf & _
			"��}�o�̡G�~�Ʒs�@������ wanfu (2009-2010)" & vbCrLf & vbCrLf & vbCrLf
	Ement = "���������ҡ�" & vbCrLf & _
			"============" & vbCrLf & _
			"- �䴩�����B�z�� Passolo 6.0 �ΥH�W�����A����" & vbCrLf & _
			"- Windows Script Host (WSH) ���� (VBS)�A����" & vbCrLf & _
			"- Adodb.Stream ���� (VBS)�A�䴩�۰ʧ�s�һ�" & vbCrLf & _
			"- Microsoft.XMLHTTP ����A�䴩�۰ʧ�s�һ�" & vbCrLf & vbCrLf & vbCrLf
	Dec = "���n��²����" & vbCrLf & _
			"============" & vbCrLf & _
			"�K����B�פ�ũM�[�t���ˬd�����O�@�ӥΩ� Passolo ½Ķ�ˬd�������{���C���㦳�H�U�\��G" & vbCrLf & _
			"- �ˬd½Ķ���K����B�פ�šB�[�t���M�Ů�" & vbCrLf & _
			"- �ˬd�íץ��ˬd½Ķ���K����B�פ�šB�[�t���M�Ů�" & vbCrLf & _
			"- �R��½Ķ�����K����" & vbCrLf & _
			"- ���m�i�ۭq���۰ʧ�s�\��" & vbCrLf & vbCrLf & _
			"���{���]�t�U�C�ɮסG" & vbCrLf & _
			"- �۰ʥ����GPslAutoAccessKey.bas" & vbCrLf & _
			"  �b½Ķ�r��ɡA�۰ʧ󥿿��~��½Ķ�C�Q�θӥ����A�z�i�H������J�K����B�פ�šB�[�t���A" & vbCrLf & _
			"  �t�αN�ھڱz������]�w�۰����z�s�W�M���@�˪��K����B�פ�šB�[�t���A��½Ķ�פ�šC" & vbCrLf & _
			"  �Q�Υ��i�H����½Ķ�t�סA�ô��½Ķ���~�C" & vbCrLf & _
			"  ���`�N�G�ѩ� Passolo ������A�ӥ������]�w�ݳq�L�����ˬd�����ӿ���M�]�w�C" & vbCrLf & vbCrLf & _
			"- �ˬd�����GPSLCheckAccessKeys.bas" & vbCrLf & _
			"  �q�L�I�s�s�W�� Passolo ��椤���ӥ����A��U�z�ˬd�M�ץ�½Ķ�����K����B�פ�šB�[�t���A" & vbCrLf & _
			"  ��½Ķ�פ�šC���~�A���ٴ��Ѧۭq�]�w�M�]�w�����\��C" & vbCrLf & vbCrLf & _
			"- ²�餤�廡���ɮסGAccessKey.txt" & vbCrLf & vbCrLf & vbCrLf
	Setup = "���w�ˤ�k��" & vbCrLf & _
			"============" & vbCrLf & _
			"- �p�G�ϥΤF Wanfu �� Passolo �~�ƪ��A�æw�ˤF���[�����ե�A�h������Ӫ��ɮקY�i�A�_�h�G" & vbCrLf & _
			"  (1) �N�����᪺�ɮ׽ƻs�� Passolo �t�θ�Ƨ����w�q�� Macros ��Ƨ���" & vbCrLf & _
			"  (2) �۰ʥ����G�}�� Passolo ���u�� -> ������ܤ���A�N���]�w���t�Υ������I���D�������k�U�����t��" & vbCrLf & _
			"  �@  �����ҥο��ҥΥ�" & vbCrLf & _
			"  (3) �ˬd�����G�b Passolo ���u�� -> �ۭq�u���椤�s�W���ɮ�" & vbCrLf & _
			"- �ѩ�۰ʥ����L�k�b����L�{���i��]�w�A�ҥH�Шϥ��ˬd�����Ӧۭq�]�w�M�סC" & vbCrLf & _
			"- ���ˬd��ȥ��A�v����u�Ƭd�A�H�K�{���B�z���~�C" & vbCrLf & vbCrLf & vbCrLf
	MainUse = "���]�w�����" & vbCrLf & _
			"============" & vbCrLf & _
			"�{�����ѤF�w�]���]�w�A�ӳ]�w�i�H�A�Ω�j�h�Ʊ��p�C�z�]�i�H�� [�]�w] ���s�ۭq�]�w�C" & vbCrLf & _
			"�s�W�ۭq�]�w��A�z�i�H�b�]�w�M�椤����Q�ϥΪ��]�w�C" & vbCrLf & _
			"�����ۭq�]�w�A�ж}�ҳ]�w��ܤ���A�I�� [����] ���s�A�Ѿ\�������������C" & vbCrLf & vbCrLf & _
			"- �۰ʥ����]�w" & vbCrLf & _
			"  ������]�w�N�Ω�۰ʥ����C�Ъ`�N�x�s�A���M�N�ϥο���e���]�w�C" & vbCrLf & vbCrLf & _
			"- �ˬd�����]�w" & vbCrLf & _
			"  ������]�w�N�Ω��ˬd�����C�n�b�U���ϥο�����]�w�A�ݭn�x�s�C" & vbCrLf & vbCrLf & _
			"- �۰ʥ����M�ˬd�����ۦP" & vbCrLf & _
			"  ����ӿﶵ�ɡA�N�۰ʨϦ۰ʥ������]�w�P�ˬd�������]�w�@�P�C" & vbCrLf & vbCrLf & _
			"- �۰ʿ��" & vbCrLf & _
			"  ��ܸӿﶵ�ɱN�ھڳ]�w�����A�λy���M��۰ʿ���P½Ķ�M��ؼлy���ŦX���]�w�C" & vbCrLf & _
			"  ���`�N�G�ӿﶵ�ȹ�K����B�פ�ũM�[�t���ˬd�������ġC" & vbCrLf & _
			"  �@�@�@�@�n�N�]�w�P�ثe½Ķ�M�檺�ؼлy���ŦX�A�� [�]�w] ���s�A�b�r��B�z���A�λy����" & vbCrLf & _
			"  �@�@�@�@�s�W�������y���C" & vbCrLf & _
			"  �@�@�@�@�۰ʥ����{���N���u�۰� - �ۿ� - �w�]�v���ǿ���������]�w�C" & vbCrLf & vbCrLf & _
			"���ˬd�аO��" & vbCrLf & _
			"============" & vbCrLf & _
			"�ӥ\��q�L�O���r���ˬd�T���A�îھڸӰT���u����~�r��i���ˬd�A�i�j�T�����A�ˬd�t�סC" & vbCrLf & _
			"���H�U 4 �ӿﶵ�i�ѿ���G" & vbCrLf & vbCrLf & _
			"- ��������" & vbCrLf & _
			"  �N���Ҽ{�����{���������A�Ȯھڨ䥦�O������~�r��i���ˬd�C" & vbCrLf & vbCrLf & _
			"- �����]�w" & vbCrLf & _
			"  �N���Ҽ{�]�w�O�_�ۦP�A�Ȯھڨ䥦�O������~�r��i���ˬd�C" & vbCrLf & vbCrLf & _
			"- �������" & vbCrLf & _
			"  �N���Ҽ{�ˬd����M½Ķ����A�Ȯھڨ䥦�O������~�r��i���ˬd�C" & vbCrLf & vbCrLf & _
			"- ��������" & vbCrLf & _
			"  �N���Ҽ{�����ˬd�O���A�ӹ�Ҧ��r��i���ˬd�C" & vbCrLf & vbCrLf & _
			"���`�N�G�ˬd�аO�\��b���ճ]�w�ɵL�ġA�H�K�i��X�{����|�����յ��G�C" & vbCrLf & _
			"�@�@�@�@�p�G�ܧ�]�w���e�Ӥ��ܧ�]�w�W�٪��ܡA�п�ܥ��������ΧR���ˬd�аO�ﶵ�C" & vbCrLf & vbCrLf & _
			"���]�w�����" & vbCrLf & _
			"============" & vbCrLf & _
			"���H�U 3 �ӿﶵ�i�ѿ���G" & vbCrLf & vbCrLf & _
			"- ���ˬd" & vbCrLf & _
			"  �u��½Ķ�i���ˬd�A�Ӥ��ץ����~��½Ķ�C" & vbCrLf & vbCrLf & _
			"- �ˬd�íץ�" & vbCrLf & _
			"  ��½Ķ�i���ˬd�A�æ۰ʭץ����~��½Ķ�C" & vbCrLf & vbCrLf & _
			"- �R���K����" & vbCrLf & _
			"  �R��½Ķ���{�����K����C" & vbCrLf & vbCrLf & _
			"���r��������" & vbCrLf & _
			"============" & vbCrLf & _
			"���ѤF�����B���B��ܤ���B�r���B�[�t���B�����B��L�B�ȿ�ܵ��ﶵ�C" & vbCrLf & vbCrLf & _
			"- �p�G��������A�h��L�涵�N�Q�۰ʨ�������C" & vbCrLf & _
			"- �p�G����涵�A�h�����ﶵ�N�Q�۰ʨ�������C" & vbCrLf & _
			"- �涵�i�H�h��C�䤤����ȿ�ܮɡA��L���Q�۰ʨ�������C" & vbCrLf & vbCrLf & _
			"���r�ꤺ�e��" & vbCrLf & _
			"============" & vbCrLf & _
			"���ѤF�����B�K����B�פ�šB�[�t�� 4 �ӿﶵ�C" & vbCrLf & vbCrLf & _
			"- �p�G��������A�h��L�涵�N�Q�۰ʨ�������C" & vbCrLf & _
			"- �p�G����涵�A�h�����ﶵ�N�Q�۰ʨ�������C" & vbCrLf & _
			"- �涵�i�H�h��C" & vbCrLf & vbCrLf & _
			"����L�ﶵ��" & vbCrLf & _
			"============" & vbCrLf & _
			"- ���ܧ��l½Ķ���A" & vbCrLf & _
			"  ��ܸӶ��ɡA�N�b�ˬd�B�ˬd�íץ��B�R���K����ɤ��ܧ�r�ꪺ��l½Ķ���A�C�_�h�A" & vbCrLf & _
			"  �N�ܧ�L���~�εL�ܧ�r�ꪺ½Ķ���A���w���Ҫ��A�A�����~�Τw�ܧ�r�ꪺ½Ķ���A" & vbCrLf & _
			"  ���ݽƼf���A�A�H�K�z�@���N�i�H���D���Ǧr�꦳���~�Τw�Q�ܧ�C" & vbCrLf & vbCrLf & _
			"- ���إߩΧR���ˬd�аO" & vbCrLf & _
			"  ��ܸӶ��ɡA�N���b�M�פ��x�s�ˬd�аO�T���A�p�G�w�s�b�ˬd�аO�T���A�N�Q�R���C" & vbCrLf & _
			"  ���`�N�G��ܸӶ��ɡA�ˬd�аO������������������N�Q��ܡC" & vbCrLf & vbCrLf & _
			"- �~��ɦ۰��x�s�Ҧ����" & vbCrLf & _
			"  ��ܸӶ��ɡA�N�b�� [�~��] ���s�ɦ۰��x�s�Ҧ�����A�U������ɱNŪ�J�x�s������C" & vbCrLf & _
			"  ���`�N�G�p�G�۰ʥ����]�w�Q�ܧ�A�t�αN�۰ʿ�ܸӿﶵ�A�H�Ͽ�ܪ��۰ʥ����]�w�ͮġC" & vbCrLf & vbCrLf & _
			"- �����S�w�r��" & vbCrLf & _
  			"  �b�ˬd�íץ��L�{���ϥγ]�w���w�q���n�۰ʴ������r���A�����r�ꤤ�S�w���r���C" & vbCrLf & vbCrLf & _
			"����L�\�ࡸ" & vbCrLf & _
			"============" & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �I���ӫ��s�A�N�����ܤ���C����ܵ{�����СB�������ҡB�}�o�ӤΪ��v���T���C" & vbCrLf& vbCrLf & _
			"- ����" & vbCrLf & _
			"  �I���ӫ��s�A�N�u�X�ثe�����������T���C" & vbCrLf& vbCrLf & _
			"- �x�s�Ҧ����" & vbCrLf & _
			"  �ӫ��s�i�H�b�ܧ�]�w�Ӥ��i���ˬd�ɨϥΡC" & vbCrLf & _
			"  �p�G���@�ﶵ�Q�ܧ�A�ӿﶵ�N�۰��ܬ��i�Ϊ��A�A�_�h�N�۰��ܬ����i�Ϊ��A�C" & vbCrLf & vbCrLf & _
			"- �T�w" & vbCrLf & _
			"  �I���ӫ��s�A�N�����D��ܤ���A�ë���ܪ��ﶵ�i��r��½Ķ�C" & vbCrLf& vbCrLf & _
			"- ����" & vbCrLf & _
			"  �I���ӫ��s�A�N�����{���C" & vbCrLf& vbCrLf & vbCrLf
	SetUse ="���]�w�M�桸" & vbCrLf & _
			"============" & vbCrLf & _
			"- ����]�w" & vbCrLf & _
			"  �n����]�w�A�I���]�w�M��C" & vbCrLf & vbCrLf & _
			"- �]�w�]�w���u����" & vbCrLf & _
			"  �]�w�u���ťΩ���]�w���A�λy�����۰ʿ���]�w�\��C" & vbCrLf & _
			"  ���`�N�G���h�ӳ]�w�]�t�F�ۦP���A�λy���ɡA�ݭn�]�w���u���šC" & vbCrLf & _
			"  �@�@�@�@�b�ۦP�A�λy�����]�w���A�e�����]�w�Q�u������ϥΡC" & vbCrLf & _
			"  �n�]�w�]�w���u���šA�I���k�䪺 [...] ���s�C" & vbCrLf & vbCrLf & _
			"- �s�W�]�w" & vbCrLf & _
			"  �n�s�W�]�w�A�I�� [�s�W] ���s�A�b�u�X����ܤ������J�W�١C" & vbCrLf & vbCrLf & _
			"- �ܧ�]�w" & vbCrLf & _
			"  �n�ܧ�]�w�W�١A�п���]�w�M�椤�n��W���]�w�A�M���I�� [�ܧ�] ���s�C" & vbCrLf & vbCrLf & _
			"- �R���]�w" & vbCrLf & _
			"  �n�R���]�w�A�п���]�w�M�椤�n�R�����]�w�A�M���I�� [�R��] ���s�C" & vbCrLf & vbCrLf & _
			"�s�W�]�w��A�N�b�M�椤��ܷs���]�w�A�]�w���e�N��ܪŭȡC" & vbCrLf & _
			"�ܧ�]�w��A�N�b�M�椤��ܧ�W���]�w�A�]�w���e�����]�w�Ȥ��ܡC" & vbCrLf & _
			"�R���]�w��A�N�b�M�椤��ܹw�]�]�w�A�]�w���e�N��ܹw�]�]�w�ȡC" & vbCrLf & vbCrLf & _
			"���x�s������" & vbCrLf & _
			"============" & vbCrLf & _
			"- �ɮ�" & vbCrLf & _
			"  �]�w�N�H�ɮקΦ��x�s�b�����Ҧb��Ƨ��U�� Data ��Ƨ����C" & vbCrLf & vbCrLf & _
			"- ���U��" & vbCrLf & _
			"  �]�w�N�Q�x�s���U���� HKCU\Software\VB and VBA Program Settings\AccessKey ���U�C" & vbCrLf & vbCrLf & _
			"- �פJ�]�w" & vbCrLf & _
			"  ���\�q��L�]�w�ɮפ��פJ�]�w�C�פJ�³]�w�ɱN�Q�۰ʤɯšA�{���]�w�M�椤�w�����]�w�N�Q" & vbCrLf & _
			"  �ܧ�A�S�����]�w�N�Q�s�W�C" & vbCrLf & vbCrLf & _
			"- �ץX�]�w" & vbCrLf & _
			"  ���\�ץX�Ҧ��]�w���r�ɮסA�H�K�i�H�洫���ಾ�]�w�C" & vbCrLf & vbCrLf & _
			"���`�N�G�����x�s�����ɡA�N�۰ʧR���즳��m�����]�w���e�C" & vbCrLf & vbCrLf & _
			"���]�w���e��" & vbCrLf & _
			"============" & vbCrLf & _
			"<�K����>" & vbCrLf & _
			"  - �n�ư����t & �Ÿ����D�K����զX" & vbCrLf & _
			"    �K����H & ���X�вšA���Ǧr�����M�]�t�ӲŸ������O�K����A�ݭn�ư����C�b����J�o��" & vbCrLf & _
			"    �n�ư��]�t & �Ÿ����D�K����զX�C" & vbCrLf & vbCrLf & _
			"  - �r����ΥκX�в�" & vbCrLf & _
			"    �o�Ǧr���Ω���Χt���h�ӫK����A�פ�ũΥ[�t�����r��A�H�K�ˬd�r�ꤤ�Ҧ����K����B" & vbCrLf & _
			"    �פ�ũΥ[�t���C�_�h�u��B�z�r��̫᳡�����K����B�פ�ũΥ[�t���C" & vbCrLf & vbCrLf & _
			"  - �n�ˬd���K����e��A��" & vbCrLf & _
			"    �w�]���K����e��A���� ()�A�b�����w���K����e��A���A���N�Q�������w�]���A���C" & vbCrLf & vbCrLf & _
			"  - �n�O�d���D�K����e�ᦨ��r��" & vbCrLf & _
			"    �p�G���M½Ķ�����s�b�o�Ǧr���C�N�Q�O�d�A�_�h�N�Q�{���O�K����A�ñN�s�W�K����ò�" & vbCrLf & _
			"    ���r��̫�C" & vbCrLf & vbCrLf & _
			"  - �b��r�᭱��ܱa�A�����K���� (�q�`�Ω�Ȭw�y��)" & vbCrLf & _
			"    �q�`�b�Ȭw�y���p����B��嵥�n�餤�ϥ� (&X) �Φ����K����A�ñN��m��r�굲���]�b��" & vbCrLf & _
			"    ��ũΥ[�t���e�^�C" & vbCrLf & _
			"    ��ܸӿﶵ��A�N�ˬd�t���K���䪺½Ķ�r�ꤤ���K����O�_�ŦX�D�ҡC�p�G���ŦX�N�Q�۰�" & vbCrLf & _
			"    �ܧ�øm��C" & vbCrLf & vbCrLf & _
			"<�פ��>" & vbCrLf & _
			"  - �n�ˬd���פ��" & vbCrLf & _
			"    ½Ķ�����פ�ũM��夣�ۦP�ɡA�N�Q�ܧ󬰭�夤���פ�šA���O�ŦX�n�۰ʴ������פ��" & vbCrLf & _
			"    �襤���פ�Ű��~�C" & vbCrLf & vbCrLf & _
			"    �����䴩�U�Φr���A���O���O�ҽk���ӬO��T���C�Ҧp�GA*C ���ŦX XAYYCZ�A�u�ŦX AXYYC�C" & vbCrLf & _
			"    �n�ŦX XAYYCZ�A���Ӭ� *A*C* �� *A??C*�C" & vbCrLf & vbCrLf & _
			"    ���`�N�G�ѩ� Sax Basic ���������D�A�G�ӫD�^��r�������� ? �U�Φr�����Q�䴩�C" & vbCrLf & _
			"    �@�@�@�@�Ҧp�G�u�}��??�ɮסv���ŦX�u�}�ҨϥΪ��ɮסv�C" & vbCrLf & vbCrLf & _
			"  - �n�O�d���פ�ŲզX" & vbCrLf & _
			"    �Ҧ��Q�]�t�b�ӲզX�����n�ˬd���פ�űN�Q�O�d�C�]�N�O�o�ǲפ�Ť��Q�{���O�פ�šC" & vbCrLf & vbCrLf & _
			"    �����䴩�U�Φr���A���O���O�ҽk���ӬO��T���C�Ҧp�GA*C ���ŦX XAYYCZ�A�u�ŦX AXYYC�C" & vbCrLf & _
			"    �n�ŦX XAYYCZ�A���Ӭ� *A*C* �� *A??C*�C" & vbCrLf & vbCrLf & _
			"    ���`�N�G�ѩ� Sax Basic ���������D�A�G�ӫD�^��r�������� ? �U�Φr�����Q�䴩�C" & vbCrLf & _
			"    �@�@�@�@�Ҧp�G�u�}��??�ɮסv���ŦX�u�}�ҨϥΪ��ɮסv�C" & vbCrLf & vbCrLf & _
			"  - �n�۰ʴ������פ�Ź�" & vbCrLf & _
			"    �ŦX�פ�Ź襤�e�@�Ӧr�����פ�šA���N�Q�������פ�Ź襤��@�Ӧr�����פ�šC" & vbCrLf & _
			"    �Q�Φ����i�H�۰�½Ķ�έץ��@�ǲפ�šC" & vbCrLf & vbCrLf & _
			"<�[�t��>" & vbCrLf & _
			"  - �n�ˬd���[�t���X�в�" & vbCrLf & _
			"    �[�t���q�`�H \t ���X�в� (�]���ҥ~��)�A�p�G�r�ꤤ�]�t�o�Ǧr���A�N�Q�{���]�t�[�t���A" & vbCrLf & _
			"    ���ݭn�ھڭn�ˬd���[�t���r���i�@�B�P�_�C" & vbCrLf & vbCrLf & _
			"  - �n�ˬd���[�t���r��" & vbCrLf & _
			"    �]�t�[�t���X�вŪ��r�ꤤ�A�p�G�X�вū᭱���r���ŦX����쪺�r���A�N�Q���Ѭ��[�t���A" & vbCrLf & _
			"    �G�ӥH�W�r���զX�Ӧ����[�t�����\�䤤�@�Ӥ��ŦX�C" & vbCrLf & vbCrLf & _
			"    �����䴩�U�Φr���A���O���O�ҽk���ӬO��T���C�Ҧp�GA*C ���ŦX XAYYCZ�A�u�ŦX AXYYC�C" & vbCrLf & _
			"    �n�ŦX XAYYCZ�A���Ӭ� *A*C* ��  *A??C*�C" & vbCrLf & vbCrLf & _
			"    ���`�N�G�ѩ� Sax Basic ���������D�A�G�ӫD�^��r�������� ? �U�Φr�����Q�䴩�C" & vbCrLf & _
			"    �@�@�@�@�Ҧp�G�u�}��??�ɮסv���ŦX�u�}�ҨϥΪ��ɮסv�C" & vbCrLf & vbCrLf & _
			"  - �n�O�d���[�t���r��" & vbCrLf & _
			"    �ŦX�o�Ǧr�����[�t���N�Q�O�d�A�_�h�N�Q�����C�Q�Φ����i�O�d�Y�ǥ[�t����½Ķ�C" & vbCrLf & vbCrLf & _
			"<�r������>" & vbCrLf & _
			"  �r�ꤤ�]�t�C�Ӵ����r���襤���u|�v�e���r���ɡA�N�Q�������u|�v�᪺�r���C" & vbCrLf & vbCrLf & _
			"  - ½Ķ�e�n�������r��" & vbCrLf & _
			"    �w�q�b½Ķ�e�n�Q�������r���H�δ����᪺�r���C" & vbCrLf & vbCrLf & _
			"  - ½Ķ��n�������r��" & vbCrLf & _
			"    �w�q�b½Ķ��n�Q�������r���H�δ����᪺�r���C" & vbCrLf & vbCrLf & _
			"  ���`�N�G�����ɰϤ��j�p�g�C" & vbCrLf & _
			"  �@�@�@�@�p�G�n�h���o�Ǧr���A�i�H�N�u|�v�᪺�r���m�šC" & vbCrLf & vbCrLf & _
			"<�A�λy��>" & vbCrLf & _
			"  �o�̪��A�λy���O��½Ķ�M�檺�ؼлy���A���Ω�ھ�½Ķ�M�檺�ؼлy���۰ʿ�������]�w��" & vbCrLf & _
			"  �۰ʿ���\��C" & vbCrLf & vbCrLf & _
			"  - �s�W" & vbCrLf & _
			"    �n�s�W�A�λy���A����i�λy���M�椤���y���A�M���I�� [�s�W] ���s�C" & vbCrLf & _
			"    �I���ӫ��s��A�i�λy���M�椤����ܻy���N���ʨ�A�λy���M�椤�C" & vbCrLf & vbCrLf & _
			"  - �����s�W" & vbCrLf & _
			"    �I���ӫ��s��A�i�λy���M�椤���Ҧ��y���N�������ʨ�A�λy���M�椤�C" & vbCrLf & vbCrLf & _
			"  - �R��" & vbCrLf & _
			"    �n�R���A�λy���A����A�λy���M�椤���y���A�M���I�� [�R��] ���s�C" & vbCrLf & _
			"    �I���ӫ��s��A�A�λy���M�椤����ܻy���N���ʨ�i�λy���M�椤�C" & vbCrLf & vbCrLf & _
			"  - �����R��" & vbCrLf & _
			"    �I���ӫ��s��A�A�λy���M�椤���Ҧ��y���N�������ʨ�i�λy���M�椤�C" & vbCrLf & vbCrLf & _
			"  - �W�[�i�λy��" & vbCrLf & _
			"    �I���ӫ��s��A�N�u�X�i��J�y���W�٩M�N�X��ܤ���A�T�w��N�s�W��i�λy���M�椤�C" & vbCrLf & vbCrLf & _
			"  - �s��i�λy��" & vbCrLf & _
			"    �n�s��i�λy���A����i�λy���M�椤���y���A�M���I�� [�s��i�λy��] ���s�C" & vbCrLf & _
			"    �I���ӫ��s��A�N�u�X�i�s��y���W�٩M�N�X��ܤ���A�T�w��N�ק�i�λy���M�椤��ܪ��y���C" & vbCrLf & vbCrLf & _
			"  - �R���i�λy��" & vbCrLf & _
			"    �n�R���i�λy���A����i�λy���M�椤�n�R�����y���A�M���I�� [�R���i�λy��] ���s�C" & vbCrLf & vbCrLf & _
			"  - �W�[�A�λy��" & vbCrLf & _
			"    �I���ӫ��s��A�N�u�X�i��J�y���W�٩M�N�X��ܤ���A�T�w��N�s�W��A�λy���M�椤�C" & vbCrLf & vbCrLf & _
			"  - �s��A�λy��" & vbCrLf & _
			"    �n�s��A�λy���A����A�λy���M�椤���y���A�M���I�� [�s��A�λy��] ���s�C" & vbCrLf & _
			"    �I���ӫ��s��A�N�u�X�i�s��y���W�٩M�N�X��ܤ���A�T�w��N�ק�A�λy���M�椤��ܪ��y���C" & vbCrLf & vbCrLf & _
			"  - �R���A�λy��" & vbCrLf & _
			"    �n�R���A�λy���A����A�λy���M�椤�n�R�����y���A�M���I�� [�R���A�λy��] ���s�C" & vbCrLf & vbCrLf & _
			"  ���`�N�G�s�W�B�s��y���ȥΩ� Passolo ���Ӫ����s�W���䴩�y���C" & vbCrLf & _
			"  �@�@�@�@�y���N�X�ЩM Passolo �� ISO 396-1 �N�X�O���@�P�A�]�A�j�p�g�C" & vbCrLf & vbCrLf & _
			"����L�\�ࡸ" & vbCrLf & _
			"============" & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �I���ӫ��s�A�N�u�X�ثe�����������T���C" & vbCrLf & vbCrLf & _
			"- Ū��" & vbCrLf & _
			"  �I���ӫ��s�A�N�ھڿ�ܳ]�w�����P�u�X�U�C���:" & vbCrLf & _
			"  (1) �w�]��" & vbCrLf & _
			"      Ū���w�]�]�w�ȡA����ܦb�]�w���e���C" & vbCrLf & _
			"      ���ȷ��ܪ��]�w���t�ιw�]���]�w�ɡA�~��ܸӿ��C" & vbCrLf & vbCrLf & _
			"  (2) ���" & vbCrLf & _
			"      Ū����ܳ]�w����l�ȡA����ܦb�]�w���e���C" & vbCrLf & _
			"      ���ȷ��ܳ]�w����l�Ȭ��D�ŮɡA�~��ܸӿ��C" & vbCrLf & vbCrLf & _
			"  (3) �ѷӭ�" & vbCrLf & _
			"      Ū����ܪ��ѷӳ]�w�ȡA����ܦb�]�w���e���C" & vbCrLf & _
			"      ���ӿ����ܰ���ܳ]�w�~���Ҧ��]�w�M��C" & vbCrLf & vbCrLf & _
			"- �M��" & vbCrLf & _
			"  �I���ӫ��s�A�N�M�Ų{���]�w�������ȡA�H��K���s��J�]�w�ȡC" & vbCrLf & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �I���ӫ��s�A�N�u�X���չ�ܤ���A�H�K�ˬd�]�w�����T�ʡC" & vbCrLf & vbCrLf & _
			"- �T�w" & vbCrLf & _
			"  �I���ӫ��s�A�N�x�s�]�w�������������ܧ�A�����]�w�����ê�^�D�����C" & vbCrLf & _
			"  �{���N�ϥ��ܧ�᪺�]�w�ȡC" & vbCrLf & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �I���ӫ��s�A���x�s�]�w�������������ܧ�A�����]�w�����ê�^�D�����C" & vbCrLf & _
			"  �{���N�ϥέ�Ӫ��]�w�ȡC" & vbCrLf & vbCrLf & vbCrLf
	TestUse ="���]�w�W�١�" & vbCrLf & _
			"============" & vbCrLf & _
			"- �n���ժ��]�w�W�١C�n����]�w�A�I���]�w�M��C" & vbCrLf & vbCrLf & _
			"��½Ķ�M�桸" & vbCrLf & _
			"============" & vbCrLf & _
			"- �ӲM��N��ܱM�פ����Ҧ�½Ķ�M��C�п���P�z���ۭq�]�w�ŦX��½Ķ�M��i����աC" & vbCrLf & vbCrLf & _
			"���۰ʴ����r����" & vbCrLf & _
			"============" & vbCrLf & _
			"- �b�ˬd�ɦ۰ʴ����r�ꤤ�ŦX�]�w���ҩw�q�������r���C" & vbCrLf & vbCrLf & _
			"��Ū�J��ơ�" & vbCrLf & _
			"============" & vbCrLf & _
			"- ��ܭn��ܪ����~½Ķ�r��ơC��ĳ���n��J�Ӥj���ȡA�H�K�b�r����h�ɵ��ݮɶ��L���C" & vbCrLf & vbCrLf & _
			"���]�t���e��" & vbCrLf & _
			"============" & vbCrLf & _
			"- ���w�u�ˬd�]�t���e���r��C�Q�θӶ��i�H���w��ʪ����աA�åB�[�ִ��ծɶ��C" & vbCrLf & _
			"- �����䴩�ҽk���U�Φr���C�Ҧp�GA*C �i�H�ŦX XAYYCZ�C" & vbCrLf & vbCrLf & _
			"���`�N�G�ѩ� Sax Basic ���������D�A�G�ӫD�^��r�������� ? �U�Φr�����Q�䴩�C" & vbCrLf & _
			"�@�@�@�@�Ҧp�G�u�}��??�ɮסv���ŦX�u�}�ҨϥΪ��ɮסv�C" & vbCrLf & vbCrLf & _
			"���r�ꤺ�e��" & vbCrLf & _
			"============" & vbCrLf & _
			"���ѤF�����B�K����B�פ�šB�[�t�� 4 �ӿﶵ�C" & vbCrLf & vbCrLf & _
			"- �p�G��������A�h��L�涵�N�Q�۰ʨ�������C" & vbCrLf & _
			"- �p�G����涵�A�h�����ﶵ�N�Q�۰ʨ�������C" & vbCrLf & _
			"- �涵�i�H�h��C" & vbCrLf & vbCrLf & _
			"����L�\�ࡸ" & vbCrLf & _
			"============" & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �I���ӫ��s�A�N�u�X�ثe�����������T���C" & vbCrLf & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �I���ӫ��s�A�N���ӿ�ܪ�����i����աC" & vbCrLf & vbCrLf & _
			"- �M��" & vbCrLf & _
			"  �I���ӫ��s�A�N�M�Ų{�������յ��G�C" & vbCrLf & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �I���ӫ��s�A�N�������յ{���ê�^�]�w�����C" & vbCrLf & vbCrLf & vbCrLf
	CopyRight = "�����v�ŧi��" & vbCrLf & _
			"============" & vbCrLf & _
			"- ���n�骺���v�k�}�o�̩M�ק�̩Ҧ��A����H�i�H�K�O�ϥΡB�ק�B�ƻs�B���G���n��C" & vbCrLf & _
			"- �ק�B���G���n�饲���H���������ɮסA�õ����n���l�}�o�̥H�έק�̡C" & vbCrLf & _
			"- ���g�}�o�̩M�ק�̦P�N�A�����´�έӤH�A���o�Ω�ӷ~�n��B�ӷ~�άO�䥦��Q�ʬ��ʡC" & vbCrLf & _
			"- ��ϥΥ��n�骺��l�����A�H�Ψϥθg�L�H�ק諸�D��l�����ҳy�����l���M�l�`�A�}�o�̤�" & vbCrLf & _
			"  �Ӿ����d���C" & vbCrLf & _
			"- �ѩ󬰧K�O�n��A�}�o�̩M�ק�̨S���q�ȴ��ѳn��޳N�䴩�A�]�L�q�ȧ�i�Χ�s�����C" & vbCrLf & _
			"- �w��������~�ô��X��i�N���C�p�����~�Ϋ�ĳ�A�жǰe��: z_shangyi@163.com�C" & vbCrLf & vbCrLf & vbCrLf
	Thank = "���P�@�@�¡�" & vbCrLf & _
			"============" & vbCrLf & _
			"- ���n��b�ק�L�{���o��~�Ʒs�@���|�������աA�b����ܰJ�ߪ��P�¡I" & vbCrLf & _
			"- �P�¥x�W Heaven ���ʹ��X����λy�ק�N���I" & vbCrLf & vbCrLf & vbCrLf
	Contact = "���P���pô��" & vbCrLf & _
			"============" & vbCrLf & _
			"wanfu�Gz_shangyi@163.com" & vbCrLf & vbCrLf & vbCrLf
	Logs = "�P�¤���I�z������O�ڳ̤j���ʤO�I�P���w��ϥΧڭ̻s�@���n��I" & vbCrLf & vbCrLf & _
			"**************************************************" & vbCrLf & _
			"�ݭn��h�B��s�B��n���~�ơA�Ы��X:" & vbCrLf & _
			"�~�Ʒs�@�� -- http://www.hanzify.org" & vbCrLf & _
			"�~�Ʒs�@���׾� -- http://bbs.hanzify.org" & vbCrLf & _
			"**************************************************" & vbCrLf
	Else
	AboutTitle = "����"
	HelpTitle = "����"
	HelpTipTitle = "��ݼ�����ֹ���ͼ���������"
	AboutWindows = " ���� "
	MainWindows = " ������ "
	SetWindows = " ���ô��� "
	TestWindows = " ���Դ��� "
	Lines = "-----------------------"
	Sys = "����汾��" & Version & vbCrLf & _
			"����ϵͳ��Windows XP/2000 ����ϵͳ" & vbCrLf & _
			"���ð汾������֧�ֺ괦��� Passolo 6.0 �����ϰ汾" & vbCrLf & _
			"�������ԣ��������ĺͷ������� (�Զ�ʶ��)" & vbCrLf & _
			"��Ȩ���У�����������" & vbCrLf & _
			"��Ȩ��ʽ��������" & vbCrLf & _
			"�ٷ���ҳ��http://www.hanzify.org" & vbCrLf & _
			"ǰ�����ߣ����������ͳ�Ա gnatix (2007-2008)" & vbCrLf & _
			"�󿪷��ߣ����������ͳ�Ա wanfu (2009-2010)" & vbCrLf & vbCrLf & vbCrLf
	Ement = "�����л�����" & vbCrLf & _
			"============" & vbCrLf & _
			"- ֧�ֺ괦��� Passolo 6.0 �����ϰ汾������" & vbCrLf & _
			"- Windows Script Host (WSH) ���� (VBS)������" & vbCrLf & _
			"- Adodb.Stream ���� (VBS)��֧���Զ���������" & vbCrLf & _
			"- Microsoft.XMLHTTP ����֧���Զ���������" & vbCrLf & vbCrLf & vbCrLf
	Dec = "���������" & vbCrLf & _
			"============" & vbCrLf & _
			"��ݼ�����ֹ���ͼ�����������һ������ Passolo ������ĺ�������������¹��ܣ�" & vbCrLf & _
			"- ��鷭���п�ݼ�����ֹ�����������Ϳո�" & vbCrLf & _
			"- ��鲢������鷭���п�ݼ�����ֹ�����������Ϳո�" & vbCrLf & _
			"- ɾ�������еĿ�ݼ�" & vbCrLf & _
			"- ���ÿ��Զ�����Զ����¹���" & vbCrLf & vbCrLf & _
			"��������������ļ���" & vbCrLf & _
			"- �Զ��꣺PslAutoAccessKey.bas" & vbCrLf & _
			"  �ڷ����ִ�ʱ���Զ���������ķ��롣���øú꣬�����Բ��������ݼ�����ֹ������������" & vbCrLf & _
			"  ϵͳ��������ѡ��������Զ�������Ӻ�ԭ��һ���Ŀ�ݼ�����ֹ��������������������ֹ����" & vbCrLf & _
			"  ������������߷����ٶȣ������ٷ������" & vbCrLf & _
			"  ��ע�⣺���� Passolo �����ƣ��ú��������ͨ�����м�����ѡ������á�" & vbCrLf & vbCrLf & _
			"- ���꣺PSLCheckAccessKeys.bas" & vbCrLf & _
			"  ͨ��������ӵ� Passolo �˵��еĸú꣬�������������������еĿ�ݼ�����ֹ������������" & vbCrLf & _
			"  ��������ֹ�������⣬�����ṩ�Զ������ú����ü�⹦�ܡ�" & vbCrLf & vbCrLf & _
			"- ��������˵���ļ���AccessKey.txt" & vbCrLf & vbCrLf & vbCrLf
	Setup = "�װ������" & vbCrLf & _
			"============" & vbCrLf & _
			"- ���ʹ���� Wanfu �� Passolo �����棬����װ�˸��Ӻ���������滻ԭ�����ļ����ɣ�����" & vbCrLf & _
			"  (1) ����ѹ����ļ����Ƶ� Passolo ϵͳ�ļ����ж���� Macros �ļ�����" & vbCrLf & _
			"  (2) �Զ��꣺�� Passolo �Ĺ��� -> ��Ի��򣬽�������Ϊϵͳ�겢���������ڵ����½ǵ�ϵͳ" & vbCrLf & _
			"  ��  �꼤��˵�������" & vbCrLf & _
			"  (3) ���꣺�� Passolo �Ĺ��� -> �Զ��幤�߲˵�����Ӹ��ļ�" & vbCrLf & _
			"- �����Զ����޷������й����н������ã�������ʹ�ü������Զ������÷�����" & vbCrLf & _
			"- ���������������ֹ����飬������������" & vbCrLf & vbCrLf & vbCrLf
	MainUse = "������ѡ���" & vbCrLf & _
			"============" & vbCrLf & _
			"�����ṩ��Ĭ�ϵ����ã������ÿ��������ڴ�����������Ҳ���԰� [����] ��ť�Զ������á�" & vbCrLf & _
			"����Զ������ú��������������б���ѡ����ʹ�õ����á�" & vbCrLf & _
			"�й��Զ������ã�������öԻ��򣬵��� [����] ��ť�����İ����е�˵����" & vbCrLf & vbCrLf & _
			"- �Զ�������" & vbCrLf & _
			"  ѡ������ý������Զ��ꡣ��ע�Ᵽ�棬��Ȼ��ʹ��ѡ��ǰ�����á�" & vbCrLf & vbCrLf & _
			"- ��������" & vbCrLf & _
			"  ѡ������ý����ڼ��ꡣҪ���´�ʹ��ѡ������ã���Ҫ���档" & vbCrLf & vbCrLf & _
			"- �Զ���ͼ�����ͬ" & vbCrLf & _
			"  ѡ���ѡ��ʱ�����Զ�ʹ�Զ������������������һ�¡�" & vbCrLf & vbCrLf & _
			"- �Զ�ѡ��" & vbCrLf & _
			"  ѡ����ѡ��ʱ�����������е����������б��Զ�ѡ���뷭���б�Ŀ������ƥ������á�" & vbCrLf & _
			"  ��ע�⣺��ѡ����Կ�ݼ�����ֹ���ͼ�����������Ч��" & vbCrLf & _
			"  ��������Ҫ�������뵱ǰ�����б��Ŀ������ƥ�䣬�� [����] ��ť�����ִ����������������" & vbCrLf & _
			"  �������������Ӧ�����ԡ�" & vbCrLf & _
			"  ���������Զ�����򽫰����Զ� - ��ѡ - Ĭ�ϡ�˳��ѡ����Ӧ�����á�" & vbCrLf & vbCrLf & _
			"�����ǡ�" & vbCrLf & _
			"============" & vbCrLf & _
			"�ù���ͨ����¼�ִ������Ϣ�������ݸ���Ϣֻ�Դ����ִ����м�飬�ɴ������ټ���ٶȡ�" & vbCrLf & _
			"������ 4 ��ѡ��ɹ�ѡ��" & vbCrLf & vbCrLf & _
			"- ���԰汾" & vbCrLf & _
			"  �������Ǻ����İ汾��������������¼�Դ����ִ����м�顣" & vbCrLf & vbCrLf & _
			"- ��������" & vbCrLf & _
			"  �������������Ƿ���ͬ��������������¼�Դ����ִ����м�顣" & vbCrLf & vbCrLf & _
			"- ��������" & vbCrLf & _
			"  �������Ǽ�����ںͷ������ڣ�������������¼�Դ����ִ����м�顣" & vbCrLf & vbCrLf & _
			"- ȫ������" & vbCrLf & _
			"  ���������κμ���¼�����������ִ����м�顣" & vbCrLf & vbCrLf & _
			"��ע�⣺����ǹ����ڲ�������ʱ��Ч��������ܳ�������©�Ĳ��Խ����" & vbCrLf & _
			"����������������������ݶ��������������ƵĻ�����ѡ��ȫ�����Ի�ɾ�������ѡ�" & vbCrLf & vbCrLf & _
			"������ѡ���" & vbCrLf & _
			"============" & vbCrLf & _
			"������ 3 ��ѡ��ɹ�ѡ��" & vbCrLf & vbCrLf & _
			"- �����" & vbCrLf & _
			"  ֻ�Է�����м�飬������������ķ��롣" & vbCrLf & vbCrLf & _
			"- ��鲢����" & vbCrLf & _
			"  �Է�����м�飬���Զ���������ķ��롣" & vbCrLf & vbCrLf & _
			"- ɾ����ݼ�" & vbCrLf & _
			"  ɾ�����������еĿ�ݼ���" & vbCrLf & vbCrLf & _
			"���ִ����͡�" & vbCrLf & _
			"============" & vbCrLf & _
			"�ṩ��ȫ�����˵����Ի����ַ��������������汾����������ѡ����ѡ�" & vbCrLf & vbCrLf & _
			"- ���ѡ��ȫ����������������Զ�ȡ��ѡ��" & vbCrLf & _
			"- ���ѡ�����ȫ��ѡ����Զ�ȡ��ѡ��" & vbCrLf & _
			"- ������Զ�ѡ������ѡ���ѡ��ʱ�����������Զ�ȡ��ѡ��" & vbCrLf & vbCrLf & _
			"���ִ����ݡ�" & vbCrLf & _
			"============" & vbCrLf & _
			"�ṩ��ȫ������ݼ�����ֹ���������� 4 ��ѡ�" & vbCrLf & vbCrLf & _
			"- ���ѡ��ȫ����������������Զ�ȡ��ѡ��" & vbCrLf & _
			"- ���ѡ�����ȫ��ѡ����Զ�ȡ��ѡ��" & vbCrLf & _
			"- ������Զ�ѡ��" & vbCrLf & vbCrLf & _
			"������ѡ���" & vbCrLf & _
			"============" & vbCrLf & _
			"- ������ԭʼ����״̬" & vbCrLf & _
			"  ѡ������ʱ�����ڼ�顢��鲢������ɾ����ݼ�ʱ�������ִ���ԭʼ����״̬������" & vbCrLf & _
			"  �������޴�����޸����ִ��ķ���״̬Ϊ����֤״̬���д�����Ѹ����ִ��ķ���״̬" & vbCrLf & _
			"  Ϊ������״̬���Ա���һ�۾Ϳ���֪����Щ�ִ��д�����ѱ����ġ�" & vbCrLf & vbCrLf & _
			"- ��������ɾ�������" & vbCrLf & _
			"  ѡ������ʱ�������ڷ����б���������Ϣ������Ѵ��ڼ������Ϣ������ɾ����" & vbCrLf & _
			"  ��ע�⣺ѡ������ʱ������ǿ��е�ȫ���������ѡ����" & vbCrLf & vbCrLf & _
			"- ����ʱ�Զ���������ѡ��" & vbCrLf & _
			"  ѡ������ʱ�����ڰ� [����] ��ťʱ�Զ���������ѡ���´�����ʱ�����뱣���ѡ��" & vbCrLf & _
			"  ��ע�⣺����Զ������ñ����ģ�ϵͳ���Զ�ѡ����ѡ���ʹѡ�����Զ���������Ч��" & vbCrLf & vbCrLf & _
			"- �滻�ض��ַ�" & vbCrLf & _
  			"  �ڼ�鲢����������ʹ�������ж����Ҫ�Զ��滻���ַ����滻�ִ����ض����ַ���" & vbCrLf & vbCrLf & _
			"���������ܡ�" & vbCrLf & _
			"============" & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �����ð�ť�������ڶԻ��򡣲���ʾ������ܡ����л����������̼���Ȩ����Ϣ��" & vbCrLf& vbCrLf & _
			"- ����" & vbCrLf & _
			"  �����ð�ť����������ǰ���ڵİ�����Ϣ��" & vbCrLf& vbCrLf & _
			"- ��������ѡ��" & vbCrLf & _
			"  �ð�ť�����ڸ������ö������м��ʱʹ�á�" & vbCrLf & _
			"  �����һѡ����ģ���ѡ��Զ���Ϊ����״̬�������Զ���Ϊ������״̬��" & vbCrLf & vbCrLf & _
			"- ȷ��" & vbCrLf & _
			"  �����ð�ť�����ر����Ի��򣬲���ѡ����ѡ������ִ����롣" & vbCrLf& vbCrLf & _
			"- ȡ��" & vbCrLf & _
			"  �����ð�ť�����˳�����" & vbCrLf& vbCrLf & vbCrLf
	SetUse ="�������б��" & vbCrLf & _
			"============" & vbCrLf & _
			"- ѡ������" & vbCrLf & _
			"  Ҫѡ�����ã����������б�" & vbCrLf & vbCrLf & _
			"- �������õ����ȼ�" & vbCrLf & _
			"  �������ȼ����ڻ������õ��������Ե��Զ�ѡ�����ù��ܡ�" & vbCrLf & _
			"  ��ע�⣺�ж�����ð�������ͬ����������ʱ����Ҫ���������ȼ���" & vbCrLf & _
			"  ������������ͬ�������Ե������У�ǰ������ñ�����ѡ��ʹ�á�" & vbCrLf & _
			"  Ҫ�������õ����ȼ��������ұߵ� [...] ��ť��" & vbCrLf & vbCrLf & _
			"- �������" & vbCrLf & _
			"  Ҫ������ã����� [���] ��ť���ڵ����ĶԻ������������ơ�" & vbCrLf & vbCrLf & _
			"- ��������" & vbCrLf & _
			"  Ҫ�����������ƣ���ѡ�������б���Ҫ���������ã�Ȼ�󵥻� [����] ��ť��" & vbCrLf & vbCrLf & _
			"- ɾ������" & vbCrLf & _
			"  Ҫɾ�����ã���ѡ�������б���Ҫɾ�������ã�Ȼ�󵥻� [ɾ��] ��ť��" & vbCrLf & vbCrLf & _
			"������ú󣬽����б�����ʾ�µ����ã��������ݽ���ʾ��ֵ��" & vbCrLf & _
			"�������ú󣬽����б�����ʾ���������ã����������е�����ֵ���䡣" & vbCrLf & _
			"ɾ�����ú󣬽����б�����ʾĬ�����ã��������ݽ���ʾĬ������ֵ��" & vbCrLf & vbCrLf & _
			"������͡�" & vbCrLf & _
			"============" & vbCrLf & _
			"- �ļ�" & vbCrLf & _
			"  ���ý����ļ���ʽ�����ں������ļ����µ� Data �ļ����С�" & vbCrLf & vbCrLf & _
			"- ע���" & vbCrLf & _
			"  ���ý�������ע����е� HKCU\Software\VB and VBA Program Settings\AccessKey ���¡�" & vbCrLf & vbCrLf & _
			"- ��������" & vbCrLf & _
			"  ��������������ļ��е������á����������ʱ�����Զ����������������б������е����ý���" & vbCrLf & _
			"  ���ģ�û�е����ý�����ӡ�" & vbCrLf & vbCrLf & _
			"- ��������" & vbCrLf & _
			"  �������������õ��ı��ļ����Ա���Խ�����ת�����á�" & vbCrLf & vbCrLf & _
			"��ע�⣺�л���������ʱ�����Զ�ɾ��ԭ��λ���е��������ݡ�" & vbCrLf & vbCrLf & _
			"���������ݡ�" & vbCrLf & _
			"============" & vbCrLf & _
			"<��ݼ�>" & vbCrLf & _
			"  - Ҫ�ų��ĺ� & ���ŵķǿ�ݼ����" & vbCrLf & _
			"    ��ݼ��� & Ϊ��־������Щ�ִ���Ȼ�����÷��ŵ����ǿ�ݼ�����Ҫ�ų������ڴ�������Щ" & vbCrLf & _
			"    Ҫ�ų����� & ���ŵķǿ�ݼ���ϡ�" & vbCrLf & vbCrLf & _
			"  - �ִ�����ñ�־��" & vbCrLf & _
			"    ��Щ�ַ����ڲ�ֺ��ж����ݼ�����ֹ������������ִ����Ա����ִ������еĿ�ݼ���" & vbCrLf & _
			"    ��ֹ���������������ֻ�ܴ����ִ���󲿷ֵĿ�ݼ�����ֹ�����������" & vbCrLf & vbCrLf & _
			"  - Ҫ���Ŀ�ݼ�ǰ������" & vbCrLf & _
			"    Ĭ�ϵĿ�ݼ�ǰ������Ϊ ()���ڴ�ָ���Ŀ�ݼ�ǰ�����ţ��������滻ΪĬ�ϵ����š�" & vbCrLf & vbCrLf & _
			"  - Ҫ�����ķǿ�ݼ�ǰ��ɶ��ַ�" & vbCrLf & _
			"    ���ԭ�ĺͷ����ж�������Щ�ַ����������������򽫱���Ϊ�ǿ�ݼ���������ӿ�ݼ�����" & vbCrLf & _
			"    λ���ִ����" & vbCrLf & vbCrLf & _
			"  - ���ı�������ʾ�����ŵĿ�ݼ� (ͨ��������������)" & vbCrLf & _
			"    ͨ�����������������ġ����ĵ������ʹ�� (&X) ��ʽ�Ŀ�ݼ��������������ִ���β������" & vbCrLf & _
			"    ֹ���������ǰ����" & vbCrLf & _
			"    ѡ����ѡ��󣬽���麬�п�ݼ��ķ����ִ��еĿ�ݼ��Ƿ���Ϲ�������������Ͻ����Զ�" & vbCrLf & _
			"    ���Ĳ��ú�" & vbCrLf & vbCrLf & _
			"<��ֹ��>" & vbCrLf & _
			"  - Ҫ������ֹ��" & vbCrLf & _
			"    �����е���ֹ����ԭ�Ĳ�һ��ʱ����������Ϊԭ���е���ֹ�������Ƿ���Ҫ�Զ��滻����ֹ��" & vbCrLf & _
			"    ���е���ֹ�����⡣" & vbCrLf & vbCrLf & _
			"    ���ֶ�֧��ͨ��������ǲ���ģ���Ķ��Ǿ�ȷ�ġ����磺A*C ��ƥ�� XAYYCZ��ֻƥ�� AXYYC��" & vbCrLf & _
			"    Ҫƥ�� XAYYCZ��Ӧ��Ϊ *A*C* �� *A??C*��" & vbCrLf & vbCrLf & _
			"    ��ע�⣺���� Sax Basic ��������⣬������Ӣ����ĸ֮��� ? ͨ�������֧�֡�" & vbCrLf & _
			"    �����������磺����??�ļ�����ƥ�䡰���û��ļ�����" & vbCrLf & vbCrLf & _
			"  - Ҫ��������ֹ�����" & vbCrLf & _
			"    ���б������ڸ�����е�Ҫ������ֹ������������Ҳ������Щ��ֹ��������Ϊ����ֹ����" & vbCrLf & vbCrLf & _
			"    ���ֶ�֧��ͨ��������ǲ���ģ���Ķ��Ǿ�ȷ�ġ����磺A*C ��ƥ�� XAYYCZ��ֻƥ�� AXYYC��" & vbCrLf & _
			"    Ҫƥ�� XAYYCZ��Ӧ��Ϊ *A*C* �� *A??C*��" & vbCrLf & vbCrLf & _
			"    ��ע�⣺���� Sax Basic ��������⣬������Ӣ����ĸ֮��� ? ͨ�������֧�֡�" & vbCrLf & _
			"    �����������磺����??�ļ�����ƥ�䡰���û��ļ�����" & vbCrLf & vbCrLf & _
			"  - Ҫ�Զ��滻����ֹ����" & vbCrLf & _
			"    ������ֹ������ǰһ���ַ�����ֹ�����������滻����ֹ�����к�һ���ַ�����ֹ����" & vbCrLf & _
			"    ���ô�������Զ����������һЩ��ֹ����" & vbCrLf & vbCrLf & _
			"<������>" & vbCrLf & _
			"  - Ҫ���ļ�������־��" & vbCrLf & _
			"    ������ͨ���� \t Ϊ��־�� (Ҳ�������)������ִ��а�����Щ�ַ���������Ϊ������������" & vbCrLf & _
			"    ����Ҫ����Ҫ���ļ������ַ���һ���жϡ�" & vbCrLf & vbCrLf & _
			"  - Ҫ���ļ������ַ�" & vbCrLf & _
			"    ������������־�����ִ��У������־��������ַ����ϸ��ֶε��ַ�������ʶ��Ϊ��������" & vbCrLf & _
			"    ���������ַ���϶��ɵļ�������������һ����ƥ�䡣" & vbCrLf & vbCrLf & _
			"    ���ֶ�֧��ͨ��������ǲ���ģ���Ķ��Ǿ�ȷ�ġ����磺A*C ��ƥ�� XAYYCZ��ֻƥ�� AXYYC��" & vbCrLf & _
			"    Ҫƥ�� XAYYCZ��Ӧ��Ϊ *A*C* ��  *A??C*��" & vbCrLf & vbCrLf & _
			"    ��ע�⣺���� Sax Basic ��������⣬������Ӣ����ĸ֮��� ? ͨ�������֧�֡�" & vbCrLf & _
			"    �����������磺����??�ļ�����ƥ�䡰���û��ļ�����" & vbCrLf & vbCrLf & _
			"  - Ҫ�����ļ������ַ�" & vbCrLf & _
			"    ������Щ�ַ��ļ������������������򽫱��滻�����ô���ɱ���ĳЩ�������ķ��롣" & vbCrLf & vbCrLf & _
			"<�ַ��滻>" & vbCrLf & _
			"  �ִ��а���ÿ���滻�ַ����еġ�|��ǰ���ַ�ʱ�������滻�ɡ�|������ַ���" & vbCrLf & vbCrLf & _
			"  - ����ǰҪ�滻���ַ�" & vbCrLf & _
			"    �����ڷ���ǰҪ���滻���ַ��Լ��滻����ַ���" & vbCrLf & vbCrLf & _
			"  - �����Ҫ�滻���ַ�" & vbCrLf & _
			"    �����ڷ����Ҫ���滻���ַ��Լ��滻����ַ���" & vbCrLf & vbCrLf & _
			"  ��ע�⣺�滻ʱ���ִ�Сд��" & vbCrLf & _
			"  �����������Ҫȥ����Щ�ַ������Խ���|������ַ��ÿա�" & vbCrLf & vbCrLf & _
			"<��������>" & vbCrLf & _
			"  ���������������ָ�����б��Ŀ�����ԣ������ڸ��ݷ����б��Ŀ�������Զ�ѡ����Ӧ���õ�" & vbCrLf & _
			"  �Զ�ѡ���ܡ�" & vbCrLf & vbCrLf & _
			"  - ���" & vbCrLf & _
			"    Ҫ����������ԣ�ѡ����������б��е����ԣ�Ȼ�󵥻� [���] ��ť��" & vbCrLf & _
			"    �����ð�ť�󣬿��������б��е�ѡ�����Խ��ƶ������������б��С�" & vbCrLf & vbCrLf & _
			"  - ȫ�����" & vbCrLf & _
			"    �����ð�ť�󣬿��������б��е��������Խ�ȫ���ƶ������������б��С�" & vbCrLf & vbCrLf & _
			"  - ɾ��" & vbCrLf & _
			"    Ҫɾ���������ԣ�ѡ�����������б��е����ԣ�Ȼ�󵥻� [ɾ��] ��ť��" & vbCrLf & _
			"    �����ð�ť�����������б��е�ѡ�����Խ��ƶ������������б��С�" & vbCrLf & vbCrLf & _
			"  - ȫ��ɾ��" & vbCrLf & _
			"    �����ð�ť�����������б��е��������Խ�ȫ���ƶ������������б��С�" & vbCrLf & vbCrLf & _
			"  - ���ӿ�������" & vbCrLf & _
			"    �����ð�ť�󣬽������������������ƺʹ���Ի���ȷ������ӵ����������б��С�" & vbCrLf & vbCrLf & _
			"  - �༭��������" & vbCrLf & _
			"    Ҫ�༭�������ԣ�ѡ����������б��е����ԣ�Ȼ�󵥻� [�༭��������] ��ť��" & vbCrLf & _
			"    �����ð�ť�󣬽������ɱ༭�������ƺʹ���Ի���ȷ�����޸Ŀ��������б���ѡ�������ԡ�" & vbCrLf & vbCrLf & _
			"  - ɾ����������" & vbCrLf & _
			"    Ҫɾ���������ԣ�ѡ����������б���Ҫɾ�������ԣ�Ȼ�󵥻� [ɾ����������] ��ť��" & vbCrLf & vbCrLf & _
			"  - ������������" & vbCrLf & _
			"    �����ð�ť�󣬽������������������ƺʹ���Ի���ȷ������ӵ����������б��С�" & vbCrLf & vbCrLf & _
			"  - �༭��������" & vbCrLf & _
			"    Ҫ�༭�������ԣ�ѡ�����������б��е����ԣ�Ȼ�󵥻� [�༭��������] ��ť��" & vbCrLf & _
			"    �����ð�ť�󣬽������ɱ༭�������ƺʹ���Ի���ȷ�����޸����������б���ѡ�������ԡ�" & vbCrLf & vbCrLf & _
			"  - ɾ����������" & vbCrLf & _
			"    Ҫɾ���������ԣ�ѡ�����������б���Ҫɾ�������ԣ�Ȼ�󵥻� [ɾ����������] ��ť��" & vbCrLf & vbCrLf & _
			"  ��ע�⣺��ӡ��༭���Խ����� Passolo δ���汾������֧�����ԡ�" & vbCrLf & _
			"  �����������Դ������ Passolo �� ISO 396-1 ���뱣��һ�£�������Сд��" & vbCrLf & vbCrLf & _
			"���������ܡ�" & vbCrLf & _
			"============" & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �����ð�ť����������ǰ���ڵİ�����Ϣ��" & vbCrLf & vbCrLf & _
			"- ��ȡ" & vbCrLf & _
			"  �����ð�ť��������ѡ�����õĲ�ͬ�������в˵�:" & vbCrLf & _
			"  (1) Ĭ��ֵ" & vbCrLf & _
			"      ��ȡĬ������ֵ������ʾ�����������С�" & vbCrLf & _
			"      �����ѡ��������ΪϵͳĬ�ϵ�����ʱ������ʾ�ò˵���" & vbCrLf & vbCrLf & _
			"  (2) ԭֵ" & vbCrLf & _
			"      ��ȡѡ�����õ�ԭʼֵ������ʾ�����������С�" & vbCrLf & _
			"      �����ѡ�����õ�ԭʼֵΪ�ǿ�ʱ������ʾ�ò˵���" & vbCrLf & vbCrLf & _
			"  (3) ����ֵ" & vbCrLf & _
			"      ��ȡѡ���Ĳ�������ֵ������ʾ�����������С�" & vbCrLf & _
			"      ��ò˵���ʾ��ѡ������������������б�" & vbCrLf & vbCrLf & _
			"- ���" & vbCrLf & _
			"  �����ð�ť��������������õ�ȫ��ֵ���Է���������������ֵ��" & vbCrLf & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �����ð�ť�����������ԶԻ����Ա������õ���ȷ�ԡ�" & vbCrLf & vbCrLf & _
			"- ȷ��" & vbCrLf & _
			"  �����ð�ť�����������ô����е��κθ��ģ��˳����ô��ڲ����������ڡ�" & vbCrLf & _
			"  ����ʹ�ø��ĺ������ֵ��" & vbCrLf & vbCrLf & _
			"- ȡ��" & vbCrLf & _
			"  �����ð�ť�����������ô����е��κθ��ģ��˳����ô��ڲ����������ڡ�" & vbCrLf & _
			"  ����ʹ��ԭ��������ֵ��" & vbCrLf & vbCrLf & vbCrLf
	TestUse ="���������ơ�" & vbCrLf & _
			"============" & vbCrLf & _
			"- Ҫ���Ե��������ơ�Ҫѡ�����ã����������б�" & vbCrLf & vbCrLf & _
			"����б��" & vbCrLf & _
			"============" & vbCrLf & _
			"- ���б���ʾ�����е����з����б���ѡ���������Զ�������ƥ��ķ����б���в��ԡ�" & vbCrLf & vbCrLf & _
			"���Զ��滻�ַ���" & vbCrLf & _
			"============" & vbCrLf & _
			"- �ڼ��ʱ�Զ��滻�ִ��з�����������������滻�ַ���" & vbCrLf & vbCrLf & _
			"�����������" & vbCrLf & _
			"============" & vbCrLf & _
			"- ��ʾҪ��ʾ�Ĵ������ִ��������鲻Ҫ����̫���ֵ���������ִ��϶�ʱ�ȴ�ʱ�������" & vbCrLf & vbCrLf & _
			"��������ݡ�" & vbCrLf & _
			"============" & vbCrLf & _
			"- ָ��ֻ���������ݵ��ִ������ø������������ԵĲ��ԣ����Ҽӿ����ʱ�䡣" & vbCrLf & _
			"- ���ֶ�֧��ģ����ͨ��������磺A*C ����ƥ�� XAYYCZ��" & vbCrLf & vbCrLf & _
			"��ע�⣺���� Sax Basic ��������⣬������Ӣ����ĸ֮��� ? ͨ�������֧�֡�" & vbCrLf & _
			"�����������磺����??�ļ�����ƥ�䡰���û��ļ�����" & vbCrLf & vbCrLf & _
			"���ִ����ݡ�" & vbCrLf & _
			"============" & vbCrLf & _
			"�ṩ��ȫ������ݼ�����ֹ���������� 4 ��ѡ�" & vbCrLf & vbCrLf & _
			"- ���ѡ��ȫ����������������Զ�ȡ��ѡ��" & vbCrLf & _
			"- ���ѡ�����ȫ��ѡ����Զ�ȡ��ѡ��" & vbCrLf & _
			"- ������Զ�ѡ��" & vbCrLf & vbCrLf & _
			"���������ܡ�" & vbCrLf & _
			"============" & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �����ð�ť����������ǰ���ڵİ�����Ϣ��" & vbCrLf & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �����ð�ť��������ѡ�����������в��ԡ�" & vbCrLf & vbCrLf & _
			"- ���" & vbCrLf & _
			"  �����ð�ť����������еĲ��Խ����" & vbCrLf & vbCrLf & _
			"- ȡ��" & vbCrLf & _
			"  �����ð�ť�����˳����Գ��򲢷������ô��ڡ�" & vbCrLf & vbCrLf & vbCrLf
	CopyRight = "���Ȩ������" & vbCrLf & _
			"============" & vbCrLf & _
			"- ������İ�Ȩ�鿪���ߺ��޸������У��κ��˿������ʹ�á��޸ġ����ơ�ɢ���������" & vbCrLf & _
			"- �޸ġ�ɢ������������渽��˵���ļ�����ע�����ԭʼ�������Լ��޸��ߡ�" & vbCrLf & _
			"- δ�������ߺ��޸���ͬ�⣬�κ���֯����ˣ�����������ҵ�������ҵ��������Ӫ���Ի��" & vbCrLf & _
			"- ��ʹ�ñ������ԭʼ�汾���Լ�ʹ�þ������޸ĵķ�ԭʼ�汾����ɵ���ʧ���𺦣������߲�" & vbCrLf & _
			"  �е��κ����Ρ�" & vbCrLf & _
			"- ����Ϊ�������������ߺ��޸���û�������ṩ�������֧�֣�Ҳ������Ľ�����°汾��" & vbCrLf & _
			"- ��ӭָ����������Ľ���������д�����飬�뷢�͵�: z_shangyi@163.com��" & vbCrLf & vbCrLf & vbCrLf
	Thank = "���¡���л��" & vbCrLf & _
			"============" & vbCrLf & _
			"- ��������޸Ĺ����еõ����������ͻ�Ա�Ĳ��ԣ��ڴ˱�ʾ���ĵĸ�л��" & vbCrLf & _
			"- ��л̨�� Heaven ����������������޸������" & vbCrLf & vbCrLf & vbCrLf
	Contact = "��������ϵ��" & vbCrLf & _
			"============" & vbCrLf & _
			"wanfu��z_shangyi@163.com" & vbCrLf & vbCrLf & vbCrLf
	Logs = "��л֧�֣�����֧���������Ķ�����ͬʱ��ӭʹ�����������������" & vbCrLf & vbCrLf & _
			"**************************************************" & vbCrLf & _
			"��Ҫ���ࡢ���¡����õĺ����������:" & vbCrLf & _
			"���������� -- http://www.hanzify.org" & vbCrLf & _
			"������������̳ -- http://bbs.hanzify.org" & vbCrLf & _
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


'�������
Sub EngineHelp(HelpTip As String)
	If OSLanguage = "0404" Then
	AboutTitle = "����"
	HelpTitle = "����"
	HelpTipTitle = "Passolo �s�u½Ķ����"
	AboutWindows = " ���� "
	MainWindows = " �D���� "
	SetWindows = " �]�w���� "
	TestWindows = " ���յ��� "
	Lines = "-----------------------"
	Sys = "�n�骩���G" & Version & vbCrLf & _
			"�A�Ψt�ΡGWindows XP/2000 �H�W�t��" & vbCrLf & _
			"�A�Ϊ����G�Ҧ��䴩�����B�z�� Passolo 6.0 �ΥH�W����" & vbCrLf & _
			"�����y���G²�餤��M���餤�� (�۰ʿ���)" & vbCrLf & _
			"���v�Ҧ��G�~�Ʒs�@��" & vbCrLf & _
			"���v�Φ��G�K�O�n��" & vbCrLf & _
			"�x�譺���Ghttp://www.hanzify.org" & vbCrLf & _
			"�} �o �̡G�~�Ʒs�@������ wanfu (2010)" & vbCrLf & vbCrLf & vbCrLf
	Ement = "���������ҡ�" & vbCrLf & _
			"============" & vbCrLf & _
			"- �䴩�����B�z�� Passolo 6.0 �ΥH�W�����A����" & vbCrLf & _
			"- Windows Script Host (WSH) ���� (VBS)�A����" & vbCrLf & _
			"- Adodb.Stream ����A�䴩 Utf-8�BUnicode ����" & vbCrLf & _
			"- Microsoft.XMLHTTP ����A����" & vbCrLf & _
			"- Microsoft.XMLDOM ����A��R responseXML ��^�榡�һ�" & vbCrLf & vbCrLf & vbCrLf
	Dec = "���n��²����" & vbCrLf & _
			"============" & vbCrLf & _
			"Passolo �s�u½Ķ�����O�@�ӥΩ� Passolo ½Ķ�M��r�ꪺ�s�u½Ķ�����{���C���㦳�H�U�\��G" & vbCrLf & _
			"- �Q�γs�u½Ķ�����۰�½Ķ Passolo ½Ķ�M�椤���r��" & vbCrLf & _
			"- ��X�F�@�ǵۦW���s�u½Ķ�����A�åi�ۭq��L�s�u½Ķ����" & vbCrLf & _
			"- �i����r�������B���L�r��H�ι�½Ķ�e�᪺�r��i��B�z" & vbCrLf & _
			"- ��X�K����B�פ�šB�[�t���ˬd�����B�i�b½Ķ���ˬd�êȥ�½Ķ�������~" & vbCrLf & _
			"- ���m�i�ۭq���۰ʧ�s�\��" & vbCrLf & vbCrLf & _
			"���{���]�t�U�C�ɮסG" & vbCrLf & _
			"- PSLWebTrans.bas (Passolo �s�u½Ķ�����ɮ� - ��ܤ���覡����)" & vbCrLf & _
			"- PSLWebTrans_Silent.bas (Passolo �s�u½Ķ�����ɮ� - �R�q�覡����)" & vbCrLf & _
			"- PSLWebTrans.txt (²�餤�廡���ɮ�)" & vbCrLf & vbCrLf & vbCrLf
	Setup = "���w�ˤ�k��" & vbCrLf & _
			"============" & vbCrLf & _
			"- �N�����᪺�ɮ׽ƻs�� Passolo �t�θ�Ƨ����w�q�� Macros ��Ƨ���" & vbCrLf & _
			"- �b Passolo ���u�� -> �ۭq�u���椤�s�W�ӥ����ɮרéw�q�ӿ��W�١A" & vbCrLf & _
			"  ����N�i�H�I���ӿ�檽���I�s" & vbCrLf & _
			"- �R�q�覡���檺�s�u½Ķ�����ȧ@�Ω��ܦr��A���L�r��M�r��B�z����ܤ���覡����" & vbCrLf & _
 			"  ���s�u½Ķ�����]�w����C�n�w�q�ӥ���������ѼơA�ݨϥι�ܤ���覡�]�w" & vbCrLf & vbCrLf & vbCrLf
	MainUse = "��½Ķ������" & vbCrLf & _
			"============" & vbCrLf & _
			"�{���w�g���Ѥ@�ǵۦW���s�u½Ķ�����C�z�]�i�H���U���� [�]�w] ���s�ۭq�]�w�C" & vbCrLf & _
			"�s�W�ۭq�]�w��A�z�i�H�b½Ķ�����M�椤����Q�ϥΪ�½Ķ�����C" & vbCrLf & _
			"�����ۭq½Ķ�����A�ж}�ҳ]�w��ܤ���A�I�� [����] ���s�A�Ѿ\�������������C" & vbCrLf & vbCrLf & _
			"��½Ķ��塸" & vbCrLf & _
			"============" & vbCrLf & _
			"�{���۰ʦC�X�ثe½Ķ�M������ɮפ����ثe½Ķ�M�檺�ؼлy���~���Ҧ��ؼлy���C" & vbCrLf & _
			"�q�L������P���y���A�i�H����ثe½Ķ�M�椤���ӷ��r��A�Ϊ̦P�@�ɮפ�����L½Ķ�M�椤" & vbCrLf & _
			"��½Ķ�r��C" & vbCrLf & vbCrLf & _
			"- ����P�ثe½Ķ�M��ؼлy���ۦP���y���ɡA�N�ϥΥثe½Ķ�M�椤���ӷ��r��@��½Ķ���C" & vbCrLf & _
			"- ����P�ثe½Ķ�M��ؼлy�����P���y���ɡA�N�ϥο�ܪ���L½Ķ�M�椤��½Ķ�r��@��½Ķ" & vbCrLf & _
			"  ���C�H��K�N�w����½Ķ�ഫ����L�y���C��p�N²�餤��½Ķ�����餤��C" & vbCrLf & vbCrLf & _
			"��½Ķ�r�꡸" & vbCrLf & _
			"============" & vbCrLf & _
			"���ѤF�����B���B��ܤ���B�r���B�[�t���B�����B��L�B�ȿ�ܵ��ﶵ�C" & vbCrLf & vbCrLf & _
			"- �p�G��������A�h��L�涵�N�Q�۰ʨ�������C" & vbCrLf & _
			"- �p�G����涵�A�h�����ﶵ�N�Q�۰ʨ�������C" & vbCrLf & _
			"- �涵�i�H�h��C�䤤����ȿ�ܮɡA��L���Q�۰ʨ�������C" & vbCrLf & vbCrLf & _
			"�����L�r�꡸" & vbCrLf & _
			"============" & vbCrLf & _
			"���ѤF�ѽƼf�B�w���ҡB��½Ķ�B�����Ʀr�M�Ÿ��B�����j�g�^��B�����p�g�^�嵥�ﶵ�C" & vbCrLf & vbCrLf & _
			"- �Ҧ����ا��i�H�h��C" & vbCrLf & _
			"- ��������j�g�^��B�����p�g�^��ﶵ�ɱN�۰ʿ�ܥ����Ʀr�M�Ÿ��ﶵ�A�åB�����Ʀr" & vbCrLf & _
			"  �M�Ÿ��ﶵ�N�۰��ܬ����i�Ϊ��A�C" & vbCrLf & vbCrLf & _
			"���r��B�z��" & vbCrLf & _
			"============" & vbCrLf & _
			"- �]�w" & vbCrLf & _
			"  �{�����ѤF�M�K����B�פ�šB�[�t���ˬd�����ۦP���w�]�]�w�A�ӳ]�w�i�H�A�Ω�j�h�Ʊ��p�C" & vbCrLf & _
			"  �z�]�i�H���U���� [�]�w] ���s�ۭq�]�w�C" & vbCrLf & _
			"  �s�W�ۭq�]�w��A�z�i�H�b�]�w�M�椤����Q�ϥΪ��]�w�C" & vbCrLf & _
			"  �����ۭq�]�w�A�ж}�ҳ]�w��ܤ���A�I�� [����] ���s�A�Ѿ\�������������C" & vbCrLf & vbCrLf & _
			"- �۰ʿ���]�w" & vbCrLf & _
			"  ��ܸӿﶵ�ɱN�ھڳ]�w�����i�λy���M��۰ʿ���P½Ķ�M��ؼлy���ŦX���]�w�C" & vbCrLf & _
			"  ���`�N�G�ӿﶵ�ȹ�s�u½Ķ�������ġC" & vbCrLf & _
			"  �@�@�@�@�n�N�]�w�P�ثe½Ķ�M�檺�ؼлy���ŦX�A�� [�]�w] ���s�A�b�r��B�z���A�λy����" & vbCrLf & _
			"  �@�@�@�@�s�W�������y���C" & vbCrLf & _
			"  �@�@�@�@�b�R�q�覡���p�U�A�{���N���u�۰� - �ۿ� - �w�]�v���ǿ���������]�w�C" & vbCrLf & vbCrLf & _
			"- �h���K����" & vbCrLf & _
			"  �b½Ķ�e�N�r�ꤤ���K����R���A�H�K½Ķ�����i�H���T½Ķ�C" & vbCrLf & vbCrLf & _
			"- �h���[�t��" & vbCrLf & _
			"  �b½Ķ�e�N�r�ꤤ���[�t���R���A�H�K½Ķ�����i�H���T½Ķ�C" & vbCrLf & vbCrLf & _
			"- �����S�w�r���æb½Ķ���٭�" & vbCrLf & _
			"  �b½Ķ�e�ϥο�ܪ��]�w�A�����S�w�r���æb½Ķ���٭�C" & vbCrLf & vbCrLf & _
			"  �n�s�W�o�ǯS�w�r���A�Цb�]�w��ܤ�����r��������½Ķ�e�n�������r�����w�q�C" & vbCrLf & _
			"  ���`�N�G�ѩ󥨶����������D�A���Ǧr���O���i�Q��J�Ϊ̵L�k�Q���ѡC" & vbCrLf & _
			"  �@�@�@�@�p�G½Ķ����½Ķ�F�o�Ǵ����᪺�r���A�Q�������r���L�k�Q�٭�C" & vbCrLf & vbCrLf & _
			"- ����½Ķ" & vbCrLf & _
			"  �b½Ķ�e�N�h�檺�r����ά����r��i��½Ķ�A�H�K����½Ķ����½Ķ�����T�סC" & vbCrLf & vbCrLf & _
			"- �ȥ��K����B�פ�šB�[�t��" & vbCrLf & _
			"  �b½Ķ��ϥο�ܪ��]�w�A�ˬd�êȥ�½Ķ�������~�C" & vbCrLf & _
			"  �ȥ��覡�M�K����B�פ�šB�[�t���ˬd���������ۦP�C" & vbCrLf & vbCrLf & _
			"- �����S�w�r��" & vbCrLf & _
			"  �b½Ķ��ϥο�ܪ��]�w�A����½Ķ���S�w���r���C" & vbCrLf & vbCrLf & _
			"  �n�s�W�o�ǯS�w�r���A�Цb�]�w��ܤ�����r��������½Ķ��n�������r�����w�q�C" & vbCrLf & _
			"  ���`�N�G�ѩ󥨶����������D�A���Ǧr���O���i�Q��J�Ϊ̵L�k�Q���ѡC" & vbCrLf & _
			"  �@�@�@�@�p�G½Ķ����½Ķ�F�o�Ǵ����e���r���A�N�L�k�Q�����C" & vbCrLf & vbCrLf & _
			"����L�ﶵ��" & vbCrLf & _
			"============" & vbCrLf & _
			"- �~��ɦ۰��x�s�Ҧ����" & vbCrLf & _
			"  ��ܸӶ��ɡA�N�b�� [�~��] ���s�ɦ۰��x�s�Ҧ�����A�U������ɱNŪ�J�x�s������C" & vbCrLf & vbCrLf & _
			"- ��ܿ�X�T��" & vbCrLf & _
			"  ��ܸӶ��ɡA�N�b Passolo �T����X��������X½Ķ�B�ȥ��M�����T���C" & vbCrLf & _
			"  �b����ܸӶ����p�U�A�N��۴����{��������t�סC" & vbCrLf & vbCrLf & _
			"- �s�W½Ķ����" & vbCrLf & _
			"  ��ܸӶ��ɡA�N�b½Ķ�M�檺�r����Ѥ��W�[���½Ķ�����۰�½Ķ�o�˪����ѡC" & vbCrLf & _
			"  �Q�θӵ��ѥi�H�Ϥ�½Ķ�r�ꪺ½Ķ�X�B�C" & vbCrLf& vbCrLf & _
			"����L�\�ࡸ" & vbCrLf & _
			"============" & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �I���ӫ��s�A�N�����ܤ���C����ܵ{�����СB�������ҡB�}�o�ӤΪ��v���T���C" & vbCrLf & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �I���ӫ��s�A�N�u�X�ثe�����������T���C" & vbCrLf & vbCrLf & _
			"- �]�w" & vbCrLf & _
			"  �I���ӫ��s�A�N�u�X�]�w��ܤ���C�i�H�b�]�w��ܤ�����]�w�U�ذѼơC" & vbCrLf & vbCrLf & _
			"- �T�w" & vbCrLf & _
			"  �I���ӫ��s�A�N�����D��ܤ���A�ë���ܪ��ﶵ�i��r��½Ķ�C" & vbCrLf & _
			"  ���`�N�G½Ķ�᪺�r��½Ķ���A�N�ܧ󬰫ݽƼf���A�A�H�ܰϧO�C" & vbCrLf& vbCrLf & _
			"- ����" & vbCrLf & _
			"  �I���ӫ��s�A�N�����{���C" & vbCrLf& vbCrLf & vbCrLf
	SetUse ="���]�w�M�桸" & vbCrLf & _
			"============" & vbCrLf & _
			"- ����]�w" & vbCrLf & _
			"  �n����]�w���A�I���]�w�M��C" & vbCrLf & vbCrLf & _
			"- �s�W�]�w" & vbCrLf & _
			"  �n�s�W�]�w���A�I�� [�s�W] ���s�A�b�u�X����ܤ������J�W�١C" & vbCrLf & vbCrLf & _
			"- �ܧ�]�w" & vbCrLf & _
			"  �n�ܧ�]�w�W�١A����]�w�M�椤�n��W���]�w�A�M���I�� [�ܧ�] ���s�C" & vbCrLf & vbCrLf & _
			"- �R���]�w" & vbCrLf & _
			"  �n�R���]�w���A����]�w�M�椤�n�R�����]�w�A�M���I�� [�R��] ���s�C" & vbCrLf & vbCrLf & _
			"�s�W�]�w��A�N�b�M�椤��ܷs���]�w�A�]�w���e�N��ܪŭȡC" & vbCrLf & _
			"�ܧ�]�w��A�N�b�M�椤��ܧ�W�᪺�]�w�A�]�w���e�����]�w�Ȥ��ܡC" & vbCrLf & _
			"�R���]�w��A�N�b�M�椤��ܤW�@�ӳ]�w�A�]�w���e�N��ܤW�@�ӳ]�w�ȡC" & vbCrLf & vbCrLf & _
			"���x�s������" & vbCrLf & _
			"============" & vbCrLf & _
			"- �ɮ�" & vbCrLf & _
			"  �]�w�N�H�ɮקΦ��x�s�b�����Ҧb��Ƨ��U�� Data ��Ƨ����C" & vbCrLf & vbCrLf & _
			"- ���U��" & vbCrLf & _
			"  �]�w�N�Q�x�s���U���� HKCU\Software\VB and VBA Program Settings\WebTranslate ���U�C" & vbCrLf & vbCrLf & _
			"- �פJ�]�w" & vbCrLf & _
			"  ���\�q��L�]�w�ɮפ��פJ�]�w�C�פJ�³]�w�ɱN�Q�۰ʤɯšA�{���]�w�M�椤�w�����]�w�N" & vbCrLf & _
			"  �Q�ܧ�A�S�����]�w�N�Q�s�W�C" & vbCrLf & vbCrLf & _
			"- �ץX�]�w" & vbCrLf & _
			"  ���\�ץX�Ҧ��]�w���r�ɮסA�H�K�i�H�洫���ಾ�]�w�C" & vbCrLf & vbCrLf & _
			"���`�N�G�����x�s�����ɡA�N�۰ʧR���즳��m�����]�w���e�C" & vbCrLf & vbCrLf & _
			"���]�w���e��" & vbCrLf & _
			"============" & vbCrLf & _
			"<�����Ѽ�>" & vbCrLf & _
			"  - �ϥΪ���" & vbCrLf & _
			"    �Ӫ���w�]���uMicrosoft.XMLHTTP�v�A���i�ܧ�C" & vbCrLf & _
			"    �����Ӫ��󪺤@�ǨϥΤ�k�A�аѾ\�������C" & vbCrLf & vbCrLf & _
			"  - �������U ID" & vbCrLf & _
			"    ����½Ķ�����ݭn���U�~��ϥΡC�������D�лP½Ķ�������Ѱ��pô�C" & vbCrLf & _
			"    �{���w�g���ѤF Microsoft ½Ķ���������U ID�C�����O�ҥi�H�ä[�ϥΡC" & vbCrLf & vbCrLf & _
			"  - �������}" & vbCrLf & _
			"    ½Ķ�������s�����}�C" & vbCrLf & _
			"    ���`�N�G�\�h�s�u½Ķ���ѰӪ����}�ä��O�u����½Ķ�����s�����}�C�ݭn���R�����N�X��" & vbCrLf & _
			"    �@�@�@�@�߰ݯu����½Ķ�������ѰӤ~����o�C" & vbCrLf & vbCrLf & _
			"  - �ǰe���e�d��" & vbCrLf & _
			"    ������ڤW�O Microsoft.XMLHTTP �� Open ��k���� bstrUrl �ѼơC" & vbCrLf & _
			"    ���`�N�G�j�A�� {} �����r�����t�����A���i�ܧ�A�_�h�t�εL�k���ѡC" & vbCrLf & _
			"    �@�@�@�@�n�s�W�o�Ǩt�����A���I���k�䪺�u>�v���s�A����ݭn���t�����C" & vbCrLf & vbCrLf & _
			"  - ��ƶǰe�覡" & vbCrLf & _
			"    ������ڤW�O Microsoft.XMLHTTP ���� Open ��k���� bstrMethod �ѼơC" & vbCrLf & _
			"    �Y GET �� POST�C�� POST �覡�ǰe���,�i�H�j�� 4MB�A�]�i�H�� GET�A�u�� 256KB�C" & vbCrLf & vbCrLf & _
			"  - �P�B�覡" & vbCrLf & _
			"    ������ڤW�O Microsoft.XMLHTTP ���� Open ��k���� varAsync �ѼơC�i�ٲ��C" & vbCrLf & _
			"    �w�]�� True�A�Y�P�B����A���u��b DOM ����I�P�B����C�@��N��]�w�� False�A�Y���B����C" & vbCrLf & vbCrLf & _
			"  - �ϥΪ̦W" & vbCrLf & _
			"    ������ڤW�O Microsoft.XMLHTTP ���� Open ��k���� bstrUser �ѼơC�i�ٲ��C" & vbCrLf & vbCrLf & _
			"  - �K�X" & vbCrLf & _
			"    ������ڤW�O Microsoft.XMLHTTP ���� Open ��k���� bstrPassword �ѼơC�i�ٲ��C" & vbCrLf & vbCrLf & _
			"  - ���O��" & vbCrLf & _
			"    ������ڤW�O Microsoft.XMLHTTP ���� Send ��k���� varBody �ѼơC�i�ٲ��C" & vbCrLf & _
			"    ���i�H�O XML �榡��ơA�]�i�H�O�r��B��Ƭy�A�Ϊ̤@�ӵL�Ÿ���ư}�C�C�����O�q�L" & vbCrLf & _
			"    Open ��k���� URL �ѼƥN�J�C" & vbCrLf & vbCrLf & _
			"    �ǰe��ƪ��覡�����P�B�M���B��ءG" & vbCrLf & _
			"    (1) ���B�覡�G��ƥ]�@���ǰe�����A�N���� Send �i�{�A�Ȥ�������L���ާ@�C" & vbCrLf & _
			"    (2) �P�B�覡�G�Ȥ���n������A����^�T�{�T����~���� Send �i�{�C" & vbCrLf & vbCrLf & _
			"  - HTTP �Y�M��" & vbCrLf & _
			"    ������ڤW�O Microsoft.XMLHTTP ���󤤪� setRequestHeader ��k�AGET �覡�U�L�ġC" & vbCrLf & _
			"    (1) �����䴩�h�� setRequestHeader ���ءA�C�Ӷ��إΤ�����j�C" & vbCrLf & _
			"    (2) �C�Ӷ��ت��Y�W�� (bstrHeader) �M�� (bstrValue) �Υb�γr�����j�C" & vbCrLf & _
			"    (3) �ݭn�ǰe ContentT-Length ���خɡA��ȥi�H�ϥΨt�����A�t�αN�۰ʭp�����סC" & vbCrLf & vbCrLf & _
			"  - ��^���G�榡�Ѽ�" & vbCrLf & _
			"    ������ڤW�O Microsoft.XMLHTTP ���󤤪��@�Ӫ�^���G�榡�ݩʡC���H�U�X�ءG" & vbCrLf & _
			"    responseBody	Variant	��	���G��^���L�Ÿ���ư}�C" & vbCrLf & _
			"    responseStream	Variant	��	���G��^�� ADO Stream ����" & vbCrLf & _
			"    responseText	String	��	���G��^���r��" & vbCrLf & _
			"    responseXML	Object	��	���G��^�� XML �榡���" & vbCrLf & vbCrLf & _
			"    ���`�N�G���� responseXML �ɡA�u½Ķ�}�l�ۡv�M�u½Ķ������v�|���O��ܬ��u�� ID �j���v" & vbCrLf & _
			"    �@�@�@�@�M�u�����ҦW�j���v�C" & vbCrLf & vbCrLf & _
			"  - ½Ķ�}�l��" & vbCrLf & _
			"    �����Ω���Ѫ�^���G����½Ķ�r��e���r��C" & vbCrLf & _
			"    �i�I���k�䪺�u...�v���s�˵��q���A����^���Ҧ���r�����ӿ�J���r���C" & vbCrLf & _
			"    ���`�N�G�����䴩�Ρu|�v���j���h�өζ����ءC" & vbCrLf & vbCrLf & _
			"  - ½Ķ������" & vbCrLf & _
			"    �����Ω���Ѫ�^���G����½Ķ�r��᪺�r��C" & vbCrLf & _
			"    �i�I���k�䪺�u...�v���s�˵��q���A����^���Ҧ���r�����ӿ�J���r���C" & vbCrLf & _
			"    ���`�N�G�����䴩�Ρu|�v���j���h�өζ����ءC" & vbCrLf & vbCrLf & _
			"  - �� ID �j��" & vbCrLf & _
			"    �����ϥ� XML DOM �� getElementById ��k�j�M XML �ɮפ��㦳���w ID ������������r�ȡC" & vbCrLf & _
			"    �i�I���k�䪺�u...�v���s�˵��q���A����^���Ҧ���r�����ӿ�J�� ID ���C" & vbCrLf & _
			"    ���`�N�G�����䴩�Ρu|�v���j���h�өζ����ءC" & vbCrLf & vbCrLf & _
			"  - �����ҦW�j��" & vbCrLf & _
			"    �����ϥ� XML DOM �� getElementsByTagName ��k�j�M XML �ɮפ��Ҧ��㦳���w���ҦW�٪�" & vbCrLf & _
			"    ��������r�ȡC" & vbCrLf & _
			"    �i�I���k�䪺�u...�v���s�˵��q���A����^���Ҧ���r�����ӿ�J�����ҦW�١C" & vbCrLf & _
			"    ���`�N�G�����䴩�Ρu|�v���j���h�өζ����ءC" & vbCrLf & vbCrLf & _
			"<�y���t��>" & vbCrLf & _
			"  - �y���W��" & vbCrLf & _
			"    �w�]���y���W�ٲM��O�N Passolo ���y���M��i��F��²�A�h���F��a/�a�Ϫ��y�����C" & vbCrLf & vbCrLf & _
			"  - Passolo �N�X" & vbCrLf & _
			"    �w�]���y���N�X���� Passolo �y���M�椤�� ISO 936-1 �y���N�X�C�䤤�G" & vbCrLf & _
			"    ²�餤��M���餤�媺�y���N�X���� Passolo �y���M�椤����a/�a�ϻy���N�X�C" & vbCrLf & vbCrLf & _
			"  - ½Ķ�����N�X" & vbCrLf & _
			"    ½Ķ�����ҳW�w���y���N�X�C�ݭn���R�����N�X�θ߰�½Ķ�������ѰӤ~����o�C" & vbCrLf & vbCrLf & _
			"  - �W�[�y��" & vbCrLf & _
			"    �n�W�[�y���t��M�涵�A�I�� [�s�W] ���s�A�N�u�X�s�W��ܤ���C" & vbCrLf & vbCrLf & _
			"  - �R���y��" & vbCrLf & _
			"    �n�R���y���t��M�涵�A����M�椤�n�R�����y���A�M���I�� [�R��] ���s�C" & vbCrLf & vbCrLf & _
			"  - �s��y��" & vbCrLf & _
			"    �n�s��y���t��M�涵�A����M�椤�n�s�誺�y���A�M���I�� [�s��] ���s�C" & vbCrLf & vbCrLf & _
			"  - �~���s��y��" & vbCrLf & _
			"    �I�s���m�Υ~���s��{���s���ӻy���t��M��C" & vbCrLf & vbCrLf & _
			"  - �m�Ży��" & vbCrLf & _
			"    �n�N½Ķ�����N�X�m�šA����M�椤�n�m�Ū��y���A�M���I�� [�m��] ���s�C" & vbCrLf & _
			"    [�m��] ���s�|�ھ�½Ķ�����N�X�O�_���� (Null) ����ܤ��P���i�Ϊ��A�C" & vbCrLf & vbCrLf & _
			"  - ���]�y��" & vbCrLf & _
			"    �n�N½Ķ�����N�X���]����l�ȡA����M�椤�n���]���y���A�M���I�� [���]] ���s�C" & vbCrLf & _
			"    [���]] ���s�u���b½Ķ�����N�X�Q�ܧ��~���ର�i�Ϊ��A�C" & vbCrLf & vbCrLf & _
			"  - ��ܫD�Ŷ�" & vbCrLf & _
			"    �n�����½Ķ�����N�X���D�Ū��y�����A�I�� [��ܫD�Ŷ�] ���s�C�I����ӫ��s" & vbCrLf & _
			"    �N�ର���i�Ϊ��A�A[�������] �M [��ܪŶ�] ���s�N�ର�i�Ϊ��A�C" & vbCrLf & vbCrLf & _
			"  - ��ܪŶ�" & vbCrLf & _
			"    �n�����½Ķ�����N�X���� (Null) ���y�����A�I�� [��ܪŶ�] ���s�C�I����" & vbCrLf & _
			"    �ӫ��s�N�ର���i�Ϊ��A�A[�������] �M [��ܫD�Ŷ�] ���s�N�ର�i�Ϊ��A�C" & vbCrLf & vbCrLf & _
			"  - �������" & vbCrLf & _
			"    �n��ܥ����y�����A�I�� [�������] ���s�C�I����ӫ��s�N�ର���i�Ϊ��A�A" & vbCrLf & _
			"    [��ܫD�Ŷ�] �M [��ܫD�Ŷ�] ���s�N�ର�i�Ϊ��A�C" & vbCrLf & vbCrLf & _
			"  �s�W�y����A�N�b�M�椤��ܨÿ�ܷs���y���ΥN�X��C" & vbCrLf & _
			"  �R���y����A�N�b�M�椤��ܩҧR���y�����e�@���y���C" & vbCrLf & _
			"  �s��y����A�N�b�M�椤��ܨëO����ܽs��᪺�y���ΥN�X��C" & vbCrLf & _
			"  �m�Ży����A�N�b�M�椤��ܨëO����ܸm�ū᪺�y���ΥN�X��C" & vbCrLf & _
			"  ���]�y����A�N�b�M�椤��ܨëO����ܭ��]�᪺�y���ΥN�X��C" & vbCrLf & vbCrLf & _
			"����L�\�ࡸ" & vbCrLf & _
			"============" & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �I���ӫ��s�A�N�u�X�ثe�����������T���C" & vbCrLf & vbCrLf & _
			"- Ū��" & vbCrLf & _
			"  �I���ӫ��s�A�N�ھڿ�ܳ]�w�����P�u�X�U�C���:" & vbCrLf & _
			"  (1) �w�]��" & vbCrLf & _
			"      Ū���w�]�]�w�ȡA����ܦb�]�w���e���C" & vbCrLf & _
			"      ���ȷ��ܪ��]�w���t�ιw�]���]�w�ɡA�~��ܸӿ��C" & vbCrLf & vbCrLf & _
			"  (2) ���" & vbCrLf & _
			"      Ū����ܳ]�w����l�ȡA����ܦb�]�w���e���C" & vbCrLf & _
			"      ���ȷ��ܳ]�w����l�Ȭ��D�ŮɡA�~��ܸӿ��C" & vbCrLf & vbCrLf & _
			"  (3) �ѷӭ�" & vbCrLf & _
			"      Ū����ܪ��ѷӳ]�w�ȡA����ܦb�]�w���e���C" & vbCrLf & _
			"      ���ӿ����ܰ���ܳ]�w�~���Ҧ��]�w�M��C" & vbCrLf & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �I���ӫ��s�A�N�u�X���չ�ܤ���A�H�K�ˬd�]�w�����T�ʡC" & vbCrLf & vbCrLf & _
			"- �M��" & vbCrLf & _
			"  �I���ӫ��s�A�N�M�Ų{���]�w�������ȡA�H��K���s��J�]�w�ȡC" & vbCrLf & vbCrLf & _
			"- �T�w" & vbCrLf & _
			"  �I���ӫ��s�A�N�x�s�]�w�������������ܧ�A�����]�w�����ê�^�D�����C" & vbCrLf & _
			"  �{���N�ϥ��ܧ�᪺�]�w�ȡC" & vbCrLf & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �I���ӫ��s�A���x�s�]�w�������������ܧ�A�����]�w�����ê�^�D�����C" & vbCrLf & _
			"  �{���N�ϥέ�Ӫ��]�w�ȡC" & vbCrLf & vbCrLf & vbCrLf
	TestUse ="��½Ķ������" & vbCrLf & _
			"============" & vbCrLf & _
			"- �n���ժ�½Ķ�����W�١C�n���½Ķ�����A�I��½Ķ�����M��C" & vbCrLf & vbCrLf & _
			"���ؼлy����" & vbCrLf & _
			"============" & vbCrLf & _
			"- ½Ķ���ؼлy���C�ӲM�����ܻy���t��M�椤½Ķ�����y���N�X�����Ū��y���C" & vbCrLf & vbCrLf & _
			"��½Ķ�M�桸" & vbCrLf & _
			"============" & vbCrLf & _
			"- ½Ķ��媺½Ķ�M�淽�C�ӲM��]�t�F�M�פ����Ҧ�½Ķ�M��C" & vbCrLf & vbCrLf & _
			"��Ū�J��ơ�" & vbCrLf & _
			"============" & vbCrLf & _
			"- ���w�q���½Ķ�M�椤Ū�J���r��ơC" & vbCrLf & vbCrLf & _
			"��½Ķ���e��" & vbCrLf & _
			"============" & vbCrLf & _
			"- �n½Ķ����r�C���½Ķ�M��B�ӷ��r��M½Ķ�r��ﶵ�ɦ۰ʱq��ܪ�½Ķ�M�椤Ū�J�r��C" & vbCrLf & _
			"- �QŪ�J���r��۰ʥΪŮ�s���C" & vbCrLf & vbCrLf & _
			"���ӷ��r�꡸" & vbCrLf & _
			"============" & vbCrLf & _
			"- ����ӿﶵ�ɡA�۰ʱq��ܪ�½Ķ�M�椤Ū�J�ӷ��r��C" & vbCrLf & vbCrLf & _
			"��½Ķ�r�꡸" & vbCrLf & _
			"============" & vbCrLf & _
			"- ����ӿﶵ�ɡA�۰ʱq��ܪ�½Ķ�M�椤Ū�J½Ķ�r��C" & vbCrLf & vbCrLf & _
			"��½Ķ���G��" & vbCrLf & _
			"============" & vbCrLf & _
			"- �q½Ķ������^�����G�C�ھڿ�ܪ��ӷ��r��M½Ķ�r��ﶵ�A���O���½Ķ��r�αq½Ķ����" & vbCrLf & _
			"  ��^��������r�C" & vbCrLf & vbCrLf & _
			"��½Ķ��r��" & vbCrLf & _
			"============" & vbCrLf & _
			"- �q½Ķ������^��½Ķ�C��½Ķ�O�q½Ķ������^��������r�����Ӥ����Ѽƪ��u½Ķ�}�l�ۡv�M" & vbCrLf & _
			"  �u½Ķ������v��쪺�]�w�A���X����G�̤�������r�C" & vbCrLf & vbCrLf & _
			"��������r��" & vbCrLf & _
			"============" & vbCrLf & _
			"- �q½Ķ������^��������r�C�Ӥ�r�i���٥]�A�����N�X�C�Q�θӤ�r�i�H���D�u½Ķ�}�l�ۡv�M" & vbCrLf & _
			"  �u½Ķ������v������Ӧp��]�w�C" & vbCrLf & vbCrLf & _
			"����L�\�ࡸ" & vbCrLf & _
			"============" & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �I���ӫ��s�A�N��������T���C" & vbCrLf & vbCrLf & _
			"- ½Ķ" & vbCrLf & _
			"  �I���ӫ��s�A�N���ӿ�ܪ�����i��½Ķ�C" & vbCrLf & vbCrLf & _
			"- �T���Y" & vbCrLf & _
			"  �I���ӫ��s�A�N��� HTTP �T���Y�A�H�A�Ѥ�r�s�X���T���C" & vbCrLf & vbCrLf & _
			"- �M��" & vbCrLf & _
			"  �N�{��½Ķ���e�M½Ķ���G�����M�šC�I����N�ӫ��s�N�ܬ��uŪ�J�v���A�C" & vbCrLf & vbCrLf & _
			"- Ū�J" & vbCrLf & _
			"  �A���q��ܪ�½Ķ�M�椤Ū�J�r��C�I����N�ӫ��s�N�ܬ��u�M�šv���A�C" & vbCrLf & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �������յ{���ê�^�]�w�����C" & vbCrLf & vbCrLf & vbCrLf
	CopyRight = "�����v�ŧi��" & vbCrLf & _
			"============" & vbCrLf & _
			"- ���n�骺���v�k�}�o�̩M�ק�̩Ҧ��A����H�i�H�K�O�ϥΡB�ק�B�ƻs�B���G���n��C" & vbCrLf & _
			"- �ק�B���G���n�饲���H���������ɮסA�õ����n���l�}�o�̥H�έק�̡C" & vbCrLf & _
			"- ���g�}�o�̩M�ק�̦P�N�A�����´�έӤH�A���o�Ω�ӷ~�n��B�ӷ~�άO�䥦��Q�ʬ��ʡC" & vbCrLf & _
			"- ��ϥΥ��n�骺��l�����A�H�Ψϥθg�L�H�ק諸�D��l�����ҳy�����l���M�l�`�A�}�o�̤�" & vbCrLf & _
			"  �Ӿ����d���C" & vbCrLf & _
			"- �ѩ󬰧K�O�n��A�}�o�̩M�ק�̨S���q�ȴ��ѳn��޳N�䴩�A�]�L�q�ȧ�i�Χ�s�����C" & vbCrLf & _
			"- �w��������~�ô��X��i�N���C�p�����~�Ϋ�ĳ�A�жǰe��: z_shangyi@163.com�C" & vbCrLf & vbCrLf & vbCrLf
	Thank = "���P�@�@�¡�" & vbCrLf & _
			"============" & vbCrLf & _
			"- ���n��b�ק�L�{���o��~�Ʒs�@���|�������աA�b����ܰJ�ߪ��P�¡I" & vbCrLf & vbCrLf & vbCrLf
	Contact = "���P���pô��" & vbCrLf & _
			"============" & vbCrLf & _
			"wanfu�Gz_shangyi@163.com" & vbCrLf & vbCrLf & vbCrLf
	Logs = "�P�¤���I�z������O�ڳ̤j���ʤO�I�P���w��ϥΧڭ̻s�@���n��I" & vbCrLf & vbCrLf & _
			"**************************************************" & vbCrLf & _
			"�ݭn��h�B��s�B��n���~�ơA�Ы��X:" & vbCrLf & _
			"�~�Ʒs�@�� -- http://www.hanzify.org" & vbCrLf & _
			"�~�Ʒs�@���׾� -- http://bbs.hanzify.org" & vbCrLf & _
			"**************************************************" & vbCrLf
	Else
	AboutTitle = "����"
	HelpTitle = "����"
	HelpTipTitle = "Passolo ���߷����"
	AboutWindows = " ���� "
	MainWindows = " ������ "
	SetWindows = " ���ô��� "
	TestWindows = " ���Դ��� "
	Lines = "-----------------------"
	Sys = "����汾��" & Version & vbCrLf & _
			"����ϵͳ��Windows XP/2000 ����ϵͳ" & vbCrLf & _
			"���ð汾������֧�ֺ괦��� Passolo 6.0 �����ϰ汾" & vbCrLf & _
			"�������ԣ��������ĺͷ������� (�Զ�ʶ��)" & vbCrLf & _
			"��Ȩ���У�����������" & vbCrLf & _
			"��Ȩ��ʽ��������" & vbCrLf & _
			"�ٷ���ҳ��http://www.hanzify.org" & vbCrLf & _
			"�� �� �ߣ����������ͳ�Ա wanfu (2010)" & vbCrLf & vbCrLf & vbCrLf
	Ement = "�����л�����" & vbCrLf & _
			"============" & vbCrLf & _
			"- ֧�ֺ괦��� Passolo 6.0 �����ϰ汾������" & vbCrLf & _
			"- Windows Script Host (WSH) ���� (VBS)������" & vbCrLf & _
			"- Adodb.Stream ����֧�� Utf-8��Unicode ����" & vbCrLf & _
			"- Microsoft.XMLHTTP ���󣬱���" & vbCrLf & _
			"- Microsoft.XMLDOM ���󣬽��� responseXML ���ظ�ʽ����" & vbCrLf & vbCrLf & vbCrLf
	Dec = "���������" & vbCrLf & _
			"============" & vbCrLf & _
			"Passolo ���߷������һ������ Passolo �����б��ִ������߷����������������¹��ܣ�" & vbCrLf & _
			"- �������߷��������Զ����� Passolo �����б��е��ִ�" & vbCrLf & _
			"- ������һЩ���������߷������棬�����Զ����������߷�������" & vbCrLf & _
			"- ��ѡ���ִ����͡������ִ��Լ��Է���ǰ����ִ����д���" & vbCrLf & _
			"- ���ɿ�ݼ�����ֹ�������������ꡢ���ڷ�����鲢���������еĴ���" & vbCrLf & _
			"- ���ÿ��Զ�����Զ����¹���" & vbCrLf & vbCrLf & _
			"��������������ļ���" & vbCrLf & _
			"- PSLWebTrans.bas (Passolo ���߷�����ļ� - �Ի���ʽ����)" & vbCrLf & _
			"- PSLWebTrans_Silent.bas (Passolo ���߷�����ļ� - ��Ĭ��ʽ����)" & vbCrLf & _
			"- PSLWebTrans.txt (��������˵���ļ�)" & vbCrLf & vbCrLf & vbCrLf
	Setup = "�װ������" & vbCrLf & _
			"============" & vbCrLf & _
			"- ����ѹ����ļ����Ƶ� Passolo ϵͳ�ļ����ж���� Macros �ļ�����" & vbCrLf & _
			"- �� Passolo �Ĺ��� -> �Զ��幤�߲˵�����Ӹú��ļ�������ò˵����ƣ�" & vbCrLf & _
			"  �˺�Ϳ��Ե����ò˵�ֱ�ӵ���" & vbCrLf & _
			"- ��Ĭ��ʽ���е����߷�����������ѡ���ִ��������ִ����ִ������Ի���ʽ����" & vbCrLf & _
 			"  �����߷�����������С�Ҫ����ú�����в�������ʹ�öԻ���ʽ����" & vbCrLf & vbCrLf & vbCrLf
	MainUse = "��������" & vbCrLf & _
			"============" & vbCrLf & _
			"�����Ѿ��ṩһЩ���������߷������档��Ҳ���԰������ [����] ��ť�Զ������á�" & vbCrLf & _
			"����Զ������ú��������ڷ��������б���ѡ����ʹ�õķ������档" & vbCrLf & _
			"�й��Զ��巭�����棬������öԻ��򣬵��� [����] ��ť�����İ����е�˵����" & vbCrLf & vbCrLf & _
			"���Դ�ġ�" & vbCrLf & _
			"============" & vbCrLf & _
			"�����Զ��г���ǰ�����б������ļ��г���ǰ�����б��Ŀ�������������Ŀ�����ԡ�" & vbCrLf & _
			"ͨ��ѡ��ͬ�����ԣ�����ѡ��ǰ�����б��е���Դ�ִ�������ͬһ�ļ��е����������б���" & vbCrLf & _
			"�ķ����ִ���" & vbCrLf & vbCrLf & _
			"- ѡ���뵱ǰ�����б�Ŀ��������ͬ������ʱ����ʹ�õ�ǰ�����б��е���Դ�ִ���Ϊ����Դ�ġ�" & vbCrLf & _
			"- ѡ���뵱ǰ�����б�Ŀ�����Բ�ͬ������ʱ����ʹ��ѡ�������������б��еķ����ִ���Ϊ����" & vbCrLf & _
			"  Դ�ġ��Է��㽫���еķ���ת�����������ԡ����罫�������ķ���ɷ������ġ�" & vbCrLf & vbCrLf & _
			"����ִ���" & vbCrLf & _
			"============" & vbCrLf & _
			"�ṩ��ȫ�����˵����Ի����ַ��������������汾����������ѡ����ѡ�" & vbCrLf & vbCrLf & _
			"- ���ѡ��ȫ����������������Զ�ȡ��ѡ��" & vbCrLf & _
			"- ���ѡ�����ȫ��ѡ����Զ�ȡ��ѡ��" & vbCrLf & _
			"- ������Զ�ѡ������ѡ���ѡ��ʱ�����������Զ�ȡ��ѡ��" & vbCrLf & vbCrLf & _
			"�������ִ���" & vbCrLf & _
			"============" & vbCrLf & _
			"�ṩ�˹���������֤��δ���롢ȫΪ���ֺͷ��š�ȫΪ��дӢ�ġ�ȫΪСдӢ�ĵ�ѡ�" & vbCrLf & vbCrLf & _
			"- ������Ŀ�����Զ�ѡ��" & vbCrLf & _
			"- ѡ��ȫΪ��дӢ�ġ�ȫΪСдӢ��ѡ��ʱ���Զ�ѡ��ȫΪ���ֺͷ���ѡ�����ȫΪ����" & vbCrLf & _
			"  �ͷ���ѡ��Զ���Ϊ������״̬��" & vbCrLf & vbCrLf & _
			"���ִ������" & vbCrLf & _
			"============" & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �����ṩ�˺Ϳ�ݼ�����ֹ����������������ͬ��Ĭ�����ã������ÿ��������ڴ���������" & vbCrLf & _
			"  ��Ҳ���԰������ [����] ��ť�Զ������á�" & vbCrLf & _
			"  ����Զ������ú��������������б���ѡ����ʹ�õ����á�" & vbCrLf & _
			"  �й��Զ������ã�������öԻ��򣬵��� [����] ��ť�����İ����е�˵����" & vbCrLf & vbCrLf & _
			"- �Զ�ѡ������" & vbCrLf & _
			"  ѡ����ѡ��ʱ�����������еĿ��������б��Զ�ѡ���뷭���б�Ŀ������ƥ������á�" & vbCrLf & _
			"  ��ע�⣺��ѡ��������߷������Ч��" & vbCrLf & _
			"  ��������Ҫ�������뵱ǰ�����б��Ŀ������ƥ�䣬�� [����] ��ť�����ִ����������������" & vbCrLf & _
			"  �������������Ӧ�����ԡ�" & vbCrLf & _
			"  ���������ھ�Ĭ��ʽ����£����򽫰����Զ� - ��ѡ - Ĭ�ϡ�˳��ѡ����Ӧ�����á�" & vbCrLf & vbCrLf & _
			"- ȥ����ݼ�" & vbCrLf & _
			"  �ڷ���ǰ���ִ��еĿ�ݼ�ɾ�����Ա㷭�����������ȷ���롣" & vbCrLf & vbCrLf & _
			"- ȥ��������" & vbCrLf & _
			"  �ڷ���ǰ���ִ��еļ�����ɾ�����Ա㷭�����������ȷ���롣" & vbCrLf & vbCrLf & _
			"- �滻�ض��ַ����ڷ����ԭ" & vbCrLf & _
			"  �ڷ���ǰʹ��ѡ�������ã��滻�ض��ַ����ڷ����ԭ��" & vbCrLf & vbCrLf & _
			"  Ҫ�����Щ�ض��ַ����������öԻ�����ַ��滻�ķ���ǰҪ�滻���ַ��ж��塣" & vbCrLf & _
			"  ��ע�⣺���ں���������⣬��Щ�ַ��ǲ����Ա���������޷���ʶ��" & vbCrLf & _
			"  ������������������淭������Щ�滻����ַ������滻���ַ��޷�����ԭ��" & vbCrLf & vbCrLf & _
			"- ���з���" & vbCrLf & _
			"  �ڷ���ǰ�����е��ִ����Ϊ�����ִ����з��룬�Ա���߷������淭�����ȷ�ȡ�" & vbCrLf & vbCrLf & _
			"- ������ݼ�����ֹ����������" & vbCrLf & _
			"  �ڷ����ʹ��ѡ�������ã���鲢���������еĴ���" & vbCrLf & _
			"  ������ʽ�Ϳ�ݼ�����ֹ����������������ȫ��ͬ��" & vbCrLf & vbCrLf & _
			"- �滻�ض��ַ�" & vbCrLf & _
			"  �ڷ����ʹ��ѡ�������ã��滻�������ض����ַ���" & vbCrLf & vbCrLf & _
			"  Ҫ�����Щ�ض��ַ����������öԻ�����ַ��滻�ķ����Ҫ�滻���ַ��ж��塣" & vbCrLf & _
			"  ��ע�⣺���ں���������⣬��Щ�ַ��ǲ����Ա���������޷���ʶ��" & vbCrLf & _
			"  ������������������淭������Щ�滻ǰ���ַ������޷����滻��" & vbCrLf & vbCrLf & _
			"������ѡ���" & vbCrLf & _
			"============" & vbCrLf & _
			"- ����ʱ�Զ���������ѡ��" & vbCrLf & _
			"  ѡ������ʱ�����ڰ� [����] ��ťʱ�Զ���������ѡ���´�����ʱ�����뱣���ѡ��" & vbCrLf & vbCrLf & _
			"- ��ʾ�����Ϣ" & vbCrLf & _
			"  ѡ������ʱ������ Passolo ��Ϣ���������������롢�������滻��Ϣ��" & vbCrLf & _
			"  �ڲ�ѡ����������£���������߳���������ٶȡ�" & vbCrLf & vbCrLf & _
			"- ��ӷ���ע��" & vbCrLf & _
			"  ѡ������ʱ�����ڷ����б���ִ�ע��������ѡ�����������Զ�����������ע�͡�" & vbCrLf & _
			"  ���ø�ע�Ϳ������ַ����ִ��ķ��������" & vbCrLf& vbCrLf & _
			"���������ܡ�" & vbCrLf & _
			"============" & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �����ð�ť�������ڶԻ��򡣲���ʾ������ܡ����л����������̼���Ȩ����Ϣ��" & vbCrLf & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �����ð�ť����������ǰ���ڵİ�����Ϣ��" & vbCrLf & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �����ð�ť�����������öԻ��򡣿��������öԻ��������ø��ֲ�����" & vbCrLf & vbCrLf & _
			"- ȷ��" & vbCrLf & _
			"  �����ð�ť�����ر����Ի��򣬲���ѡ����ѡ������ִ����롣" & vbCrLf & _
			"  ��ע�⣺�������ִ�����״̬������Ϊ������״̬����ʾ����" & vbCrLf& vbCrLf & _
			"- ȡ��" & vbCrLf & _
			"  �����ð�ť�����˳�����" & vbCrLf& vbCrLf & vbCrLf
	SetUse ="�������б��" & vbCrLf & _
			"============" & vbCrLf & _
			"- ѡ������" & vbCrLf & _
			"  Ҫѡ����������������б�" & vbCrLf & vbCrLf & _
			"- �������" & vbCrLf & _
			"  Ҫ������������ [���] ��ť���ڵ����ĶԻ������������ơ�" & vbCrLf & vbCrLf & _
			"- ��������" & vbCrLf & _
			"  Ҫ�����������ƣ�ѡ�������б���Ҫ���������ã�Ȼ�󵥻� [����] ��ť��" & vbCrLf & vbCrLf & _
			"- ɾ������" & vbCrLf & _
			"  Ҫɾ�������ѡ�������б���Ҫɾ�������ã�Ȼ�󵥻� [ɾ��] ��ť��" & vbCrLf & vbCrLf & _
			"������ú󣬽����б�����ʾ�µ����ã��������ݽ���ʾ��ֵ��" & vbCrLf & _
			"�������ú󣬽����б�����ʾ����������ã����������е�����ֵ���䡣" & vbCrLf & _
			"ɾ�����ú󣬽����б�����ʾ��һ�����ã��������ݽ���ʾ��һ������ֵ��" & vbCrLf & vbCrLf & _
			"������͡�" & vbCrLf & _
			"============" & vbCrLf & _
			"- �ļ�" & vbCrLf & _
			"  ���ý����ļ���ʽ�����ں������ļ����µ� Data �ļ����С�" & vbCrLf & vbCrLf & _
			"- ע���" & vbCrLf & _
			"  ���ý�������ע����е� HKCU\Software\VB and VBA Program Settings\WebTranslate ���¡�" & vbCrLf & vbCrLf & _
			"- ��������" & vbCrLf & _
			"  ��������������ļ��е������á����������ʱ�����Զ����������������б������е����ý�" & vbCrLf & _
			"  �����ģ�û�е����ý�����ӡ�" & vbCrLf & vbCrLf & _
			"- ��������" & vbCrLf & _
			"  �������������õ��ı��ļ����Ա���Խ�����ת�����á�" & vbCrLf & vbCrLf & _
			"��ע�⣺�л���������ʱ�����Զ�ɾ��ԭ��λ���е��������ݡ�" & vbCrLf & vbCrLf & _
			"���������ݡ�" & vbCrLf & _
			"============" & vbCrLf & _
			"<�������>" & vbCrLf & _
			"  - ʹ�ö���" & vbCrLf & _
			"    �ö���Ĭ��Ϊ��Microsoft.XMLHTTP�������ɸ��ġ�" & vbCrLf & _
			"    �йظö����һЩʹ�÷��������������ĵ���" & vbCrLf & vbCrLf & _
			"  - ����ע�� ID" & vbCrLf & _
			"    ��Щ����������Ҫע�����ʹ�á�����������뷭�������ṩ����ϵ��" & vbCrLf & _
			"    �����Ѿ��ṩ�� Microsoft ���������ע�� ID��������֤��������ʹ�á�" & vbCrLf & vbCrLf & _
			"  - ������ַ" & vbCrLf & _
			"    ��������ķ�����ַ��" & vbCrLf & _
			"    ��ע�⣺������߷����ṩ�̵���ַ�����������ķ������������ַ����Ҫ������ҳ�����" & vbCrLf & _
			"    ��������ѯ�������ķ��������ṩ�̲���ȡ�á�" & vbCrLf & vbCrLf & _
			"  - ��������ģ��" & vbCrLf & _
			"    ���ֶ�ʵ������ Microsoft.XMLHTTP �� Open �����е� bstrUrl ������" & vbCrLf & _
			"    ��ע�⣺������ {} �е��ַ�Ϊϵͳ�ֶΣ����ɸ��ģ�����ϵͳ�޷�ʶ��" & vbCrLf & _
			"    ��������Ҫ�����Щϵͳ�ֶΣ��뵥���ұߵġ�>����ť��ѡ����Ҫ��ϵͳ�ֶΡ�" & vbCrLf & vbCrLf & _
			"  - ���ݴ��ͷ�ʽ" & vbCrLf & _
			"    ���ֶ�ʵ������ Microsoft.XMLHTTP ����� Open �����е� bstrMethod ������" & vbCrLf & _
			"    �� GET �� POST���� POST ��ʽ��������,���Դ� 4MB��Ҳ����Ϊ GET��ֻ�� 256KB��" & vbCrLf & vbCrLf & _
			"  - ͬ����ʽ" & vbCrLf & _
			"    ���ֶ�ʵ������ Microsoft.XMLHTTP ����� Open �����е� varAsync ��������ʡ�ԡ�" & vbCrLf & _
			"    ȱʡΪ True����ͬ��ִ�У���ֻ���� DOM ��ʵʩͬ��ִ�С�һ�㽫����Ϊ False�����첽ִ�С�" & vbCrLf & vbCrLf & _
			"  - �û���" & vbCrLf & _
			"    ���ֶ�ʵ������ Microsoft.XMLHTTP ����� Open �����е� bstrUser ��������ʡ�ԡ�" & vbCrLf & vbCrLf & _
			"  - ����" & vbCrLf & _
			"    ���ֶ�ʵ������ Microsoft.XMLHTTP ����� Open �����е� bstrPassword ��������ʡ�ԡ�" & vbCrLf & vbCrLf & _
			"  - ָ�" & vbCrLf & _
			"    ���ֶ�ʵ������ Microsoft.XMLHTTP ����� Send �����е� varBody ��������ʡ�ԡ�" & vbCrLf & _
			"    �������� XML ��ʽ���ݣ�Ҳ�������ַ�������������һ���޷����������顣��ָ��ͨ��" & vbCrLf & _
			"    Open �����е� URL �������롣" & vbCrLf & vbCrLf & _
			"    �������ݵķ�ʽ��Ϊͬ�����첽���֣�" & vbCrLf & _
			"    (1) �첽��ʽ�����ݰ�һ��������ϣ��ͽ��� Send ���̣��ͻ���ִ�������Ĳ�����" & vbCrLf & _
			"    (2) ͬ����ʽ���ͻ���Ҫ�ȵ�����������ȷ����Ϣ��Ž��� Send ���̡�" & vbCrLf & vbCrLf & _
			"  - HTTP ͷ��ֵ" & vbCrLf & _
			"    ���ֶ�ʵ������ Microsoft.XMLHTTP �����е� setRequestHeader ������GET ��ʽ����Ч��" & vbCrLf & _
			"    (1) ���ֶ�֧�ֶ�� setRequestHeader ��Ŀ��ÿ����Ŀ�÷��зָ���" & vbCrLf & _
			"    (2) ÿ����Ŀ��ͷ���� (bstrHeader) ��ֵ (bstrValue) �ð�Ƕ��ŷָ���" & vbCrLf & _
			"    (3) ��Ҫ���� ContentT-Length ��Ŀʱ����ֵ����ʹ��ϵͳ�ֶΣ�ϵͳ���Զ������䳤�ȡ�" & vbCrLf & vbCrLf & _
			"  - ���ؽ����ʽ����" & vbCrLf & _
			"    ���ֶ�ʵ������ Microsoft.XMLHTTP �����е�һ�����ؽ����ʽ���ԡ������¼��֣�" & vbCrLf & _
			"    responseBody	Variant	��	�������Ϊ�޷�����������" & vbCrLf & _
			"    responseStream	Variant	��	�������Ϊ ADO Stream ����" & vbCrLf & _
			"    responseText	String	��	�������Ϊ�ַ���" & vbCrLf & _
			"    responseXML	Object	��	�������Ϊ XML ��ʽ����" & vbCrLf & vbCrLf & _
			"    ��ע�⣺��ѡ�� responseXML ʱ�������뿪ʼ�ԡ��͡��������������ֱ���ʾΪ���� ID ������" & vbCrLf & _
			"    ���������͡�����ǩ����������" & vbCrLf & vbCrLf & _
			"  - ���뿪ʼ��" & vbCrLf & _
			"    ���ֶ�����ʶ�𷵻ؽ���еķ����ִ�ǰ���ִ���" & vbCrLf & _
			"    �ɵ����ұߵġ�...����ť�鿴�ӷ��������ص������ı���Ӧ��������ַ���" & vbCrLf & _
			"    ��ע�⣺���ֶ�֧���á�|���ָ��Ķ��������Ŀ��" & vbCrLf & vbCrLf & _
			"  - ���������" & vbCrLf & _
			"    ���ֶ�����ʶ�𷵻ؽ���еķ����ִ�����ִ���" & vbCrLf & _
			"    �ɵ����ұߵġ�...����ť�鿴�ӷ��������ص������ı���Ӧ��������ַ���" & vbCrLf & _
			"    ��ע�⣺���ֶ�֧���á�|���ָ��Ķ��������Ŀ��" & vbCrLf & vbCrLf & _
			"  - �� ID ����" & vbCrLf & _
			"    ���ֶ�ʹ�� XML DOM �� getElementById �������� XML �ļ��о���ָ�� ID ��Ԫ���е��ı�ֵ��" & vbCrLf & _
			"    �ɵ����ұߵġ�...����ť�鿴�ӷ��������ص������ı���Ӧ������� ID �š�" & vbCrLf & _
			"    ��ע�⣺���ֶ�֧���á�|���ָ��Ķ��������Ŀ��" & vbCrLf & vbCrLf & _
			"  - ����ǩ������" & vbCrLf & _
			"    ���ֶ�ʹ�� XML DOM �� getElementsByTagName �������� XML �ļ������о���ָ����ǩ���Ƶ�" & vbCrLf & _
			"    Ԫ�ص��ı�ֵ��" & vbCrLf & _
			"    �ɵ����ұߵġ�...����ť�鿴�ӷ��������ص������ı���Ӧ������ı�ǩ���ơ�" & vbCrLf & _
			"    ��ע�⣺���ֶ�֧���á�|���ָ��Ķ��������Ŀ��" & vbCrLf & vbCrLf & _
			"<�������>" & vbCrLf & _
			"  - ��������" & vbCrLf & _
			"    Ĭ�ϵ����������б��ǽ� Passolo �������б�����˾���ȥ���˹���/�����������" & vbCrLf & vbCrLf & _
			"  - Passolo ����" & vbCrLf & _
			"    Ĭ�ϵ����Դ���ȡ�� Passolo �����б��е� ISO 936-1 ���Դ��롣���У�" & vbCrLf & _
			"    �������ĺͷ������ĵ����Դ���ȡ�� Passolo �����б��еĹ���/�������Դ��롣" & vbCrLf & vbCrLf & _
			"  - �����������" & vbCrLf & _
			"    �����������涨�����Դ��롣��Ҫ������ҳ�����ѯ�ʷ��������ṩ�̲���ȡ�á�" & vbCrLf & vbCrLf & _
			"  - ��������" & vbCrLf & _
			"    Ҫ������������б������ [���] ��ť����������ӶԻ���" & vbCrLf & vbCrLf & _
			"  - ɾ������" & vbCrLf & _
			"    Ҫɾ����������б��ѡ���б���Ҫɾ�������ԣ�Ȼ�󵥻� [ɾ��] ��ť��" & vbCrLf & vbCrLf & _
			"  - �༭����" & vbCrLf & _
			"    Ҫ�༭��������б��ѡ���б���Ҫ�༭�����ԣ�Ȼ�󵥻� [�༭] ��ť��" & vbCrLf & vbCrLf & _
			"  - �ⲿ�༭����" & vbCrLf & _
			"    �������û��ⲿ�༭����༭������������б�" & vbCrLf & vbCrLf & _
			"  - �ÿ�����" & vbCrLf & _
			"    Ҫ��������������ÿգ�ѡ���б���Ҫ�ÿյ����ԣ�Ȼ�󵥻� [�ÿ�] ��ť��" & vbCrLf & _
			"    [�ÿ�] ��ť����ݷ�����������Ƿ�Ϊ�� (Null) ����ʾ��ͬ�Ŀ���״̬��" & vbCrLf & vbCrLf & _
			"  - ��������" & vbCrLf & _
			"    Ҫ�����������������Ϊԭʼֵ��ѡ���б���Ҫ���õ����ԣ�Ȼ�󵥻� [����] ��ť��" & vbCrLf & _
			"    [����] ��ťֻ���ڷ���������뱻���ĺ����תΪ����״̬��" & vbCrLf & vbCrLf & _
			"  - ��ʾ�ǿ���" & vbCrLf & _
			"    Ҫ����ʾ�����������Ϊ�ǿյ���������� [��ʾ�ǿ���] ��ť��������ð�ť" & vbCrLf & _
			"    ��תΪ������״̬��[ȫ����ʾ] �� [��ʾ����] ��ť��תΪ����״̬��" & vbCrLf & vbCrLf & _
			"  - ��ʾ����" & vbCrLf & _
			"    Ҫ����ʾ�����������Ϊ�� (Null) ����������� [��ʾ����] ��ť��������" & vbCrLf & _
			"    �ð�ť��תΪ������״̬��[ȫ����ʾ] �� [��ʾ�ǿ���] ��ť��תΪ����״̬��" & vbCrLf & vbCrLf & _
			"  - ȫ����ʾ" & vbCrLf & _
			"    Ҫ��ʾȫ����������� [ȫ����ʾ] ��ť��������ð�ť��תΪ������״̬��" & vbCrLf & _
			"    [��ʾ�ǿ���] �� [��ʾ�ǿ���] ��ť��תΪ����״̬��" & vbCrLf & vbCrLf & _
			"  ������Ժ󣬽����б�����ʾ��ѡ���µ����Լ�����ԡ�" & vbCrLf & _
			"  ɾ�����Ժ󣬽����б���ѡ����ɾ�����Ե�ǰһ�����ԡ�" & vbCrLf & _
			"  �༭���Ժ󣬽����б�����ʾ������ѡ���༭������Լ�����ԡ�" & vbCrLf & _
			"  �ÿ����Ժ󣬽����б�����ʾ������ѡ���ÿպ�����Լ�����ԡ�" & vbCrLf & _
			"  �������Ժ󣬽����б�����ʾ������ѡ�����ú�����Լ�����ԡ�" & vbCrLf & vbCrLf & _
			"���������ܡ�" & vbCrLf & _
			"============" & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �����ð�ť����������ǰ���ڵİ�����Ϣ��" & vbCrLf & vbCrLf & _
			"- ��ȡ" & vbCrLf & _
			"  �����ð�ť��������ѡ�����õĲ�ͬ�������в˵�:" & vbCrLf & _
			"  (1) Ĭ��ֵ" & vbCrLf & _
			"      ��ȡĬ������ֵ������ʾ�����������С�" & vbCrLf & _
			"      �����ѡ��������ΪϵͳĬ�ϵ�����ʱ������ʾ�ò˵���" & vbCrLf & vbCrLf & _
			"  (2) ԭֵ" & vbCrLf & _
			"      ��ȡѡ�����õ�ԭʼֵ������ʾ�����������С�" & vbCrLf & _
			"      �����ѡ�����õ�ԭʼֵΪ�ǿ�ʱ������ʾ�ò˵���" & vbCrLf & vbCrLf & _
			"  (3) ����ֵ" & vbCrLf & _
			"      ��ȡѡ���Ĳ�������ֵ������ʾ�����������С�" & vbCrLf & _
			"      ��ò˵���ʾ��ѡ������������������б�" & vbCrLf & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �����ð�ť�����������ԶԻ����Ա������õ���ȷ�ԡ�" & vbCrLf & vbCrLf & _
			"- ���" & vbCrLf & _
			"  �����ð�ť��������������õ�ȫ��ֵ���Է���������������ֵ��" & vbCrLf & vbCrLf & _
			"- ȷ��" & vbCrLf & _
			"  �����ð�ť�����������ô����е��κθ��ģ��˳����ô��ڲ����������ڡ�" & vbCrLf & _
			"  ����ʹ�ø��ĺ������ֵ��" & vbCrLf & vbCrLf & _
			"- ȡ��" & vbCrLf & _
			"  �����ð�ť�����������ô����е��κθ��ģ��˳����ô��ڲ����������ڡ�" & vbCrLf & _
			"  ����ʹ��ԭ��������ֵ��" & vbCrLf & vbCrLf & vbCrLf
	TestUse ="��������" & vbCrLf & _
			"============" & vbCrLf & _
			"- Ҫ���Եķ����������ơ�Ҫѡ�������棬�������������б�" & vbCrLf & vbCrLf & _
			"��Ŀ�����ԡ�" & vbCrLf & _
			"============" & vbCrLf & _
			"- �����Ŀ�����ԡ����б����ʾ��������б��з����������Դ��벻Ϊ�յ����ԡ�" & vbCrLf & vbCrLf & _
			"����б��" & vbCrLf & _
			"============" & vbCrLf & _
			"- ����Դ�ĵķ����б�Դ�����б�����˷����е����з����б�" & vbCrLf & vbCrLf & _
			"�����������" & vbCrLf & _
			"============" & vbCrLf & _
			"- ָ����ѡ�������б��ж�����ִ�����" & vbCrLf & vbCrLf & _
			"������ݡ�" & vbCrLf & _
			"============" & vbCrLf & _
			"- Ҫ������ı���ѡ�����б���Դ�ִ��ͷ����ִ�ѡ��ʱ�Զ���ѡ���ķ����б��ж����ִ���" & vbCrLf & _
			"- ��������ִ��Զ��ÿո����ӡ�" & vbCrLf & vbCrLf & _
			"����Դ�ִ���" & vbCrLf & _
			"============" & vbCrLf & _
			"- ѡ���ѡ��ʱ���Զ���ѡ���ķ����б��ж�����Դ�ִ���" & vbCrLf & vbCrLf & _
			"����ִ���" & vbCrLf & _
			"============" & vbCrLf & _
			"- ѡ���ѡ��ʱ���Զ���ѡ���ķ����б��ж��뷭���ִ���" & vbCrLf & vbCrLf & _
			"�������" & vbCrLf & _
			"============" & vbCrLf & _
			"- �ӷ������淵�صĽ��������ѡ������Դ�ִ��ͷ����ִ�ѡ��ֱ���ʾ�����ı���ӷ�������" & vbCrLf & _
			"  ���ص�ȫ���ı���" & vbCrLf & vbCrLf & _
			"����ı���" & vbCrLf & _
			"============" & vbCrLf & _
			"- �ӷ������淵�صķ��롣�÷����Ǵӷ������淵�ص�ȫ���ı��а�����������ġ����뿪ʼ�ԡ���" & vbCrLf & _
			"  ��������������ֶε����ã�ȡ�����ڶ���֮����ı���" & vbCrLf & vbCrLf & _
			"��ȫ���ı���" & vbCrLf & _
			"============" & vbCrLf & _
			"- �ӷ������淵�ص�ȫ���ı������ı����ܻ�������ҳ���롣���ø��ı�����֪�������뿪ʼ�ԡ���" & vbCrLf & _
			"  ��������������ֶ�Ӧ��������á�" & vbCrLf & vbCrLf & _
			"���������ܡ�" & vbCrLf & _
			"============" & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �����ð�ť������ȡ������Ϣ��" & vbCrLf & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �����ð�ť��������ѡ�����������з��롣" & vbCrLf & vbCrLf & _
			"- ��Ӧͷ" & vbCrLf & _
			"  �����ð�ť������ȡ HTTP ��Ӧͷ�����˽��ı��������Ϣ��" & vbCrLf & vbCrLf & _
			"- ���" & vbCrLf & _
			"  �����з������ݺͷ�����ȫ����ա������󽫸ð�ť����Ϊ�����롱״̬��" & vbCrLf & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �ٴδ�ѡ���ķ����б��ж����ִ��������󽫸ð�ť����Ϊ����ա�״̬��" & vbCrLf & vbCrLf & _
			"- ȡ��" & vbCrLf & _
			"  �˳����Գ��򲢷������ô��ڡ�" & vbCrLf & vbCrLf & vbCrLf
	CopyRight = "���Ȩ������" & vbCrLf & _
			"============" & vbCrLf & _
			"- ������İ�Ȩ�鿪���ߺ��޸������У��κ��˿������ʹ�á��޸ġ����ơ�ɢ���������" & vbCrLf & _
			"- �޸ġ�ɢ������������渽��˵���ļ�����ע�����ԭʼ�������Լ��޸��ߡ�" & vbCrLf & _
			"- δ�������ߺ��޸���ͬ�⣬�κ���֯����ˣ�����������ҵ�������ҵ��������Ӫ���Ի��" & vbCrLf & _
			"- ��ʹ�ñ������ԭʼ�汾���Լ�ʹ�þ������޸ĵķ�ԭʼ�汾����ɵ���ʧ���𺦣������߲�" & vbCrLf & _
			"  �е��κ����Ρ�" & vbCrLf & _
			"- ����Ϊ�������������ߺ��޸���û�������ṩ�������֧�֣�Ҳ������Ľ�����°汾��" & vbCrLf & _
			"- ��ӭָ����������Ľ���������д�����飬�뷢�͵�: z_shangyi@163.com��" & vbCrLf & vbCrLf & vbCrLf
	Thank = "���¡���л��" & vbCrLf & _
			"============" & vbCrLf & _
			"- ��������޸Ĺ����еõ����������ͻ�Ա�Ĳ��ԣ��ڴ˱�ʾ���ĵĸ�л��" & vbCrLf & vbCrLf & vbCrLf
	Contact = "��������ϵ��" & vbCrLf & _
			"============" & vbCrLf & _
			"wanfu��z_shangyi@163.com" & vbCrLf & vbCrLf & vbCrLf
	Logs = "��л֧�֣�����֧���������Ķ�����ͬʱ��ӭʹ�����������������" & vbCrLf & vbCrLf & _
			"**************************************************" & vbCrLf & _
			"��Ҫ���ࡢ���¡����õĺ����������:" & vbCrLf & _
			"���������� -- http://www.hanzify.org" & vbCrLf & _
			"������������̳ -- http://bbs.hanzify.org" & vbCrLf & _
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


'�Զ����°���
Sub UpdateHelp(HelpTip As String)
	If OSLanguage = "0404" Then
	HelpTitle = "����"
	HelpTipTitle = "�۰ʧ�s"
	SetWindows = " �]�w���� "
	Lines = "-----------------------"
	SetUse ="����s�覡��" & vbCrLf & _
			"============" & vbCrLf & _
			"- �۰ʤU����s�æw��" & vbCrLf & _
			"  ��ܸӿﶵ�ɡA�{���N�ھڧ�s�W�v�����]�w�۰��ˬd��s�A�p�G�����즳�s�������i�ήɡA" & vbCrLf & _
			"  �N�b���x�D�ϥΪ̷N�������p�U�۰ʤU����s�æw�ˡC��s������N�u�X��ܤ���q���ϥΪ̡A��" & vbCrLf & _
			"  [�T�w] ���s��{���Y�����C" & vbCrLf & vbCrLf & _
			"- ����s�ɳq���ڡA�ѧڨM�w�U���æw��" & vbCrLf & _
			"  ��ܸӿﶵ�ɡA�{���N�ھڧ�s�W�v�����]�w�۰��ˬd��s�A�p�G�����즳�s�������i�ήɡA" & vbCrLf & _
			"  �N�u�X��ܤ�����ܨϥΪ̡A�p�G�ϥΪ̨M�w��s�A�{���N�U����s�æw�ˡC��s������N�u�X��" & vbCrLf & _
			"  �ܤ���q���ϥΪ̡A�� [�T�w] ���s��{���Y�����C" & vbCrLf & vbCrLf & _
			"- �����۰ʧ�s" & vbCrLf & _
			"  ��ܸӿﶵ�ɡA�{���N���ˬd��s�C" & vbCrLf & vbCrLf & _
			"���`�N�G�L�צ�ا�s�覡�A��s���\�õ����{���᳣�ݭn�ϥΪ̭��s�Ұʥ����C" & vbCrLf & vbCrLf & _
			"����s�W�v��" & vbCrLf & _
			"============" & vbCrLf & _
			"- �ˬd���j" & vbCrLf & _
			"  �ˬd��s���ɶ��g���C�{���b�ˬd�g�����u�ˬd�@���C" & vbCrLf & vbCrLf & _
			"- �̫��ˬd���" & vbCrLf & _
			"  �̫��ˬd�ɪ�����C�{���b�ˬd��s�ɦ۰ʰO���ˬd����A�æb�P�@������u�ˬd�@���C" & vbCrLf & vbCrLf & _
			"- �ˬd" & vbCrLf & _
			"  �I���ӫ��s��A�{���N�����ˬd���j�M�ˬd����i���s�ˬd�A�ç�s�ˬd����C" & vbCrLf & vbCrLf & _
			"  ���p�G�����즳�s�������i�ΡA�N�u�X��ܤ�����ܨϥΪ̡A�p�G�ϥΪ̨M�w��s�A�{���N�U����s" & vbCrLf & _
			"  �@�æw�ˡC��s������N�u�X��ܤ���q���ϥΪ̡A�� [�T�w] ���s��{���Y�����C" & vbCrLf & _
			"  ���p�G�S�������즳�s�������i�ΡA�N�u�X�ثe���W���������C�ô��ܬO�_���s�U����s�C" & vbCrLf & vbCrLf & _
			"����s���}�M�桸" & vbCrLf & _
			"============" & vbCrLf & _
			"���B����s���}�ϥΪ̥i�H�ۤv�w�q�C�ϥΪ̩w�q�����}�N�Q�u���ϥΡC" & vbCrLf & _
			"�ϥΪ̩w�q�����}�L�k�s���ɡA�N�ϥε{���}�o�̩w�q����s���}�C" & vbCrLf & vbCrLf & _
			"��RAR �����{����" & vbCrLf & _
			"================" & vbCrLf & _
			"�ѩ�{���]�ĥΤF RAR �榡���Y�A�G�ݭn�b�����W�w�˦������������{���C" & vbCrLf & _
			"�즸�ϥήɡA�{���N�۰ʷj�����U�����U�� RAR ���ɦW�w�]�����{���A�öi��A���]�w�C" & vbCrLf & _
			"�æb�C���ϥήɦ۰��ˬd�����{���O�_�٦s�b�A�p�G���s�b�N���s�j���ó]�w�C" & vbCrLf & vbCrLf & _
			"���`�N�G�{���w�]�䴩�������{�����GWinRAR�BWinZIP�B7z�C�p�G�������S���o�Ǹ����{���A" & vbCrLf & _
			"�@�@�@�@�ݭn��ʳ]�w�C" & vbCrLf & vbCrLf & _
			"- �{�����|" & vbCrLf & _
			"  �����{����������|�C�i�I���k�䪺 [...] ���s��u�s�W�C" & vbCrLf & vbCrLf & _
			"- �����Ѽ�" & vbCrLf & _
			"  �����{�������Y RAR �ɮ׮ɪ��R�O�C�ѼơC�䤤�G" & vbCrLf & _
			"  %1 �����Y�ɮסA%2 ���n�q���Y�]���^�����D�{���ɮסA%3 �������᪺�ɮ׸��|�C" & vbCrLf & _
			"  �o�ǰѼƬ����n�ѼơA���i�ʤ֨åB���i�Ψ�L�Ÿ��N���C�ܩ���ᶶ�ǡA�̷Ӹ����{����" & vbCrLf & _
			"  �R�O�C�W�h�C" & vbCrLf & _
			"  �i�I���k�䪺 [>] ���s��u�s�W�o�ǥ��n�ѼơC" & vbCrLf & vbCrLf & _
			"����L�\�ࡸ" & vbCrLf & _
			"============" & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �I���ӫ��s�A�N��������T���C" & vbCrLf & vbCrLf & _
			"- Ū��" & vbCrLf & _
			"  �I���ӫ��s�A�N�ھڿ�ܳ]�w�����P�u�X�U�C���:" & vbCrLf & _
			"  (1) �w�]��" & vbCrLf & _
			"      Ū���۰ʧ�s�]�w���w�]�ȡA����ܦb��s���}�M��M RAR �����{�����C" & vbCrLf & _
			"  (2) ���" & vbCrLf & _
			"      Ū���۰ʧ�s�]�w����l�ȡA����ܦb��s���}�M��M RAR �����{�����C" & vbCrLf & vbCrLf & _
			"- �M��" & vbCrLf & _
			"  �N��s���}�M��MRAR �����{���������e�����M�šC" & vbCrLf & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �I���ӫ��s�A�N���է�s���}�M��M RAR �����{���O�_���T�C" & vbCrLf & vbCrLf & _
			"- �T�w" & vbCrLf & _
			"  �I���ӫ��s�A�N�x�s�]�w�������������ܧ�A�����]�w�����ê�^�D�����C" & vbCrLf & _
			"  �{���N�ϥ��ܧ�᪺�]�w�ȡC" & vbCrLf & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �I���ӫ��s�A���x�s�]�w�������������ܧ�A�����]�w�����ê�^�D�����C" & vbCrLf & _
			"  �{���N�ϥέ�Ӫ��]�w�ȡC" & vbCrLf & vbCrLf & vbCrLf
	Logs = "�P�¤���I�z������O�ڳ̤j���ʤO�I�P���w��ϥΧڭ̻s�@���n��I" & vbCrLf & vbCrLf & _
			"**************************************************" & vbCrLf & _
			"�ݭn��h�B��s�B��n���~�ơA�Ы��X:" & vbCrLf & _
			"�~�Ʒs�@�� -- http://www.hanzify.org" & vbCrLf & _
			"�~�Ʒs�@���׾� -- http://bbs.hanzify.org" & vbCrLf & _
			"**************************************************" & vbCrLf
	Else
	HelpTitle = "����"
	HelpTipTitle = "�Զ�����"
	SetWindows = " ���ô��� "
	Lines = "-----------------------"
	SetUse ="����·�ʽ��" & vbCrLf & _
			"============" & vbCrLf & _
			"- �Զ����ظ��²���װ" & vbCrLf & _
			"  ѡ����ѡ��ʱ�����򽫸��ݸ���Ƶ���е������Զ������£������⵽���µİ汾����ʱ��" & vbCrLf & _
			"  ���ڲ������û������������Զ����ظ��²���װ��������Ϻ󽫵����Ի���֪ͨ�û�����" & vbCrLf & _
			"  [ȷ��] ��ť������˳���" & vbCrLf & vbCrLf & _
			"- �и���ʱ֪ͨ�ң����Ҿ������ز���װ" & vbCrLf & _
			"  ѡ����ѡ��ʱ�����򽫸��ݸ���Ƶ���е������Զ������£������⵽���µİ汾����ʱ��" & vbCrLf & _
			"  �������Ի�����ʾ�û�������û��������£��������ظ��²���װ��������Ϻ󽫵�����" & vbCrLf & _
			"  ����֪ͨ�û����� [ȷ��] ��ť������˳���" & vbCrLf & vbCrLf & _
			"- �ر��Զ�����" & vbCrLf & _
			"  ѡ����ѡ��ʱ�����򽫲������¡�" & vbCrLf & vbCrLf & _
			"��ע�⣺���ۺ��ָ��·�ʽ�����³ɹ����˳��������Ҫ�û����������ꡣ" & vbCrLf & vbCrLf & _
			"�����Ƶ�ʡ�" & vbCrLf & _
			"============" & vbCrLf & _
			"- �����" & vbCrLf & _
			"  �����µ�ʱ�����ڡ������ڼ��������ֻ���һ�Ρ�" & vbCrLf & vbCrLf & _
			"- ���������" & vbCrLf & _
			"  �����ʱ�����ڡ������ڼ�����ʱ�Զ���¼������ڣ�����ͬһ������ֻ���һ�Ρ�" & vbCrLf & vbCrLf & _
			"- ���" & vbCrLf & _
			"  �����ð�ť�󣬳��򽫺��Լ�����ͼ�����ڽ��и��¼�飬�����¼�����ڡ�" & vbCrLf & vbCrLf & _
			"  �������⵽���µİ汾���ã��������Ի�����ʾ�û�������û��������£��������ظ���" & vbCrLf & _
			"  ������װ��������Ϻ󽫵����Ի���֪ͨ�û����� [ȷ��] ��ť������˳���" & vbCrLf & _
			"  �����û�м�⵽���µİ汾���ã���������ǰ���ϵİ汾�š�����ʾ�Ƿ��������ظ��¡�" & vbCrLf & vbCrLf & _
			"�������ַ�б��" & vbCrLf & _
			"============" & vbCrLf & _
			"�˴��ĸ�����ַ�û������Լ����塣�û��������ַ��������ʹ�á�" & vbCrLf & _
			"�û��������ַ�޷�����ʱ����ʹ�ó��򿪷��߶���ĸ�����ַ��" & vbCrLf & vbCrLf & _
			"��RAR ��ѹ�����" & vbCrLf & _
			"================" & vbCrLf & _
			"���ڳ���������� RAR ��ʽѹ��������Ҫ�ڱ����ϰ�װ����Ӧ�Ľ�ѹ����" & vbCrLf & _
			"����ʹ��ʱ�������Զ�����ע�����ע��� RAR ��չ��Ĭ�Ͻ�ѹ���򣬲������ʵ������á�" & vbCrLf & _
			"����ÿ��ʹ��ʱ�Զ�����ѹ�����Ƿ񻹴��ڣ���������ڽ��������������á�" & vbCrLf & vbCrLf & _
			"��ע�⣺����Ĭ��֧�ֵĽ�ѹ����Ϊ��WinRAR��WinZIP��7z�����������û����Щ��ѹ����" & vbCrLf & _
			"����������Ҫ�ֶ����á�" & vbCrLf & vbCrLf & _
			"- ����·��" & vbCrLf & _
			"  ��ѹ���������·�����ɵ����ұߵ� [...] ��ť�ֹ���ӡ�" & vbCrLf & vbCrLf & _
			"- ��ѹ����" & vbCrLf & _
			"  ��ѹ�����ѹ�� RAR �ļ�ʱ�������в��������У�" & vbCrLf & _
			"  %1 Ϊѹ���ļ���%2 ΪҪ��ѹ��������ȡ���������ļ���%3 Ϊ��ѹ����ļ�·����" & vbCrLf & _
			"  ��Щ����Ϊ��Ҫ����������ȱ�ٲ��Ҳ������������Ŵ��档�����Ⱥ�˳�����ս�ѹ�����" & vbCrLf & _
			"  �����й���" & vbCrLf & _
			"  �ɵ����ұߵ� [>] ��ť�ֹ������Щ��Ҫ������" & vbCrLf & vbCrLf & _
			"���������ܡ�" & vbCrLf & _
			"============" & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �����ð�ť������ȡ������Ϣ��" & vbCrLf & vbCrLf & _
			"- ��ȡ" & vbCrLf & _
			"  �����ð�ť��������ѡ�����õĲ�ͬ�������в˵�:" & vbCrLf & _
			"  (1) Ĭ��ֵ" & vbCrLf & _
			"      ��ȡ�Զ��������õ�Ĭ��ֵ������ʾ�ڸ�����ַ�б�� RAR ��ѹ�����С�" & vbCrLf & _
			"  (2) ԭֵ" & vbCrLf & _
			"      ��ȡ�Զ��������õ�ԭʼֵ������ʾ�ڸ�����ַ�б�� RAR ��ѹ�����С�" & vbCrLf & vbCrLf & _
			"- ���" & vbCrLf & _
			"  ��������ַ�б��RAR ��ѹ�����е�����ȫ����ա�" & vbCrLf & vbCrLf & _
			"- ����" & vbCrLf & _
			"  �����ð�ť�������Ը�����ַ�б�� RAR ��ѹ�����Ƿ���ȷ��" & vbCrLf & vbCrLf & _
			"- ȷ��" & vbCrLf & _
			"  �����ð�ť�����������ô����е��κθ��ģ��˳����ô��ڲ����������ڡ�" & vbCrLf & _
			"  ����ʹ�ø��ĺ������ֵ��" & vbCrLf & vbCrLf & _
			"- ȡ��" & vbCrLf & _
			"  �����ð�ť�����������ô����е��κθ��ģ��˳����ô��ڲ����������ڡ�" & vbCrLf & _
			"  ����ʹ��ԭ��������ֵ��" & vbCrLf & vbCrLf & vbCrLf
	Logs = "��л֧�֣�����֧���������Ķ�����ͬʱ��ӭʹ�����������������" & vbCrLf & vbCrLf & _
			"**************************************************" & vbCrLf & _
			"��Ҫ���ࡢ���¡����õĺ����������:" & vbCrLf & _
			"���������� -- http://www.hanzify.org" & vbCrLf & _
			"������������̳ -- http://bbs.hanzify.org" & vbCrLf & _
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
