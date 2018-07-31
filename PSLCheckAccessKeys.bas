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
	Dim TypeMsg As Integer,CountMsg As Integer,HeaderID As Integer,Header As String
	If OSLanguage = "0404" Then
		Msg29 = "���~"
		Msg30 = "�п���n�B�z���r�������I"
		Msg31 = "�п���n�B�z���r�ꤺ�e�I"
		Msg32 = "�п���n�B�z���r�������M���e�I"
		Msg36 = "�L�k�x�s�I���ˬd�O�_���g�J�U�C��m���v��:" & vbCrLf & vbCrLf
	Else
		Msg29 = "����"
		Msg30 = "��ѡ��Ҫ������ִ����ͣ�"
		Msg31 = "��ѡ��Ҫ������ִ����ݣ�"
		Msg32 = "��ѡ��Ҫ������ִ����ͺ����ݣ�"
		Msg36 = "�޷����棡�����Ƿ���д������λ�õ�Ȩ��:" & vbCrLf & vbCrLf
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
	Case 2 ' ��ֵ���Ļ��߰����˰�ť
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
		'����ִ����ͺ�����ѡ���Ƿ�ͬʱѡ��ȫ������������
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
			'����ִ����ͺ�����ѡ���Ƿ�Ϊ��
			TypeMsg = mAllType + mMenu + mDialog + mString + mAccTable + mVer + mOther + mSelOnly
			CountMsg = mAllCont + mAccKey + mEndChar + mAcceler
			If TypeMsg = 0 And CountMsg <> 0 Then
				MsgBox(Msg30,vbOkOnly+vbInformation,Msg29)
				MainDlgFunc% = True ' ��ֹ���°�ť�رնԻ��򴰿�
			ElseIf TypeMsg <> 0 And CountMsg = 0 Then
				MsgBox(Msg31,vbOkOnly+vbInformation,Msg29)
				MainDlgFunc% = True ' ��ֹ���°�ť�رնԻ��򴰿�
			ElseIf TypeMsg = 0 And CountMsg = 0 Then
				MsgBox(Msg32,vbOkOnly+vbInformation,Msg29)
				MainDlgFunc% = True ' ��ֹ���°�ť�رնԻ��򴰿�
			End If
			If Join(cSelected) = nSelected Then Exit Function
			cSelected = Split(nSelected,JoinStr)
			If CheckWrite(CheckDataList,cWriteLoc,"Main") = False Then
				If WriteLocation = CheckFilePath Then Msg36 = Msg36 & CheckFilePath
				If WriteLocation = CheckRegKey Then Msg36 = Msg36 & CheckRegKey
				If MsgBox(Msg36,vbYesNo+vbInformation,Msg29) = vbNo Then Exit Function
				MainDlgFunc% = True ' ��ֹ���°�ť�رնԻ��򴰿�
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
			MainDlgFunc% = True ' ��ֹ���°�ť�رնԻ��򴰿�
		End If
	End Select
End Function


' ������
Sub Main
	Dim i As Integer,j As Integer,CheckVer As String,CheckSet As String,CheckID As Integer
	Dim srcString As String,trnString As String,OldtrnString As String,NewtrnString As String
	Dim LineNumErrCount As Integer,accKeyNumErrCount As Integer,ErrorCount As Integer
	Dim TransListOpen As Boolean,CheckDate As Date,TranDate As Date,CheckState As String
	Dim TranLang As String

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
		Msg00 = "(c) 2007-2008 by gnatix, 2009-2010 by wanfu  ����: " & Version
		Msg01 = "�K����B�פ�ũM�[�t���ˬd����"
		Msg02 = "���{���Ω��ˬd�B�ק�M�R����½Ķ�M�椤�Ҧ�½Ķ�r��" & _
				"���K����B�פ�ũM�[�t���C�Цb�ˬd��v���i��H�u�Ƭd�C"
		Msg03 = "½Ķ�M��: "

		Msg04 = "�]�w���"
		Msg05 = "�۰ʥ���:"
		Msg06 = "�ˬd����:"
		Msg07 = "�۰ʥ����M�ˬd�����ۦP(&U)"
		Msg08 = "�۰ʿ��(&X)"

		Msg09 = "�ˬd�аO"
		Msg10 = "��������(&B)"
		Msg11 = "�����]�w(&G)"
		Msg12 = "�������(&T)"
		Msg13 = "�������A(&Z)"
		Msg14 = "��������(&Q)"

		Msg15 = "�ާ@���"
		Msg16 = "���ˬd(&O)"
		Msg17 = "�ˬd�íץ�(&M)"
		Msg18 = "�R���K����(&R)"

		Msg19 = "�r������"
		Msg20 = "����(&A)"
		Msg21 = "���(&M)"
		Msg22 = "��ܤ��(&D)"
		Msg23 = "�r���(&S)"
		Msg24 = "�[�t����(&A)"
		Msg25 = "����(&V)"
		Msg26 = "��L(&O)"
		Msg27 = "�ȿ��(&L)"

		Msg28 = "�r�ꤺ�e"
		Msg29 = "����(&F)"
		Msg30 = "�K����(&K)"
		Msg31 = "�פ��(&E)"
		Msg32 = "�[�t��(&P)"

		Msg33 = "���L�r��"
		Msg34 = "���мf(&K)"
		Msg35 = "�w����(&E)"
		Msg36 = "��½Ķ(&N)"

		Msg37 = "��L�ﶵ"
		Msg38 = "�~��ɦ۰��x�s�Ҧ����(&V)"
		Msg39 = "���إߩΧR���ˬd�аO(&K)"
		Msg40 = "���ܧ��l½Ķ���A(&Y)"
		Msg41 = "�۰ʴ����S�w�r��(&L)"

		Msg42 = "����(&A)"
		Msg43 = "����(&H)"
		Msg44 = "�]�w(&S)"
		Msg45 = "�x�s���(&L)"

		Msg50 = "�T�{"
		Msg51 = "�T��"
		Msg52 = "���~"
		Msg53 =	"�z�� Passolo �����ӧC�A�������ȾA�Ω� Passolo 6.0 �ΥH�W�����A�Фɯū�A�ϥΡC"
		Msg54 = "�п���@��½Ķ�M��I"
		Msg55 = "���b�إߩM��s½Ķ�M��..."
		Msg56 = "�L�k�إߩM��s½Ķ�M��A���ˬd�z���M�׳]�w�C"
		Msg57 = "�ӲM�楼�Q�}�ҡC�����A�U�i�H�ˬd���~���L�k�ק�r��C" & vbCrLf & _
				"�z�ݭn���t�Φ۰ʶ}�Ҹ�½Ķ�M��ܡH" & vbCrLf & vbCrLf & _
				"���F�w���A�}���ˬd��A�p�G�����~�N���|�۰��x�s�ק�" & vbCrLf & _
				"������½Ķ�M��C�_�h�N�۰�����½Ķ�M��C"
		Msg58 = "���b�}��½Ķ�M��..."
		Msg59 = "�L�k�}��½Ķ�M��A���ˬd�z���M�׳]�w�C"
		Msg60 = "�ӲM��w�B��}�Ҫ��A�C�����A�U�i����~�ˬd�M�ק�N��" & vbCrLf & _
				"�z���x�s��½Ķ�L�k�٭�C���F�w���A�t�αN���x�s�z��½" & vbCrLf & _
				"Ķ�A�M��i���ˬd�M�ק�C" & vbCrLf & vbCrLf & _
				"�z�T�w�n���t�Φ۰��x�s�z��½Ķ�ܡH"
		Msg62 = "���b�ˬd�A�i��ݭn�X�����A�еy��..."
		Msg65 = "��Ķ��: "
		Msg66 = "�X�p�ή�: "
		Msg67 = "hh �p�� mm �� ss ��"
		Msg70 = "�^��줤��"
		Msg71 = "�����^��"
		Msg72 = "M:�{������"
		Msg73 = "M:�]�w�W��"
		Msg74 = "M:�ˬd���A"
		Msg75 = "M:�ˬd���"
		Msg76 = "�����~"
		Msg77 = "�L���~"
		Msg78 = "�w�ץ�"
		Msg79 = "yyyy�~m��d�� hh:mm:ss"
	Else
		Msg00 = "(c) 2007-2008 by gnatix, 2009-2010 by wanfu  �汾: " & Version
		Msg01 = "��ݼ�����ֹ���ͼ���������"
		Msg02 = "���������ڼ�顢�޸ĺ�ɾ���÷����б������з����ִ�" & _
				"�Ŀ�ݼ�����ֹ���ͼ����������ڼ������������˹����顣"
		Msg03 = "�����б�: "

		Msg04 = "����ѡ��"
		Msg05 = "�Զ���:"
		Msg06 = "����:"
		Msg07 = "�Զ���ͼ�����ͬ(&U)"
		Msg08 = "�Զ�ѡ��(&X)"

		Msg09 = "�����"
		Msg10 = "���԰汾(&B)"
		Msg11 = "��������(&G)"
		Msg12 = "��������(&T)"
		Msg13 = "����״̬(&Z)"
		Msg14 = "ȫ������(&Q)"

		Msg15 = "����ѡ��"
		Msg16 = "�����(&O)"
		Msg17 = "��鲢����(&M)"
		Msg18 = "ɾ����ݼ�(&R)"

		Msg19 = "�ִ�����"
		Msg20 = "ȫ��(&A)"
		Msg21 = "�˵�(&M)"
		Msg22 = "�Ի���(&D)"
		Msg23 = "�ִ���(&S)"
		Msg24 = "��������(&A)"
		Msg25 = "�汾(&V)"
		Msg26 = "����(&O)"
		Msg27 = "��ѡ��(&L)"

		Msg28 = "�ִ�����"
		Msg29 = "ȫ��(&F)"
		Msg30 = "��ݼ�(&K)"
		Msg31 = "��ֹ��(&E)"
		Msg32 = "������(&P)"

		Msg33 = "�����ִ�"
		Msg34 = "������(&K)"
		Msg35 = "����֤(&E)"
		Msg36 = "δ����(&N)"

		Msg37 = "����ѡ��"
		Msg38 = "����ʱ�Զ���������ѡ��(&V)"
		Msg39 = "��������ɾ�������(&K)"
		Msg40 = "������ԭʼ����״̬(&Y)"
		Msg41 = "�Զ��滻�ض��ַ�(&L)"

		Msg42 = "����(&A)"
		Msg43 = "����(&H)"
		Msg44 = "����(&S)"
		Msg45 = "����ѡ��(&L)"

		Msg50 = "ȷ��"
		Msg51 = "��Ϣ"
		Msg52 = "����"
		Msg53 =	"���� Passolo �汾̫�ͣ������������ Passolo 6.0 �����ϰ汾������������ʹ�á�"
		Msg54 = "��ѡ��һ�������б�"
		Msg55 = "���ڴ����͸��·����б�..."
		Msg56 = "�޷������͸��·����б��������ķ������á�"
		Msg57 = "���б�δ���򿪡���״̬�¿��Լ������޷��޸��ִ���" & vbCrLf & _
				"����Ҫ��ϵͳ�Զ��򿪸÷����б���" & vbCrLf & vbCrLf & _
				"Ϊ�˰�ȫ���򿪼�������ҵ����󽫲����Զ������޸�" & vbCrLf & _
				"���رշ����б������Զ��رշ����б�"
		Msg58 = "���ڴ򿪷����б�..."
		Msg59 = "�޷��򿪷����б��������ķ������á�"
		Msg60 = "���б��Ѵ��ڴ�״̬����״̬�½��д�������޸Ľ�ʹ" & vbCrLf & _
				"��δ����ķ����޷���ԭ��Ϊ�˰�ȫ��ϵͳ���ȱ������ķ�" & vbCrLf & _
				"�룬Ȼ����м����޸ġ�" & vbCrLf & vbCrLf & _
				"��ȷ��Ҫ��ϵͳ�Զ��������ķ�����"
		Msg62 = "���ڼ�飬������Ҫ�����ӣ����Ժ�..."
		Msg65 = "ԭ����: "
		Msg66 = "�ϼ���ʱ: "
		Msg67 = "hh Сʱ mm �� ss ��"
		Msg70 = "Ӣ�ĵ�����"
		Msg71 = "���ĵ�Ӣ��"
		Msg72 = "M:����汾"
		Msg73 = "M:��������"
		Msg74 = "M:���״̬"
		Msg75 = "M:�������"
		Msg76 = "�д���"
		Msg77 = "�޴���"
		Msg78 = "������"
		Msg79 = "yyyy��m��d�� hh:mm:ss"
	End If

	If PSL.Version < 600 Then
		MsgBox Msg53,vbOkOnly+vbInformation,Msg51
		Exit Sub
	End If

	Set trn = PSL.ActiveTransList
	'��ⷭ���б��Ƿ�ѡ��
	If trn Is Nothing Then
		MsgBox Msg54,vbOkOnly+vbInformation,Msg51
		Exit Sub
	End If

	'��ʼ������
	ReDim DefaultCheckList(1),CheckList(0),CheckDataList(0)
	DefaultCheckList(0) = Msg70
	DefaultCheckList(1) = Msg71

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

	'�Ի���
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
		CancelButton 510,434,90,21,.CancelButton '6 ȡ��
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

	'��ȡ�ִ��������
	If dlg.Menu = 1 Then StrTypes = "|Menu|"
	If dlg.Dialog = 1 Then StrTypes = StrTypes & "|Dialog|"
	If dlg.Strings = 1 Then StrTypes = StrTypes & "|StringTable|"
	If dlg.AccTable = 1 Then StrTypes = StrTypes & "|AcceleratorTable|"
	If dlg.Versions = 1 Then StrTypes = StrTypes & "|Version|"

	'��ʾ�򿪹رյķ����б������Ҫ�޸ĺ�ɾ��
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

	'��ʾ����򿪵ķ����б����⴦������ݲ��ɻָ�
	If trn.IsOpen = True And iVo <> 0 Then
		Massage = MsgBox(Msg03 & vbCrLf & vbCrLf & Msg60,vbYesNoCancel,Msg50)
 		If Massage = vbYes Then trn.Save
		If Massage = vbCancel Then GoTo ExitSub
	End If

	'��������б�ĸ���ʱ������ԭʼ�б��Զ�����
	If trn.SourceList.LastChange > trn.LastChange Then
		PSL.Output Msg55
		If trn.Update = False Then
			MsgBox Msg56,vbOkOnly+vbInformation,Msg51
			GoTo ExitSub
		End If
	End If

	'���ü���ר�õ��û���������
	If Not trn Is Nothing Then
		If trn.Property(19980) <> "CheckAccessKeys" Then
			trn.Property(19980) = "CheckAccessKeys"
		End If
	End If

	'��ȡPSL��Ŀ�����Դ���
	trnLng = PSL.GetLangCode(trn.Language.LangID,pslCode639_1)
	If trnLng = "zh" Or trnLng = "ja" Or trnLng = "ko" Then TranLang = "Asia"

	'�ͷŲ���ʹ�õĶ�̬������ʹ�õ��ڴ�
	Erase CheckListBak,CheckDataListBak,tempCheckList,tempCheckDataList

	'�����Ƿ�ѡ�� "��ѡ���ִ�" ������Ҫ�����ִ���
	If dlg.Seleted = 0 Then
		StringCount = trn.StringCount
	Else
		StringCount = trn.StringCount(pslSelection)
	End If

	'��ʼ�ִ�����
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
		'�����Ƿ�ѡ�� "��ѡ���ִ�" ������Ҫ������ִ�
		If dlg.Seleted = 0 Then Set TransString = trn.String(j)
		If dlg.Seleted = 1 Then Set TransString = trn.String(j,pslSelection)

		'��Ϣ���ִ���ʼ������ȡԭ�ĺͷ����ִ�
		Massage = ""
		LineMsg = ""
		AcckeyMsg = ""
		ChangeMsg = ""
		srcString = TransString.SourceText
		trnString = TransString.Text
		OldtrnString = trnString

		'ɾ�������
		If dlg.NoCheckTag = 1 Then TransString.Properties.RemoveAll

		'�ִ����ʹ���
		If dlg.AllType = 0 And dlg.Seleted = 0 Then
			If InStr("Menu|Dialog|StringTable|AcceleratorTable|Version",TransString.ResType) Then
				If InStr(StrTypes,TransString.ResType) = 0 Then GoTo Skip
			Else
				If dlg.Other = 0 Then GoTo Skip
			End If
		End If

		'����δ���롢��������ֻ�����ִ�
		If iVo <> 2 Then
			If srcString = trnString Then GoTo Skip
			If TransString.State(pslStateTranslated) = False Then GoTo Skip
		End If
		If TransString.State(pslStateLocked) = True Then GoTo Skip
		If TransString.State(pslStateReadOnly) = True Then GoTo Skip
		If Trim(srcString) = "" Then GoTo Skip

		'�����Ѽ���޴����Ҽ���������ڷ������ڵ��ִ�
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

		'��ʼ�����ִ�
		NewtrnString = CheckHanding(CheckID,srcString,trnString,TranLang)
		If dlg.AutoRepStr = 1 Then NewtrnString = ReplaceStr(CheckID,NewtrnString,0)

		'������Ϣ���
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

		'�滻�򸽼ӡ�ɾ�������ִ�
		If NewtrnString <> OldtrnString Then
			If iVo = 1 Then TransString.OutputError(Msg65 & OldtrnString)
			If iVo <> 0 Then TransString.Text = NewtrnString
		End If

		'���ĺ�ԭ�Ĳ�һ�µ������ִ�״̬,�Է���鿴
		If dlg.NoChangeState = 0 Then
			If NewtrnString <> OldtrnString Or srcLineNum <> trnLineNum Or srcAccKeyNum <> trnAccKeyNum Then
				TransString.State(pslStateReview) = True
			Else
				TransString.State(pslStateReview) = False
			End If
		End If

		'�����ִ����Զ��������Ժͼ������
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

	'�����������Ϣ���
	ErrorCount = ModifiedCount + AddedCount + WarningCount + LineNumErrCount + accKeyNumErrCount
	PSL.Output CountMassage(ErrorCount,LineNumErrCount,accKeyNumErrCount)
	If ErrorCount = 0 And TransListOpen = True Then trn.Close
	EndTimes = Timer
	PSL.Output Msg66 & Format(DateAdd("s",EndTimes - StartTimes,0),Msg67)

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


'����ִ��Ƿ���������ֺͷ���
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


' �޸���Ϣ���
Function ReplaceMassage(OldtrnString As String,NewtrnString As String) As String
	Dim AcckeyMsg As String,EndStringMsg As String,ShortcutMsg As String,Tmsg1 As String
	Dim Tmsg2 As String,Fmsg As String,Smsg As String,Massage1 As String,Massage2 As String
	Dim Massage3 As String,Massage4 As String,n As Integer,x As Integer,y As Integer

	If OSLanguage = "0404" And iVo = 0 Then
		Msg00 = "�K����P��夣�P"
		Msg01 = "�פ�ŻP��夣�P�Ϊ̻ݭn½Ķ"
		Msg02 = "�[�t���P��夣�P"
		Msg03 = "�K����M�פ�ŻP��夣�P"
		Msg04 = "�פ�ũM�[�t���P��夣�P"
		Msg05 = "�K����M�[�t���P��夣�P"
		Msg06 = "�K����B�פ�ũM�[�t���P��夣�P"
		Msg07 = "�K���䪺�j�p�g���P"
		Msg08 = "�K���䪺�j�p�g�M�פ�ŻP��夣�P"
		Msg09 = "�K���䪺�j�p�g�M�[�t���P��夣�P"
		Msg10 = "�K���䪺�j�p�g�B�פ�ũM�[�t���P��夣�P"

		Msg11 = "Ķ�夤�ʤ֫K����"
		Msg12 = "Ķ�夤�ʤֲפ��"
		Msg13 = "Ķ�夤�ʤ֥[�t��"
		Msg14 = "Ķ�夤�ʤ֫K����M�פ��"
		Msg15 = "Ķ�夤�ʤֲפ�ũM�[�t��"
		Msg16 = "Ķ�夤�ʤ֫K����M�[�t��"
		Msg17 = "Ķ�夤�ʤ֫K����B�פ�ũM�[�t��"

		Msg21 = "��Ķ�夤���K����"
		Msg22 = "��Ķ�夤���פ��"
		Msg23 = "��Ķ�夤���[�t��"
		Msg24 = "��Ķ�夤���K����M�פ��"
		Msg25 = "��Ķ�夤���פ�ũM�[�t��"
		Msg26 = "��Ķ�夤���K����M�[�t��"
		Msg27 = "��Ķ�夤���K����B�פ�ũM�[�t��"

		Msg31 = "�ݲ��ʫK�����̫�"
		Msg32 = "�ݲ��ʫK�����פ�ūe"
		Msg33 = "�ݲ��ʫK�����[�t���e"

		Msg41 = "�K����e���Ů�"
		Msg42 = "�פ�ūe���Ů�"
		Msg43 = "�[�t���e���Ů�"
		Msg44 = "�K����e���פ��"
		Msg45 = "�K����e���Ů�M�פ��"

		Msg46 = "�פ�ūe���Ů�����"
		Msg47 = "�פ�ū᪺�Ů�����"
		Msg48 = "�פ�ūe�᪺�Ů�����"
		Msg49 = "�r��᪺�Ů�����"
		Msg50 = "�פ�ūe�M�r��᪺�Ů�����"
		Msg51 = "�פ�ū�M�r��᪺�Ů�����"
		Msg52 = "�פ�ūe��M�r��᪺�Ů�����"
		Msg53 = "�פ�ūe���h�l���Ů�"
		Msg54 = "�פ�ūᦳ�h�l���Ů�"
		Msg55 = "�פ�ūe�ᦳ�h�l���Ů�"
		Msg56 = "�r��ᦳ�h�l���Ů�"
		Msg57 = "�פ�ūe�M�r��ᦳ�h�l���Ů�"
		Msg58 = "�פ�ū�M�r��ᦳ�h�l���Ů�"
		Msg59 = "�פ�ūe��M�r��ᦳ�h�l���Ů�"

		Msg61 = "�A"
		Msg62 = "��"
		Msg63 = "�C"
		Msg64 = "�B"
		Msg65 = "�w�N "
		Msg66 = " ������ "
		Msg67 = "�w�R�� "
	ElseIf OSLanguage = "0404" And iVo = 1 Then
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
	ElseIf OSLanguage <> "0404" And iVo = 0 Then
		Msg00 = "��ݼ���ԭ�Ĳ�ͬ"
		Msg01 = "��ֹ����ԭ�Ĳ�ͬ������Ҫ����"
		Msg02 = "��������ԭ�Ĳ�ͬ"
		Msg03 = "��ݼ�����ֹ����ԭ�Ĳ�ͬ"
		Msg04 = "��ֹ���ͼ�������ԭ�Ĳ�ͬ"
		Msg05 = "��ݼ��ͼ�������ԭ�Ĳ�ͬ"
		Msg06 = "��ݼ�����ֹ���ͼ�������ԭ�Ĳ�ͬ"
		Msg07 = "��ݼ��Ĵ�Сд��ͬ"
		Msg08 = "��ݼ��Ĵ�Сд����ֹ����ԭ�Ĳ�ͬ"
		Msg09 = "��ݼ��Ĵ�Сд�ͼ�������ԭ�Ĳ�ͬ"
		Msg10 = "��ݼ��Ĵ�Сд����ֹ���ͼ�������ԭ�Ĳ�ͬ"

		Msg11 = "������ȱ�ٿ�ݼ�"
		Msg12 = "������ȱ����ֹ��"
		Msg13 = "������ȱ�ټ�����"
		Msg14 = "������ȱ�ٿ�ݼ�����ֹ��"
		Msg15 = "������ȱ����ֹ���ͼ�����"
		Msg16 = "������ȱ�ٿ�ݼ��ͼ�����"
		Msg17 = "������ȱ�ٿ�ݼ�����ֹ���ͼ�����"

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

		Msg41 = "��ݼ�ǰ�пո�"
		Msg42 = "��ֹ��ǰ�пո�"
		Msg43 = "������ǰ�пո�"
		Msg44 = "��ݼ�ǰ����ֹ��"
		Msg45 = "��ݼ�ǰ�пո����ֹ��"

		Msg46 = "��ֹ��ǰ�Ŀո��ԭ����"
		Msg47 = "��ֹ����Ŀո��ԭ����"
		Msg48 = "��ֹ��ǰ��Ŀո��ԭ����"
		Msg49 = "�ִ���Ŀո��ԭ����"
		Msg50 = "��ֹ��ǰ���ִ���Ŀո��ԭ����"
		Msg51 = "��ֹ������ִ���Ŀո��ԭ����"
		Msg52 = "��ֹ��ǰ����ִ���Ŀո��ԭ����"
		Msg53 = "��ֹ��ǰ�ж���Ŀո�"
		Msg54 = "��ֹ�����ж���Ŀո�"
		Msg55 = "��ֹ��ǰ���ж���Ŀո�"
		Msg56 = "�ִ����ж���Ŀո�"
		Msg57 = "��ֹ��ǰ���ִ����ж���Ŀո�"
		Msg58 = "��ֹ������ִ����ж���Ŀո�"
		Msg59 = "��ֹ��ǰ����ִ����ж���Ŀո�"

		Msg61 = "��"
		Msg62 = "��"
		Msg63 = "��"
		Msg64 = "��"
		Msg65 = "�ѽ� "
		Msg66 = " �滻Ϊ "
		Msg67 = "��ɾ�� "
	ElseIf OSLanguage <> "0404" And iVo = 1 Then
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


' ɾ����ݼ���Ϣ���
Function DelAccKeyMassage(OldtrnString As String,NewtrnString As String) As String
	If OSLanguage = "0404" Then
		Msg01 = "�w�R���M��夣�P���K���� (%s)"
		Msg02 = "�w�R���M���ۦP���K���� (%s)"
		Msg03 = "�w�R���ȦbĶ�夤�s�b���K���� (%s)"
	Else
		Msg01 = "��ɾ����ԭ�Ĳ�ͬ�Ŀ�ݼ� (%s)"
		Msg02 = "��ɾ����ԭ����ͬ�Ŀ�ݼ� (%s)"
		Msg03 = "��ɾ�����������д��ڵĿ�ݼ� (%s)"
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


'���������Ϣ���
Function CountMassage(ErrorCount As Integer,LineNumErrCount As Integer,accKeyNumErrCount As Integer) As String
	If OSLanguage = "0404" Then
		Msg01 = "�S�������~�C"
		Msg02 = "��� " & ErrorCount & " �ӿ��~�C�䤤: "
		Msg03 = "�M��夣�P " & ModifiedCount & " �ӡA" & _
				"Ķ�夤�ʤ� " & AddedCount & " �ӡA" & _
				"�ȦbĶ�夤�s�b " & WarningCount & " �ӡA"
		Msg04 = "���MĶ�夤��Ƥ��P " & LineNumErrCount & " �ӡA" & _
				"���MĶ�夤�K����Ƥ��P " & accKeyNumErrCount & " �ӡC"
		Msg06 = "�w�ק� " & ModifiedCount & " �ӡA�w�s�W " & AddedCount & " �ӡA" & _
				"�n�ˬd " & WarningCount & " �ӡA��Ƥ��P " & LineNumErrCount & " �ӡA" & _
				"�K����Ƥ��P " & accKeyNumErrCount & " �ӡC"
		Msg07 = "�S���R������K����C"
		Msg08 = "�w�R�� " & ErrorCount & " �ӫK����C�䤤: "
		Msg09 = "�M��夣�P " & ModifiedCount & " �ӡA" & _
				"�M���ۦP " & AddedCount & " �ӡA" & _
				"�ȦbĶ�夤�s�b " & WarningCount & " �ӡC"
	Else
		Msg01 = "û���ҵ�����"
		Msg02 = "�ҵ� " & ErrorCount & " ����������: "
		Msg03 = "��ԭ�Ĳ�ͬ " & ModifiedCount & " ����" & _
				"������ȱ�� " & AddedCount & " ����" & _
				"���������д��� " & WarningCount & " ����"
		Msg04 = "ԭ�ĺ�������������ͬ " & LineNumErrCount & " ����" & _
				"ԭ�ĺ������п�ݼ�����ͬ " & accKeyNumErrCount & " ����"
		Msg06 = "���޸� " & ModifiedCount & " ��������� " & AddedCount & " ����" & _
				"Ҫ��� " & WarningCount & " ����������ͬ " & LineNumErrCount & " ����" & _
				"��ݼ�����ͬ " & accKeyNumErrCount & " ����"
		Msg07 = "û��ɾ���κο�ݼ���"
		Msg08 = "��ɾ�� " & ErrorCount & " ����ݼ�������: "
		Msg09 = "��ԭ�Ĳ�ͬ " & ModifiedCount & " ����" & _
				"��ԭ����ͬ " & AddedCount & " ����" & _
				"���������д��� " & WarningCount & " ����"
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


'������������Ϣ
Sub sysErrorMassage(sysError As ErrObject)
	If OSLanguage = "0404" Then
		Msg01 = "���~"
		Msg02 = "�o�͵{���]�p�W�����~�C���~�N�X "
	Else
		Msg01 = "����"
		Msg02 = "������������ϵĴ��󡣴������ "
	End If
	MsgBox(msg02 & CStr(sysError.Number) & ", " & sysError.Description,vbOkOnly+vbInformation,Msg01)
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


' ��;: ���˽���ת��Ϊʮ����
' ����: OctStr(�˽�����)
' �����������ͣ�String
' ���: OCTtoDEC(ʮ������)
' �����������: Long
' ���������Ϊ17777777777,��������Ϊ2147483647
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


' ��;: ��ʮ������ת��Ϊʮ����
' ����: HexStr(ʮ��������)
' ������������: String
' ���: HEXtoDEC(ʮ������)
' �����������: Long
' ���������Ϊ7FFFFFFF,��������Ϊ2147483647
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
Function Settings(CheckID As Integer) As Integer
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

		Msg33 = "�n�۰ʴ������r�� (�� | ���j�����e�᪺�r��) (�b�γr�����j):"
		Msg34 = "½Ķ��n�Q�������r�� (�� | ���j�����e�᪺�r��) (�b�γr�����j):"

		Msg35 = "����(&H)"
		Msg37 = "�r��B�z"
		Msg54 = ">"
		Msg55 = "..."

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

		Msg33 = "Ҫ�Զ��滻���ַ� (�� | �ָ��滻ǰ����ַ�) (��Ƕ��ŷָ�):"
		Msg34 = "�����Ҫ���滻���ַ� (�� | �ָ��滻ǰ����ַ�) (��Ƕ��ŷָ�):"

		Msg35 = "����(&H)"
		Msg37 = "�ִ�����"
		Msg54 = ">"
		Msg55 = "..."

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


'����ز鿴�Ի�������������˽������Ϣ��
Private Function SetFunc%(DlgItem$, Action%, SuppValue&)
	Dim Header As String,HeaderID As Integer,NewData As String,Path As String,cStemp As Boolean
	Dim i As Integer,n As Integer,TempArray() As String,Temp As String,LngName As String
	Dim LngID As Integer,LangArray() As String,AppLngList() As String,UseLngList() As String

	If OSLanguage = "0404" Then
		Msg01 = "���~"
		Msg02 = "�w�]��"
		Msg03 = "���"
		Msg04 = "�ѷӭ�"
		Msg08 = "����"
		Msg11 = "ĵ�i"
		Msg12 =	"�p�G�Y�ǰѼƬ��šA�N�ϵ{�����浲�G�����T�C" & vbCrLf & _
				"�z�T��Q�n�o�˰��ܡH"
		Msg13 = "�]�w���e�w�g�ܧ���O�S���x�s�I�O�_�ݭn�x�s�H"
		Msg14 = "�x�s�����w�g�ܧ���O�S���x�s�I�O�_�ݭn�x�s�H"
		Msg18 = "�ثe�]�w���A�ܤ֦��@���ѼƬ��šI" & vbCrLf
		Msg19 = "�Ҧ��]�w���A�ܤ֦��@���ѼƬ��šI" & vbCrLf
		Msg21 = "�T�{"
		Msg22 = "�T��n�R���]�w�u%s�v�ܡH"
		Msg24 = "�T��n�R���y���u%s�v�ܡH"
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
		Msg54 = "�i�λy��:"
		Msg55 = "�A�λy��:"
		Msg60 = "��������{��"
		Msg61 = "�i�����ɮ� (*.exe)|*.exe|�Ҧ��ɮ� (*.*)|*.*||"
		Msg62 = "�S�����w�����{���I�Э��s��J�ο���C"
		Msg63 = "�ɮװѷӰѼ�(%1)"
		Msg64 = "�n�^�����ɮװѼ�(%2)"
		Msg65 = "�������|�Ѽ�(%3)"
	Else
		Msg01 = "����"
		Msg02 = "Ĭ��ֵ"
		Msg03 = "ԭֵ"
		Msg04 = "����ֵ"
		Msg08 = "δ֪"
		Msg11 = "����"
		Msg12 =	"���ĳЩ����Ϊ�գ���ʹ�������н������ȷ��" & vbCrLf & _
				"��ȷʵ��Ҫ��������"
		Msg13 = "���������Ѿ����ĵ���û�б��棡�Ƿ���Ҫ���棿"
		Msg14 = "���������Ѿ����ĵ���û�б��棡�Ƿ���Ҫ���棿"
		Msg18 = "��ǰ�����У�������һ�����Ϊ�գ�" & vbCrLf
		Msg19 = "���������У�������һ�����Ϊ�գ�" & vbCrLf
		Msg21 = "ȷ��"
		Msg22 = "ȷʵҪɾ�����á�%s����"
		Msg24 = "ȷʵҪɾ�����ԡ�%s����"
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
		Msg54 = "��������:"
		Msg55 = "��������:"
		Msg60 = "ѡ���ѹ����"
		Msg61 = "��ִ���ļ� (*.exe)|*.exe|�����ļ� (*.*)|*.*||"
		Msg62 = "û��ָ����ѹ���������������ѡ��"
		Msg63 = "�ļ����ò���(%1)"
		Msg64 = "Ҫ��ȡ���ļ�����(%2)"
		Msg65 = "��ѹ·������(%3)"
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
	Case 1 ' �Ի��򴰿ڳ�ʼ��
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
	Case 2 ' ��ֵ���Ļ��߰����˰�ť
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
						SetFunc% = True '��ֹ���°�ť�رնԻ��򴰿�
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
			cUpdateSet = Split(Data,JoinStr)
			If DlgItem$ = "TestButton" Then Download(updateMethod,UpdateUrl,updateAsync,"4")
			If DlgItem$ = "HelpButton" Then Call UpdateHelp("SetHelp")
		End If

		If DlgItem$ = "OKButton" Then
			If CheckNullData("",CheckDataList,"4,11",1) = True Then
				If MsgBox(Msg19 & Msg12,vbYesNo+vbInformation,Msg11) = vbNo Then
					SetFunc% = True '��ֹ���°�ť�رնԻ��򴰿�
					Exit Function
				End If
			End If
			If DlgValue("cWriteType") = 0 Then cPath = CheckFilePath
			If DlgValue("cWriteType") = 1 Then cPath = CheckRegKey
			If CheckWrite(CheckDataList,cPath,"Sets") = False Then
				MsgBox(Msg36 & cPath,vbOkOnly+vbInformation,Msg01)
				SetFunc% = True '��ֹ���°�ť�رնԻ��򴰿�
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
			SetFunc% = True '��ֹ���°�ť�رնԻ��򴰿�
		End If
	Case 3 ' �ı��������Ͽ��ı�������
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
	'������º�����ݵ��ļ�
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
			'PSL.Output Key & " : " &  FindStr  '������
			KeyCode = UCase(Key) Like UCase(FindStr)
			If KeyCode = True Then CheckKeyCode = 1
			If KeyCode = True Then Exit For
		Next i
	ElseIf CheckKey = "" And Key <> "" Then
		CheckKeyCode = 1
	End If
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
			"  - �n�۰ʴ������r��" & vbCrLf & _
			"    �w�q�b�ˬd�íץ��L�{���n�Q�������r���H�δ����᪺�r���C" & vbCrLf & vbCrLf & _
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
			"  - Ҫ�Զ��滻���ַ�" & vbCrLf & _
			"    �����ڼ�鲢����������Ҫ���滻���ַ��Լ��滻����ַ���" & vbCrLf & vbCrLf & _
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
