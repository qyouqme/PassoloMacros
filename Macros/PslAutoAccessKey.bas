'' Check and Set Access Key Automatically in Form of XXX(&X)
'' (c) 2007-2008 by gnatix (Last modified on 2008.02.24)
'' Modified by wanfu (Last modified on 2010.11.11)

Public trn As PslTransList,TransString As PslTransString,OSLanguage As String

Public SpaceTrn As String,acckeyTrn As String,ExpStringTrn As String,EndStringTrn As String
Public acckeySrc As String,EndStringSrc As String,ShortcutSrc As String,ShortcutTrn As String
Public PreStringTrn As String,EndSpaceSrc As String,EndSpaceTrn As String

Public srcLineNum As Integer,trnLineNum As Integer,srcAccKeyNum As Integer,trnAccKeyNum As Integer
Public DefaultCheckList() As String,AppRepStr As String,PreRepStr As String
Public CheckDataList() As String,RepString As Integer

Private Const Version = "2010.11.11"
Private Const ToUpdateCheckVersion = "2010.09.25"
Private Const CheckRegKey = "HKCU\Software\VB and VBA Program Settings\AccessKey\"
Private Const CheckFilePath = MacroDir & "\Data\PSLCheckAccessKeys.dat"
Private Const JoinStr = vbBack
Private Const SubJoinStr = Chr$(1)
Private Const rSubJoinStr = Chr$(19) & Chr$(20)
Private Const LngJoinStr = "|"


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


' ��ȡ��ǰ�ִ������в���ִ����滻�����ַ���
Public Sub PSL_OnEditTransString(TransString As PslTransString)
	Dim srcString As String,trnString As String,OldtrnString As String,NewtrnString As String
	Dim LineNumErrCount As Integer,accKeyNumErrCount As Integer,ErrorCount As Integer
	Dim TranLang As String,Data As String

	On Error GoTo SysErrorMsg
	'������ר�õ��û����������Ƿ���ڣ����ڼ��˳�
	Set trn = PSL.ActiveTransList
	If Not trn Is Nothing Then
		If trn.Property(19980) = "CheckAccessKeys" Then Exit Sub
	End If

	'���ϵͳ����
	Dim strKeyPath As String, WshShell As Object
	Set WshShell = CreateObject("WScript.Shell")
	If WshShell Is Nothing Then
		PSL.Output("Your system is missing the Windows Script Host (WSH) object!")
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
		Msg01 = "�^��줤��"
		Msg02 = "�����^��"
		Msg03 =	"�z�� Passolo �����ӧC�A�������ȾA�Ω� Passolo 6.0 �ΥH�W�����A�Фɯū�A�ϥΡC"
	Else
		Msg01 = "Ӣ�ĵ�����"
		Msg02 = "���ĵ�Ӣ��"
		Msg03 =	"���� Passolo �汾̫�ͣ������������ Passolo 6.0 �����ϰ汾������������ʹ�á�"
	End If

	If PSL.Version < 600 Then
		PSL.Output Msg03
		Exit Sub
	End If

	'��ȡPSL��Ŀ�����Դ���
	trnLng = PSL.GetLangCode(trn.Language.LangID,pslCode639_1)
	If trnLng = "zh" Or trnLng = "ja" Or trnLng = "ko" Then TranLang = "Asia"
	If trnLng = "zh" Then
		trnLng = PSL.GetLangCode(trn.Language.LangID,pslCodeLangRgn)
		If trnLng = "zh-CHS" Or trnLng = "zh-SG" Then trnLng = "zh-CN"
		If trnLng = "zh-CHT" Or trnLng = "zh-HK" Or trnLng = "zh-MO" Then trnLng = "zh-TW"
	End If

	'��ʼ������
	ReDim DefaultCheckList(1)
	DefaultCheckList(0) = Msg01
	DefaultCheckList(1) = Msg02

	'��ȡ�ִ���������
	If CheckGet(trnLng,CheckDataList,"") = False Then
		If CheckGet("",CheckDataList,"") = False Then
			If TranLang = "Asia" Then CheckName = DefaultCheckList(0)
			If TranLang <> "Asia" Then CheckName = DefaultCheckList(1)
			If CheckGet(CheckName,CheckDataList,"") = False Then
				LangPair = Join(LangCodeList(CheckName,OSLanguage,1,107),SubJoinStr)
				Temp = CheckName & JoinStr & CheckSettings(CheckName,OSLanguage) & JoinStr & LangPair
				CheckDataList = Split(Temp,JoinStr)
			End If
		End If
	End If

	'����δ���롢��������ֻ�����ִ�
	If TransString.SourceText = TransString.Text Then GoTo ExitSub
	If TransString.State(pslStateTranslated) = False Then GoTo ExitSub
	If TransString.State(pslStateLocked) = True Then GoTo ExitSub
	If TransString.State(pslStateReadOnly) = True Then GoTo ExitSub

	'��Ϣ���ִ���ʼ������ȡԭ�ĺͷ����ִ�
	Massage = ""
	LineMsg = ""
	AcckeyMsg = ""
	ChangeMsg = ""
	srcString = TransString.SourceText
	trnString = TransString.Text
	OldtrnString = trnString

	'��ʼ�����ִ�
	NewtrnString = CheckHanding(srcString,trnString,TranLang)
	If RepString = 1 Then NewtrnString = ReplaceStr(NewtrnString,0)

	'������Ϣ���
	If srcLineNum <> trnLineNum Then
		LineMsg = LineErrMassage(srcLineNum,trnLineNum,LineNumErrCount)
	End If
	If srcAccKeyNum <> trnAccKeyNum Then
		AcckeyMsg = AccKeyErrMassage(srcAccKeyNum,trnAccKeyNum,accKeyNumErrCount)
	End If
	'If NewtrnString <> OldtrnString Then
		ChangeMsg = ReplaceMassage(OldtrnString,NewtrnString)
	'End If
	Massage = ChangeMsg & AcckeyMsg & LineMsg
	If Massage <> "" Then TransString.OutputError(Massage)

	'�滻�򸽼ӡ�ɾ�������ִ�
	If NewtrnString <> OldtrnString Then
		'TransString.OutputError(Msg32 & OldtrnString)
		TransString.Text = NewtrnString
	End If
	trn.Save

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
Function ReplaceStr(trnStr As String,fType As Integer) As String
	Dim i As Integer,BaktrnStr As String
	'��ȡѡ�����õĲ���
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
			TempArray = Split(FindStr,"|")
			If fType < 2 Then
				PreStr = TempArray(0)
				AppStr = TempArray(1)
			Else
				PreStr = TempArray(1)
				AppStr = TempArray(0)
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
Function CheckHanding(srcStr As String,trnStr As String,TranLang As String) As String
	Dim i As Integer,BaksrcStr As String,BaktrnStr As String,srcStrBak As String,trnStrBak As String
	Dim srcNum As Integer,trnNum As Integer,srcSplitNum As Integer,trnSplitNum As Integer
	Dim FindStrArr() As String,srcStrArr() As String,trnStrArr() As String,LineSplitArr() As String
	Dim posinSrc As Integer,posinTrn As Integer

	'��ȡѡ�����õĲ���
	SetsArray = Split(CheckDataList(1),SubJoinStr)
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
		BaktrnStr = StringReplace(BaksrcStr,BaktrnStr,TranLang)
	ElseIf srcSplitNum <> 0 Or trnSplitNum <> 0 Then
		LineSplitArr = MergeArray(srcStrArr,trnStrArr)
		BaktrnStr = ReplaceStrSplit(BaktrnStr,LineSplitArr,TranLang)
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
Function StringReplace(srcStr As String,trnStr As String,TranLang As String) As String
	Dim posinSrc As Integer,posinTrn As Integer,StringSrc As String,StringTrn As String
	Dim accesskeySrc As String,accesskeyTrn As String,Temp As String
	Dim ShortcutPosSrc As Integer,ShortcutPosTrn As Integer,PreTrn As String
	Dim EndStringPosSrc As Integer,EndStringPosTrn As Integer,AppTrn As String
	Dim preKeyTrn As String,appKeyTrn As String,Stemp As Boolean,FindStrArr() As String
	Dim i As Integer,j As Integer,x As Integer,y As Integer,m As Integer,n As Integer

	'��ȡѡ�����õĲ���
	SetsArray = Split(CheckDataList(1),SubJoinStr)
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
		Msg03 = " ��C"
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
		Msg03 = " "
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


'������������Ϣ
Sub sysErrorMassage(sysError As ErrObject)
	If OSLanguage = "0404" Then
		Msg01 = "���~"
		Msg02 = "�o�͵{���]�p�W�����~�C���~�N�X "
	Else
		Msg01 = "����"
		Msg02 = "������������ϵĴ��󡣴������ "
	End If
	PSL.Output(Msg02 & CStr(sysError.Number) & ", " & sysError.Description)
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


'��ȡ�����е�ÿ���ִ����滻����
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


'��ȡ�ִ��������
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
			'��ȡ Option ���ֵ
			If Header = "Option" And setPreStr <> "" Then
				If setPreStr = "Version" Then OldVersion = setAppStr
				If setPreStr = "AutoSelection" Then AutoSele = setAppStr
				If setPreStr = "AutoRepString" Then RepString = setAppStr
				If SelSet = "" Then
					If setPreStr = "AutoMacroSet" Then AutoMacroSet = setAppStr
					If AutoMacroSet = "Default" Then AutoMacroSet = DefaultCheckList(0)
				End If
			End If
			'��ȡ Option �����ȫ�����ֵ
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
				'���¾ɰ��Ĭ������ֵ
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
	'��ȡ Option ���ֵ
	OldVersion = GetSetting("AccessKey","Option","Version","")
	AutoSele = GetSetting("AccessKey","Option","AutoSelection",1)
	RepString = GetSetting("AccessKey","Option","AutoRepString",0)
	If SelSet = "" Then
		AutoMacroSet = GetSetting("AccessKey","Option","AutoMacroSet","")
		If AutoMacroSet = "Default" Then AutoMacroSet = DefaultCheckList(0)
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
						'���¾ɰ��Ĭ������ֵ
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


'����ָ��ֵ�Ƿ���������
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
