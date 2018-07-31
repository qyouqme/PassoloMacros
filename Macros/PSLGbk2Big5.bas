''This Macro is to convert translations between simplified Chinese and traditional Chinese
''Idea and implemented by gnatix 2007.07.09 (List modified on 2008.12.15)
''Modified by wanfu (List modified on 2010.12.11)
''-----------------------------------------------------------------------------------------
''
Public Prj As PslProject,Src As PslSourceList,trn As PslTransList,prjFolder As String
Public OSLanguage As String,PSLVersion As String,ConCmd As String
Public ConTypeList() As String,ConDataList() As String,ConDataListBak() As String
Public ConTypeDataList() As String,ConListDataList() As String
Public AllStrList() As String,ConStrList() As String,ConStrListBak() As String
Public UpdateSet() As String,UpdateSetBak() As String,OldTextExpCharSet As Integer

Public WriteLoc As String,cSelected() As String,AddinID As String
Public ConCmdList() As String,ConCmdListBak() As String,ConCmdDataList() As String
Public ConCmdDataListBak() As String,ExpCodeList() As String,AddinIDList() As String

Public FileNo As Integer,FileText As String,FindText As String,FindLine As String
Public AppNames() As String,AppPaths() As String,FileList() As String,CodeList() As String

Private Const STARTF_USESHOWWINDOW& = &H1
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
Private Const Version = "2010.12.11"
Private Const ToUpdateVersion = "2010.09.06"
Private Const RegKey = "HKCU\Software\VB and VBA Program Settings\Gb2Big5"
Private Const FilePath = MacroDir & "\Data\PSLGbk2Big5.dat"
Private Const InputFile = "\~psltmp.txt"
Private Const OutputFile = "\~psltms.txt"
Private Const JoinStr = vbBack
Private Const SubJoinStr = Chr$(1)
Private Const rSubJoinStr = Chr$(19) & Chr$(20)

Private Const DefaultObject = "Microsoft.XMLHTTP"
Private Const updateAppName = "PSLGbk2Big5"
Private Const updateMainFile = "PSLGbk2Big5.bas"
Private Const updateINIFile = "PSLMacrosUpdates.ini"
Private Const updateMethod = "GET"
Private Const updateINIMainUrl = "ftp://hhdown:0011@czftp.hanzify.org/update/PSLMacrosUpdates.ini"
Private Const updateINIMinorUrl = "http://www.wanfutrade.com/software/hanhua/PSLMacrosUpdates.ini"
Private Const updateMainUrl = "ftp://hhdown:0011@czftp.hanzify.org/download/PSLGbk2Big5_Modified_wanfu.rar"
Private Const updateMinorUrl = "ftp://hhdown:0011@222.76.212.240:121/downloads/PSLGbk2Big5_Modified_wanfu.rar"
Private Const updateAsync = "False"

Private Type STARTUPINFO
	cb As Long
	lpReserved As String
	lpDesktop As String
	lpTitle As String
	dwX As Long
	dwY As Long
	dwXSize As Long
	dwYSize As Long
	dwXCountChars As Long
	dwYCountChars As Long
	dwFillAttribute As Long
	dwFlags As Long
	wShowWindow As Integer
	cbReserved2 As Integer
	lpReserved2 As Long
	hStdInput As Long
	hStdOutput As Long
	hStdError As Long
End Type

Private Type PROCESS_INFORMATION
	hProcess As Long
	hThread As Long
	dwProcessID As Long
	dwThreadID As Long
End Type

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
	hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
	lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
	lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
	ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
	ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
	lpStartupInfo As STARTUPINFO, lpProcessInformation As _
	PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal _
	hObject As Long) As Long

Public Sub ShellWait(Pathname As String, Optional WindowStyle As Long)
	Dim proc As PROCESS_INFORMATION,start As STARTUPINFO,ret As Long
	' Initialize the STARTUPINFO structure:
	With start
		.cb = Len(start)
		If Not IsMissing(WindowStyle) Then
        	.dwFlags = STARTF_USESHOWWINDOW
        	.wShowWindow = WindowStyle
		End If
	End With
	' Start the shelled application:
	ret& = CreateProcessA(0&, Pathname, 0&, 0&, 1&, _
    		NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
	' Wait for the shelled application to finish:
	ret& = WaitForSingleObject(proc.hProcess, INFINITE)
	ret& = CloseHandle(proc.hProcess)
End Sub


'Ĭ��ת�������������
Function DefaultSetting(CmdName As String,CmdPath As String) As String
	If LCase(CmdName) = "concmd" Or InStr(LCase(CmdPath),"concmd.exe") Then
		GBa = "/i:gbk /o:big5 %1"
		GBu = "/i:ule /o:ule %1"
		GB8 = "/i:utf8 /o:utf8 %1"
		GBFa = "/i:gbk /o:big5 /f:t %1"
		GBFu = "/i:ule /o:ule /f:t %1"
		GBF8 = "/i:utf8 /o:utf8 /f:t %1"

		BGa = "/i:big5 /o:gbk %1"
		BGu = "/i:ule /o:ule %1"
		BG8 = "/i:utf8 /o:utf8 %1"
		BGFa = "/i:big5 /o:gbk /f:s %1"
		BGFu = "/i:ule /o:ule /f:s %1"
		BGF8 = "/i:utf8 /o:utf8 /f:s %1"
		GBK = "GBfix.dat"
		BIG = "B5fix.dat"
		rGBK = 1
		rBIG = 1
	ElseIf LCase(CmdName) = "convertz" Or InStr(LCase(CmdPath),"convertz.exe") Then
		GBa = "/i:gbk /o:big5 /f:d %1"
		GBu = "/i:ule /o:ule /f:d %1"
		GB8 = "/i:utf8 /o:utf8 /f:d %1"
		GBFa = "/i:gbk /o:big5 /f:t %1"
		GBFu = "/i:ule /o:ule /f:t %1"
		GBF8 = "/i:utf8 /o:utf8 /f:t %1"

		BGa = "/i:big5 /o:gbk /f:d %1"
		BGu = "/i:ule /o:ule /f:d %1"
		BG8 = "/i:utf8 /o:utf8 /f:d %1"
		BGFa = "/i:big5 /o:gbk /f:s %1"
		BGFu = "/i:ule /o:ule /f:s %1"
		BGF8 = "/i:utf8 /o:utf8 /f:s %1"
		GBK = "BI_SimFix.dat"
		BIG = "BI_TradFix.dat"
		rGBK = 0
		rBIG = 0
	End If
	DefaultSetting = GBa & JoinStr & BGa & JoinStr & GBFa & JoinStr & BGFa & JoinStr & _
					GBu & JoinStr & BGu & JoinStr & GBFu & JoinStr & BGFu & JoinStr & GB8 & _
					JoinStr & BG8 & JoinStr & GBF8 & JoinStr & BGF8 & JoinStr & GBK & _
					JoinStr & BIG & JoinStr & rGBK & JoinStr & rBIG & JoinStr & JoinStr
End Function


'��ȡ�Զ���ת�����������
Function GetCustomSetting(CmdList() As String,CmdDataList() As String) As Boolean
	Dim i As Integer,FindConCMD As Integer,FindConvertZ As Integer,SettingArr() As String
	GetCustomSetting = False
	If GetSettings("",CmdList,CmdDataList,"") = True Then
		If ConCmd <> "" Then
			ConCmd = RemoveBackslash(ConCmd,"""","""",1)
			ConCmd = AppendBackslash(ConCmd,"","\",1)
		End If
		FindConCMD = 0
		FindConvertZ = 0
		For i = LBound(CmdDataList) To UBound(CmdDataList)
			ConCmdData = CmdDataList(i)
			SettingArr = Split(ConCmdData,JoinStr)
			CmdPath = SettingArr(1)
			GBKFile = SettingArr(14)
			BIGFile = SettingArr(15)
			If CmdPath <> "" And GBKFile <> "" And BIGFile <> "" Then
				If InStr(LCase(CmdPath),"concmd.exe") Then FindConCMD = 1
				If InStr(LCase(CmdPath),"convertz.exe") Then FindConvertZ = 1
			End If
			If FindConCMD + FindConvertZ = 2 Then Exit For
		Next i
		If FindConCMD = 0 And FindConvertZ = 1 Then GetDefaultSetting("ConCMD",CmdList,CmdDataList)
		If FindConCMD = 1 And FindConvertZ = 0 Then GetDefaultSetting("ConvertZ",CmdList,CmdDataList)
		If FindConCMD = 0 And FindConvertZ = 0 Then GetDefaultSetting("",CmdList,CmdDataList)
		GetCustomSetting = True
	End If
End Function


'��ȡĬ��ת�����������
Function GetDefaultSetting(CmdName As String,CmdList() As String,CmdDataList() As String) As Boolean
	Dim i As Integer,Path As String,WshShell As Object
	GetDefaultSetting = False
	On Error Resume Next
	strKeyPath = "HKLM\SOFTWARE\SDL Passolo GmbH\" & PSLVersion & "\System\InstallDir"
	Set WshShell = CreateObject("WScript.Shell")
	PSLPath = WshShell.RegRead(strKeyPath)
	PSLPath = RemoveBackslash(PSLPath,"""","""",1)
	PSLPath = AppendBackslash(PSLPath,"","\",1)
	Set WshShell = Nothing
	On Error GoTo 0
	ProgramsPath = Environ("ProgramFiles")
	ProgramsPath = RemoveBackslash(ProgramsPath,"""","""",1)
	ProgramsPath = AppendBackslash(ProgramsPath,"","\",1)
	If ConCmd <> "" Then
		ConCmdPath = RemoveBackslash(ConCmd,"","\",1)
		ConCmdPath = Left(ConCmdPath,InStrRev(ConCmdPath,"\"))
	End If
	For i = 0 To 2
		If i = 0 Then Path = ConCmdPath
		If i = 1 Then Path = PSLPath
		If i = 2 Then Path = ProgramsPath
		If CmdName = "" Or LCase(CmdName) = "concmd" Then
			If Path <> "" And Dir(Path & "ConCmd\ConCmd.exe") <> "" Then
				dCmdName = "ConCMD"
				CmdPath = Path & "ConCmd\ConCmd.exe"
				DefaultData = DefaultSetting(dCmdName,CmdPath)
				DefaultData = Replace(DefaultData,"GBfix.dat",Path & "ConCmd\GBfix.dat")
				DefaultData = Replace(DefaultData,"B5fix.dat",Path & "ConCmd\B5fix.dat")
				ConCmdData = dCmdName & JoinStr & CmdPath & JoinStr & DefaultData
				CreateArray(dCmdName,ConCmdData,CmdList,CmdDataList)
				GetDefaultSetting = True
			End If
		End If
		If CmdName = "" Or LCase(CmdName) = "convertz" Then
			If Path <> "" And Dir(Path & "ConvertZ\ConvertZ.exe") <> "" Then
				dCmdName = "ConvertZ"
				CmdPath = Path & "ConvertZ\ConvertZ.exe"
				DefaultData = DefaultSetting(dCmdName,CmdPath)
				DefaultData = Replace(DefaultData,"BI_SimFix.dat",Path & "ConvertZ\BI_SimFix.dat")
				DefaultData = Replace(DefaultData,"BI_TradFix.dat",Path & "ConvertZ\BI_TradFix.dat")
				ConCmdData = dCmdName & JoinStr & CmdPath & JoinStr & DefaultData
				CreateArray(dCmdName,ConCmdData,CmdList,CmdDataList)
				GetDefaultSetting = True
			End If
		End If
	Next i
	If CmdName = "" And GetDefaultSetting = False Then
		ConCmdData = JoinStr & JoinStr & JoinStr & JoinStr & JoinStr & JoinStr & JoinStr & _
					JoinStr & JoinStr & JoinStr &	JoinStr & JoinStr & JoinStr & JoinStr & _
					JoinStr & JoinStr & JoinStr & JoinStr & JoinStr
		CreateArray(dCmdName,ConCmdData,CmdList,CmdDataList)
		GetDefaultSetting = True
	End If
End Function


'��ȡ AddinID ����
Function GetAddinID(gAddinID As String) As String
	Dim i As Integer
	If gAddinID = "" Then
		If PSL.Option(pslOptionSystemLanguage) = 2052 Then
			gAddinID = "Passolo �ı���ʽ"
		ElseIf PSL.Option(pslOptionSystemLanguage) = 1028 Then
			gAddinID = "Passolo ��r�榡"
		Else
			gAddinID = "Passolo text format"
		End If
	End If

	If AddinIDTest(gAddinID) = True Then
		GetAddinID = gAddinID
	Else
		For i = 0 To 3
			If i = 0 Then gAddinID = "Passolo �ı���ʽ"
			If i = 1 Then gAddinID = "Passolo ���ָ�ʽ"
			If i = 2 Then gAddinID = "Passolo ��r�榡"
			If i = 3 Then gAddinID = "Passolo text format"
			If AddinIDTest(gAddinID) = True Then
				GetAddinID = gAddinID
				Exit For
			End If
		Next i
	End If
End Function


'������
Sub Main
	Dim WshShell As Object,strKeyPath As String,TextExpWriteTranslated As Integer,objStream As Object
	Dim i As Integer,n As Integer,ConCmdID As Integer,ConTypeID As Integer,ConID As Integer
	Dim SrcStemp As Boolean,TrnStemp As Boolean,Lan As PslLanguage
	Dim AllList() As String,ConList() As String,UseList() As String
	Dim AllHandle As Integer,FixID As Integer,UpdatedNum As Integer,TextExpCharSet As Integer
	Dim WordFixSelect As Integer,AllTypeSame As Integer,AllListSame As Integer,ExpCharSet As Integer

	On Error GoTo SysErrorMsg
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

	If OSLanguage = "0404" Then
		Msg00 = "(c) 2007-2008 by gnatix, 2009-2010 by wanfu  ����: " & Version
		Msg01 = "���~"
		Msg02 = "�ثe�S���}�ұM�סA�ж}�Ҥ@�ӱM�׫�A�աI"
		Msg03 = "�M�פ��S������ӷ��M½Ķ�M��I"
		Msg04 = "�T��"
		Msg05 =	"�z�� Passolo �����ӧC�A�������ȾA�Ω� Passolo 5.0 �ΥH�W�����A�Фɯū�A�ϥΡC"
		Msg06 = "²�餤��½Ķ --> ���餤��½Ķ"
		Msg07 = "²�餤���� --> ���餤��½Ķ"
		Msg08 = "���餤��½Ķ --> ²�餤��½Ķ"
		Msg09 = "���餤���� --> ²�餤��½Ķ"
		Msg10 = "²���餤���ഫ����"
		Msg11 = "�ഫ�ާ@"
		Msg12 = "�ഫ�M��"
		Msg13 = "�s��(&E)"
		Msg14 = "�����s�W"
		Msg15 = "�R��(&D)"
		Msg16 = "�M��(&K)"
		Msg17 = "���](&R)"
		Msg20 = "�ഫ�]�w"
		Msg21 = "�{���]�w:"
		Msg22 = "�ץX�W�q��:"
		Msg23 = "�ܧ�(&G)"
		Msg24 = "�ץX�r����:"
		Msg25 = "�ഫ�ɦ۰ʭץ��ؼлy�����J(&F)"
		Msg26 = "�B�z�Ҧ��ഫ�ާ@(&L)"
		Msg27 = "�Ҧ��ഫ�ާ@���]�w�ۦP(&M)"
		Msg28 = "�P�@�ഫ�ާ@���]�w�ۦP(&N)"
		Msg29 = "�~��ɦ۰��x�s�ഫ�]�w(&L)"
		Msg30 = "�B�z�������^(&Y)"
		Msg31 = "����(&A)"
		Msg32 = "�]�w(&P)"
		Msg33 = "�x�s�]�w(&S)"
		Msg36 = "�`�p�ഫ�F %s ��½Ķ�M��C"
		Msg66 = "�X�p�ή�: "
		Msg67 = "hh �p�� mm �� ss ��"
	Else
		Msg00 = "(c) 2007-2008 by gnatix, 2009-2010 by wanfu  �汾: " & Version
		Msg01 = "����"
		Msg02 = "Ŀǰû�д򿪷��������һ�����������ԣ�"
		Msg03 = "������û���κ���Դ�ͷ����б�"
		Msg04 = "��Ϣ"
		Msg05 =	"���� Passolo �汾̫�ͣ������������ Passolo 5.0 �����ϰ汾������������ʹ�á�"
		Msg06 = "�������ķ��� --> �������ķ���"
		Msg07 = "��������Դ�� --> �������ķ���"
		Msg08 = "�������ķ��� --> �������ķ���"
		Msg09 = "��������Դ�� --> �������ķ���"
		Msg10 = "��������ת����"
		Msg11 = "ת������"
		Msg12 = "ת���б�"
		Msg13 = "�༭(&E)"
		Msg14 = "ȫ�����"
		Msg15 = "ɾ��(&D)"
		Msg16 = "���(&K)"
		Msg17 = "����(&R)"
		Msg20 = "ת������"
		Msg21 = "��������:"
		Msg22 = "�������:"
		Msg23 = "����(&G)"
		Msg24 = "�����ַ���:"
		Msg25 = "ת��ʱ�Զ�����Ŀ�����Դʻ�(&F)"
		Msg26 = "��������ת������(&L)"
		Msg27 = "����ת��������������ͬ(&M)"
		Msg28 = "ͬһת��������������ͬ(&N)"
		Msg29 = "����ʱ�Զ�����ת������(&L)"
		Msg30 = "������ɺ󷵻�(&Y)"
		Msg31 = "����(&A)"
		Msg32 = "����(&P)"
		Msg33 = "��������(&S)"
		Msg36 = "�ܹ�ת���� %s �������б�"
		Msg66 = "�ϼ���ʱ: "
		Msg67 = "hh Сʱ mm �� ss ��"
	End If

	If PSL.Version < 500 Then
		MsgBox Msg05,vbOkOnly+vbInformation,Msg04
		Exit Sub
	End If

	Set Prj = PSL.ActiveProject
	If Prj Is Nothing Then
		MsgBox Msg02,vbOkOnly+vbInformation,Msg01
		Exit Sub
	End If
	If Prj.SourceLists.Count = 0 Then
		MsgBox Msg03,vbOkOnly+vbInformation,Msg01
		Exit Sub
	End If
	prjFolder = Prj.Location

	If Int(PSL.Version / 100) = 4 Then PSLVersion = "Passolo 4"
	If Int(PSL.Version / 100) = 5 Then PSLVersion = "Passolo 5"
	If Int(PSL.Version / 100) = 6 Then PSLVersion = "Passolo 6"
	If Int(PSL.Version / 100) = 7 Then PSLVersion = "Passolo 2007"
	If Int(PSL.Version / 100) = 8 Then PSLVersion = "Passolo 2009"
	If Int(PSL.Version / 100) = 11 Then PSLVersion = "Passolo 2011"
	On Error GoTo 0

	'��ȡ����� PSL WriteTranslated ע���ֵ
	On Error Resume Next
	strKeyPath = "HKCU\Software\PASS Engineering\" & PSLVersion & "\TextExport\WriteTranslated"
	TextExpWriteTranslated = WshShell.RegRead(strKeyPath)
	If TextExpWriteTranslated <> 1 Then WshShell.RegWrite(strKeyPath,"1","REG_DWORD")
	strKeyPath = "HKCU\Software\PASS Engineering\" & PSLVersion & "\TextExport\CharSet"
	OldTextExpCharSet = WshShell.RegRead(strKeyPath)
	Set objStream = CreateObject("Adodb.Stream")
	If objStream Is Nothing Then WshShell.RegWrite(strKeyPath,"0","REG_DWORD")
	On Error GoTo 0

	On Error GoTo SysErrorMsg
	'��ȡ�ַ������б�
	If objStream Is Nothing Then CodeList = CodePageList(0,0)
	If Not objStream Is Nothing Then CodeList = CodePageList(0,49)

	'��ȡ�����ַ���������
	ReDim ConCmdList(0),ConCmdDataList(0),AddinIDList(0)
	If objStream Is Nothing Then
		ReDim ExpCodeList(0)
		ExpCodeList(0) = "ANSI"
		ExpCode = "ANSI"
	Else
		ReDim ExpCodeList(2)
		ExpCodeList(0) = "ANSI"
		ExpCodeList(1) = "Unicode"
		ExpCodeList(2) = "UTF-8"
	End If

	'��ȡѡ���б��ȫ�������б�
	ReDim ConTypeList(0),ConDataList(0),AllList(0),ConList(0),AllStrList(0),ConStrList(0)
	ConTypeID = 0
	n = 0
	Set Lan = Prj.Languages("chs")
	If Not Lan Is Nothing Then
		If getList(0,2052,AllList,ConList) = True Then ConTypeID = n
		ConData = Msg06 & JoinStr & Join(AllList,SubJoinStr) & JoinStr & Join(ConList,SubJoinStr)
		CreateArray(Msg06,ConData,ConTypeList,ConDataList)
		n = n + 1
	End If
	Set Lan = Prj.Languages("cht")
	If Not Lan Is Nothing Then
		If getList(2,1028,AllList,ConList) = True Then ConTypeID = n
		ConData = Msg08 & JoinStr & Join(AllList,SubJoinStr) & JoinStr & Join(ConList,SubJoinStr)
		CreateArray(Msg08,ConData,ConTypeList,ConDataList)
		n = n + 1
	End If
	SrcStemp = False
	TrnStemp = False
	For i = 1 To Prj.SourceLists.Count
		Set Src = Prj.SourceLists(i)
		If Src.LangID = 2052 And SrcStemp = False Then
			If getList(1,2052,AllList,ConList) = True Then SrcStemp = True
			ConData = Msg07 & JoinStr & Join(AllList,SubJoinStr) & JoinStr & Join(ConList,SubJoinStr)
			CreateArray(Msg07,ConData,ConTypeList,ConDataList)
			If SrcStemp = True Then ConTypeID = n
			n = n + 1
			SrcStemp = True
		End If
		If Src.LangID = 1028 And TrnStemp = False Then
			If getList(3,1028,AllList,ConList) = True Then TrnStemp = True
			ConData = Msg09 & JoinStr & Join(AllList,SubJoinStr) & JoinStr & Join(ConList,SubJoinStr)
			CreateArray(Msg09,ConData,ConTypeList,ConDataList)
			If TrnStemp = True Then ConTypeID = n
			n = n + 1
			TrnStemp = True
		End If
		If SrcStemp = True And TrnStemp = True Then Exit For
	Next i

	'��ȡ������������б��ת����������б�
	If GetCustomSetting(ConCmdList,ConCmdDataList) = False Then
		GetDefaultSetting("",ConCmdList,ConCmdDataList)
	End If
	AddinID = GetAddinID(AddinID)

	'��ȡ�������ݲ�����°汾
	If Join(UpdateSet) <> "" Then
		updateMode = UpdateSet(0)
		updateUrl = UpdateSet(1)
		CmdPath = UpdateSet(2)
		CmdArg = UpdateSet(3)
		updateCycle = UpdateSet(4)
		updateDate = UpdateSet(5)
		If updateMode = "" Then
			UpdateSet(0) = "1"
			updateMode = "1"
		End If
		If updateUrl = "" Then UpdateSet(1) = updateMainUrl & vbCrLf & updateMinorUrl
		If CmdPath = "" Or (CmdPath <> "" And Dir(CmdPath) = "") Then
			CmdPathArgList = Split(getCMDPath(".rar","",""),JoinStr)
			UpdateSet(2) = CmdPathArgList(0)
			UpdateSet(3) = CmdPathArgList(1)
		End If
	Else
		updateMode = "1"
		updateUrl = updateMainUrl & vbCrLf & updateMinorUrl
		updateCycle = "7"
		Temp = updateMode & JoinStr & updateUrl & JoinStr & getCMDPath(".rar","","") & _
				JoinStr & updateCycle & JoinStr & updateDate
		UpdateSet = Split(Temp,JoinStr)
	End If
	If updateMode <> "" And updateMode <> "2" Then
		If updateDate <> "" Then
			i = CInt(DateDiff("d",CDate(updateDate),Date))
			m = StrComp(Format(Date,"yyyy-MM-dd"),updateDate)
			If updateCycle <> "" Then n = i - CInt(updateCycle)
		End If
		If updateDate = "" Or (m = 1 And n >= 0) Then
			If Download(updateMethod,updateUrl,updateAsync,updateMode) = True Then
				UpdateSet(5) = Format(Date,"yyyy-MM-dd")
				WriteSettings(ConCmdDataList,WriteLoc,"Update")
				GoTo ExitSub
			Else
				UpdateSet(5) = Format(Date,"yyyy-MM-dd")
				WriteSettings(ConCmdDataList,WriteLoc,"Update")
			End If
		End If
	End If

	If Join(cSelected) <> "" Then
		For i = LBound(ConCmdList) To UBound(ConCmdList)
			If ConCmdList(i) = cSelected(0) Then
				ConCmdID = i
				Exit For
			End If
		Next i
		ExpCharSet = StrToInteger(cSelected(2))
		WordFixSelect = StrToInteger(cSelected(3))
		AllTypeSame = StrToInteger(cSelected(5))
		AllListSame = StrToInteger(cSelected(6))
	Else
		ConCmdID = 0
		ExpCharSet = 0
		WordFixSelect = 0
		AllTypeSame = 1
		AllListSame = 1
	End If
	For i = LBound(ConDataList) To UBound(ConDataList)
		TempArray = Split(ConDataList(i),JoinStr)
		UseList = Split(TempArray(2),SubJoinStr)
		For n = LBound(UseList) To UBound(UseList)
			UseList(n) = UseList(n) & rSubJoinStr & ConCmdID & rSubJoinStr & _
							ExpCharSet & rSubJoinStr & WordFixSelect
		Next n
		ConDataList(i) = ConDataList(i) & JoinStr & Join(UseList,SubJoinStr) & JoinStr & _
							AllTypeSame & JoinStr & AllListSame
	Next i
	ConDataListBak = ConDataList

	MainDlg:
	Begin Dialog UserDialog 620,448,Msg10,.MainDlgFunc ' %GRID:10,7,1,1
		Text 20,7,580,14,Msg00,.Text1,2
		GroupBox 20,28,580,42,Msg11,.ConTypeGroup
		DropListBox 130,42,360,21,ConTypeList(),.ConTypeList

		GroupBox 20,77,580,161,Msg12,.ConListGroup
		ListBox 40,98,430,126,ConList(),.ConList
		PushButton 490,98,90,21,Msg13,.EditButton
		PushButton 490,119,90,21,Msg14,.AddAllButton
		PushButton 490,147,90,21,Msg15,.DelButton
		PushButton 490,168,90,21,Msg16,.CleanButton
		PushButton 490,196,90,21,Msg17,.ResetButton

		GroupBox 20,245,580,147,Msg20,.Settings
		Text 40,266,90,14,Msg21,.Text2
		Text 140,266,330,14,ConCmdName,.ConCmdNameText
		DropListBox 140,262,330,21,ConCmdList(),.ConCmdList

		Text 40,291,90,14,Msg22,.Text3
		Text 140,290,330,14,AddinID,.AddinIDText
		PushButton 490,287,90,21,Msg23,.AddinChangButton
		Text 40,315,90,14,Msg24,.Text4
		Text 140,315,330,14,ExpCode,.ExpCodeText
		DropListBox 140,311,330,21,ExpCodeList(),.ExpCodeList
		CheckBox 40,343,300,14,Msg25,.WordFixCheckBox
		CheckBox 350,343,230,14,Msg26,.AllHandleBox
		CheckBox 40,364,300,14,Msg27,.AllTypeSameBox
		CheckBox 350,364,230,14,Msg28,.AllListSameBox

		CheckBox 40,399,300,14,Msg29,.KeepSelet
		CheckBox 350,399,230,14,Msg30,.CycleBox
		PushButton 20,420,90,21,Msg31,.AboutButton
		PushButton 120,420,90,21,Msg32,.SetButton
		PushButton 220,420,110,21,Msg33,.SaveButton
		OKButton 410,420,90,21,.OKButton
		CancelButton 510,420,90,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	dlg.ConTypeList = ConTypeID
	If Dialog(dlg) = 0 Then GoTo ExitSub
	ConTypeID =	dlg.ConTypeList
	AllHandle = dlg.AllHandleBox

	'�ͷŲ���ʹ�õĶ�̬������ʹ�õ��ڴ�
	Erase TempArray,AllList,UseList
	Erase AllStrList,ConStrList,ConStrListBak
	Erase ConCmdListBak,ConCmdDataListBak,AddinIDList
	Erase AppNames,AppPaths,FileList

	StartTimes = Timer
	If AllHandle = 0 Then
		If ConTypeList(ConTypeID) = Msg06 Then ConID = 0
		If ConTypeList(ConTypeID) = Msg07 Then ConID = 1
		If ConTypeList(ConTypeID) = Msg08 Then ConID = 2
		If ConTypeList(ConTypeID) = Msg09 Then ConID = 3
		TempArray = Split(ConDataList(ConTypeID),JoinStr)
		UseList = Split(TempArray(3),SubJoinStr)
		UpdatedNum = Config(ConID,UseList,TextExpCharSet)
	Else
		For i = LBound(ConTypeList) To UBound(ConTypeList)
			If ConTypeList(i) = Msg06 Then ConID = 0
			If ConTypeList(i) = Msg07 Then ConID = 1
			If ConTypeList(i) = Msg08 Then ConID = 2
			If ConTypeList(i) = Msg09 Then ConID = 3
			TempArray = Split(ConDataList(i),JoinStr)
			UseList = Split(TempArray(3),SubJoinStr)
			UpdatedNum = UpdatedNum + Config(ConID,UseList,TextExpCharSet)
		Next i
	End If
	EndTimes = Timer
	PSL.Output Replace(Msg36,"%s",CStr(UpdatedNum))
	PSL.Output Msg66 & Format(DateAdd("s",EndTimes - StartTimes,0),Msg67)
	If dlg.CycleBox = 1 Then GoTo MainDlg

	ExitSub:
	If TextExpWriteTranslated <> 1 Then
		strKeyPath = "HKCU\Software\PASS Engineering\" & PSLVersion & "\TextExport\WriteTranslated"
		WshShell.RegWrite(strKeyPath,TextExpWriteTranslated,"REG_DWORD")
	End If
	If TextExpCharSet <> OldTextExpCharSet Then
		strKeyPath = "HKCU\Software\PASS Engineering\" & PSLVersion & "\TextExport\CharSet"
		WshShell.RegWrite(strKeyPath,OldTextExpCharSet,"REG_DWORD")
	End If
	Set WshShell = Nothing
	Set objStream = Nothing
	On Error GoTo 0
	Exit Sub

	'��ʾ���������Ϣ
	SysErrorMsg:
	Call sysErrorMassage(Err)
	GoTo ExitSub
End Sub


'������Ի�����
Private Function MainDlgFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Integer,n As Integer,ConTypeID As Integer,j As Integer,ConCmdID As Integer
	Dim AllList() As String,ConList() As String,UseList() As String
	Dim DataList() As String,TempDataList() As String
	Dim KeepItemSelect As Integer,TextExpCharSet As Integer,WordFixSelect As Integer
	Dim CycleSelect As Integer,AllHandle As Integer,Temp As String,Stemp As Boolean

	If OSLanguage = "0404" Then
		Msg01 = "���~"
		Msg02 = "�����w�Υ�������"
		Msg03 = "���(&S)"
		Msg04 = "�ܧ�(&G)"
		Msg07 = "�ഫ�M��"
		Msg08 = "�s�W(&A)"
		Msg09 = "�s��(&E)"
		Msg10 = "�T�{"
		Msg11 = "�T��n�N�ഫ�ާ@���ഫ�]�w�������Ҧ��M�泣�ۦP�ܡH" & vbCrLf & vbCrLf & _
				"�`�N�G�]�w���ۦP��A�Y�ӲM�檺�ഫ�]�w�Q�ܧ��A��" & vbCrLf & _
				"�L�M�泣�|�P���ܧ�C" & vbCrLf
		Msg12 = "�T��n�N�ഫ�ާ@���ഫ�]�w�������C�ӲM�泣���P�ܡH" & vbCrLf & vbCrLf  & _
				"�`�N�G�]�w�����P��A�Y�ӲM�檺�ഫ�]�w�Q�ܧ��A��" & vbCrLf & _
				"�L�M�椣�|�P���ܧ�C" & vbCrLf
		Msg13 = "�{�����M�|�۰ʦa�ʺA�O���C�ӲM�檺�ഫ�]�w����A�H" & vbCrLf & _
				"�K�i�H�H�ɦb�ۦP�M���ۦP���A���������Ӥ��|�򥢨C��" & vbCrLf & _
				"�M�檺�ഫ�]�w�C"
	Else
		Msg01 = "����"
		Msg02 = "δָ����δ��⵽"
		Msg03 = "ѡ��(&S)"
		Msg04 = "����(&G)"
		Msg07 = "ת���б�"
		Msg08 = "���(&A)"
		Msg09 = "�༭(&E)"
		Msg10 = "ȷ��"
		Msg11 = "ȷʵҪ��ת��������ת�������л�Ϊ�����б���ͬ��" & vbCrLf & vbCrLf & _
				"ע�⣺���ó���ͬ��ĳ���б��ת�����ñ����ĺ���" & vbCrLf & _
				"���б���ͬʱ���ġ�" & vbCrLf
		Msg12 = "ȷʵҪ��ת��������ת�������л�Ϊÿ���б���ͬ��" & vbCrLf & vbCrLf  & _
				"ע�⣺���óɲ�ͬ��ĳ���б��ת�����ñ����ĺ���" & vbCrLf & _
				"���б���ͬʱ���ġ�" & vbCrLf
		Msg13 = "������Ȼ���Զ��ض�̬��¼ÿ���б��ת������ѡ����" & vbCrLf & _
				"�������ʱ����ͬ�Ͳ���ͬ״̬֮���л������ᶪʧÿ��" & vbCrLf & _
				"�б��ת�����á�"
	End If
	Select Case Action%
	Case 1 ' �Ի��򴰿ڳ�ʼ��
		If Join(cSelected) <> "" Then
			DlgText "ConCmdList",cSelected(0)
			DlgValue "ExpCodeList",StrToInteger(cSelected(2))
			DlgValue "WordFixCheckBox",StrToInteger(cSelected(3))
			DlgValue "AllHandleBox",StrToInteger(cSelected(4))
			DlgValue "AllTypeSameBox",StrToInteger(cSelected(5))
			DlgValue "AllListSameBox",StrToInteger(cSelected(6))
			DlgValue "CycleBox",StrToInteger(cSelected(7))
			DlgValue "KeepSelet",StrToInteger(cSelected(8))
		End If

		If DlgText("ConTypeList") = "" Then DlgValue "ConTypeList",0
		ConTypeID = DlgValue("ConTypeList")
		TempArray = Split(ConDataList(ConTypeID),JoinStr)
		AllList = Split(TempArray(1),SubJoinStr)
		ConList = Split(TempArray(2),SubJoinStr)
		DlgListBoxArray "ConList",ConList()
		DlgValue "ConList",0
		If Join(ConList) = Join(AllList) Then DlgEnable "AddAllButton",False

		If DlgText("ConList") = "" Then
			DlgText "ConListGroup",Msg07 & "(0)"
			DlgText "EditButton",Msg08
			DlgEnable "DelButton",False
			DlgEnable "CleanButton",False
		Else
			DlgText "ConListGroup",Msg07 & "(" & UBound(ConList) + 1 & ")"
			DlgText "EditButton",Msg09
		End If

		CmdName = DlgText("ConCmdList")
		If CmdName = "" Then CmdName = ConCmdList(0)
		If CmdName <> "" Then
			Stemp = False
			For i = LBound(ConCmdList) To UBound(ConCmdList)
				If LCase(ConCmdList(i)) = LCase(CmdName) Then
					Stemp = True
					Exit For
				End If
			Next i
			If Stemp = False Then CmdName = ConCmdList(0)
		End If
		If UBound(ConCmdList) = 0 Then
			DlgVisible "ConCmdList",False
			DlgValue "ConCmdList",0
			If ConCmdList(0) = "" Then DlgText "ConCmdNameText",Msg02
			If ConCmdList(0) <> "" Then DlgText "ConCmdNameText",CmdName
		Else
			DlgVisible "ConCmdNameText",False
			DlgText "ConCmdList",CmdName
		End If

		If AddinID = "" Then
			DlgText "AddinIDText",Msg02
			DlgText "AddinChangButton",Msg03
		Else
			DlgText "AddinIDText",AddinID
			DlgText "AddinChangButton",Msg04
		End If

		If DlgValue("AllTypeSameBox") = 1 Then
			DlgValue "AllListSameBox",1
			DlgEnable "AllListSameBox",False
		End If

		If DlgText("ConCmdList") <> "" Then
			ConCmdID = DlgValue("ConCmdList")
			TextExpCharSet = DlgValue("ExpCodeList")
			TempArray = Split(ConCmdDataList(ConCmdID),JoinStr)
			ConCmdPath = TempArray(1)
			If TextExpCharSet = 0 Then
				GBF = TempArray(4)
				BGF = TempArray(5)
			ElseIf TextExpCharSet = 1 Then
				GBF = TempArray(8)
				BGF = TempArray(9)
			ElseIf TextExpCharSet = 2 Then
				GBF = TempArray(12)
				BGF = TempArray(13)
			End If
			GBKFixPath = TempArray(14)
			Big5FixPath = TempArray(15)
		End If
		Stemp = False
		If GBF = "" Or BGF = "" Then Stemp = True
		If InStr(LCase(ConCmdPath),"concmd.exe") Or InStr(LCase(ConCmdPath),"convertz.exe") Then
			If GBKFixPath = "" Or Big5FixPath = "" Then Stemp = True
		End If
		If Stemp = True Then
			DlgValue "WordFixCheckBox",0
			DlgEnable "WordFixCheckBox",False
		End If

		If CmdName = "" Or AddinID = Msg02 Or DlgText("ConList") = "" Then
			DlgEnable "OKButton",False
		Else
			If AllHandle = 0 Then
				If CheckNullData(CmdName,ConCmdDataList,"4,5,8,9,12,13,14,15,18,19",1) = True Then
					DlgEnable "OKButton",False
				End If
			Else
				If CheckNullData("",ConCmdDataList,"4,5,8,9,12,13,14,15,18,19",1) = True Then
					DlgEnable "OKButton",False
				End If
			End If
		End If

		If UBound(ExpCodeList) = 0 Then	DlgVisible "ExpCodeList",False
		If UBound(ExpCodeList) > 0 Then	DlgVisible "ExpCodeText",False
		TempArray = Split(ConDataListBak(ConTypeID),JoinStr)
		If TempArray(2) = Join(ConList,SubJoinStr) Then
			DlgEnable "ResetButton",False
		Else
			DlgEnable "ResetButton",True
		End If
		DlgEnable "SaveButton",False

		If Join(ConTypeDataList) = "" Then ConTypeDataList = ConDataList
		If Join(ConListDataList) = "" Then ConListDataList = ConDataList
	Case 2 ' ��ֵ���Ļ��߰����˰�ť
		ConTypeID = DlgValue("ConTypeList")
		DataList = Split(ConDataList(ConTypeID),JoinStr)
		AllList = Split(DataList(1),SubJoinStr)
		ConList = Split(DataList(2),SubJoinStr)
		UseList = Split(DataList(3),SubJoinStr)

		Stemp = False
		If DlgItem$ = "ConTypeList" Then Stemp = True
		If DlgItem$ = "EditButton" Then Stemp = True
		If DlgItem$ = "AddAllButton" Then Stemp = True
		If DlgItem$ = "DelButton" Then Stemp = True
		If DlgItem$ = "CleanButton" Then Stemp = True
		If DlgItem$ = "ResetButton" Then Stemp = True
		If Stemp = True Then
			If DlgItem$ = "ConTypeList" Then
				ConTypeID = DlgValue("ConTypeList")
				DataList = Split(ConDataList(ConTypeID),JoinStr)
				AllList = Split(DataList(1),SubJoinStr)
				ConList = Split(DataList(2),SubJoinStr)
				UseList = Split(DataList(3),SubJoinStr)
				TempArray = Split(UseList(0),rSubJoinStr)
				DlgListBoxArray "ConList",ConList()
				DlgValue "ConList",0
				DlgValue "ConCmdList",StrToInteger(TempArray(1))
				DlgValue "ExpCodeList",StrToInteger(TempArray(2))
				DlgValue "WordFixCheckBox",StrToInteger(TempArray(3))
				DlgValue "AllListSameBox",StrToInteger(DataList(5))
				If DlgValue("AllTypeSameBox") = 0 Then
					If DlgValue("AllListSameBox") = 0 Then ConDataList = ConTypeDataList
					If DlgValue("AllListSameBox") = 1 Then ConDataList = ConListDataList
				End If
			End If
			If DlgItem$ = "EditButton" Then
				If EditList(AllList,ConList) = True Then
					DlgListBoxArray "ConList",ConList()
					DlgValue "ConList",0
				End If
			ElseIf DlgItem$ = "AddAllButton" Then
				ConList = AllList
				DlgListBoxArray "ConList",ConList()
				DlgValue "ConList",0
			ElseIf DlgItem$ = "DelButton" Then
				n = DlgValue("ConList")
				i = UBound(ConList)
				ConList = DelArray(DlgText("ConList"),ConList,"",0)
				If n > 0 And n = i Then n = n - 1
				DlgListBoxArray "ConList",ConList()
				DlgValue "ConList",n
			ElseIf DlgItem$ = "CleanButton" Then
				ReDim ConList(0)
				DlgListBoxArray "ConList",ConList()
				DlgValue "ConList",0
			ElseIf DlgItem$ = "ResetButton" Then
				DataList = Split(ConDataListBak(ConTypeID),JoinStr)
				AllList = Split(DataList(1),SubJoinStr)
				ConList = Split(DataList(2),SubJoinStr)
				UseList = Split(DataList(3),SubJoinStr)
				DlgListBoxArray "ConList",ConList()
				DlgValue "ConList",0
			End If
			For i = 0 To 2
				If i = 0 Then TempArray = ConDataList
				If i = 1 Then TempArray = ConTypeDataList
				If i = 2 Then TempArray = ConListDataList
				DataList = Split(TempArray(ConTypeID),JoinStr)
				UseList = Split(DataList(3),SubJoinStr)
				For j = LBound(ConList) To UBound(ConList)
					ReDim Preserve TempDataList(j)
					Stemp = False
					For n = LBound(UseList) To UBound(UseList)
						oArray = Split(UseList(n),rSubJoinStr)
						If ConList(j) = oArray(0) Then
							Stemp = True
							Exit For
						End If
					Next n
					If Stemp = True Then
						TempDataList(j) = UseList(n)
					Else
						ConCmdID = DlgValue("ConCmdList")
						TextExpCharSet = DlgValue("ExpCodeList")
						WordFixSelect = DlgValue("WordFixCheckBox")
						AllTypeSame = DlgValue("AllTypeSameBox")
						AllListSame = DlgValue("AllListSameBox")
						TempDataList(j) = ConList(j) & rSubJoinStr & ConCmdID & rSubJoinStr & _
										TextExpCharSet & rSubJoinStr & WordFixSelect
					End If
				Next j
				DataList(2) = Join(ConList,SubJoinStr)
				DataList(3) = Join(TempDataList,SubJoinStr)
				TempArray(ConTypeID) = Join(DataList,JoinStr)
				If i = 0 Then ConDataList = TempArray
				If i = 1 Then ConTypeDataList = TempArray
				If i = 2 Then ConListDataList = TempArray
			Next i
			If DlgText("ConList") <> "" Then
				DataList = Split(ConDataList(ConTypeID),JoinStr)
				UseList = Split(DataList(3),SubJoinStr)
				TempArray = Split(UseList(DlgValue("ConList")),rSubJoinStr)
				DlgValue "ConCmdList",StrToInteger(TempArray(1))
				DlgValue "ExpCodeList",StrToInteger(TempArray(2))
				DlgValue "WordFixCheckBox",StrToInteger(TempArray(3))

				DlgText "ConListGroup",Msg07 & "(" & UBound(ConList) + 1 & ")"
				DlgText "EditButton",Msg09
				DlgEnable "DelButton",True
				DlgEnable "CleanButton",True
				If Join(ConList) = Join(AllList) Then
					DlgEnable "AddAllButton",False
				Else
					DlgEnable "AddAllButton",True
				End If
			Else
				DlgText "ConListGroup",Msg07 & "(0)"
				DlgText "EditButton",Msg08
				DlgEnable "DelButton",False
				DlgEnable "CleanButton",False
				DlgEnable "AddAllButton",True
			End If
			DataList = Split(ConDataListBak(ConTypeID),JoinStr)
			If DataList(2) = Join(ConList,SubJoinStr) Then
				DlgEnable "ResetButton",False
			Else
				DlgEnable "ResetButton",True
			End If
		End If

		If DlgItem$ = "SetButton" Then
			ConCmdListBak = ConCmdList
			ConCmdDataListBak = ConCmdDataList
			UpdateSetBak = UpdateSet
			ConCmdID = DlgValue("ConCmdList")
			Call ConCmdInput(ConCmdID,TextExpCharSet)
			DlgListBoxArray "ConCmdList",ConCmdList()
			DlgValue "ConCmdList",ConCmdID
			If DlgText("ConCmdList") = "" Then DlgValue "ConCmdList",0
		End If

		If DlgItem$ = "AddinChangButton" Then
			AddinID = AddinIDInput(DlgText("AddinIDText"))
			If AddinID = "" Then
				DlgText "AddinIDText",Msg02
				DlgText "AddinChangButton",Msg03
			Else
				DlgText "AddinIDText",AddinID
				DlgText "AddinChangButton",Msg04
			End If
		End If

		If DlgItem$ = "AllTypeSameBox" Or DlgItem$ = "AllListSameBox" Then
			Stemp = False
			If UBound(ConList) > 0 Then Stemp = True
			If Stemp = False Then
				For i = LBound(ConDataList) To UBound(ConDataList)
					DataList = Split(ConDataList(i),JoinStr)
					UseList = Split(DataList(3),SubJoinStr)
					If UBound(UseList) > 0 Then
						Stemp = True
						Exit For
					End If
				Next i
			End If
			If Stemp = True Then
				AllTypeSame = DlgValue("AllTypeSameBox")
				AllListSame = DlgValue("AllListSameBox")
				If DlgItem$ = "AllTypeSameBox" And AllTypeSame = 1 Then Stemp = False
				If DlgItem$ = "AllListSameBox" And AllListSame = 1 Then Stemp = False
				If Stemp = False Then
					If MsgBox(Msg11 & vbCrLf & Msg13,vbYesNo+vbInformation,Msg10) = vbNo Then
						If DlgItem$ = "AllTypeSameBox" Then DlgValue "AllTypeSameBox",0
						If DlgItem$ = "AllListSameBox" Then DlgValue "AllListSameBox",0
					End If
				Else
					If MsgBox(Msg12 & vbCrLf & Msg13,vbYesNo+vbInformation,Msg10) = vbNo Then
						If DlgItem$ = "AllTypeSameBox" Then DlgValue "AllTypeSameBox",1
						If DlgItem$ = "AllListSameBox" Then DlgValue "AllListSameBox",1
					End If
				End If
			End If
			If DlgValue("AllTypeSameBox") = 0 Then
				If DlgValue("AllListSameBox") = 0 Then ConDataList = ConTypeDataList
				If DlgValue("AllListSameBox") = 1 Then ConDataList = ConListDataList
			End If
			If DlgText("ConList") <> "" Then
				DataList = Split(ConDataList(ConTypeID),JoinStr)
				UseList = Split(DataList(3),SubJoinStr)
				TempArray = Split(UseList(DlgValue("ConList")),rSubJoinStr)
				DlgValue "ConCmdList",StrToInteger(TempArray(1))
				DlgValue "ExpCodeList",StrToInteger(TempArray(2))
				DlgValue "WordFixCheckBox",StrToInteger(TempArray(3))
			End If
		End If

		If DlgItem$ = "ConList" And DlgText("ConList") <> "" Then
			ConTypeID = DlgValue("ConTypeList")
			DataList = Split(ConDataList(ConTypeID),JoinStr)
			UseList = Split(DataList(3),SubJoinStr)
			TempArray = Split(UseList(DlgValue("ConList")),rSubJoinStr)
			DlgValue "ConCmdList",StrToInteger(TempArray(1))
			DlgValue "ExpCodeList",StrToInteger(TempArray(2))
			DlgValue "WordFixCheckBox",StrToInteger(TempArray(3))
		End If

		If DlgItem$ <> "CancelButton" Then
			If UBound(ConCmdList) = 0 Then
				DlgVisible "ConCmdNameText",True
				DlgVisible "ConCmdList",False
				DlgValue "ConCmdList",0
				If ConCmdList(0) <> "" Then	DlgText "ConCmdNameText",DlgText("ConCmdList")
			Else
				DlgVisible "ConCmdNameText",False
				DlgVisible "ConCmdList",True
				DlgText "ConCmdNameText",DlgText("ConCmdList")
			End If

			If DlgText("ConCmdList") <> "" Then
				ConCmdID = DlgValue("ConCmdList")
				TextExpCharSet = DlgValue("ExpCodeList")
				TempArray = Split(ConCmdDataList(ConCmdID),JoinStr)
				ConCmdPath = TempArray(1)
				If TextExpCharSet = 0 Then
					GBF = TempArray(4)
					BGF = TempArray(5)
				ElseIf TextExpCharSet = 1 Then
					GBF = TempArray(8)
					BGF = TempArray(9)
				ElseIf TextExpCharSet = 2 Then
					GBF = TempArray(12)
					BGF = TempArray(13)
				End If
				GBKFixPath = TempArray(14)
				Big5FixPath = TempArray(15)
			End If
			Stemp = False
			If GBF = "" Or BGF = "" Then Stemp = True
			If InStr(LCase(ConCmdPath),"concmd.exe") Or InStr(LCase(ConCmdPath),"convertz.exe") Then
				If GBKFixPath = "" Or Big5FixPath = "" Then Stemp = True
			End If
			If Stemp = True Then
				DlgValue "WordFixCheckBox",0
				DlgEnable "WordFixCheckBox",False
			Else
				DlgEnable "WordFixCheckBox",True
			End If

			CmdName = DlgText("ConCmdList")
			AddinID = DlgText("AddinIDText")
			TextExpCharSet = DlgValue("ExpCodeList")
			WordFixSelect = DlgValue("WordFixCheckBox")
			AllHandle = DlgValue("AllHandleBox")
			AllTypeSame = DlgValue("AllTypeSameBox")
			AllListSame = DlgValue("AllListSameBox")
			CycleSelect = DlgValue("CycleBox")
			KeepItemSelect = DlgValue("KeepSelet")
			nSelected = CmdName & JoinStr & AddinID & JoinStr & TextExpCharSet & JoinStr & _
						WordFixSelect & JoinStr & AllHandle & JoinStr & AllTypeSame & JoinStr & _
						AllListSame & JoinStr & CycleSelect & JoinStr & KeepItemSelect
			If Join(cSelected,JoinStr) = nSelected Then
				DlgEnable "SaveButton",False
			Else
				DlgEnable "SaveButton",True
			End If
			If AllTypeSame = 0 Then DlgEnable "AllListSameBox",True
			If AllTypeSame = 1 Then DlgEnable "AllListSameBox",False

			If CmdName = "" Or AddinID = Msg02 Or DlgText("ConList") = "" Then
				DlgEnable "OKButton",False
			Else
				If AllHandle = 0 Then
					If CheckNullData(CmdName,ConCmdDataList,"4,5,8,9,12,13,14,15,18,19",1) = True Then
						DlgEnable "OKButton",False
					Else
						DlgEnable "OKButton",True
					End If
				Else
					If CheckNullData("",ConCmdDataList,"4,5,8,9,12,13,14,15,18,19",1) = True Then
						DlgEnable "OKButton",False
					Else
						DlgEnable "OKButton",True
					End If
				End If
			End If
		End If

		If DlgItem$ = "SaveButton" Or DlgItem$ = "OKButton" Then
			If DlgItem$ = "SaveButton" Then KeepItemSelect = 1
			If KeepItemSelect = 1 Then
				cSelected = Split(nSelected,JoinStr)
				WriteSettings(ConCmdDataList,WriteLoc,"Main")
				ConCmdID = DlgValue("ConCmdList")
				TempArray = Split(ConCmdDataList(ConCmdID),JoinStr)
				ConCmdPath = TempArray(1)
				ConCmd = Left(ConCmdPath,InStrRev(ConCmdPath,"\"))
				SaveSetting("gb2big5","Settings","ConCmd",RemoveBackslash(ConCmd,"","\",1))
				DlgEnable "SaveButton",False
			End If
		End If

		If DlgItem$ <> "OKButton" And DlgItem$ <> "CancelButton" Then
			If DlgText("ConList") <> "" Then
				ConTypeID = DlgValue("ConTypeList")
				ConID = DlgValue("ConList")
				If DlgValue("AllTypeSameBox") = 0 Then
					For i = 0 To 2
						If i = 0 Then TempArray = ConDataList
						If i = 1 Then TempArray = ConTypeDataList
						If i = 2 Then TempArray = ConListDataList
						DataList = Split(TempArray(ConTypeID),JoinStr)
						UseList = Split(DataList(3),SubJoinStr)
						If DlgValue("AllListSameBox") = 0 And i <> 2 Then
							TempDataList = Split(UseList(ConID),rSubJoinStr)
							TempDataList(1) = DlgValue("ConCmdList")
							TempDataList(2) = DlgValue("ExpCodeList")
							TempDataList(3) = DlgValue("WordFixCheckBox")
							UseList(ConID) = Join(TempDataList,rSubJoinStr)
						ElseIf DlgValue("AllListSameBox") = 1 And i <> 1 Then
							For j = LBound(UseList) To UBound(UseList)
								TempDataList = Split(UseList(j),rSubJoinStr)
								TempDataList(1) = DlgValue("ConCmdList")
								TempDataList(2) = DlgValue("ExpCodeList")
								TempDataList(3) = DlgValue("WordFixCheckBox")
								UseList(j) = Join(TempDataList,rSubJoinStr)
							Next j
						End If
						DataList(3) = Join(UseList,SubJoinStr)
						DataList(4) = DlgValue("AllTypeSameBox")
						DataList(5) = DlgValue("AllListSameBox")
						TempArray(ConTypeID) = Join(DataList,JoinStr)
						If i = 0 Then ConDataList = TempArray
						If i = 1 Then ConTypeDataList = TempArray
						If i = 2 Then ConListDataList = TempArray
					Next i
				Else
					For i = LBound(ConDataList) To UBound(ConDataList)
						DataList = Split(ConDataList(i),JoinStr)
						UseList = Split(DataList(3),SubJoinStr)
						For n = LBound(UseList) To UBound(UseList)
							TempDataList = Split(UseList(n),rSubJoinStr)
							TempDataList(1) = DlgValue("ConCmdList")
							TempDataList(2) = DlgValue("ExpCodeList")
							TempDataList(3) = DlgValue("WordFixCheckBox")
							UseList(n) = Join(TempDataList,rSubJoinStr)
						Next n
						DataList(3) = Join(UseList,SubJoinStr)
						DataList(4) = DlgValue("AllTypeSameBox")
						ConDataList(i) = Join(DataList,JoinStr)
					Next i
				End If
			End If
			MainDlgFunc = True ' ��ֹ���°�ť�رնԻ��򴰿�
		End If
		If DlgItem$ = "AboutButton" Then Call Help("About")
	End Select
End Function


'����ѡ���ִ��б�����
Function getList(ConType As Integer,ID As Integer,AllList() As String,ConList() As String) As Boolean
	Dim i As Integer,j As Integer,n As Integer
	getList = False
	n = 0
	j = 0
	ReDim AllList(0),ConList(0)
	If ConType = 1 Or ConType = 3 Then
		If Prj.SourceLists.Count <> 0 Then
			For i = 1 To Prj.SourceLists.Count
				Set Src = Prj.SourceLists(i)
				If Src.LangID = ID Then
					If Src.Selected = True Then
						ReDim Preserve ConList(j)
						ConList(j) = Src.Title & " - " & PSL.GetLangCode(Src.LangID,pslCodeText)
						j = j + 1
					End If
					ReDim Preserve AllList(n)
					AllList(n) = Src.Title & " - " & PSL.GetLangCode(Src.LangID,pslCodeText)
					n = n + 1
				End If
			Next i
		End If
	End If
	If ConType = 0 Or ConType = 2 Then
		If Prj.TransLists.Count <> 0 Then
			For i = 1 To Prj.TransLists.Count
				Set trn = Prj.TransLists(i)
				If trn.Language.LangID = ID Then
					If trn.Selected = True Then
						ReDim Preserve ConList(j)
						ConList(j) = trn.Title & " - " & PSL.GetLangCode(trn.Language.LangID,pslCodeText)
						j = j + 1
					End If
					ReDim Preserve AllList(n)
					AllList(n) = trn.Title & " - " & PSL.GetLangCode(trn.Language.LangID,pslCodeText)
					n = n + 1
 				End If
			Next i
		End If
	End If
	If j > 0 Then ConList = ClearArray(ConList)
	If n > 0 Then AllList = ClearArray(AllList)
	If ConList(0) <> "" Then getList = True
End Function


'�༭ת���б�
Function EditList(AllList() As String,ConList() As String) As Boolean
	Dim AppStrList() As String
	If OSLanguage = "0404" Then
		Msg01 = "�s���ഫ�M��"
		Msg04 = "�s�W >"
		Msg05 = "�����s�W >>"
		Msg06 = "< �R��"
		Msg07 = "<< �����R��"
		Msg08 = "���]"
	Else
		Msg01 = "�༭ת���б�"
		Msg04 = "��� >"
		Msg05 = "ȫ����� >>"
		Msg06 = "< ɾ��"
		Msg07 = "<< ȫ��ɾ��"
		Msg08 = "����"
	End If
	EditList = False
	AllStrList = AllList
	ConStrList = ConList
	ConStrListBak = ConStrList
	AppStrList = ChangeList(AllStrList,ConStrList)
	Begin Dialog UserDialog 760,406,Msg01,.EditListFun ' %GRID:10,7,1,1
		Text 10,7,310,14,Msg02,.AppStrListText
		Text 440,7,310,14,Msg03,.ConStrListText
		ListBox 10,21,310,350,AppStrList(),.AppStrList
		ListBox 440,21,310,350,ConStrList(),.ConStrList
		PushButton 330,21,100,21,Msg04,.AddButton
		PushButton 330,42,100,21,Msg05,.AddAllButton
		PushButton 330,77,100,21,Msg06,.DelButton
		PushButton 330,98,100,21,Msg07,.CleanButton
		PushButton 330,133,100,21,Msg08,.ResetButton
		OKButton 250,378,90,21,.OKButton
		CancelButton 420,378,90,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	If Dialog(dlg) = 0 Then
		ConList = ConStrListBak
	Else
		ConList = ConStrList
		EditList = True
	End If
End Function


'����ز鿴�Ի�������������˽������Ϣ��
Private Function EditListFun(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Integer,ID As Integer,AppStrList() As String,Temp As String
	If OSLanguage = "0404" Then
		Msg02 = "�i�βM��:"
		Msg03 = "�ഫ�M��:"
	Else
		Msg02 = "�����б�:"
		Msg03 = "ת���б�:"
	End If
	Select Case Action%
	Case 1 ' �Ի��򴰿ڳ�ʼ��
		If DlgText("AppStrList") = "" Then
			DlgEnable "AddButton",False
			DlgEnable "AddAllButton",False
			DlgText "AppStrListText",Msg02 & "(0)"
		Else
			DlgEnable "AddButton",True
			DlgEnable "AddAllButton",True
			i = UBound(AllStrList) - UBound(ConStrList)
			DlgText "AppStrListText",Msg02 & "(" & i & ")"
		End If
		If DlgText("ConStrList") = "" Then
			DlgEnable "DelButton",False
			DlgEnable "CleanButton",False
			DlgText "ConStrListText",Msg03 & "(0)"
		Else
			DlgEnable "DelButton",True
			DlgEnable "CleanButton",True
			i = UBound(ConStrList) + 1
			DlgText "ConStrListText",Msg03 & "(" & i & ")"
		End If
		DlgEnable "ResetButton",False
	Case 2 ' ��ֵ���Ļ��߰��°�ťʱ
		AppStrList = ChangeList(AllStrList,ConStrList)
		If DlgItem$ = "AddButton" Or DlgItem$ = "DelButton" Then
			If DlgItem$ = "AddButton" Then
				Temp = DlgText("AppStrList")
				If Temp <> "" Then
					ID = DlgValue("AppStrList")
					i = UBound(AppStrList)
					AppStrList = DelArray(Temp,AppStrList,"",0)
					ConStrList = ChangeList(AllStrList,AppStrList)
				End If
			Else
				Temp = DlgText("ConStrList")
				If Temp <> "" Then
					ID = DlgValue("ConStrList")
					i = UBound(ConStrList)
					ConStrList = DelArray(Temp,ConStrList,"",0)
					AppStrList = ChangeList(AllStrList,ConStrList)
				End If
			End If
			If Temp <> "" Then
				If ID > 0 And ID = i Then ID = ID - 1
				DlgListBoxArray "AppStrList",AppStrList()
				DlgListBoxArray "ConStrList",ConStrList()
				If DlgItem$ = "AddButton" Then
					DlgValue "AppStrList",ID
					DlgText "ConStrList",Temp
				Else
					DlgText "AppStrList",Temp
					DlgValue "ConStrList",ID
				End If
			End If
		End If
		If DlgItem$ = "AddAllButton" Or DlgItem$ = "CleanButton" Or DlgItem$ = "ResetButton" Then
			If DlgItem$ = "AddAllButton" Then
				ReDim AppStrList(0)
				ConStrList = ChangeList(AllStrList,AppStrList)
			ElseIf DlgItem$ = "CleanButton" Then
				ReDim ConStrList(0)
				AppStrList = ChangeList(AllStrList,ConStrList)
			Else
				ConStrList = ConStrListBak
				AppStrList = ChangeList(AllStrList,ConStrList)
			End If
			DlgListBoxArray "AppStrList",AppStrList()
			DlgListBoxArray "ConStrList",ConStrList()
			DlgValue "AppStrList",0
			DlgValue "ConStrList",0
		End If
		If DlgText("AppStrList") = "" Then
			DlgEnable "AddButton",False
			DlgEnable "AddAllButton",False
			DlgText "AppStrListText",Msg02 & "(0)"
		Else
			DlgEnable "AddButton",True
			DlgEnable "AddAllButton",True
			i = UBound(AppStrList) + 1
			DlgText "AppStrListText",Msg02 & "(" & i & ")"
		End If
		If DlgText("ConStrList") = "" Then
			DlgEnable "DelButton",False
			DlgEnable "CleanButton",False
			DlgText "ConStrListText",Msg03 & "(0)"
		Else
			DlgEnable "DelButton",True
			DlgEnable "CleanButton",True
			i = UBound(ConStrList) + 1
			DlgText "ConStrListText",Msg03 & "(" & i & ")"
		End If
		If Join(ConStrList) <> Join(ConStrListBak) Then
			DlgEnable "ResetButton",True
		Else
			DlgEnable "ResetButton",False
		End If
		If DlgItem$ <> "OKButton" And DlgItem$ <> "CancelButton" Then
			EditListFun = True ' ��ֹ���°�ťʱ�رնԻ��򴰿�
		End If
	End Select
End Function


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
	If Join(UpdateSet) <> "" Then
		If Mode = "" Then Mode = UpdateSet(0)
		If Url = "" Then Url = UpdateSet(1)
		ExePath = UpdateSet(2)
		Argument = UpdateSet(3)
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
			If Mode = "1" Or Mode = "3" Then MsgBox(Msg08,vbOkOnly+vbInformation,Msg01)
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
		If Url <> Join(UrlList,vbCrLf) Then UpdateSet(1) = Join(UrlList,vbCrLf)
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


'�����������ļ�
Function RenameFixFile(Source As String,Target As String) As Boolean
	RenameFixFile = False
	If Source <> "" Then
		If FileRename(Source,Target) = True Then
			On Error GoTo ExitFunction
			Open Source For Output As #1
			If InStr(Source,"GBfix.dat") Or InStr(Source,"BI_SimFix.dat") Then
				Print #1,"-1,-1,-1,-1,-1,-1" & vbCrLf & "����ǰ,������"
			ElseIf InStr(Source,"B5fix.dat") Or InStr(Source,"BI_TradFix.dat") Then
				Print #1,"-1,-1,-1,-1,-1,-1" & vbCrLf & "�ץ��e,�ץ���"
			Else
				Print #1,""
			End If
			Close #1
			On Error GoTo 0
			RenameFixFile = True
		End If
	End If
	ExitFunction:
End Function


'��ȡת������
Function GetArgument(Data As String,FixFile As String) As String
	GetArgument = Data
	If GetArgument = "" Then Exit Function
	If InStr(GetArgument,"%1") Then
		GetArgument = Replace(GetArgument,"%1"," """ & prjFolder & InputFile & """")
	Else
		'GetArgument = GetArgument & " """ & prjFolder & InputFile & """"
	End If
	If InStr(GetArgument,"%2") Then
		GetArgument = Replace(GetArgument,"%2"," """ & prjFolder & OutputFile & """")
	Else
		'GetArgument = GetArgument & " """ & prjFolder & OutputFile & """"
	End If
	If InStr(GetArgument,"%3") Then GetArgument = Replace(GetArgument,"%3","""" & FixFile & """")
End Function


'����ת������
Function Config(ConID As Integer,ConList() As String,TextExpCharSet As Integer) As Integer
	Dim Srctrn As PslTransList,Trgtrn As PslTransList,Lan As PslLanguage,WshShell As Object
	Dim i As Integer,j As Integer,Argument As String,Temp As String,FixFileList() As String
	Dim FixID As Integer,ConCmdID As Integer,Code As String
	Dim SrcID As String,TrgID As String,Code_1 As String,Code_2 As String,Stemp As Boolean
	Dim ConArg As String,ConArgFix As String,FixPath As String,ReNameID As String

	If OSLanguage = "0404" Then
		Msg01 = "���~"
		Msg02 = "�L�k���s�R�W�U�C�ɮסA�нT�{�O�_�s�b�Υ��Q��L�{���ϥΡC" & vbCrLf
		Msg03 = "�L�k�٭�U�C�ɮסA�нT�{�ؼЦ�m�O�_���g�J�v���C" & vbCrLf
		Msg04 = "�L�k�ק��ഫ�{�����]�w�ɮסA�нT�{�ؼ��ɮ׬O�_���g�J�v���C" & vbCrLf
	Else
		Msg01 = "����"
		Msg02 = "�޷������������ļ�����ȷ���Ƿ���ڻ�������������ʹ�á�" & vbCrLf
		Msg03 = "�޷���ԭ�����ļ�����ȷ��Ŀ��λ���Ƿ���д��Ȩ�ޡ�" & vbCrLf
		Msg04 = "�޷��޸�ת������������ļ�����ȷ��Ŀ���ļ��Ƿ���д��Ȩ�ޡ�" & vbCrLf
	End If

	Config = 0
	Stemp = False
	Set WshShell = CreateObject("WScript.Shell")
	strKeyPath = "HKCU\Software\PASS Engineering\" & PSLVersion & "\TextExport\CharSet"

	For i = LBound(ConList) To UBound(ConList)
		TempArray = Split(ConList(i),rSubJoinStr)
		ConName = TempArray(0)
		ConCmdID = StrToInteger(TempArray(1))
		TextExpCharSet = StrToInteger(TempArray(2))
		FixID = StrToInteger(TempArray(3))

		If TextExpCharSet <> OldTextExpCharSet Then
			WshShell.RegWrite(strKeyPath,TextExpCharSet,"REG_DWORD")
		End If

		TempArray = Split(ConCmdDataList(ConCmdID),JoinStr)
		CmdPath = TempArray(1)
		If TextExpCharSet = 0 Then
			Code = "ANSI"
			GBKToBig5 = TempArray(2)
			Big5ToGBK = TempArray(3)
			GBKToBig5Fix = TempArray(4)
			Big5ToGBKFix = TempArray(5)
		ElseIf TextExpCharSet = 1 Then
			Code = "Unicode"
			GBKToBig5 = TempArray(6)
			Big5ToGBK = TempArray(7)
			GBKToBig5Fix = TempArray(8)
			Big5ToGBKFix = TempArray(9)
		ElseIf TextExpCharSet = 2 Then
			Code = "UTF-8"
			GBKToBig5 = TempArray(10)
			Big5ToGBK = TempArray(11)
			GBKToBig5Fix = TempArray(12)
			Big5ToGBKFix = TempArray(13)
		End If
		GBKFixPath = TempArray(14)
		Big5FixPath = TempArray(15)
		RenGBKFix = TempArray(16)
		RenBig5Fix = TempArray(17)
		FixFileSlipStr = TempArray(18)
		FixFileSaveCode = TempArray(19)

		If ConID = 0 Or ConID = 1 Then
			Code_1 = "1252"
			Code_2 = "950"
			Code_3 = 950
			Code_4 = 936
			If ConID = 0 Then LangID_1 = 2052
			If ConID = 1 Then LangID_1 = 65535
			LangID_2 = 1028
			LangName = "cht"
			ConArg = GBKToBig5
			ConArgFix = GBKToBig5Fix
			FixPath = Big5FixPath
			ReNameID = RenBig5Fix
			BuiltInFile = "bi_tradfix.dat"
		Else
			Code_1 = "1252"
			Code_2 = "936"
			Code_3 = 936
			Code_4 = 950
			If ConID = 2 Then LangID_1 = 1028
			If ConID = 3 Then LangID_1 = 65535
			LangID_2 = 2052
			LangName = "chs"
			ConArg = Big5ToGBK
			ConArgFix = Big5ToGBKFix
			FixPath = GBKFixPath
			ReNameID = RenGBKFix
			BuiltInFile = "bi_simfix.dat"
		End If

		Set Lan = Prj.Languages(LangName)
		If Lan Is Nothing Then
			Prj.Languages.Add(LangID_2)
			Prj.Languages(LangName).Option(pslOptionCodepage) = Code_3
		End If
		If ConID = 1 Or ConID = 3 Then
			Prj.Languages.Add(65535)
			Prj.Languages("3ff3f").Option(pslOptionCodepage) = Code_4
		End If

		FixFile = FixPath
		If FixFileSlipStr <> "" Then
			FixFile = Replace(FixPath,SubJoinStr,Convert(FixFileSlipStr))
			FixFilePath = prjFolder & "\~temp.txt"
			If FixFileSaveCode <> "" Then
				WriteToFile(FixFilePath,FixFile,FixFileSaveCode)
				FixFile = FixFilePath
			End If
		End If

		If FixID = 0 Then Argument = GetArgument(ConArg,FixFile)
		If FixID = 1 Then Argument = GetArgument(ConArgFix,FixFile)

		If FixPath <> "" Then
			FixFileList = Split(FixPath,SubJoinStr)
			If ReNameID = "1" Then
				For j = LBound(FixFileList) To UBound(FixFileList)
					FixFile = FixFileList(j)
					If FixID = 1 And Dir(FixFile & ".bak") <> "" Then
						If FileRename(FixFile & ".bak",FixFile) = False Then
							MsgBox Msg03 & FixFile,vbOkOnly+vbInformation,Msg01
							GoTo ExitFunction
						End If
					ElseIf FixID = 0 And Dir(FixFile) <> "" Then
						If RenameFixFile(FixFile,FixFile & ".bak") = False Then
							MsgBox Msg02 & FixFile,vbOkOnly+vbInformation,Msg01
							GoTo ExitFunction
						End If
					End If
				Next j
			End If
			If InStr(LCase(CmdPath),"convertz.exe") Then
				If Not (UBound(FixFileList) = 0 And LCase(FixFileList(0)) = BuiltInFile) Then
					INIFile = Left(CmdPath,InStrRev(LCase(CmdPath),".exe")) & "ini"
					If Dir(INIFile) <> "" Then
						If FileRename(INIFile,INIFile & ".bak") = False Then
							MsgBox Msg02 & INIFile,vbOkOnly+vbInformation,Msg01
							GoTo RevertFile
						End If
					End If
					If ChangeCMDSetings(INIFile,INIFile & ".bak",FixPath,ConID) = False Then
						MsgBox Msg04 & INIFile,vbOkOnly+vbInformation,Msg01
						GoTo RevertFile
					End If
				End If
			End If
		End If

		For j = 1 To Prj.SourceLists.Count
			Set Src = Prj.SourceLists(j)
			Set Srctrn = Prj.TransLists(Src,LangID_1)
			Set Trgtrn = Prj.TransLists(Src,LangID_2)
			If ConID = 0 Or ConID = 2 Then
				Temp = Srctrn.Title & " - " & PSL.GetLangCode(Srctrn.Language.LangID,pslCodeText)
			Else
				Temp = Src.Title & " - " & PSL.GetLangCode(Src.LangID,pslCodeText)
			End If
			If Temp = ConName Then
				SrcID = "@ID" & Str(Srctrn.ListID)
				TrgID = "@ID" & Str(Trgtrn.ListID)
				Argument = Code & JoinStr & Code_1 & JoinStr & Code_2 & JoinStr & SrcID & _
							JoinStr & TrgID & JoinStr & CmdPath & JoinStr & Argument
				If ConID = 1 Or ConID = 3 Then
					If Srctrn.SourceList.LastChange > Srctrn.LastUpdate Then Srctrn.Update
					If Srctrn.IsOpen Then Srctrn.Close(pslSaveChanges)
				End If
				If Trgtrn.SourceList.LastChange > Trgtrn.LastUpdate Then Trgtrn.Update
				If Trgtrn.IsOpen Then Trgtrn.Close(pslSaveChanges)
				If TrnConvert(Srctrn,Argument) = False Then GoTo ExitFunction
				Config = Config + 1
				Exit For
			End If
		Next j

		RevertFile:
		If FixPath <> "" Then
			If ReNameID = "1" And FixID = 0 Then
				For j = LBound(FixFileList) To UBound(FixFileList)
					FixFile = FixFileList(j)
					If Dir(FixFile & ".bak") <> "" Then
						If FileRename(FixFile & ".bak",FixFile) = False Then
							MsgBox Msg03 & FixFile,vbOkOnly+vbInformation,Msg01
						End If
					End If
				Next j
			End If
			If InStr(LCase(CmdPath),"convertz.exe") Then
				If Dir(INIFile & ".bak") <> "" Then
					If FileRename(INIFile & ".bak",INIFile) = False Then
						MsgBox Msg03 & INIFile,vbOkOnly+vbInformation,Msg01
					End If
				End If
			End If
		End If
		If FixFileSaveCode <> "" And FixFilePath <> "" Then
			On Error Resume Next
			If Dir(FixFilePath) <> "" Then Kill FixFilePath
			On Error GoTo 0
		End If

		ExitFunction:
		If ConID = 1 Or ConID = 3 Then
			Prj.Languages.Remove("3ff3f")
			PSL.OutputWnd.Clear
		End If
	Next i
	Set WshShell = Nothing
End Function


'����ת�����������
Function ChangeCMDSetings(OldPath As String,BakPath As String,FixPath As String,ConID As Integer) As Boolean
	Dim i As Integer,j As Integer,n As Integer,k As Integer,SectionName As String,FixFile As String
	Dim Code As String,FileLines() As String,TempArray() As String,Stemp As Boolean
	ChangeCMDSetings = False
	If ConID = 0 Or ConID = 1 Then
		SectionName = "Fix-Trad"
		BuiltInFile = "bi_tradfix.dat"
	ElseIf ConID = 2 Or ConID = 3 Then
		SectionName = "Fix-Sim"
		BuiltInFile = "bi_simfix.dat"
	End If
	FixFileList = Split(FixPath,SubJoinStr)
	n = 0
	Stemp = False
	If Dir(BakPath) <> "" Then
		Code = CheckCode(BakPath)
		Temp = ReadFile(BakPath,Code)
		FileLines = Split(Temp,vbCrLf,-1)
		k = 0
		For i = LBound(FileLines) To UBound(FileLines)
			L$ = FileLines(i)
			If Left(Trim(L$),1) = "[" And Right(Trim(L$),1) = "]" Then
				Header = Mid(Trim(L$),2,Len(Trim(L$))-2)
			End If
			setPreStr = ""
			setAppStr = ""
			If Header = SectionName Then
				If InStr(L$,"=") Then setPreStr = Trim(Left(L$,InStr(L$,"=")-1))
				If InStr(L$,"=") Then setAppStr = Trim(Mid(L$,InStr(L$,"=")+1))
				If InStr(setPreStr,"Age_") And setPreStr <> "Age_BuiltIn" Then L$ = "DelItem"
				If InStr(setPreStr,"Enable_") And setPreStr <> "Enable_BuiltIn" Then L$ = "DelItem"
				If InStr(setPreStr,"UserFile_") Then L$ = "DelItem"
				If setPreStr = "Enable_BuiltIn" Then
					For j = LBound(FixFileList) To UBound(FixFileList)
						FixFile = FixFileList(j)
						If LCase(FixFile) = BuiltInFile Then
							If setAppStr = "0" Then L$ = "Enable_BuiltIn=1"
						Else
							If setAppStr = "1" Then L$ = "Enable_BuiltIn=0"
							n = n + 1
							Temp = Mid(FixFile,InStrRev(FixFile,"\") + 1)
							UserFile = "UserFile_" & n & "=" & Temp
							Enable = "Enable_" & n & "=1"
							If LineStr <> "" Then LineStr = LineStr & vbCrLf & UserFile & vbCrLf & Enable
							If LineStr = "" Then LineStr = UserFile & vbCrLf & Enable
						End If
					Next j
					UserFile_Count = "UserFile_Count=" & n
					L$ = L$ & vbCrLf & UserFile_Count & vbCrLf & LineStr
					Stemp = True
				End If
				If L$ = "" And Stemp = False Then
					For j = LBound(FixFileList) To UBound(FixFileList)
						FixFile = FixFileList(j)
						If LCase(FixFile) = BuiltInFile Then
							If L$ = "" Then L$ = "Enable_BuiltIn=1"
						Else
							If L$ = "" Then L$ = "Enable_BuiltIn=0"
							n = n + 1
							FileName = Mid(FixFile,InStrRev(FixFile,"\") + 1)
							UserFile = "UserFile_" & n & "=" & FileName
							Enable = "Enable_" & n & "=1"
							If LineStr <> "" Then LineStr = LineStr & vbCrLf & UserFile & vbCrLf & Enable
							If LineStr = "" Then LineStr = UserFile & vbCrLf & Enable
						End If
					Next j
					UserFile_Count = "UserFile_Count=" & n
					L$ = L$ & vbCrLf & UserFile_Count & vbCrLf & LineStr & vbCrLf
					Stemp = True
				End If
			End If
			If L$ <> "DelItem" Then
				ReDim Preserve TempArray(k)
				TempArray(k) = L$
				k = k + 1
			End If
		Next i
		Temp = Join(TempArray,vbCrLf)
	Else
		Header = "[" & SectionName & "]"
		For j = LBound(FixFileList) To UBound(FixFileList)
			FixFile = FixFileList(j)
			If LCase(FixFile) = BuiltInFile Then
				If L$ = "" Then L$ = "Enable_BuiltIn=1"
			Else
				If L$ = "" Then L$ = "Enable_BuiltIn=0"
				n = n + 1
				FileName = Mid(FixFile,InStrRev(FixFile,"\") + 1)
				UserFile = "UserFile_" & n & "=" & FileName
				Enable = "Enable_" & n & "=1"
				If LineStr <> "" Then LineStr = LineStr & vbCrLf & UserFile & vbCrLf & Enable
				If LineStr = "" Then LineStr = UserFile & vbCrLf & Enable
			End If
		Next j
		UserFile_Count = "UserFile_Count=" & n
		Temp = Header & vbCrLf & L$ & vbCrLf & UserFile_Count & vbCrLf & LineStr
	End If
	If WriteToFile(OldPath,Temp,"utf-8") = True Then
		If Dir(OldPath) <> "" Then ChangeCMDSetings = True
	End If
End Function


'������ת������
Function TrnConvert(TrnList As PslTransList,Argument As String) As Boolean
	Dim objStream As Object,Code As String,FileLines() As String
	If OSLanguage = "0404" Then
		Msg01 = "���~"
		Msg02 = "���b�ഫ�A�еy��..."
		Msg03 = "�ഫ���ѡI�S����� Passolo �ץX�ɮסC"
		Msg04 = "�ഫ���ѡI�ഫ�{���ΩR�O�C�ѼƳ]�w�i�঳���D�C"
		Msg05 = "�ഫ���ѡI�ץX�ɮת��榡�i�ण�O ANSI �s�X�C"
		Msg06 = "�ഫ���ѡI�L�k�g�J�U�C�ɮסA�i��O�L�g�J�v���C" & vbCrLf
	Else
		Msg01 = "����"
		Msg02 = "����ת�������Ժ�..."
		Msg03 = "ת��ʧ�ܣ�û���ҵ� Passolo �����ļ���"
		Msg04 = "ת��ʧ�ܣ�ת������������в������ÿ��������⡣"
		Msg05 = "ת��ʧ�ܣ������ļ��ĸ�ʽ���ܲ��� ANSI ���롣"
		Msg06 = "ת��ʧ�ܣ��޷�д�������ļ�����������д��Ȩ�ޡ�" & vbCrLf
	End If

	TrnConvert = False
	TrnList.Export(AddinID, prjFolder & OutputFile, expAll)
	If Dir(prjFolder & OutputFile) = "" Then
		MsgBox Msg03,vbOkOnly+vbInformation,Msg01
		Exit Function
	End If

	TempArray = Split(Argument,JoinStr)
	Code = TempArray(0)
	Code_1 = TempArray(1)
	Code_2 = TempArray(2)
	SrcID = TempArray(3)
	TrgID = TempArray(4)
	ConCmdPath = TempArray(5)
	ConCmdArg = TempArray(6)

	PSL.Output(Msg02)
	Set objStream = CreateObject("Adodb.Stream")
	If objStream Is Nothing Or Code = "ANSI" Then
		On Error GoTo ErrorFormat
		Open prjFolder & OutputFile For Input As #1
   		Open prjFolder & InputFile For Output As #2
		While Not EOF(1)
   			Line Input #1, l$
			If InStr(l$,SrcID) = 1 Then l$ = TrgID
			If InStr(l$,"@CodePage1") = 1 Then l$ = "@CodePage1 " & Code_1 & ""
			If InStr(l$,"@CodePage2") = 1 Then l$ = "@CodePage2 " & Code_2 & ""
			Print #2, l$
		Wend
		Close #1
		Close #2
		On Error GoTo 0
	Else
		textStr = ReadFile(prjFolder & OutputFile,Code)
		FileLines = Split(textStr, vbCrLf, -1)
		n = 0
		For i = LBound(FileLines) To UBound(FileLines)
			l$ = FileLines(i)
			If InStr(l$,SrcID) = 1 Then
				l$ = TrgID
				n = n + 1
			ElseIf InStr(l$,"@CodePage1") = 1 Then
				l$ = "@CodePage1 " & Code_1 & ""
				n = n + 1
			ElseIf InStr(l$,"@CodePage2") = 1 Then
				l$ = "@CodePage2 " & Code_2 & ""
				n = n + 1
			End If
			FileLines(i) = l$
			If n > 3 Then Exit For
		Next i
		newTextStr = Join(FileLines,vbCrLf)
		If WriteToFile(prjFolder & InputFile,newTextStr,Code) = False Then
			MsgBox Msg06 & prjFolder & InputFile,vbOkOnly+vbInformation,Msg01
			GoTo DelFile
		End If
	End If

	On Error Resume Next
	If Dir(prjFolder & InputFile) <> "" Then Kill prjFolder & OutputFile
	On Error GoTo 0

	On Error GoTo ErrorConCmd
	Dim WshShell As Object
	Set WshShell = CreateObject("WScript.Shell")
	If WshShell Is Nothing Then
		ShellWait("""" & ConCmdPath & """ " & ConCmdArg)
	Else
		Return = WshShell.Run("""" & ConCmdPath & """ " & ConCmdArg, 0, True)
		Set WshShell = Nothing
		If Return <> 0 Then GoTo ErrorConCmd
	End If
	On Error GoTo 0

	On Error Resume Next
	If Not objStream Is Nothing And Code = "UTF-8" Then
		If Dir(prjFolder & OutputFile) <> "" Then
			Code = CheckCode(prjFolder & OutputFile)
			If LCase(Code) = "utf-8" Then
				textStr = ReadFile(prjFolder & OutputFile,Code)
				Kill prjFolder & OutputFile
				WriteToFile(prjFolder & OutputFile,textStr,Code)
			End If
		ElseIf Dir(prjFolder & InputFile) <> "" Then
			Code = CheckCode(prjFolder & InputFile)
			If LCase(Code) = "utf-8" Then
				textStr = ReadFile(prjFolder & InputFile, Code)
				Kill prjFolder & InputFile
				WriteToFile(prjFolder & InputFile,textStr,Code)
			End If
		End If
	End If
	Set objStream = Nothing
	On Error GoTo 0

	On Error Resume Next
	If InStr(ConCmdArg,OutputFile) <> 0 Then
		If Dir(prjFolder & InputFile) <> "" Then Kill prjFolder & InputFile
	End If
	On Error GoTo 0

	If Dir(prjFolder & OutputFile) <> "" Then
		PSL.Import(AddinID, prjFolder & OutputFile, pslImportValidate)
	ElseIf Dir(prjFolder & InputFile) <> "" Then
		PSL.Import(AddinID, prjFolder & InputFile, pslImportValidate)
	End If
	TrnConvert = True
	GoTo DelFile
	Exit Function

	ErrorConCmd:
	MsgBox Msg04,vbOkOnly+vbInformation,Msg01
	GoTo DelFile
	Exit Function

	ErrorFormat:
	Close #1
    Close #2
	MsgBox Msg05,vbOkOnly+vbInformation,Msg01

	DelFile:
	On Error Resume Next
	If Dir(prjFolder & InputFile) <> "" Then Kill prjFolder & InputFile
	If Dir(prjFolder & OutputFile) <> "" Then Kill prjFolder & OutputFile
	On Error GoTo 0
End Function


'���� AddinID ����
Function AddinIDTest(Data As String) As Boolean
	Dim dummyTrn As PslTransList,Lan As PslLanguage
	AddinIDTest = False
	If Data = "" Then Exit Function
	Set Lan = Prj.Languages(1)
	If Lan Is Nothing Then
		Prj.Languages.Add(65535)
		Prj.Languages("3ff3f").Option(pslOptionCodepage) = 936
		Set dummyTrn = Prj.TransLists(1)
		If dummyTrn.SourceList.LastChange > dummyTrn.LastUpdate Then dummyTrn.Update
	Else
		Set dummyTrn = Prj.TransLists(1)
	End If

	On Error Resume Next
	If Dir(prjFolder & OutputFile) <> "" Then Kill prjFolder & OutputFile
	On Error GoTo 0

	If dummyTrn.Export(Data, prjFolder & OutputFile, expAll) = True Then
		AddinIDTest = True
	End If

	Prj.Languages.Remove("3ff3f")
	PSL.OutputWnd.Clear
	On Error Resume Next
	If Dir(prjFolder & OutputFile) <> "" Then Kill prjFolder & OutputFile
	On Error GoTo 0
End Function


'���� AddinID ����
Function AddinIDInput(Data As String) As String
	If OSLanguage = "0404" Then
		Msg01 = "���~"
		Msg02 = "�x�s�W�١A�H�K�U������ɥi�H�۰ʩI�s"
		Msg03 = "��r�榡�ץX�W�q���W��"
		Msg04 = "����(&T)"
		Msg07 = "���(&O)"
		Msg08 = ">"
		Msg10 = "�п�J�uPassolo ��r�榡�v�ץX�W�q�����W�١C" & vbCrLf & vbCrLf & _
				"�ӶץX�W�q�����W�٥i��|�] Passolo �����λy���Ӥ��P�C" & vbCrLf & vbCrLf & _
				"�z�i�H:" & vbCrLf & _
				"�@�@- �ഫ��^�夶���U�A�� (������)�C" & vbCrLf & _
				"�@�@- �ˬd�ץX�r��M���ܤ���� Passolo ��r�榡" & vbCrLf & _
				"�@�@  ���W�١A�M��b�U�C��r�������J�ӦW�١C"
		Msg11 = "�z�� Passolo ��r�榡�ץX�W�q���W��:"
		AddinID0 = "Passolo ��r�榡"
		AddinID1 = "Passolo ��r�榡"
		AddinID2 = "Passolo text format"
	Else
		Msg01 = "����"
		Msg02 = "�������ƣ��Ա��´�����ʱ�����Զ�����"
		Msg03 = "�ı���ʽ�����������"
		Msg04 = "����(&T)"
		Msg07 = "ԭֵ(&O)"
		Msg08 = ">"
		Msg10 = "�����롰Passolo �ı���ʽ��������������ơ�" & vbCrLf & vbCrLf & _
				"�õ�����������ƿ��ܻ��� Passolo �汾�����Զ���ͬ��" & vbCrLf & vbCrLf & _
				"������:" & vbCrLf & _
				"����- ת����Ӣ�Ľ��������� (���Ƽ�)��" & vbCrLf & _
				"����- ��鵼���ִ��б�Ի����� Passolo �ı���ʽ" & vbCrLf & _
				"����  �����ƣ�Ȼ���������ı�������������ơ�"
		Msg11 = "���� Passolo �ı���ʽ�����������:"
		AddinID0 = "Passolo �ı���ʽ"
		AddinID1 = "Passolo ���ָ�ʽ"
		AddinID2 = "Passolo text format"
	End If

	AddinIDInput = Data
	ReDim AddinIDList(2)
	AddinIDList(0) = AddinID0
	AddinIDList(1) = AddinID1
	AddinIDList(2) = AddinID2

	Begin Dialog UserDialog 450,238,Msg03,.AddinIDInputFunc ' %GRID:10,7,1,1
		Text 20,14,410,112,Msg10
		Text 20,140,410,14,Msg11
		TextBox 20,154,380,21,.SetBox
		PushButton 400,154,30,21,Msg08,.SetsButton
		PushButton 20,210,90,21,Msg04,.TestButton
		PushButton 120,210,90,21,Msg07,.ResetButton
		OKButton 240,210,90,21,.OKButton
		CancelButton 340,210,90,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	dlg.SetBox = Data
	If Dialog(dlg) = 0 Then Exit Function
	AddinIDInput = dlg.SetBox
End Function


'����ز鿴�Ի�������������˽������Ϣ��
Private Function AddinIDInputFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim Data As String,i As Integer,Stemp As Boolean
	If OSLanguage = "0404" Then
		Msg01 = "���~"
		Msg02 = "�T��"
		Msg10 = "�z�S����J���󤺮e�I"
		Msg11 = "���զ��\�I"
		Msg12 = "���ե��ѡI"
		Msg13 = "�ӼW�q���W�ٴ��ե��ѡA�Э��s��J�C"
	Else
		Msg01 = "����"
		Msg02 = "��Ϣ"
		Msg10 = "��û�������κ����ݣ�"
		Msg11 = "���Գɹ���"
		Msg12 = "����ʧ�ܣ�"
		Msg13 = "�ò�����Ʋ���ʧ�ܣ����������롣"
	End If

	Select Case Action%
	Case 1 ' �Ի��򴰿ڳ�ʼ��
		If DlgText("SetBox") = "" Then DlgEnable "ResetButton",False
	Case 2 ' ��ֵ���Ļ��߰����˰�ť
		If DlgItem$ = "OKButton" Then
			Data = DlgText("SetBox")
			If Data = "" Then
				MsgBox Msg10,vbOkOnly+vbInformation,Msg01
				AddinIDInputFunc = True ' ��ֹ���°�ťʱ�رնԻ��򴰿�
				Exit Function
			End If
			If Data <> AddinID Then
				If AddinIDTest(Data) = True Then
					AddinID = Data
				Else
					MsgBox Msg13,vbOkOnly+vbInformation,Msg01
					AddinIDInputFunc = True ' ��ֹ���°�ťʱ�رնԻ��򴰿�
				End If
			End If
		End If
		If DlgItem$ = "TestButton" Then
			Data = DlgText("SetBox")
			If Data = "" Then
				MsgBox Msg10,vbOkOnly+vbInformation,Msg01
				AddinIDInputFunc = True ' ��ֹ���°�ťʱ�رնԻ��򴰿�
				Exit Function
			End If
			If AddinIDTest(Data) = True Then
				MsgBox Msg11,vbOkOnly+vbInformation,Msg02
			Else
				MsgBox Msg12,vbOkOnly+vbInformation,Msg02
			End If
		End If
		If DlgItem$ = "ResetButton" Then DlgText "SetBox",AddinID
		If DlgItem$ = "SetsButton" Then
	  		i = ShowPopupMenu(AddinIDList,vbPopupVCenterAlign)
			If i >= 0 Then DlgText "SetBox",AddinIDList(i)
		End If
		If DlgItem$ <> "OKButton" And DlgItem$ <> "CancelButton" Then
			AddinIDInputFunc = True ' ��ֹ���°�ťʱ�رնԻ��򴰿�
		End If
	Case 3 ' �ı��������Ͽ��ı�������
		Data = DlgText("SetBox")
		Stemp = False
		For i = LBound(AddinIDList) To UBound(AddinIDList)
			If AddinIDList(i) = Data Then
				Stemp = True
				Exit For
			End If
		Next i
		If Stemp = False And Data <> "" Then
			i = UBound(AddinIDList) + 1
			ReDim Preserve AddinIDList(i)
			AddinIDList(i) = Data
		End If
	End Select
End Function


'���� ConCmd ת�����������
Function ConCmdInput(ConCmdID,CodeID) As Integer
	Dim GBKFileList() As String,BIGFileList() As String
	If OSLanguage = "0404" Then
		Msg01 = "�ۭq�ഫ�{��"
		Msg02 = "�Ы��w�ഫ�{���ó]�w���n���ѼơA���յL�~��A�M�Ω��ھާ@�C" & vbCrLf & vbCrLf & _
				"�`�N: " & vbCrLf & _
				"�ѼƤ��ݭn��J�ο�X�ɮ׮ɡA���I���k�䪺���s��J�C" & _
				"��J�ɮ� (%1) �M��X�ɮ� (%2) ��쬰�t�ΰѼơA���i�ܧ�C" & _
				"���{���䴩�Ҧ��r���s�X����J�M��X�ɮסC"
		Msg03 = "�]�w�M��"
		Msg04 = "�s�W(&A)"
		Msg05 = "�ܧ�(&C)"
		Msg06 = "�R��(&D)"
		Msg07 = "�x�s����"
		Msg08 = "�ɮ�"
		Msg09 = "���U��"
		Msg10 = "�פJ�]�w"
		Msg11 = "�ץX�]�w"
		Msg12 = "�ഫ�{�����|:"
		Msg13 = "..."
		Msg14 = "ANSI"
		Msg15 = "Unicode"
		Msg16 = "UTF-8"
		Msg17 = "²���c�R�O�C�Ѽ�"
		Msg18 = "�c��²�R�O�C�Ѽ�"
		Msg19 = "���ץ����J��(���i��):"
		Msg20 = "�ץ����J��(�i��):"
		Msg21 = ">"
		Msg22 = "�x�s�]�w�A�H�K�U������ɥi�H�۰ʩI�s"
		Msg23 = "����(&H)"
		Msg24 = "����(&T)"
		Msg25 = "�M��(&C)"
		Msg26 = "�x�s(&S)"
		Msg27 = "����(&E)"
		Msg28 = "�D�{��"
		Msg29 = "�ץ���"
		Msg30 = "²��ץ����ɮײM��"
		Msg31 = "����ץ����ɮײM��"
		Msg33 = "������²�餣�ץ����J�ɡA���ݭn�ϥ�²��ץ����ɮ�(&B)"
		Msg34 = "²���ॿ�餣�ץ����J�ɡA���ݭn�ϥΥ���ץ����ɮ�(&G)"
		Msg35 = "�s��ץ���(&M)"
		Msg36 = "Ū��(&R)"
		Msg37 = "�s��(&E)"
		Msg38 = "���](&R)"
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
		Tools01 = "���m�{��(&E)"
		Tools02 = "�O�ƥ�(&N)"
		Tools03 = "�t�ιw�]�{��(&M)"
		Tools04 = "�ۭq�{��(&C)"
	Else
		Msg01 = "�Զ���ת������"
		Msg02 = "��ָ��ת���������ñ�Ҫ�Ĳ����������������Ӧ����ʵ�ʲ�����" & vbCrLf & vbCrLf & _
				"ע��: " & vbCrLf & _
				"��������Ҫ���������ļ�ʱ���뵥���ұߵİ�ť���롣" & _
				"�����ļ� (%1) ������ļ� (%2) �ֶ�Ϊϵͳ���������ɸ��ġ�" & _
				"������֧�������ַ���������������ļ���"
		Msg03 = "�����б�"
		Msg04 = "���(&A)"
		Msg05 = "����(&C)"
		Msg06 = "ɾ��(&D)"
		Msg07 = "��������"
		Msg08 = "�ļ�"
		Msg09 = "ע���"
		Msg10 = "��������"
		Msg11 = "��������"
		Msg12 = "ת������·��:"
		Msg13 = "..."
		Msg14 = "ANSI"
		Msg15 = "Unicode"
		Msg16 = "UTF-8"
		Msg17 = "��ת�������в���"
		Msg18 = "��ת�������в���"
		Msg19 = "�������ʻ�ʱ(���ɿ�):"
		Msg20 = "�����ʻ�ʱ(�ɿ�):"
		Msg21 = ">"
		Msg22 = "�������ã��Ա��´�����ʱ�����Զ�����"
		Msg23 = "����(&H)"
		Msg24 = "����(&T)"
		Msg25 = "���(&C)"
		Msg26 = "����(&S)"
		Msg27 = "�˳�(&E)"
		Msg28 = "������"
		Msg29 = "������"
		Msg30 = "�����������ļ��б�"
		Msg31 = "�����������ļ��б�"
		Msg33 = "����ת���岻�����ʻ�ʱ������Ҫʹ�ü����������ļ�(&B)"
		Msg34 = "����ת���岻�����ʻ�ʱ������Ҫʹ�÷����������ļ�(&G)"
		Msg35 = "�༭������(&M)"
		Msg36 = "��ȡ(&R)"
		Msg37 = "�༭(&E)"
		Msg38 = "����(&R)"
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
		Tools01 = "���ó���(&E)"
		Tools02 = "���±�(&N)"
		Tools03 = "ϵͳĬ�ϳ���(&M)"
		Tools04 = "�Զ������(&C)"
	End If

	ReDim AppNames(3),AppPaths(3),GBKFileList(0),GBKFileList(0)
	AppNames(0) = Tools01
   	AppNames(1) = Tools02
	AppNames(2) = Tools03
   	AppNames(3) = Tools04
	AppPaths(0) = ""
  	AppPaths(1) = "notepad.exe"
	AppPaths(2) = ""
  	AppPaths(3) = ""
	Begin Dialog UserDialog 620,448,Msg01,.ConCmdInputFunc ' %GRID:10,7,1,1
		Text 20,7,580,14,Msg02
		OptionGroup .Options
			OptionButton 130,28,130,14,Msg28,.MainApp
			OptionButton 270,28,130,14,Msg29,.FixFile
			OptionButton 410,28,130,14,Msg84,.AutoUpdate

		GroupBox 20,49,330,70,Msg03,.ConComGroup
		DropListBox 40,66,290,21,ConCmdList(),.ConCmdName
		PushButton 40,91,90,21,Msg04,.AddButton
		PushButton 140,91,90,21,Msg05,.ChangButton
		PushButton 240,91,90,21,Msg06,.DelButton

		GroupBox 370,49,230,70,Msg07,.SaveTypeGroup
		OptionGroup .WriteType
			OptionButton 390,69,90,21,Msg08,.SaveToFile
			OptionButton 490,69,90,21,Msg09,.SaveToRegistry
		PushButton 390,91,90,21,Msg10,.ImportButton
		PushButton 490,91,90,21,Msg11,.ExportButton

		Text 20,133,550,14,Msg12,.ConCmdPathText
		TextBox 20,147,550,21,.ConCmdPath
		PushButton 570,147,30,21,Msg13,.BrowseButton

		OptionGroup .Encoding
			OptionButton 130,189,130,14,Msg14,.ANSIType
			OptionButton 270,189,130,14,Msg15,.UTF8Type
			OptionButton 410,189,130,14,Msg16,.UnicodeType

		GroupBox 20,217,580,91,Msg17,.GBKToBigGroup
		Text 50,241,170,14,Msg19,.GBKToBig5Text
		Text 50,276,170,14,Msg20,.GBKToBig5FixText
		TextBox 230,238,320,21,.GBKToBig5ANSI
		TextBox 230,238,320,21,.GBKToBig5Unicode
		TextBox 230,238,320,21,.GBKToBig5UTF8
		TextBox 230,273,320,21,.GBKToBig5FixANSI
		TextBox 230,273,320,21,.GBKToBig5FixUnicode
		TextBox 230,273,320,21,.GBKToBig5FixUTF8
		PushButton 550,238,30,21,Msg21,.GBBtn
		PushButton 550,273,30,21,Msg21,.GBFBtn

		GroupBox 20,315,580,91,Msg18,.Big5ToGBKGroup
		Text 50,339,170,14,Msg19,.Big5ToGBKText
		Text 50,374,170,14,Msg20,.Big5ToGBKFixText
		TextBox 230,336,320,21,.Big5ToGBKANSI
		TextBox 230,336,320,21,.Big5ToGBKUnicode
		TextBox 230,336,320,21,.Big5ToGBKUTF8
		TextBox 230,371,320,21,.Big5ToGBKFixANSI
		TextBox 230,371,320,21,.Big5ToGBKFixUnicode
		TextBox 230,371,320,21,.Big5ToGBKFixUTF8
		PushButton 550,336,30,21,Msg21,.BGBtn
		PushButton 550,371,30,21,Msg21,.BGFBtn

		GroupBox 20,133,580,133,Msg30,.GBKGroup
		ListBox 30,147,460,91,GBKFileList(),.GBKFileList,1
		PushButton 500,147,90,21,Msg04,.GBKAddButton
		PushButton 500,168,90,21,Msg05,.GBKChangeButton
		PushButton 500,189,90,21,Msg06,.GBKDelButton
		PushButton 500,210,90,21,Msg37,.GBKEditButton
		PushButton 500,231,90,21,Msg38,.GBKResetButton
		CheckBox 30,245,460,14,Msg33,.GBKRename

		GroupBox 20,273,580,133,Msg31,.BIGGroup
		ListBox 30,287,460,91,BIGFileList(),.BIGFileList,1
		PushButton 500,287,90,21,Msg04,.BIGAddButton
		PushButton 500,308,90,21,Msg05,.BIGChangeButton
		PushButton 500,329,90,21,Msg06,.BIGDelButton
		PushButton 500,350,90,21,Msg37,.BIGEditButton
		PushButton 500,371,90,21,Msg38,.BIGResetButton
		CheckBox 30,385,460,14,Msg34,.BIGRename

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
		GroupBox 20,154,580,112,Msg93,.WebSiteGroup
		TextBox 40,175,540,77,.WebSiteBox,1
		GroupBox 20,280,580,126,Msg94,.CmdGroup
		Text 40,301,510,14,Msg95,.CmdPathBoxText
		Text 40,350,510,14,Msg96,.ArgumentBoxText
		TextBox 40,322,510,21,.CmdPathBox
		TextBox 40,371,510,21,.ArgumentBox
		PushButton 550,322,30,21,Msg13,.ExeBrowseButton
		PushButton 550,371,30,21,Msg21,.ArgumentButton

		PushButton 120,420,90,21,Msg25,.CleanButton
		PushButton 20,420,90,21,Msg36,.ResetButton
		PushButton 20,420,90,21,Msg36,.FixResetButton
		PushButton 220,420,90,21,Msg24,.TestButton
		PushButton 120,420,90,21,Msg24,.FixTestButton
		PushButton 220,420,140,21,Msg35,.EditButton
		OKButton 410,420,90,21,.OKButton
		CancelButton 510,420,90,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	dlg.ConCmdName = ConCmdID
	dlg.Encoding = CodeID
	If Dialog(dlg) = 0 Then Exit Function
	ConCmdInput = dlg.ConCmdName
End Function


'���� ConCMD ת������Ի�����
Private Function ConCmdInputFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Integer,x As Integer,n As Integer,ConCmdID As Integer,CodeID As Integer
	Dim CmdName As String,CmdPath As String,Stemp As Boolean,Temp As String,Path As String
	Dim SettingArr() As String,GBKFileList() As String,BIGFileList() As String
	Dim TempArray() As String,FixFileSeparator As String,FixFileSaveCode As String

	If OSLanguage = "0404" Then
		Msg01 = "���~"
		Msg02 = "����ഫ�{��"
		Msg03 = "�T��"
		Msg04 = "�i�����ɮ� (*.exe)|*.exe|�Ҧ��ɮ� (*.*)|*.*||"
		Msg05 = "�w�]��"
		Msg06 = "���"
		Msg07 = "�ѷӭ�"
		Msg08 = "����"
		Msg11 = "�x�s(&S)"
		Msg12 = "����(&E)"
		Msg13 = "..."
		Msg14 = ">"
		Msg15 = "�]�w���e�w�g�ܧ���O�S���x�s�I�O�_�ݭn�x�s�H"
		Msg18 = "�x�s�����w�g�ܧ���O�S���x�s�I�O�_�ݭn�x�s�H"
		Msg21 = "�T�{"
		Msg22 = "�T��n�R���]�w�u%s�v�ܡH"
		Msg23 = "�T��n�R���ץ��ɮסu%s�v�ܡH"
		Msg24 = "��ܪ��{���w�b�]�w���s�b�I�O�_���ݷs�W�H"
		Msg25 = "�������ا����šI�Э��s��J�C"
		Msg26 = "�ثe�]�w���A���|�M���ץ����J�ɪ��ѼƦܤ֦��@�����šI���ˬd�ÿ�J�C"
		Msg27 = "�Ҧ��]�w���A���|�M���ץ����J�ɪ��ѼƦܤ֦��@�����šI���ˬd�ÿ�J�C"
		Msg28 = "��ܪ��ɮפw�b�M�椤�s�b�I�Э��s����C"
		Msg29 = "�ݭn�N�ץ��ɮ׽ƻs���ഫ�{���Ҧb��Ƨ��ܡH" & vbCrLf & _
				"�Юھ��ഫ�{�����n�D�i��T�{�I"
		Msg30 = "��ܪ��ɮפw�b�ഫ�{���Ҧb��Ƨ����s�b�I�O�_�٭n�ƻs�H"
		Msg32 = "�ץX�]�w���\�I"
		Msg33 = "�פJ�]�w���\�I"
		Msg36 = "�L�k�x�s�]�w�I���ˬd�O�_���g�J�U�C��m���v��:" & vbCrLf & vbCrLf
		Msg39 = "�פJ�]�w���ѡI���ˬd�O�_���g�J�U�C��m���v��" & vbCrLf & _
				"�ζפJ�ɮת��榡�O�_���T:" & vbCrLf & vbCrLf
		Msg40 = "�ץX�]�w���ѡI���ˬd�O�_���g�J�U�C��m���v��" & vbCrLf & _
				"�ζץX�ɮת��榡�O�_���T:" & vbCrLf & vbCrLf
		Msg42 = "����n�פJ���ɮ�"
		Msg43 = "����n�ץX���ɮ�"
		Msg44 = "�]�w�ɮ� (*.dat)|*.dat|�Ҧ��ɮ� (*.*)|*.*||"
		Msg45 = "����ץ��ɮ�"
		Msg46 = "�ץ��ɮ� (*.dat)|*.dat|�Ҧ��ɮ� (*.*)|*.*||"
		Msg60 = "��������{��"
		Msg61 = "�i�����ɮ� (*.exe)|*.exe|�Ҧ��ɮ� (*.*)|*.*||"
		Msg62 = "�S�����w�����{���I�Э��s��J�ο���C"
		Msg63 = "�ɮװѷӰѼ�(%1)"
		Msg64 = "�n�^�����ɮװѼ�(%2)"
		Msg65 = "�������|�Ѽ�(%3)"
		InFile = "��J�ɮ�(%1)"
		OutFile = "��X�ɮ�(%2)"
		FixFile = "�ץ��ɮ�(%3)"
		AddFile = "�s��(&G)"
		EditFile = "�s��(&E)"
	Else
		Msg01 = "����"
		Msg02 = "ѡ��ת������"
		Msg03 = "��Ϣ"
		Msg04 = "��ִ���ļ� (*.exe)|*.exe|�����ļ� (*.*)|*.*||"
		Msg05 = "Ĭ��ֵ"
		Msg06 = "ԭֵ"
		Msg07 = "����ֵ"
		Msg08 = "δ֪"
		Msg11 = "����(&S)"
		Msg12 = "�˳�(&E)"
		Msg13 = "..."
		Msg14 = ">"
		Msg15 = "���������Ѿ����ĵ���û�б��棡�Ƿ���Ҫ���棿"
		Msg18 = "���������Ѿ����ĵ���û�б��棡�Ƿ���Ҫ���棿"
		Msg21 = "ȷ��"
		Msg22 = "ȷʵҪɾ�����á�%s����"
		Msg23 = "ȷʵҪɾ�������ļ���%s����"
		Msg24 = "ѡ���ĳ������������д��ڣ��Ƿ�������ӣ�"
		Msg25 = "ȫ����Ŀ��Ϊ�գ����������롣"
		Msg26 = "��ǰ�����У�·���Ͳ������ʻ�ʱ�Ĳ���������һ��Ϊ�գ����鲢���롣"
		Msg27 = "���������У�·���Ͳ������ʻ�ʱ�Ĳ���������һ��Ϊ�գ����鲢���롣"
		Msg28 = "ѡ�����ļ������б��д��ڣ�������ѡ��"
		Msg29 = "��Ҫ�������ļ����Ƶ�ת�����������ļ�����" & vbCrLf & _
				"�����ת�������Ҫ�����ȷ�ϣ�"
		Msg30 = "ѡ�����ļ�����ת�����������ļ����д��ڣ��Ƿ�Ҫ���ƣ�"
		Msg32 = "�������óɹ���"
		Msg33 = "�������óɹ���"
		Msg36 = "�޷��������ã������Ƿ���д������λ�õ�Ȩ��:" & vbCrLf & vbCrLf
		Msg39 = "��������ʧ�ܣ������Ƿ���д������λ�õ�Ȩ��" & vbCrLf & _
				"�����ļ��ĸ�ʽ�Ƿ���ȷ:" & vbCrLf & vbCrLf
		Msg40 = "��������ʧ�ܣ������Ƿ���д������λ�õ�Ȩ��" & vbCrLf & _
				"�򵼳��ļ��ĸ�ʽ�Ƿ���ȷ:" & vbCrLf & vbCrLf
		Msg42 = "ѡ��Ҫ������ļ�"
		Msg43 = "ѡ��Ҫ�������ļ�"
		Msg44 = "�����ļ� (*.dat)|*.dat|�����ļ� (*.*)|*.*||"
		Msg45 = "ѡ�������ļ�"
		Msg46 = "�����ļ� (*.dat)|*.dat|�����ļ� (*.*)|*.*||"
		Msg60 = "ѡ���ѹ����"
		Msg61 = "��ִ���ļ� (*.exe)|*.exe|�����ļ� (*.*)|*.*||"
		Msg62 = "û��ָ����ѹ���������������ѡ��"
		Msg63 = "�ļ����ò���(%1)"
		Msg64 = "Ҫ��ȡ���ļ�����(%2)"
		Msg65 = "��ѹ·������(%3)"
		InFile = "�����ļ�(%1)"
		OutFile = "����ļ�(%2)"
		FixFile = "�����ļ�(%3)"
		AddFile = "���(&G)"
		EditFile = "�༭(&E)"
	End If

	If DlgValue("Options") = 0 Then
		DlgVisible "ConComGroup",True
		DlgVisible "ConCmdName",True
		DlgVisible "AddButton",True
		DlgVisible "ChangButton",True
		DlgVisible "DelButton",True
		DlgVisible "SaveTypeGroup",True
		DlgVisible "WriteType",True
		DlgVisible "ImportButton",True
		DlgVisible "ExportButton",True

		DlgVisible "ConCmdPathText",True
		DlgVisible "ConCmdPath",True
		DlgVisible "BrowseButton",True

		DlgVisible "ANSIType",True
		DlgVisible "UnicodeType",True
		DlgVisible "UTF8Type",True
		DlgVisible "GBBtn",True
		DlgVisible "GBFBtn",True
		DlgVisible "BGBtn",True
		DlgVisible "BGFBtn",True
		DlgVisible "GBKToBigGroup",True
		DlgVisible "Big5ToGBKGroup",True
		DlgVisible "GBKToBig5Text",True
		DlgVisible "GBKToBig5FixText",True
		DlgVisible "Big5ToGBKText",True
		DlgVisible "Big5ToGBKFixText",True
		DlgVisible "ResetButton",True
		DlgVisible "CleanButton",True
		DlgVisible "TestButton",True
		If DlgValue("Encoding") = 0 Then
			DlgVisible "GBKToBig5ANSI",True
			DlgVisible "Big5ToGBKANSI",True
			DlgVisible "GBKToBig5FixANSI",True
			DlgVisible "Big5ToGBKFixANSI",True
			DlgVisible "GBKToBig5Unicode",False
			DlgVisible "Big5ToGBKUnicode",False
			DlgVisible "GBKToBig5FixUnicode",False
			DlgVisible "Big5ToGBKFixUnicode",False
			DlgVisible "GBKToBig5UTF8",False
			DlgVisible "Big5ToGBKUTF8",False
			DlgVisible "GBKToBig5FixUTF8",False
			DlgVisible "Big5ToGBKFixUTF8",False
			DlgVisible "TestButton",True
		ElseIf DlgValue("Encoding") = 1 Then
			DlgVisible "GBKToBig5ANSI",False
			DlgVisible "Big5ToGBKANSI",False
			DlgVisible "GBKToBig5FixANSI",False
			DlgVisible "Big5ToGBKFixANSI",False
			DlgVisible "GBKToBig5Unicode",True
			DlgVisible "Big5ToGBKUnicode",True
			DlgVisible "GBKToBig5FixUnicode",True
			DlgVisible "Big5ToGBKFixUnicode",True
			DlgVisible "GBKToBig5UTF8",False
			DlgVisible "Big5ToGBKUTF8",False
			DlgVisible "GBKToBig5FixUTF8",False
			DlgVisible "Big5ToGBKFixUTF8",False
		ElseIf DlgValue("Encoding") = 2 Then
			DlgVisible "GBKToBig5ANSI",False
			DlgVisible "Big5ToGBKANSI",False
			DlgVisible "GBKToBig5FixANSI",False
			DlgVisible "Big5ToGBKFixANSI",False
			DlgVisible "GBKToBig5Unicode",False
			DlgVisible "Big5ToGBKUnicode",False
			DlgVisible "GBKToBig5FixUnicode",False
			DlgVisible "Big5ToGBKFixUnicode",False
			DlgVisible "GBKToBig5UTF8",True
			DlgVisible "Big5ToGBKUTF8",True
			DlgVisible "GBKToBig5FixUTF8",True
			DlgVisible "Big5ToGBKFixUTF8",True
		End If
		DlgVisible "GBKGroup",False
		DlgVisible "GBKFileList",False
		DlgVisible "GBKAddButton",False
		DlgVisible "GBKChangeButton",False
		DlgVisible "GBKDelButton",False
		DlgVisible "GBKEditButton",False
		DlgVisible "GBKResetButton",False

		DlgVisible "GBKRename",False
		DlgVisible "BIGGroup",False
		DlgVisible "BIGFileList",False
		DlgVisible "BIGAddButton",False
		DlgVisible "BIGChangeButton",False
		DlgVisible "BIGDelButton",False
		DlgVisible "BIGEditButton",False
		DlgVisible "BIGResetButton",False
		DlgVisible "BIGRename",False

		DlgVisible "FixResetButton",False
		DlgVisible "FixTestButton",False
		DlgVisible "EditButton",False

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
		DlgVisible "ConComGroup",True
		DlgVisible "ConCmdName",True
		DlgVisible "AddButton",True
		DlgVisible "ChangButton",True
		DlgVisible "DelButton",True
		DlgVisible "SaveTypeGroup",True
		DlgVisible "WriteType",True
		DlgVisible "ImportButton",True
		DlgVisible "ExportButton",True

		DlgVisible "ConCmdPathText",False
		DlgVisible "ConCmdPath",False
		DlgVisible "BrowseButton",False

		DlgVisible "ANSIType",False
		DlgVisible "UnicodeType",False
		DlgVisible "UTF8Type",False
		DlgVisible "GBBtn",False
		DlgVisible "GBFBtn",False
		DlgVisible "BGBtn",False
		DlgVisible "BGFBtn",False
		DlgVisible "GBKToBigGroup",False
		DlgVisible "Big5ToGBKGroup",False
		DlgVisible "GBKToBig5Text",False
		DlgVisible "GBKToBig5FixText",False
		DlgVisible "Big5ToGBKText",False
		DlgVisible "Big5ToGBKFixText",False
		DlgVisible "ResetButton",False
		DlgVisible "CleanButton",False
		DlgVisible "TestButton",False

		DlgVisible "GBKToBig5ANSI",False
		DlgVisible "Big5ToGBKANSI",False
		DlgVisible "GBKToBig5FixANSI",False
		DlgVisible "Big5ToGBKFixANSI",False
		DlgVisible "GBKToBig5Unicode",False
		DlgVisible "Big5ToGBKUnicode",False
		DlgVisible "GBKToBig5FixUnicode",False
		DlgVisible "Big5ToGBKFixUnicode",False
		DlgVisible "GBKToBig5UTF8",False
		DlgVisible "Big5ToGBKUTF8",False
		DlgVisible "GBKToBig5FixUTF8",False
		DlgVisible "Big5ToGBKFixUTF8",False

		DlgVisible "GBKGroup",True
		DlgVisible "GBKFileList",True
		DlgVisible "GBKAddButton",True
		DlgVisible "GBKChangeButton",True
		DlgVisible "GBKDelButton",True
		DlgVisible "GBKEditButton",True
		DlgVisible "GBKResetButton",True
		DlgVisible "GBKRename",True

		DlgVisible "BIGGroup",True
		DlgVisible "BIGFileList",True
		DlgVisible "BIGAddButton",True
		DlgVisible "BIGChangeButton",True
		DlgVisible "BIGDelButton",True
		DlgVisible "BIGEditButton",True
		DlgVisible "BIGResetButton",True
		DlgVisible "BIGRename",True

		DlgVisible "FixResetButton",True
		DlgVisible "FixTestButton",True
		DlgVisible "EditButton",True

		DlgVisible "UpdateSetGroup",False
		DlgVisible "UpdateSet",False
		DlgVisible "AutoButton",False
		DlgVisible "ManualButton",False
		DlgVisible "OffButton",False
		DlgVisible "CheckGroup",False
		DlgVisible "UpdateCycleText",False
		DlgVisible "UpdateCycleBox",False
		DlgVisible "UpdateDatesText",False
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
		DlgVisible "ConComGroup",False
		DlgVisible "ConCmdName",False
		DlgVisible "AddButton",False
		DlgVisible "ChangButton",False
		DlgVisible "DelButton",False
		DlgVisible "SaveTypeGroup",False
		DlgVisible "WriteType",False
		DlgVisible "ImportButton",False
		DlgVisible "ExportButton",False

		DlgVisible "ConCmdPathText",False
		DlgVisible "ConCmdPath",False
		DlgVisible "BrowseButton",False

		DlgVisible "ANSIType",False
		DlgVisible "UnicodeType",False
		DlgVisible "UTF8Type",False
		DlgVisible "GBBtn",False
		DlgVisible "GBFBtn",False
		DlgVisible "BGBtn",False
		DlgVisible "BGFBtn",False
		DlgVisible "GBKToBigGroup",False
		DlgVisible "Big5ToGBKGroup",False
		DlgVisible "GBKToBig5Text",False
		DlgVisible "GBKToBig5FixText",False
		DlgVisible "Big5ToGBKText",False
		DlgVisible "Big5ToGBKFixText",False
		DlgVisible "ResetButton",True
		DlgVisible "CleanButton",True
		DlgVisible "TestButton",True

		DlgVisible "GBKToBig5ANSI",False
		DlgVisible "Big5ToGBKANSI",False
		DlgVisible "GBKToBig5FixANSI",False
		DlgVisible "Big5ToGBKFixANSI",False
		DlgVisible "GBKToBig5Unicode",False
		DlgVisible "Big5ToGBKUnicode",False
		DlgVisible "GBKToBig5FixUnicode",False
		DlgVisible "Big5ToGBKFixUnicode",False
		DlgVisible "GBKToBig5UTF8",False
		DlgVisible "Big5ToGBKUTF8",False
		DlgVisible "GBKToBig5FixUTF8",False
		DlgVisible "Big5ToGBKFixUTF8",False

		DlgVisible "GBKGroup",False
		DlgVisible "GBKFileList",False
		DlgVisible "GBKAddButton",False
		DlgVisible "GBKChangeButton",False
		DlgVisible "GBKDelButton",False
		DlgVisible "GBKEditButton",False
		DlgVisible "GBKResetButton",False
		DlgVisible "GBKRename",False

		DlgVisible "BIGGroup",False
		DlgVisible "BIGFileList",False
		DlgVisible "BIGAddButton",False
		DlgVisible "BIGChangeButton",False
		DlgVisible "BIGDelButton",False
		DlgVisible "BIGEditButton",False
		DlgVisible "BIGResetButton",False
		DlgVisible "BIGRename",False

		DlgVisible "FixResetButton",False
		DlgVisible "FixTestButton",False
		DlgVisible "EditButton",False

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
		If Join(ConCmdList) <> "" Then
			ConCmdID = DlgValue("ConCmdName")
			SettingArr = Split(ConCmdDataList(ConCmdID),JoinStr)
			GBKFileList = Split(SettingArr(14),SubJoinStr)
			BIGFileList = Split(SettingArr(15),SubJoinStr)
			DlgText "ConCmdPath",SettingArr(1)
			DlgText "GBKToBig5ANSI",SettingArr(2)
			DlgText "Big5ToGBKANSI",SettingArr(3)
			DlgText "GBKToBig5FixANSI",SettingArr(4)
			DlgText "Big5ToGBKFixANSI",SettingArr(5)
			DlgText "GBKToBig5Unicode",SettingArr(6)
			DlgText "Big5ToGBKUnicode",SettingArr(7)
			DlgText "GBKToBig5FixUnicode",SettingArr(8)
			DlgText "Big5ToGBKFixUnicode",SettingArr(9)
			DlgText "GBKToBig5UTF8",SettingArr(10)
			DlgText "Big5ToGBKUTF8",SettingArr(11)
			DlgText "GBKToBig5FixUTF8",SettingArr(12)
			DlgText "Big5ToGBKFixUTF8",SettingArr(13)
			DlgListBoxArray "GBKFileList",GBKFileList
			DlgListBoxArray "BIGFileList",BIGFileList
			DlgValue "GBKRename",StrToInteger(SettingArr(16))
			DlgValue "BIGRename",StrToInteger(SettingArr(17))
			DlgValue "GBKFileList",0
			DlgValue "BIGFileList",0
		End If

		If WriteLoc = FilePath Then DlgValue "WriteType",0
		If WriteLoc = RegKey Then DlgValue "WriteType",1
		If WriteLoc = "" Then DlgValue "WriteType",0

		If Join(UpdateSet) <> "" Then
			DlgValue "UpdateSet",StrToInteger(UpdateSet(0))
			DlgText "WebSiteBox",UpdateSet(1)
			DlgText "CmdPathBox",UpdateSet(2)
			DlgText "ArgumentBox",UpdateSet(3)
			DlgText "UpdateCycleBox",UpdateSet(4)
			DlgText "UpdateDateBox",UpdateSet(5)
		End If
		If DlgText("UpdateDateBox") = "" Then DlgText "UpdateDateBox",Msg08
		DlgEnable "UpdateDateBox",False

		If DlgText("GBKFileList") = "" Then
			DlgEnable "GBKDelButton",False
			DlgEnable "GBKChangeButton",False
			DlgEnable "GBKEditButton",False
		End If
		If DlgText("BIGFileList") = "" Then
			DlgEnable "BIGDelButton",False
			DlgEnable "BIGChangeButton",False
			DlgEnable "BIGEditButton",False
		End If
		If DlgText("ConCmdName") = "" Then
			DlgEnable "ResetButton",False
			DlgEnable "FixResetButton",False
		End If
		If DlgText("GBKFileList") & DlgText("BIGFileList") = "" Then DlgEnable "EditButton",False
		If UBound(ConCmdList) = 0 Then DlgEnable "DelButton",False
	Case 2 ' ��ֵ���Ļ��߰����˰�ť
		ConCmdID = DlgValue("ConCmdName")
		If ConCmdID < 0 Then ConCmdID = 0
		SettingArr = Split(ConCmdDataList(ConCmdID),JoinStr)
		GBKFileList = Split(SettingArr(14),SubJoinStr)
		BIGFileList = Split(SettingArr(15),SubJoinStr)
		FixFileSeparator = SettingArr(18)
		FixFileSaveCode = SettingArr(19)

		If DlgValue("Options") = 0 Or DlgValue("Options") = 1 Then
			If DlgItem$ = "ConCmdName" Then
				DlgText "ConCmdPath",SettingArr(1)
				DlgText "GBKToBig5ANSI",SettingArr(2)
				DlgText "Big5ToGBKANSI",SettingArr(3)
				DlgText "GBKToBig5FixANSI",SettingArr(4)
				DlgText "Big5ToGBKFixANSI",SettingArr(5)
				DlgText "GBKToBig5Unicode",SettingArr(6)
				DlgText "Big5ToGBKUnicode",SettingArr(7)
				DlgText "GBKToBig5FixUnicode",SettingArr(8)
				DlgText "Big5ToGBKFixUnicode",SettingArr(9)
				DlgText "GBKToBig5UTF8",SettingArr(10)
				DlgText "Big5ToGBKUTF8",SettingArr(11)
				DlgText "GBKToBig5FixUTF8",SettingArr(12)
				DlgText "Big5ToGBKFixUTF8",SettingArr(13)
				DlgListBoxArray "GBKFileList",GBKFileList
				DlgListBoxArray "BIGFileList",BIGFileList
				DlgValue "GBKRename",StrToInteger(SettingArr(16))
				DlgValue "BIGRename",StrToInteger(SettingArr(17))
				DlgValue "GBKFileList",0
				DlgValue "BIGFileList",0
			End If

			If DlgItem$ = "AddButton" Or DlgItem$ = "BrowseButton" Then
				If PSL.SelectFile(CmdPath,True,Msg04,Msg02) = True Then
					Stemp = False
					For i = LBound(ConCmdDataList) To UBound(ConCmdDataList)
						TempArray = Split(ConCmdDataList(i),JoinStr)
						If LCase(TempArray(1)) = LCase(CmdPath) Then
							Stemp = True
							Exit For
						End If
					Next i
					If Stemp = True Then
						If MsgBox(Msg24,vbYesNo+vbInformation,Msg21) = vbNo Then
							ConCmdInputFunc = True '��ֹ���°�ť�رնԻ��򴰿�
							Exit Function
						End If
					End If
					Temp = Mid(CmdPath,InStrRev(CmdPath,"\")+1)
					CmdName = Left(Temp,InStrRev(Temp,".")-1)
					For i = LBound(ConCmdList) To UBound(ConCmdList)
						If InStr(LCase(ConCmdList(i)),LCase(CmdName)) Then x = x + 1
					Next i
					If x > 0 Then
						If InStr(LCase(Join(ConCmdList,JoinStr)),LCase(CmdName & "_" & x)) Then
							CmdName = CmdName & "_" & x +1
						Else
							CmdName = CmdName & "_" & x
						End If
					End If
					If DlgText("ConCmdName") <> "" Then
						If DlgItem$ = "BrowseButton" Then CmdName = DlgText("ConCmdName")
					End If
					If Stemp = True Then
						TempArray(0) = CmdName
						TempArray(1) = CmdPath
						ConCmdData = Join(TempArray,JoinStr)
					Else
						If InStr(LCase(CmdPath),"concmd.exe") Then Stemp = True
						If InStr(LCase(CmdPath),"convertz.exe") Then Stemp = True
						If Stemp = True Then
							Path = Left(CmdPath,InStrRev(CmdPath,"\"))
							DefaultData = DefaultSetting(CmdName,CmdPath)
							If InStr(LCase(CmdPath),"concmd.exe") Then
								DefaultData = Replace(DefaultData,"GBfix.dat",Path & "GBfix.dat")
								DefaultData = Replace(DefaultData,"B5fix.dat",Path & "B5fix.dat")
							Else
								DefaultData = Replace(DefaultData,"BI_SimFix.dat",Path & "BI_SimFix.dat")
								DefaultData = Replace(DefaultData,"BI_TradFix.dat",Path & "BI_TradFix.dat")
							End If
							ConCmdData = CmdName & JoinStr & CmdPath & JoinStr & DefaultData
						Else
							ConCmdData = CmdName & JoinStr & CmdPath & JoinStr & JoinStr & JoinStr & _
									JoinStr & JoinStr & JoinStr & JoinStr &	JoinStr & JoinStr & JoinStr & _
									JoinStr & JoinStr & JoinStr & JoinStr & JoinStr & "0" & JoinStr & "0" & _
									JoinStr & JoinStr
						End If
					End If
					CreateArray(CmdName,ConCmdData,ConCmdList,ConCmdDataList)
					DlgListBoxArray "ConCmdName",ConCmdList()
					DlgText "ConCmdName",CmdName
					ConCmdID = DlgValue("ConCmdName")
					If ConCmdID < 0 Then ConCmdID = 0
					SettingArr = Split(ConCmdDataList(ConCmdID),JoinStr)
					GBKFileList = Split(SettingArr(14),SubJoinStr)
					BIGFileList = Split(SettingArr(15),SubJoinStr)
					DlgText "ConCmdPath",SettingArr(1)
					DlgText "GBKToBig5ANSI",SettingArr(2)
					DlgText "Big5ToGBKANSI",SettingArr(3)
					DlgText "GBKToBig5FixANSI",SettingArr(4)
					DlgText "Big5ToGBKFixANSI",SettingArr(5)
					DlgText "GBKToBig5Unicode",SettingArr(6)
					DlgText "Big5ToGBKUnicode",SettingArr(7)
					DlgText "GBKToBig5FixUnicode",SettingArr(8)
					DlgText "Big5ToGBKFixUnicode",SettingArr(9)
					DlgText "GBKToBig5UTF8",SettingArr(10)
					DlgText "Big5ToGBKUTF8",SettingArr(11)
					DlgText "GBKToBig5FixUTF8",SettingArr(12)
					DlgText "Big5ToGBKFixUTF8",SettingArr(13)
					DlgListBoxArray "GBKFileList",GBKFileList
					DlgListBoxArray "BIGFileList",BIGFileList
					DlgValue "GBKRename",StrToInteger(SettingArr(16))
					DlgValue "BIGRename",StrToInteger(SettingArr(17))
					DlgValue "GBKFileList",0
					DlgValue "BIGFileList",0
				End If
			End If

			If DlgItem$ = "ChangButton" Then
				CmdName = DlgText("ConCmdName")
				NewCmdName = EditSet(ConCmdList,CmdName)
				If NewCmdName <> "" Then
					ConCmdList(ConCmdID) = NewCmdName
					SettingArr = Split(ConCmdDataList(ConCmdID),JoinStr)
					SettingArr(0) = NewCmdName
					ConCmdDataList(ConCmdID) = Join(SettingArr,JoinStr)
					DlgListBoxArray "ConCmdName",ConCmdList()
					DlgValue "ConCmdName",ConCmdID
				End If
			End If

	    	If DlgItem$ = "DelButton" Then
	    		CmdName = DlgText("ConCmdName")
	    		Msg = Replace(Msg22,"%s",CmdName)
				If MsgBox(Msg,vbYesNo+vbInformation,Msg21) = vbYes Then
					ConCmdList = DelArray(CmdName,ConCmdList,"",0)
					ConCmdDataList = DelArray(CmdName,ConCmdDataList,JoinStr,0)
					DlgListBoxArray "ConCmdName",ConCmdList()
					DlgValue "ConCmdName",0
					SettingArr = Split(ConCmdDataList(0),JoinStr)
					GBKFileList = Split(SettingArr(14),SubJoinStr)
					BIGFileList = Split(SettingArr(15),SubJoinStr)
					DlgText "ConCmdPath",SettingArr(1)
					DlgText "GBKToBig5ANSI",SettingArr(2)
					DlgText "Big5ToGBKANSI",SettingArr(3)
					DlgText "GBKToBig5FixANSI",SettingArr(4)
					DlgText "Big5ToGBKFixANSI",SettingArr(5)
					DlgText "GBKToBig5Unicode",SettingArr(6)
					DlgText "Big5ToGBKUnicode",SettingArr(7)
					DlgText "GBKToBig5FixUnicode",SettingArr(8)
					DlgText "Big5ToGBKFixUnicode",SettingArr(9)
					DlgText "GBKToBig5UTF8",SettingArr(10)
					DlgText "Big5ToGBKUTF8",SettingArr(11)
					DlgText "GBKToBig5FixUTF8",SettingArr(12)
					DlgText "Big5ToGBKFixUTF8",SettingArr(13)
					DlgListBoxArray "GBKFileList",GBKFileList
					DlgListBoxArray "BIGFileList",BIGFileList
					DlgValue "GBKRename",StrToInteger(SettingArr(16))
					DlgValue "BIGRename",StrToInteger(SettingArr(17))
					DlgValue "GBKFileList",0
					DlgValue "BIGFileList",0
				End If
			End If

			If DlgItem$ = "GBBtn" Or DlgItem$ = "BGBtn" Or DlgItem$ = "GBFBtn" Or DlgItem$ = "BGFBtn" Then
				CmdPath = DlgText("ConCmdPath")
				If InStr(LCase(CmdPath),"concmd.exe") Or InStr(LCase(CmdPath),"convertz.exe") Then
					ReDim TempArray(1)
					TempArray(0) = InFile
					TempArray(1) = OutFile
				Else
					ReDim TempArray(2)
					TempArray(0) = InFile
					TempArray(1) = OutFile
					TempArray(2) = FixFile
				End If
				x = ShowPopupMenu(TempArray,vbPopupVCenterAlign)
				If x = 2 Then
					If FixFileListSet(FixFileSeparator,FixFileSaveCode) = False Then
						ConCmdInputFunc = True '��ֹ���°�ť�رնԻ��򴰿�
						Exit Function
					End If
				End If
				If x >= 0 Then
					GBa = DlgText("GBKToBig5ANSI")
					BGa = DlgText("Big5ToGBKANSI")
					GBFa = DlgText("GBKToBig5FixANSI")
					BGFa = DlgText("Big5ToGBKFixANSI")
					GBu = DlgText("GBKToBig5Unicode")
					BGu = DlgText("Big5ToGBKUnicode")
					GBFu = DlgText("GBKToBig5FixUnicode")
					BGFu = DlgText("Big5ToGBKFixUnicode")
					GB8 = DlgText("GBKToBig5UTF8")
					BG8 = DlgText("Big5ToGBKUTF8")
					GBF8 = DlgText("GBKToBig5FixUTF8")
					BGF8 = DlgText("Big5ToGBKFixUTF8")
					Temp = TempArray(x)
					Temp = Mid(Temp,InStr(Temp,"%"),2)
					If DlgValue("Encoding") = 0 Then
						If DlgItem$ = "GBBtn" Then DlgText "GBKToBig5ANSI",GBa & " " & Temp
						If DlgItem$ = "BGBtn" Then DlgText "Big5ToGBKANSI",BGa & " " & Temp
						If DlgItem$ = "GBFBtn" Then DlgText "GBKToBig5FixANSI",GBFa & " " & Temp
						If DlgItem$ = "BGFBtn" Then DlgText "Big5ToGBKFixANSI",BGFa & " " & Temp
					ElseIf DlgValue("Encoding") = 1 Then
						If DlgItem$ = "GBBtn" Then DlgText "GBKToBig5Unicode",GBu & " " & Temp
						If DlgItem$ = "BGBtn" Then DlgText "Big5ToGBKUnicode",BGu & " " & Temp
						If DlgItem$ = "GBFBtn" Then DlgText "GBKToBig5FixUnicode",GBFu & " " & Temp
						If DlgItem$ = "BGFBtn" Then DlgText "Big5ToGBKFixUnicode",BGFu & " " & Temp
					ElseIf DlgValue("Encoding") = 2 Then
						If DlgItem$ = "GBBtn" Then DlgText "GBKToBig5UTF8",GB8 & " " & Temp
						If DlgItem$ = "BGBtn" Then DlgText "Big5ToGBKUTF8",BG8 & " " & Temp
						If DlgItem$ = "GBFBtn" Then DlgText "GBKToBig5FixUTF8",GBF8 & " " & Temp
						If DlgItem$ = "BGFBtn" Then DlgText "Big5ToGBKFixUTF8",BGF8 & " " & Temp
					End If
				End If
			End If

			If DlgItem$ = "CleanButton" Then
				If DlgValue("Encoding") = 0 Then
					DlgText "GBKToBig5ANSI",""
					DlgText "Big5ToGBKANSI",""
					DlgText "GBKToBig5FixANSI",""
					DlgText "Big5ToGBKFixANSI",""
				ElseIf DlgValue("Encoding") = 1 Then
					DlgText "GBKToBig5Unicode",""
					DlgText "Big5ToGBKUnicode",""
					DlgText "GBKToBig5FixUnicode",""
					DlgText "Big5ToGBKFixUnicode",""
				ElseIf DlgValue("Encoding") = 2 Then
					DlgText "GBKToBig5UTF8",""
					DlgText "Big5ToGBKUTF8",""
					DlgText "GBKToBig5FixUTF8",""
					DlgText "Big5ToGBKFixUTF8",""
				End If
			End If

			If DlgItem$ = "ResetButton" Or DlgItem$ = "FixResetButton" Then
				CmdName = DlgText("ConCmdName")
				CmdPath = DlgText("ConCmdPath")
				ReDim TempArray(1)
				If InStr(LCase(CmdPath),"concmd.exe") Or InStr(LCase(CmdPath),"convertz.exe") Then
					TempArray(0) = Msg05
				End If
				Stemp = CheckNullData(CmdName,ConCmdDataListBak,"4,5,8,9,12,13,14,15,18,19",0)
				If Stemp = False Then TempArray(1) = Msg06
				For i = LBound(ConCmdList) To UBound(ConCmdList)
					If i <> ConCmdID Then
						ReDim Preserve TempArray(i+2)
						TempArray(i+2) = Msg07 & " - " & ConCmdList(i)
					End If
				Next i
				i = ShowPopupMenu(TempArray,vbPopupUseRightButton)
				If i = 0 Then
					DefaultData = DefaultSetting(CmdName,CmdPath)
					SettingArr = Split(DefaultData,JoinStr)
				ElseIf i = 1 Then
					SettingArr = Split(ConCmdDataListBak(ConCmdID),JoinStr)
				ElseIf i >= 2 Then
					For n = LBound(ConCmdList) To UBound(ConCmdList)
						Temp = Mid(TempArray(i),InStr(TempArray(i),Msg07 & " - ") + Len(Msg07 & " - "))
						If Temp = ConCmdList(n) Then
							ConCmdID = n
							Exit For
						End If
					Next n
					SettingArr = Split(ConCmdDataList(ConCmdID),JoinStr)
				End If
				If i = 0 Then
					Temp = Left(CmdPath,InStrRev(CmdPath,"\"))
					GBKFileList = Split(Temp & SettingArr(12),SubJoinStr)
					BIGFileList = Split(Temp & SettingArr(13),SubJoinStr)
					If DlgValue("Options") = 0 Then
						If DlgValue("Encoding") = 0 Then
							DlgText "GBKToBig5ANSI",SettingArr(0)
							DlgText "Big5ToGBKANSI",SettingArr(1)
							DlgText "GBKToBig5FixANSI",SettingArr(2)
							DlgText "Big5ToGBKFixANSI",SettingArr(3)
						ElseIf DlgValue("Encoding") = 1 Then
							DlgText "GBKToBig5Unicode",SettingArr(4)
							DlgText "Big5ToGBKUnicode",SettingArr(5)
							DlgText "GBKToBig5FixUnicode",SettingArr(6)
							DlgText "Big5ToGBKFixUnicode",SettingArr(7)
						ElseIf DlgValue("Encoding") = 2 Then
							DlgText "GBKToBig5UTF8",SettingArr(8)
							DlgText "Big5ToGBKUTF8",SettingArr(9)
							DlgText "GBKToBig5FixUTF8",SettingArr(10)
							DlgText "Big5ToGBKFixUTF8",SettingArr(11)
						End If
					ElseIf DlgValue("Options") = 1 Then
						DlgListBoxArray "GBKFileList",GBKFileList
						DlgListBoxArray "BIGFileList",BIGFileList
						DlgValue "GBKRename",StrToInteger(SettingArr(14))
						DlgValue "BIGRename",StrToInteger(SettingArr(15))
						DlgValue "GBKFileList",0
						DlgValue "BIGFileList",0
					End If
				ElseIf i > 0 Then
					GBKFileList = Split(SettingArr(14),SubJoinStr)
					BIGFileList = Split(SettingArr(15),SubJoinStr)
					If DlgValue("Options") = 0 Then
						If DlgValue("Encoding") = 0 Then
							DlgText "GBKToBig5ANSI",SettingArr(2)
							DlgText "Big5ToGBKANSI",SettingArr(3)
							DlgText "GBKToBig5FixANSI",SettingArr(4)
							DlgText "Big5ToGBKFixANSI",SettingArr(5)
						ElseIf DlgValue("Encoding") = 1 Then
							DlgText "GBKToBig5Unicode",SettingArr(6)
							DlgText "Big5ToGBKUnicode",SettingArr(7)
							DlgText "GBKToBig5FixUnicode",SettingArr(8)
							DlgText "Big5ToGBKFixUnicode",SettingArr(9)
						ElseIf DlgValue("Encoding") = 2 Then
							DlgText "GBKToBig5UTF8",SettingArr(10)
							DlgText "Big5ToGBKUTF8",SettingArr(11)
							DlgText "GBKToBig5FixUTF8",SettingArr(12)
							DlgText "Big5ToGBKFixUTF8",SettingArr(13)
						End If
					ElseIf DlgValue("Options") = 1 Then
						DlgListBoxArray "GBKFileList",GBKFileList
						DlgListBoxArray "BIGFileList",BIGFileList
						DlgValue "GBKRename",StrToInteger(SettingArr(16))
						DlgValue "BIGRename",StrToInteger(SettingArr(17))
						DlgValue "GBKFileList",0
						DlgValue "BIGFileList",0
					End If
				End If
 			End If

			If DlgItem$ = "GBKAddButton" Or DlgItem$ = "BIGAddButton" Then
				If PSL.SelectFile(Path,True,Msg46,Msg45) = True Then
					CmdPath = DlgText("ConCmdPath")
					If CmdPath <> "" Then
						NewPath = Left(CmdPath,InStrRev(CmdPath,"\")) & Mid(Path,InStrRev(Path,"\")+1)
					Else
						NewPath = Path
					End If
					Stemp = False
					If LCase(Path) <> LCase(NewPath) Then
						If InStr(LCase(CmdPath),"concmd.exe") Then Stemp = True
						If InStr(LCase(CmdPath),"convertz.exe") Then Stemp = True
						If Stemp = False Then
							Msg = MsgBox(Msg29,vbYesNoCancel+vbInformation,Msg21)
							If Msg = vbYes Then
								Stemp = True
							ElseIf Msg = vbNo Then
								NewPath = Path
							Else
								ConCmdInputFunc = True '��ֹ���°�ť�رնԻ��򴰿�
								Exit Function
							End If
						End If
					End If
					If DlgItem$ = "GBKAddButton" Then Temp = Join(GBKFileList,SubJoinStr)
					If DlgItem$ = "BIGAddButton" Then Temp = Join(BIGFileList,SubJoinStr)
					If InStr(LCase(Temp),LCase(NewPath)) Then
						MsgBox(Msg28,vbOkOnly+vbInformation,Msg01)
						ConCmdInputFunc = True '��ֹ���°�ť�رնԻ��򴰿�
						Exit Function
					End If
					If LCase(Path) <> LCase(NewPath) And Stemp = True Then
						If Dir(NewPath) = "" Then
							FileCopy Path,NewPath
						Else
							Msg = MsgBox(Msg30,vbYesNoCancel+vbInformation,Msg21)
							If Msg = vbYes Then
								SetAttr NewPath,vbNormal
								FileCopy Path,NewPath
							ElseIf Msg = vbCancel Then
								ConCmdInputFunc = True '��ֹ���°�ť�رնԻ��򴰿�
								Exit Function
							End If
						End If
					End If
					If DlgItem$ = "GBKAddButton" Then
						If GBKFileList(0) = "" Then i = UBound(GBKFileList)
						If GBKFileList(0) <> "" Then i = UBound(GBKFileList) + 1
						ReDim Preserve GBKFileList(i)
						GBKFileList(i) = NewPath
						DlgListBoxArray "GBKFileList",GBKFileList
						DlgValue "GBKFileList",i
					Else
						If BIGFileList(0) = "" Then i = UBound(BIGFileList)
						If BIGFileList(0) <> "" Then i = UBound(BIGFileList) + 1
						ReDim Preserve BIGFileList(i)
						BIGFileList(i) = NewPath
						DlgListBoxArray "BIGFileList",BIGFileList
						DlgValue "BIGFileList",i
					End If
				End If
			End If

			If DlgItem$ = "GBKDelButton" Or DlgItem$ = "BIGDelButton" Then
				If DlgItem$ = "GBKDelButton" Then Temp = DlgText("GBKFileList")
				If DlgItem$ = "BIGDelButton" Then Temp = DlgText("BIGFileList")
				If Temp <> "" Then
					Msg = Replace(Msg23,"%s",Temp)
					If MsgBox(Msg,vbYesNo+vbInformation,Msg21) = vbYes Then
						If DlgItem$ = "GBKDelButton" Then
							i = DlgValue("GBKFileList")
							If i > 0 And i = UBound(GBKFileList) Then i = i - 1
							GBKFileList = DelArray(Temp,GBKFileList,"",0)
							DlgListBoxArray "GBKFileList",GBKFileList
							DlgValue "GBKFileList",i
						Else
							i = DlgValue("BIGFileList")
							If i > 0 And i = UBound(BIGFileList) Then i = i - 1
							BIGFileList = DelArray(Temp,BIGFileList,"",0)
							DlgListBoxArray "BIGFileList",BIGFileList
							DlgValue "BIGFileList",i
						End If
					End If
				End If
			End If

			If DlgItem$ = "GBKChangeButton" Or DlgItem$ = "BIGChangeButton" Then
				If DlgItem$ = "GBKChangeButton" Then Temp = DlgText("GBKFileList")
				If DlgItem$ = "BIGChangeButton" Then Temp = DlgText("BIGFileList")
				If Temp <> "" Then
					If PSL.SelectFile(Path,True,Msg46,Msg45) = True Then
						CmdPath = DlgText("ConCmdPath")
						If CmdPath <> "" Then
							NewPath = Left(CmdPath,InStrRev(CmdPath,"\")) & Mid(Path,InStrRev(Path,"\")+1)
						Else
							NewPath = Path
						End If
						Stemp = False
						If LCase(Path) <> LCase(NewPath) Then
							If InStr(LCase(CmdPath),"concmd.exe") Then Stemp = True
							If InStr(LCase(CmdPath),"convertz.exe") Then Stemp = True
							If Stemp = False Then
								Msg = MsgBox(Msg29,vbYesNoCancel+vbInformation,Msg21)
								If Msg = vbYes Then
									Stemp = True
								ElseIf Msg = vbNo Then
									NewPath = Path
								Else
									ConCmdInputFunc = True '��ֹ���°�ť�رնԻ��򴰿�
									Exit Function
								End If
							End If
						End If
						If DlgItem$ = "GBKChangeButton" Then Temp = Join(GBKFileList,SubJoinStr)
						If DlgItem$ = "BIGChangeButton" Then Temp = Join(BIGFileList,SubJoinStr)
						If InStr(LCase(Temp),LCase(NewPath)) Then
							MsgBox(Msg28,vbOkOnly+vbInformation,Msg21)
							ConCmdInputFunc = True '��ֹ���°�ť�رնԻ��򴰿�
							Exit Function
						End If
						If LCase(Path) <> LCase(NewPath) And Stemp = True Then
						If Dir(NewPath) = "" Then
							FileCopy Path,NewPath
						Else
							Msg = MsgBox(Msg30,vbYesNoCancel+vbInformation,Msg21)
							If Msg = vbYes Then
								SetAttr NewPath,vbNormal
								FileCopy Path,NewPath
							ElseIf Msg = vbCancel Then
								ConCmdInputFunc = True '��ֹ���°�ť�رնԻ��򴰿�
								Exit Function
							End If
						End If
					End If
						If DlgItem$ = "GBKChangeButton" Then
							i = DlgValue("GBKFileList")
							GBKFileList(i) = NewPath
							DlgListBoxArray "GBKFileList",GBKFileList
							DlgValue "GBKFileList",i
						Else
							i = DlgValue("BIGFileList")
							BIGFileList(n) = NewPath
							DlgListBoxArray "BIGFileList",BIGFileList
							DlgValue "BIGFileList",i
						End If
					End If
				End If
			End If

			If DlgItem$ = "GBKEditButton" Or DlgItem$ = "BIGEditButton" Then
				If DlgItem$ = "GBKEditButton" Then Temp = DlgText("GBKFileList")
				If DlgItem$ = "BIGEditButton" Then Temp = DlgText("BIGFileList")
				If Temp <> "" Then
					n = 0
					For i = LBound(ConCmdDataList) To UBound(ConCmdDataList)
						SettingArr = Split(ConCmdDataList(i),JoinStr)
						TempArray = Split(SettingArr(14),SubJoinStr)
						For x = LBound(TempArray) To UBound(TempArray)
							If TempArray(x) <> "" Then
								ReDim Preserve FileList(n)
								FileList(n) = TempArray(x)
								n = n + 1
							End If
						Next x
						TempArray = Split(SettingArr(15),SubJoinStr)
						For x = LBound(TempArray) To UBound(TempArray)
							If TempArray(x) <> "" Then
								ReDim Preserve FileList(n)
								FileList(n) = TempArray(x)
								n = n + 1
							End If
						Next x
					Next i
					x = ShowPopupMenu(AppNames,vbPopupVCenterAlign)
					OpenFile(Temp,FileList,x)
				End If
			End If

			If DlgItem$ = "GBKResetButton" Or DlgItem$ = "BIGResetButton" Then
				SettingArr = Split(ConCmdDataListBak(ConCmdID),JoinStr)
				GBKFileList = Split(SettingArr(14),SubJoinStr)
				BIGFileList = Split(SettingArr(15),SubJoinStr)
				If DlgItem$ = "GBKResetButton" Then
					DlgListBoxArray "GBKFileList",GBKFileList
					DlgValue "GBKFileList",0
				Else
					DlgListBoxArray "BIGFileList",BIGFileList
					DlgValue "BIGFileList",0
				End If
			End If

			If DlgItem$ = "EditButton" Then
				n = 0
				For i = LBound(ConCmdDataList) To UBound(ConCmdDataList)
					SettingArr = Split(ConCmdDataList(i),JoinStr)
					TempArray = Split(SettingArr(14),SubJoinStr)
					For x = LBound(TempArray) To UBound(TempArray)
						If TempArray(x) <> "" Then
							ReDim Preserve FileList(n)
							FileList(n) = TempArray(x)
							If i = ConCmdID And Temp = "" Then Temp = TempArray(x)
							n = n + 1
						End If
					Next x
					TempArray = Split(SettingArr(15),SubJoinStr)
					For x = LBound(TempArray) To UBound(TempArray)
						If TempArray(x) <> "" Then
							ReDim Preserve FileList(n)
							FileList(n) = TempArray(x)
							If i = ConCmdID And Temp = "" Then Temp = TempArray(x)
							n = n + 1
						End If
					Next x
				Next i
				If Temp = "" Then Temp = FileList(0)
				If Temp <> "" Then
					x = ShowPopupMenu(AppNames,vbPopupVCenterAlign)
					OpenFile(Temp,FileList,x)
				End If
			End If

			If DlgItem$ = "ImportButton" Then
				If PSL.SelectFile(Path,True,Msg44,Msg42) = True Then
					If GetSettings("",ConCmdList,ConCmdDataList,Path) = True Then
						'GetSettings("",ConCmdListBak,ConCmdDataListBak,Path)
						DlgListBoxArray "ConCmdName",ConCmdList()
						ConCmdID = UBound(ConCmdList)
						DlgValue "ConCmdName",ConCmdID
						SettingArr = Split(ConCmdDataList(ConCmdID),JoinStr)
						GBKFileList = Split(SettingArr(14),SubJoinStr)
						BIGFileList = Split(SettingArr(15),SubJoinStr)
						DlgText "ConCmdPath",SettingArr(1)
						DlgText "GBKToBig5ANSI",SettingArr(2)
						DlgText "Big5ToGBKANSI",SettingArr(3)
						DlgText "GBKToBig5FixANSI",SettingArr(4)
						DlgText "Big5ToGBKFixANSI",SettingArr(5)
						DlgText "GBKToBig5Unicode",SettingArr(6)
						DlgText "Big5ToGBKUnicode",SettingArr(7)
						DlgText "GBKToBig5FixUnicode",SettingArr(8)
						DlgText "Big5ToGBKFixUnicode",SettingArr(9)
						DlgText "GBKToBig5UTF8",SettingArr(10)
						DlgText "Big5ToGBKUTF8",SettingArr(11)
						DlgText "GBKToBig5FixUTF8",SettingArr(12)
						DlgText "Big5ToGBKFixUTF8",SettingArr(13)
						DlgListBoxArray "GBKFileList",GBKFileList
						DlgListBoxArray "BIGFileList",BIGFileList
						DlgValue "GBKRename",StrToInteger(SettingArr(16))
						DlgValue "BIGRename",StrToInteger(SettingArr(17))
						DlgValue "GBKFileList",0
						DlgValue "BIGFileList",0
						MsgBox(Msg33,vbOkOnly+vbInformation,Msg03)
					Else
						MsgBox(Msg39 & WriteLoc,vbOkOnly+vbInformation,Msg01)
					End If
				End If
			End If

			If DlgItem$ = "ExportButton" Then
				If CheckNullData("",ConCmdDataList,"4,5,8,9,12,13,14,15,18,19",1) = True Then
					MsgBox(Msg27,vbOkOnly+vbInformation,Msg01)
					ConCmdInputFunc = True '��ֹ���°�ť�رնԻ��򴰿�
					Exit Function
				End If
				If PSL.SelectFile(Temp,False,Msg44,Msg43) = True Then
					If WriteSettings(ConCmdDataList,Temp,"All") = False Then
						MsgBox(Msg40 & Temp,vbOkOnly+vbInformation,Msg01)
					Else
						MsgBox(Msg32,vbOkOnly+vbInformation,Msg03)
					End If
				End If
			End If

			If DlgItem$ <> "CancelButton" And DlgItem$ <> "OKButton" And DlgItem$ <> "ExportButton" Then
				CmdName = DlgText("ConCmdName")
				CmdPath = DlgText("ConCmdPath")
				GBa = DlgText("GBKToBig5ANSI")
				BGa = DlgText("Big5ToGBKANSI")
				GBFa = DlgText("GBKToBig5FixANSI")
				BGFa = DlgText("Big5ToGBKFixANSI")
				GBu = DlgText("GBKToBig5Unicode")
				BGu = DlgText("Big5ToGBKUnicode")
				GBFu = DlgText("GBKToBig5FixUnicode")
				BGFu = DlgText("Big5ToGBKFixUnicode")
				GB8 = DlgText("GBKToBig5UTF8")
				BG8 = DlgText("Big5ToGBKUTF8")
				GBF8 = DlgText("GBKToBig5FixUTF8")
				BGF8 = DlgText("Big5ToGBKFixUTF8")
				If DlgValue("GBKFileList") < 0 Then GBK = "" Else GBK = Join(GBKFileList,SubJoinStr)
				If DlgValue("BIGFileList") < 0 Then BIG = "" Else BIG = Join(BIGFileList,SubJoinStr)
				rGBK = DlgValue("GBKRename")
				rBIG = DlgValue("BIGRename")
				Temp = GBa & BGa & GBFa & BGFa & GBu & BGu & GBFu & BGFu & GB8 & BG8 & GBF8 & BGF8 & GBK & BIG
				If Temp <> "" Or DlgItem$ = "GBKRename" Or DlgItem$ = "BIGRename" Then
					ConCmdData = CmdName & JoinStr & CmdPath & JoinStr & GBa & JoinStr & BGa & JoinStr & _
								GBFa & JoinStr & BGFa & JoinStr & GBu & JoinStr & BGu & JoinStr & GBFu & _
								JoinStr & BGFu & JoinStr & GB8 & JoinStr & BG8 & JoinStr & GBF8 & JoinStr & _
								BGF8 & JoinStr & GBK & JoinStr & BIG & JoinStr & rGBK & JoinStr & rBIG & _
								JoinStr & FixFileSeparator & JoinStr & FixFileSaveCode
					CreateArray(CmdName,ConCmdData,ConCmdList,ConCmdDataList)
				End If

				If InStr(LCase(CmdPath),"concmd.exe") Then
					If InStr(LCase(GBK),"gbfix.dat") Then
						DlgEnable "GBKAddButton",False
					Else
						DlgEnable "GBKAddButton",True
					End If
					If InStr(LCase(BIG),"b5fix.dat") Then
						DlgEnable "BIGAddButton",False
					Else
						DlgEnable "BIGAddButton",True
					End If
				Else
					DlgEnable "GBKAddButton",True
					DlgEnable "BIGAddButton",True
				End If

				If CheckNullData(CmdName,ConCmdDataListBak,"4,5,8,9,12,13,14,15,18,19",0) = False Then
					ConCmdID = DlgValue("ConCmdName")
					If ConCmdID < 0 Then ConCmdID = 0
					TempArray = Split(ConCmdDataListBak(ConCmdID),JoinStr)
					If Join(GBKFileList,SubJoinStr) = TempArray(14) Then
						DlgEnable "GBKResetButton",False
					Else
						DlgEnable "GBKResetButton",True
					End If
					If Join(BIGFileList,SubJoinStr) = TempArray(15) Then
						DlgEnable "BIGResetButton",False
					Else
						DlgEnable "BIGResetButton",True
					End If
				Else
					DlgEnable "GBKResetButton",False
					DlgEnable "BIGResetButton",False
				End If

				If DlgText("GBKFileList") = "" Then
					DlgEnable "GBKDelButton",False
					DlgEnable "GBKChangeButton",False
					DlgEnable "GBKEditButton",False
				Else
					DlgEnable "GBKDelButton",True
					DlgEnable "GBKChangeButton",True
					DlgEnable "GBKEditButton",True
				End If
				If DlgText("BIGFileList") = "" Then
					DlgEnable "BIGDelButton",False
					DlgEnable "BIGChangeButton",False
					DlgEnable "BIGEditButton",False
				Else
					DlgEnable "BIGDelButton",True
					DlgEnable "BIGChangeButton",True
					DlgEnable "BIGEditButton",True
				End If
				If DlgText("ConCmdName") = "" Then
					DlgEnable "ResetButton",False
					DlgEnable "FixResetButton",False
				Else
					DlgEnable "ResetButton",True
					DlgEnable "FixResetButton",True
				End If
				If DlgText("GBKFileList") & DlgText("BIGFileList") = "" Then
					DlgEnable "EditButton",False
				Else
					DlgEnable "EditButton",True
				End If
				If UBound(ConCmdList) = 0 Then
					DlgEnable "DelButton",False
				Else
					DlgEnable "DelButton",True
				End If
			End If

			If DlgItem$ = "TestButton" Or DlgItem$ = "FixTestButton" Then
				If CheckNullData("",ConCmdDataList,"4,5,8,9,12,13,14,15,18,19",1) = True Then
					MsgBox(Msg27,vbOkOnly+vbInformation,Msg01)
					ConCmdInputFunc = True '��ֹ���°�ť�رնԻ��򴰿�
					Exit Function
				End If
				ConCmdID = DlgValue("ConCmdName")
				If UBound(ExpCodeList) = 0 Then CodeID = 0
				If UBound(ExpCodeList) <> 0 Then CodeID = DlgValue("Encoding")
				Call ConCmdTest(ConCmdID,CodeID)
			End If
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
				TempArray(0) = Msg05
				TempArray(1) = Msg06
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
					DlgValue "UpdateSet",StrToInteger(UpdateSetBak(0))
					DlgText "WebSiteBox",UpdateSetBak(1)
					DlgText "CmdPathBox",UpdateSetBak(2)
					DlgText "ArgumentBox",UpdateSetBak(3)
					DlgText "UpdateCycleBox",UpdateSetBak(4)
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
					UpdateSet(5) = DlgText("UpdateDateBox")
					If DlgValue("WriteType") = 0 Then Path = FilePath
					If DlgValue("WriteType") = 1 Then Path = RegKey
					If WriteSettings(ConCmdDataList,Path,"Update") = False Then
						MsgBox(Msg36 & Path,vbOkOnly+vbInformation,Msg01)
						ConCmdInputFunc = True '��ֹ���°�ť�رնԻ��򴰿�
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
			UpdateSet = Split(Data,JoinStr)
			If DlgItem$ = "TestButton" Then Download(updateMethod,updateUrl,updateAsync,"4")
		End If

		If DlgItem$ = "OKButton" Then
			If CheckNullData("",ConCmdDataList,"0,4,5,8,9,12,13,14,15,18,19",1) = True Then
				MsgBox(Msg27,vbOkOnly+vbInformation,Msg01)
				ConCmdInputFunc = True '��ֹ���°�ť�رնԻ��򴰿�
				Exit Function
			End If
			If DlgValue("WriteType") = 0 Then Path = FilePath
			If DlgValue("WriteType") = 1 Then Path = RegKey
			If WriteSettings(ConCmdDataList,Path,"ConCmd") = False Then
				MsgBox(Msg36 & Path,vbOkOnly+vbInformation,Msg01)
				ConCmdInputFunc = True '��ֹ���°�ť�رնԻ��򴰿�
				Exit Function
			Else
				ConCmdListBak = ConCmdList
				ConCmdDataListBak = ConCmdDataList
				UpdateSetBak = UpdateSet
			End If
		End If

		If DlgItem$ = "CancelButton" Then
			ConCmdList = ConCmdListBak
			ConCmdDataList = ConCmdDataListBak
			UpdateSet = UpdateSetBak
		End If

		If DlgItem$ <> "OKButton" And DlgItem$ <> "CancelButton" Then
			ConCmdInputFunc = True ' ��ֹ���°�ť�رնԻ��򴰿�
		End If
	Case 3 ' �ı��������Ͽ��ı�������
		If DlgValue("Options") = 0 Or DlgValue("Options") = 1 Then
			ConCmdID = DlgValue("ConCmdName")
			If ConCmdID < 0 Then ConCmdID = 0
			SettingArr = Split(ConCmdDataList(ConCmdID),JoinStr)
			CmdName = DlgText("ConCmdName")
			CmdPath = DlgText("ConCmdPath")
			GBa = DlgText("GBKToBig5ANSI")
			BGa = DlgText("Big5ToGBKANSI")
			GBFa = DlgText("GBKToBig5FixANSI")
			BGFa = DlgText("Big5ToGBKFixANSI")
			GBu = DlgText("GBKToBig5Unicode")
			BGu = DlgText("Big5ToGBKUnicode")
			GBFu = DlgText("GBKToBig5FixUnicode")
			BGFu = DlgText("Big5ToGBKFixUnicode")
			GB8 = DlgText("GBKToBig5UTF8")
			BG8 = DlgText("Big5ToGBKUTF8")
			GBF8 = DlgText("GBKToBig5FixUTF8")
			BGF8 = DlgText("Big5ToGBKFixUTF8")
			If DlgValue("GBKFileList") < 0 Then GBK = "" Else GBK = SettingArr(14)
			If DlgValue("BIGFileList") < 0 Then BIG = "" Else BIG = SettingArr(15)
			rGBK = DlgValue("GBKRename")
			rBIG = DlgValue("BIGRename")
			FixFileSeparator = SettingArr(18)
			FixFileSaveCode = SettingArr(19)
			Temp = GBa & BGa & GBFa & BGFa & GBu & BGu & GBFu & BGFu & GB8 & BG8 & GBF8 & BGF8 & GBK & BIG
			If Temp <> "" Then
				ConCmdData = CmdName & JoinStr & CmdPath & JoinStr & GBa & JoinStr & BGa & JoinStr & _
							GBFa & JoinStr & BGFa & JoinStr & GBu & JoinStr & BGu & JoinStr & GBFu & _
							JoinStr & BGFu & JoinStr & GB8 & JoinStr & BG8 & JoinStr & GBF8 & JoinStr & _
							BGF8 & JoinStr & GBK & JoinStr & BIG & JoinStr & rGBK & JoinStr & rBIG & _
							JoinStr & FixFileSeparator & JoinStr & FixFileSaveCode
				CreateArray(CmdName,ConCmdData,ConCmdList,ConCmdDataList)
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
			UpdateSet = Split(Data,JoinStr)
		End If
	End Select
End Function


'�����ļ���������
Function FixFileListSet(Separator As String,FileSaveCode As String) As Boolean
	Dim i As Integer,FileCode(2) As String
	If OSLanguage = "0404" Then
		Msg01 = "�ץ��ɮװѼƳ]�w"
		Msg02 = "�ץ��ɮװѼƴN�O�n�N�ץ����ɮײM�椤���ɮ׫����ഫ�{�����n�D�ഫ���@�Ӧr��" & _
				"�Ϊ��x�s���@�ӥ]�t�ץ��ɮײM�檺��r�ɮסC�нT�{�ഫ�{�����R�O�C�O�_�䴩�ץ�" & _
				"�ɮײM��H�ή榡��s�W�ץ��ɮװѼơC"
		Msg03 = "���j�ץ��ɮײM��һݪ����j�Ÿ� (�䴩�h�X��)"
		Msg04 = "�����(&L)"
		Msg05 = "TAB ��(&T)"
		Msg06 = "��L(&O):"
		Msg07 = "�ץ����ɮײM�檺�x�s"
		Msg08 = "�ݭn�x�s�M�欰��r�ɮ�(&S)"
		Msg09 = "��r�ɮת��s�X(&C):"
		Msg10 = "ANSI"
		Msg11 = "Unicode"
		Msg12 = "UTF-8"
	Else
		Msg01 = "�����ļ���������"
		Msg02 = "�����ļ���������Ҫ���������ļ��б��е��ļ�����ת�������Ҫ��ת����һ���ַ���" & _
				"���߱���Ϊһ�����������ļ��б���ı��ļ�����ȷ��ת��������������Ƿ�֧������" & _
				"�ļ��б��Լ���ʽ����������ļ�������"
		Msg03 = "�ָ������ļ��б�����ķָ��� (֧��ת���)"
		Msg04 = "���з�(&L)"
		Msg05 = "TAB ��(&T)"
		Msg06 = "����(&O):"
		Msg07 = "�������ļ��б�ı���"
		Msg08 = "��Ҫ�����б�Ϊ�ı��ļ�(&S)"
		Msg09 = "�ı��ļ��ı���(&C):"
		Msg10 = "ANSI"
		Msg11 = "Unicode"
		Msg12 = "UTF-8"
	End If
	FixFileListSet = False
	FileCode(0) = Msg10
	FileCode(1) = Msg11
	FileCode(2) = Msg12
	Begin Dialog UserDialog 460,259,Msg01 ' %GRID:10,7,1,1
		Text 20,7,420,56,Msg02,.Text1
		GroupBox 20,70,420,56,Msg03,.SeparatorGroupBox
		OptionGroup .Separator
			OptionButton 50,91,100,21,Msg04,.vbCrLfButton
			OptionButton 160,91,100,21,Msg05,.vbTabButton
			OptionButton 270,91,90,21,Msg06,.OtherButton
		TextBox 370,91,40,21,.OtherTextBox
		GroupBox 20,140,420,77,Msg07,.SaveFileGroupBox
		CheckBox 50,161,360,14,Msg08,.SaveFileBox
		Text 50,185,240,14,Msg09,.Text2
		DropListBox 300,182,110,21,FileCode(),.FileCodeList
		OKButton 130,231,90,21,.OKButton
		CancelButton 260,231,90,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	If Separator = "\r\n" Then
		dlg.Separator = 0
	ElseIf Separator = "\t" Then
		dlg.Separator = 1
	Else
		dlg.Separator = 2
		dlg.OtherTextBox = Separator
	End If
	If FileSaveCode <> "" Then
		dlg.SaveFileBox = 1
		For i = LBound(FileCode) To UBound(FileCode)
			If FileCode(i) = FileSaveCode Then
				dlg.FileCodeList = i
				Exit For
			End If
		Next i
	End If
	If Dialog(dlg) = 0 Then Exit Function
	If dlg.Separator = 0 Then
		Separator = "\r\n"
	ElseIf dlg.Separator = 1 Then
		Separator = "\t"
	Else
		Separator = dlg.OtherTextBox
	End If
	If dlg.SaveFileBox = 1 Then
		FileSaveCode = FileCode(dlg.FileCodeList)
	Else
		FileSaveCode = ""
	End If
	FixFileListSet = True
End Function

'���ļ�
Function OpenFile(FilePath As String,FileList() As String,x As Integer) As Boolean
	Dim ExePathStr As String,Argument As String
	If OSLanguage = "0404" Then
		Msg01 = "���~"
		Msg02 = "�O�ƥ����b�t�Τ����C�п����L�}�Ҥ�k�C"
		Msg03 = "�t�ΨS�����U�}�Ҹ��ɮת��M�ε{���C�п����L�}�Ҥ�k�C"
		Msg04 = "�{�������I�i��O�{�����|��R���~�Τ��Q�䴩�A�п����L�}�Ҥ�k�C"
		Msg05 = "�L�k�}���ɮסI�M�ε{����^�F���~�N�X�A�п����L�}�Ҥ�k�C"
		Msg06 = "�L�k�}���ɮסI�M�ε{����^�F���~�N�X�A�i��O���ɮפ��Q�䴩�ΰ���ѼƦ����D�C"
		Msg07 = "�{���W��: "
		Msg08 = "��R���|: "
		Msg09 = "����Ѽ�: "
	Else
		Msg01 = "����"
		Msg02 = "���±�δ��ϵͳ���ҵ�����ѡ�������򿪷�����"
		Msg03 = "ϵͳû��ע��򿪸��ļ���Ӧ�ó�����ѡ�������򿪷�����"
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
				Return = WshShell.Run("""" & ExePath & """ " & File,1,False)
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
			ExtName = Mid(File,InStrRev(File,"."))
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
				Return = WshShell.Run("""" & ExePath & """ " & ArgumentFile,1,False)
				If Return <> 0 Then
					MsgBox Msg05,vbOkOnly+vbInformation,Msg01
				Else
					If LCase(ExeName) <> "notepad.exe" And InStr(Tools02,ExeName) = 0 Then
						Call AddArray(AppNames,AppPaths,ExeName,ExePath & JoinStr & Argument)
					End If
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
				Return = WshShell.Run("""" & ExePath & """ " & ArgumentFile,1,False)
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
				Return = WshShell.Run("""" & ExePath & """ " & ArgumentFile,1,False)
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
		PushButton 20,238,90,21,Msg06,.CleanButton
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
		DlgEnable "CleanButton",False
	Case 2 ' ��ֵ���Ļ��߰����˰�ť
		If DlgItem$ = "BrowseButton" Then
			If PSL.SelectFile(File,True,Msg03,Msg02) = True Then
 				DlgText "CmdPath",File
 			Else
 				DlgText "CmdPath",""
 			End If
		End If
		If DlgItem$ = "CleanButton" Then
 			DlgText "CmdPath",""
 			DlgText "Argument",""
 			DlgEnable "CleanButton",False
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
 				DlgEnable "CleanButton",False
 			Else
 				DlgEnable "CleanButton",True
 			End If
 			CmdInputFunc = True ' ��ֹ���°�ť�رնԻ��򴰿�
		End If
	Case 3 ' �ı��������Ͽ��ı�������
 		If DlgText("CmdPath") = "" And DlgText("Argument") = "" Then
 			DlgEnable "CleanButton",False
 		Else
 			DlgEnable "CleanButton",True
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
		Text 10,7,180,14,Msg04
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


'������������
Function EditSet(DataArr() As String,Header As String) As String
	Dim tempHeader As String,NewHeader As String
	If OSLanguage = "0404" Then
		Msg01 = "�ܧ�"
		Msg04 = "�s�W��:"
		Msg06 = "���~"
		Msg07 = "�z�S����J���󤺮e�I�Э��s��J�C"
		Msg08 = "�ӦW�٤w�g�s�b�I�п�J�@�Ӥ��P���W�١C"
		Msg09 = "�¦W��:"
	Else
		Msg01 = "����"
		Msg04 = "������:"
		Msg06 = "����"
		Msg07 = "��û�������κ����ݣ����������롣"
		Msg08 = "�������Ѿ����ڣ�������һ����ͬ�����ơ�"
		Msg09 = "������:"
	End If
	tempHeader = Header
	If InStr(Header,"&") Then
		tempHeader = Replace(Header,"&","&&")
	End If

	Begin Dialog UserDialog 310,126,Msg01 ' %GRID:10,7,1,1
		GroupBox 10,17,290,28,"",.GroupBox1
		Text 10,7,185,14,Msg09,.Text1
		Text 20,28,270,14,tempHeader,.oldNameText
		Text 10,53,180,14,Msg04,.newNameText
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


'��ȡ����
Function GetSettings(SelSet As String,CmdList() As String,DataList() As String,Path As String) As Boolean
	Dim i As Integer,j As Integer,Header As String,HeaderIDs As String,HeaderIDArr() As String,Stemp As Boolean
	GetSettings = False
	NewVersion = ToUpdateVersion

	If Path = RegKey Then GoTo GetFromRegistry
	If Path = "" Then Path = FilePath
	If Dir(Path) = "" Then GoTo GetFromRegistry
	If Path = FilePath Then On Error GoTo GetFromRegistry
	Open Path For Input As #1
	i = 0
	Header = ""
	While Not EOF(1)
		Line Input #1,L$
		If Left(Trim(L$),1) = "[" And Right(Trim(L$),1) = "]" Then
			Header = Mid(Trim(L$),2,Len(Trim(L$))-2)
		End If
		If L$ <> "" And Header <> "" Then
			setPreStr = ""
			setAppStr = ""
			Site = ""
			GBK = ""
			BIG = ""
			If InStr(L$,"=") Then setPreStr = Trim(Left(L$,InStr(L$,"=")-1))
			If InStr(L$,"=") Then setAppStr = LTrim(Mid(L$,InStr(L$,"=")+1))
			'��ȡ Option ���ֵ
			If setPreStr = "Version" Then OldVersion = setAppStr
			If Header = "Option" And SelSet = "" And setPreStr <> "" Then
				If setPreStr = "ConCmdSeleted" Then ConCmdName = setAppStr
				If setPreStr = "AddinID" Then AddinID = setAppStr
				If setPreStr = "TextExpCharSet" Then TextExpCharSet = setAppStr
				If setPreStr = "WordFixSelect" Then WordFixSelect = setAppStr
				If setPreStr = "AllHandle" Then AllHandle = setAppStr
				If setPreStr = "AllConTypeSame" Then AllTypeSame = setAppStr
				If setPreStr = "AllConListSame" Then AllListSame = setAppStr
				If setPreStr = "CycleSelect" Then CycleSelect = setAppStr
				If setPreStr = "KeepItemSelect" Then KeepItemSelect = setAppStr
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
				If setPreStr = "ConCmdPath" Then CmdPath = setAppStr
				If setPreStr = "GBKToBig5ANSI" Then GBa = setAppStr
				If setPreStr = "Big5ToGBKANSI" Then BGa = setAppStr
				If setPreStr = "GBKToBig5FixANSI" Then GBFa = setAppStr
				If setPreStr = "Big5ToGBKFixANSI" Then BGFa = setAppStr
				If setPreStr = "GBKToBig5Unicode" Then GBu = setAppStr
				If setPreStr = "Big5ToGBKUnicode" Then BGu = setAppStr
				If setPreStr = "GBKToBig5FixUnicode" Then GBFu = setAppStr
				If setPreStr = "Big5ToGBKFixUnicode" Then BGFu = setAppStr
				If setPreStr = "GBKToBig5UTF8" Then GB8 = setAppStr
				If setPreStr = "Big5ToGBKUTF8" Then BG8 = setAppStr
				If setPreStr = "GBKToBig5FixUTF8" Then GBF8 = setAppStr
				If setPreStr = "Big5ToGBKFixUTF8" Then BGF8 = setAppStr
				If setPreStr = "GBKFixFilePath" Then GBK = setAppStr
				If setPreStr = "Big5FixFilePath" Then BIG = setAppStr
				If setPreStr = "RenameGBKFixFile" Then rGBK = setAppStr
				If setPreStr = "RenameBig5FixFile" Then rBIG = setAppStr
				If InStr(setPreStr,"GBKFixFile_") Then GBK = setAppStr
				If InStr(setPreStr,"Big5FixFile_") Then BIG = setAppStr
				If GBK <> "" And Dir(GBK) <> "" Then
					If GBKFiles <> "" Then GBKFiles = GBKFiles & SubJoinStr & GBK
					If GBKFiles = "" Then GBKFiles = GBK
				End If
				If BIG <> "" And Dir(BIG) <> "" Then
					If BIGFiles <> "" Then BIGFiles = BIGFiles & SubJoinStr & BIG
					If BIGFiles = "" Then BIGFiles = BIG
				End If
				If setPreStr = "FixFileListSeparator" Then SprtStr = setAppStr
				If setPreStr = "FixFileListSaveCode" Then SaveCode = setAppStr
			End If
		End If
		If (L$ = "" Or EOF(1)) And Header = "Option" Then
			Data = ConCmdName & JoinStr & AddinID & JoinStr & TextExpCharSet & JoinStr & _
					WordFixSelect & JoinStr & AllHandle & JoinStr & AllTypeSame & JoinStr & _
					AllListSame & JoinStr & CycleSelect & JoinStr & KeepItemSelect
			cSelected = Split(Data,JoinStr)
		End If
		If (L$ = "" Or EOF(1)) And Header = "Update" Then
			Data = UpdateMode & JoinStr & UpdateSite & JoinStr & CmdPath & JoinStr & CmdArg & _
					JoinStr & UpdateCycle & JoinStr & UpdateDate
			UpdateSet = Split(Data,JoinStr)
		End If
		If (L$ = "" Or EOF(1)) And Header <> "" And Header <> "Option" And Header <> "Update" Then
			If CmdPath <> "" And Dir(CmdPath) <> "" Then
				Data = Header & JoinStr & CmdPath & JoinStr & GBa & JoinStr & BGa & JoinStr & _
						GBFa & JoinStr & BGFa & JoinStr & GBu & JoinStr & BGu & JoinStr & GBFu & _
						JoinStr & BGFu & JoinStr & GB8 & JoinStr & BG8 & JoinStr & GBF8 & JoinStr & _
						BGF8 & JoinStr & GBKFiles & JoinStr & BIGFiles & JoinStr & rGBK & JoinStr & _
						rBIG & JoinStr & SprtStr & JoinStr & SaveCode
				'���¾ɰ��Ĭ������ֵ
				If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 Then
					Data = UpdateSetting(Header,CmdPath,Data)
				End If
				'�������ݵ�������
				CreateArray(Header,Data,CmdList,DataList)
				GetSettings = True
			End If
			'���ݳ�ʼ��
			CmdPath = ""
			GBa = ""
			BGa = ""
			GBFa = ""
			BGFa = ""
			GBu = ""
			BGu = ""
			GBFu = ""
			BGFu = ""
			GB8 = ""
			BG8 = ""
			GBF8 = ""
			BGF8 = ""
			GBKFiles = ""
			BIGFiles = ""
			rGBK = ""
			rBIG = ""
			SprtStr = ""
			SaveCode = ""
		End If
	Wend
	Close #1
	If Path = FilePath Then On Error GoTo 0
	'������º͵��������ݵ��ļ�
	If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 And Path = FilePath Then
		If Dir(FilePath) <> "" Then WriteSettings(DataList,FilePath,"All")
	End If
	If WriteLoc = "" Then WriteLoc = FilePath
	Exit Function

	GetFromRegistry:
	'��ȡ Option ���ֵ
	On Error Resume Next
	OldVersion = GetSetting("gb2big5","Settings","Version")
	If SelSet = "" Then
		ConCmd = GetSetting("gb2big5","Settings","ConCmd")
		ConCmdName = GetSetting("gb2big5","Settings","Name")
		AddinID = GetSetting("gb2big5","Settings","AddinID","")
		TextExpCharSet = GetSetting("gb2big5","Settings","TextExpCharSet",0)
		WordFixSelect = GetSetting("gb2big5","Settings","WordFixSelect",1)
		AllHandle =  GetSetting("gb2big5","Settings","AllHandle",0)
		CycleSelect = GetSetting("gb2big5","Settings","CycleSelect",0)
		AllTypeSame = GetSetting("gb2big5","Settings","AllConTypeSame",1)
		AllListSame = GetSetting("gb2big5","Settings","AllConListSame",1)
		KeepItemSelect = GetSetting("gb2big5","Settings","KeepItemSelect",1)
		ConCmd = RemoveBackslash(ConCmd,"""","""",1)
		ConCmd = AppendBackslash(ConCmd,"","\",1)
		Data = ConCmdName & JoinStr & AddinID & JoinStr & TextExpCharSet & JoinStr & _
				WordFixSelect & JoinStr & AllHandle & JoinStr & AllTypeSame & JoinStr & _
				AllListSame & JoinStr & CycleSelect & JoinStr & KeepItemSelect
		cSelected = Split(Data,JoinStr)
		'��ȡ Update ���ֵ
		UpdateMode = GetSetting("gb2big5","Update","UpdateMode",1)
		Count = GetSetting("gb2big5","Update","Count",0)
		For i = 0 To Count
			Site = GetSetting("gb2big5","Update",CStr(i),"")
			If Site <> "" Then
				If i > 0 Then UpdateSite = UpdateSite & vbCrLf & Site
				If i = 0 Then UpdateSite = Site
			End If
		Next i
		CmdPath = GetSetting("gb2big5","Update","Path","")
		CmdArg = GetSetting("gb2big5","Update","Argument","")
		UpdateCycle = GetSetting("gb2big5","Update","UpdateCycle",7)
		UpdateDate = GetSetting("gb2big5","Update","UpdateDate","")
		Data = UpdateMode & JoinStr & UpdateSite & JoinStr & CmdPath & JoinStr & CmdArg & _
				JoinStr & UpdateCycle & JoinStr & UpdateDate
		UpdateSet = Split(Data,JoinStr)
	End If

	HeaderIDs = GetSetting("Gb2Big5","Settings","ConCmdList")
	If HeaderIDs = "" Then
		'��ȡ�ɰ�Ĳ������ֵ
		Header = GetSetting("gb2big5","Settings","Name")
		CmdPath = GetSetting("gb2big5","Settings","Path")
		If CmdPath <> "" And Dir(CmdPath) <> "" Then
			'�������ݵ�������
			DefaultData = DefaultSetting(Header,CmdPath)
			Data = Header & JoinStr & CmdPath & JoinStr & DefaultData
			Data = UpdateSetting(Header,CmdPath,Data)
			CreateArray(Header,Data,CmdList,DataList)
			GetSettings = True
		End If
	Else
		'��ȡ Option ������ֵ
		HeaderIDArr = Split(HeaderIDs,";",-1)
		For i = LBound(HeaderIDArr) To UBound(HeaderIDArr)
			HeaderID = HeaderIDArr(i)
			'��ȡ�������ֵ
			Header = GetSetting("gb2big5",HeaderID,"CmdName")
			CmdPath = GetSetting("gb2big5",HeaderID,"ConCmdPath")
			GBa = GetSetting("gb2big5",HeaderID,"GBKToBig5ANSI")
			BGa = GetSetting("gb2big5",HeaderID,"Big5ToGBKANSI")
			GBFa = GetSetting("gb2big5",HeaderID,"GBKToBig5FixANSI")
			BGFa = GetSetting("gb2big5",HeaderID,"Big5ToGBKFixANSI")
			GBu = GetSetting("gb2big5",HeaderID,"GBKToBig5Unicode")
			BGu = GetSetting("gb2big5",HeaderID,"Big5ToGBKUnicode")
			GBFu = GetSetting("gb2big5",HeaderID,"GBKToBig5FixUnicode")
			BGFu = GetSetting("gb2big5",HeaderID,"Big5ToGBKFixUnicode")
			GB8 = GetSetting("gb2big5",HeaderID,"GBKToBig5UTF8")
			BG8 = GetSetting("gb2big5",HeaderID,"Big5ToGBKUTF8")
			GBF8 = GetSetting("gb2big5",HeaderID,"GBKToBig5FixUTF8")
			BGF8 = GetSetting("gb2big5",HeaderID,"Big5ToGBKFixUTF8")
			rGBK = GetSetting("gb2big5",HeaderID,"RenameGBKFixFile")
			rBIG = GetSetting("gb2big5",HeaderID,"RenameBig5FixFile")
			Count = GetSetting("gb2big5",HeaderID,"GBKFixFileCount",0)
			For j = 0 To Count
				GBK = GetSetting("gb2big5",HeaderID,"GBKFixFile_" & CStr(j),"")
				If GBK = "" Then GBK = GetSetting("gb2big5",HeaderID,"GBKFixFilePath","")
				If GBK <> "" And Dir(GBK) <> "" Then
					If j > 0 Then GBKFiles = GBKFiles & SubJoinStr & GBK
					If j = 0 Then GBKFiles = GBK
				End If
			Next j
			Count = GetSetting("gb2big5",HeaderID,"Big5FixFileCount",0)
			For j = 0 To Count
				BIG = GetSetting("gb2big5",HeaderID,"Big5FixFile_" & CStr(j),"")
				If BIG = "" Then BIG = GetSetting("gb2big5",HeaderID,"Big5FixFilePath","")
				If BIG <> "" And Dir(BIG) <> "" Then
					If j > 0 Then BIGFiles = BIGFiles & SubJoinStr & BIG
					If j = 0 Then BIGFiles = BIG
				End If
			Next j
			SprtStr = GetSetting("gb2big5",HeaderID,"FixFileListSeparator")
			SaveCode = GetSetting("gb2big5",HeaderID,"FixFileListSaveCode")
			If CmdPath <> "" And Dir(CmdPath) <> "" Then
				Data = Header & JoinStr & CmdPath & JoinStr & GBa & JoinStr & BGa & JoinStr & _
						GBFa & JoinStr & BGFa & JoinStr & GBu & JoinStr & BGu & JoinStr & GBFu & _
						JoinStr & BGFu & JoinStr & GB8 & JoinStr & BG8 & JoinStr & GBF8 & JoinStr & _
						BGF8 & JoinStr & GBKFiles & JoinStr & BIGFiles & JoinStr & rGBK & JoinStr & _
						rBIG & JoinStr & SprtStr & JoinStr & SaveCode
				'���¾ɰ��Ĭ������ֵ
				If OldVersion <> "" And StrComp(NewVersion,OldVersion) = 1 Then
					Data = UpdateSetting(Header,CmdPath,Data)
				End If
				'�������ݵ�������
				CreateArray(Header,Data,CmdList,DataList)
				GetSettings = True
			End If
		Next i
	End If
	On Error GoTo 0
	If WriteLoc = "" Then WriteLoc = RegKey
End Function


'д������
Function WriteSettings(DataList() As String,Path As String,WriteType As String) As Boolean
	Dim i As Integer,j As Integer,Header As String,HeaderIDs As String,TempPath As String
	Dim SetsArray() As String,HeaderIDArr() As String
	WriteSettings = False
	If ConCmd <> "" Then ConCmdPath = RemoveBackslash(ConCmd,"","\",1)
	KeepSet = cSelected(UBound(cSelected))

	'д���ļ�
	If Path <> "" And Path <> RegKey Then
		Header = ""
		FindHeader = 0
   		On Error Resume Next
		TempPath = Left(Path,InStrRev(Path,"\"))
   		If Dir(TempPath & "*.*") = "" Then MkDir TempPath
		If Dir(Path) <> "" Then SetAttr Path,vbNormal
		On Error GoTo 0
		On Error GoTo ExitFunction
		Open Path For Output As #2
			Print #2,";------------------------------------------------------------"
			Print #2,";Settings for PSLGbk2Big5.bas"
			Print #2,";------------------------------------------------------------"
			Print #2,""
			Print #2,"[Option]"
			Print #2,"Version = " & Version
			If KeepSet = 1 Then
				Print #2,"ConCmdSeleted = " & cSelected(0)
				Print #2,"AddinID = " & cSelected(1)
				Print #2,"TextExpCharSet = " & cSelected(2)
				Print #2,"WordFixSelect = " & cSelected(3)
				Print #2,"AllHandle = " & cSelected(4)
				Print #2,"AllConTypeSame = " & cSelected(5)
				Print #2,"AllConListSame = " & cSelected(6)
				Print #2,"CycleSelect = " & cSelected(7)
				Print #2,"KeepItemSelect = " & cSelected(8)
			End If
			Print #2,""
			If Join(UpdateSet) <> "" Then
				UpdateSiteList = Split(UpdateSet(1),vbCrLf,-1)
				Print #2,"[Update]"
				Print #2,"UpdateMode = " & UpdateSet(0)
				For i = LBound(UpdateSiteList) To UBound(UpdateSiteList)
					Print #2,"Site_" & CStr(i) & " = " & UpdateSiteList(i)
				Next i
				Print #2,"Path = " & UpdateSet(2)
				Print #2,"Argument = " & UpdateSet(3)
				Print #2,"UpdateCycle = " & UpdateSet(4)
				Print #2,"UpdateDate = " & UpdateSet(5)
				Print #2,""
			End If
			For i = LBound(DataList) To UBound(DataList)
				ConCmdData = DataList(i)
				SetsArray = Split(ConCmdData,JoinStr)
				Print #2,"[" & SetsArray(0) & "]"
				Print #2,"ConCmdPath = " & SetsArray(1)
				Print #2,"GBKToBig5ANSI = " & SetsArray(2)
				Print #2,"Big5ToGBKANSI = " & SetsArray(3)
				Print #2,"GBKToBig5FixANSI = " & SetsArray(4)
				Print #2,"Big5ToGBKFixANSI = " & SetsArray(5)
				Print #2,"GBKToBig5Unicode = " & SetsArray(6)
				Print #2,"Big5ToGBKUnicode = " & SetsArray(7)
				Print #2,"GBKToBig5FixUnicode = " & SetsArray(8)
				Print #2,"Big5ToGBKFixUnicode = " & SetsArray(9)
				Print #2,"GBKToBig5UTF8 = " & SetsArray(10)
				Print #2,"Big5ToGBKUTF8 = " & SetsArray(11)
				Print #2,"GBKToBig5FixUTF8 = " & SetsArray(12)
				Print #2,"Big5ToGBKFixUTF8 = " & SetsArray(13)
				TempArray = Split(SetsArray(14),SubJoinStr)
				For j = LBound(TempArray) To UBound(TempArray)
					Print #2,"GBKFixFile_" & CStr(j) & " = " & TempArray(j)
				Next j
				TempArray = Split(SetsArray(15),SubJoinStr)
				For j = LBound(TempArray) To UBound(TempArray)
					Print #2,"Big5FixFile_" & CStr(j) & " = " & TempArray(j)
				Next j
				Print #2,"RenameGBKFixFile = " & SetsArray(16)
				Print #2,"RenameBig5FixFile = " & SetsArray(17)
				Print #2,"FixFileListSeparator = " & SetsArray(18)
				Print #2,"FixFileListSaveCode = " & SetsArray(19)
				If i <> UBound(DataList) Then Print #2,""
			Next i
		Close #2
		On Error GoTo 0
		WriteSettings = True
		If Path = FilePath Then WriteLoc = FilePath
		If Path = FilePath Then GoTo RemoveRegKey
	'д��ע���
	ElseIf Path = RegKey Then
		On Error GoTo ExitFunction
		SaveSetting("gb2big5","Settings","Version",Version)
		If WriteType = "Main" Or WriteType = "All" Then
			If KeepSet = 1 Then
				SaveSetting("gb2big5","Settings","ConCmd",ConCmdPath)
				SaveSetting("gb2big5","Settings","Name",cSelected(0))
				SaveSetting("gb2big5","Settings","AddinID",cSelected(1))
				SaveSetting("gb2big5","Settings","TextExpCharSet",cSelected(2))
				SaveSetting("gb2big5","Settings","WordFixSelect",cSelected(3))
				SaveSetting("gb2big5","Settings","AllHandle",cSelected(4))
				SaveSetting("gb2big5","Settings","AllConTypeSame",cSelected(5))
				SaveSetting("gb2big5","Settings","AllConListSame",cSelected(6))
				SaveSetting("gb2big5","Settings","CycleSelect",cSelected(7))
				SaveSetting("gb2big5","Settings","KeepItemSelect",cSelected(8))
			End If
		End If
		If WriteType = "ConCmd" Or WriteType = "All" Then
			'ɾ��ԭ������
			HeaderIDs = GetSetting("gb2big5","Settings","ConCmdList")
			If HeaderIDs <> "" Then
				HeaderIDArr = Split(HeaderIDs,";",-1)
				For i = LBound(HeaderIDArr) To UBound(HeaderIDArr)
					On Error Resume Next
					DeleteSetting("gb2big5",HeaderIDArr(i))
					On Error GoTo 0
				Next i
			End If
			'д����������
			For i = LBound(DataList) To UBound(DataList)
				ReDim Preserve HeaderIDArr(i)
				HeaderID = CStr(i)
				HeaderIDArr(i) = HeaderID
				SetsArray = Split(DataList(i),JoinStr)
				SaveSetting("gb2big5",HeaderID,"CmdName",SetsArray(0))
				SaveSetting("gb2big5",HeaderID,"ConCmdPath",SetsArray(1))
				SaveSetting("gb2big5",HeaderID,"GBKToBig5ANSI",SetsArray(2))
				SaveSetting("gb2big5",HeaderID,"Big5ToGBKANSI",SetsArray(3))
				SaveSetting("gb2big5",HeaderID,"GBKToBig5FixANSI",SetsArray(4))
				SaveSetting("gb2big5",HeaderID,"Big5ToGBKFixANSI",SetsArray(5))
				SaveSetting("gb2big5",HeaderID,"GBKToBig5Unicode",SetsArray(6))
				SaveSetting("gb2big5",HeaderID,"Big5ToGBKUnicode",SetsArray(7))
				SaveSetting("gb2big5",HeaderID,"GBKToBig5FixUnicode",SetsArray(8))
				SaveSetting("gb2big5",HeaderID,"Big5ToGBKFixUnicode",SetsArray(9))
				SaveSetting("gb2big5",HeaderID,"GBKToBig5UTF8",SetsArray(10))
				SaveSetting("gb2big5",HeaderID,"Big5ToGBKUTF8",SetsArray(11))
				SaveSetting("gb2big5",HeaderID,"GBKToBig5FixUTF8",SetsArray(12))
				SaveSetting("gb2big5",HeaderID,"Big5ToGBKFixUTF8",SetsArray(13))
				TempArray = Split(SetsArray(14),SubJoinStr)
				For j = LBound(TempArray) To UBound(TempArray)
					SaveSetting("gb2big5",HeaderID,"GBKFixFile_" & CStr(j),TempArray(j))
				Next j
				SaveSetting("gb2big5",HeaderID,"GBKFixFileCount",UBound(TempArray))
				TempArray = Split(SetsArray(15),SubJoinStr)
				For j = LBound(TempArray) To UBound(TempArray)
					SaveSetting("gb2big5",HeaderID,"Big5FixFile_" & CStr(j),TempArray(j))
				Next j
				SaveSetting("gb2big5",HeaderID,"Big5FixFileCount",UBound(TempArray))
				SaveSetting("gb2big5",HeaderID,"RenameGBKFixFile",SetsArray(16))
				SaveSetting("gb2big5",HeaderID,"RenameBig5FixFile",SetsArray(17))
				SaveSetting("gb2big5",HeaderID,"FixFileListSeparator",SetsArray(18))
				SaveSetting("gb2big5",HeaderID,"FixFileListSaveCode",SetsArray(19))
			Next i
			HeaderIDs = Join(HeaderIDArr,";")
			SaveSetting("gb2big5","Settings","ConCmdList",HeaderIDs)
		End If
		If WriteType = "Update" Or WriteType = "ConCmd" Or WriteType = "All" Then
			If Join(UpdateSet) <> "" Then
				On Error Resume Next
				DeleteSetting("gb2big5","Update")
				On Error GoTo 0
				UpdateSiteList = Split(UpdateSet(1),vbCrLf,-1)
				SaveSetting("gb2big5","Update","UpdateMode",UpdateSet(0))
				For i = LBound(UpdateSiteList) To UBound(UpdateSiteList)
					SaveSetting("gb2big5","Update",CStr(i),UpdateSiteList(i))
				Next i
				SaveSetting("gb2big5","Update","Count",UBound(UpdateSiteList))
				SaveSetting("gb2big5","Update","Path",UpdateSet(2))
				SaveSetting("gb2big5","Update","Argument",UpdateSet(3))
				SaveSetting("gb2big5","Update","UpdateCycle",UpdateSet(4))
				SaveSetting("gb2big5","Update","UpdateDate",UpdateSet(5))
			End If
		End If
		WriteSettings = True
		WriteLoc = RegKey
		GoTo RemoveFilePath
	'ɾ�����б��������
	ElseIf Path = "" Then
		'ɾ���ļ�������
 		RemoveFilePath:
		On Error Resume Next
		If Dir(FilePath) <> "" Then
			SetAttr FilePath,vbNormal
			Kill FilePath
		End If
		TempFilePath = Left(FilePath,InStrRev(FilePath,"\"))
		If Dir(TempFilePath & "*.*") = "" Then RmDir TempFilePath
		On Error GoTo 0
		If Path = RegKey Then Exit Function
		'ɾ��ע���������
		RemoveRegKey:
		If GetSetting("gb2big5","Settings","Version") <> "" Then
			HeaderIDs = GetSetting("gb2big5","Settings","ConCmdList")
			On Error Resume Next
			If HeaderIDs <> "" Then
				HeaderIDArr = Split(HeaderIDs,";",-1)
				For i = LBound(HeaderIDArr) To UBound(HeaderIDArr)
					DeleteSetting("gb2big5",HeaderIDArr(i))
				Next i
			End If
			DeleteSetting("gb2big5","Settings","Version")
			DeleteSetting("gb2big5","Settings","ConCmdList")
			DeleteSetting("gb2big5","Settings","AddinID")
			DeleteSetting("gb2big5","Settings","TextExpCharSet")
			DeleteSetting("gb2big5","Settings","WordFixSelect")
			DeleteSetting("gb2big5","Settings","AllHandle")
			DeleteSetting("gb2big5","Settings","AllConTypeSame")
			DeleteSetting("gb2big5","Settings","AllConListSame")
			DeleteSetting("gb2big5","Settings","CycleSelect")
			DeleteSetting("gb2big5","Settings","KeepItemSelect")
			DeleteSetting("gb2big5","Settings","Name")
			DeleteSetting("gb2big5","Settings","Path")
			DeleteSetting("gb2big5","Settings","GBKToBig5Argument")
			DeleteSetting("gb2big5","Settings","GBKToBig5ArgumentFix")
			DeleteSetting("gb2big5","Settings","Big5ToGBKArgument")
			DeleteSetting("gb2big5","Settings","Big5ToGBKArgumentFix")
			DeleteSetting("gb2big5","Settings","KeepConCmdSet")
			DeleteSetting("gb2big5","Settings","KeepAddinIDSet")
			DeleteSetting("gb2big5","Update")
			On Error GoTo 0
		End If
		If Path = FilePath Then Exit Function
		'����д��λ������Ϊ��
		WriteSettings = True
		WriteLoc = ""
	End If
	ExitFunction:
End Function


'���¼��ɰ汾����ֵ
Function UpdateSetting(CmdName As String,CmdPath As String,Data As String) As String
	Dim i As Integer,Path As String,Stemp As Boolean
	Stemp = False
	If InStr(LCase(CmdName),"concmd") Then Stemp = True
	If InStr(LCase(CmdName),"convertz") Then Stemp = True
	If InStr(LCase(CmdPath),"concmd.exe") Then Stemp = True
	If InStr(LCase(CmdPath),"convertz.exe") Then Stemp = True
	If Stemp = False Then
		UpdateSetting = Data
		Exit Function
	Else
		SetsArray = Split(Data,JoinStr)
		DefaultData = DefaultSetting(CmdName,CmdPath)
		TempArray = Split(DefaultData,JoinStr)
		For i = 2 To UBound(SetsArray)
			If i < 14 Then SetsArray(i) = TempArray(i-2)
			If i > 15 Then SetsArray(i) = TempArray(i-2)
		Next i
		UpdateSetting = Join(SetsArray,JoinStr)
	End If
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


'����������Ŀ
Function ChangeList(AllList() As String,UseList() As String) As Variant
	Dim i As Integer,n As Integer,TempList() As String,Stemp As Boolean
	n = 0
	ReDim TempList(0)
	For i = LBound(AllList) To UBound(AllList)
		Stemp = False
		For j = LBound(UseList) To UBound(UseList)
			If UseList(j) = AllList(i) Then
				Stemp = True
				Exit For
			End If
		Next j
		If Stemp = False Then
			ReDim Preserve TempList(n)
			TempList(n) = AllList(i)
			n = n + 1
		End If
	Next i
	ChangeList = TempList
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
		SetsArray = Split(DataList(i),JoinStr)
		If Header <> "" And SetsArray(0) = Header Then hStemp = True
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
			If fType = 0 Then
				If Header <> "" Then
					If n = UBound(SetsArray)-UBound(SkipNumArray) Then CheckNullData = True
				Else
					If n = UBound(SetsArray)-UBound(SkipNumArray) Then m = m + 1
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
	If rMax = 0 Or CompType = "" Or Operator ="" Then
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


'��Ӳ���ͬ������Ԫ��
Sub AddArray(AppNames() As String,AppPaths() As String,CmdName As String,CmdPath As String)
	Dim i As Integer,n As Integer,Stemp As Boolean
	If CmdName = "" And CmdPath = "" Then Exit Sub
	n = UBound(AppNames)
	Stemp = False
	For i = LBound(AppNames) To UBound(AppNames)
		If LCase(CmdName) = LCase(AppNames(i)) Then
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
	If folderPath = "" Then	Exit Function
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


'�ִ�ǰ�󸽼�ָ���� PreStr �� AppStr
Function AppendBackslash(folderPath As String,PreStr As String,AppStr As String,Spase As Integer) As String
	Dim folderPathBak As String
	If folderPath = "" Then Exit Function
	folderPathBak = folderPath
	If Spase = 1 Then folderPathBak = Trim(folderPathBak)
	If PreStr <> "" And Left(folderPathBak,Len(PreStr)) <> PreStr Then
		folderPathBak = PreStr & folderPathBak
	End If
	If AppStr <> "" And Right(folderPathBak,Len(AppStr)) <> AppStr Then
		folderPathBak = folderPathBak & AppStr
	End If
	AppendBackslash = folderPathBak
End Function


'�������ļ���
Function FileRename(Source As String,Target As String) As Boolean
	FileRename = False
	If Dir(Source) = "" Then Exit Function
	On Error GoTo SysErrorMsg
	If Dir(Target) <> "" Then
		SetAttr Source,vbNormal
		SetAttr Target,vbNormal
		Kill Target
		Name Source As Target
	Else
		SetAttr Source,vbNormal
		Name Source As Target
	End If
	On Error GoTo 0
	If Dir(Source) = "" And Dir(Target) <> "" Then
		FileRename = True
	End If
	Exit Function
	SysErrorMsg:
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


'ת���ַ�Ϊ������ֵ
Function StrToInteger(mStr As String) As Integer
	If mStr = "" Then mStr = "0"
	StrToInteger = CInt(mStr)
End Function


'����ת������
Sub ConCmdTest(ConCmdID As Integer,CodeID As Integer)
	Dim i As Integer,FixList(1) As String
	If OSLanguage = "0404" Then
		Msg01 = "����"
		Msg02 = "�ഫ�������i�ﶵ�ط|�ھڱM�ײ{�����y���۰ʽT�w�C�n�W�[���ؽзs�W�������y���C"
		Msg03 = "�ഫ����:"
		Msg04 = "���J�ץ�:"
		Msg05 = "��J���(�۰ʱq��ܲM�椤Ū�J�r��A�]�i�H��ʿ�J):"
		Msg06 = "�ഫ���G(�����ի��s��b����X���G�A�P�B�ഫ�ɦ۰ʦP�B):"
		Msg07 = "����(&T)"
		Msg08 = "�M��(&C)"
		Msg09 = "����(&E)"
		Msg15 = "���ץ����J"
		Msg16 = "�ץ����J"
		Msg17 = "Ū�J���:"
		Msg18 = "�P�B�ഫ"
		Msg21 = "�ഫ�{��:"
		Msg22 = "�ץX�r����:"
		Msg23 = "�s��ץ���(&F)"
	Else
		Msg01 = "����"
		Msg02 = "ת�����͵Ŀ�ѡ��Ŀ����ݷ������е������Զ�ȷ����Ҫ������Ŀ�������Ӧ�����ԡ�"
		Msg03 = "ת������:"
		Msg04 = "�ʻ�����:"
		Msg05 = "����Դ��(�Զ���ѡ���б��ж����ִ���Ҳ�����ֶ�����):"
		Msg06 = "ת�����(�����԰�ť���ڴ���������ͬ��ת��ʱ�Զ�ͬ��):"
		Msg07 = "����(&T)"
		Msg08 = "���(&C)"
		Msg09 = "�˳�(&E)"
		Msg15 = "�������ʻ�"
		Msg16 = "�����ʻ�"
		Msg17 = "��������:"
		Msg18 = "ͬ��ת��"
		Msg21 = "ת������:"
		Msg22 = "�����ַ���:"
		Msg23 = "�༭������(&F)"
	End If

	FixList(0) = Msg15
	FixList(1) = Msg16
	Begin Dialog UserDialog 650,483,Msg01,.ConCmdTestFunc ' %GRID:10,7,1,1
		Text 20,7,610,21,Msg02
		Text 20,35,300,14,Msg03
		Text 330,35,300,14,Msg04
		DropListBox 20,49,300,21,ConTypeList(),.TypeListBox
		DropListBox 330,49,300,21,FixList(),.FixListBox
		Text 20,77,300,14,Msg21
		Text 330,77,300,14,Msg22
		DropListBox 20,91,300,21,ConCmdList(),.CmdNameList
		DropListBox 330,91,300,21,ExpCodeList(),.ExpCodeList
		Text 20,119,470,21,Msg05
		Text 20,287,470,21,Msg06
		Text 500,119,80,21,Msg17
		TextBox 580,117,50,18,.LineNumBox
		CheckBox 500,285,130,21,Msg18,.CheckBox,1
		TextBox 20,140,610,140,.InTextBox,1
		TextBox 20,308,610,140,.OutTextBox,1
		PushButton 20,455,90,21,Msg07,.TestButton
		PushButton 120,455,90,21,Msg08,.CleanButton
		PushButton 270,455,140,21,Msg23,.EditButton
		CancelButton 540,455,90,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	dlg.CmdNameList = ConCmdID
	dlg.ExpCodeList = CodeID
	If Dialog(dlg) = 0 Then Exit Sub
End Sub


'����ת������Ի�����
Private Function ConCmdTestFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim dummyTrn As PslTransList,inText As String,outText As String,TypeList As String
	Dim TypeListID As Integer,FixListID As Integer,LineNum As Integer,Sync As Integer
	Dim Code As String,ConCmdID As Integer,x As Integer,ConTypeID As Integer
	Dim TempArray() As String,ConList() As String

	If OSLanguage = "0404" Then
		Msg01 = "���~"
		Msg08 = "�M��(&C)"
		Msg09 = "Ū�J(&R)"
		Msg10 = "²�餤��½Ķ --> ���餤��½Ķ"
		Msg11 = "²�餤���� --> ���餤��½Ķ"
		Msg12 = "���餤��½Ķ --> ²�餤��½Ķ"
		Msg13 = "���餤���� --> ²�餤��½Ķ"
		Msg14 = "����Ū���r��A�i��ӻy�����r��M��S���r��Υ��Q��s�C"
	Else
		Msg01 = "����"
		Msg08 = "���(&C)"
		Msg09 = "����(&R)"
		Msg10 = "�������ķ��� --> �������ķ���"
		Msg11 = "��������Դ�� --> �������ķ���"
		Msg12 = "�������ķ��� --> �������ķ���"
		Msg13 = "��������Դ�� --> �������ķ���"
		Msg14 = "δ�ܶ�ȡ�ִ������ܸ����Ե��ִ��б�û���ִ���δ�����¡�"
	End If

	Code = DlgText("ExpCodeList")
	ConCmdID = DlgValue("CmdNameList")
	TypeListID = DlgValue("TypeListBox")
	FixListID = DlgValue("FixListBox")
	If DlgText("TypeListBox") = Msg10 Then ConTypeID = 0
	If DlgText("TypeListBox") = Msg11 Then ConTypeID = 1
	If DlgText("TypeListBox") = Msg12 Then ConTypeID = 2
	If DlgText("TypeListBox") = Msg13 Then ConTypeID = 3
	TempArray = Split(ConDataList(TypeListID),JoinStr)
	ConList = Split(TempArray(2),SubJoinStr)

	Select Case Action%
	Case 1
		LineNum = 20
		DlgText "LineNumBox",CStr(LineNum)
		DlgValue "CheckBox",1
		Sync = DlgValue("CheckBox")
		inText = ReadStrings(ConTypeID,LineNum,ConList)
    	DlgText "InTextBox",inText
    	If inText = "" Then DlgText "inTextBox",Msg14
    	If inText = "" Then DlgText "OutTextBox",""
    	If inText <> "" And inText <> Msg14 Then
    		DlgEnable "TestButton",True
    		DlgText "CleanButton",Msg08
    	Else
    		DlgText "OutTextBox",""
    		DlgEnable "TestButton",False
    		DlgText "CleanButton",Msg09
    	End If
    	If Sync = 1 And inText <> "" Then
    		outText = TestConvert(inText,ConTypeID,FixListID,Code,ConCmdID)
			DlgText "OutTextBox",outText
    	End If
	Case 2 ' ��ֵ���Ļ��߰����˰�ť
		If DlgItem$ = "CmdNameList" Then
			ConCmdID = DlgValue("CmdNameList")
			LineNum = CLng(DlgText("LineNumBox"))
			Sync = DlgValue("CheckBox")
			inText = DlgText("InTextBox")
    		If inText = "" Or inText = Msg14 Then DlgText "OutTextBox",""
    		If Sync = 1 And inText <> "" And inText <> Msg14 Then
    			DlgText "OutTextBox",""
    			outText = TestConvert(inText,ConTypeID,FixListID,Code,ConCmdID)
				DlgText "OutTextBox",outText
    		End If
		End If

		If DlgItem$ = "ExpCodeList" Then
			Code = DlgText("ExpCodeList")
			LineNum = CLng(DlgText("LineNumBox"))
			Sync = DlgValue("CheckBox")
			inText = DlgText("InTextBox")
    		If inText = "" Or inText = Msg14 Then DlgText "OutTextBox",""
    		If Sync = 1 And inText <> "" And inText <> Msg14 Then
    			DlgText "OutTextBox",""
    			outText = TestConvert(inText,ConTypeID,FixListID,Code,ConCmdID)
				DlgText "OutTextBox",outText
    		End If
		End If

		If DlgItem$ = "TypeListBox" Then
			LineNum = CLng(DlgText("LineNumBox"))
			Sync = DlgValue("CheckBox")
			TypeList = DlgText("TypeListBox")
			FixListID = DlgValue("FixListBox")
			If TypeList = Msg10 Then ConTypeID = 0
			If TypeList = Msg11 Then ConTypeID = 1
			If TypeList = Msg12 Then ConTypeID = 2
			If TypeList = Msg13 Then ConTypeID = 3
			inText = ReadStrings(ConTypeID,LineNum,ConList)
    		DlgText "InTextBox",inText
    		If inText = "" Then DlgText "inTextBox",Msg14
    		If inText = "" Then DlgText "OutTextBox",""
    		If Sync = 1 And inText <> "" And inText <> Msg14 Then
    			DlgText "OutTextBox",""
    			outText = TestConvert(inText,ConTypeID,FixListID,Code,ConCmdID)
				DlgText "OutTextBox",outText
    		End If
		End If

		If DlgItem$ = "FixListBox" Then
			FixListID = DlgValue("FixListBox")
			LineNum = CLng(DlgText("LineNumBox"))
			Sync = DlgValue("CheckBox")
			inText = DlgText("InTextBox")
    		If inText = "" Or inText = Msg14 Then DlgText "OutTextBox",""
    		If Sync = 1 And inText <> "" And inText <> Msg14 Then
    			DlgText "OutTextBox",""
    			outText = TestConvert(inText,ConTypeID,FixListID,Code,ConCmdID)
				DlgText "OutTextBox",outText
    		End If
		End If

		If DlgItem$ = "TestButton" Then
			LineNum = CLng(DlgText("LineNumBox"))
			Sync = DlgValue("CheckBox")
			inText = DlgText("InTextBox")
			If inText = "" Or inText = Msg14 Then DlgText "OutTextBox",""
    		If inText <> "" And inText <> Msg14 Then
    			DlgText "OutTextBox",""
    			outText = TestConvert(inText,ConTypeID,FixListID,Code,ConCmdID)
				DlgText "OutTextBox",outText
    		End If
		End If

		If DlgItem$ = "CleanButton" Then
			If DlgText("CleanButton") = Msg08 Then
				DlgText "InTextBox",""
				DlgText "OutTextBox",""
				DlgEnable "TestButton",False
				DlgText "CleanButton",Msg09
			Else
				LineNum = CLng(DlgText("LineNumBox"))
				Sync = DlgValue("CheckBox")
				inText = ReadStrings(ConTypeID,LineNum,ConList)
    			DlgText "InTextBox",inText
    			If inText = "" Then DlgText "inTextBox",Msg14
    			If inText = "" Then DlgText "OutTextBox",""
    			If inText <> "" And inText <> Msg14 Then
    				DlgEnable "TestButton",True
    				DlgText "CleanButton",Msg08
    			Else
    				DlgEnable "TestButton",False
    				DlgText "CleanButton",Msg09
    			End If
    			If Sync = 1 And inText <> "" Then
    				outText = TestConvert(inText,ConTypeID,FixListID,Code,ConCmdID)
					DlgText "OutTextBox",outText
    			End If
			End If
		End If

		If DlgItem$ = "EditButton" Then
			n = 0
			For i = LBound(ConCmdDataList) To UBound(ConCmdDataList)
				TempArray = Split(ConCmdDataList(i),JoinStr)
				If TempArray(14) <> "" Then
					ReDim Preserve FileList(n)
					FileList(n) = TempArray(14)
					If i = ConCmdID Then
						If ConTypeID = 2 Or ConTypeID = 3 Then File = TempArray(14)
					End If
					n = n + 1
				End If
				If TempArray(15) <> "" Then
					ReDim Preserve FileList(n)
					FileList(n) = TempArray(15)
					If i = ConCmdID Then
						If ConTypeID = 0 Or ConTypeID = 1 Then File = TempArray(15)
					End If
					n = n + 1
				End If
			Next i
			If File = "" Then File = FileList(0)
   			If File <> "" Then
   				x = ShowPopupMenu(AppNames,vbPopupVCenterAlign)
				OpenFile(File,FileList,x)
			End If
		End If

		If DlgItem$ <> "CancelButton" Then
		    inText = DlgText("InTextBox")
		    If inText = "" Or inText = Msg14 Then DlgText "OutTextBox",""
		    If inText <> "" And inText <> Msg14 Then
    			DlgEnable "TestButton",True
    			DlgText "CleanButton",Msg08
    		Else
    			DlgEnable "TestButton",False
    			DlgText "CleanButton",Msg09
    		End If
			ConCmdTestFunc = True '��ֹ���°�ť�رնԻ��򴰿�
		End If
	Case 3 ' �ı��������Ͽ��ı�������
		If DlgItem$ = "LineNumBox" Then
			LineNum = CLng(DlgText("LineNumBox"))
			Sync = DlgValue("CheckBox")
			DlgText "InTextBox",""
			inText = ReadStrings(ConTypeID,LineNum,ConList)
			DlgText "InTextBox",inText
			If inText = "" Then DlgText "inTextBox",Msg14
			If inText = "" Then DlgText "OutTextBox",""
    		If Sync = 1 And inText <> "" And inText <> Msg14 Then
    			DlgText "OutTextBox",""
    			outText = TestConvert(inText,ConTypeID,FixListID,Code,ConCmdID)
				DlgText "OutTextBox",outText
    		End If
		End If
		If DlgItem$ = "InTextBox" Then
			LineNum = CLng(DlgText("LineNumBox"))
			Sync = DlgValue("CheckBox")
			inText = DlgText("InTextBox")
    		If inText = "" Or inText = Msg14 Then DlgText "OutTextBox",""
    		If Sync = 1 And inText <> "" And inText <> Msg14 Then
    			DlgText "OutTextBox",""
    			outText = TestConvert(inText,ConTypeID,FixListID,Code,ConCmdID)
				DlgText "OutTextBox",outText
    		End If
		End If
		If DlgItem$ <> "OutTextBox" Then
			inText = DlgText("InTextBox")
			If inText = "" Or inText = Msg14 Then DlgText "OutTextBox",""
			If inText <> "" And inText <> Msg14 Then
    			DlgEnable "TestButton",True
    			DlgText "CleanButton",Msg08
    		Else
    			DlgEnable "TestButton",False
    			DlgText "CleanButton",Msg09
    		End If
    	End If
	End Select
End Function


'��ȡ�ִ��б��е��ִ�
Function ReadStrings(ConTypeID As Integer,LineNum As Integer,ConList() As String) As String
	Dim SrcTitle As String,TrnTitle As String,TestTitle As String,StringText As String
	Dim i As Integer,j As Integer,k As Integer,t As Integer,Stemp As Boolean
	Dim SrcString As PslSourceString,TrnString As PslTransString,Temp As String

	ReadStrings = ""
	k = 0
	If ConTypeID = 0 Or ConTypeID = 2 Then
		For i = 1 To Prj.TransLists.Count
			Set trn = Prj.TransLists(i)
			Stemp = False
			If ConTypeID = 0 And trn.Language.LangID = 2052 Then Stemp = True
			If ConTypeID = 2 And trn.Language.LangID = 1028 Then Stemp = True
			If Stemp = True And ConList(0) <> "" Then
				For j = LBound(ConList) To UBound(ConList)
					Temp = Left(ConList(j),InStr(ConList(j)," - ")-1)
					If trn.Title = Temp Then
						If trn.SourceList.LastChange > trn.LastUpdate Then trn.Update
						For t = 1 To trn.StringCount
							Set TrnString = trn.String(t)
							If TrnString.Text <> "" Then
								StringText = TrnString.Text
								If tText <> "" Then tText = tText & vbCrLf & StringText
		 						If tText = "" Then tText = StringText
								k = k + 1
							End If
							If k = LineNum Then Exit For
						Next t
					End If
					If k = LineNum Then Exit For
				Next j
			End If
			If k = LineNum Then Exit For
		Next i
		If tText = "" Then
			For i = 1 To Prj.TransLists.Count
				Set trn = Prj.TransLists(i)
				Stemp = False
				If ConTypeID = 0 And trn.Language.LangID = 2052 Then Stemp = True
				If ConTypeID = 2 And trn.Language.LangID = 1028 Then Stemp = True
				If Stemp = True Then
					For t = 1 To trn.StringCount
						Set TrnString = trn.String(t)
						If TrnString.Text <> "" Then
							StringText = TrnString.Text
							If tText <> "" Then tText = tText & vbCrLf & StringText
							If tText = "" Then tText = StringText
							k = k + 1
						End If
						If k = LineNum Then Exit For
					Next t
					If k = LineNum Then Exit For
				End If
			Next i
		End If
	ElseIf ConTypeID = 1 Or ConTypeID = 3 Then
		For i = 1 To Prj.SourceLists.Count
			Set Src = Prj.SourceLists(i)
			Stemp = False
			If ConTypeID = 1 And Src.LangID = 2052 Then Stemp = True
			If ConTypeID = 3 And Src.LangID = 1028 Then Stemp = True
			If Stemp = True And ConList(0) <> "" Then
				For j = LBound(ConList) To UBound(ConList)
					Temp = Left(ConList(j),InStr(ConList(j)," - ")-1)
					If Src.Title = Temp Then
						If Src.FileDate > Src.LastUpdate Then Src.Update
						For t = 1 To Src.StringCount
							Set SrcString = Src.String(t)
							If SrcString.Text <> "" Then
								StringText = SrcString.Text
	 							If tText <> "" Then tText = tText & vbCrLf & StringText
	 							If tText = "" Then tText = StringText
								k = k + 1
							End If
							If k = LineNum Then Exit For
						Next t
					End If
					If k = LineNum Then Exit For
				Next j
			End If
			If k = LineNum Then Exit For
		Next i
		If tText = "" Then
			For i = 1 To Prj.SourceLists.Count
				Set Src = Prj.SourceLists(i)
				Stemp = False
				If ConTypeID = 1 And Src.LangID = 2052 Then Stemp = True
				If ConTypeID = 3 And Src.LangID = 1028 Then Stemp = True
				If Stemp = True Then
					For t = 1 To Src.StringCount
						Set SrcString = Src.String(t)
						If SrcString.Text <> "" Then
							StringText = SrcString.Text
							If tText <> "" Then tText = tText & vbCrLf & StringText
							If tText = "" Then tText = StringText
							k = k + 1
						End If
						If k = LineNum Then Exit For
					Next t
					If k = LineNum Then Exit For
				End If
			Next i
		End If
	End If
	If tText <> "" Then ReadStrings = tText
	If tText = "" Then ReadStrings = ""
End Function


'�����ı�ת��
Function TestConvert(inText As String,ConID As Integer,FixID As Integer,Code,ConCmdID As Integer) As String
	Dim Argument As String,CodeID As Integer,objStream As Object
	Dim i As Integer,j As Integer,FixFileList() As String,BuiltInFile As String
	Dim ConArg As String,ConArgFix As String,FixPath As String,ReNameID As String
	If OSLanguage = "0404" Then
		Msg01 = "���~"
		Msg02 = "���b�ഫ�A�еy��..."
		Msg04 = "�ഫ���ѡI�ഫ�{���ΩR�O�C�ѼƳ]�w�i�঳���D�C"
		Msg05 = "�ഫ���ѡI�L�k�g�J�U�C�ɮסC�i��L�g�J�v���C" & vbCrLf
		Msg18 = "�L�k���s�R�W�U�C�ɮסA�нT�{�O�_�s�b�Υ��Q��L�{���ϥΡC" & vbCrLf
		Msg19 = "�L�k�٭�U�C�ɮסA�нT�{�ؼЦ�m�O�_���g�J�v���C" & vbCrLf
		Msg20 = "�L�k�ק��ഫ�{�����]�w�ɮסA�нT�{�ؼ��ɮ׬O�_���g�J�v���C" & vbCrLf
	Else
		Msg01 = "����"
		Msg02 = "����ת�������Ժ�..."
		Msg04 = "ת��ʧ�ܣ�ת������������в������ÿ��������⡣"
		Msg05 = "ת��ʧ�ܣ��޷�д�������ļ���������д��Ȩ�ޡ�" & vbCrLf
		Msg18 = "�޷������������ļ�����ȷ���Ƿ���ڻ�������������ʹ�á�" & vbCrLf
		Msg19 = "�޷���ԭ�����ļ�����ȷ��Ŀ��λ���Ƿ���д��Ȩ�ޡ�" & vbCrLf
		Msg20 = "�޷��޸�ת������������ļ�����ȷ��Ŀ���ļ��Ƿ���д��Ȩ�ޡ�" & vbCrLf
	End If

	TestConvert = ""
	TempArray = Split(ConCmdDataList(ConCmdID),JoinStr)
	CmdPath = TempArray(1)
	If Code = "ANSI" Then
		GBKToBig5 = TempArray(2)
		Big5ToGBK = TempArray(3)
		GBKToBig5Fix = TempArray(4)
		Big5ToGBKFix = TempArray(5)
	ElseIf Code = "Unicode" Then
		GBKToBig5 = TempArray(6)
		Big5ToGBK = TempArray(7)
		GBKToBig5Fix = TempArray(8)
		Big5ToGBKFix = TempArray(9)
	ElseIf Code = "UTF-8" Then
		GBKToBig5 = TempArray(10)
		Big5ToGBK = TempArray(11)
		GBKToBig5Fix = TempArray(12)
		Big5ToGBKFix = TempArray(13)
	End If
	GBKFixPath = TempArray(14)
	Big5FixPath = TempArray(15)
	RenGBKFix = TempArray(16)
	RenBig5Fix = TempArray(17)
	FixFileSlipStr = TempArray(18)
	FixFileSaveCode = TempArray(19)

	If ConID = 0 Or ConID = 1 Then
		ConArg = GBKToBig5
		ConArgFix = GBKToBig5Fix
		FixPath = Big5FixPath
		ReNameID = RenBig5Fix
		BuiltInFile = "bi_tradfix.dat"
	Else
		ConArg = Big5ToGBK
		ConArgFix = Big5ToGBKFix
		FixPath = GBKFixPath
		ReNameID = RenGBKFix
		BuiltInFile = "bi_simfix.dat"
	End If

	FixFile = FixPath
	If FixFileSlipStr <> "" Then
		FixFile = Replace(FixPath,SubJoinStr,Convert(FixFileSlipStr))
		FixFilePath = prjFolder & "\~temp.txt"
		If FixFileSaveCode <> "" Then
			WriteToFile(FixFilePath,FixFile,FixFileSaveCode)
			FixFile = FixFilePath
		End If
	End If

	If FixID = 0 Then Argument = GetArgument(ConArg,FixFile)
	If FixID = 1 Then Argument = GetArgument(ConArgFix,FixFile)

	If FixPath <> "" Then
		FixFileList = Split(FixPath,SubJoinStr)
		If ReNameID = "1" Then
			For j = LBound(FixFileList) To UBound(FixFileList)
				FixFile = FixFileList(j)
				If FixID = 1 And Dir(FixFile & ".bak") <> "" Then
					If FileRename(FixFile & ".bak",FixFile) = False Then
						MsgBox Msg19 & FixFile,vbOkOnly+vbInformation,Msg01
						Exit Function
					End If
				ElseIf FixID = 0 And Dir(FixFile) <> "" Then
					If RenameFixFile(FixFile,FixFile & ".bak") = False Then
						MsgBox Msg18 & FixFile,vbOkOnly+vbInformation,Msg01
						Exit Function
					End If
				End If
			Next j
		End If
		If InStr(LCase(CmdPath),"convertz.exe") Then
			If Not (UBound(FixFileList) = 0 And LCase(FixFileList(0)) = BuiltInFile) Then
				INIFile = Left(CmdPath,InStrRev(LCase(CmdPath),".exe")) & "ini"
				If Dir(INIFile) <> "" Then
					If FileRename(INIFile,INIFile & ".bak") = False Then
						MsgBox Msg18 & INIFile,vbOkOnly+vbInformation,Msg01
						GoTo RevertFile
					End If
				End If
				If ChangeCMDSetings(INIFile,INIFile & ".bak",FixPath,ConID) = False Then
					MsgBox Msg20 & INIFile,vbOkOnly+vbInformation,Msg01
					GoTo RevertFile
				End If
			End If
		End If
	End If

	On Error Resume Next
	If Dir(prjFolder & InputFile) <> "" Then Kill prjFolder & InputFile
	If Dir(prjFolder & OutputFile) <> "" Then Kill prjFolder & OutputFile
	On Error GoTo 0

	On Error GoTo RevertFile
	Set objStream = CreateObject("Adodb.Stream")
	If objStream Is Nothing Or Code = "ANSI" Then
		If ConID = 0 Then inText = PSL.ConvertUnicode2ASCII(inText,936)
		If ConID = 1 Then inText = PSL.ConvertUnicode2ASCII(inText,936)
		If ConID = 2 Then inText = PSL.ConvertUnicode2ASCII(inText,950)
		If ConID = 3 Then inText = PSL.ConvertUnicode2ASCII(inText,950)
		Open prjFolder & InputFile For Output As #1
			Print #1,inText;
		Close #1
	Else
		If WriteToFile(prjFolder & InputFile,inText,Code) = False Then
			MsgBox Msg05 & prjFolder & InputFile,vbOkOnly+vbInformation,Msg01
		End If
	End If
	On Error GoTo 0

	On Error GoTo ErrorConCmd
	Dim WshShell As Object
	Set WshShell = CreateObject("WScript.Shell")
	If WshShell Is Nothing Then
		ShellWait("""" & CmdPath & """ " & Argument)
	Else
		Return = WshShell.Run("""" & CmdPath & """ " & Argument, 0, True)
		Set WshShell = Nothing
		If Return <> 0 Then GoTo ErrorConCmd
	End If
	On Error GoTo 0

	On Error Resume Next
	If InStr(Argument,OutputFile) <> 0 Then
		If Dir(prjFolder & InputFile) <> "" Then Kill prjFolder & InputFile
	End If
	On Error GoTo 0

	If Not objStream Is Nothing Then
		If Dir(prjFolder & OutputFile) <> "" Then
			Code = CheckCode(prjFolder & OutputFile)
			For i = LBound(CodeList) To UBound(CodeList)
				TempArray = Split(CodeList(i),JoinStr)
				If TempArray(1) = Code Then
 					CodeID = i
 					Exit For
 				End If
			Next i
			If CodeID > 1 And CodeID < 39 Then Code = "ANSI"
			If Code <> "ANSI" Then TestConvert = ReadFile(prjFolder & OutputFile, Code)
		ElseIf Dir(prjFolder & InputFile) <> "" Then
			Code = CheckCode(prjFolder & InputFile)
			For i = LBound(CodeList) To UBound(CodeList)
				TempArray = Split(CodeList(i),JoinStr)
				If TempArray(1) = Code Then
 					CodeID = i
 					Exit For
 				End If
			Next i
			If CodeID > 1 And CodeID < 39 Then Code = "ANSI"
			If Code <> "ANSI" Then TestConvert = ReadFile(prjFolder & InputFile, Code)
		End If
	End If
	On Error GoTo DelFile
	If objStream Is Nothing Or Code = "ANSI" Then
		If Dir(prjFolder & OutputFile) <> "" Then
			Open prjFolder & OutputFile For Input As #1
		ElseIf Dir(prjFolder & InputFile) <> "" Then
			Open prjFolder & InputFile For Input As #1
		End If
		While Not EOF(1)
			Line Input #1,l$
			If TestConvert <> "" Then TestConvert = TestConvert & vbCrLf & l$
			If TestConvert = "" Then TestConvert = l$
		Wend
		Close #1
		If ConID = 0 Then TestConvert = PSL.ConvertASCII2Unicode(TestConvert,950)
		If ConID = 1 Then TestConvert = PSL.ConvertASCII2Unicode(TestConvert,950)
		If ConID = 2 Then TestConvert = PSL.ConvertASCII2Unicode(TestConvert,936)
		If ConID = 3 Then TestConvert = PSL.ConvertASCII2Unicode(TestConvert,936)
	End If
	Set objStream = Nothing
	On Error GoTo 0
	GoTo DelFile

	ErrorConCmd:
	MsgBox Msg04,vbOkOnly+vbInformation,Msg01

	DelFile:
	On Error Resume Next
	If Dir(prjFolder & InputFile) <> "" Then Kill prjFolder & InputFile
	If Dir(prjFolder & OutputFile) <> "" Then Kill prjFolder & OutputFile
	On Error GoTo 0

	RevertFile:
	If FixPath <> "" Then
		If ReNameID = "1" And FixID = 0 Then
			For j = LBound(FixFileList) To UBound(FixFileList)
				FixFile = FixFileList(j)
				If Dir(FixFile & ".bak") <> "" Then
					If FileRename(FixFile & ".bak",FixFile) = False Then
						MsgBox Msg19 & FixFile,vbOkOnly+vbInformation,Msg01
						Exit Function
					End If
				End If
			Next j
		End If
		If InStr(LCase(CmdPath),"convertz.exe") Then
			If Dir(INIFile & ".bak") <> "" Then
				If FileRename(INIFile & ".bak",INIFile) = False Then
					MsgBox Msg19 & INIFile,vbOkOnly+vbInformation,Msg01
				End If
			End If
		End If
	End If
	If FixFileSaveCode <> "" And FixFilePath <> "" Then
		On Error Resume Next
		If Dir(FixFilePath) <> "" Then Kill FixFilePath
		On Error GoTo 0
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
	ReDim CodePage(MaxNum - MinNum) As String
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


'����
Sub Help(HelpTip As String)
	If OSLanguage = "0404" Then
	AboutTitle = "����"
	HelpTitle = "����"
	HelpTipTitle = "Passolo ²�����ഫ����"
	AboutWindows = " ���� "
	MainWindows = " �D���� "
	SetWindows = " �]�w���� "
	TestWindows = " ���յ��� "
	Lines = "-----------------------"
	Sys = "�n�骩���G" & Version & vbCrLf & _
			"�A�Ψt�ΡGWindows XP/2000 �H�W�t��" & vbCrLf & _
			"�A�Ϊ����G�Ҧ��䴩�����B�z�� Passolo 5.0 �ΥH�W����" & vbCrLf & _
			"���v�Ҧ��G�~�Ʒs�@��" & vbCrLf & _
			"���v�Φ��G�K�O�n��" & vbCrLf & _
			"�x�譺���Ghttp://www.hanzify.org" & vbCrLf & _
			"�e�}�o�̡G�~�Ʒs�@������ gnatix (2007-2008)" & vbCrLf & _
			"��}�o�̡G�~�Ʒs�@������ wanfu (2009-2010)" & vbCrLf & vbCrLf & vbCrLf
	Ement = "���������ҡ�" & vbCrLf & _
			"============" & vbCrLf & _
			"- �䴩�����B�z�� Passolo 5.0 �ΥH�W�����A����" & vbCrLf & _
			"- Windows Script Host (WSH) ���� (VBS)�A����" & vbCrLf & _
			"- Windows ���O�X Adodb.Stream ���� (VBS)�A�䴩 Utf-8�BUnicode ����" & vbCrLf & _
			"- Microsoft.XMLHTTP ����A�䴩�۰ʧ�s�һ�" & vbCrLf & _
			"- ���Ӧ��}�o�� ConCmd 1.5 �� ConvertZ 8.02 �Ψ�L�䴩�R�O�C��²���ഫ�{���A����" & vbCrLf & vbCrLf & vbCrLf
	Dec = "���n��²����" & vbCrLf & _
			"============" & vbCrLf & _
			"�Ω� Passolo �r��M�檺²����ۤ��ഫ�C���㦳�H�U�\��G" & vbCrLf & _
			"- �䴩�q²�������²����½Ķ���ഫ" & vbCrLf & _
			"- �䴩²����½Ķ�������ۤ��ഫ" & vbCrLf & _
			"- �P���Ӧ��}�o�� ConCmd 1.5 �� ConvertZ 8.02 ��K��X" & vbCrLf & _
			"- �䴩�ۭq²���ഫ�{���Ψ�R�O�C�Ѽ�" & vbCrLf & _
			"- �䴩�Ҧ��r���s�X�A�æ۰ʿ���" & vbCrLf & _
			"- �����ഫ�{�������աC�i�H�ߧY�A���ഫ�{�����]�w�O�_���T" & vbCrLf & _
			"- �i�ϥΤ��m�{���B�O�ƥ��B�t�ιw�]�{���B�ۭq�{���s����J�ץ��ɮ�" & vbCrLf & _
			"- ���m�i�ۭq���۰ʧ�s�\��" & vbCrLf & vbCrLf & _
			"���{���]�t�U�C�ɮסG" & vbCrLf & _
			"- PSLGbk2Big5.bas (�����ɮ�)" & vbCrLf & _
			"- PSLGbk2Big5.txt (²�餤�廡���ɮ�)" & vbCrLf & _
			"- ConCmd1.5.rar (���Ӧ��}�o�A�W�K�M�ץ��F²����λy�ץ���)" & vbCrLf & _
			"- ConvertZ8.02.rar (���Ӧ��}�o�A�W�K�M�ץ��F²����λy�ץ���)" & vbCrLf & vbCrLf & vbCrLf
	Setup = "���w�ˤ�k��" & vbCrLf & _
			"============" & vbCrLf & _
			"- �p�G�ϥΤF Wanfu �� Passolo �~�ƪ��A�æw�ˤF���[�����ե�A�h�G" & vbCrLf & _
			"  1) �N�����᪺�����ɮפ��O�����w�˵{���ؿ��U�� Macros ��Ƨ��M Passolo �t�θ�Ƨ���" & vbCrLf & _
            "     �w�q�� Macros ��Ƨ�����Ӫ��ɮסC" & vbCrLf & _
			"  2) �N���Ӧ��}�o���ഫ�{������²����λy�ץ��� DAT �ɮ׽ƻs���ഫ�{���Ҧb��Ƨ��C" & vbCrLf & _
            "- �p�G�ϥΤF��L Passolo �����A�h�G" & vbCrLf & _
			"  1) �N�����᪺�ɮ׽ƻs�� Passolo �t�θ�Ƨ����w�q�� Macros ��Ƨ���" & vbCrLf & _
			"  2) �b Passolo ���u�� -> �ۭq�u���椤�s�W���ɮרéw�q�ӿ��W�١A" & vbCrLf & _
			"     ����N�i�H�I���ӿ�檽���I�s" & vbCrLf & _
			"  3) �w�˥ѧ��Ӧ��}�o���ഫ�{���A�îھڰ���ɼu�X����ܤ���]�w���m�C" & vbCrLf & vbCrLf & vbCrLf
	CopyRight = "�����v�ŧi��" & vbCrLf & _
			"============" & vbCrLf & _
			"- ���n�骺���v�k�}�o�̩M�ק�̩Ҧ��A����H�i�H�K�O�ϥΡB�ק�B�ƻs�B���G���n��C" & vbCrLf & _
			"- �ק�B���G���n�饲���H���������ɮסA�õ����n���l�}�o�̥H�έק�̡C" & vbCrLf & _
			"- ���g�}�o�̩M�ק�̦P�N�A�����´�έӤH�A���o�Ω�ӷ~�n��B�ӷ~�άO�䥦��Q�ʬ��ʡC" & vbCrLf & _
			"- ��ϥΥ��n�骺��l�����A�H�Ψϥθg�L�H�ק諸�D��l�����ҳy�����l���M�l�`�A�}�o�̤�" & vbCrLf & _
			"  �Ӿ����d���C" & vbCrLf & _
			"- �ѩ󬰧K�O�n��A�}�o�̩M�ק�̨S���q�ȴ��ѳn��޳N�䴩�A�]�L�q�ȧ�i�Χ�s�����C" & vbCrLf & _
			"- �w������~�~�ô��X��i�N���C�p�����~�Ϋ�ĳ�A�жǰe��: z_shangyi@163.com�C" & vbCrLf & vbCrLf & vbCrLf
	Thank = "���P�@�@�¡�" & vbCrLf & _
			"============" & vbCrLf & _
			"- �P�§��Ӧ��}�o�F ConCmd 1.5 �� ConvertZ 8.02 �u�q���ഫ�n��I" & vbCrLf & _
			"- �P�¥x�W�~�ƬɤͤH�����͡B�~�Ʒs�@���|�������o���ʹ��ѤF�ץ���I" & vbCrLf & _
			"- �P�º~�Ʒs�@���|�������i���ͩM Heaven ���ʹ��X����λy�ץ��N���I" & vbCrLf & _
			"- �P�º~�Ʒs�@���|�����ըô��X�ק�N���I" & vbCrLf & vbCrLf & vbCrLf
	Contact = "���P���pô��" & vbCrLf & _
			"============" & vbCrLf & _
			"wanfu�Gz_shangyi@163.com" & vbCrLf & vbCrLf & vbCrLf
	Logs = "�P�¤䴩�I�z���䴩�O�ڳ̤j���ʤO�I�P���w��ϥΧڭ̻s�@���n��I" & vbCrLf & vbCrLf & _
			"**************************************************" & vbCrLf & _
			"�ݭn��h�B��s�B��n���~�ơA�Ы��X:" & vbCrLf & _
			"�~�Ʒs�@�� -- http://www.hanzify.org" & vbCrLf & _
			"�~�Ʒs�@���׾� -- http://bbs.hanzify.org" & vbCrLf & _
			"**************************************************" & vbCrLf
	Else
	AboutTitle = "����"
	HelpTitle = "����"
	HelpTipTitle = "Passolo ����ת����"
	AboutWindows = " ���� "
	MainWindows = " ������ "
	SetWindows = " ���ô��� "
	TestWindows = " ���Դ��� "
	Lines = "-----------------------"
	Sys = "����汾��" & Version & vbCrLf & _
			"����ϵͳ��Windows XP/2000 ����ϵͳ" & vbCrLf & _
			"���ð汾������֧�ֺ괦��� Passolo 5.0 �����ϰ汾" & vbCrLf & _
			"��Ȩ���У�����������" & vbCrLf & _
			"��Ȩ��ʽ��������" & vbCrLf & _
			"�ٷ���ҳ��http://www.hanzify.org" & vbCrLf & _
			"ǰ�����ߣ����������ͳ�Ա gnatix (2007-2008)" & vbCrLf & _
			"�󿪷��ߣ����������ͳ�Ա wanfu (2009-2010)" & vbCrLf & vbCrLf & vbCrLf
	Ement = "�����л�����" & vbCrLf & _
			"============" & vbCrLf & _
			"- ֧�ֺ괦��� Passolo 5.0 �����ϰ汾������" & vbCrLf & _
			"- Windows Script Host (WSH) ���� (VBS)������" & vbCrLf & _
			"- Windows �ű� Adodb.Stream ���� (VBS)��֧�� Utf-8��Unicode ����" & vbCrLf & _
			"- Microsoft.XMLHTTP ����֧���Զ���������" & vbCrLf & _
			"- ��־�ɿ����� ConCmd 1.5 �� ConvertZ 8.02 ������֧�������еļ�ת�����򣬱���" & vbCrLf & vbCrLf & vbCrLf
	Dec = "���������" & vbCrLf & _
			"============" & vbCrLf & _
			"���� Passolo �ִ��б�ļ����໥ת�������������¹��ܣ�" & vbCrLf & _
			"- ֧�ִӼ���ԭ�ĵ����巭���ת��" & vbCrLf & _
			"- ֧�ּ��巭��֮����໥ת��" & vbCrLf & _
			"- ����־�ɿ����� ConCmd 1.5 �� ConvertZ 8.02 ���ܼ���" & vbCrLf & _
			"- ֧���Զ����ת�������������в���" & vbCrLf & _
			"- ֧�������ַ����룬���Զ�ʶ��" & vbCrLf & _
			"- �ṩת������Ĳ��ԡ����������˽�ת������������Ƿ���ȷ" & vbCrLf & _
			"- ��ʹ�����ó��򡢼��±���ϵͳĬ�ϳ����Զ������༭�ʻ������ļ�" & vbCrLf & _
			"- ���ÿ��Զ�����Զ����¹���" & vbCrLf & vbCrLf & _
			"��������������ļ���" & vbCrLf & _
			"- PSLGbk2Big5.bas (���ļ�)" & vbCrLf & _
			"- PSLGbk2Big5.txt (��������˵���ļ�)" & vbCrLf & _
			"- ConCmd1.5.rar (��־�ɿ���������������˼�������������)" & vbCrLf & _
			"- ConvertZ8.02.rar (��־�ɿ���������������˼�������������)" & vbCrLf & vbCrLf & vbCrLf
	Setup = "�װ������" & vbCrLf & _
			"============" & vbCrLf & _
			"- ���ʹ���� Wanfu �� Passolo �����棬����װ�˸��Ӻ��������" & vbCrLf & _
			"  1) ����ѹ��ĺ��ļ��ֱ��滻��װ����Ŀ¼�µ� Macros �ļ��к� Passolo ϵͳ�ļ�����" & vbCrLf & _
            "     ����� Macros �ļ�����ԭ�����ļ���" & vbCrLf & _
			"  2) ����־�ɿ�����ת�������еļ������������� DAT �ļ����Ƶ�ת�����������ļ��С�" & vbCrLf & _
            "- ���ʹ�������� Passolo �汾����" & vbCrLf & _
			"  1) ����ѹ����ļ����Ƶ� Passolo ϵͳ�ļ����ж���� Macros �ļ�����" & vbCrLf & _
			"  2) �� Passolo �Ĺ��� -> �Զ��幤�߲˵�����Ӹ��ļ�������ò˵����ƣ�" & vbCrLf & _
			"     �˺�Ϳ��Ե����ò˵�ֱ�ӵ���" & vbCrLf & _
			"  3) ��װ����־�ɿ�����ת�����򣬲���������ʱ�����ĶԻ���������λ�á�" & vbCrLf & vbCrLf & vbCrLf
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
			"- ��л��־�ɿ����� ConCmd 1.5 �� ConvertZ 8.02 �����ת�������" & vbCrLf & _
			"- ��л̨�庺�������˸����������������ͻ�Ա��˹�������ṩ��������" & vbCrLf & _
			"- ��л���������ͻ�Ա���չ������ Heaven ������������������������" & vbCrLf & _
			"- ��л���������ͻ�Ա���Բ�����޸������" & vbCrLf & vbCrLf & vbCrLf
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
