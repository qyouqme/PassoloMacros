''This Macro is to copy right-side'strings of the like INI file to left-side
''Idea and implemented by wanfu 2009.12.19 (Last modified on 2012.02.10)
''Support all encodings
''-----------------------------------------------------------------------------------------
''
Public OSLanguage As String,Prj As PslProject,TempFile As String,CovertTypeList() As String
Public CovertFileList() As String,CovertDataList() As String,CovertDataListBak() As String

Public UIFileList() As String,UIDataList() As String,UILangList() As String,LangFile As String
Public UpdateSet() As String,UpdateSetBak() As String,WriteLoc As String,Selected() As String

Public FileText As String,FindText As String,FindLine As String,CodeList() As String
Public AppNames() As String,AppPaths() As String,FileDataList() As String

Private Const Version = "2012.02.10"
Private Const Extension = "ini;lng"
Private Const Separator = "="
Private Const JoinStr = vbFormFeed  'vbBack
Private Const SubJoinStr = vbVerticalTab  'Chr$(1)
Private Const RegKey = "HKCU\Software\VB and VBA Program Settings\CovertINIFile"

Private Const DefaultObject = "Microsoft.XMLHTTP;Msxml2.XMLHTTP"
Private Const updateAppName = "CovertINIFile"
Private Const updateMainFile = "CovertINIFile.bas"
Private Const updateINIFile = "PSLMacrosUpdates.ini"
Private Const updateMethod = "GET"
Private Const updateINIMainUrl = "ftp://hhdown:886@czftp.hanzify.org/update/PSLMacrosUpdates.ini"
Private Const updateINIMinorUrl = "http://www.wanfutrade.com/software/hanhua/PSLMacrosUpdates.rar"
Private Const updateMainUrl = "ftp://hhdown:886@czftp.hanzify.org/download/CovertINIFile.rar"
Private Const updateMinorUrl = "http://www.wanfutrade.com/software/hanhua/CovertINIFile.rar"
Private Const updateAsync = False


' 主程序
Sub Main
	Dim WshShell As Object,objStream As Object,MsgList() As String
	Dim i As Long,j As Long,n As Long,CodeNameList() As String,Temp As String,TempList() As String

	'检测系统语言
	On Error Resume Next
	Set WshShell = CreateObject("WScript.Shell")
	If WshShell Is Nothing Then
		MsgBox(Err.Description & " - " & "WScript.Shell",vbInformation)
		Exit Sub
	End If
	strKeyPath = "HKLM\SYSTEM\CurrentControlSet\Control\Nls\Language\Default"
	OSLanguage = WshShell.RegRead(strKeyPath)
	If OSLanguage = "" Then
		strKeyPath = "HKLM\SYSTEM\CurrentControlSet\Control\Nls\Language\InstallLanguage"
		OSLanguage = WshShell.RegRead(strKeyPath)
		If Err.Source = "WshShell.RegRead" Then
			MsgBox(Err.Description,vbInformation)
			Exit Sub
		End If
	End If

	'检测 Adodb.Stream 是否存在
	Set objStream = CreateObject("Adodb.Stream")
	If objStream Is Nothing Then
		MsgBox(Err.Description & " - " & "Adodb.Stream",vbInformation)
		Exit Sub
	End If
	On Error GoTo SysErrorMsg

	'初始化数组
	ReDim UIFileList(0),UIDataList(0),UILangList(0),Selected(0),UpdateSet(5)
	UIFileList(0) = "Auto"
	UIDataList(0) = "Auto" & JoinStr & "0" & JoinStr

	'读取设置
	getSettings("",UpdateSet)

	'读取界面语言字串
	If GetUIList(UIFileList,UIDataList) = True Then
		If Join(Selected,"") <> "" Then UILangID = LCase(Selected(0))
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
	If PSL.Version < 500 Then
		MsgBox MsgList(36),vbOkOnly+vbInformation,MsgList(35)
		Exit Sub
	End If

	'获取字符代码列表
	If objStream Is Nothing Then CodeList = CodePageList(0,0)
	If Not objStream Is Nothing Then CodeList = CodePageList(0,49)
	Set objStream = Nothing
	For i = LBound(CodeList) To UBound(CodeList)
		ReDim Preserve CodeNameList(i)
		TempArray = Split(CodeList(i),JoinStr)
		CodeNameList(i) = TempArray(0)
	Next i

	'初始化数组
	ReDim AppNames(3),AppPaths(3),FileDataList(0)
	ReDim CovertTypeList(7),CovertFileList(0),CovertDataList(0),CovertDataListBak(0)
	CovertTypeList(0) = MsgList(37)
	CovertTypeList(1) = MsgList(38)
	CovertTypeList(2) = MsgList(39)
	CovertTypeList(3) = MsgList(40)
	CovertTypeList(4) = MsgList(51)
	CovertTypeList(5) = MsgList(52)
	CovertTypeList(6) = MsgList(53)
	CovertTypeList(7) = MsgList(54)
	AppNames(0) = MsgList(41)
   	AppNames(1) = MsgList(42)
	AppNames(2) = MsgList(43)
   	AppNames(3) = MsgList(44)
   	AppPaths(1) = "notepad.exe"

   	'获取更新数据并检查新版本
	If Join(UpdateSet,"") <> "" Then
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
			getCMDPath(".rar",UpdateSet(2),UpdateSet(3))
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
		j = 0
		n = 0
		If updateDate <> "" Then
			i = CLng(DateDiff("d",CDate(updateDate),Date))
			j = StrComp(Format(Date,"yyyy-MM-dd"),updateDate)
			If updateCycle <> "" Then n = i - CLng(updateCycle)
		End If
		If updateDate = "" Or (j = 1 And n >= 0) Then
			i = Download(updateMethod,updateUrl,updateAsync,updateMode)
			If i > 0 Then
				If UpdateSet(5) < Format(Date,"yyyy-MM-dd") Then
					UpdateSet(5) = Format(Date,"yyyy-MM-dd")
					WriteSettings(UpdateSet,"Update")
				End If
				If i = 3 Then Exit Sub
			End If
		End If
	End If

	Begin Dialog UserDialog 660,518,MsgList(1),.MainDlgFunc ' %GRID:10,7,1,1
		Text 10,7,640,14,MsgList(0) & Version,.Text1,2
		Text 10,28,640,35,MsgList(2),.Text2

		GroupBox 10,70,640,245,MsgList(3),.SelectGroupBox
		ListBox 20,91,500,161,CovertFileList(),.CovertFileList,3
		PushButton 530,91,110,21,MsgList(4),.AddButton
		PushButton 530,112,110,21,MsgList(5),.BatchAddButton
		PushButton 530,140,110,21,MsgList(6),.ChangeButton
		PushButton 530,161,110,21,MsgList(7),.DelButton
		PushButton 530,182,110,21,MsgList(8),.ClearButton
		PushButton 530,259,110,21,MsgList(46),.PreViewButton
		Text 20,262,130,14,MsgList(45),.CodeText
		DropListBox 160,259,360,21,CodeNameList(),.CodeNameList
		Text 20,290,240,14,MsgList(9),.ExtensionText
		TextBox 270,287,160,21,.ExtNameList
		CheckBox 436,287,200,21,MsgList(10),.SubFolder

		GroupBox 10,322,640,42,MsgList(11),.ConfigGroupBox
		Text 40,339,80,14,MsgList(12),.CovertTypeText
		DropListBox 130,336,190,21,CovertTypeList(),.CovertType
		Text 340,339,90,14,MsgList(13),.StartLineText
		TextBox 440,336,50,21,.StartLine
		Text 500,339,70,14,MsgList(14),.BreakText
		TextBox 580,336,50,21,.Break

		GroupBox 10,371,640,84,MsgList(15),.AdditionGroupBox
		Text 40,388,80,14,MsgList(16),.PreAddition
		Text 130,388,80,14,MsgList(17),.PreAddStrText
		TextBox 220,385,150,21,.PreAddStr
		Text 380,388,80,14,MsgList(18),.PreAppStrText
		TextBox 470,385,160,21,.PreAppStr
		Text 40,409,80,14,MsgList(19),.AppAddition
		Text 130,409,80,14,MsgList(20),.AppAddStrText
		TextBox 220,406,150,21,.AppAddStr
		Text 380,409,80,14,MsgList(21),.AppAppStrText
		TextBox 470,406,160,21,.AppAppStr

		Text 40,430,80,14,MsgList(22),.Replaces
		Text 130,430,80,14,MsgList(23),.RepStrText
		TextBox 220,427,150,21,.RepStr
		Text 380,430,80,14,MsgList(24),.RepAsStrText
		TextBox 470,427,160,21,.RepAsStr

		CheckBox 20,462,410,14,MsgList(25),.Backup
		CheckBox 440,462,200,14,MsgList(34),.AllSame

		PushButton 10,490,90,21,MsgList(26),.AboutButton
		PushButton 100,490,90,21,MsgList(31),.SetButton
		PushButton 190,490,90,21,MsgList(27),.TestButton
		PushButton 280,490,90,21,MsgList(28),.AllEditButton
		PushButton 380,490,90,21,MsgList(29),.AllRestoreButton
		PushButton 530,210,110,21,MsgList(32),.CovertButton
		PushButton 530,231,110,21,MsgList(33),.RestoreButton
		PushButton 470,490,90,21,MsgList(30),.AllCovertButton
		CancelButton 560,490,90,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	If Dialog(dlg) = 0 Then Exit Sub

	'显示程序错误消息
	SysErrorMsg:
	If Err.Source <> "ExitSub" Then Call sysErrorMassage(Err,0)
End Sub


' 主对话框函数
Private Function MainDlgFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Long,m As Long,n As Long,sf As Long,Stemp As Boolean
	Dim Folder As String,Code As String,File As String,Temp As String
	Dim FilesArray() As String,CovertFileListBak() As String,MsgList() As String
	Dim TempArray() As String,TempList() As String

	If Action% < 3 Then
		If getMsgList(UILangList,MsgList,"MainDlgFunc",0) = False Then
			MainDlgFunc = True '防止按下按钮关闭对话框窗口
			Exit Function
		End If
	End If

	Select Case Action%
	Case 1 ' 对话框窗口初始化
		DlgValue "SubFolder",1
		DlgText "SelectGroupBox",MsgList(7) & "(0)"
		DlgText "ExtNameList",Extension
		DlgValue "CodeNameList",1
		DlgValue "CovertType",0
		DlgText "StartLine","1"
		DlgText "Break",Separator
		DlgValue "Backup",1
		DlgValue "AllSame",0
		DlgEnable "CodeNameList",False
		DlgEnable "CovertType",False
		DlgEnable "StartLine",False
		DlgEnable "Break",False
		DlgEnable "PreAddStr",False
		DlgEnable "PreAppStr",False
		DlgEnable "AppAddStr",False
		DlgEnable "AppAppStr",False
		DlgEnable "RepStr",False
		DlgEnable "RepAsStr",False
		DlgEnable "ChangeButton",False
		DlgEnable "DelButton",False
		DlgEnable "ClearButton",False
		DlgEnable "PreViewButton",False
		DlgEnable "AllCovertButton",False
		DlgEnable "CovertButton",False
		DlgEnable "TestButton",False
		DlgEnable "AllEditButton",False
		DlgEnable "AllRestoreButton",False
		DlgEnable "RestoreButton",False
	Case 2 ' 数值更改或者按下了按钮
		If DlgItem$ = "AboutButton" Then
			Call Help("About")
			MainDlgFunc = True '防止按下按钮关闭对话框窗口
			Exit Function
		End If

		If DlgItem$ = "SetButton" Then
			UpdateSetBak = UpdateSet
			UILangID = Selected(0)
			Temp = Join(UIDataList,JoinStr)
			Call Settings(0)

			Stemp = False
			If UILangID <> Selected(0) Then Stemp = True
			If Temp <> Join(UIDataList,JoinStr) Then Stemp = True
			If Stemp = True Then
				If getMsgList(UILangList,MsgList,"Main",1) = False Then
					MainDlgFunc = True ' 防止按下按钮关闭对话框窗口
					Exit Function
				End If

				'重置编辑菜单名称
				ReDim AppNames(3)
   				AppNames(0) = MsgList(41)
   				AppNames(1) = MsgList(42)
				AppNames(2) = MsgList(43)
   				AppNames(3) = MsgList(44)

				'重置对话框字串
				DlgText -1,MsgList(1)
				DlgText "Text1",MsgList(0) & Version
				DlgText "Text2",MsgList(2)

				DlgText "SelectGroupBox",MsgList(3)
				DlgText "AddButton",MsgList(4)
				DlgText "BatchAddButton",MsgList(5)
				DlgText "ChangeButton",MsgList(6)
				DlgText "DelButton",MsgList(7)
				DlgText "ClearButton",MsgList(8)
				DlgText "PreViewButton",MsgList(46)
				DlgText "CodeText",MsgList(45)
				DlgText "ExtensionText",MsgList(9)
				DlgText "SubFolder",MsgList(10)

				DlgText "ConfigGroupBox",MsgList(11)
				DlgText "CovertTypeText",MsgList(12)
				DlgText "StartLineText",MsgList(13)
				DlgText "BreakText",MsgList(14)

				DlgText "AdditionGroupBox",MsgList(15)
				DlgText "PreAddition",MsgList(16)
				DlgText "PreAddStrText",MsgList(17)
				DlgText "PreAppStrText",MsgList(18)
				DlgText "AppAddition",MsgList(19)
				DlgText "AppAddStrText",MsgList(20)
				DlgText "AppAppStrText",MsgList(21)

				DlgText "Replaces",MsgList(22)
				DlgText "RepStrText",MsgList(23)
				DlgText "RepAsStrText",MsgList(24)

				DlgText "Backup",MsgList(25)
				DlgText "AllSame",MsgList(34)

				DlgText "AboutButton",MsgList(26)
				DlgText "SetButton",MsgList(31)
				DlgText "TestButton",MsgList(27)
				DlgText "AllEditButton",MsgList(28)
				DlgText "AllRestoreButton",MsgList(29)
				DlgText "CovertButton",MsgList(32)
				DlgText "RestoreButton",MsgList(33)
				DlgText "AllCovertButton",MsgList(30)

				n = DlgValue("CodeNameList")
				For i = LBound(CodeList) To UBound(CodeList)
					ReDim Preserve TempList(i)
					TempArray = Split(CodeList(i),JoinStr)
					TempList(i) = TempArray(0)
				Next i
				DlgListBoxArray "CodeNameList",TempList()
				DlgValue "CodeNameList",n

				n = DlgValue("CovertType")
				CovertTypeList(0) = MsgList(37)
				CovertTypeList(1) = MsgList(38)
				CovertTypeList(2) = MsgList(39)
				CovertTypeList(3) = MsgList(40)
				CovertTypeList(4) = MsgList(51)
				CovertTypeList(5) = MsgList(52)
				CovertTypeList(6) = MsgList(53)
				CovertTypeList(7) = MsgList(54)
				DlgListBoxArray "CovertType",CovertTypeList()
				DlgValue "CovertType",n
			End If
			MainDlgFunc = True '防止按下按钮关闭对话框窗口
			Exit Function
		End If

		If DlgItem$ = "CovertFileList" Then
			File = DlgText("CovertFileList")
			For i = LBound(CovertFileList) To UBound(CovertFileList)
				If CovertFileList(i) = File Then
					TempList = Split(CovertDataList(i),JoinStr)
					Code = TempList(1)
					Exit For
				End If
			Next i
			If Code <> "" Then
				For i = LBound(CodeList) To UBound(CodeList)
					TempArray = Split(CodeList(i),JoinStr)
					If TempArray(1) = Code Then
						DlgValue "CodeNameList",i
						Exit For
					End If
				Next i
				DlgValue "CovertType",TempList(2)
				DlgText "StartLine",TempList(3)
				DlgText "Break",TempList(4)
				DlgText "PreAddStr",TempList(5)
				DlgText "PreAppStr",TempList(6)
				DlgText "AppAddStr",TempList(7)
				DlgText "AppAppStr",TempList(8)
				DlgText "RepStr",TempList(9)
				DlgText "RepAsStr",TempList(10)
			Else
				DlgValue "CodeNameList",1
			End If
			If File <> TempFile Then TempFile = ""
		End If

		If DlgItem$ = "AddButton" Or DlgItem$ = "BatchAddButton" Then
			ExtList = DlgText("ExtNameList")
			If InStr(ExtList,";") Then
				ExtName = "*." & Replace(ExtList,";","; *.")
			Else
				ExtName = "*." & ExtList
			End If

			Stemp = False
			If DlgItem$ = "AddButton" Then
				MsgList(3) = Replace(MsgList(3),"INI",UCase(Replace(ExtList,";","、")))
				MsgList(3) = Replace(MsgList(3),"*.ini",ExtName)
				If PSL.SelectFile(File,True,MsgList(3),MsgList(1)) = True Then
					ReDim FilesArray(0)
					FilesArray(0) = File
					Stemp = True
				End If
			Else
				If PSL.SelectFolder(Folder,Msg03) = True Then
					Folder = AppendBackslash(Folder,"","\",1)
					sf = DlgValue("SubFolder")
					FilesArray = GetFiles(Folder,ExtList,sf)
					If Join(FilesArray,"") = "" Then
						If sf = 0 Then
							MsgList(4) = Replace(MsgList(4),"INI",UCase(Replace(ExtList,";","、")))
							MsgBox(MsgList(4),vbOkOnly+vbInformation,MsgList(0))
						Else
							MsgList(5) = Replace(MsgList(5),"INI",UCase(Replace(ExtList,";","、")))
							MsgBox(MsgList(5),vbOkOnly+vbInformation,MsgList(0))
						End If
					Else
						Stemp = True
					End If
				End If
			End If

			If Stemp = True Then
				Files = Join(CovertFileList,JoinStr)
				For i = LBound(FilesArray) To UBound(FilesArray)
					File = FilesArray(i)
					If InStr(Files,File) Then
						MsgBox(File & vbCrLf & MsgList(6),vbOkOnly+vbInformation,MsgList(0))
					Else
						If CovertFileList(0) = "" Or DlgValue("AllSame") = 0 Then
							TempFile = File
							DlgValue "CovertType",0
							DlgText "StartLine","1"
							DlgText "Break",Separator
							DlgText "PreAddStr",""
							DlgText "PreAppStr",""
							DlgText "AppAddStr",""
							DlgText "AppAppStr",""
							DlgText "RepStr",""
							DlgText "RepAsStr",""
						End If
						Code = CheckCode(File)
						For n = LBound(CodeList) To UBound(CodeList)
							TempArray = Split(CodeList(n),JoinStr)
							If TempArray(1) = Code Then
 								DlgValue "CodeNameList",n
 								Exit For
 							End If
						Next n
						ReDim TempList(9)
						TempList(0) = Code
						TempList(1) = DlgValue("CovertType")
						TempList(2) = DlgText("StartLine")
						TempList(3) = DlgText("Break")
						TempList(4) = DlgText("PreAddStr")
						TempList(5) = DlgText("PreAppStr")
						TempList(6) = DlgText("AppAddStr")
						TempList(7) = DlgText("AppAppStr")
						TempList(8) = DlgText("RepStr")
						TempList(9) = DlgText("RepAsStr")
						Temp = File & JoinStr & Join(TempList,JoinStr)
						CovertFileListBak = CovertFileList
						CreateArray(File,Temp,CovertFileList,CovertDataList)
						CreateArray(File,Temp,CovertFileListBak,CovertDataListBak)
					End If
				Next i
				DlgListBoxArray "CovertFileList",CovertFileList()
				DlgText "CovertFileList",File
			End If
		End If

		If DlgItem$ = "ChangeButton" Then
			If DlgText("CovertFileList") <> "" Then
				ExtList = DlgText("ExtNameList")
				If InStr(ExtList,";") Then
					ExtName = "*." & Replace(ExtList,";","; *.")
				Else
					ExtName = "*." & ExtList
				End If
				MsgList(3) = Replace(MsgList(3),"INI",UCase(Replace(ExtList,";","、")))
				MsgList(3) = Replace(MsgList(3),"*.ini",ExtName)
				If PSL.SelectFile(TempFile,True,MsgList(3),MsgList(1)) = True Then
					Files = Join(CovertFileList,JoinStr)
					If InStr(Files,TempFile) Then
						MsgBox(TempFile & vbCrLf & MsgList(10),vbOkOnly+vbInformation,MsgList(0))
					Else
						File = DlgText("CovertFileList")
						Stemp = False
						For i = LBound(CovertFileList) To UBound(CovertFileList)
							If CovertFileList(i) = File Then
								Stemp = True
								Exit For
							End If
						Next i
						If Stemp = True Then
							CovertFileList(i) = TempFile
							TempList = Split(CovertDataList(i),JoinStr)
							TempList(0) = TempFile
							CovertDataList(i) = Join(TempList,JoinStr)
							TempList = Split(CovertDataListBak(i),JoinStr)
							TempList(0) = TempFile
							CovertDataListBak(i) = Join(TempList,JoinStr)
							DlgListBoxArray "CovertFileList",CovertFileList()
							DlgText "CovertFileList",TempFile
							TempFile = ""
						End If
					End If
				End If
			End If
		End If

		If DlgItem$ = "DelButton" And DlgText("CovertFileList") <> "" Then
			File = DlgText("CovertFileList")
			If MsgBox(MsgList(9) & vbCrLf & File,vbYesNo+vbInformation,MsgList(8)) = vbYes Then
				i = DlgValue("CovertFileList")
				CovertFileList = DelArray(File,CovertFileList,"",0)
				CovertDataList = DelArray(File,CovertDataList,JoinStr,0)
				CovertDataListBak = DelArray(File,CovertDataListBak,JoinStr,0)
				DlgListBoxArray "CovertFileList",CovertFileList()
				If i > UBound(CovertFileList) Then i = UBound(CovertFileList)
				If Join(CovertFileList,"") <> "" Then
					DlgValue "CovertFileList",i
					TempList = Split(CovertDataList(i),JoinStr,-1)
					If TempList(1) <> "" Then
						For i = LBound(CodeList) To UBound(CodeList)
							TempArray = Split(CodeList(i),JoinStr)
							If TempArray(1) = TempList(1) Then
								DlgValue "CodeNameList",i
								Exit For
							End If
						Next i
					Else
						DlgValue "CodeNameList",1
					End If
					DlgValue "CovertType",TempList(2)
					DlgText "StartLine",TempList(3)
					DlgText "Break",TempList(4)
					DlgText "PreAddStr",TempList(5)
					DlgText "PreAppStr",TempList(6)
					DlgText "AppAddStr",TempList(7)
					DlgText "AppAppStr",TempList(8)
					DlgText "RepStr",TempList(9)
					DlgText "RepAsStr",TempList(10)
				Else
					DlgValue "CodeNameList",0
					DlgValue "CovertType",0
					DlgText "StartLine","1"
					DlgText "Break",Separator
					DlgText "PreAddStr",""
					DlgText "PreAppStr",""
					DlgText "AppAddStr",""
					DlgText "AppAppStr",""
					DlgText "RepStr",""
					DlgText "RepAsStr",""
				End If
			End If
		End If

		If DlgItem$ = "ClearButton" Then
			If MsgBox(MsgList(11),vbYesNo+vbInformation,MsgList(8)) = vbYes Then
				ReDim CovertFileList(0)
				ReDim CovertDataListBak(0)
				ReDim CovertDataList(0)
				DlgListBoxArray "CovertFileList",CovertFileList()
				DlgValue "CodeNameList",1
				DlgValue "CovertType",0
				DlgText "StartLine","1"
				DlgText "Break",Separator
				DlgText "PreAddStr",""
				DlgText "PreAppStr",""
				DlgText "AppAddStr",""
				DlgText "AppAppStr",""
				DlgText "RepStr",""
				DlgText "RepAsStr",""
			End If
		End If

		If DlgItem$ = "CovertType" And DlgText("CovertFileList") <> "" Then
			File = DlgText("CovertFileList")
			Stemp = False
			For i = LBound(CovertFileList) To UBound(CovertFileList)
				If CovertFileList(i) = File Then
					Stemp = True
					Exit For
				End If
			Next i
			If Stemp = True Then
				Temp = DlgValue("CovertType")
				TempList = Split(CovertDataListBak(i),JoinStr)
				TempList(2) = Temp
				CovertDataListBak(i) = Join(TempList,JoinStr)
				If DlgValue("AllSame") = 0 Then
					TempList = Split(CovertDataList(i),JoinStr)
					TempList(2) = Temp
					CovertDataList(i) = Join(TempList,JoinStr)
				Else
					For i = LBound(CovertDataList) To UBound(CovertDataList)
						TempList = Split(CovertDataList(i),JoinStr)
						TempList(2) = Temp
						CovertDataList(i) = Join(TempList,JoinStr)
					Next i
				End If
			End If
		End If

		If DlgItem$ = "CodeNameList" Then
			File = DlgText("CovertFileList")
			i = DlgValue("CodeNameList")
			TempList = Split(CodeList(i),JoinStr)
			Code = TempList(1)
			If Code = "_autodetect_all" Or Code = "_autodetect" Or Code = "_autodetect_kr" Then
				Code = CheckCode(File)
				For i = LBound(CodeList) To UBound(CodeList)
					TempArray = Split(CodeList(i),JoinStr)
					If TempArray(1) = Code Then
 						DlgValue "CodeNameList",i
 						Exit For
 					End If
				Next i
			End If
			Stemp = False
			For i = LBound(CovertFileList) To UBound(CovertFileList)
				If CovertFileList(i) = File Then
					Stemp = True
					Exit For
				End If
			Next i
			If Stemp = True Then
				TempList = Split(CovertDataList(i),JoinStr)
				TempList(1) = Code
				CovertDataList(i) = Join(TempList,JoinStr)
				TempList = Split(CovertDataListBak(i),JoinStr)
				TempList(1) = Code
				CovertDataListBak(i) = Join(TempList,JoinStr)
			End If
		End If

		If DlgItem$ = "AllSame" And UBound(CovertFileList) <> 0 Then
			If DlgValue("AllSame") = 1 Then
				If MsgBox(MsgList(12) & MsgList(14),vbYesNo+vbInformation,MsgList(8)) = vbYes Then
					ReDim TempList(8)
					TempList(0) = DlgValue("CovertType")
					TempList(1) = DlgText("StartLine")
					TempList(2) = DlgText("Break")
					TempList(3) = DlgText("PreAddStr")
					TempList(4) = DlgText("PreAppStr")
					TempList(5) = DlgText("AppAddStr")
					TempList(6) = DlgText("AppAppStr")
					TempList(7) = DlgText("RepStr")
					TempList(8) = DlgText("RepAsStr")
					Temp = Join(TempList,JoinStr)
					For i = LBound(CovertDataList) To UBound(CovertDataList)
						TempList = Split(CovertDataList(i),JoinStr)
						CovertDataList(i) = CovertFileList(i) & JoinStr & TempList(1) & JoinStr & Temp
					Next i
				Else
					DlgValue "AllSame",0
				End If
			Else
				If MsgBox(MsgList(13) & MsgList(14),vbYesNo+vbInformation,MsgList(8)) = vbYes Then
					CovertDataList = CovertDataListBak
				Else
					DlgValue "AllSame",1
				End If
			End If
		End If

		If DlgItem$ = "PreViewButton" Or DlgItem$ = "TestButton" Or DlgItem$ = "AllEditButton" Then
			File = DlgText("CovertFileList")
			If File <> "" Then
				FileDataList = CovertDataList
				If DlgItem$ = "PreViewButton" Then
					Stemp = EditFile(File,FileDataList,1)
				ElseIf DlgItem$ = "TestButton" Then
					Stemp = CovertTest(File,FileDataList)
				Else
					n = ShowPopupMenu(AppNames)
   					If n = 0 Then
						Stemp = OpenFile(File,FileDataList,n,False)
					ElseIf n > 0 Then
						OpenFile(File,FileDataList,n,False)
					End If
				End If
				If Stemp = True Then
					If Join(CovertDataList,JoinStr) <> Join(FileDataList,JoinStr) Then
						CovertDataList = FileDataList
						For i = LBound(CovertFileList) To UBound(CovertFileList)
							TempList = Split(CovertDataList(i),JoinStr)
							Temp = TempList(1)
							TempArray = Split(CovertDataListBak(i),JoinStr)
							TempArray(1) = Temp
							CovertDataListBak(i) = Join(TempArray,JoinStr)
							If TempList(0) = File Then
								Code = Temp
								If DlgItem$ = "TestButton" Then
									DlgValue "CovertType",TempList(2)
									DlgText "StartLine",TempList(3)
									DlgText "Break",TempList(4)
									DlgText "PreAddStr",TempList(5)
									DlgText "PreAppStr",TempList(6)
									DlgText "AppAddStr",TempList(7)
									DlgText "AppAppStr",TempList(8)
									DlgText "RepStr",TempList(9)
									DlgText "RepAsStr",TempList(10)
								End If
							End If
						Next i
						If Code <> "" Then
							For i = LBound(CodeList) To UBound(CodeList)
								TempArray = Split(CodeList(i),JoinStr)
								If TempArray(1) = Code Then
 									DlgValue "CodeNameList",i
 									Exit For
 								End If
							Next i
						Else
							DlgValue "CodeNameList",1
						End If
					End If
				End If
			End If
		End If

		If DlgItem$ = "AllRestoreButton" Then
			PSL.OutputWnd.Clear
			n = 0
			BackupFile = ""
			For i = LBound(CovertFileList) To UBound(CovertFileList)
				m = 0
				File = CovertFileList(i)
				BackupFile = File & ".bak"
				If Dir(BackupFile) <> "" Then
					On Error GoTo AllRestoreError
					If Dir(File) <> "" Then
						SetAttr File,vbNormal
						Kill File
					End If
					Name BackupFile As File
					On Error GoTo 0
					n = n + 1
					Call CovertMassage("Restored",File,BackupFile,m,n)
					GoTo AllRestoreNext
					AllRestoreError:
					m = 1
					Call CovertMassage("Restored",File,BackupFile,m,n)
					AllRestoreNext:
				End If
			Next i
			Call CovertMassage("AllRestored",File,BackupFile,m,n)
			If DlgValue("Backup") = 1 Then
				If n <> 0 Then DlgEnable "AllCovertButton",True
				If n <> 0 Then DlgEnable "AllRestoreButton",False
			End If
		End If

		If DlgItem$ = "RestoreButton" And DlgText("CovertFileList") <> "" Then
			n = 0
			m = 0
			File = DlgText("CovertFileList")
			BackupFile = File & ".bak"
			If Dir(BackupFile) <> "" Then
				On Error GoTo RestoreError
				If Dir(File) <> "" Then
					SetAttr File,vbNormal
					Kill File
				End If
				Name BackupFile As File
				On Error GoTo 0
				n = n + 1
				Call CovertMassage("Restored",File,BackupFile,m,n)
				GoTo RestoreNext
				RestoreError:
				m = 1
				Call CovertMassage("Restored",File,BackupFile,m,n)
				RestoreNext:
			End If
		End If

		If DlgItem$ = "AllCovertButton" Then
			PSL.OutputWnd.Clear
			n = 0
			BackupFile = ""
			For i = LBound(CovertDataList) To UBound(CovertDataList)
				m = 0
				File = CovertFileList(i)
				BackupFile = File & ".bak"
				'Call CovertMassage("Coverting",File,BackupFile,m,n)
				TempList = Split(CovertDataList(i),JoinStr)
				Code = TempList(1)
				If Code = "" Then Code = CheckCode(File)
				inText = ReadFile(File,Code)  '不会返回 code
				If inText <> "" Then
					outText = CovertText(inText,CovertDataList(i),m)
					If m <> 0 Then
						If Dir(BackupFile) = "" Then
							On Error GoTo AllCovertError
							SetAttr File,vbNormal
							If DlgValue("Backup") = 1 Then
								Name File As BackupFile
							Else
								BackupFile = ""
							End If
							On Error GoTo 0
							If WriteToFile(File,outText,Code) = True Then
								n = n + 1
							End If
							GoTo AllCovertNext
							AllCovertError:
							m = 0
							Call CovertMassage("NotCoverted",File,BackupFile,m,n)
							AllCovertNext:
						Else
							Call CovertMassage("ExitBackupFile",File,BackupFile,m,n)
						End If
					End If
				End If
				Call CovertMassage("Coverted",File,BackupFile,m,n)
			Next i
			Call CovertMassage("AllCoverted",File,BackupFile,m,n)
			If DlgValue("Backup") = 1 Then
				If n <> 0 Then DlgEnable "AllCovertButton",False
				If n <> 0 Then DlgEnable "AllRestoreButton",True
			End If
		End If

		If DlgItem$ = "CovertButton" And DlgText("CovertFileList") <> "" Then
			n = 0
			m = 0
			File = DlgText("CovertFileList")
			BackupFile = File & ".bak"
			'Call CovertMassage("Coverting",File,BackupFile,m,n)
			TempList = Split(CodeList(DlgValue("CodeNameList")),JoinStr)
			Code = TempList(1)
			If Code = "" Then Code = CheckCode(File)
			inText = ReadFile(File,Code)
			If inText <> "" Then
				i = DlgValue("CovertFileList")
				outText = CovertText(inText,CovertDataList(i),m)
				If m <> 0 Then
					If Dir(BackupFile) = "" Then
						On Error GoTo CovertError
						SetAttr File,vbNormal
						If DlgValue("Backup") = 1 Then
							Name File As BackupFile
						Else
							BackupFile = ""
						End If
						On Error GoTo 0
						If WriteToFile(File,outText,Code) = True Then
							n = n + 1
						End If
						GoTo CovertNext
						CovertError:
						m = 0
						Call CovertMassage("NotCoverted",File,BackupFile,m,n)
						CovertNext:
					Else
						Call CovertMassage("ExitBackupFile",File,BackupFile,m,n)
					End If
				End If
			End If
			Call CovertMassage("Coverted",File,BackupFile,m,n)
		End If

		If DlgItem$ <> "CancelButton" Then
			If DlgText("CovertFileList") <> "" Then
				DlgEnable "CodeNameList",True
				DlgEnable "CovertType",True
				DlgEnable "StartLine",True
				DlgEnable "Break",True
				DlgEnable "PreAddStr",True
				DlgEnable "PreAppStr",True
				DlgEnable "AppAddStr",True
				DlgEnable "AppAppStr",True
				DlgEnable "RepStr",True
				DlgEnable "RepAsStr",True
				DlgEnable "ChangeButton",True
				DlgEnable "DelButton",True
				DlgEnable "ClearButton",True
				DlgEnable "PreViewButton",True
				DlgEnable "TestButton",True
				DlgEnable "AllEditButton",True
				BackupFile = DlgText("CovertFileList") & ".bak"
				If Dir(BackupFile) <> "" Then
					DlgEnable "CovertButton",False
					DlgEnable "RestoreButton",True
					DlgEnable "AllCovertButton",False
					DlgEnable "AllRestoreButton",True
				Else
					DlgEnable "CovertButton",True
					DlgEnable "RestoreButton",False
					DlgEnable "AllCovertButton",True
					DlgEnable "AllRestoreButton",False
				End If
				DlgText "SelectGroupBox",MsgList(7) & "(" & UBound(CovertFileList)+1 & ")"
			Else
				DlgEnable "CodeNameList",False
				DlgEnable "CovertType",False
				DlgEnable "StartLine",False
				DlgEnable "Break",False
				DlgEnable "PreAddStr",False
				DlgEnable "PreAppStr",False
				DlgEnable "AppAddStr",False
				DlgEnable "AppAppStr",False
				DlgEnable "RepStr",False
				DlgEnable "RepAsStr",False
				DlgEnable "ChangeButton",False
				DlgEnable "DelButton",False
				DlgEnable "ClearButton",False
				DlgEnable "PreViewButton",False
				DlgEnable "TestButton",False
				DlgEnable "AllEditButton",False
				DlgEnable "CovertButton",False
				DlgEnable "RestoreButton",False
				DlgEnable "AllCovertButton",False
				DlgEnable "AllRestoreButton",False
				DlgText "SelectGroupBox",MsgList(7) & "(0)"
			End If
			MainDlgFunc = True '防止按下按钮关闭对话框窗口
		End If
	Case 3 ' 文本框或者组合框文本被更改
		If DlgItem$ <> "ExtNameList" And DlgText("CovertFileList") <> "" Then
			If TempFile = "" Then TempFile = DlgText("CovertFileList")
			ReDim TempList(8)
			TempList(0) = DlgValue("CovertType")
			TempList(1) = DlgText("StartLine")
			TempList(2) = DlgText("Break")
			TempList(3) = DlgText("PreAddStr")
			TempList(4) = DlgText("PreAppStr")
			TempList(5) = DlgText("AppAddStr")
			TempList(6) = DlgText("AppAppStr")
			TempList(7) = DlgText("RepStr")
			TempList(8) = DlgText("RepAsStr")
			Temp = Join(TempList,JoinStr)
			For i = LBound(CovertFileList) To UBound(CovertFileList)
				File = CovertFileList(i)
				If File = TempFile Then
					TempList = Split(CovertDataListBak(i),JoinStr)
					CovertDataListBak(i) = File & JoinStr & TempList(1) & JoinStr & Temp
					If DlgValue("AllSame") = 0 Then
						TempList = Split(CovertDataList(i),JoinStr)
						CovertDataList(i) = File & JoinStr & TempList(1) & JoinStr & Temp
						Exit For
					End If
				End If
				If DlgValue("AllSame") = 1 Then
					TempList = Split(CovertDataList(i),JoinStr)
					CovertDataList(i) = File & JoinStr & TempList(1) & JoinStr & Temp
				End If
			Next i
		End If
	Case 4 ' 焦点被更改
		If DlgItem$ <> "ExtNameList" And DlgText("CovertFileList") <> "" Then
			If TempFile = "" Then TempFile = DlgText("CovertFileList")
		End If
	End Select
End Function


'检测并下载新版本
Function Download(Method As String,Url As String,Async As Boolean,Mode As String) As Long
	Dim i As Long,j As Long,m As Long,n As Long,updateINI As String,TempList() As String
	Dim TempPath As String,File As String,OpenFile As Boolean,Body As Variant
	Dim xmlHttp As Object,UrlList() As String,Stemp As Boolean,MsgList() As String
	Dim mErrUrlList() As String,nErrUrlList() As String,BodyBak As Variant

	If getMsgList(UILangList,MsgList,"Download",1) = False Then Exit Function
	Download = 0
	OpenFile = False
	If Join(UpdateSet,"") <> "" Then
		If Mode = "" Then Mode = UpdateSet(0)
		If Url = "" Then Url = UpdateSet(1)
		ExePath = Trim(UpdateSet(2))
		Argument = UpdateSet(3)
	End If
	If Mode = "2" Then Exit Function

	PSL.OutputWnd.Clear
	If Mode = "4" Then PSL.Output MsgList(22)
	If Mode <> "4" Then PSL.Output MsgList(23)
	If ExePath = "" Then
		If Mode <> "4" Then Msg = MsgList(1) & vbCrLf & MsgList(3)
		If Mode = "4" Then Msg = MsgList(2) & vbCrLf & MsgList(3)
		MsgBox(Msg,vbOkOnly+vbInformation,MsgList(0))
		Exit Function
	End If
	Stemp = False
	If Url = "" Or Mode = "" Or Trim(Argument) = "" Then Stemp = True
	If Stemp = False Then
		If InStr(Argument,"%1") = 0 Then Stemp = True
		If InStr(Argument,"%2") = 0 Then Stemp = True
		If InStr(Argument,"%3") = 0 Then Stemp = True
	End If
	If Stemp = True Then
		If Mode <> "4" Then Msg = MsgList(1) & vbCrLf & MsgList(4)
		If Mode = "4" Then Msg = MsgList(2) & vbCrLf & MsgList(4)
		MsgBox(Msg,vbOkOnly+vbInformation,MsgList(0))
		Exit Function
	End If

	On Error Resume Next
	TempList = Split(DefaultObject,";")
	For i = 0 To UBound(TempList)
		Set xmlHttp = CreateObject(TempList(i))
		If Not xmlHttp Is Nothing Then Exit For
	Next i
	If xmlHttp Is Nothing Then
		Err.Source = Join(TempList,"; ")
		Call sysErrorMassage(Err,2)
		Exit Function
	End If
	On Error GoTo 0
	If Mode = "4" Then GoTo getFile
	Stemp = False
	updateINIUrl = updateINIMainUrl & vbCrLf & updateINIMinorUrl
	UrlList = Split(updateINIUrl,vbCrLf)
	ReGetUpdateINIFile:
	For i = LBound(UrlList) To UBound(UrlList)
		updateINI = ""
		Body = ""
		updateUrl = Trim(UrlList(i))
		If updateUrl <> "" Then
			On Error GoTo SkipINIUrl
			xmlHttp.Open Method,updateUrl,Async,User,Password
			xmlHttp.setRequestHeader("If-Modified-Since","0")
			xmlHttp.send()
			If xmlHttp.readyState = 4 Then
				If xmlHttp.Status = 200 Then Body = xmlHttp.responseBody
			End If
			On Error GoTo 0
			If Body <> "" Then
				If LenB(Body) > 0 Then
					updateINI = BytesToBstr(Body,"utf-8")
					If updateINI <> "" Then
						If InStr(LCase(updateINI),LCase(updateAppName)) Then
							Stemp = True
							Exit For
						Else
							updateINI = ""
						End If
					End If
				End If
			End If
		End If
		SkipINIUrl:
		If Err.Number <> 0 Then Err.Clear
	Next i
	xmlHttp.Abort
	If Stemp = False And Url <> "" Then
		UrlList = Split(Url,vbCrLf)
		For i = LBound(UrlList) To UBound(UrlList)
			updateUrl = Trim(UrlList(i))
			If updateUrl <> "" Then
				j = InStrRev(LCase(updateUrl),"/download")
				If j = 0 Then j = InStrRev(updateUrl,"/")
				If j <> 0 Then
					UrlList(i) = Left(updateUrl,j) & "update/" & updateINIFile
					Stemp = True
				End If
			End If
		Next i
		If Stemp = True Then GoTo ReGetUpdateINIFile
	End If
	If updateINI <> "" Then
		UrlList = Split(updateINI,vbCrLf)
		LangName = ""
		For i = LBound(UrlList) To UBound(UrlList)
			L$ = Trim(UrlList(i))
			If L$ <> "" Then
				If Left(L$,1) = "[" And Right(L$,1) = "]" Then
					Header = Trim(Mid(L$,2,Len(L$)-2))
				End If
				setAppStr = ""
				setAppStr = ""
				If Header <> "" Then
					j = InStr(L$,"=")
					If j > 0 Then
						setPreStr = Trim(Left(L$,j - 1))
						setAppStr = Mid(L$,j + 1)
					End If
				End If
				If Header = "Option" And setPreStr <> "" Then
					If setPreStr = "DefaultLanguage" Then DefaultLanguage = Trim(setAppStr)
				End If
				If Header = "Language" And setPreStr <> "" And LangName = "" Then
					TempList = Split(Trim(setAppStr),";")
					For j = LBound(TempList) To UBound(TempList)
						If LCase(TempList(j)) = LCase(OSLanguage) Then
							LangName = setPreStr
							Exit For
						End If
					Next j
				End If
				If Header = updateAppName And setPreStr <> "" Then
					If setPreStr = "Version" Then NewVersion = Trim(setAppStr)
					If InStr(setPreStr,"URL_") Then
						Site = Trim(setAppStr)
						If UpdateSite <> "" Then UpdateSite = UpdateSite & vbCrLf & Site
						If UpdateSite = "" Then UpdateSite = Site
					End If
					If LangName = "" Then LangName = DefaultLanguage
					If InStr(LCase(setPreStr),LCase("Des_" & LangName)) Then
						Des = setAppStr
						If UpdateDes <> "" Then UpdateDes = UpdateDes & vbCrLf & Des
						If UpdateDes = "" Then UpdateDes = Des
					End If
				End If
			End If
			If Header <> updateAppName And NewVersion <> "" Then Exit For
		Next i
		If NewVersion <> "" Then
			n = StrComp(NewVersion,Version)
			If n = 1 Then
				If Mode = "1" Or Mode = "3" Then
					Msg = Replace(MsgList(15),"%s",NewVersion) & vbCrLf & vbCrLf & MsgList(20)
					OKMsg = MsgBox(Msg & vbCrLf & UpdateDes,vbYesNo+vbInformation,MsgList(17))
					If OKMsg = vbNo Then NewVersion = ""
				End If
			ElseIf n = 0 Then
				If Mode = "3" Then
					Msg = Replace(MsgList(14),"%s",NewVersion) & vbCrLf & MsgList(21)
					OKMsg = MsgBox(Msg,vbYesNo+vbInformation,MsgList(17))
					If OKMsg = vbNo Then NewVersion = ""
				Else
					NewVersion = ""
				End If
			ElseIf n < 0 Then
				If Mode = "3" Then
					Msg = Replace(MsgList(14),"%s",NewVersion)
					MsgBox(Msg,vbOkOnly+vbInformation,MsgList(16))
				End If
				NewVersion = ""
			End If
			Download = 1
		End If
		If NewVersion = "" Then
			Set xmlHttp = Nothing
			Exit Function
		End If
	End If

	getFile:
	If Mode <> "4" Then PSL.Output MsgList(24)
	If UpdateSite = "" Then UrlList = Split(Url,vbCrLf)
	If UpdateSite <> "" Then UrlList = ClearArray(Split(Url & vbCrLf & UpdateSite,vbCrLf))
	m = 0
	n = 0
	j = 0
	For i = LBound(UrlList) To UBound(UrlList)
		Body = ""
		updateUrl = Trim(UrlList(i))
		If updateUrl <> "" Then
			j = j + 1
			On Error GoTo Skip
			xmlHttp.Open Method,updateUrl,Async,User,Password
			xmlHttp.setRequestHeader("If-Modified-Since","0")
			xmlHttp.send()
			If xmlHttp.readyState = 4 Then
				If xmlHttp.Status = 200 Then Body = xmlHttp.responseBody
			End If
			On Error GoTo 0
			If Body <> "" Then
				If LenB(Body) > 0 Then
					If Mode <> "4" Then
						Exit For
					Else
						If BodyBak = "" Then BodyBak = Body
					End If
				Else
					If Mode = "4" Then
						ReDim Preserve mErrUrlList(m)
						mErrUrlList(m) = updateUrl
					End If
					m = m + 1
				End If
			Else
				If Mode = "4" Then
					ReDim Preserve nErrUrlList(n)
					nErrUrlList(n) = updateUrl
				End If
				n = n + 1
			End If
			Skip:
			If Err.Number <> 0 Then
				If Mode = "4" Then
					ReDim Preserve nErrUrlList(n)
					nErrUrlList(n) = updateUrl
				End If
				n = n + 1
				Err.Clear
			End If
		End If
	Next i
	xmlHttp.Abort
	Set xmlHttp = Nothing
	If m <> 0 Or n <> 0 Then
		If Mode <> "4" Then
			If n = j Then Msg = MsgList(1) & vbCrLf & MsgList(5)
			If m = j Then Msg = MsgList(1) & vbCrLf & MsgList(6)
			If m = j Or n = j Then
				MsgBox(Msg,vbOkOnly+vbInformation,MsgList(0))
				Exit Function
			End If
		Else
			If m <> 0 Then mMsg = MsgList(34) & vbCrLf & Join(mErrUrlList,vbCrLf)
			If n <> 0 Then nMsg = MsgList(33) & vbCrLf & Join(nErrUrlList,vbCrLf)
			If m <> 0 And n = 0 Then
				Msg = MsgList(2) & vbCrLf & mMsg
			ElseIf m = 0 And n <> 0 Then
				Msg = MsgList(2) & vbCrLf & nMsg
			ElseIf m <> 0 And n <> 0 Then
				Msg = MsgList(2) & vbCrLf & nMsg & vbCrLf & vbCrLf & mMsg
			End If
			MsgBox(Msg,vbOkOnly+vbInformation,MsgList(12))
			Exit Function
		End If
	End If

	If Mode = "4" Then Body = BodyBak
	TempPath = MacroDir & "\temp\"
	File = TempPath & "temp.rar"
	If LenB(Body) > 0 Then
		On Error Resume Next
		If Dir(TempPath & "*.*") = "" Then MkDir TempPath
		If BytesToFile(Body,File) = False Then
			FN = FreeFile
			Open File For Binary Access Write As #FN
			Put #FN,,Body
			Close #FN
		End If
		On Error GoTo 0
	End If
	Set xmlHttp = Nothing

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
			If InStr(Argument,"""%2""") Then Argument = Replace(Argument,"%2","*.*")
			If InStr(Argument,"""%3""") Then Argument = Replace(Argument,"%3",TempPath)
			If InStr(Argument,"%1") Then Argument = Replace(Argument,"%1","""" & File & """")
			If InStr(Argument,"%2") Then Argument = Replace(Argument,"%2","*.*")
			If InStr(Argument,"%3") Then Argument = Replace(Argument,"%3","""" & TempPath & """")
		End If
		If ExePath <> "" And Dir(ExePath) <> "" Then
			If Mode <> "4" Then PSL.Output MsgList(25)
			On Error Resume Next
			Set WshShell = CreateObject("WScript.Shell")
			If WshShell Is Nothing Then
				Err.Source = "WScript.Shell"
				Call sysErrorMassage(Err,2)
				Exit Function
			End If
			On Error GoTo 0
			Return = WshShell.Run("""" & ExePath & """ " & Argument,0,True)
			Set WshShell = Nothing
			If Return = 0 Then
				File = TempPath & updateMainFile
				If Dir(File) <> "" Then OpenFile = True
			Else
				If Mode <> "4" Then Msg = MsgList(1) & vbCrLf & MsgList(7)
				If Mode = "4" Then Msg = MsgList(2) & vbCrLf & MsgList(7)
				MsgBox(Msg,vbOkOnly+vbInformation,MsgList(0))
			End If
		ElseIf ExePath <> "" And Dir(ExePath) = "" Then
			ExeName = Mid(Left(ExePath,InStrRev(ExePath,".")-1),InStrRev(ExePath,"\")+1)
			Msg = MsgList(8) & ExeName & vbCrLf & MsgList(9) & ExePath & vbCrLf & MsgList(10) & _
					Argument & vbCrLf & vbCrLf & MsgList(11)
			MsgBox(Msg,vbOkOnly+vbInformation,MsgList(0))
		End If
	End If

	If OpenFile = True Then
		If Mode <> "4" Then PSL.Output MsgList(26)
		On Error Resume Next
		FN = FreeFile
		Open File For Input As #FN
		Do While Not EOF(FN)
			Line Input #FN,L$
			FindStr = "Private Const Version = "
			n = InStr(L$,FindStr)
			If n > 0 Then
				WebVersion = Mid(L$,n+Len(FindStr)+1,10)
				Exit Do
			End If
		Loop
		Close #FN
		On Error GoTo 0
		If WebVersion = "" Then MsgBox(MsgList(19),vbOkCancel+vbInformation,MsgList(0))
	End If

	If WebVersion <> "" Then
		If Url <> Join(UrlList,vbCrLf) Then UpdateSet(1) = Join(UrlList,vbCrLf)
		If Mode = "4" Then
			Msg = MsgList(13) & vbCrLf & Replace(MsgList(14),"%s",WebVersion)
			MsgBox(Msg,vbOkOnly+vbInformation,MsgList(12))
			Download = 2
			GoTo delFile
		ElseIf Mode = "0" Or Mode = "1" Or Mode = "3" Then
			n = StrComp(WebVersion,Version)
			If n = 1 Or (n = 0 And Mode = "3") Then
				If NewVersion = "" Then
					If Mode = "1" Then
						Msg = Replace(MsgList(15),"%s",WebVersion)
					ElseIf Mode = "3" Then
						If n = 0 Then
							Msg = Replace(MsgList(14),"%s",WebVersion) & vbCrLf & MsgList(21)
						Else
							Msg = Replace(MsgList(15),"%s",WebVersion)
						End If
					End If
					OKMsg = MsgBox(Msg,vbYesNo+vbInformation,MsgList(17))
					If OKMsg = vbNo Then GoTo delFile
				Else
					If NewVersion <> WebVersion Then
						Msg = MsgList(1) & vbCrLf & Replace(MsgList(29),"%s",WebVersion) & vbCrLf & MsgList(32)
						MsgBox(Msg,vbOkOnly+vbInformation,MsgList(0))
						GoTo delFile
					End If
				End If
			Else
				If Mode = "0" Or Mode = "1" Then
					PSL.Output Replace(MsgList(30),"%s",WebVersion)
				ElseIf Mode = "3" Then
					Msg = Replace(MsgList(31),"%s",WebVersion)
					MsgBox(Msg,vbOkOnly+vbInformation,MsgList(16))
				End If
				GoTo delFile
			End If
		End If

		PSL.Output MsgList(27)
		On Error Resume Next
		lngFile = Dir$(TempPath & "*.lng")
		iniFile = Dir$(TempPath & "*.ini")
		txtFile = Dir$(TempPath & "*.txt")
		If lngFile <> "" Or iniFile <> "" Then
			If Dir(MacroDir & "\Data\" & "*.*") = "" Then MkDir MacroDir & "\Data\"
		End If
		If txtFile <> "" Then
			If Dir(MacroDir & "\Doc\" & "*.*") = "" Then MkDir MacroDir & "\Doc\"
		End If
		n = 0
		File = Dir$(TempPath & "*.*")
		Do While File <> ""
			ReDim Preserve TempList(n)
			TempList(n) = File
			n = n + 1
			File = Dir$()
		Loop
		If n > 0 Then
			For i = LBound(TempList) To UBound(TempList)
				File = TempList(i)
				sExtName = LCase(Mid(File,InStrRev(File,".")+1))
				If sExtName = "bas" Then
					FileCopy TempPath & File,MacroDir & "\" & File
					Kill TempPath & File
				ElseIf sExtName = "lng" Or sExtName = "dat" Then
					FileCopy TempPath & File,MacroDir & "\Data\" & File
					Kill TempPath & File
				ElseIf sExtName = "txt" Then
					FileCopy TempPath & File,MacroDir & "\Doc\" & File
					Kill TempPath & File
					If Dir(MacroDir & "\Data\" & File) <> "" Then
						Kill MacroDir & "\Data\" & File
					End If
				ElseIf sExtName = "ini" Then
					If Dir(MacroDir & "\Data\" & File) <> "" Then
						sdate = FileDateTime(TempPath & File)
						tdate = FileDateTime(MacroDir & "\Data\" & File)
						If sdate > tdate Then
							FileCopy TempPath & File,MacroDir & "\Data\" & File
							Kill TempPath & File
						End If
					Else
						FileCopy TempPath & File,MacroDir & "\Data\" & File
						Kill TempPath & File
					End If
				End If
			Next i
			PSL.Output MsgList(28)
			MsgBox(MsgList(18),vbOkOnly+vbInformation,MsgList(17))
			Download = 3
		End If
		On Error GoTo 0
	End If

	delFile:
	File = Dir$(TempPath & "*.*")
	If File <> "" Then
		On Error Resume Next
		Do While File <> ""
			Kill TempPath & File
			File = Dir$()
		Loop
		If Dir(TempPath & "*.*") = "" Then RmDir TempPath
		On Error GoTo 0
	End If
End Function


'从注册表中获取 RAR 扩展名的默认程序
Function getCMDPath(ExtName As String,CmdPath As String,Argument As String) As String
	Dim i As Long,ExePathStr As String,ExeArg As String,WshShell As Object
	On Error Resume Next
	Set WshShell = CreateObject("WScript.Shell")
	If WshShell Is Nothing Then
		Err.Source = "WScript.Shell"
		Call sysErrorMassage(Err,2)
		Exit Function
	End If
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
		AppExePath = Mid(ExePathStr,Len(PreExePath)+1)
		i = InStr(AppExePath," ")
		If i > 0 Then AppExePath = Left(AppExePath,i-1)
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
	End If
	CmdPath = ExePath
	Argument = ExeArg
	getCMDPath = ExePath & JoinStr & ExeArg
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


'自定义参数
Function Settings(OptionID As Long) As Long
	Dim MsgList() As String

	If getMsgList(UILangList,MsgList,"Settings",1) = False Then Exit Function
	If Join(UIFileList,"") = "" Then
		Call GetUIList(UIFileList,UIDataList)
	End If
  	UIFileList(0) = MsgList(31)

	Begin Dialog UserDialog 660,518,MsgList(0),.SetFunc ' %GRID:10,7,1,1
		Text 20,7,620,28,MsgList(1),.MainText
		OptionGroup .Options
			OptionButton 170,42,210,14,MsgList(8),.AutoUpdate
			OptionButton 390,42,210,14,MsgList(26),.UIFileListSet

		GroupBox 20,70,620,91,MsgList(9),.UpdateSetGroup
		OptionGroup .UpdateSet
			OptionButton 40,91,580,14,MsgList(10),.AutoButton
			OptionButton 40,112,580,14,MsgList(11),.ManualButton
			OptionButton 40,133,580,14,MsgList(12),.OffButton
		GroupBox 20,175,620,49,MsgList(13),.CheckGroup
		Text 40,196,100,14,MsgList(14),.UpdateCycleText
		TextBox 150,192,40,21,.UpdateCycleBox
		Text 200,196,60,14,MsgList(15),.UpdateDatesText
		Text 280,196,130,14,MsgList(16),.UpdateDateText
		TextBox 420,192,100,21,.UpdateDateBox
		PushButton 540,192,80,21,MsgList(21),.CheckButton
		GroupBox 20,238,620,98,MsgList(17),.WebSiteGroup
		TextBox 40,259,580,63,.WebSiteBox,1
		GroupBox 20,350,620,126,MsgList(18),.CmdGroup
		Text 40,371,550,14,MsgList(19),.CmdPathBoxText
		Text 40,420,550,14,MsgList(20),.ArgumentBoxText
		TextBox 40,392,550,21,.CmdPathBox
		TextBox 40,441,550,21,.ArgumentBox
		PushButton 590,392,30,21,MsgList(2),.ExeBrowseButton
		PushButton 590,441,30,21,MsgList(3),.ArgumentButton

		GroupBox 20,70,620,406,MsgList(27),.UIFileSetGroup
		Text 40,91,580,231,MsgList(28),.UIFileSetText1
		Text 40,336,580,14,MsgList(29),.UIFileSetText2
		DropListBox 40,357,580,21,UIFileList(),.UIFileList
		Text 40,392,580,70,MsgList(30),.UIFileSetText3

		PushButton 20,490,90,21,MsgList(4),.HelpButton
		PushButton 120,490,100,21,MsgList(7),.ResetButton
		PushButton 330,490,90,21,MsgList(5),.TestButton
		PushButton 230,490,90,21,MsgList(6),.CleanButton
		PushButton 120,490,160,21,MsgList(32),.EditUIFileButton
		OKButton 460,490,90,21,.OKButton
		CancelButton 560,490,80,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	dlg.Options = OptionID
	If Dialog(dlg) = 0 Then Exit Function
	Settings = dlg.Options
End Function


'请务必查看对话框帮助主题以了解更多信息。
Private Function SetFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Long,j As Long,n As Long,HeaderID As Long,Path As String,Stemp As Boolean
	Dim TempList() As String,TempArray() As String,MsgList() As String

	If getMsgList(UILangList,MsgList,"SetFunc",1) = False Then
		SetFunc = True '防止按下按钮关闭对话框窗口
		Exit Function
	End If

	If Action% < 3 Then
		If DlgValue("Options") = 0 Then
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

			DlgVisible "UIFileSetGroup",False
			DlgVisible "UIFileSetText1",False
			DlgVisible "UIFileSetText2",False
			DlgVisible "UIFileList",False
			DlgVisible "UIFileSetText3",False

			DlgVisible "ResetButton",True
			DlgVisible "TestButton",True
			DlgVisible "CleanButton",True
			DlgVisible "EditUIFileButton",False
		ElseIf DlgValue("Options") = 1 Then
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

			DlgVisible "UIFileSetGroup",True
			DlgVisible "UIFileSetText1",True
			DlgVisible "UIFileSetText2",True
			DlgVisible "UIFileList",True
			DlgVisible "UIFileSetText3",True

			DlgVisible "ResetButton",False
			DlgVisible "TestButton",False
			DlgVisible "CleanButton",False
			DlgVisible "EditUIFileButton",True
		End If
	End If

	Select Case Action%
	Case 1 ' 对话框窗口初始化
		If Join(UIFileList,"") <> "" Then
			Stemp = False
			If Join(Selected,"") <> "" Then UILangID = LCase(Selected(0))
			If UILangID <> "" And UILangID <> "0" Then
				TempList = Split(UILangID,";")
				For i = 1 To UBound(UIDataList)
					TempArray = Split(UIDataList(i),JoinStr)
					Temp = LCase(TempArray(1))
					If Temp = UILangID Then
						DlgValue "UIFileList",i
						Stemp = True
						Exit For
					End If
					TempArray = Split(Temp,";")
					For j = 0 To UBound(TempList)
						For n = 0 To UBound(TempArray)
							If TempList(j) = TempArray(n) Then
								DlgValue "UIFileList",i
								Stemp = True
								Exit For
							End If
						Next n
						If Stemp =True Then Exit For
					Next j
					If Stemp =True Then Exit For
				Next i
			End If
			If Stemp = False Then DlgValue "UIFileList",0
		End If

		If Join(UpdateSet,"") <> "" Then
			DlgValue "UpdateSet",StrToLong(UpdateSet(0))
			DlgText "WebSiteBox",UpdateSet(1)
			DlgText "CmdPathBox",UpdateSet(2)
			DlgText "ArgumentBox",UpdateSet(3)
			DlgText "UpdateCycleBox",UpdateSet(4)
			DlgText "UpdateDateBox",UpdateSet(5)
		End If
		If DlgText("UpdateDateBox") = "" Then DlgText "UpdateDateBox",MsgList(3)
		DlgEnable "UpdateDateBox",False
	Case 2 ' 数值更改或者按下了按钮
		If DlgItem$ = "HelpButton" Then
			If DlgValue("Options") = 0 Then Call Help("UpdateSetHelp")
			If DlgValue("Options") = 1 Then Call Help("UILangSetHelp")
			SetFunc = True '防止按下按钮关闭对话框窗口
			Exit Function
		End If

		If DlgValue("Options") = 0 Then
			If DlgItem$ = "ExeBrowseButton" Then
				If PSL.SelectFile(Path,True,MsgList(6),MsgList(5)) = True Then DlgText "CmdPathBox",Path
			End If
			If DlgItem$ = "ArgumentButton" Then
				ReDim TempArray(2)
				TempArray(0) = MsgList(8)
				TempArray(1) = MsgList(9)
				TempArray(2) = MsgList(10)
				Temp = DlgText("ArgumentBox")
				i = ShowPopupMenu(TempArray)
				If i = 0 Then DlgText "ArgumentBox",Temp  & " " & """%1"""
				If i = 1 Then DlgText "ArgumentBox",Temp  & " " & """%2"""
				If i = 2 Then DlgText "ArgumentBox",Temp  & " " & """%3"""
			End If
			If DlgItem$ = "ResetButton" Then
				ReDim TempArray(1)
				TempArray(0) = MsgList(1)
				TempArray(1) = MsgList(2)
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
					DlgValue "UpdateSet",StrToLong(UpdateSetBak(0))
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
	    		If DlgText("CmdPathBox") = "" Or DlgText("ArgumentBox") = "" Then
	    			MsgBox(MsgList(7) & RegKey,vbOkOnly+vbInformation,MsgList(0))
	    			SetFunc = True '防止按下按钮关闭对话框窗口
					Exit Function
	    		End If
				UpdateUrl = DlgText("WebSiteBox")
				i = Download(updateMethod,updateUrl,updateAsync,"3")
	    		If i > 0 Then
	    			Stemp = False
	    			If UpdateSet(5) < Format(Date,"yyyy-MM-dd") Then Stemp = True
	    			If i = 3 And Join(UpdateSet,"") = Join(UpdateSetBak,"") Then Stemp = True
	    			If Stemp = True Then
						DlgText "UpdateDateBox",Format(Date,"yyyy-MM-dd")
						UpdateSet(5) = DlgText("UpdateDateBox")
						If WriteSettings(UpdateSet,"Update") = False Then
							MsgBox(MsgList(11) & RegKey,vbOkOnly+vbInformation,MsgList(0))
							SetFunc = True '防止按下按钮关闭对话框窗口
							Exit Function
						Else
							UpdateSetBak(5) = DlgText("UpdateDateBox")
						End If
					End If
					If i = 3 Then Exit All
				End If
			End If
			If DlgItem$ = "TestButton" Then
				If DlgText("CmdPathBox") = "" Or DlgText("ArgumentBox") = "" Then
	    			MsgBox(MsgList(7) & RegKey,vbOkOnly+vbInformation,MsgList(0))
	    			SetFunc = True '防止按下按钮关闭对话框窗口
					Exit Function
	    		End If
	    		Download(updateMethod,updateUrl,updateAsync,"4")
	    	End If
			UpdateSet(0) = DlgValue("UpdateSet")
			UpdateSet(1) = DlgText("WebSiteBox")
			UpdateSet(2) = DlgText("CmdPathBox")
			UpdateSet(3) = DlgText("ArgumentBox")
			UpdateSet(4) = DlgText("UpdateCycleBox")
			UpdateSet(5) = DlgText("UpdateDateBox")
			If UpdateSet(5) = MsgList(3) Then UpdateSet(5) = ""
		ElseIf DlgValue("Options") = 1 Then
			Stemp = False
			If DlgItem$ = "UIFileList" Then
				HeaderID = DlgValue("UIFileList")
				TempArray = Split(UIDataList(HeaderID),JoinStr)
				LangID = TempArray(1)
				If LCase(Selected(0)) <> LCase(LangID) Then Stemp = True
			End If
			If DlgItem$ = "EditUIFileButton" Then
				HeaderID = DlgValue("UIFileList")
				n = UBound(UIDataList)
				ReDim FileDataList(n-1)
				For i = 1 To n
					TempArray = Split(UIDataList(i),JoinStr)
					Path = TempArray(2)
					FileDataList(i-1) = MacroDir & "\Data\" & Path & JoinStr & "unicodeFFFE"
				Next i
				If EditFile(LangFile,FileDataList,0) = True Then
					If GetUIList(UIFileList,UIDataList) = True Then
						TempArray = Split(UIDataList(HeaderID),JoinStr)
						Stemp = True
					End If
				End If
			End If

			If Stemp = True Then
				If HeaderID <> 0 Then
					TempLangFile = MacroDir & "\Data\" & TempArray(2)
				Else
					LangID = LCase(TempArray(1))
					If LangID = "" Or LangID = "0" Then LangID = LCase(OSLanguage)
					TempList = Split(LangID,";")
					For i = 1 To UBound(UIDataList)
						TempArray = Split(UIDataList(i),JoinStr)
						Temp = LCase(TempArray(1))
						Path = TempArray(2)
						If Temp = LangID Then
							TempLangFile = MacroDir & "\Data\" & Path
							Exit For
						End If
						TempArray = Split(Temp,";")
						For j = 0 To UBound(TempList)
							For n = 0 To UBound(TempArray)
								If TempList(j) = TempArray(n) Then
									TempLangFile = MacroDir & "\Data\" & Path
									Exit For
								End If
							Next n
							If TempLangFile <> "" Then Exit For
						Next j
						If TempLangFile <> "" Then Exit For
					Next i
				End If
				Stemp = False
				If TempLangFile <> "" Then
					If Dir(TempLangFile) <> "" Then
						If getUILangList(TempLangFile,UILangList) = True Then
							If DlgItem$ = "UIFileList" Then
								If HeaderID = 0 Then Selected(0) = "0"
								If HeaderID <> 0 Then Selected(0) = LangID
							End If
							LangFile = TempLangFile
							Stemp = True
						End If
					End If
				End If
			End If

			If Stemp = True Then
				'重置字符代码列表
				On Error Resume Next
				Set objStream = CreateObject("Adodb.Stream")
				On Error GoTo 0
				If objStream Is Nothing Then CodeList = CodePageList(0,0)
				If Not objStream Is Nothing Then CodeList = CodePageList(0,49)

				'重置设置对话框字串
				If getMsgList(UILangList,MsgList,"Settings",1) = False Then
					SetFunc = True ' 防止按下按钮关闭对话框窗口
					Exit Function
				End If

				'重置对话框项目名称
				UIFileList(0) = MsgList(31)
				DlgListBoxArray "UIFileList",UIFileList
				DlgValue "UIFileList",HeaderID

				DlgText -1,MsgList(0)
				DlgText "MainText",MsgList(1)
				DlgText "AutoUpdate",MsgList(8)
				DlgText "UIFileListSet",MsgList(26)

				DlgText "UpdateSetGroup",MsgList(9)
				DlgText "AutoButton",MsgList(10)
				DlgText "ManualButton",MsgList(11)
				DlgText "OffButton",MsgList(12)
				DlgText "CheckGroup",MsgList(13)
				DlgText "UpdateCycleText",MsgList(14)
				DlgText "UpdateDatesText",MsgList(15)
				DlgText "UpdateDateText",MsgList(16)
				DlgText "CheckButton",MsgList(21)
				DlgText "WebSiteGroup",MsgList(17)
				DlgText "CmdGroup",MsgList(18)
				DlgText "CmdPathBoxText",MsgList(19)
				DlgText "ArgumentBoxText",MsgList(20)
				DlgText "ExeBrowseButton",MsgList(3)
				DlgText "ArgumentButton",MsgList(2)

				DlgText "UIFileSetGroup",MsgList(27)
				DlgText "UIFileSetText1",MsgList(28)
				DlgText "UIFileSetText2",MsgList(29)
				DlgText "UIFileSetText3",MsgList(30)

				DlgText "HelpButton",MsgList(4)
				DlgText "ResetButton",MsgList(7)
				DlgText "TestButton",MsgList(5)
				DlgText "CleanButton",MsgList(6)
				DlgText "EditUIFileButton",MsgList(32)
			End If
		End If

		If DlgItem$ = "OKButton" Then
			If DlgText("CmdPathBox") = "" Or DlgText("ArgumentBox") = "" Then
    			MsgBox(MsgList(7) & RegKey,vbOkOnly+vbInformation,MsgList(0))
    			SetFunc = True '防止按下按钮关闭对话框窗口
				Exit Function
    		End If
			If WriteSettings(UpdateSet,"Sets") = False Then
				MsgBox(MsgList(4) & RegKey,vbOkOnly+vbInformation,MsgList(0))
				SetFunc = True '防止按下按钮关闭对话框窗口
				Exit Function
			Else
				UpdateSetBak = UpdateSet
			End If
		End If
		If DlgItem$ = "CancelButton" Then UpdateSet = UpdateSetBak
		If DlgItem$ <> "OKButton" And DlgItem$ <> "CancelButton" Then
			SetFunc = True '防止按下按钮关闭对话框窗口
		End If
	Case 3 ' 文本框或者组合框文本被更改
		If DlgValue("Options") = 0 Then
			UpdateSet(0) = DlgValue("UpdateSet")
			UpdateSet(1) = DlgText("WebSiteBox")
			UpdateSet(2) = DlgText("CmdPathBox")
			UpdateSet(3) = DlgText("ArgumentBox")
			UpdateSet(4) = DlgText("UpdateCycleBox")
			UpdateSet(5) = DlgText("UpdateDateBox")
			If UpdateSet(5) = MsgList(3) Then UpdateSet(5) = ""
		End If
	End Select
	Exit Function
End Function


'获取设置
Function getSettings(SelSet As String,DataList() As String) As Long
	Dim i As Long,j As Long,SetsArray() As String,Temp As String
	getSettings = 0
	ReDim SetsArray(5)
	On Error GoTo ExitFunction
	OldVersion = GetSetting(updateAppName,"Option","Version","")
	If SelSet = "" Or SelSet = "Option" Then
		Selected(0) = GetSetting(updateAppName,"Option","UILanguageID",0)
		If SelSet = "Option" Then
			If Join(Selected,"") <> "" Then getSettings = 1
			Exit Function
		End If
	End If
	If SelSet = "" Or SelSet = "Update" Then
		SetsArray(0) = GetSetting(updateAppName,"Update","UpdateMode",1)
		j = GetSetting(updateAppName,"Update","Count",0)
		For i = 0 To j
			Temp = GetSetting(updateAppName,"Update",CStr(i),"")
			If Temp <> "" Then
				If SetsArray(1) <> "" Then SetsArray(1) = SetsArray(1) & vbCrLf & Temp
				If SetsArray(1) = "" Then SetsArray(1) = Temp
			End If
		Next i
		SetsArray(2) = GetSetting(updateAppName,"Update","Path","")
		SetsArray(3) = GetSetting(updateAppName,"Update","Argument","")
		SetsArray(4) = GetSetting(updateAppName,"Update","UpdateCycle",7)
		SetsArray(5) = GetSetting(updateAppName,"Update","UpdateDate","")
		If Join(SetsArray,"") <> "" Then
			DataList = SetsArray
			getSettings = 2
		End If
	End If
	ExitFunction:
End Function


'写入设置
Function WriteSettings(DataList() As String,WriteType As String) As Boolean
	Dim i As Long,UpdateSiteList() As String
	WriteSettings = False
	On Error GoTo ExitFunction
	SaveSetting(updateAppName,"Option","Version",Version)
	If WriteType = "Sets" Or WriteType = "All" Then
		SaveSetting(updateAppName,"Option","UILanguageID",Selected(0))
	End If
	If WriteType = "Update" Or WriteType = "Sets" Or WriteType = "All" Then
		If Join(DataList,"") <> "" Then
			On Error Resume Next
			DeleteSetting(updateAppName,"Update")
			On Error GoTo 0
			UpdateSiteList = Split(DataList(1),vbCrLf,-1)
			SaveSetting(updateAppName,"Update","UpdateMode",DataList(0))
			For i = LBound(UpdateSiteList) To UBound(UpdateSiteList)
				SaveSetting(updateAppName,"Update",CStr(i),UpdateSiteList(i))
			Next i
			SaveSetting(updateAppName,"Update","Count",UBound(UpdateSiteList))
			SaveSetting(updateAppName,"Update","Path",DataList(2))
			SaveSetting(updateAppName,"Update","Argument",DataList(3))
			SaveSetting(updateAppName,"Update","UpdateCycle",DataList(4))
			SaveSetting(updateAppName,"Update","UpdateDate",DataList(5))
		End If
	End If
	WriteSettings = True
	ExitFunction:
End Function


'打开文件
Function OpenFile(FilePath As String,FileDataList() As String,x As Long,RunStemp As Boolean) As Boolean
	Dim i As Long,ExePathStr As String,Argument As String,MsgList() As String

	OpenFile = False
	If getMsgList(UILangList,MsgList,"OpenFile",1) = False Then Exit Function
	File = FilePath

	If x > 0 Then
		On Error Resume Next
		Set WshShell = CreateObject("WScript.Shell")
		If WshShell Is Nothing Then
			Err.Source = "WScript.Shell"
			Call sysErrorMassage(Err,2)
			Exit Function
		End If
		On Error GoTo 0
	End If

	If x = 0 Then
		If EditFile(File,FileDataList,0) = True Then OpenFile = True
	ElseIf x = 1 Then
		If Dir(Environ("SystemRoot") & "\system32\notepad.exe") = "" Then
			If Dir(Environ("SystemRoot") & "\notepad.exe") = "" Then
				MsgBox MsgList(1),vbOkOnly+vbInformation,MsgList(0)
			Else
				ExePath = "%SystemRoot%\notepad.exe"
			End If
		Else
			ExePath = "%SystemRoot%\system32\notepad.exe"
		End If
		If ExePath <> "" Then
			File = """" & File & """"
			Return = WshShell.Run("""" & ExePath & """ " & File,1,RunStemp)
			If Return <> 0 Then
				MsgBox MsgList(4),vbOkOnly+vbInformation,MsgList(0)
			Else
				OpenFile = True
			End If
		End If
	ElseIf x = 2 Then
		i = InStrRev(File,".")
		If i <> 0 Then ExtName = Mid(File,i)
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
			AppExePath = Mid(ExePathStr,Len(PreExePath)+1)
			i = InStr(AppExePath," ")
			If i > 0 Then AppExePath = Left(AppExePath,i-1)
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
			Return = WshShell.Run("""" & ExePath & """ " & ArgumentFile,1,RunStemp)
			If Return <> 0 Then
				MsgBox MsgList(4),vbOkOnly+vbInformation,MsgList(0)
			Else
				If LCase(ExeName) <> "notepad.exe" And InStr(Join(AppNames,"|"),ExeName) = 0 Then
					Call AddArray(AppNames,AppPaths,ExeName,ExePath & JoinStr & Argument)
				End If
				OpenFile = True
			End If
		End If
		If ExePath = "" Then MsgBox MsgList(2),vbOkOnly+vbInformation,MsgList(0)
		If ExePath <> "" And Dir(ExePath) = "" Then
			Msg = MsgList(6) & ExeName & vbCrLf & MsgList(7) & ExePath & vbCrLf & MsgList(8) & _
					Argument & vbCrLf & vbCrLf & MsgList(3)
			MsgBox Msg,vbOkOnly+vbInformation,MsgList(0)
		End If
  	ElseIf x = 3 Then
		Call CommandInput(ExePathStr,Argument)
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
			Return = WshShell.Run("""" & ExePath & """ " & ArgumentFile,1,RunStemp)
			If Return <> 0 Then
				MsgBox MsgList(5),vbOkOnly+vbInformation,MsgList(0)
			Else
				ExeName = Mid(ExePath,InStrRev(ExePath,"\")+1)
				If LCase(ExeName) <> "notepad.exe" And InStr(Join(AppNames,"|"),ExeName) = 0 Then
					Call AddArray(AppNames,AppPaths,ExeName,ExePath & JoinStr & Argument)
				End If
				OpenFile = True
			End If
		End If
		If ExePath <> "" And Dir(ExePath) = "" Then
			MsgBox ExeName & MsgList(3),vbOkOnly+vbInformation,MsgList(0)
		End If
	ElseIf x > 3 Then
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
			Return = WshShell.Run("""" & ExePath & """ " & ArgumentFile,1,RunStemp)
			If Return <> 0 Then
				MsgBox MsgList(4),vbOkOnly+vbInformation,MsgList(0)
			Else
				OpenFile = True
			End If
		End If
		If ExePath = "" Then MsgBox MsgList(2),vbOkOnly+vbInformation,MsgList(0)
		If ExePath <> "" And Dir(ExePath) = "" Then
			MsgBox ExeName & MsgList(3),vbOkOnly+vbInformation,MsgList(0)
		End If
	End If
	Set WshShell = Nothing
End Function


'增加或更改数组项目
Function CreateArray(Header As String,Data As String,HeaderList() As String,DataList() As String) As Boolean
	Dim i As Long,n As Long
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
Function DelArray(dName As String,dList() As String,Separator As String,Num As Long) As Variant
	Dim i As Long,n As Long,TempList() As String
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


'清理数组中重复的数据
Function ClearArray(xArray() As String) As Variant
	Dim yArray() As String,Stemp As Boolean,i As Long,j As Long,y As Long
	ClearArray = xArray
	If UBound(xArray) = 0 Then Exit Function
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


'添加不相同的数组元素
Sub AddArray(AppNames() As String,AppPaths() As String,CmdName As String,CmdPath As String)
	Dim i As Long,n As Long,Stemp As Boolean
	If CmdName = "" And CmdPath = "" Then Exit Sub
	n = UBound(AppNames)
	Stemp = False
	For i = LBound(AppNames) To UBound(AppNames)
		If LCase(AppNames(i)) = LCase(CmdName) Then
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


'测试预览
Function CovertTest(File As String,FileDataList() As String) As Boolean
	Dim i As Long,CodeNameList() As String,MsgList() As String
	Dim FileDataListBak() As String

	CovertTest = False
	If getMsgList(UILangList,MsgList,"CovertTest",1) = False Then Exit Function

	'Dim objStream As Object
	'Set objStream = CreateObject("Adodb.Stream")
	'If objStream Is Nothing Then CodeList = CodePageList(0,0)
	'If Not objStream Is Nothing Then CodeList = CodePageList(0,49)
	'Set objStream = Nothing

	For i = LBound(CodeList) To UBound(CodeList)
		ReDim Preserve CodeNameList(i)
		TempArray = Split(CodeList(i),JoinStr)
		CodeNameList(i) = TempArray(0)
	Next i
	FileDataListBak = FileDataList

	Begin Dialog UserDialog 660,518,MsgList(0),.CovertTestFunc ' %GRID:10,7,1,1
		Text 10,7,640,14,File,.FileName,2

		GroupBox 10,28,640,70,MsgList(1),.CovertTypeGroupBox
		DropListBox 190,42,280,21,CovertTypeList(),.CovertType

		Text 40,70,90,14,MsgList(2),.StartLineText
		TextBox 140,68,50,19,.StartLine
		Text 200,70,70,14,MsgList(3),.BreakText
		TextBox 280,68,50,19,.Break
		Text 340,70,80,14,MsgList(4),.CodeNameText
		DropListBox 430,68,200,19,CodeNameList(),.CodeNameList

		GroupBox 10,105,640,84,MsgList(13),.AdditionGroupBox
		Text 40,123,80,14,MsgList(14),.PreAddition
		Text 130,123,80,14,MsgList(15),.PreAddStrText
		TextBox 220,119,150,21,.PreAddStr
		Text 380,123,80,14,MsgList(16),.PreAppStrText
		TextBox 470,119,160,21,.PreAppStr
		Text 40,144,80,14,MsgList(17),.AppAddition
		Text 130,144,80,14,MsgList(18),.AppAddStrText
		TextBox 220,140,150,21,.AppAddStr
		Text 380,144,80,14,MsgList(19),.AppAppStrText
		TextBox 470,140,160,21,.AppAppStr

		Text 40,165,80,14,MsgList(20),.Replaces
		Text 130,165,80,14,MsgList(21),.RepStrText
		TextBox 220,161,150,21,.RepStr
		Text 380,165,80,14,MsgList(22),.RepAsStrText
		TextBox 470,161,160,21,.RepAsStr

		Text 10,196,490,14,MsgList(5)
		TextBox 10,217,640,119,.InTextBox,1
		Text 520,196,80,14,MsgList(6),.LineNumText
		TextBox 600,194,50,19,.LineNum

		Text 10,345,500,14,MsgList(7)
		TextBox 10,364,640,119,.OutTextBox,1
		CheckBox 520,345,130,14,MsgList(8),.Sync,1

		PushButton 20,490,90,21,MsgList(9),.TestButton
		PushButton 120,490,90,21,MsgList(10),.ClearButton
		PushButton 220,490,90,21,MsgList(11),.PreviousButton
		PushButton 320,490,90,21,MsgList(12),.NextButton
		OKButton 460,490,80,21,.OKButton
		CancelButton 550,490,90,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	If Dialog(dlg) = 0 Then
		FileDataList = FileDataListBak
		Exit Function
	End If
	CovertTest = True
End Function


'测试转换程序对话框函数
Private Function CovertTestFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i As Long,m As Long,n As Long,LineNum As Long,Sync As Long,FileNo As Long
	Dim Code As String,CodeID As Long,TempList() As String,MsgList() As String

	If getMsgList(UILangList,MsgList,"CovertTestFunc",1) = False Then
		PreViewFunc = True '防止按下按钮关闭对话框窗口
		Exit Function
	End If

	Select Case Action%
	Case 1
		LineNum = 20
		DlgText "LineNum",CStr(LineNum)
		DlgValue "Sync",1
		File = DlgText("FileName")
		For i = LBound(FileDataList) To UBound(FileDataList)
			TempList = Split(FileDataList(i),JoinStr)
			If TempList(0) = File Then
				Code = TempList(1)
				DlgValue "CovertType",StrToLong(TempList(2))
				DlgText "StartLine",TempList(3)
				DlgText "Break",TempList(4)
				DlgText "PreAddStr",TempList(5)
				DlgText "PreAppStr",TempList(6)
				DlgText "AppAddStr",TempList(7)
				DlgText "AppAppStr",TempList(8)
				DlgText "RepStr",TempList(9)
				DlgText "RepAsStr",TempList(10)
				FileNo = i
				Exit For
			End If
		Next i
		If Code = "" Then Code = CheckCode(File)
		For i = LBound(CodeList) To UBound(CodeList)
			TempArray = Split(CodeList(i),JoinStr)
			If TempArray(1) = Code Then
				DlgValue "CodeNameList",i
				Exit For
			End If
		Next i
		FileText = ReadFile(File,Code)
		If FileText <> "" Then
			FileTextLines = Split(FileText,vbCrLf,-1)
			For i = LBound(FileTextLines) To UBound(FileTextLines)
				l$ = FileTextLines(i)
				If i <> 0 Then inText = inText & vbCrLf & l$
				If i = 0 Then inText = l$
				If i = LineNum Then Exit For
			Next i
			DlgText "InTextBox",inText
		End If
    	Sync = DlgValue("Sync")
    	inText = DlgText("InTextBox")
    	If Sync = 1 And inText <> "" Then
    		outText = CovertText(inText,FileDataList(FileNo),m)
			DlgText "OutTextBox",outText
    	End If
		If DlgText("InTextBox") <> "" Then
    		DlgEnable "TestButton",True
    		DlgText "ClearButton",MsgList(0)
    	Else
    		DlgEnable "TestButton",False
    		DlgText "ClearButton",MsgList(1)
    	End If
		If UBound(FileDataList) = 0 Then
			DlgEnable "PreviousButton",False
			DlgEnable "NextButton",False
		ElseIf FileNo = 0 Then
			DlgEnable "PreviousButton",False
			DlgEnable "NextButton",True
		ElseIf FileNo = UBound(FileDataList) Then
			DlgEnable "PreviousButton",True
			DlgEnable "NextButton",False
		Else
    		DlgEnable "PreviousButton",True
    		DlgEnable "NextButton",True
    	End If
	Case 2 ' 数值更改或者按下了按钮
		If DlgItem$ = "TestButton" Then
			File = DlgText("FileName")
			Sync = DlgValue("Sync")
			inText = DlgText("InTextBox")
    		If inText <> "" Then
				ReDim TempList(9)
				CodeID = DlgValue("CodeNameList")
				TempArray = Split(CodeList(CodeID),JoinStr)
				TempList(0) = TempArray(1)
				TempList(1) = DlgValue("CovertType")
				TempList(2) = DlgText("StartLine")
				TempList(3) = DlgText("Break")
				TempList(4) = DlgText("PreAddStr")
				TempList(5) = DlgText("PreAppStr")
				TempList(6) = DlgText("AppAddStr")
				TempList(7) = DlgText("AppAppStr")
				TempList(8) = DlgText("RepStr")
				TempList(9) = DlgText("RepAsStr")
    			CovertData = File & JoinStr & Join(TempList,JoinStr)
    			outText = CovertText(inText,CovertData,m)
				DlgText "OutTextBox",outText
    		End If
		End If

		If DlgItem$ = "CovertType" Then
			File = DlgText("FileName")
			Sync = DlgValue("Sync")
			inText = DlgText("InTextBox")
    		If Sync = 1 And inText <> "" Then
				ReDim TempList(9)
				CodeID = DlgValue("CodeNameList")
				TempArray = Split(CodeList(CodeID),JoinStr)
				TempList(0) = TempArray(1)
				TempList(1) = DlgValue("CovertType")
				TempList(2) = DlgText("StartLine")
				TempList(3) = DlgText("Break")
				TempList(4) = DlgText("PreAddStr")
				TempList(5) = DlgText("PreAppStr")
				TempList(6) = DlgText("AppAddStr")
				TempList(7) = DlgText("AppAppStr")
				TempList(8) = DlgText("RepStr")
				TempList(9) = DlgText("RepAsStr")
    			CovertData = File & JoinStr & Join(TempList,JoinStr)
    			outText = CovertText(inText,CovertData,m)
				DlgText "OutTextBox",outText
    		End If
    		For i = LBound(FileDataList) To UBound(FileDataList)
				TempList = Split(FileDataList(i),JoinStr)
				If TempList(0) = File Then
					TempList(2) = DlgValue("CovertType")
					FileDataList(i) = Join(TempList,JoinStr)
					Exit For
				End If
			Next i
		End If

		If DlgItem$ = "ClearButton" Then
			If DlgText("ClearButton") = MsgList(0) Then
				DlgText "InTextBox",""
				DlgText "OutTextBox",""
				DlgText "ClearButton",MsgList(1)
				DlgEnable "TestButton",False
			Else
				File = DlgText("FileName")
				If FileText = "" Then
					CodeID = DlgValue("CodeNameList")
					TempList = Split(CodeList(CodeID),JoinStr)
					FileText = ReadFile(File,TempList(1))
				Else
					LineNum = StrToLong(DlgText("LineNum"))
					FileTextLines = Split(FileText,vbCrLf,-1)
					For i = LBound(FileTextLines) To UBound(FileTextLines)
						l$ = FileTextLines(i)
						If i <> 0 Then inText = inText & vbCrLf & l$
						If i = 0 Then inText = l$
						If i = LineNum Then Exit For
					Next i
					DlgText "InTextBox",inText
    			End If

    			Sync = DlgValue("Sync")
    			inText = DlgText("InTextBox")
    			If Sync = 1 And inText <> "" Then
					ReDim TempList(9)
					CodeID = DlgValue("CodeNameList")
					TempArray = Split(CodeList(CodeID),JoinStr)
					TempList(0) = TempArray(1)
					TempList(1) = DlgValue("CovertType")
					TempList(2) = DlgText("StartLine")
					TempList(3) = DlgText("Break")
					TempList(4) = DlgText("PreAddStr")
					TempList(5) = DlgText("PreAppStr")
					TempList(6) = DlgText("AppAddStr")
					TempList(7) = DlgText("AppAppStr")
					TempList(8) = DlgText("RepStr")
					TempList(9) = DlgText("RepAsStr")
					CovertData = File & JoinStr & Join(TempList,JoinStr)
    				outText = CovertText(inText,CovertData,m)
					DlgText "OutTextBox",outText
    			End If
				DlgText "ClearButton",MsgList(0)
				DlgEnable "TestButton",True
    		End If
		End If

		If DlgItem$ = "CodeNameList" Then
			File = DlgText("FileName")
			CodeID = DlgValue("CodeNameList")
			TempList = Split(CodeList(CodeID),JoinStr)
			Code = TempList(1)
			If Code = "_autodetect_all" Or Code = "_autodetect" Or Code = "_autodetect_kr" Then
				Code = CheckCode(File)
				For i = LBound(CodeList) To UBound(CodeList)
					TempList = Split(CodeList(i),JoinStr)
					If TempList(1) = Code Then
 						DlgValue "CodeNameList",i
 						Exit For
 					End If
				Next i
			End If
			FileText = ReadFile(File,Code)
			If FileText <> "" Then
				LineNum = StrToLong(DlgText("LineNum"))
				FileTextLines = Split(FileText, vbCrLf, -1)
				For i = LBound(FileTextLines) To UBound(FileTextLines)
					l$ = FileTextLines(i)
					If i <> 0 Then inText = inText & vbCrLf & l$
					If i = 0 Then inText = l$
					If i = LineNum Then Exit For
				Next i
				For i = LBound(FileDataList) To UBound(FileDataList)
					TempList = Split(FileDataList(i),JoinStr)
					If TempList(0) = File Then
						TempList(1) = Code
						FileDataList(i) = Join(TempList,JoinStr)
						Exit For
					End If
				Next i
				DlgText "InTextBox",inText
			Else
				DlgText "InTextBox",FileText
    		End If

        	Sync = DlgValue("Sync")
        	inText = DlgText("InTextBox")
        	If Sync = 1 And inText <> "" Then
        		ReDim TempList(9)
				TempList(0) = Code
				TempList(1) = DlgValue("CovertType")
				TempList(2) = DlgText("StartLine")
				TempList(3) = DlgText("Break")
				TempList(4) = DlgText("PreAddStr")
				TempList(5) = DlgText("PreAppStr")
				TempList(6) = DlgText("AppAddStr")
				TempList(7) = DlgText("AppAppStr")
				TempList(8) = DlgText("RepStr")
				TempList(9) = DlgText("RepAsStr")
				CovertData = File & JoinStr & Join(TempList,JoinStr)
    			outText = CovertText(inText,CovertData,m)
				DlgText "OutTextBox",outText
    		End If
    	End If

		If DlgItem$ = "PreviousButton" Or DlgItem$ = "NextButton" Then
			File = DlgText("FileName")
			For i = LBound(FileDataList) To UBound(FileDataList)
				TempList = Split(FileDataList(i),JoinStr)
				If TempList(0) = File Then
					FileNo = i
					Exit For
				End If
			Next i
			i = FileNo
			If DlgItem$ = "PreviousButton" And FileNo <> 0 Then FileNo = FileNo - 1
			If DlgItem$ = "NextButton" And FileNo < UBound(FileDataList) Then FileNo = FileNo + 1
			If FileNo <> i Then
				TempList = Split(FileDataList(FileNo),JoinStr)
				File = TempList(0)
				DlgText "FileName",File
				Code = TempList(1)
				If Code = "" Then Code = CheckCode(File)
				For i = LBound(CodeList) To UBound(CodeList)
					TempArray = Split(CodeList(i),JoinStr)
					If TempArray(1) = Code Then
 						DlgValue "CodeNameList",i
 						Exit For
 					End If
				Next i
				DlgValue "CovertType",StrToLong(TempList(2))
				DlgText "StartLine",TempList(3)
				DlgText "Break",TempList(4)
				DlgText "PreAddStr",TempList(5)
				DlgText "PreAppStr",TempList(6)
				DlgText "AppAddStr",TempList(7)
				DlgText "AppAppStr",TempList(8)
				DlgText "RepStr",TempList(9)
				DlgText "RepAsStr",TempList(10)
				FileText = ReadFile(File,Code)
				If FileText <> "" Then
					LineNum = StrToLong(DlgText("LineNum"))
					FileTextLines = Split(FileText, vbCrLf, -1)
					For i = LBound(FileTextLines) To UBound(FileTextLines)
						l$ = FileTextLines(i)
						If i <> 0 Then inText = inText & vbCrLf & l$
						If i = 0 Then inText = l$
						If i = LineNum Then Exit For
					Next i
					DlgText "InTextBox",inText
				Else
					DlgText "InTextBox",FileText
				End If
        		Sync = DlgValue("Sync")
        		inText = DlgText("InTextBox")
        		If Sync = 1 And inText <> "" Then
					CovertData = Join(TempList,JoinStr)
    				outText = CovertText(inText,CovertData,m)
					DlgText "OutTextBox",outText
				End If
    		End If
    		If UBound(FileDataList) = 0 Then
				DlgEnable "PreviousButton",False
				DlgEnable "NextButton",False
			ElseIf FileNo = 0 Then
				DlgEnable "PreviousButton",False
				DlgEnable "NextButton",True
			ElseIf FileNo = UBound(FileDataList) Then
				DlgEnable "PreviousButton",True
				DlgEnable "NextButton",False
			Else
    			DlgEnable "PreviousButton",True
    			DlgEnable "NextButton",True
    		End If
    	End If

    	If DlgItem$ = "OKButton" Then
    		If Join(FileDataList,JoinStr) <> Join(CovertDataList,JoinStr) Then
    			Msg = MsgBox(MsgList(3),vbYesNoCancel+vbInformation,MsgList(2))
    			If Msg = vbNo Then FileDataList = CovertDataList
    			If Msg = vbCancel Then
    				CovertTestFunc = True '防止按下按钮关闭对话框窗口
    			End If
    		End If
    	End If

		If DlgItem$ <> "OKButton" And DlgItem$ <> "CancelButton" Then
		    inText = DlgText("InTextBox")
		    If inText = "" Then DlgText "OutTextBox",""
		    If inText <> "" Then
    			DlgEnable "TestButton",True
    			DlgText "ClearButton",MsgList(0)
    		Else
    			DlgEnable "TestButton",False
    			DlgText "ClearButton",MsgList(1)
    		End If
			CovertTestFunc = True '防止按下按钮关闭对话框窗口
		End If
	Case 3 ' 文本框或者组合框文本被更改
		File = DlgText("FileName")
		If DlgItem$ = "LineNum" Or DlgItem$ = "StartLine" Then
			If FileText = "" Then
				CodeID = DlgValue("CodeNameList")
				TempArray = Split(CodeList(CodeID),JoinStr)
				FileText = ReadFile(File,TempArray(1))
			End If
			If FileText <> "" Then
				LineNum = StrToLong(DlgText("LineNum"))
				FileTextLines = Split(FileText, vbCrLf, -1)
				For i = LBound(FileTextLines) To UBound(FileTextLines)
					l$ = FileTextLines(i)
					If i <> 0 Then inText = inText & vbCrLf & l$
					If i = 0 Then inText = l$
					If i = LineNum Then Exit For
				Next i
				DlgText "InTextBox",inText
			Else
				DlgText "InTextBox",FileText
			End If
		End If

   		If DlgItem$ <> "OutTextBox" Then
   			Sync = DlgValue("Sync")
   			inText = DlgText("InTextBox")
   			If Sync = 1 And inText <> "" Then
   				ReDim TempList(9)
				CodeID = DlgValue("CodeNameList")
				TempArray = Split(CodeList(CodeID),JoinStr)
				TempList(0) = TempArray(1)
				TempList(1) = DlgValue("CovertType")
				TempList(2) = DlgText("StartLine")
				TempList(3) = DlgText("Break")
				TempList(4) = DlgText("PreAddStr")
				TempList(5) = DlgText("PreAppStr")
				TempList(6) = DlgText("AppAddStr")
				TempList(7) = DlgText("AppAppStr")
				TempList(8) = DlgText("RepStr")
				TempList(9) = DlgText("RepAsStr")
				CovertData = File & JoinStr & Join(TempList,JoinStr)
   				outText = CovertText(inText,CovertData,m)
				DlgText "OutTextBox",outText
			End If

			inText = DlgText("InTextBox")
			If inText = "" Then DlgText "OutTextBox",""
			If inText <> "" Then
    			DlgEnable "TestButton",True
    			DlgText "ClearButton",MsgList(0)
    		Else
    			DlgEnable "TestButton",False
    			DlgText "ClearButton",MsgList(1)
    		End If
    	End If

		If DlgItem$ <> "inTextBox" And DlgItem$ <> "OutTextBox" Then
    		inText = DlgText("InTextBox")
   			ReDim TempList(8)
			TempList(0) = DlgValue("CovertType")
			TempList(1) = DlgText("StartLine")
			TempList(2) = DlgText("Break")
			TempList(3) = DlgText("PreAddStr")
			TempList(4) = DlgText("PreAppStr")
			TempList(5) = DlgText("AppAddStr")
			TempList(6) = DlgText("AppAppStr")
			TempList(7) = DlgText("RepStr")
			TempList(8) = DlgText("RepAsStr")
			CodeID = DlgValue("CodeNameList")
			TempArray = Split(CodeList(CodeID),JoinStr)
			Temp = TempArray(1) & JoinStr & Join(TempList,JoinStr)
			For i = LBound(FileDataList) To UBound(FileDataList)
				TempList = Split(FileDataList(i),JoinStr)
				If TempList(0) = File Then
					FileDataList(i) = File & JoinStr & Temp
					Exit For
				End If
			Next i
		End If
	End Select
End Function


'编辑文件
'fType = 0 编辑模式，如果打开文件成功返回字符编码和 True
'fType = 1 查看和确认字符编码模式，如果打开文件成功并按 [确定] 按钮返回字符编码和 True
Function EditFile(File As String,FileDataList() As String,fType As Long) As Boolean
	Dim i As Long,CodeNameList() As String,MsgList() As String
	Dim FileDataListBak() As String

	EditFile = False
	If getMsgList(UILangList,MsgList,"EditFile",1) = False Then Exit Function
	If fType = 1 Then
		MsgList(0) = MsgList(9)
		MsgList(8) = MsgList(10)
	End If

    'Dim objStream As Object
	'Set objStream = CreateObject("Adodb.Stream")
	'If objStream Is Nothing Then CodeList = CodePageList(0,0)
	'If Not objStream Is Nothing Then CodeList = CodePageList(0,49)
	'Set objStream = Nothing

	For i = LBound(CodeList) To UBound(CodeList)
		ReDim Preserve CodeNameList(i)
		TempArray = Split(CodeList(i),JoinStr)
		CodeNameList(i) = TempArray(0)
	Next i
	FileDataListBak = FileDataList

	Begin Dialog UserDialog 820,518,MsgList(0),.EditFileFunc ' %GRID:10,7,1,1
		Text 10,7,800,14,File,.FileName,2
		Text 510,28,80,14,MsgList(1),.CodeText
		DropListBox 600,24,210,21,CodeNameList(),.CodeNameList
		OptionGroup .Options
			OptionButton 290,119,90,14,"",.OptionButton1
			OptionButton 420,119,90,14,"",.OptionButton2
		TextBox 0,49,820,434,.InTextBox,1
		Text 10,28,80,14,MsgList(2),.FindText
		TextBox 100,25,310,19,.FindBox
		PushButton 420,24,80,21,MsgList(3),.FindButton
		PushButton 20,490,90,21,MsgList(4),.ReadButton
		PushButton 120,490,90,21,MsgList(5),.PreviousButton
		PushButton 220,490,90,21,MsgList(6),.NextButton
		PushButton 380,490,140,21,MsgList(7),.EditButton
		PushButton 600,490,100,21,MsgList(8),.SaveButton
		PushButton 710,490,90,21,MsgList(11),.ExitButton
		CancelButton 710,490,90,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	dlg.Options = fType
	If Dialog(dlg) = 0 Then
		FileDataList = FileDataListBak
		Exit Function
	End If
	EditFile = True
End Function


'编辑对话框函数
Private Function EditFileFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim File As String,inText As String,outText As String
	Dim i As Long,n As Long,FileNo As Long,CodeID As Long
	Dim Code As String,TempList() As String,MsgList() As String

	If getMsgList(UILangList,MsgList,"EditFileFunc",1) = False Then
		EditFileFunc = True '防止按下按钮关闭对话框窗口
		Exit Function
	End If

	Select Case Action%
	Case 1
		DlgVisible "Options",False
		If DlgValue("Options") = 0 Then
			DlgVisible "CancelButton",False
		Else
			DlgVisible "ExitButton",False
		End If
		File = DlgText("FileName")
		For i = LBound(FileDataList) To UBound(FileDataList)
			TempArray = Split(FileDataList(i),JoinStr)
			If TempArray(0) = File Then
				Code = TempArray(1)
				FileNo = i
				Exit For
			End If
		Next i
		If Code = "" Then Code = CheckCode(File)
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
			DlgText "ReadButton",MsgList(9)
    		DlgEnable "FindButton",True
			DlgEnable "SaveButton",True
			DlgEnable "EditButton",False
    	Else
    		DlgText "ReadButton",MsgList(8)
    		DlgEnable "FindButton",False
    		DlgEnable "SaveButton",False
    		DlgEnable "EditButton",False
    	End If
    	If UBound(FileDataList) = 0 Then
			DlgEnable "PreviousButton",False
			DlgEnable "NextButton",False
		ElseIf FileNo = 0 Then
			DlgEnable "PreviousButton",False
			DlgEnable "NextButton",True
		ElseIf FileNo = UBound(FileDataList) Then
			DlgEnable "PreviousButton",True
			DlgEnable "NextButton",False
		Else
    		DlgEnable "PreviousButton",True
    		DlgEnable "NextButton",True
    	End If
	Case 2 ' 数值更改或者按下了按钮
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
			End If
			FileText = ReadFile(File,Code)
			DlgText "InTextBox",FileText
			FindText = ""
			FindLine = ""
			If FileText <> "" Then
				For i = LBound(FileDataList) To UBound(FileDataList)
					TempArray = Split(FileDataList(i),JoinStr)
					If TempArray(0) = File Then
						TempArray(1) = Code
						FileDataList(i) = Join(TempArray,JoinStr)
						Exit For
					End If
				Next i
			End If
		End If

		If DlgItem$ = "ReadButton" Then
			FindText = ""
			FindLine = ""
			If DlgText("ReadButton") = MsgList(8) Then
				File = DlgText("FileName")
				CodeID = DlgValue("CodeNameList")
				TempArray = Split(CodeList(CodeID),JoinStr)
				Code = TempArray(1)
				FileText = ReadFile(File,Code)
				DlgText "InTextBox",FileText
				If FindText <> "" Then DlgText "ReadButton",MsgList(9)
			Else
				DlgText "InTextBox",""
				DlgText "ReadButton",MsgList(8)
			End If
		End If

		If DlgItem$ = "FindButton" And DlgText("FindBox") <> "" Then
			outText = FindText
			inText = FindLine
			FindText = ""
			FindLine = ""
			n = 0
			toFindText = "*" & DlgText("FindBox") & "*"
			InTextArray = Split(FileText,vbCrLf,-1)
			For i = LBound(InTextArray) To UBound(InTextArray)
				tempText = InTextArray(i)
				If tempText Like toFindText Then
					Temp = "【" & i+1 & MsgList(10) & "】" & tempText
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
				FindText = outText
				FindLine = inText
				MsgBox(MsgList(4),vbOkOnly+vbInformation,MsgList(0))
			End If
    	End If

		If DlgItem$ = "EditButton" And DlgText("InTextBox") <> "" Then
			inText = DlgText("InTextBox")
			If FindLine <> "" And inText <> FindText Then
				If MsgBox(MsgList(2),vbYesNo+vbInformation,MsgList(0)) = vbYes Then
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
							Temp = MsgList(10) & "】"
							LineNoStr = Left(NewString,InStr(NewString,Temp)+1)
							Temp = "【" & "*" & MsgList(10) & "】"
							If LineNoStr Like Temp Then
								NewString = Mid(NewString,Len(LineNoStr)+1)
								InTextArray(OldLineNum) = NewString
							End If
						Next i
						inText = Join(InTextArray,vbCrLf)
						DlgText "InTextBox",inText
						FileText = inText
						FindText = ""
						FindLine = ""
					Else
						MsgBox(MsgList(3),vbOkOnly+vbInformation,MsgList(0))
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
			File = DlgText("FileName")
			For i = LBound(FileDataList) To UBound(FileDataList)
				TempArray = Split(FileDataList(i),JoinStr)
				If TempArray(0) = File Then
					FileNo = i
					Exit For
				End If
			Next i
			i = FileNo
			If DlgItem$ = "PreviousButton" And FileNo <> 0 Then FileNo = FileNo - 1
			If DlgItem$ = "NextButton" And FileNo < UBound(FileDataList) Then FileNo = FileNo + 1
			If FileNo <> i Then
				TempArray = Split(FileDataList(FileNo),JoinStr)
				File = TempArray(0)
				DlgText "FileName",File
				Code = TempArray(1)
				If Code = "" Then Code = CheckCode(File)
				For i = LBound(CodeList) To UBound(CodeList)
					TempArray = Split(CodeList(i),JoinStr)
					If TempArray(1) = Code Then
 						DlgValue "CodeNameList",i
 						Exit For
 					End If
				Next i
				FileText = ReadFile(File,Code)
				DlgText "InTextBox",FileText
				FindText = ""
				FindLine = ""
    		End If
			If UBound(FileDataList) = 0 Then
				DlgEnable "PreviousButton",False
				DlgEnable "NextButton",False
			ElseIf FileNo = 0 Then
				DlgEnable "PreviousButton",False
				DlgEnable "NextButton",True
			ElseIf FileNo = UBound(FileDataList) Then
				DlgEnable "PreviousButton",True
				DlgEnable "NextButton",False
			Else
    			DlgEnable "PreviousButton",True
    			DlgEnable "NextButton",True
    		End If
    	End If

		If DlgItem$ = "SaveButton" And DlgText("InTextBox") <> "" Then
			If DlgValue("Options") = 0 Then
				File = DlgText("FileName")
				inText = DlgText("InTextBox")
				CodeID = DlgValue("CodeNameList")
				TempArray = Split(CodeList(CodeID),JoinStr)
				Code = TempArray(1)
				If Dir(File) <> "" Then SetAttr File,vbNormal
				If WriteToFile(File,inText,Code) = True Then
					MsgBox(MsgList(5),vbOkOnly+vbInformation,MsgList(0))
					FileText = inText
					FindText = ""
					FindLine = ""
				Else
					MsgBox(MsgList(6),vbOkOnly+vbInformation,MsgList(0))
				End If
			Else
				Exit Function
			End If
    	End If

		If DlgItem$ = "CancelButton" And DlgText("InTextBox") <> "" Then
			File = DlgText("FileName")
			inText = DlgText("InTextBox")
			If inText <> FileText And DlgValue("Options") = 0 Then
				If MsgBox(MsgList(1),vbYesNo+vbInformation,MsgList(0)) = vbYes Then
					If Dir(File) <> "" Then SetAttr File,vbNormal
					CodeID = DlgValue("CodeNameList")
					TempArray = Split(CodeList(CodeID),JoinStr)
					Code = TempArray(1)
					If WriteToFile(File,inText,Code) = True Then
						MsgBox(MsgList(5),vbOkOnly+vbInformation,MsgList(0))
					Else
						MsgBox(MsgList(6),vbOkOnly+vbInformation,MsgList(0))
					End If
				End If
			End If
			FileText = inText
			FindText = ""
			FindLine = ""
    	End If

		If DlgItem$ <> "CancelButton" And DlgItem$ <> "ExitButton" Then
		    inText = DlgText("InTextBox")
		    If inText <> "" Then
    			DlgText "ReadButton",MsgList(9)
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
    			DlgText "ReadButton",MsgList(8)
    			DlgEnable "FindButton",False
    			DlgEnable "SaveButton",False
    			DlgEnable "EditButton",False
    			DlgEnable "CancelButton",True
    		End If
			EditFileFunc = True '防止按下按钮关闭对话框窗口
		End If
	Case 3 ' 文本框或者组合框文本被更改
		If DlgItem$ = "InTextBox" Then
		    inText = DlgText("InTextBox")
		    If inText <> "" Then
    			DlgText "ReadButton",MsgList(9)
    			DlgEnable "FindButton",True
				If FindLine = "" Then
					DlgEnable "SaveButton",True
					DlgEnable "EditButton",False
				Else
					DlgEnable "SaveButton",False
					DlgEnable "EditButton",True
				End If
    		Else
    			DlgText "ReadButton",MsgList(8)
    			DlgEnable "FindButton",False
    			DlgEnable "SaveButton",False
    			DlgEnable "EditButton",False
    		End If
    		If DlgEnable("EditButton") = False Then FileText = inText
    	End If
	End Select
End Function


'输入编辑程序
Sub CommandInput(CmdPath As String,Argument As String)
	Dim MsgList() As String
	If getMsgList(UILangList,MsgList,"CommandInput",1) = False Then Exit Sub
	Begin Dialog UserDialog 540,294,MsgList(0),.CommandInputFunc ' %GRID:10,7,1,1
		Text 10,7,520,140,MsgList(1)
		Text 10,154,490,14,MsgList(2)
		TextBox 10,175,490,21,.CmdPath
		PushButton 500,175,30,21,MsgList(3),.BrowseButton
		Text 10,210,490,14,MsgList(4)
		TextBox 10,231,490,21,.Argument
		PushButton 500,231,30,21,MsgList(6),.FileArgButton
		PushButton 20,266,90,21,MsgList(5),.ClearButton
		OKButton 300,266,100,21,.OKButton
		CancelButton 420,266,100,21,.CancelButton
	End Dialog
	Dim dlg As UserDialog
	dlg.CmdPath = CmdPath
	dlg.Argument = Argument
	If Dialog(dlg) = 0 Then Exit Sub
	CmdPath = dlg.CmdPath
	Argument = dlg.Argument
End Sub


'获取编辑程序对话框函数
Private Function CommandInputFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim x As Long,File As String,Items(0) As String,MsgList() As String

	If Action% < 3 Then
		If getMsgList(UILangList,MsgList,"CommandInputFunc",1) = False Then
			CommandInputFunc = True '防止按下按钮关闭对话框窗口
			Exit Function
		End If
	End If

	Select Case Action%
	Case 1 ' 对话框窗口初始化
		DlgEnable "ClearButton",False
	Case 2 ' 数值更改或者按下了按钮
		If DlgItem$ = "BrowseButton" Then
			If PSL.SelectFile(File,True,MsgList(2),MsgList(1)) = True Then
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
			Items(0) = MsgList(4)
			x = ShowPopupMenu(Items)
			If x = 0 Then
				Argument = DlgText("Argument")
				DlgText "Argument",Argument & " " & """%1"""
			End If
		End If
		If DlgItem$ = "OKButton" Then
			If DlgText("CmdPath") = "" Then
 				MsgBox MsgList(3),vbOkOnly+vbInformation,MsgList(0)
				CommandInputFunc = True ' 防止按下按钮关闭对话框窗口
			End If
		End If
		If DlgItem$ <> "OKButton" And DlgItem$ <> "CancelButton" Then
 			If DlgText("CmdPath") = "" And DlgText("Argument") = "" Then
 				DlgEnable "ClearButton",False
 			Else
 				DlgEnable "ClearButton",True
 			End If
 			CommandInputFunc = True ' 防止按下按钮关闭对话框窗口
		End If
	Case 3 ' 文本框或者组合框文本被更改
 		If DlgText("CmdPath") = "" And DlgText("Argument") = "" Then
 			DlgEnable "ClearButton",False
 		Else
 			DlgEnable "ClearButton",True
 		End If
	End Select
End Function


'读取当前文件夹中的每个文件
Function GetFiles(Folder As String,ExtList As String,sf As Long) As String()
	Dim i As Long,j As Long,objFS As Object,gFiles() As String
	ReDim gFiles(0)
	gFiles(0) = ""
	j = 0
	On Error Resume Next
	Set objFS = CreateObject("Scripting.FileSystemObject")
	On Error GoTo 0
	If objFS Is Nothing Then
		ExtListArry = Split(ExtList,";",-1)
		For i = LBound(ExtListArry) To UBound(ExtListArry)
			If Trim(Join(gFiles)) = "" Then j = 0 Else j = UBound(gFiles) + 1
			ExtName = ExtListArry(i)
			File = Dir$(Folder & "*." & ExtName)
			Do While File <> ""
				sExtName = Mid(File,InStrRev(File,".")+1)
				If InStr(ExtName,"?") Or InStr(ExtName,"*") Then
					If UCase(sExtName) Like UCase(ExtName) Then
						ReDim Preserve gFiles(j)
						gFiles(j) = Folder & File
						j = j + 1
					End If
				Else
					If UCase(sExtName) = UCase(ExtName) Then
						ReDim Preserve gFiles(j)
						gFiles(j) = Folder & File
						j = j + 1
					End If
				End If
				File = Dir$()
			Loop
			If sf = 1 Then
				Call FindSubFiles(Folder,ExtName,gFiles)
			End If
		Next i
	Else
		ExtListArry = Split(ExtList,";",-1)
		For i = LBound(ExtListArry) To UBound(ExtListArry)
			If Trim(Join(gFiles)) = "" Then j = 0 Else j = UBound(gFiles) + 1
			ExtName = ExtListArry(i)
			Set objFolder = objFS.GetFolder(Folder)
			For Each File In objFolder.Files
				sExtName = objFS.GetExtensionName(File.Path)
				If InStr(ExtName,"?") Or InStr(ExtName,"*") Then
					If UCase(sExtName) Like UCase(ExtName) Then
						ReDim Preserve gFiles(j)
						gFiles(j) = File.Path
						j = j + 1
					End If
				Else
					If UCase(sExtName) = UCase(ExtName) Then
						ReDim Preserve gFiles(j)
						gFiles(j) = File.Path
						j = j + 1
					End If
				End If
			Next File
			If sf = 1 Then
				Call GetSubFiles(objFS,Folder,ExtName,gFiles)
			End If
		Next i
	End If
	Set objFS = Nothing
	GetFiles = gFiles
End Function


'读取子文件夹中的每个文件
Sub FindSubFiles(Folder As String,ExtName As String,gFiles() As String)
	Dim i As Long,j As Long,k As Long,File As String,Path As String
	Dim sFolder As String,subFolders() As String
	ReDim subFolders(0)
	subFolders(0) = Folder
	If Trim(Join(gFiles)) = "" Then j = 0 Else j = UBound(gFiles) + 1
	i = 0
	k = 0
	Do
		sFolder = subFolders(i)
		File = Dir$(sFolder & "*.*",vbDirectory)
		While File <> ""
			If File <> "." And File <> ".." Then
				If GetAttr(sFolder & File) And vbDirectory Then
            	 	k = k + 1
            	 	ReDim Preserve subFolders(k)
             		subFolders(k) = sFolder & File & "\"
				ElseIf sFolder <> Folder Then
					sExtName = Mid(File,InStrRev(File,".")+1)
					If InStr(ExtName,"?") Or InStr(ExtName,"*") Then
						If UCase(sExtName) Like UCase(ExtName) Then
							ReDim Preserve gFiles(j)
							gFiles(j) = sFolder & File
							j = j + 1
						End If
					Else
						If UCase(sExtName) = UCase(ExtName) Then
							ReDim Preserve gFiles(j)
							gFiles(j) = sFolder & File
							j = j + 1
						End If
					End If
				End If
			End If
			File = Dir$()
		Wend
		i = i + 1
	Loop Until i = k + 1
End Sub


'读取子文件夹中的每个文件
Sub GetSubFiles(objFS As Object,Folder As String,ExtName As String,gFiles() As String)
	Dim j As Long,sExtName As String
	Set objFolder = objFS.GetFolder(Folder)
	Set colSubFolders = objFolder.SubFolders
	For Each objSubFolder In colSubFolders
		If Trim(Join(gFiles)) = "" Then j = 0 Else j = UBound(gFiles) + 1
		For Each File In objSubFolder.Files
			sExtName = objFS.GetExtensionName(File.Path)
			If InStr(ExtName,"?") Or InStr(ExtName,"*") Then
				If UCase(sExtName) Like UCase(ExtName) Then
					ReDim Preserve gFiles(j)
					gFiles(j) = File.Path
					j = j + 1
				End If
			Else
				If UCase(sExtName) = UCase(ExtName) Then
					ReDim Preserve gFiles(j)
					gFiles(j) = File.Path
					j = j + 1
				End If
			End If
		Next File
		Call GetSubFiles(objFS,objSubFolder,ExtName,gFiles)
	Next objSubFolder
End Sub


'转换文本
Function CovertText(inText As String,CovertData As String,n As Long) As String
	Dim i As Long,j As Long,k As Long,pos As Long,CovertType As Long,StartLine As Long
	Dim TempList() As String,pStemp As Boolean,aStemp As Boolean,Stemp As Boolean

	TempList = Split(CovertData,JoinStr,-1)
	CovertType = StrToLong(TempList(2))
	StartLine = StrToLong(TempList(3))
	Break = Convert(TempList(4))
	PreAddStr = Convert(TempList(5))
	PreAppStr = Convert(TempList(6))
	AppAddStr = Convert(TempList(7))
	AppAppStr = Convert(TempList(8))
	RepStr = Convert(TempList(9))
	RepAsStr = Convert(TempList(10))

	If StartLine = 0 Then StartLine = 1
	FileLines = Split(inText,vbCrLf,-1)
    n = 0
    For i = StartLine - 1 To UBound(FileLines)
    	Stemp = False
		l$ = FileLines(i)
		If InStr(l$,Break) Then
			pStemp = False
			aStemp = False
			PreTxt = ""
			AppTxt = ""
			pos = InStr(l$,Break)
			k = pos
			Do While pos > 0
				PreTxt = Mid(l$,1,pos - 1)
				AppTxt = Mid(l$,pos + Len(Break))
				tPreTxt = Trim(PreTxt)
				tAppTxt = Trim(AppTxt)
				For j = 0 To 1
					If j = 0 Then Marks = """"
					If j = 1 Then Marks = "'"
					If Left(tPreTxt,1) = Marks And Right(tPreTxt,1) = Marks Then
						pStemp = True
					End If
					If Left(tAppTxt,1) = Marks And Right(tAppTxt,1) = Marks Then
						aStemp = True
					End If
					If pStemp = True Or aStemp = True Then Exit For
				Next j
				If pStemp = True Or aStemp = True Then
					k = 1
					Exit Do
				End If
				pos = InStr(pos + Len(Break),l$,Break)
				If pos = 0 Then k = 1
			Loop
			If k > 0 Then
				tPreTxt = RTrim(PreTxt)
				tAppTxt = LTrim(AppTxt)
				j = Len(PreTxt) - Len(tPreTxt)
				k = Len(AppTxt) - Len(tAppTxt)
				cBreak = Space(j) & Break & Space(k)
				If pStemp = True Then tPreTxt = RemoveBackslash(tPreTxt,Marks,Marks,2)
				If aStemp = True Then tAppTxt = RemoveBackslash(tAppTxt,Marks,Marks,2)
				If CovertType = 0 Then
					PreTxt = PreAddStr & tPreTxt & PreAppStr
					AppTxt = AppAddStr & tAppTxt & AppAppStr
					If pStemp = True Then PreTxt = AppendBackslash(PreTxt,Marks,Marks,0)
					If aStemp = True Then AppTxt = AppendBackslash(AppTxt,Marks,Marks,0)
				ElseIf CovertType = 1 Then
					PreTxt = PreAddStr & tPreTxt & PreAppStr
					AppTxt = AppAddStr & tPreTxt & AppAppStr
					If pStemp = True Then PreTxt = AppendBackslash(PreTxt,Marks,Marks,0)
					If pStemp = True Then AppTxt = AppendBackslash(AppTxt,Marks,Marks,0)
				ElseIf CovertType = 2 Then
					PreTxt = PreAddStr & tAppTxt & PreAppStr
					AppTxt = AppAddStr & tAppTxt & AppAppStr
					If aStemp = True Then PreTxt = AppendBackslash(PreTxt,Marks,Marks,0)
					If aStemp = True Then AppTxt = AppendBackslash(AppTxt,Marks,Marks,0)
				ElseIf CovertType = 3 Then
					PreTxt = PreAddStr & tAppTxt & PreAppStr
					AppTxt = AppAddStr & tPreTxt & AppAppStr
					If aStemp = True Then PreTxt = AppendBackslash(PreTxt,Marks,Marks,0)
					If pStemp = True Then AppTxt = AppendBackslash(AppTxt,Marks,Marks,0)
				ElseIf CovertType = 4 Then
					PreTxt = PreAddStr & "" & PreAppStr
					AppTxt = AppAddStr & tPreTxt & AppAppStr
					If pStemp = True Then AppTxt = AppendBackslash(AppTxt,Marks,Marks,0)
				ElseIf CovertType = 5 Then
					PreTxt = PreAddStr & tAppTxt & PreAppStr
					AppTxt = AppAddStr & "" & AppAppStr
					If aStemp = True Then PreTxt = AppendBackslash(PreTxt,Marks,Marks,0)
				ElseIf CovertType = 6 Then
					PreTxt = PreAddStr & "" & PreAppStr
					AppTxt = AppAddStr & tAppTxt & AppAppStr
					If pStemp = True Then PreTxt = AppendBackslash(PreTxt,Marks,Marks,0)
					If aStemp = True Then AppTxt = AppendBackslash(AppTxt,Marks,Marks,0)
				ElseIf CovertType = 7 Then
					PreTxt = PreAddStr & tPreTxt & PreAppStr
					AppTxt = AppAddStr & "" & AppAppStr
					If pStemp = True Then PreTxt = AppendBackslash(PreTxt,Marks,Marks,0)
					If aStemp = True Then AppTxt = AppendBackslash(AppTxt,Marks,Marks,0)
				End If
				l$ = PreTxt & cBreak & AppTxt
				Stemp = True
			End If
		End If
		If l$ <> "" Then
			If RepStr <> "" And InStr(l$,RepStr) Then
				l$ = Replace(l$,RepStr,RepAsStr)
				Stemp = True
			End If
			FileLines(i) = l$
			If Stemp = True Then n = n + 1
		End If
	Next i
	CovertText = Join(FileLines,vbCrLf)
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


'字串前后附加指定的 PreStr 和 AppStr
'fType = 0 不去除字串前后空格，但在字串前后附加指定的 PreStr 和 AppStr
'fType = 1 去除字串前后空格，并在字串前后附加指定的 PreStr 和 AppStr
Function AppendBackslash(Path As String,PreStr As String,AppStr As String,fType As Long) As String
	AppendBackslash = Path
	If fType = 1 Then AppendBackslash = Trim(AppendBackslash)
	If AppendBackslash = "" And PreStr = AppStr Then
		AppendBackslash = PreStr & AppendBackslash & AppStr
	Else
		If PreStr <> "" And Left(AppendBackslash,Len(PreStr)) <> PreStr Then
			AppendBackslash = PreStr & AppendBackslash
		End If
		If AppStr <> "" And Right(AppendBackslash,Len(AppStr)) <> AppStr Then
			AppendBackslash = AppendBackslash & AppStr
		End If
	End If
End Function


' 检查文件编码
' ----------------------------------------------------
' ANSI      无格式定义
' EFBB BF   UTF-8
' FFFE      UTF-16LE/UCS-2, Little Endian with BOM
' FEFF      UTF-16BE/UCS-2, Big Endian with BOM
' XX00 XX00 UTF-16LE/UCS-2, Little Endian without BOM
' 00XX 00XX UTF-16BE/UCS-2, Big Endian without BOM
' FFFE 0000 UTF-32LE/UCS-4, Little Endian with BOM
' 0000 FEFF UTF-32BE/UCS-4, Big Endian with BOM
' XX00 0000 UTF-32LE/UCS-4, Little Endian without BOM
' 0000 00XX UTF-32BE/UCS-4, Big Endian without BOM
' 上述中的 XX 表示任意十六进制字符

Function CheckCode(FilePath As String) As String
	Dim objStream As Object,i As Long,Code As String
	If FilePath = "" Then Exit Function
	On Error Resume Next
	Set objStream = CreateObject("Adodb.Stream")
	On Error GoTo 0
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


' 读取文件
Function ReadFile(FilePath As String,CharSet As String) As String
	Dim objStream As Object,Code As String
	If FilePath = "" Then Exit Function
	On Error Resume Next
	Set objStream = CreateObject("Adodb.Stream")
	On Error GoTo ErrorMsg
	Code = CharSet
	If Not objStream Is Nothing Then
		If Code = "" Then Code = CheckCode(FilePath)
		If Code = "" Then Code = "_autodetect_all"
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


' 写入文件
Function WriteToFile(FilePath As String,textStr As String,CharSet As String) As Boolean
	Dim objStream As Object,Code As String,CodeBak As String,m As Long,n As Long,Bin As Variant
	WriteToFile = False
	If FilePath = "" Or textStr = "" Then Exit Function
	On Error Resume Next
	Set objStream = CreateObject("Adodb.Stream")
	On Error GoTo ErrorMsg
	Code = CharSet
	If LCase(Code) = "_autodetect_all" Then Code = "ANSI"
	If Not objStream Is Nothing Then
		If Code = "" Then Code = CheckCode(FilePath)
		CodeBak = Code
		If Code = "utf-8EFBB" Then Code = "utf-8"
		If Code <> "ANSI" Then
			With objStream
				.Type = 2
				.Mode = 3
				.CharSet = Code
				.Open
				.WriteText textStr
				If CodeBak = "utf-16LE" Or CodeBak = "utf-8" Then
					If CodeBak = "utf-16LE" Then n = 2 Else n = 3
					.Position = 0
					.Type = 1
					m = .Size
					.Position = n
					Bin = .Read(m - n)
					.Position = 0
					.SetEOS
					.Write Bin
				End If
				.SaveToFile FilePath,2
				.Close
			End With
			WriteToFile = True
		End If
	End If
	If objStream Is Nothing Or Code = "ANSI" Then
		FN = FreeFile
		Open FilePath For Output As #FN
		Print #FN,textStr
		Close #FN
		WriteToFile = True
	End If
	If CharSet = "" Then CharSet = Code
	Set objStream = Nothing
	Exit Function

	ErrorMsg:
	Set objStream = Nothing
	Err.Source = "NotWriteFile"
	Err.Description = Err.Description & JoinStr & FilePath
	Call sysErrorMassage(Err,1)
End Function


' 输出转换结果
Sub CovertMassage(MsgType As String,InputFile As String,OutputFile As String,m As Long,n As Long)
	Dim CovertMsg As String,MsgList() As String
	If getMsgList(UILangList,MsgList,"CovertMassage",1) = False Then Exit Sub

	If MsgType = "Coverting" Then
		CovertMsg = MsgList(1)
	ElseIf MsgType = "Coverted" Then
		If m <> 0 And OutputFile = "" Then CovertMsg = MsgList(2)
		If m <> 0 And OutputFile <> "" Then CovertMsg = MsgList(3)
		If m = 0 Then CovertMsg = MsgList(4)
	ElseIf MsgType = "NotCoverted" Then
		CovertMsg = MsgList(5)
	ElseIf MsgType = "ExitBackupFile" Then
		CovertMsg = MsgList(6)
	ElseIf MsgType = "AllCoverted" Then
		If n <> 0 Then CovertMsg = MsgList(7)
		If n = 0 Then CovertMsg = MsgList(8)

	ElseIf MsgType = "Saved" Then
		CovertMsg = MsgList(9)
	ElseIf MsgType = "SaveError" Then
		CovertMsg = MsgList(10)

	ElseIf MsgType = "Restored" Then
		If m = 0 Then CovertMsg = MsgList(11)
		If m = 1 Then CovertMsg = MsgList(12)
	ElseIf MsgType = "AllRestored" Then
		If n <> 0 Then CovertMsg = MsgList(13)
		If n = 0 Then CovertMsg = MsgList(14)
	End If

	CovertMsg = Replace(CovertMsg,"%s",InputFile)
	CovertMsg = Replace(CovertMsg,"%d",OutputFile)
	CovertMsg = Replace(CovertMsg,"%n",CStr(n))
	If MsgType = "Saved" Or MsgType = "SaveError" Then
		MsgBox(CovertMsg,vbOkOnly+vbInformation,MsgList(0))
	Else
		PSL.Output CovertMsg
	End If
End Sub


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
				TitleMsg = MsgList(47)
				If fType <> 0 Then ContinueMsg = MsgList(49) Else ContinueMsg = MsgList(50)
				Msg = Replace(Replace(MsgList(48),"%s",ItemList$),"%d",LangFile)
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
		Msg = Msg & ContinueMsg
		If fType = 0 Then
			MsgBox(Msg,vbOkOnly+vbInformation,TitleMsg)
			Exit All
		ElseIf fType = 1 Then
			If MsgBox(Msg,vbYesNo+vbInformation,TitleMsg) = vbNo Then Exit All 'Err.Raise(1,"ExitSub")
		Else
			MsgBox(Msg,vbOkOnly+vbInformation,TitleMsg)
		End If
	End If
End Sub


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


'fType = 0  检查字串是否包含指定字符，并找出指定字符的位置
'fType <> 0  检查字串是否只包含指定字符
Function CheckStr(textStr As String,AscRange As String,fType As Long) As Boolean
	Dim i As Long,j As Long,n As Long,k As Long,InpAsc As Long,Length As Long,Pos As Long
	Dim MinV As Long,MaxV As Long,Temp As String,Stemp As Boolean,FindStemp As Boolean
	CheckStr = False
	If Trim(textStr) = "" Or AscRange = "" Then Exit Function
	k = 0
	Stemp = False
	Length = Len(textStr)
	AscValue = Split(AscRange,",",-1)
	n = UBound(AscValue)
	For i = 1 To Length
		InpAsc = AscW(Mid(textStr,i,1))
		FindStemp = False
		For j = 0 To n
			Temp = AscValue(j)
			Pos = InStr(Temp,"-")
			If Pos > 0 Then
				Min = Left(Temp,Pos-1)
				Max = Mid(Temp,Pos+1)
			Else
				Min = Temp
				Max = Temp
			End If
			If Min <> "" And Max <> "" Then
				MinV = CLng(Min)
				MaxV = CLng(Max)
				If InpAsc >= MinV And InpAsc <= MaxV Then
					If fType = 0 Then k = i
					FindStemp = True
				End If
			ElseIf Min <> "" And Max = "" Then
				MinV = CLng(Min)
				If InpAsc >= MinV Then
					If fType = 0 Then k = i
					FindStemp = True
				End If
			ElseIf Min = "" And Max <> "" Then
				MaxV = CLng(Max)
				If InpAsc <= MaxV Then
					If fType = 0 Then k = i
					FindStemp = True
				End If
			End If
			If FindStemp = True Then Exit For
		Next j
		If fType = 0 And k > 0 Then Exit For
		If fType <> 0 And FindStemp = False Then
			Stemp = True
			Exit For
		End If
	Next i
	If fType <> 0 And Stemp = False Then CheckStr = True
	If fType = 0 And k > 0 Then
		CheckStr = True
		fType = k
	End If
End Function


'转换字符为整数数值
Function StrToLong(mStr As String) As Long
	If mStr = "" Then mStr = "0"
	StrToLong = CLng(mStr)
End Function


'创建代码页数组
Public Function CodePageList(MinNum As Long,MaxNum As Long) As Variant
	Dim i As Long,CodePage() As String,MsgList() As String

	If getMsgList(UILangList,MsgList,"CodePageList",0) = False Then Exit Function
	CodePageCode = "ANSI|_autodetect_all|gb2312|hz-gb-2312|gb18030|big5|euc-jp|iso-2022-jp|shift_jis|" & _
					"_autodetect|ks_c_5601-1987|euc-kr|iso-2022-kr|_autodetect_kr|windows-874|" & _
					"windows-1258|iso-8859-4|windows-1257|ASMO-708|DOS-720|iso-8859-6|windows-1256|" & _
					"DOS-862|iso-8859-8-i|iso-8859-8|windows-1255|iso-8859-9|iso-8859-7|windows-1253|" & _
					"iso-8859-1|cp866|iso-8859-5|koi8-r|koi8-ru|windows-1251|ibm852|iso-8859-2|" & _
					"windows-1250|iso-8859-3|utf-7|utf-8EFBB|utf-8|unicodeFFFE|unicodeFEFF|utf-16LE|" & _
					"utf-16BE|unicode-32FFFE|unicode-32FEFF|utf-32LE|utf-32BE"
	CodePageCodeList = Split(CodePageCode,"|")

	i = UBound(MsgList)
	If MaxNum > i Then MaxNum = i
	ReDim CodePage(MaxNum - MinNum)
	For i = MinNum To MaxNum
		CodePage(i - MinNum) = MsgList(i) & JoinStr & CodePageCodeList(i)
	Next i
	CodePageList = CodePage
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


'关于和帮助
Sub Help(HelpTip As String)
	Dim i As Long,MsgList() As String,TempList() As String

	For i = 0 To UBound(UILangList)
		TempList = Split(UILangList(i),JoinStr)
		Header$ = TempList(0)
		If Header$ = "Windows" Then
			MsgList = Split(TempList(2),SubJoinStr)
			AboutTitle = MsgList(0)
			HelpTitle = MsgList(1)
			AboutWindows = MsgList(2)
			MainWindows = MsgList(3)
			SetWindows = MsgList(4)
			TestWindows = MsgList(5)
			Lines = MsgList(6)
			MainTipTitle = MsgList(7)
			UpdateTipTitle = MsgList(8)
			UILangTipTitle = MsgList(9)
		End If
		If Header$ = "System" Then
			Sys = Replace(Replace(TempList(2),"%s",Version),SubJoinStr,vbCrLf) & vbCrLf & vbCrLf
		End If
		If Header$ = "Description" Then Dec = Replace(TempList(2),SubJoinStr,vbCrLf) & vbCrLf & vbCrLf
		If Header$ = "Precondition" Then Ement = Replace(TempList(2),SubJoinStr,vbCrLf) & vbCrLf & vbCrLf
		If Header$ = "Setup" Then Setup = Replace(TempList(2),SubJoinStr,vbCrLf) & vbCrLf & vbCrLf
		If Header$ = "CopyRight" Then CopyRight = Replace(TempList(2),SubJoinStr,vbCrLf) & vbCrLf & vbCrLf
		If Header$ = "Thank" Then Thank = Replace(TempList(2),SubJoinStr,vbCrLf) & vbCrLf & vbCrLf
		If Header$ = "Contact" Then Contact = Replace(TempList(2),SubJoinStr,vbCrLf) & vbCrLf & vbCrLf
		If Header$ = "Logs" Then Logs = Replace(TempList(2),SubJoinStr,vbCrLf)
		If Header$ = "MainHelp" Then MainHelp = Replace(TempList(2),SubJoinStr,vbCrLf) & vbCrLf & vbCrLf
		If Header$ = "MainSetHelp" Then MainSetHelp = Replace(TempList(2),SubJoinStr,vbCrLf) & vbCrLf & vbCrLf
		If Header$ = "MainTestHelp" Then MainTestHelp = Replace(TempList(2),SubJoinStr,vbCrLf) & vbCrLf & vbCrLf
		If Header$ = "UpdateSetHelp" Then UpdateSetHelp = Replace(TempList(2),SubJoinStr,vbCrLf) & vbCrLf & vbCrLf
		If Header$ = "UILangSetHelp" Then UILangSetHelp = Replace(TempList(2),SubJoinStr,vbCrLf) & vbCrLf & vbCrLf
	Next i

	preLines = Lines & Lines & Lines
	appLines = Lines & Lines & Lines & vbCrLf & vbCrLf
	If HelpTip = "About" Then
		Title = AboutTitle & " " & MainTipTitle
		HelpTipTitle = AboutWindows
		HelpMsg = Sys & Dec & Ement & Setup & CopyRight & Thank & Contact & Logs
	ElseIf HelpTip = "MainHelp" Then
		Title = HelpTitle & " - " & MainTipTitle
		HelpTipTitle = MainWindows
		HelpMsg = MainHelp & Logs
	ElseIf HelpTip = "MainSetHelp" Then
		Title = HelpTitle & " - " & MainTipTitle
		HelpTipTitle = SetWindows
		HelpMsg = MainSetHelp & Logs
	ElseIf HelpTip = "MainTestHelp" Then
		Title = HelpTitle & " - " & MainTipTitle
		HelpTipTitle = TestWindows
		HelpMsg = MainTestHelp & Logs
	ElseIf HelpTip = "UpdateSetHelp" Then
		Title = HelpTitle & " - " & UpdateTipTitle
		HelpTipTitle = SetWindows
		HelpMsg = UpdateSetHelp & Logs
	ElseIf HelpTip = "UILangSetHelp" Then
		Title = HelpTitle & " - " & UILangTipTitle
		HelpTipTitle = SetWindows
		HelpMsg = UILangSetHelp & Logs
	End If

	Begin Dialog UserDialog 760,518,Title ' %GRID:10,7,1,1
		Text 0,7,760,14,HelpTipTitle,.Text,2
		TextBox 0,28,760,455,.TextBox,1
		OKButton 330,490,100,21
	End Dialog
	Dim dlg As UserDialog
	dlg.TextBox = HelpMsg
	Dialog dlg
End Sub
