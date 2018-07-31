'#Reference {EA544A21-C82D-11D1-A3E4-00A0C90AEA82}#6.0#9#C:\Windows\system32\msvbvm60.dll\3#Visual Basic runtime objects and procedures
'#Reference {000204EF-0000-0000-C000-000000000046}#6.0#9#C:\Windows\system32\msvbvm60.dll#Visual Basic For Applications
'#Reference {000204EF-0000-0000-C000-000000000046}#4.0#9#C:\Program Files\Common Files\Microsoft Shared\VBA\VBA6\VBE6.DLL#Visual Basic For Applications
'#Reference {000204EF-0000-0000-C000-000000000046}#2.1#9#C:\Windows\system32\VEN2232.OLB#Visual Basic For Applications
'#Reference {000204F1-0000-0000-C000-000000000046}#1.0#9#C:\Windows\system32\VBAEND32.OLB#Visual Basic For Applications
'#Reference {000204F3-0000-0000-C000-000000000046}#1.0#9#C:\Windows\system32\VBAEN32.OLB#Visual Basic For Applications
'#Reference {000204F3-0000-0000-C000-000000000046}#1.0#804#C:\Windows\system32\VBACHS32.OLB#Visual Basic For Applications
'#Reference {3F4DACA7-160D-11D2-A8E9-00104B365C9F}#5.5#0#C:\Windows\system32\vbscript.dll\3#Microsoft VBScript Regular Expressions 5.5
'#Reference {3F4DACA7-160D-11D2-A8E9-00104B365C9F}#1.0#0#C:\Windows\system32\vbscript.dll\2#Microsoft VBScript Regular Expressions 1.0
'#Reference {3F258A9B-F235-42F8-9DE3-02F6A7ED6862}#1.0#0#C:\Program Files\Windows Live\Photo Gallery\MicrosoftEffects.dll#MicrosoftEffects 1.0 Type Library
'#Reference {00020813-0000-0000-C000-000000000046}#1.6#0#C:\Program Files\Microsoft Office\Office12\EXCEL.EXE#Microsoft Excel 12.0 Object Library
'#Reference {00020813-0000-0000-C000-000000000046}#1.0#804#C:\Program Files\Microsoft Office\Office12\XL5CHS32.OLB#Microsoft Excel 5.0 ╢╘╧≤│╠╨≥┐Γ
'#Reference {00020813-0000-0000-C000-000000000046}#1.0#9#C:\Program Files\Microsoft Office\Office12\XL5EN32.OLB#Microsoft Excel 5.0 ╢╘╧≤│╠╨≥┐Γ
'#Reference {023C8B7F-0664-4CDE-A87F-1E2633A507DD}#1.0#0#C:\Program Files\SDL International\T2007\TT\TradosExcelFileSniffer.dll#ExcelFileSniffer 1.0 Type Library
Sub Main
Dim ex As Excel.Application
Set ex = CreateObject("Excel.Application")
Dim wb As Excel.Workbook
Set wb = ex.Workbooks.Add




Dim col As Integer

col = 1

With wb.Sheets.Add(,,1,xlWorksheet)


For i = 1 To PSL.ActiveProject.TransLists.Count		'翻译列表的个数



	Set src = PSL.ActiveProject.TransLists(i)		'取得某个翻译列表

	If src.Selected Then							'判定翻译列表是否是选中状态



		.Range("a1") = "enString"

		.Range("b1") = "jaString"

		.Range("c1") = "StringID"

		.Range("d1") = "FileName"

		For j = 1 To src.StringCount 				'遍历字串列表





			If src.String(j).State(pslStateReview) And src.String(j).State(pslStateReadOnly)<>1 Then	'对已翻译的和需确认的字串的字数限制进行判定

		col = col + 1

      				.Range("a" & CStr(col) ) = src.String(j).SourceText


      				.Range("b" & CStr(col)) =src.String(j).Text


					.Range("c" & CStr(col)) =src.String(j).ID


					.Range("d" & CStr(col)) =src.TargetFile



		End If
Next j
	End If

Next i

End With
ex.Visible = True

End Sub

