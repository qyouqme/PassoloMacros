'#Reference {831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0#C:\Windows\system32\mscomctl.ocx#Microsoft Windows Common Controls 6.0 (SP6)
'#Reference {00020813-0000-0000-C000-000000000046}#1.6#0#C:\Program Files\Microsoft Office\Office12\EXCEL.EXE#Microsoft Excel 12.0 Object Library
'#Reference {00020813-0000-0000-C000-000000000046}#1.0#804#C:\Program Files\Microsoft Office\Office12\XL5CHS32.OLB#Microsoft Excel 5.0 ╢╘╧≤│╠╨≥┐Γ
'#Reference {3F4DACA7-160D-11D2-A8E9-00104B365C9F}#5.5#0#C:\Windows\system32\vbscript.dll\3#Microsoft VBScript Regular Expressions 5.5
'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\system32\scrrun.dll#Microsoft Scripting Runtime
''Export the update [en->target] stringlist for each sync Loop
Sub Main
	start_time = Timer()

'defind related var stirngs
Dim prj As PslProject
Dim trnlst As PslTransList
Dim  i, j, n As Long
Dim row As Long
Dim trans_str As PslTransString
Dim fso
Dim myfile As String
Dim s As String





'建立Excel对象
Dim ea As Excel.Application
Dim wb As Excel.Workbook
Dim shet



Dim flag As Boolean

Set prj = PSL.ActiveProject

flag = False


Set ea = CreateObject("Excel.Application")
Set wb = ea.Workbooks.Add
'ea.Visible = True





With shet

For n=1 To prj.Languages.Count

Set trnlst = prj.TransLists(n)

Set shet = wb.Sheets.Add(,,1,xlWorksheet)

shet.Name = trnlst.Language.LangCode

		.Range("a1") = "Title"

		.Range("b1") = "Resource"

		.Range("c1") = "Number"

		.Range("d1") = "ID"

		.Range("e1") = "English"

		.Range("f1") = "Localized"

		.Range("g1") = "Comments"


Next n

End With


For i = 1 To prj.TransLists.Count		'翻译列表的个数


Set trnlst = prj.TransLists(i)		'取得某个翻译列表

	 With wb.Sheets(trnlst.Language.LangCode)

		For j = 1 To trnlst.StringCount 				'遍历字串列表

			If trnlst.String(j).State(pslStateReview) = True Then'处理need to review string

                    row = wb.Sheets(trnlst.Language.LangCode).UsedRange.Rows.Count

      				.Range("a" & CStr(row+1) ) = trnlst.Title

					.Range("b" & CStr(row+1)).NumberFormatLocal="@"
      				.Range("b" & CStr(row+1)) = trnlst.String(j).ResType

					.Range("c" & CStr(row+1)).NumberFormatLocal="@"
					.Range("c" & CStr(row+1)) = trnlst.String(j).Number

					.Range("d" & CStr(row+1)).NumberFormatLocal="@"
					.Range("d" & CStr(row+1)) = trnlst.String(j).ID

					.Range("e" & CStr(row+1)).NumberFormatLocal="@"
      				.Range("e" & CStr(row+1)) = trnlst.String(j).SourceText

					.Range("f" & CStr(row+1)).NumberFormatLocal="@"
					.Range("f" & CStr(row+1)) = trnlst.String(j).Text

					.Range("g" & CStr(row+1)).NumberFormatLocal="@"
					.Range("g" & CStr(row+1)) = trnlst.String(j).Comment

		End If
Next j

End With

Next i



myfile = prj.Location & "\" & prj.Name & "_Termlist.xlsx"

Set fso = CreateObject("Scripting.FileSystemObject")
If Dir(myfile) <> "" Then fso.DeleteFile(myfile)

wb.SaveAs(myfile)
wb.Close

MsgBox "Done!"

End Sub
