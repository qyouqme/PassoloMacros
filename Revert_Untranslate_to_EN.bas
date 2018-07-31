'' Set translated strings validated (for review),only apply to selected lists
Sub SetLock()

Dim time_start As Date
Dim time_end As Date
Dim time_spent

Dim prj As PslProject
'get the active Passolo project
Set prj = PSL.ActiveProject
' exit sub if no active project
If prj Is Nothing Then Exit Sub

PSL.OutputWnd.Clear

time_start=Now()

prj.SuspendSaving

Dim tarlst As PslTransList
Dim n As Double
Dim m As Double

'loop the translation list
'----------for loop--------------
For i = 1 To prj.TransLists.Count
If prj.TransLists(i).Selected = True Then
'loop the string list for each translation list
	For j=1 To prj.TransLists(i).StringCount
	Set tarlst = prj.TransLists(i)
	If tarlst.String(j).State(pslStateLocked) = False Then
	m = m + 1
		If tarlst.String(j).State(pslStateTranslated)= False Then
		tarlst.String(j).Text=tarlst.String(j).SourceText
                n = n + 1
		End If
	End If
	Next j
	'PSL.Output (tarlst.Title)
	tarlst.Save
End If
Next i
'-------------for end---------------

prj.ResumeSaving

time_end=Now()

time_spent=DateDiff("s",time_start,time_end)

MsgBox ("Total strings: " & m & vbTab & "Untranslate to EN: " & n,vbOkCancel,"Info")
PSL.Output ("Total strings: " & m & vbTab)
PSL.Output ("Untranslate to EN: " & n & vbTab)
PSL.Output ("Total spent time: " & time_spent &" s")

End Sub
Sub Main

	SetLock

End Sub
