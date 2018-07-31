'' Set translated strings locked (translated&validated),only apply to selected lists
Sub SetLock()
Dim prj As PslProject
'get the active Passolo project
Set prj = PSL.ActiveProject
' exit sub if no active project
If prj Is Nothing Then Exit Sub

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
		If tarlst.String(j).State(pslStateTranslated) = True Then
			If tarlst.String(j).State(pslStateReview) = False Then
				tarlst.String(j).State(pslStateLocked) = True
			n = n + 1
			End If
		End If
	End If
	Next j
	PSL.Output (tarlst.Title)
	tarlst.Save
End If
Next i
'-------------for end---------------

MsgBox ("Total strings: " & m & vbTab & "Set Lock: " & n,vbOkCancel,"Info")
PSL.Output ("Total strings: " & m & vbTab & "Set Lock: " & n)
End Sub
Sub Main
SetLock
End Sub
