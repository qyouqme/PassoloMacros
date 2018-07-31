'This macro is used to modify the target path of the selected translist.
'We need to set the new path to the var newPath before running the macro.

Sub Main

	Dim prj As PslProject
	Dim trnlst As PslTransList
	Dim newPath As String
	Dim oriPath As String
	Dim repPath As String

	PSL.OutputWnd.Clear

	'original path which will be replaced with repPath

	oriPath = "\PP 459\en\"

	repPath = "\en\"

	Set prj = PSL.ActiveProject

	If prj Is Nothing Then

	Else

		For Each trnlst In prj.TransLists

			If trnlst.Selected =True  Then

				'newPath = Replace(trnlst.TargetFile, oriPath, repPath)
				newPath = Replace(trnlst.SourceList.SourceFile, oriPath, repPath)

				'trnlst.TargetFile = formatPath(newPath) + getFileName(trnlst.TargetFile)

				trnlst.TargetFile = newPath
				'trnlst.SourceList.SourceFile = newPath

				'PSL.Output getFileName(trnlst.TargetFile)

				'PSL.Output formatPath(newPath)

			End If

		Next

	End If

End Sub


Function getFileName(path)

	sFlag=InStr(1,path,"\")

	While sFlag>0

		path=Right(path,Len(path)-sFlag)

		sFlag=InStr(1,path,"\")

	Wend

	getFileName=path

End Function


Function formatPath(path)

	formatPath = path

	If Right(formatPath,1) <> "\" Then

		formatPath = formatPath + "\"

	End If

End Function
