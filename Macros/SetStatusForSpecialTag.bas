'#Reference {3F4DACA7-160D-11D2-A8E9-00104B365C9F}#5.5#0#C:\Windows\system32\vbscript.dll\3#Microsoft VBScript Regular Expressions 5.5
'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\system32\scrrun.dll#Microsoft Scripting Runtime
Option Explicit
Sub Main
	Dim prj As PslProject
  	Dim trnlst As PslTransList
  	Dim strStatus As String
  	Dim strCount As Integer
  	Dim i As Integer
  	Dim trnstr As PslTransString
  	Dim trnstrResType As String

  	Set prj=PSL.ActiveProject
  	For Each trnlst In prj.TransLists

      'Only process the selected files
      If trnlst.Selected =True  Then
      		strCount = trnlst.StringCount
			For i = 1 To strCount
				Set trnstr = trnlst.String(i)
				'If trnstr.State(pslStateReadOnly)=False _
				 '  	And
				 If  	trnstr.State(pslStateTranslated)=False And  trnstr.State(pslStateHidden)=False Then
				   	'如果字符串的类型是Jscript，直接置状态为"Review"
					trnstrResType=trnstr.ResType
					If (trnstrResType="JavaScript") Then
							trnstr.State(pslStateReview)=True
							trnstr.Property ("HiSoftSpecialTagStatusSetting")="Set js script to Review status"
							strStatus=GetCorrectStatusForJSString(trnstr.Text)
						   	If strStatus="Review" Then
								trnstr.State(pslStateReview)=True
								trnstr.Property ("HiSoftSpecialTagStatusSetting")="Set JS status to review"
							Else
								If(strStatus="Validated") Then
									trnstr.State(pslStateTranslated)=True
									trnstr.Property ("HiSoftSpecialTagStatusSetting")="Set JS status to Validated status"
								End If
							End If


					Else
					   	strStatus=GetCorrectStatusForString(trnstr.Text)
					   	If strStatus="Review" Then
							trnstr.State(pslStateReview)=True
							'trnstr.State(pslStateTranslated)=True
							trnstr.Property ("HiSoftSpecialTagStatusSetting")="Set Special tag(include 'width' and so on) to translated status"
						Else
							If(strStatus="Validated") Then
								trnstr.State(pslStateTranslated)=True
								trnstr.Property ("HiSoftSpecialTagStatusSetting")="Set Special tag to Validated status"
							End If
						End If
					End If

                 End If

			Next i
            trnlst.Save
        End If
	Next

End Sub
Function GetCorrectStatusForJSString(trnstr)
	trnstr=regReplaceVariant(trnstr,"##")
	trnstr=regReplaceVariant(trnstr,"_")
	trnstr=regReplaceNumber(trnstr)
	trnstr=Trim(trnstr)
	Dim tempStr As String
	tempStr=trnstr

	'&lt;hisofttd&gt; height="20"&lt;/hisofttd&gt;
	If trnstr="" Then
		GetCorrectStatusForJSString="Validated"
	Else
		If (Asc(tempStr)=9) Then
			tempStr=""
		End If
		If tempStr="" Or tempStr="/" Then
			GetCorrectStatusForJSString="Review"
		End If

	End If
End Function

Function GetCorrectStatusForString(trnstr)
	trnstr=Replace(trnstr,"&lt;hisofttr&gt;","")
	trnstr=Replace(trnstr,"&lt;/hisofttr&gt;","")

	trnstr=Replace(trnstr,"&lt;hisofttd&gt;","")

	trnstr=Replace(trnstr,"&lt;/hisofttd&gt;","")
	trnstr=Replace(trnstr,"&nbsp;","")
	trnstr=Replace(trnstr,"hisoftreturn","")
	trnstr=Replace(trnstr,"hisofttab","")
	trnstr=Replace(trnstr,"&lt;hibreak/&gt;","")
	trnstr=Replace(trnstr,"\t","")
	trnstr=Replace(trnstr,"\n","")
	trnstr=Replace(trnstr,"<","")
	trnstr=Replace(trnstr,">","")


	trnstr=regReplaceVariant(trnstr,"##")
	trnstr=regReplaceVariant(trnstr,"_")
	trnstr=regReplaceNumber(trnstr)

	trnstr=Trim(trnstr)
	Dim tempStr As String
	tempStr=trnstr

	'&lt;hisofttd&gt; height="20"&lt;/hisofttd&gt;
	If trnstr="" Then
		GetCorrectStatusForString="Validated"
	Else
		'如果是直包含类似height="20" 信息，设置状态为Review

		tempStr=regReplace(tempStr,"height")
		tempStr=regReplace(tempStr,"class")
		tempStr=regReplace(tempStr,"width")
		tempStr=regReplace(tempStr,"valign")
		tempStr=regReplace(tempStr,"align")
		tempStr=regReplace(tempStr,"bgcolor")
		tempStr=regReplace(tempStr,"colspan")
		tempStr=regReplace(tempStr,"rowspan")
		tempStr=regReplace(tempStr,"cellpadding")
		tempStr=regReplace(tempStr,"cellspacing")
		tempStr=regReplace(tempStr,"padding-right")
		tempStr=regReplace(tempStr,"padding-left")
		tempStr=regReplace(tempStr,"border-bottom-color")
		tempStr=regReplace(tempStr,"border-bottom-style")
		tempStr=regReplace(tempStr,"border-right-style")
		tempStr=regReplace(tempStr,"border-right-width")
		tempStr=regReplace(tempStr,"border-right-color")
		tempStr=regReplace(tempStr,"frameborder")
		tempStr=regReplace(tempStr,"scrolling")
		tempStr=regReplace(tempStr,"target")
		'remove src="../images/bar-arrow-gray.png" type tag
		tempStr=regReplace(tempStr,"src")
		tempStr=regReplace(tempStr,"background")
		tempStr=regReplace(tempStr,"onClick")
		tempStr=regReplace(tempStr,"href")

		tempStr=regReplace(tempStr,".style")
		tempStr=regReplace(tempStr,"style")
		tempStr=LCase(Trim(tempStr))
		'Tab key
		If (Asc(tempStr)=9) Then
			tempStr=""
		End If
		tempStr=Replace(tempStr,"/","")
		tempStr=Replace(tempStr,"iframe","")
		tempStr=Replace(tempStr,"img","")
		If tempStr="" Or  tempStr="*" Then
			GetCorrectStatusForString="Validated"
		ElseIf  tempStr="nowrap" Or tempStr="input" Then
			GetCorrectStatusForString="Review"
		Else
			tempStr=regReplace(tempStr,"name")
			tempStr=regReplace(tempStr,"id")
			tempStr=regReplace(tempStr,"value")
			tempStr=regReplace(tempStr,"type")
			tempStr=regReplace(tempStr,"input")
			tempStr=regReplace(tempStr,"textarea")
			tempStr=regReplace(tempStr,"span")
			tempStr=Replace(tempStr,"MULTIPLE","")
			If tempStr="" Then
				GetCorrectStatusForString="Review"
			End If
		End If
	End If

End Function


Function regReplace(StrSource,ReplaceKeyStr)
	Dim regex
	Set regex = New RegExp
	'regex.Pattern    = "<script[^>]*>[\s\D]*?</script>"
	'对于Style格式的正则过程

	If ReplaceKeyStr=".style" Then
		'.style {
		regex.Pattern    = "\..*{(?:.|\n)*}"

	Else
		If ReplaceKeyStr="style" Then
			'style="width: 80%;hisoftreturn hisofttab\t\t\tbackground-color: #c0c0c0;
			regex.Pattern    = "style=""(?:.|\n)*"""
		ElseIf ReplaceKeyStr="background"	Or ReplaceKeyStr="href" Then
			'style="width: 80%;hisoftreturn hisofttab\t\t\tbackground-color: #c0c0c0;
			regex.Pattern    = ReplaceKeyStr+"=""[\._\-/a-zA-Z0-9]+"""
		ElseIf ReplaceKeyStr="src" Then
			regex.Pattern    = "\b"+ReplaceKeyStr+"\s*=\s*""[\._\-/a-zA-Z0-9]+"""
		ElseIf ReplaceKeyStr="onClick" Then
			'style="width: 80%;hisoftreturn hisofttab\t\t\tbackground-color: #c0c0c0;
			regex.Pattern    = "onClick="".*\(.*\);?"""
		Else
			regex.Pattern    = ReplaceKeyStr+"=""?[#_\-%a-zA-Z0-9]*""?"

		End If
	End If
	regex.IgnoreCase = True
	regex.Global     = True
	StrSource = regex.Replace(StrSource, "")
	StrSource=Trim(StrSource)
	Set regex = Nothing
	regReplace=StrSource
End Function
Function regReplaceVariant(StrSource,ReplaceKeyStr)
	Dim regex
	Set regex = New RegExp

	'处理类似:##asfsafasdf##
	'__asdfasdfsadf__等认为是变量东西
	If ReplaceKeyStr="##" Then
		regex.Pattern    = "\w+"+ReplaceKeyStr+"\w+"+ReplaceKeyStr
	Else
		regex.Pattern    = "\b.+=""\w*"+ReplaceKeyStr+"\w*"""
	End If

	regex.IgnoreCase = True
	regex.Global     = True
	StrSource = regex.Replace(StrSource, "")
	StrSource=Trim(StrSource)
	Set regex = Nothing
	regReplaceVariant=StrSource
End Function
Function regReplaceNumber(StrSource)
	Dim regex
	Set regex = New RegExp

	'处理类似:##asfsafasdf##
	'__asdfasdfsadf__等认为是变量东西

	regex.Pattern    = "[0-9]+"


	regex.IgnoreCase = True
	regex.Global     = True
	StrSource = regex.Replace(StrSource, "")
	StrSource=Trim(StrSource)
	Set regex = Nothing
	regReplaceNumber=StrSource
End Function
