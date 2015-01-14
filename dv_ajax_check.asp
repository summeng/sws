<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
Response.CharSet="gb2312"
Response.ContentType="text/html"
Dim XMLDom,Dvbbs
Dim CheckType,CheckValue
CheckType = Request("rs")
If Request("rsargs[]")<>"" Then
	CheckValue = Request("rsargs[]")
	CheckValue = Split(CheckValue,",")
Else
	CheckValue = Array()
End If

If CheckType<>"" Then
	Select Case LCase(CheckType)
	Case "checkusername" : CheckUserName()
	Case "checke_mail" : CheckUserEmail()
	'o start 08.1.18
	'Case "checke_regcode" : CheckRegCode()
	Case "checke_dvcode" : CheckDvCode()
	'o end
	End Select
End If
'Dvbbs.PageEnd()
Set Dvbbs = Nothing


Function ErrCode(Str)
	ErrCode = "<img src=""images/note_error.gif"" border=""0""/>&nbsp;<font class=""redfont"">"&Str&"</font>"
End Function

Function SucCode(Str)
	SucCode = "<img src=""images/note_ok.gif"" border=""0""/>&nbsp;<font class=""bluefont"">"&Str&"</font>"
End Function



Sub CheckDvCode()
	Dim CodeStr
	CodeStr = LCase(Trim(CheckValue(Ubound(CheckValue))))
	If CStr(Session("GetCode"))=CStr(CodeStr) And CodeStr<>""  Then
		Response.Write SucCode("")
	Else
		Response.Write ErrCode("")
	End If
End Sub
%>