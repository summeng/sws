<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
dim dsnstr1,dsnstr,rs
dim db
db="../data/#data#.asp"
Set dsnstr = Server.CreateObject("ADODB.Connection")
dsnstr1="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
dsnstr.Open dsnstr1
set rs=server.CreateObject("adodb.recordset")
%>

<%
Function htmlCode(content)
	TempContent = content
	TempContent = replace(TempContent,"<","&lt;")
	TempContent = replace(TempContent,">","&gt;")
	TempContent = replace(TempContent,vbcrlf,"&nbsp;<br>")
	TempContent = replace(TempContent,chr(32),"&nbsp;")
	TempContent = replace(TempContent,chr(34),"&quot;")
	htmlCode=TempContent
End Function

function code(content)
	TempContent = content
	TempContent = replace(TempContent,chr(34),"&quot;")
	code=TempContent
end function
%>