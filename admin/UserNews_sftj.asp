<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file=../Include/Function.asp-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("04")%>
<%Dim isse,ClassName,page,id,sf,selecttype,searchkeys,se_ClassName,caozuocs
isse=ReplaceUnsafe(request.QueryString("isse"))
ClassName=ReplaceUnsafe(Request.QueryString("ClassName"))
page=ReplaceUnsafe(Request.QueryString("page"))
id =ReplaceUnsafe(Request.QueryString("id"))
sf=ReplaceUnsafe(request.QueryString("sf"))
selecttype=ReplaceUnsafe(request.QueryString("selecttype"))
searchkeys=ReplaceUnsafe(request.QueryString("searchkeys"))
se_ClassName=ReplaceUnsafe(request.QueryString("se_ClassName"))
Set rs=Server.CreateObject("ADODB.RecordSet")
sql="select tj from DNJY_userNews where id = "&id
rs.Open sql,conn,1,3
rs("tj")=sf
rs.update
Call CloseRs(rs)

if isse=0 then
	caozuocs="isse=0&page="&page&"&ClassName="&ClassName
	end if
	if isse=1 then
		 if selecttype=1 then
			caozuocs="isse=1&page="&page&"&selecttype="&selecttype&"&searchkeys="&searchkeys&"&se_ClassName="&se_ClassName
		 end if
		 if selecttype=2 then
			caozuocs="isse=1&page="&page&"&selecttype="&selecttype&"&searchkeys="&searchkeys
		 end if
	end if
response.Redirect "UserNews_Modifylist.asp?"&caozuocs
%>

