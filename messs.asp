<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=top.asp-->
<!--#include file=usercookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%dim mess%>
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<link rel="stylesheet" type="text/css" href="../1.CSS">
<style type="text/css">
<!--
.style1 {color: #FF0000}
.style5 {color: #FF000; font-weight: bold; }
-->
</style>
</head>

<body background="img/bg1.gif" leftmargin="0" topmargin="0">

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="980" height="426">
    <tr>
      <td width="214" height="354" valign="top" bgcolor="#FFFFFF"><div align="center">
      
 <!--#include file=userleft.asp--></div></td>
 <td width="5">&nbsp;</td>
 <td width="760" height="354" valign="top" bgcolor="#FFFFFF">
 <%Dim rsmess,pages,allPages,page,neirong,riqi,isnew,fname,delid

if request.cookies("dick")("username")="" then
response.Write "<center>请先登录</center>"
response.End
end if
username=request.cookies("dick")("username")

If request("action")="msgdel" Then
   Set rs = conn.Execute("select * from DNJY_Message where id="&request("delid")) 
	if rs.eof and rs.bof then 
	response.write "<script language='javascript'>"
	response.write "alert('出错了，数据库操作失败！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
    elseif ucase(rs("name"))="ALL" then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，您不能删除这条消息，可能这条消息属于公共消息！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
    elseif Ucase(request.cookies("dick")("username"))<>Ucase(rs("name")) then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，您不能删除这条消息，可能这条消息属于公共消息！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
    else
    conn.Execute ("update DNJY_Message set del='1' where id="&request("delid")) 
	response.write "<script language='javascript'>"
	response.write "alert('操作成功，所选消息已被删除！');"
	response.write "location.href='?"&C_ID&"';"
	response.write "</script>"
	response.end
    end if
conn.close
set conn=nothing
End If

set rs=server.createobject("adodb.recordset")
sql = "select mess from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,1
mess=rs("mess")
Call CloseRs(rs)

If mess=0 then
response.write "<table width=100% border=0 align=center><tr><td><b>欢迎消息：</b><p>&nbsp;&nbsp;&nbsp;&nbsp;尊敬的 "&username&" 您好!祝贺您成为["&webname&"]会员，在这里您可发布和查询信息。请记住本站网址"&web&"、您的用户名和密码，以后发布信息更快捷。建议选择本站提供的VIP会员、竞价信息等有偿服务，建立有自己独立界面的商家店铺、企业名片和发布竞价信息等，使您的信息得到更好的推荐和在显著位置显示。详见帮助中心。您的支持是我们发展的动力。欢迎您和您的朋友经常回家看看！</td></tr></table>"
conn.execute "UPDATE DNJY_user SET mess=1 WHERE username='"&request.cookies("dick")("username")&"'" '标记欢迎短信已阅读
End if
set rsmess=server.CreateObject("adodb.recordset")
rsmess.Open "select * from DNJY_Message where del='0' and (name='ALL' or name='"&username&"') order by name asc, riqi "&DN_OrderType&"",conn,1,1

if rsmess.eof and rsmess.bof  then 
If mess<>0 then
	response.write "<table width=100% border=0 align=center class=""ty1""><tr><td><marquee width=100% height=10>暂时没有站内消息</marquee></td></tr></table>"
	Else
response.write ""
End if
	else
response.write "<marquee><font color=red>节约每一分空间，请及时删除无用短信，谢谢合作。重要资料请及时转移保存，确保安全！</font></marquee>"
pages = 10			'每页记录数
rsmess.pageSize = pages
allPages = rsmess.pageCount		'计算一共能分多少页
page = Request("page")
'if语句属于基本的排错处理
if isEmpty(page) or clng(page) < 1 then
page = 1
elseif clng(page) >= allPages then
page = allPages 
end if
rsmess.AbsolutePage = page

do while not rsmess.eof and  pages>0
neirong=rsmess("neirong")
riqi=rsmess("riqi")
isnew=rsmess("isnew")
fname=rsmess("fname")
id=rsmess("id")
%>

<form name=msgdel method=post action=saveprofile.asp?action=msgdel&<%=C_ID%>>
<table width="100%" border="1" cellpadding="4" cellspacing="0" bordercolor="#C0C0C0" align=center bordercolordark="#FFFFFF" style="word-break:break-all" class="ty1">
<tr>
<td onMouseOver="bgColor='#e7e7e7'" onMouseOut="bgColor='#FFFFFF'">
<table border=0 width=100%>
<tr><td width=150 rowspan=2>
<a href='messh.asp?id=<%=id%>&fname=<%=fname%>&<%=C_ID%>'><img alt="回复此消息" src=images/m_replyp.gif border=0></a> &nbsp;  
<span  style="cursor:hand" onclick="{if(confirm('提示：消息删除后不可恢复，\n\n您确实删除这条消息吗？ ')){location.href='?action=msgdel&delid=<%=id%>&<%=C_ID%>';}}"><img ALT="删除此消息" src=images/m_delete.gif border=0></span>&nbsp;&nbsp;
</td><td>
<%if rsmess("name")="ALL" then
response.write "<font color=blue>这是一条公共短信</font>"
end if
%>
</td></tr>
<tr><td>发送人：<%=fname%>  &nbsp;  发送时间：<%=riqi%></td></tr>
<tr><td colspan=2><hr color="#FFFFFF" WIDTH=100%>
<%=HTMLEncode(replace(neirong,vbCRLF,"<BR>"))%>
</td></tr>
</table>
</td></tr>
</table>
</form>
<%
pages = pages - 1
rsmess.movenext
if rsmess.eof then exit do
loop

conn.execute "UPDATE DNJY_Message SET isnew ='1' WHERE name ='"&request.cookies("dick")("username")&"'" '标记短信已阅读
response.write "<table width=100% border=0 align=center><tr><td height=50 valign=top>"
if page = 1 then
response.write "<font color=darkgray>首页 前页</font>"
else
response.write "<a href="&request.ServerVariables("script_name")&"?action=my_msg&page=1&"&C_ID&">首页</a> <a href="&request.ServerVariables("script_name")&"?action=my_msg&page="&page-1&"&"&C_ID&">前页</a>"
end if
if page = allpages then
response.write "<font color=darkgray> 下页 末页</font>"
else
response.write " <a href="&request.ServerVariables("script_name")&"?action=my_msg&page="&page+1&"&"&C_ID&">下页</a> <a href="&request.ServerVariables("script_name")&"?action=my_msg&page="&allpages&"&"&C_ID&">末页</a>"
end if
response.write " 第"&page&"页 共"&allpages&"页"
response.write "</td></tr></table>"

end if
rsmess.close
set rsmess=nothing
%></td>
      <td width="4" valign="top" bgcolor="#e6e6e6"></td>
    </tr>
    <tr>
      <td height="26" colspan="3"><!--#include file=footer.asp--></td>
    </tr>
  </table>
  </center>
</div>
</body>
</html>
