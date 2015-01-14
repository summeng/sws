<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/tt_auto_lock.asp" -->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
dim newsid
newsid=clng(request.querystring("newsid"))
title=request.querystring("title")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=title%></title>
<meta name="keywords" content="<%=title%>">
<meta name="description" content="<%=title%>">
<link href="/<%=strInstallDir%>css/style.css" rel="stylesheet" type="text/css">
</head>
<body>

<table width="980" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="40" align="center">相关新闻标题：&nbsp;<%=title%></td>
  </tr>
</table>
<table width="980" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF">
            <tr>
              <td width="7%" align="center" bgcolor="#F3F3F3"><img src="../img/review.gif" align="absmiddle" /></td>
              <td width="14%" align="left" bgcolor="#F3F3F3">评论人</td>
              <td width="60%" align="left" bgcolor="#F3F3F3">评论内容</td>
              <td width="19%" align="left" bgcolor="#F3F3F3">评论时间</td>
            </tr>
            <%Call OpenConn
		dim rs1
		set rs1=server.CreateObject("adodb.recordset") 
		rs1.open "select * from DNJY_news_pinglun where sh=1 and id="&newsid&"  order by pinglundate "&DN_OrderType&"",conn,1,1 
		if rs1.eof and rs1.bof then 
		response.write "<tr><td align=center colspan=4>暂时没有评论或未以审核</td></tr>"
		response.End
		else
		do while not rs1.eof  
		%>
            <tr>
              <td height="20" align="center" bgcolor="#F3F3F3"><img src="../img/b.gif" width="10" height="10" /></td>
              <td align="left" bgcolor="#F3F3F3"><b><%=rs1("pinglunname")%></b></td>
              <td align="left" bgcolor="#F3F3F3"><%=rs1("pingluncontent")%></td>
              <td align="left" bgcolor="#F3F3F3"><%=rs1("time")%></td>
            </tr>
            <%rs1.movenext
		loop
		rs1.close
		set rs1=nothing
		end If
		Call CloseDB(conn)%>
            <tr>
              <td height="100%" colspan="4" align="center" bgcolor="#F3F3F3"><font color="#FF0000">您的评论将被网络上成千上万的读者所共享，请对自己言论负责。</font> </td>
            </tr>
</table>
</body>
</html>
