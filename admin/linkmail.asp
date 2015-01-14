<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../Config.asp" -->
<!--#include file=../Include/mail.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%'邮件发送
If trim(request("dnjymail"))="ok" Then
Dim Email,HostName,E_Subject,Information,S_Type,C_M_Chk,e_info
Email=request.form("maildizhi")
HostName=webname
E_Subject=request.form("mailbiaoti")
Information=request.form("S1")
S_Type=True
C_M_Chk=True
e_info=DNJYEmail(Email,HostName,E_Subject,Information,S_Type,C_M_Chk)
response.write e_info
Call CloseDB(conn)
response.end
End If
'邮件发送END%>
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<title>发送邮件</title>
<body topmargin="3" leftmargin="0">
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#CCCCCC" width="352" height="64">
    <form action="?id=<%=id%>&dnjymail=ok" method="POST">
    <tr>
      <td width="80" height="25">
      <p align="center"><font color="#FF0000">邮件标题：</font></td>
      <td width="273" height="25">
      <input type="text" name="mailbiaoti" size="39" maxlength="50" value="<%=title%>与贵站建立友情链接意向"></td>
    </tr>
    <tr>
      <td width="80" height="25">
      <p align="center"><font color="#FF0000">用户信箱：</font></td>
      <td width="273" height="25">
      <input type="text" name="maildizhi" size="39" value="" maxlength="40"></td>
    </tr>
    <tr>
      <td width="80" height="25">
      <p align="center"><font color="#FF0000">邮件内容：<br>
      （纯文字）</font></td>
      <td width="273" height="25">
      <textarea rows="16" name="S1" cols="37">我站－[<%=title%>]有意与贵站建立友情链接，并已做好了链接，如果贵站同意，请做与本站的链接，代码如下： 站名：<%=title%>  网址：http://<%=web%>  图标：http://<%=web%>/images/logo.gif  简介:专业的供求信息发布平台</textarea></td>
    </tr>
    <tr>
      <td width="353" height="35" colspan="2">
      <p align="center">
      <input type="submit" value="提交" name="B1" style="color: #FFFFFF; border: 1px solid #000000; background-color: #666666"></td>
    </tr>
    </form>
  </table>
  </center>
</div>

