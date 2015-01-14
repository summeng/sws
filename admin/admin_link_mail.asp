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


<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<title>发送邮件</title>
<body topmargin="3" leftmargin="0">
<div align="center">
  <center>
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
'邮件发送END
If request("key")="ok" then
dim mail,name,qq
mail=trim(request("mail"))
name=trim(request("name"))
%>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#CCCCCC" width="352" height="64">
    <form action="?dnjymail=ok" method="POST">
    <tr>
      <td width="80" height="25">
      <p align="center"><font color="#FF0000">邮件标题：</font></td>
      <td width="273" height="25">
      <input type="text" name="mailbiaoti" size="39" maxlength="50" value="<%=webname%>与贵站[<%=name%>]友情链接事宜"></td>
    </tr>
    <tr>
      <td width="80" height="25">
      <p align="center"><font color="#FF0000">邮件地址：</font></td>
      <td width="273" height="25">
      <input type="text" name="maildizhi" size="39" value="<%=mail%>" maxlength="40"></td>
    </tr>
    <tr>
      <td width="80" height="25">
      <p align="center"><font color="#FF0000">邮件内容：<br>
      （纯文字）</font></td>
      <td width="273" height="25">
      <textarea rows="16" name="S1" cols="37">尊敬的<%=name%>管理员：为加强合作推广，本站－<%=webname%>有意与贵站[<%=name%>]建立友情链接，并已做好链接，贵站若同意请做相关链接，谢谢！本网站名称：<%=webname%>，网址：http://<%=web%>  LOGO地址：http://<%=web%>/images/logo.gif  联系QQ：<%=qq%>。----<%=webname%>管理员</textarea></td>
    </tr>
    <tr>
      <td width="353" height="35" colspan="2">
      <p align="center">
      <input type="submit" value="提交" name="B1" style="color: #FFFFFF; border: 1px solid #000000; background-color: #666666"></td>
    </tr>
    </form>
  </table>
  <%End if%>
  </center>
</div>
<%Call CloseDB(conn)%>
