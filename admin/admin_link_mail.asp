<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../Config.asp" -->
<!--#include file=../Include/mail.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>


<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<title>�����ʼ�</title>
<body topmargin="3" leftmargin="0">
<div align="center">
  <center>
<%'�ʼ�����
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
'�ʼ�����END
If request("key")="ok" then
dim mail,name,qq
mail=trim(request("mail"))
name=trim(request("name"))
%>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#CCCCCC" width="352" height="64">
    <form action="?dnjymail=ok" method="POST">
    <tr>
      <td width="80" height="25">
      <p align="center"><font color="#FF0000">�ʼ����⣺</font></td>
      <td width="273" height="25">
      <input type="text" name="mailbiaoti" size="39" maxlength="50" value="<%=webname%>���վ[<%=name%>]������������"></td>
    </tr>
    <tr>
      <td width="80" height="25">
      <p align="center"><font color="#FF0000">�ʼ���ַ��</font></td>
      <td width="273" height="25">
      <input type="text" name="maildizhi" size="39" value="<%=mail%>" maxlength="40"></td>
    </tr>
    <tr>
      <td width="80" height="25">
      <p align="center"><font color="#FF0000">�ʼ����ݣ�<br>
      �������֣�</font></td>
      <td width="273" height="25">
      <textarea rows="16" name="S1" cols="37">�𾴵�<%=name%>����Ա��Ϊ��ǿ�����ƹ㣬��վ��<%=webname%>�������վ[<%=name%>]�����������ӣ������������ӣ���վ��ͬ������������ӣ�лл������վ���ƣ�<%=webname%>����ַ��http://<%=web%>  LOGO��ַ��http://<%=web%>/images/logo.gif  ��ϵQQ��<%=qq%>��----<%=webname%>����Ա</textarea></td>
    </tr>
    <tr>
      <td width="353" height="35" colspan="2">
      <p align="center">
      <input type="submit" value="�ύ" name="B1" style="color: #FFFFFF; border: 1px solid #000000; background-color: #666666"></td>
    </tr>
    </form>
  </table>
  <%End if%>
  </center>
</div>
<%Call CloseDB(conn)%>
