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
dim biaoti,id
email=trim(request("email"))
biaoti=trim(request("biaoti"))
%>
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<title>�����ʼ�</title>
<body topmargin="3" leftmargin="0">
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#CCCCCC" width="452" height="64">
    <form action="?id=<%=id%>&dnjymail=ok" method="POST">
    <tr>
      <td width="80" height="25">
      <p align="center"><font color="#FF0000">�ʼ����⣺</font></td>
      <td width="373" height="25">
      <input type="text" name="mailbiaoti" size="55" maxlength="50" value="Re������[<%=title%>]�ٱ�����Ϣ-<%=biaoti%>"></td>
    </tr>
    <tr>
      <td width="80" height="25">
      <p align="center"><font color="#FF0000">�û����䣺</font></td>
      <td width="373" height="25">
      <input type="text" name="maildizhi" size="39" value="<%=email%>" maxlength="40"></td>
    </tr>
    <tr>
      <td width="80" height="25">
      <p align="center"><font color="#FF0000">�ʼ����ݣ�<br>
      �������֣�</font></td>
      <td width="373" height="25">
      <textarea rows="16" name="S1" cols="55">���ã���������[<%=title%>]��http://<%=web%>�������ѷ�������Ϣ����<%=biaoti%>�������˾ٱ���</textarea> <br>�ɼ�����������</td>
    </tr>
    <tr>
      <td width="453" height="35" colspan="2">
      <p align="center">
      <input type="submit" value="�ύ" name="B1" style="color: #FFFFFF; border: 1px solid #000000; background-color: #666666"></td>
    </tr>
    </form>
  </table>
  </center>
</div>

