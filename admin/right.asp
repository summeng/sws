<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/Version.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<html>
<head>
<title>����Աҳ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
body {
	font-size:9pt;
	background-color: #E3EBF9;
}
td  { font-size: 9pt  }
.style2 {
	color: #FFFFFF;
	font-weight: bold;
}
.style3 {color: #FF0000}
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
}
a:hover {
	text-decoration: none;
}
a:active {
	text-decoration: none;
}
.table{BORDER: #3F8805 1px solid;background-color:#3F8805;margin-left:12px}
tr{BACKGROUND-COLOR: #EEFEE0}
.backs{BACKGROUND-COLOR: #3F8805;COLOR: #ffffff;text-align:center}
-->
</style>

<base target="_self">
</head>
<BODY>
<body text="#000000">
<table width="90%" >

<tr><td width="10%"><b>�ٷ����棺</b></a></td><td><iframe scrolling="no" frameborder="0" marginheight="0" marginwidth="0" width="100%" height="15" src="ttscunion.htm"></iframe></td></tr>

</table>
<table width="100%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#ADAED6">
<TR>
    <TD height="20" bgcolor="#799AE1" align="center" colspan="13"><FONT color="#FFFFFF" style="font-size:14px">�� �� �� ҳ</FONT></TD>
 </TR>
  <tr>
    <td height="28" bgcolor="#BDBEDE"><span style="font-size: 9pt"><font size="3"><font size="2">&nbsp;</font></font></span><%=title%><b>����˵��</b></td>
  </tr>
  <tr>
    <td><table width="100%"  border="0" cellpadding="6" cellspacing="1" bgcolor="#eeeeee">
      <tr>
        <td  width="60"><span class="style2"><span style="font-size: 9pt"><font size="3"></font></span></span>��&nbsp;&nbsp;&nbsp;&nbsp;����</td>
        <td>�û�����������Ƽ������ҹ��򱾳����ӵ�С������û����Э�顷�涨�µ�ʹ��Ȩ���û�����ʹ�ñ������������վ�����й���Ȩ�ޡ� ��һ��ʹ��ʱ�뵽[<A href="add_administrator.asp">�����ʺ�����</A>]-[����һ���µĹ���Ա��Ҫ�������Ȩ�ޣ�]֮���˳����¹���Ա��¼ɾ��ϵͳĬ�Ͻ�����[admin]�����ʺţ�ע�⣬��Ҫʹ�ü򵥵����ֻ򵥴���Ϊ���룬�����ñ��˲��ײµõ������ּ��ַ����Ҳ�Ҫ���ض���˼���Է���µ������¼�ƻ��� </td>
      </tr>
      <tr>
        <td width="60">ʹ�����ã�</td>
        <td >�벻Ҫ���������˽��ϵͳ�����޹ص��˹���ϵͳ�������ƻ����ݿ⣬ʹ���ݿ����ݶ�ʧ��</td>
      </tr>
      <tr>
        <td width="60">ע�����</td>
        <td>����ϵͳ�ǲ���COOKIE���й���Ա��֤�ģ�����ÿ�ε�½�������������������Ͻǵ�[<span class="style3"><a href="exit.asp">��ȫ�˳�</a></span>]������ڿͻ���Ĳ������˴����COOKIE�л�ù������룡</td>
      </tr>
    </table></td>
  </tr>
</table>
<center>
<br>
  <table width="100%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#ADAED6">
    <tr>
      <td height="28" bgcolor="#BDBEDE"><span style="font-size: 9pt"><font size="3"><font size="2">&nbsp;</font></font></span><b>��Ȩ����</b></td>
    </tr>
    <tr>
      <td><table width="100%"  border="0" cellpadding="6" cellspacing="1" bgcolor="#eeeeee">

          <tr>
            <td width="60">��Ȩ������</td>
            <td>
              <p>1���������������Ƽ������ģ���Ȩ��<a href="http://www.ip126.com/" target="_blank"><FONT color=#0000ff>��������Ƽ�������</FONT></a>���У�<BR>
              2����������л����񹲺͹�������Ȩ����������������������������ط��ɡ����汣�������߱���һ��Ȩ���� <BR>
              3����ʽ�汾��׼�Ƿ����������ۡ�ת�ã�����һ�к���Ը��������û�ֻ��һ�׳�����һ�������°�װ�����ö�������װ���Ƿ�ת�á����ۡ�����ͽ��������о�������ֹͣ��ط���׷���������Ρ�<br>
			  4���û��������������ܱ�����ġ����ʹ��Э�顷����Э���ĵ��ɵ�<a href="http://www.ip126.com/" target="_blank"><FONT color=#0000ff>�ͷ�����</FONT></a>������������Ŀ�����Ķ�����������뿴������˵���顣</p></td>
          </tr>
      </table></td>
    </tr>
  </table>
  <center>
<br>
    <table width="100%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#ADAED6">
    <tr>
      <td height="28" bgcolor="#BDBEDE"><span style="font-size: 9pt"><font size="3"><font size="2">&nbsp;</font></font></span><b>��������</b></td>
    </tr>
    <tr>
      <td><table width="100%"  border="0" cellpadding="6" cellspacing="1" bgcolor="#eeeeee">

          <tr>
            <td width="60">����������</td>
            <td>
              <p>1��ͬ��������������һ�������ܵ����ְ�ȫ��������ţ��û�ʹ�á���������漰���������������أ����ܻ��ܵ��������ڲ��ȶ����ص�Ӱ�죬�����򲻿ɿ�����������������ڿ͹�����ϵͳ���ȶ����û�����λ���Լ������κ����硢������ͨ����·��ԭ����ɵ���������жϡ����ݶ�ʧ����Ϣй¶�������ʧ���������û�Ҫ��ķ��գ��û������ײ����ге����Ϸ��ա�<br>
2���û�Ӧ�����ñ����������վ֮��ҳ���������ݵ���ʵ�ԡ�׼ȷ�ԡ��Ϸ��Ը�ȫ�����Σ�һ���ɴ�������ľ��ס����鼰���漰�ķ������ξ����û��е���<br>
3���û��������ع��Һ͵����йط��ɡ����桢�������£��������ñ�������������ơ������������κη��ɷ����ֹ���к���Ϣ���û����侭Ӫ��Ϊ�ͷ�������ϢΥ���йع涨��������κε��������Ρ���������Ҫ�е�ȫ�����Σ�����������޹ء�</p></td>
          </tr>
      </table></td>
    </tr>
  </table>
  <br>
   <table width="100%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#ADAED6">
    <tr>
      <td height="28" bgcolor="#BDBEDE"><span style="font-size: 9pt"><font size="3"><font size="2">&nbsp;</font></font></span><b>�ۺ����</b></td>
    </tr>
    <tr>
      <td><table width="100%"  border="0" cellpadding="6" cellspacing="1" bgcolor="#eeeeee">
	  	   <tr>
            <td width="60"><span class="style2"><span style="font-size: 9pt"><font size="3"></font></span></span>��ǰ�汾��</td>
            <td><%Response.Write "<A HREF='http://www.ip126.com' TARGET='_blank'>TianTian &reg; INFO of SAD&#8482; " & SystemVersion & " "&DatabaseType &" Build " & SystemBuildDate & "</a>"%></td>
          </tr>
          <tr>
            <td width="60"><span class="style2"><span style="font-size: 9pt"><font size="3"></font></span></span>����˵����</td>
            <td>�����û����ݲ�ͬ�ļ�λ��ò�ͬ�ķ�����<a href="http://www.ip126.com/" target="_blank">�ٷ��ͷ�������վ</a>Ϊ׼</td>
          </tr>
         <tr>
            <td width="60">��ϵ��ʽ��</td>
            <td><p>
			�ٷ��ͷ�������վ��<a href="http://www.ip126.com/" target="_blank">http://www.ip126.com/</a><BR>
              �����ʼ���bdunion@126.com<BR>
              ��ѶQQ��530051328���Ӻ���ע������Դ�룩<BR>
            </p></td>
          </tr>
      </table></td>
    </tr>
  </table>

    <br>
  <table width="100%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#ADAED6">
    <tr>
      <td height="28" bgcolor="#BDBEDE"><span style="font-size: 9pt"><font size="3"><font size="2">&nbsp;</font></font></span><b>����������</b> ��������������<font color=green><b>��</b></font>��֧�����ǵ�ϵͳ��</td>
    </tr>
</table>
<%
Function IsObjInstalled(strClassString)'���庯��
 On Error Resume Next'�д�������
 IsObjInstalled = False'��������ʼֵ��ΪFALSE
 Err = 0'�������������
 Dim xTestObj'�������
 Set xTestObj = Server.CreateObject(strClassString)'��ʼ������
 If 0 = Err Then IsObjInstalled = True'������������Ϊ0 ������Ϊ�� ��˼���Ѿ���װ��
 Set xTestObj = Nothing'���ٶ���
 Err = 0'�������������
End Function'��������

Function GetVer(ClassStr)
 On Error Resume Next
 GetVer = ""
 Err = 0
 Dim xTestObj
 Set xTestObj = Server.CreateObject(ClassStr)
 If 0 = Err Then GetVer = xTestObj.Version
 Set xTestObj = Nothing
 Err = 0
End Function

Sub WRITE_LINE(str)
 Response.Write LTrim(str)'���str ȥ����߿ո��
End Sub

Dim InstallObj(4)'��������� 0-3 ��˼�Լ��ٶ�
InstallObj(0) = "adodb.connection"
InstallObj(1) = "Scripting.FileSystemObject"
InstallObj(2) = "Persits.Jpeg"
InstallObj(3) = "JMail.Message"

Dim name(4)'����������
name(0) = "Access"
name(1) = "FSO"
name(2) = "ASPJpeg"
name(3) = "JMail"

Dim Effect(4)'�������������
Effect(0) = "ʹ��Access��ȡ����"
Effect(1) = "FSO �ļ�ϵͳ�����ı��ļ���д���ļ����ɻ�ɾ��"
Effect(2) = "ϵͳͨ��ASPJpeg�������ͼƬ���⡢ˮӡ����ֹαװͼƬľ���ϴ���"
Effect(3) = "ϵͳͨ��JMail������͵����ʼ�"


Dim Suggestion(4)'����
Suggestion(0) = "��ϵ�ռ��̽��Access���ݿ�����"
Suggestion(1) = "��֧��FSO�Ŀռ�"
Suggestion(2) = "������Ҫ��װASPJpeg���������֧��ASPJpeg�Ŀռ䣬�����޷��ϴ�ͼƬ��"
Suggestion(3) = "������Ҫ��װJMail���������֧��JMail�Ŀռ䣬�����޷������ʼ�"
%> 
    <tr>
      <td><table border=0 width="100%" cellspacing=1 cellpadding=3>
  <tr class="backs"><td width="200">�������</td> <td width="100">���</td> <td width="400" >�� ��</td><td width="80">֧��/�汾</td><td width="100">����</td></tr>
  <tr><td width="200">Active Server Page</td> <td width="100">ASP</td> <td width="400" >�����ݿ������������н���</td><td width="80"><%response.write "<font color=green><b>��</b></font>"%></td><td width="100">��֧��</td></tr> 

<%Dim i,FoundFso
i=0
for i=0 to 3%>
  <tr><td width="200"><%=InstallObj(i)%></td>
  <td width="100" ><%=name(i)%></td>
  <td width="400" ><%=Effect(i)%></td>
   <td width="80" ><%
FoundFso = IsObjInstalled(InstallObj(i))
 
If FoundFso Then
 Response.Write "<font color=green><b>��</b></font>"&GetVer(InstallObj(i))&""
Else
 Response.Write "<font color=red><b>��</b></font>"
End If
%></td>
<td width="100"><%If FoundFso Then
Response.Write "��֧��"
Else
 Response.Write Suggestion(i)
End If
%></a>
  </tr>
 <%
 next%>
	  <tr>
		<td>IIS�汾</td><td colspan="4"><%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
	  </tr>
   <tr><td width="200" colspan="5"><a href="admin_server.asp" TARGET=right>������Ŀ���>>></a></td>
  </tr>
 </table></td>
    </tr>
  </table>

</center>
</body>
</html>