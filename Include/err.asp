<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%sub errdick(id)%>
<title>����ҳ��</title>
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
<link href="/<%=strInstallDir%>css/css.CSS" rel="stylesheet" type="text/css" />
<div align="center">
  <center>


<TABLE border=0 align="center" cellSpacing=0>
  <TBODY>
    <TR>
      <TD  background="/<%=strInstallDir%>images/ereer.gif" width="750" height="450" align="center">
	  <TABLE height="150">
	  <TR>
		<TD> </TD>
	  </TR>
	  </TABLE>
<TABLE  cellSpacing=0 cellPadding=0 border=0 >
          <TBODY>
            <TR><TD width="250"> &nbsp;</td>
<TD width="300">
<table cellSpacing="0" cellPadding="0" width="500" border="0" style="border-collapse: collapse" bordercolor="#111111">
  <tr>
    <td width="300">
    <table cellSpacing="0" cellPadding="3" width="100%" border="0">
      <tr>
        <td width="100%">��</td>
      </tr>
      <tr>
        <td>
<%'�ͷ�QQ���ļ�����Ҫ���κ��޸�
Dim n,ppp,qq,mymail,qq_name
Call OpenConn
ppp=Conn.Execute("Select WebSetting From [DNJY_config]")(0)
ppp=Split(ppp,"|")
qq=ppp(10)
mymail=ppp(4)

if qq<>"" Then 
qq=replace(qq,"��",",")
if isnull(qq_name) or qq_name="" then qq_name=qq
qq_name=replace(qq_name,"��",",")
qq_name=split(qq_name,",")
qq=split(qq,",")
end If%>
          <span class="STYLE1">
          <%
if id=0 then
response.write "<li>�Ƿ��ύ��"
elseif id=1 then
response.write "<li>�ʼ���ַ����"
elseif id=2 then
response.write "<li>�Ƿ��ַ�,���������룡"
elseif id=3 then
response.write "<li>����ʹ�������û�����"
elseif id=4 then
response.write "<li>�������벻������"
elseif id=5 then
response.write "<li>���������������"
elseif id=6 then
response.write "<li>���֤������д����"
elseif id=7 then
response.write "<li>�û������ڻ��������"
elseif id=8 then
response.write "<li>�������̫�࣬�����Ѿ�ֹͣ�������ˣ�"
elseif id=9 then
response.write "<li>���������Ҳ����û����ϣ�"
elseif id=10 then
response.write "<li>һ������û��ѡ��"
elseif id=11 then
response.write "<li>��������û��ѡ��"
elseif id=12 then
response.write "<li>������Ϣʱ���ⲻ��Ϊ�գ�"
elseif id=13 then
response.write "<li>�����к��зǷ��������ַ�����ע���飡"
elseif id=14 then
response.write "<li>�۸��������"
elseif id=15 then
response.write "<li>���׵������󣬲�����Ҫѡ����������������û����������㽻�׵���Ʒ��"
elseif id=16 then
response.write "<li>������Ϣ��˵��û����д��"
elseif id=17 then
response.write "<li>��ϵ�˱�����д������Ը���Ϊ�������֣�"
elseif id=18 then
response.write "<li>EMAIL������д������Ը���Ϊ�����ʼ���ַ��"
elseif id=19 then
response.write "<li>��������ʼ���ַ��"
elseif id=20 then
response.write "<li>��ϵ��ʽ/�绰����Ϊ�գ�"
elseif id=21 then
response.write "<li>������������ܶ���30���ַ���"
elseif id=22 then
response.write "<li>��������û����Ѿ�ע�����������Ѿ�ע���ˣ����һ����룡"
elseif id=23 then
response.write "<li>��������ʼ���ַ�Ѿ�ע�����������Ѿ�ע���ˣ����һ����룡"
elseif id=24 then
response.write "<li>δ֪��½����"
elseif id=25 then
response.write "<li>��û���㹻�ı����ɫ���ߣ��벻Ҫ�޸ĳ�������в�����"
elseif id=26 then
response.write "<li>��û���㹻����Ϣ�ö����ߣ��벻Ҫ�޸ĳ�������в�����"
elseif id=27 then
response.write "<li>��û���㹻����Ϣ��ͼ���ߣ��벻Ҫ�޸ĳ�������в�����"
elseif id=28 then
response.write "<li>��û���㹻����Ϣ��֤���ߣ��벻Ҫ�޸ĳ�������в�����"
elseif id=29 then
response.write "<li>ϵͳ���������ύ����̫�죬ϵͳ��ֹ���У���ȴ�5���ӣ�"
elseif id=30 then
response.write "<li>���벻һ�»����"
elseif id=31 then
response.write "<li>���������QQ���룡"
elseif id=32 then
response.write "<li>�����������ϸ��ַ��"
elseif id=33 then
response.write "<li>��½�ʺŲ��ܰ������û������ַ���"
elseif id=34 then
response.write "<li>������Ѷ������д��ȷ���䣡"
elseif id=35 Then
 response.write "��ʾ������IP��ַ<font color=red> "&Request.serverVariables("REMOTE_ADDR")&" </font>���ӷǷ��½��Ա���룬�ѱ�ϵͳ���Ʒ��ʣ��������Ա��ϵ������֪��������IP��ַ��<br><br>�ͷ�����,������ţ�<img src=/"&strInstallDir&"images/email.gif border=0 align=center> <a href='mailto:"&mymail&"?subject=�ҵ�IP "&Request.serverVariables("REMOTE_ADDR")&" ��"&webname&"����'>"&mymail&"</a><br>�ͷ�QQ,�����ʱ�Ự��"
for n=0 to UBound(qq)
 response.write "<a class='c' target=blank href='tencent://message/?uin="&qq(n)&"&Site="&webname&"&Menu=yes'>"&qq(n)&"</a> "
Next

 elseif id=36 Then
  response.write "��ʾ������IP��ַ<font color=red> "&Request.serverVariables("REMOTE_ADDR")&" </font>���ӹ�ˮ���ѱ�ϵͳ���Ʒ��ʣ��������Ա��ϵ������֪��������IP��ַ��<br><br>�ͷ�����,������ţ�<img src=/"&strInstallDir&"images/email.gif border=0 align=center> <a href='mailto:"&mymail&"?subject=�ҵ�IP "&Request.serverVariables("REMOTE_ADDR")&" ��"&webname&"����'>"&mymail&"</a><br>�ͷ�QQ,�����ʱ�Ự��"
for n=0 to UBound(qq)
 response.write "<a class='c' target=blank href='tencent://message/?uin="&qq(n)&"&Site="&webname&"&Menu=yes'>"&qq(n)&"</a> "
Next
 elseif id=37 Then
 response.write "��ʾ������IP��ַ<font color=red> "&Request.serverVariables("REMOTE_ADDR")&" </font>����SQLע��Σ����վ���ѱ�ϵͳ���Ʒ��ʣ��������Ա��ϵ������֪��������IP��ַ��<br><br>�ͷ�����,������ţ�<img src=/"&strInstallDir&"images/email.gif border=0 align=center> <a href='mailto:"&mymail&"?subject=�ҵ�IP "&Request.serverVariables("REMOTE_ADDR")&" ��"&webname&"����'>"&mymail&"</a><br>�ͷ�QQ,�����ʱ�Ự��"
for n=0 to UBound(qq)
 response.write "<a class='c' target=blank href='tencent://message/?uin="&qq(n)&"&Site="&webname&"&Menu=yes'>"&qq(n)&"</a> "
Next 
elseif id=38 then
response.write "<li>��ַ���ô�http://��"
elseif id=39 then
response.write "<li>������֤�벻�ԣ�"
elseif id=40 then
response.write "<li>������֤�벻�ԣ�"
elseif id=41 then
response.write "<li>������֤���ԣ�"
elseif id=42 then
response.write "<li>�㷢������Ϣ��������Υ����վ����֮����"
elseif id=43 Then
 response.write "��ʾ������IP��ַ<font color=red> "&Request.serverVariables("REMOTE_ADDR")&" </font>�ѱ�ϵͳ���Ʒ�����Ϣ���������Ա��ϵ������֪��������IP��ַ��<br><br>�ͷ�����,������ţ�<img src=/"&strInstallDir&"images/email.gif border=0 align=center> <a href='mailto:"&mymail&"?subject=�ҵ�IP "&Request.serverVariables("REMOTE_ADDR")&" ��"&webname&"����'>"&mymail&"</a><br>�ͷ�QQ,�����ʱ�Ự��"
for n=0 to UBound(qq)
 response.write "<a class='c' target=blank href='tencent://message/?uin="&qq(n)&"&Site="&webname&"&Menu=yes'>"&qq(n)&"</a> "
Next
elseif id=44 then
response.write "<li>�ʴ���֤���ԣ�"
elseif id=45 then
response.write "<li>��ʽ��֤��𰸲��ԣ�"
elseif id=46 Then
 response.write "��ʾ������IP��ַ<font color=red> "&Request.serverVariables("REMOTE_ADDR")&" </font>�ѱ�ϵͳ���Ʒ��ʣ��������Ա��ϵ������֪��������IP��ַ��<br><br>�ͷ�����,������ţ�<img src=/"&strInstallDir&"images/email.gif border=0 align=center> <a href='mailto:"&mymail&"?subject=�ҵ�IP "&Request.serverVariables("REMOTE_ADDR")&" ��"&webname&"����'>"&mymail&"</a><br>�ͷ�QQ,�����ʱ�Ự��"
for n=0 to UBound(qq)
 response.write "<a class='c' target=blank href='tencent://message/?uin="&qq(n)&"&Site="&webname&"&Menu=yes'>"&qq(n)&"</a> "
Next
elseif id=47 then
response.write "<li>��ϵ�˻��ַ���ܰ������û������ַ���"
elseif id=48 then
response.write "<li>��ϵ��ַ���ܰ��������ַ���"
elseif id=49 then
response.write "<li>��Ϣ��ϸ���ݲ��ܰ������û������ַ���"
elseif id=50 then
response.write "<li>��վ��ǰֻ��״̬������ע���Ա��������Ϣ�����Եȣ���Ҫ����������Ϣ����QQ��ʱ�Ự��ϵ�ͷ���"
for n=0 to UBound(qq)
 response.write "<a class='c' target=blank href='tencent://message/?uin="&qq(n)&"&Site="&webname&"&Menu=yes'>"&qq(n)&"</a> "
Next

end if
%>
          </span></td>
      </tr>
      <tr>
        <td>��</td>
      </tr>
      <tr>
        <td><div align="center"><a href="javascript:history.back(-1);">
          <img height="15" src="/<%=strInstallDir%>images/_back.gif" width="56" border="0">����</a></div></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
  </center>
</div>
<%end sub%>