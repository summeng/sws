<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("14")%>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<style type="text/css">
<!--
.style1 {
	color: #FFFFFF;
	font-weight: bold;
}
body {
	background-color: #E3EBF9;
}
-->
</style>
<style type="text/css">
/*��ʾ��Ϣ*/
.info {position:relative;background:#;color:#666; text-decoration:none;font-size:12px;width:0px;height:0px;text-align:center;border:0px solid #ccc;line-height:18px;}/*�������ӵ�����,һ��Ҫ����Ϊrelative����ʹ��ʾ�����������*/
.info:hover {background:#eee;color:#333;}
.info span {display: none }/*���������µ�spanΪ����״̬*/
.info:hover span /*����hover�µ�span����Ϊ����״̬,��������ʾ���λ��*/{display:block;position:absolute;top:0px;left:25px;width:200px;
border:1px solid #DC5826; background:#FFF3B4; color:#333;padding:5px;text-align:left;}
</style>
<%Call OpenConn
dim rows,action,ip,i,url,rsip
rows=request.Form("rows")
if rows="" or rows=0 then rows=10
action=request.Form("ok")
if action="" then 
	set rs=server.createobject("adodb.recordset")
	sql="select * from [DNJY_config]"
	rs.open sql,conn,1,3
%>
<div id="dHTMLToolTip" style="position: absolute; visibility: hidden; width:10; height: 10; z-index: 1000; left: 0; top: 0"></div>

<form action=admin_ip.asp method=post name=addip>
<table width="98%" border="1"  style="border-collapse: collapse; border-style: dotted; border-width: 0px"  bordercolor="#333333" cellspacing="0" cellpadding="2">

<tr><td width=25% align=right>    
����IP�б�<br>(�������ɫѡ����Ч) &nbsp;</td><td><TEXTAREA alt="���������Ʒ��ʱ�վ��IP��ַ<br>�磺205.158.40.15<br>ÿ��IP��ַռһ��" NAME="IP" ROWS="<%=rows%>" COLS="25" style="overflow:auto;">
<%
if rs("IP")<>"" then response.write Replace(rs("IP"),"@",vbCRLF)
%>
</TEXTAREA><br>����ʾ��������IP��ַ��һ���֣��磺220.50����ô�κΰ���220.50��IP��ַ������ֹ���ʱ�վ��</td></TR>
<tr><td colspan=2 width="568" align="center"><INPUT name="ok" TYPE="hidden" value="ip"><INPUT name=action TYPE="submit" value="ȷ���޸�"></form></table>

<table width="98%" border="1"  style="border-collapse: collapse; border-style: dotted; border-width: 0px"  bordercolor="#333333" bgcolor="#cccccc" cellspacing="0" cellpadding="2"><form action=admin_ip.asp method=post name=setup>
              <tr> 
            <td height="22" align="right"><font color=red>�Ƿ�����IP���Ʒ�����վ��</font></td>
            <td height="22" colspan="3"><input type="radio" name="lockip" value="1" <%if rs("lockip")="1" then%>checked<%end if%>>�� <input type="radio" name="lockip" value="0" <%if rs("lockip")<>"1" then%>checked<%end if%>>��&nbsp;&nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>������IP�б��е�IP��ַ���޷����ʱ�վ</span></a></td>
          </tr>
              <tr> 
            <td height="22" align="right"><font color=red>�Ƿ�����IP���Ʒ�����Ϣ��</font></td>
            <td height="22" colspan="3"><input type="radio" name="addip" value="1" <%if rs("addip")="1" then%>checked<%end if%>>�� <input type="radio" name="addip" value="0" <%if rs("addip")<>"1" then%>checked<%end if%>>��&nbsp;&nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>������IP�б��е�IP��ַ���޷�������Ϣ</span></a></td>
          </tr>
              <tr> 
            <td height="22" align="right"><font color=red>��Ա��¼����������ƣ�</font></td>
            <td height="22" colspan="3"><INPUT NAME="hydlip" TYPE="text" ID="hydlip" SIZE="3" VALUE="<% = rs("hydlip") %>">&nbsp;&nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>������½X�������������IP��0Ϊ������</span></a></td>
          </tr>
              <tr> 
            <td height="22" align="right"><font color=red>��Աȡ���������������ƣ�</font></td>
            <td height="22" colspan="3"><INPUT NAME="gpwip" TYPE="text" ID="gpwip" SIZE="3" VALUE="<% = rs("gpwip") %>">&nbsp;&nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>����ȡ��X�������������IP��0Ϊ������</span></a></td>
          </tr> 
              <tr> 
            <td height="22" align="right"><font color=red>��̨��¼����������ƣ�</font></td>
            <td height="22" colspan="3"><INPUT NAME="htdlip" TYPE="text" ID="htdlip" SIZE="3" VALUE="<% = rs("htdlip") %>">&nbsp;&nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>��¼��̨������������IP��0Ϊ�����ơ�</span></a></td>
          </tr>           
              <tr> 
            <td height="22" align="right"><font color=red>�������Դ������ƣ�</font></td>
            <td height="22" colspan="3"><INPUT NAME="bookip" TYPE="text" ID="bookip" SIZE="3" VALUE="<% = rs("bookip") %>">&nbsp;&nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>һ���Ự��20�������������Դ�������������IP��0Ϊ������</span></a></td>
          </tr>
    
              <tr> 
            <td height="22" align="right"><font color=red>����������Ϣ�������ƣ�</font></td>
            <td height="22" colspan="3"><INPUT NAME="infoip" TYPE="text" ID="infoip" SIZE="3" VALUE="<% = rs("infoip") %>">&nbsp;&nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>һ���Ự��20����������������Ϣ��������������IP��0Ϊ������</span></a></td>
          </tr>
              <tr> 
            <td height="22" align="right"><font color=red>��������SQLע����Ϊ�������ƣ�</font></td>
            <td height="22" colspan="3"><INPUT NAME="sqlip" TYPE="text" ID="sqlip" SIZE="3" VALUE="<% = rs("sqlip") %>">&nbsp;&nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>һ���Ự��20��������������SQLע����Ϊ��������������IP��0Ϊ������</span></a></td>
          </tr>
              <tr> 
            <td height="22" align="right">&nbsp;</td>
            <td height="22" colspan="3">&nbsp;</td>
          </tr>

              <tr> 
            <td height="22" align="right">�û�ע����ʱ�����ƣ�</td>
            <td height="22" colspan="3"><INPUT NAME="times" TYPE="text" ID="times" SIZE="3" VALUE="<% = rs("times") %>"> Сʱ��ֻ��ע��<INPUT NAME="zcNum" TYPE="text" ID="zcNum" SIZE="3" VALUE="<% = rs("zcNum") %>">&nbsp;&nbsp;��<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>ͬһip�޶�Сʱ��ע�����</span></a></td>
          </tr>
              <tr> 
            <td height="22" align="right">������Ϣ���ʱ�����ƣ�</td>
            <td height="22" colspan="3"><INPUT NAME="infosj" TYPE="text" ID="infosj" SIZE="3" VALUE="<% = rs("infosj") %>">&nbsp;&nbsp;��<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>����������Ϣ���ʱ�䣬��λΪ��</span></a></td>
          </tr>
              <tr> 
            <td height="22" colspan="4" align="center"><font color=DD5A26>�еķ�������װ�������ַ�����������������޷����棬�����������ջ򲿷�ɾ���������Ƶ��ַ����ɷ�����������ˡ�</font></td>
          </tr>
              <tr> 
            <td height="22" align="right">��Ϣ�����ַ����ƣ�</td>
            <td height="22" colspan="3"><INPUT NAME="killword" TYPE="text" ID="killword" SIZE="80" VALUE="<% = rs("killword") %>">&nbsp;&nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>��Ҫ���������Ʒ����Ƿ��ؼ��֣�����Ϣ���⺬�������ַ��Ķ����ܷ��������ַ����á�|'�ָ���ע�ⲻ���ԡ�|'��β�����еķ�������װ�������ַ����ƿ������޷����棬��ɾ��</span></a></td>
          </tr>
              <tr> 
            <td height="22" align="right">��Ϣ��ϸ�����ַ����ƣ�</td>
            <td height="22" colspan="3"><INPUT NAME="killmemo" TYPE="text" ID="killmemo" SIZE="80" VALUE="<% = rs("killmemo") %>">&nbsp;&nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>��Ҫ���������Ʒ����Ƿ��ؼ��֣�����Ϣ��ϸ���ݺ��������ַ��Ķ����ܷ������ַ����á�|'�ָ���ע�ⲻ���ԡ�|'��β��</span></a></td>
          </tr>  		  
              <tr> 
            <td height="22" align="right">��Ϣ��ϵ�˺͵�ַ�ַ����ƣ�</td>
            <td height="22" colspan="3"><INPUT NAME="killname" TYPE="text" ID="killname" SIZE="80" VALUE="<% = rs("killname") %>">&nbsp;&nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>��Ҫ���������Ʒ����Ƿ��ؼ��֣�����Ϣ��ϵ�˺��������ַ��Ķ����ܷ������ַ����á�|'�ָ���ע�ⲻ���ԡ�|'��β��</span></a></td>
          </tr>  
          
                                   
<tr><td colspan=4 width="568" align="center"><INPUT name="ok" TYPE="hidden" value="ok"><INPUT name=action TYPE="submit" value="��������"></form></table>


<DIV style="position:absolute; top:120; left:500"> 
<table border=0 width=100><tr><td width=50% align=center>
<form action=admin_ip.asp method=post name=setup2>
<INPUT name="rows" TYPE="hidden" value="<%=rows+10%>">
<input type=image src=images/jia.gif border=0 alt="���ӱ༭����">
</form>
</td>
<td width=100 align=center>
<form action=admin_ip.asp method=post name=setup3>
<INPUT name="rows" TYPE="hidden" value="<%=rows-10%>">
<input type=image src=images/jian.gif border=0 alt="���ٱ༭����">
</form>
</td></tr></table>
</td></tr>
</table>


<%Call CloseRs(rs)
conn.close
set conn=nothing
end If

if action="ok" then
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from DNJY_config"
rs.open sql,conn,1,3
rs("lockip")=request.form("lockip")
rs("addip")=request.form("addip")
rs("hydlip")=request.form("hydlip")
rs("gpwip")=request.form("gpwip")
rs("htdlip")=request.form("htdlip")
rs("bookip")=request.form("bookip")
rs("infoip")=request.form("infoip")
rs("sqlip")=request.form("sqlip")
rs("infosj")=request.form("infosj")
rs("killword")=request.form("killword")
rs("killmemo")=request.form("killmemo")
rs("killname")=request.form("killname")
rs("times")=request.form("times")
rs("zcNum")=request.form("zcNum")

rs.update
url="admin_ip.asp"
Call CloseRs(rs)
response.Write "<script language=javascript>alert('�����ɹ������Զ����������ļ�config.asp');location.href='admin_config.asp?page="&url&"';</script>"
end If

if action="ip" Then
Set rsip=Server.CreateObject("ADODB.Recordset")
sql="select * from DNJY_config"
rsip.open sql,conn,1,3
ip=replace(request.form("ip")," ","")
ip=replace(ip,vbCRLF,"@")
for i=1 to 5
ip=replace(ip,"@@","@")
next
if right(ip,1)="@" then ip=left(ip,len(ip)-1)
if left(ip,1)="@" then ip=right(ip,len(ip)-1)
rsip("IP")=ip
rsip.update
rsip.close
set rsip=nothing
Call CloseDB(conn)
response.Write "<script language=javascript>alert('�����ɹ���');location.href='admin_ip.asp';</script>"
end If
%>
