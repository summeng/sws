<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("02")%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
<style type="text/css">
/*��ʾ��Ϣ*/
.info {position:relative;background:#;color:#666; text-decoration:none;font-size:12px;width:0px;height:0px;text-align:center;border:0px solid #ccc;line-height:18px;}/*�������ӵ�����,һ��Ҫ����Ϊrelative����ʹ��ʾ�����������*/
.info:hover {background:#eee;color:#333;}
.info span {display: none }/*���������µ�spanΪ����״̬*/
.info:hover span /*����hover�µ�span����Ϊ����״̬,��������ʾ���λ��*/{display:block;position:absolute;top:0px;left:25px;width:200px;
border:1px solid #DC5826; background:#FFF3B4; color:#333;padding:5px;text-align:left;}
</style>
</head>
<BODY>

<%Call OpenConn
dim action,show,er
action=request("ok")
if action="" then
Set rs = conn.Execute("select * from DNJY_pay") 
%>
<form action=pay1.asp method=post name=pay>
<table width="98%" border="1" style="border-collapse: collapse; border-style: dotted; border-width: 0px" bordercolor="#333333" cellspacing="0" cellpadding="2">
<tr><td height=35  colspan=2> <a href=http://www.alipay.com target="_blank">֧����</a> </td></tr>
<tr><td height=35  width=20% align=right >֧�����ʻ�&nbsp;</td><td  > <input type="password" value="<%=rs("alipay_id")%>" name=alipay_id size="30"> &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>��д����֧�����ʻ��ʺ�����Ӧ�뵽�ʷ�ʽ��Ӧ����֧�����̳���</span></a>&nbsp;&nbsp;<a href=http://www.alipay.com target="_blank">����</a></td></tr>
<tr><td height=35  width=20% align=right >��ȫУ���� &nbsp;</td><td  > <input type="password" value="<%=rs("alipay_securityCode")%>" name=alipay_securityCode size="30"> &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>���½֧������վ���ڡ��̼ҷ����в�ð�ȫУ����</span></a></td></tr>
<tr><td height=35  width=20% align=right >�������ID &nbsp;</td><td  > <input type="password" value="<%=rs("partner")%>" name=partner size="30"> &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>���½֧������վ���ڡ��̼ҷ����в���ʻ���������ݣ�partnerID��</span></a></td></tr>
<tr><td height=35  width=20% align=right >���ʷ�ʽ &nbsp;</td><td  >
<%if rs("alipay_dz")=0 then%>
<input type="radio" name="alipay_dz" value="0" id="alipay_dz" checked>
��ʱ�����տ�&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="alipay_dz" value="1" id="alipay_dz">
���������տ�&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="alipay_dz" value="2" id="alipay_dz">
˫�����տ�&nbsp;&nbsp;&nbsp;&nbsp;
<%ElseIf rs("alipay_dz")=1 then%>                         
<input type="radio" name="alipay_dz" value="0" id="alipay_dz">
��ʱ�����տ�&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="alipay_dz" value="1" id="alipay_dz" checked>
���������տ�&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="alipay_dz" value="2" id="alipay_dz">
˫�����տ�&nbsp;&nbsp;&nbsp;&nbsp;
<%else%>
<input type="radio" name="alipay_dz" value="0" id="alipay_dz">
��ʱ�����տ�
<input type="radio" name="alipay_dz" value="1" id="alipay_dz">
���������տ�
<input type="radio" name="alipay_dz" value="2" id="alipay_dz" checked>
˫�����տ�
<%end if%>  &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>�տ�ʷ�ʽ������֧�����ʻ�������̼ҷ���������Ӧ</span></a></td></tr>
<tr><td height=35  width=20% align=right >���ÿ��� &nbsp;</td><td  >
<%if rs("alipay_kg")=0 then%>
<input type="radio" name="alipay_kg" value="0" id="alipay_kg" checked>
�ر�&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="alipay_kg" value="1" id="alipay_kg">
����
<%else%>                         
<input type="radio" name="alipay_kg" value="0" id="alipay_kg">
�ر�&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="alipay_kg" value="1" id="alipay_kg" checked>
����<%end if%>  &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>���û�رմ��տʽ</span></a></td></tr>


<tr><td height=35  colspan=2> <a href=http://www.tenpay.com target="_blank">�Ƹ�ͨ</a> </td></tr>
<tr><td height=35  width=20% align=right >�Ƹ�ͨ�̻���&nbsp;</td><td  > <input type="password" value="<%=rs("tenpay_id")%>" name=tenpay_id size="30"> &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>��д���ĲƸ�ͨ�ʻ�</span></a>&nbsp;&nbsp;<a href=http://www.tenpay.com target="_blank">����</a></td></tr>
<tr><td height=35  width=20% align=right >�Ƹ�ͨ�̻���Կ &nbsp;</td><td  > <input type="password" value="<%=rs("tenpay_key")%>" name=tenpay_key size="30"> &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>Ӧ��Ƹ�ͨ��վ�û���¼��̨���õ���Կ��ͬ</span></a></td></tr>
<tr><td height=35  width=20% align=right >��������˵�� &nbsp;</td><td  > <input type=text value="<%=rs("tenpay_name")%>" name=tenpay_name size="30"> &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>��д��Ʒ���߷����˵�������������д</span></a></td></tr>
<tr><td height=35  width=20% align=right >���ʷ�ʽ &nbsp;</td><td  >
<%if rs("tenpay_dz")=0 then%>
<input type="radio" name="tenpay_dz" value="0" id="tenpay_dz" checked>
��ʱ����&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="tenpay_dz" value="1" id="tenpay_dz">
��������
<%else%>                         
<input type="radio" name="tenpay_dz" value="0" id="tenpay_dz">
��ʱ����&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="tenpay_dz" value="1" id="tenpay_dz" checked>
��������<%end if%>  &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>�տ�ʷ�ʽ������Ƹ�ͨ�ʻ������������Ӧ</span></a></td></tr>
<tr><td height=35  width=20% align=right >���ÿ��� &nbsp;</td><td  >
<%if rs("tenpay_kg")=0 then%>
<input type="radio" name="tenpay_kg" value="0" id="tenpay_kg" checked>
�ر�&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="tenpay_kg" value="1" id="tenpay_kg">
����
<%else%>                         
<input type="radio" name="tenpay_kg" value="0" id="tenpay_kg">
�ر�&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="tenpay_kg" value="1" id="tenpay_kg" checked>
����<%end if%>  &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>���û�رմ��տʽ</span></a></td></tr>
<tr><td height=35  width=20% align=right >�����ʺ� &nbsp;</td><td  > ע���Ƹ�ͨ�ṩ�ļ�ʱ���ʡ��н鵣����˫�ӿ�ͨ�ò����ʺ�<br>
�̻��ţ�1900000113<br>
�̻����ƣ������̻������ʻ�<br>
��Կ��e82573dc7e6136ba414f2e2affbe39fa<br>
ע�⣺�벻Ҫ֧���������ϲ����ʺţ���Ѷ�Ƹ�ͨˡ���˿,�мǣ���ʽ����ʱ�����������̻��ż���Կ�滻��&nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>�Ƹ�ͨ�ṩ�ļ�ʱ���ʡ��н鵣����˫�ӿ�ͨ�ò����ʺ�</span></a></td></tr>

<tr><td height=35  colspan=2> <a href=http://www.yeepay.com target="_blank">�ױ�֧��</a>&nbsp;</td></tr> 
<tr><td height=35  width=20% align=right>�̻����&nbsp;</td><td  > <input type="password" value="<%=rs("yeepayid")%>" name=yeepayid size="30"> &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>��д�����ױ�֧��������̻�ID</span></a>&nbsp;&nbsp;<a href=http://www.yeepay.com target="_blank">����</a></td></tr>
<tr><td height=35  width=20% align=right >��Կ(key)&nbsp;</td><td  > <input type="password" value="<%=rs("yeepaykey")%>" name=yeepaykey size="60"> &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>Ӧ���ױ�֧������Կһ��</span></a></td></tr>

<tr><td height=35  width=20% align=right >���ÿ��� &nbsp;</td><td  >
<%if rs("yeepay_kg")=0 then%>
<input type="radio" name="yeepay_kg" value="0" id="yeepay_kg" checked>
�ر�&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="yeepay_kg" value="1" id="yeepay_kg">
����
<%else%>                         
<input type="radio" name="yeepay_kg" value="0" id="yeepay_kg">
�ر�&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="yeepay_kg" value="1" id="yeepay_kg" checked>
����<%end if%>  &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>���û�رմ��տʽ</span></a></td></tr>


<tr><td height=35  colspan=2 > <a href=http://www.chinabank.com.cn target="_blank">��������</a>&nbsp;</td></tr>
<tr><td height=35  width=20% align=right >�� �� ��&nbsp;</td><td  > <input type="password" value="<%=rs("chinaid")%>" name=chinaid size="30"> &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>��д������������������̻���һ��Ϊ4��5λ����</span></a>&nbsp;&nbsp;<a href=http://www.chinabank.com.cn target="_blank">����</a></td></tr>
<tr><td height=35  width=20% align=right >MD5 ˽Կ&nbsp;</td><td  > <input type="password" value="<%=rs("chinakey")%>" name=chinakey size="30"> &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>Ӧ��������̨���õ���Կһ��</span></a></td></tr>
<tr><td height=35  width=20% align=right rowspan=1 >���ص�ַ&nbsp;</td><td  > <input type=text value="http://<%=web%>/<%=strInstallDir%>chinabank/Receive.asp" name=chinaback size=45> &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>����֧���ɹ���ķ���֪ͨҳ��ϵͳ�����ô˴�û����</span></a>&nbsp;</td></tr>
<tr><td height=35  width=20% align=right >���ÿ��� &nbsp;</td><td  >
<%if rs("china_kg")=0 then%>
<input type="radio" name="china_kg" value="0" id="china_kg" checked>
�ر�&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="china_kg" value="1" id="china_kg">
����
<%else%>                         
<input type="radio" name="china_kg" value="0" id="china_kg">
�ر�&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" name="china_kg" value="1" id="china_kg" checked>
����<%end if%>  &nbsp; &nbsp;&nbsp; <a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>���û�رմ��տʽ</span></a></td></tr>




<tr><td height=35  colspan=2><INPUT name="ok" TYPE="hidden" value="ok"><INPUT name=action TYPE="submit" value="��������">&nbsp;</td></tr>
</table>
</form>
<%
rs.close
conn.close
set rs=nothing
set conn=nothing
end if
              
if action="ok" then

if request.form("yeepay_kg")=1 and request.form("yeepayid")="" then er=1
if request.form("china_kg")=1 then
if request.form("chinaid")="" or request.form("chinakey")="" then er=2
end if
if request.form("alipay_kg")=1 then
if request.form("alipay_id")="" or request.form("alipay_securityCode")="" or request.form("partner")="" then er=7
end if
if request.form("tenpay_kg")=1 then
if request.form("tenpay_id")="" or request.form("tenpay_key")="" or request.form("tenpay_name")="" then er=9
end if
if er=1 then
response.write "<script language='javascript'>"
response.write "alert('�����ˣ���ѡ���ˡ��ױ�֧������������дyeepaypay�̻���ţ�');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
if er=2 then
response.write "<script language='javascript'>"
response.write "alert('�����ˣ���ѡ���ˡ��������ߡ���������д�����̻�ID��MD5˽Կ��');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if

if er=7 then
response.write "<script language='javascript'>"
response.write "alert('�����ˣ���ѡ���ˡ�֧��������������д֧�����ʻ�����ȫУ���롢�������ID��');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if

if er=9 then
response.write "<script language='javascript'>"
response.write "alert('�����ˣ���ѡ���ˡ��Ƹ�ͨ����������д�Ƹ�ͨ�̻��ʺš���Կ����Ʒ�ͷ���˵����');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from DNJY_pay"
rs.open sql,conn,1,3

rs("yeepayid")=request.form("yeepayid")
rs("yeepaykey")=request.form("yeepaykey")
'rs("yeepay_dz")=request.form("yeepay_dz")
rs("yeepay_kg")=request.form("yeepay_kg")

rs("chinaid")=request.form("chinaid")
rs("chinakey")=request.form("chinakey")
rs("chinaback")=request.form("chinaback")
'rs("china_dz")=request.form("china_dz")
rs("china_kg")=request.form("china_kg")

rs("alipay_id")=Trim(request.form("alipay_id"))
rs("alipay_securityCode")=Trim(request.form("alipay_securityCode"))
rs("partner")=Trim(request.form("partner"))
rs("alipay_dz")=request.form("alipay_dz")
rs("alipay_kg")=request.form("alipay_kg")

rs("tenpay_id")=request.form("tenpay_id")
rs("tenpay_name")=request.form("tenpay_name")
rs("tenpay_key")=request.form("tenpay_key")
rs("tenpay_dz")=request.form("tenpay_dz")
rs("tenpay_kg")=request.form("tenpay_kg")

rs.update
Call CloseRs(rs)
Call CloseDB(conn)
response.write "<script language='javascript'>"
response.write "alert('�����ɹ��������õ�����֧����Ϣ�ѱ��棡');"
response.write "location.href='pay1.asp';"
response.write "</script>"
end if%>
