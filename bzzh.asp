<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=usercookies.asp-->
<!--#include file=Include/mail.asp-->
<!--#include file=config.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
Call OpenConn
ddhm=request("ddhm")
Dim username,aliname,alimoney,dianhua,CompPhone,hkuse,email,RecAddress,Zipcode,oicq,alibody,ddhm,cash,shijian,dizhi,youbian,partner,n
set rs=server.createobject("adodb.recordset")
sql="select * from [DNJY_order] where ddhm='"&ddhm&"'" 
rs.open sql,conn,1,1
if rs.eof then
response.write "<tr><td><li>���޼�¼</td></tr>"
else
username=rs("username")
aliname=rs("aliname")
ddhm=rs("ddhm")
alimoney=rs("cash")
dianhua=rs("dianhua")
CompPhone=rs("CompPhone")
hkuse=rs("hkuse")
email=rs("email")
RecAddress=rs("RecAddress")
Zipcode=rs("Zipcode")
oicq=rs("oicq")
alibody=rs("alibody")
end if
Call CloseRs(rs)
%>
<table width="760" border="1" cellspacing="0" cellpadding="1" bordercolor="#111111" align="center"> 
<tr> <td>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1"> 
<tr><a href="index.asp">������ҳ</a> <td height="25"  align="center"><FONT  SIZE="3"><B><font color=#FF6600><% =aliname %></font>�����Ķ�����ϸ��Ϣ���£�</B></FONT></td></tr> 
<tr> <td height="18" >
<table border="0" cellspacing="1" cellpadding="0" align="center" valign=absmiddle width=100%>

              <tr align="center" BGCOLOR=#74C200> 
                <td>�����ţ�</td>
                <td>��ֵ�ʺ�</td>
                <td>����</td>
                <td>��ֵ��Ԫ��</td>
                <td>��������</td>                                 
                <td>��ϵ�绰</td> 
                <td>��ϵ�ֻ�</td>                
                <td>��������</td>
                <td>����OICQ</td>              
                
              </tr>
              <tr BGCOLOR=#74C200>                 
                <td align="center"><%=ddhm %></td>
                <td align="center"><%=username%></td>
                <td align="center"><%=aliname%></td> 
                <td align="center"><%=alimoney%></td>
                <td align="center"><%=hkuse%></td>                
                <td align="center"><%=dianhua%></td>
                <td align="center"><%=CompPhone%></td>
                <td align="center"><%=email%></td>
                <td align="center"><%=oicq%></td>               
   
              </tr>
            </table></td></tr> 
           <table align="center">��ѡ�����Ļ�ʽ������������ע������<p>
            </table>
<%dim alipay_id,chinaid,chinakey,chinaback,china_dz,china_kg,yeepayid,yeepaykey,yeepay_dz,yeepay_kg,pay,alipay_securityCode,aliorder,alipay_name,alipay_dz,alipay_kg,tenpay_id,tenpay_key,tenpay_name,tenpay_dz,tenpay_kg
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_pay",conn,1,1   
pay= rs("pay")
chinaid=rs("chinaid")
chinakey=rs("chinakey")
chinaback=rs("chinaback")
china_dz=rs("china_dz")
china_kg=rs("china_kg")
alipay_id=rs("alipay_id")
alipay_securityCode=rs("alipay_securityCode")
partner=rs("partner")
alipay_dz=rs("alipay_dz")
alipay_kg=rs("alipay_kg")
yeepayid=rs("yeepayid")
yeepaykey=rs("yeepaykey")
yeepay_dz=rs("yeepay_dz")
yeepay_kg=rs("yeepay_kg")
tenpay_id=rs("tenpay_id")
tenpay_key=rs("tenpay_key")
tenpay_name=rs("tenpay_name")
tenpay_dz=rs("tenpay_dz")
tenpay_kg=rs("tenpay_kg")    
Call CloseRs(rs)

'����֧����ʼ
dim inBillNo,total
inBillNo=ddhm '������
total=alimoney '֧�����
%>
<table border="0" cellspacing="1" cellpadding="0" align="center" valign=absmiddle width=100%>
<tr><td  width=50% height=120>
<br>
<%if alipay_kg=1 then'=====֧����֧����ʼ%>                            
<table border="0" cellspacing="1" cellpadding="0" align="center" valign=absmiddle width=760>
<tr><td  width=20% height=80>
<%If alipay_dz=0 then%>
<form method="post" action="alipay/js_asp_gb/index.asp?ddhm=<%=ddhm%>&hkuse=<%=hkuse%>&alimoney=<%=alimoney%>" target=new><input type=submit  value="֧����֧��--��ʱ�����տ�"></form>
<%ElseIf alipay_dz=1 then%>
<form method="post" action="alipay/db_asp_gb/index.asp?ddhm=<%=ddhm%>&hkuse=<%=hkuse%>&alimoney=<%=alimoney%>" target=new><input type=submit  value="֧����֧��--���������տ�"></form>
<%else%>
<form method="post" action="alipay/sj_asp_gb/index.asp?ddhm=<%=ddhm%>&hkuse=<%=hkuse%>&alimoney=<%=alimoney%>" target=new><input type=submit  value="֧����֧��--˫�����տ�"></form>
<%End if%>
</td></tr>
</table>
<%End If

if tenpay_kg=1 then%>
<table border="0" cellspacing="1" cellpadding="0" align="center" valign=absmiddle width=760>
<tr><td  width=20% height=80>
<%If tenpay_dz=0 then%>
<form method="post" action="tenpay/b2c/index.asp?ddhm=<%=ddhm%>&hkuse=<%=hkuse%>&alimoney=<%=alimoney%>" target=new><input type=submit  value="�Ƹ�֧ͨ��--��ʱ����"></form>
<%else%>
<form method="post" action="tenpay/med/index.asp?ddhm=<%=ddhm%>&hkuse=<%=hkuse%>&alimoney=<%=alimoney%>" target=new><input type=submit  value="�Ƹ�֧ͨ��--��������"></form>
<%End if%>
</td></tr>
</table>
<%End If
if china_kg=1 then
dim v_rcvtel,v_rcvmobile,v_orderemail,v_rcvname,v_rcvaddr,v_rcvpost,v_oid,v_amount,remark1
v_rcvtel=dianhua'�绰
v_rcvmobile=CompPhone'�ֻ�
v_orderemail = email '����
v_rcvname =  aliname'�ջ���
v_rcvaddr =  RecAddress'��ַ
v_rcvpost = Zipcode'�ʱ�
v_oid=inBillNo'������.
v_amount=replace(total,",","")'������
remark1=webname&inBillNo&"�Ŷ���"'��עһ
%>
<table border="0" cellspacing="1" cellpadding="0" align="center" valign=absmiddle width=760>
<tr><td  width=20% height=80>
<FORM name=formbill action="chinabank/Send.asp" method=post target=new>
<input type="submit" name="v_action" value="��������--��ʱ����">
<input type="hidden" name="v_oid" value="<%=v_oid%>">
<input type="hidden" name="v_amount" value="<%=v_amount%>">
<input type="hidden" name="v_rcvtel" value="<%=v_rcvtel%>">
<input type="hidden" name="v_rcvmobile" value="<%=v_rcvmobile%>">
<input type="hidden" name="v_rcvname" value="<%=v_rcvname%>">
<input type="hidden" name="v_rcvaddr" value="<%=v_rcvaddr%>">
<input type="hidden" name="v_rcvpost" value="<%=v_rcvpost%>">
<input type="hidden" name="v_orderemail"  value="<%=v_orderemail%>">
<input type="hidden" name="remark1"  value="<%=remark1%>">
</form>
</td></tr>
</table>
<%End if

If yeepay_kg=1 then%>
<table border="0" cellspacing="1" cellpadding="0" align="center" valign=absmiddle width=100%>
<tr><td width=50% height=120>
<form method="post" action="yeepay/pay.asp?ddhm=<%=ddhm%>&hkuse=<%=hkuse%>&alimoney=<%=alimoney%>" target=new><input type=submit  value="�ױ�֧��--��ʱ����"></form>
</table>
</td></tr>
</table>
<%End if%>

<table border="0" cellspacing="1" cellpadding="0" align="center" valign=absmiddle width=100%>
<tr>��������߸���������ѡ�����´�ͳ��ʽ:<td  width=50%  height=150>
<center>
<%Dim er,bankN
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_pay ",conn,1,3
er=0
if rs("bank1")="" or isnull(rs("bank1")) then er=er+1
if rs("bank2")="" or isnull(rs("bank2")) then er=er+1
if rs("bank3")="" or isnull(rs("bank3")) then er=er+1
if rs("bank4")="" or isnull(rs("bank4")) then er=er+1
if rs("bank5")="" or isnull(rs("bank5")) then er=er+1
if rs("bank6")="" or isnull(rs("bank6")) then er=er+1
if rs("bank7")="" or isnull(rs("bank7")) then er=er+1
if rs("bank8")="" or isnull(rs("bank8")) then er=er+1
if rs("bank9")="" or isnull(rs("bank9")) then er=er+1
if rs("bank10")="" or isnull(rs("bank10")) then er=er+1
if rs("bank11")="" or isnull(rs("bank11")) then er=er+1
if rs("bank12")="" or isnull(rs("bank12")) then er=er+1
if rs("bank13")="" or isnull(rs("bank13")) then er=er+1
if rs("bank14")="" or isnull(rs("bank14")) then er=er+1
if rs("bank15")="" or isnull(rs("bank15")) then er=er+1


  if er=15 then '���er=15�������û��ȫ�������ʺ���ϢΪ��
response.write "<script language='javascript'>"
response.write "alert('����û����������ʺ���Ϣ���뵥����ȷ������ʼ��ӡ���');"
response.write "location.href='pay2.asp?ok=edit';"
response.write "</script>"
response.end
  end if

'��ʾ�����ʺ���Ϣ
response.write "<table width=98% border='1'  style='border-collapse: collapse; border-style: dotted; border-width: 0px'  bordercolor='#333333' cellspacing='0' cellpadding='2'>"
response.write "<form action=pay2.asp method=post name=setup>"


for N=1 to 15
bankN="bank"&N
if rs(bankN)="" or isnull(rs(bankN)) then
else
response.write "<tr><td width=20% valign='middle' align='center'><img border=0 src='IMAGES/bank/"&bankN&".gif'></td><td  valign='middle'>"&rs(bankN)&"</td></tr>"
end if
next
response.write "</table>"
Call CloseRs(rs)
Call CloseDB(conn)%>
</td></tr>
</table>  


