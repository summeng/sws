<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="config.asp"-->
<!--#include file=usercookies.asp-->
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
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>֧��ȷ��</title>
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<style type="text/css">
<!--
.style2 {color: #333333}
.style3 {color: #FF0000}
-->
</style>
</head>
<script language="javascript">
<!--

function urlform_onsubmit() {
if (document.urlform.zffs.value==0)
	{
	  alert("��ѡ��֧����ʽ��")
	  document.urlform.zffs.focus()
	  return false
	 }
}
// --></script>
<%Call OpenConn
dim username,ddhm,zffs,edit,aliname,cash,hkuse,Payment
username=request.cookies("dick")("username")
zffs=trim(request("zffs"))
ddhm=request.form("ddhm")
if request("edit")="yes" Then
If zffs="1" Then Payment="����֧��"
If zffs="2" then Payment="ũ�л��"
If zffs="3" then Payment="���л��"
If zffs="4" then Payment="���л��"
If zffs="5" then Payment="ũ�Ż��"
If zffs="6" then Payment="�ʾֻ��"
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_order] where username='"&username&"' and ddhm='"&ddhm&"'" 
rs.open sql,conn,1,3
if rs.eof or rs.bof then
response.write "<li>��������"
response.end
end if
rs("zffs")=zffs
rs("zfqr")=1
rs("Payment")=Payment
rs("data")=now()
rs.update
Call CloseRs(rs)
response.write "<script language=JavaScript>alert('�޸ĳɹ�,�ȴ�����Ա���ȷ�ϣ�');window.location='user_order_save.asp?ddhm="&ddhm&"&zffs="&zffs&"' </script>"
response.end
end if
%>

<%ddhm=request("ddhm")
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_order] where username='"&username&"' and ddhm='"&ddhm&"'" 
rs.open sql,conn,1,1
aliname=rs("aliname")
cash=rs("cash")
hkuse=rs("hkuse")
Call CloseRs(rs)
Call CloseDB(conn)
%>
<body topmargin="0" leftmargin="0">
<form name="urlform" method="post" action="user_order_save.asp?edit=yes" language="javascript" onSubmit="return urlform_onsubmit()">
          <tr>
          <td width="99%" height="5" colspan="3">
          <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td>˵�������������<%=hkuse%>���������ѳɹ�֧��������ң��������������֧��ȷ�ϣ����ѹ���Ա���Ѿ������Ǻ�ʵ��Ŀ�󽫾�������ʺ��ṩ��ط���</td>
              </tr>
          </table></td>
        </tr>    
        <br>  
           <tr>
            <td width="14%" height="25" align="left"><div align="center">�����ˣ�<%=aliname%></td>
            <td width="48%" height="25" align="left"> ��</td>
          </tr>
           <tr>
            <td width="14%" height="25" align="left"><div align="center">�������룺<%=ddhm%></td>
			<input type="hidden" name="ddhm" value="<%=ddhm%>">
            <td width="48%" height="25" align="left"> ��</td>
          </tr>
          <tr>
            <td width="14%" height="25" align="left"><div align="center">֧����<%=cash%>Ԫ�����</td>
            
          </tr>
          <tr>
            <td width="14%" height="25" align="left"><div align="center">֧����ʽ��</div></td>
            <td width="30%" height="25" align="left"><select size="1" name="zffs">
			    <option value="0" <%If zffs="0" then%>selected<%end if%>>ѡ��֧����ʽ</option>
                <option value="1" <%If zffs="1" then%>selected<%end if%>>����֧��</option>
                <option value="2" <%If zffs="2" then%>selected<%end if%>>ũ�л��</option>
                <option value="3" <%If zffs="3" then%>selected<%end if%>>���л��</option>
                <option value="4" <%If zffs="4" then%>selected<%end if%>>���л��</option>
                <option value="5" <%If zffs="5" then%>selected<%end if%>>ũ�Ż��</option>
                <option value="6" <%If zffs="6" then%>selected<%end if%>>�ʾֻ��</option>
            </select></td>
            <td width="48%" height="25" align="left"> ��</td>
          </tr>

          <tr>
            <td width="62%" height="25" align="left" colspan="2"><p align="center">
              <input class="inputb" type="submit" value="ȷ���ύ" name="B1">
            </td>
            <td width="20%" height="25" align="left"> ��</td>
          </tr>
        </form>
        <tr>
          <td width="99%" height="25" align="left" colspan="3"> ��</td>
        </tr>



