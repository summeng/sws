<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../Include/Function.asp"-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%
Dim id,i,M_Picw,M_Pich,picw,pich,dbpath
id=request("id")
dbpath="../"
Call OpenConn
IF Request.Form("submit")<>"" Then
	For I=1 to 50
	picw=picw&Strint(Request.Form("picw"&i))&"|"
	pich=pich&Strint(Request.form("pich"&i))&"|"
	Next
	picw="|"&Left(picw,Len(picw)-1)
	pich="|"&Left(pich,Len(pich)-1)
	Set Rs=server.createobject("adodb.recordset")
	Rs.Open "Select * From DNJY_Template Where id="&id&"",conn,1,3
	Rs("M_Picw")=picw
	Rs("M_Pich")=pich
	Rs.Update
	Closers(rs)
End IF	
Set Rs=Conn.Execute("Select M_Picw,M_Pich From DNJY_Template Where id="&id&"")
M_Picw=Split(Rs(0),"|")
M_Pich=Split(Rs(1),"|")
Closers(rs)
%>
<HTML>
<HEAD>
<LINK rel="stylesheet" type="text/css" href="../css/style.css">
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<META name="GENERATOR" content="Microsoft FrontPage 6.0">
<META name="ProgId" content="FrontPage.Editor.Document">
<TITLE>�������</title>
<style type="text/css">
<!--
body {
	background-color: #D6DFF7;
}
body,td,th {
	color: #135294;
}
-->
</style></HEAD>
<form name="form1" method="post" action="">
  <table width="101%" border="0" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
    <tr>
      <td height="20" colspan="3" align="center"><FONT color="#FFFFFF" style="font-size:14px">��  ��  ��  ��  ��  ��</FONT></td>
    </tr>
    <tr>
      <td height="20" colspan="3" align="center" bgcolor="#FFFFFF"><br>
        <table width="98%" border="0" cellpadding="1" cellspacing="1" bgcolor="#D6DFF7">
		<%for i=1 to 50%>
          <tr>
            <td width="18%" bgcolor="#FFFFFF">���<%=I%></td>
            <td width="34%" bgcolor="#FFFFFF">
			<input name="picw<%=i%>" value="<%=M_Picw(i)%>">
			��</td>
            <td width="48%" bgcolor="#FFFFFF">
			<input name="pich<%=i%>" value="<%=M_Pich(i)%>">
			�� <a href="city_photo_ad.asp">�ϴ�ͼƬ</a> </td>
          </tr>
<%Next%> <tr>
            <td colspan="3" align="center" bgcolor="#FFFFFF"><input type="submit" name="Submit" value="   ��   ��   ">
            <input type="button" name="Submit2" value="   ��   ��   " onClick="location='Template.asp'"/></td>
          </tr>
        </table>
      <br></td>
    </tr>
  </table>
</form>
</BODY>                    
</HTML> 
<%Call CloseDB(conn)%>