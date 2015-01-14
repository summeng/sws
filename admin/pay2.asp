<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("02")%>
<script language=javascript src=../Include/mouse_on_title.js></script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/manage.css" type="text/css">
</head>
<BODY>
<%dim key,action,er,bank,url,bank1,bank2,bank3,bank4,bank5,bank6,bank7,bank8,bank9,bank10,bank11,bank12,bank13,bank14,bank15,bankN,N
'on error resume next
key=request("key")
'显示银行帐号信息
Call OpenConn
if key="" Then

set rs=server.CreateObject("adodb.recordset")
sql="select * from DNJY_pay"
rs.open sql,conn,1,1
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


  if er=15 then'如果er=15，则表明没有全部银行帐号信息为空
response.write "<script language='javascript'>"
response.write "alert('您还没有添加银行帐号信息，请单击“确定”开始添加。！');"
response.write "location.href='pay2.asp?ok=edit';"
response.write "</script>"
response.end
  end if

'显示银行帐号信息
response.write "<table width=98% border='1'  style='border-collapse: collapse; border-style: dotted; border-width: 0px'  bordercolor='#333333' cellspacing='0' cellpadding='2'>"
response.write "<form name=""edit"" method=""POST"" action=pay2.asp?key=edit>"
response.write "<tr class=backs><td colspan=2 bgcolor='#BDBEDE' height=38>银行帐号设置</td></tr>"

for N=1 to 15
bankN="bank"&N
if rs(bankN)="" or isnull(rs(bankN)) then
else
response.write "<tr><td width=""20%"" valign='middle' align='center'><img border=0 src='../IMAGES/bank/"&bankN&".gif'></td><td  valign='middle'>"&rs(bankN)&"</td></tr>"
end if
next

response.write "<tr><td colspan=2 height='30' ><input type=""submit"" name=""edit"" value=""修改内容"" ></td></tr></form></table>"

Call CloseRs(rs)
Call CloseDB(conn)
end if


'编辑帐号信息
if key="edit" then 

set rs=server.CreateObject("adodb.recordset")
sql="select * from DNJY_pay"
rs.open sql,conn,1,1 
  
%>
<table width="98%" border="1"  style="border-collapse: collapse; border-style: dotted; border-width: 0px"  bordercolor="#333333" cellspacing="0" cellpadding="2">
<form name="yes" method="POST" action=pay2.asp?key=yes >
<tr class=backs><td colspan=2 class=td height=18>银行帐号设置</td></tr>

<tr><td width="20%" align=right valign="middle" align="center"><img border="0" src="../IMAGES/bank/bank1.gif">&nbsp;</td><td  valign="middle"> <textarea NAME="bank1" ROWS=8 cols="50"  style="overflow:auto;width=70%"><%=replace(rs("bank1"),"<BR>",vbCRLF)%></textarea>&nbsp;</td></tr>
  <tr>
  <td width="20%" align=right valign="middle" align="center"><img border="0" src="../IMAGES/bank/bank2.gif">&nbsp;</td><td  valign="middle"> <textarea NAME="bank2" ROWS=8 cols="50"  style="overflow:auto;width=70%"><%=replace(rs("bank2"),"<BR>",vbCRLF)%></textarea>
  &nbsp;</td>
  </tr>
  <tr>
  <td width="20%" align=right valign="middle" align="center"><img border="0" src="../IMAGES/bank/bank3.gif">&nbsp;</td><td  valign="middle"> <textarea NAME="bank3" ROWS=8 cols="50"  style="overflow:auto;width=70%"><%=replace(rs("bank3"),"<BR>",vbCRLF)%></textarea>
  &nbsp;</td>
  </tr>
  <tr>
  <td width="20%" align=right valign="middle" align="center"><img border="0" src="../IMAGES/bank/bank4.gif">&nbsp;</td><td  valign="middle"> <textarea NAME="bank4" ROWS=8 cols="50"  style="overflow:auto;width=70%"><%=replace(rs("bank4"),"<BR>",vbCRLF)%></textarea>
  &nbsp;</td>
  </tr>
  <tr>
  <td width="20%" align=right valign="middle" align="center"><img border="0" src="../IMAGES/bank/bank5.gif">&nbsp;</td><td  valign="middle"> <textarea NAME="bank5" ROWS=8 cols="50"  style="overflow:auto;width=70%"><%=replace(rs("bank5"),"<BR>",vbCRLF)%></textarea>
  &nbsp;</td>
  </tr>
  <tr>
  <td width="20%" align=right valign="middle" align="center"><img border="0" src="../IMAGES/bank/bank6.gif">&nbsp;</td><td  valign="middle"> <textarea NAME="bank6" ROWS=8 cols="50"  style="overflow:auto;width=70%"><%=replace(rs("bank6"),"<BR>",vbCRLF)%></textarea>
  &nbsp;</td>
  </tr>
  <tr>
  <td width="20%" align=right valign="middle" align="center"><img border="0" src="../IMAGES/bank/bank7.gif">&nbsp;</td><td  valign="middle"> <textarea NAME="bank7" ROWS=8 cols="50"  style="overflow:auto;width=70%"><%=replace(rs("bank7"),"<BR>",vbCRLF)%></textarea>
  &nbsp;</td>
  </tr>
  <tr>
  <td width="20%" align=right valign="middle" align="center"><img border="0" src="../IMAGES/bank/bank8.gif">&nbsp;</td><td  valign="middle"> <textarea NAME="bank8" ROWS=8 cols="50"  style="overflow:auto;width=70%"><%=replace(rs("bank8"),"<BR>",vbCRLF)%></textarea>
  &nbsp;</td>
  </tr>
  <tr>
  <td width="20%" align=right valign="middle" align="center"><img border="0" src="../IMAGES/bank/bank9.gif">&nbsp;</td><td  valign="middle"> <textarea NAME="bank9" ROWS=8 cols="50"  style="overflow:auto;width=70%"><%=replace(rs("bank9"),"<BR>",vbCRLF)%></textarea>
  &nbsp;</td>
  </tr>
  <tr>
  <td width="20%" align=right valign="middle" align="center"><img border="0" src="../IMAGES/bank/bank10.gif">&nbsp;</td><td  valign="middle"> <textarea NAME="bank10" ROWS=8 cols="50" style="overflow:auto;width=70%"><%=replace(rs("bank10"),"<BR>",vbCRLF)%></textarea>
  &nbsp;</td>
  </tr>
  <tr>
  <td width="20%" align=right valign="middle" align="center"><img border="0" src="../IMAGES/bank/bank11.gif">&nbsp;</td><td  valign="middle"> <textarea NAME="bank11" ROWS=8 cols="50" style="overflow:auto;width=70%"><%=replace(rs("bank11"),"<BR>",vbCRLF)%></textarea>
  &nbsp;</td>
  </tr>
  <tr>
  <td width="20%" align=right valign="middle" align="center"><img border="0" src="../IMAGES/bank/bank12.gif">&nbsp;</td><td  valign="middle"> <textarea NAME="bank12" ROWS=8 cols="50" style="overflow:auto;width=70%"><%=replace(rs("bank12"),"<BR>",vbCRLF)%></textarea>
  &nbsp;</td>
  </tr>
  <tr>
  <td width="20%" align=right valign="middle" align="center"><img border="0" src="../IMAGES/bank/bank13.gif">&nbsp;</td><td  valign="middle"> <textarea NAME="bank13" ROWS=8 cols="50" style="overflow:auto;width=70%"><%=replace(rs("bank13"),"<BR>",vbCRLF)%></textarea>
  &nbsp;</td>
  </tr>
  <tr>
  <td width="20%" align=right valign="middle" align="center"><img border="0" src="../IMAGES/bank/bank14.gif">&nbsp;</td><td  valign="middle"> <textarea NAME="bank14" ROWS=8 cols="50" style="overflow:auto;width=70%"><%=replace(rs("bank14"),"<BR>",vbCRLF)%></textarea>
  &nbsp;</td>
  </tr>

  <tr>
  <td width="20%" align=right valign="middle" align="center"><img border="0" src="../IMAGES/bank/bank15.gif">&nbsp;</td><td  valign="middle"> <textarea NAME="bank15" ROWS=8 cols="50" style="overflow:auto;width=70%"><%=replace(rs("bank15"),"<BR>",vbCRLF)%></textarea>
  &nbsp;</td>
  </tr>

<tr><td colspan=2 height="30" ><INPUT name="yes" TYPE="submit" value="保存设置">&nbsp;</td></tr>
</form></table>
<%Call CloseRs(rs)
Call CloseDB(conn)
end if

'修改保存
if key="yes" Then

if request("bank1")="" and  request("bank2")="" and  request("bank3")="" and  request("bank4")="" and request("bank5")="" and  request("bank6")="" and  request("bank7")="" and  request("bank8")="" and  request("bank9")="" and  request("bank10")="" and  request("bank11")="" and  request("bank12")="" and  request("bank13")="" and  request("bank14")="" then
response.write "<script language='javascript'>"
response.write "alert('至少应填写一种银行帐号，单击“确定”返回上一页！');"
response.write "location.href='javascript:history.go(-1)';"
response.write "</script>"
response.end
end if
 Set rs=Server.CreateObject("ADODB.Recordset")
 sql="select bank1,bank2,bank3,bank4,bank5,bank6,bank7,bank8,bank9,bank10,bank11,bank12,bank13,bank14,bank15 from DNJY_pay"
 rs.open sql,conn,1,3 

  bank=replace(request("bank1"),vbCRLF,"<BR>")
 rs("bank1")=bank

  bank=replace(request("bank2"),vbCRLF,"<BR>")

 rs("bank2")=bank

  bank=replace(request("bank3"),vbCRLF,"<BR>")

 rs("bank3")=bank

  bank=replace(request("bank4"),vbCRLF,"<BR>")

 rs("bank4")=bank

  bank=replace(request("bank5"),vbCRLF,"<BR>")

 rs("bank5")=bank

  bank=replace(request("bank6"),vbCRLF,"<BR>")

 rs("bank6")=bank

  bank=replace(request("bank7"),vbCRLF,"<BR>")

 rs("bank7")=bank

  bank=replace(request("bank8"),vbCRLF,"<BR>")

 rs("bank8")=bank

  bank=replace(request("bank9"),vbCRLF,"<BR>")

 rs("bank9")=bank

  bank=replace(request("bank10"),vbCRLF,"<BR>")

 rs("bank10")=bank

  bank=replace(request("bank11"),vbCRLF,"<BR>")

 rs("bank11")=bank

  bank=replace(request("bank12"),vbCRLF,"<BR>")

 rs("bank12")=bank

  bank=replace(request("bank13"),vbCRLF,"<BR>")

 rs("bank13")=bank

  bank=replace(request("bank14"),vbCRLF,"<BR>")

 rs("bank14")=bank

  bank=replace(request("bank15"),vbCRLF,"<BR>")

 rs("bank15")=bank

rs.update
Call CloseRs(rs)
Call CloseDB(conn)
response.Write "<script language=javascript>alert('更新成功!');location.replace('pay2.asp');</script>"
end If
%>

&nbsp;</td></tr>
</table> 
&nbsp;</td></tr>
</table>
</body></html>
