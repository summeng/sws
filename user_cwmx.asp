<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=top.asp-->
<!--#include file=usercookies.asp-->
<%
'=====================================
'���칩����Ϣ��վ����ϵͳ
'��������Ƽ������ҳ�Ʒ
'�ͷ�����http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Call OpenConn
dim aliorder,alimoney
aliorder=request.form("aliorder")
alimoney=request.form("alimoney")
username=request.cookies("dick")("username")
dim rscw,Amount,tAmount
set rscw=server.createobject("adodb.recordset")
sql = "select Amount,tAmount from [DNJY_user] where username='"&username&"'"
rscw.open sql,conn,1,1
Amount=rscw("Amount")
tAmount=rscw("tAmount")
rscw.close
set rscw=nothing%>
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<style type="text/css">
<!--
.style2 {color: #333333}
.style3 {color: #FF0000}
-->
</style>
</head>

<body topmargin="0" leftmargin="0">

<div align="center">
  <center>
  <table width="980" height="371" border="0" cellpadding="0" cellspacing="0" bordercolor="#111111" bgcolor="#FFFFFF" style="border-collapse: collapse">
    <tr>
      <td width="214" height="299" valign="top"><div align="center">
            <!--#include file=userleft.asp-->��
        </div></td>
		<td width="5">&nbsp;</td>
      <td width="760" height="299" align="center" valign="top"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" height="139">
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%" height="130" class="ty1">
  <tr bgcolor="#BDBEDE">
    <td height="28" colspan="12" align="center"><div align="left"><span class="style1"><span style="font-size: 9pt"><font size="3"><font size="2">&nbsp;</font>��Ա <%=username%> ������ϸ</span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��ʹ�ý��<%=Amount%>Ԫ&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�������ܶ�<%=tAmount%>Ԫ</div></td>
  </tr>

  <tr bgcolor="#eeeeee">
    <td width="5%" height="24" align="center" bgcolor="#eeeeee">���</td> 
    <td width="10%" height="24" align="center">ҵ����</td>
    <td width="10%" height="24" align="center">ҵ������</td> 
    <td width="20%" height="24" align="center">����ʱ��</td>                
    <td width="30%" height="24" align="center">��ע</td> 
    <td width="5%" height="24" align="center">������</td>
  
    </tr>
<%
dim k,dick,uid
k=0
dim ThisPage,Pagesize,Allrecord,Allpage
if request("page")="" then
  ThisPage=1		
else
  ThisPage=request("page")
end if
dick=request("dick")
set rs=server.createobject("adodb.recordset")
sql = "select * from DNJY_Financial where admincl=1 and username='"&username&"' order by id "&DN_OrderType&""
rs.open sql,conn,1,1
if rs.eof then
response.write "<table>��û�м�¼��</table>"
Call CloseRs(rs)
Call CloseDB(conn)
response.end
end If
rs.Pagesize=20
Pagesize=rs.Pagesize
Allrecord=rs.Recordcount
Allpage=rs.Pagecount
if ThisPage<1 then                           
ThisPage=1
end if
rs.move (ThisPage-1)*Pagesize
k=0
do while not rs.eof
uid=rs("id")
%>
  <tr bgcolor="#FFFFFF">
    <td  height="24" align="center" bordercolor="#999999" bgcolor="#FAFAFA" style="border-bottom-style: solid; border-bottom-width: 1"><div align="center"><%=k+1%></div></td>
    <td  height="24" align="center" bordercolor="#999999" style="border-bottom-style: solid; border-bottom-width: 1"><font color="#FF0000"><%=rs("ywje")%></font></td>
    <td  height="24" align="center" bordercolor="#999999" bgcolor="#FAFAFA" style="border-bottom-style: solid; border-bottom-width: 1"><%If rs("ywlx")="֧��" then%><font color="#FF0000"><%elseif rs("ywlx")="����" then%><font color="#0000FF"><%End if%><%=rs("ywlx")%></font></td>
    <td  height="24" align="center" bordercolor="#999999" bgcolor="#FAFAFA" style="border-bottom-style: solid; border-bottom-width: 1"><%=rs("addTime")%></td>	
   <td  height="24" align="center" bordercolor="#999999" bgcolor="#FAFAFA" style="border-bottom-style: solid; border-bottom-width: 1"><%=rs("czbz")%></td>
   <td  height="24" align="center" bordercolor="#999999" bgcolor="#FAFAFA" style="border-bottom-style: solid; border-bottom-width: 1"><%=rs("czz")%></td>
    </tr>
<%rs.movenext
k=k+1
if k>=Pagesize then exit do
loop
%>
</table>
 <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
<tr> 
<td width="595" height="25" align="center">

  ����&nbsp;<font color="#CC5200"><%=Allrecord%></font>&nbsp;����¼ �� <font color="#CC5200"><%=Allpage%></font> ҳ �����ǵ� <font color="#CC5200"><%=ThisPage%></font> ҳ
<%
if ThisPage<2 then     
response.write "<font color=""#808080"">��ҳ</font>&nbsp;"
response.write "<font color=""#808080"">��һҳ</font>&nbsp;"     
else     
response.write "<a href=?page=1&t1="&trim(request("t1"))&"&"&C_ID&">��ҳ</a>&nbsp;"
response.write "<a href=?page="&ThisPage-1&"&t1="&trim(request("t1"))&"&"&C_ID&">��һҳ</a>&nbsp;"     
end if
if Allpage-ThisPage<1 then     
response.write "<font color=""#808080"">��һҳ</font>&nbsp;"
response.write "<font color=""#808080"">βҳ</font>&nbsp;"  
else     
response.write "<a href=?page="&(ThisPage+1)&"&t1="&trim(request("t1"))&"&"&C_ID&">��һҳ</a>&nbsp;"   
response.write "<a href=?page="&Allpage&"&t1="&trim(request("t1"))&"&"&C_ID&">βҳ</a>&nbsp;"     
end if
%></td>
</tr>
          </table></td>
        </tr>
        <tr>
          <td width="99%" height="25" colspan="3"><p align="center"><font color="#0000FF">˵����������Բ�����ˮ�������������Ա��ϵ��</font></td>
        </tr>
      </table></td>
      <td width="4" align="center" valign="top" bgcolor="#e6e6e6"></td>
    </tr>
    <tr><%Call CloseRs(rs)%>
      <td height="26" colspan="3"><!--#include file=footer.asp--></td>
    </tr>
  </table>
  </center>
</div>
</body>
</html>

