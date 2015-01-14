<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=top.asp-->
<!--#include file=usercookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
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
            <!--#include file=userleft.asp-->　
        </div></td>
		<td width="5">&nbsp;</td>
      <td width="760" height="299" align="center" valign="top"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" height="139">
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%" height="130" class="ty1">
  <tr bgcolor="#BDBEDE">
    <td height="28" colspan="12" align="center"><div align="left"><span class="style1"><span style="font-size: 9pt"><font size="3"><font size="2">&nbsp;</font>会员 <%=username%> 财务明细</span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;可使用金额<%=Amount%>元&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;已消费总额<%=tAmount%>元</div></td>
  </tr>

  <tr bgcolor="#eeeeee">
    <td width="5%" height="24" align="center" bgcolor="#eeeeee">编号</td> 
    <td width="10%" height="24" align="center">业务金额</td>
    <td width="10%" height="24" align="center">业务类型</td> 
    <td width="20%" height="24" align="center">操作时间</td>                
    <td width="30%" height="24" align="center">备注</td> 
    <td width="5%" height="24" align="center">操作者</td>
  
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
response.write "<table>还没有记录！</table>"
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
    <td  height="24" align="center" bordercolor="#999999" bgcolor="#FAFAFA" style="border-bottom-style: solid; border-bottom-width: 1"><%If rs("ywlx")="支出" then%><font color="#FF0000"><%elseif rs("ywlx")="撤帐" then%><font color="#0000FF"><%End if%><%=rs("ywlx")%></font></td>
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

  共有&nbsp;<font color="#CC5200"><%=Allrecord%></font>&nbsp;条记录 共 <font color="#CC5200"><%=Allpage%></font> 页 现在是第 <font color="#CC5200"><%=ThisPage%></font> 页
<%
if ThisPage<2 then     
response.write "<font color=""#808080"">首页</font>&nbsp;"
response.write "<font color=""#808080"">上一页</font>&nbsp;"     
else     
response.write "<a href=?page=1&t1="&trim(request("t1"))&"&"&C_ID&">首页</a>&nbsp;"
response.write "<a href=?page="&ThisPage-1&"&t1="&trim(request("t1"))&"&"&C_ID&">上一页</a>&nbsp;"     
end if
if Allpage-ThisPage<1 then     
response.write "<font color=""#808080"">下一页</font>&nbsp;"
response.write "<font color=""#808080"">尾页</font>&nbsp;"  
else     
response.write "<a href=?page="&(ThisPage+1)&"&t1="&trim(request("t1"))&"&"&C_ID&">下一页</a>&nbsp;"   
response.write "<a href=?page="&Allpage&"&t1="&trim(request("t1"))&"&"&C_ID&">尾页</a>&nbsp;"     
end if
%></td>
</tr>
          </table></td>
        </tr>
        <tr>
          <td width="99%" height="25" colspan="3"><p align="center"><font color="#0000FF">说明：如果您对财务流水有疑问请与管理员联系！</font></td>
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

