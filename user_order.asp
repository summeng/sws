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
<%
dim aliorder,alimoney
%>
<%
aliorder=request.form("aliorder")
alimoney=request.form("alimoney")
%>

<head><link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
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
      <td width="760" height="299" align="center" valign="top"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" height="139" class="ty1">
<%
dim email,dianhua,name,zffs,zfqr,ddhm
username=request.cookies("dick")("username")
'set rs = Server.CreateObject("ADODB.RecordSet")
'sql="select name,email,dianhua from [DNJY_user] where username='"&username&"'" 
'rs.open sql,conn,1,1
'if rs.eof then
'response.write "<br>"
'response.write "<li>参数错误！"
'response.end
'end if
'email=rs("email")
'dianhua=rs("dianhua")
'name=rs("name")
'rs.close
%>
        <tr>
          <td width="99%" height="25" colspan="3"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="25">
              <tr>
                <td width="100%" height="25" bgcolor="#F2F2F2"><span class="style3">　<strong>我的订单：</strong></span></td>
              </tr>
              <tr>
                <td height="1" bgcolor="#CCCCCC"></td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td width="99%" height="55" colspan="3"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="100%">
              <tr>
                
                <td width="10%" align="center"><span class="style2">定单号码</span></td>
                <td width="10%" align="center"><span class="style2">购买项目</td>
                <td width="20%" align="center"><span class="style2">提交时间</span></td>
                <td width="10%" align="center"><span class="style2">应付金额</span></td>
                <td width="10%" align="center"><span class="style2">我要支付</td>                 
                <td width="10%" align="center"><span class="style2">支付方式</td>
				<td width="10%" align="center"><span class="style2">支付渠道</td> 
                <td width="10%" align="center"><span class="style2">汇款确认</td>
                <td width="10%" align="center"><span class="style2">审核状态</span></td>
              </tr>
<%dim k,ThisPage,Pagesize,Allrecord,Allpage
if request("page")="" then
  ThisPage=1		
else
  ThisPage=request("page")
end If
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_order] where username='"&username&"' order by data "&DN_OrderType&""
rs.open sql,conn,1,1
if rs.eof then
response.write "<tr><td><li>暂无记录</td></tr>"
Else
rs.Pagesize=20
Pagesize=rs.Pagesize
Allrecord=rs.Recordcount
Allpage=rs.Pagecount
if ThisPage<1 then                           
ThisPage=1
end if
rs.move (ThisPage-1)*Pagesize
k=0

do while not rs.eof%>
              <tr>

                <td align="center"><%=rs("ddhm")%></td>
                <td  align="center"><%=rs("hkuse")%></td>
                <td  align="center"><%=rs("data")%></td>                                
                <td  align="center"><%=rs("cash")%> 元</td>
                <td  align="center"><%if rs("admincl")=1 Or rs("zfqr")=1 then%>
                    <font color="#0000FF">已经支付</font>
                  <%else%>
                  <font color="#FF0000"><a href="bzzh.asp?ddhm=<%=rs("ddhm")%>&<%=C_ID%>">我要支付</a></font>
                  <%end if%></td>                
                <td  align="center">
                <%if rs("zffs")=0 then
                response.write "尚未汇款"
                elseif rs("zffs")=1 then
                response.write "在线支付"
                elseif rs("zffs")=2 then
                response.write "农行汇款"
                elseif rs("zffs")=3 then
                response.write "工行汇款"
                elseif rs("zffs")=4 then
                response.write "建行汇款"
                elseif rs("zffs")=5 then
                response.write "农信汇款"                                                
                elseif rs("zffs")=6 then
                response.write "邮局汇款"
                end if%>
                </td>
                <td align="center"><%=rs("Payment")%></td>
                <td  align="center"><%if rs("zfqr")=0 then%>
                <A href="javascript:win=open('user_order_save.asp?ddhm=<%=rs("ddhm")%>&<%=C_ID%>','offer','width=350,height=250,status=no,menubar=no,scrollbars=yes,top=150,left=200'); win.focus()">确认</a>
                <%else
                response.write "<font color='#0000FF'>已汇款</font>"
                end if%>
                </td>
                <td  align="center"><%if rs("admincl")=1 then%>
                    <font color="#0000FF">√</font>
                  <%else%>
                  <font color="#FF0000">ⅹ</font>
                  <%end if%>
                </td>
              </tr>
              <%
rs.movenext
k=k+1
if k>=Pagesize then exit do
loop
end if
%><tr> 
<td width="595" height="25" align="center" colspan="8">

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
          </table>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">

          </table>		  </td>
        </tr>
        <tr>
          <td width="99%" height="25" colspan="3"><p align="center"><font color="#0000FF">说明：您汇款后请点击汇款确认。√代表款已到且管理员已确认并入帐，</font><font color="#FF0000">ⅹ代表正在等待管理员处理</font></td>
        </tr>
      </table></td>
      <td width="4" align="center" valign="top" bgcolor="#e6e6e6"></td>
    </tr>
    <tr><%
Call CloseRs(rs)%>
      <td height="26" colspan="3"><!--#include file=footer.asp--></td>
    </tr>
  </table>
  </center>
</div>
</body>
</html>

