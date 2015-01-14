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
<link rel="stylesheet" type="text/css" href="/<%=strInstallDir%>css/css.css">
<style type="text/css">
<!--
.style1 {
	color: #FF0000;
	font-weight: bold;
}
-->
</style>
</head>
<SCRIPT language=JavaScript>
function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.name != 'chkall' )
       e.checked = form.chkall.checked;
    }
  }
//-->
</SCRIPT>
<body topmargin="0" leftmargin="0">
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="980" height="371">
    <tr>
      <td width="214" valign="top" bgcolor="#FFFFFF"><div align="center"><!--#include file=userleft.asp--></div></td>
	  <td width="5">&nbsp;</td>

<%If request("key")="del" then
dim str2
username=request.cookies("dick")("username")
id=trim(request("selectedid"))
if trim(id)="" then
response.write "<script language=JavaScript>" & chr(13) & "alert('没有选择记录！');" & "history.back()" & "</script>"
response.end
end if
str2=split(id,",")
set rs=server.createobject("adodb.recordset")
for i=0 to ubound(str2)
sql="delete from [DNJY_sc_shops] where username='"&username&"' and id="&cstr(str2(i))
rs.open sql,conn,1,3
next
response.write "<script language=JavaScript>" & chr(13) & "alert('删除收藏成功！');" & "window.location='user_favshops.asp?"&C_ID&"'" & "</script>"
response.end
Call CloseRs(rs)
End if%>
      <td width="780" valign="top" bgcolor="#FFFFFF"><table width="100%" height="114" border="0" align="right" cellpadding="0" cellspacing="0" bordercolor="#111111" style="border-collapse: collapse" class="ty1">
        <FORM name=thisForm action="user_favshops.asp?key=del&<%=C_ID%>" method=POST>
          <tr>
            <td width="50" height="26" align="center" class="ty2"><p class="style1">编号</td>
			<td width="100" height="13" align="center" class="ty2"><span class="style1"> 掌柜ID</span></td>
            <td width="100" height="13" align="center" class="ty2"><span class="style1"> 掌柜姓名</span></td>
            <td width="150" height="13" align="center" class="ty2"><span class="style1"> 店铺名称</span></td>
            <td width="150" height="13" align="center" class="ty2"><span class="style1"> 企业名片</span></td>
            <td width="200" height="13" align="center" class="ty2"><span class="style1"> 收藏时间</span></td>
            <td width="5" height="13" align="center" class="ty2"> 　</td>
          </tr>

          <%
dim ThisPage,Pagesize,Allrecord,Allpage
i=1
if request("page")="" then
  ThisPage=1		
else
  ThisPage=request("page")
end if
username=request.cookies("dick")("username")
set rs = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_sc_shops] where username='"&username&"' order by scsj "&DN_OrderType&"" 
rs.open sql,conn,1,1
if rs.eof then
response.write "<tr><td></td><td><li>没有收藏记录</td></tr>"
else
rs.Pagesize=20
Pagesize=rs.Pagesize
Allrecord=rs.Recordcount
Allpage=rs.Pagecount
if ThisPage<1 then                           
ThisPage=1
end if
On Error Resume Next
rs.move (ThisPage-1)*Pagesize
do while not rs.eof
%>
          <tr>
            <td width="50" height="26" align="center" background="img/22.gif" ><%=i%> </td>
            <td width="100" height="25" align="center" background="img/22.gif" ><%=rs("scid")%> </td>
            <td width="100" height="25" align="center" background="img/22.gif" ><%=rs("name")%></td>
            <td width="150" height="25" align="center" background="img/22.gif" ><p align="left"><a target="_blank" href="user/com.asp?com=<%=rs("scid")%>&<%=C_ID%>"><%=rs("sdname")%></a></td>
            <td width="150" height="25" align="center" background="img/22.gif" ><p align="left"><a target="_blank" href="company.asp?username=<%=rs("scid")%>&<%=C_ID%>"><%=rs("mpname")%></a></td>
            <td width="200" height="25" align="center" background="img/22.gif" ><%= year(rs("scsj"))&"年"&month(rs("scsj"))&"月"&day(rs("scsj"))&"日" %> </td>
            <td width="5" height="25" align="center" background="img/22.gif" ><input type="checkbox" name="selectedid" value="<%=rs("id")%>"></td>
          </tr>
          <%
i=i+1
rs.movenext
if i>=Pagesize then exit do
loop
end if
%>
          <tr>
            <td height="25" colspan="7"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                <tr>
                  <td height="25" colspan="7"><p align="right">
                      <INPUT onclick=CheckAll(this.form) type=checkbox value=on name=chkall>
                    选中所有记录&nbsp;&nbsp;
                    <input  type="submit" value="删除" name="B1">
                        </td>
                </tr>
                <tr>
                  <td height="25"  colspan="7" align="center"> 共有&nbsp;<font color="#CC5200"><%=Allrecord%></font>&nbsp;条记录  共 <font color="#CC5200"><%=Allpage%></font> 页 现在是第 <font color="#CC5200"><%=ThisPage%></font> 页 <%
if ThisPage<2 then     
response.write "<font color=""#808080"">首页</font>&nbsp;"
response.write "<font color=""#808080"">上一页</font>&nbsp;"     
else     
response.write "<a href=?page=1&"&C_ID&">首页</a>&nbsp;"
response.write "<a href=?page="&ThisPage-1&"&"&C_ID&">上一页</a>&nbsp;"     
end if
if Allpage-ThisPage<1 then     
response.write "<font color=""#808080"">下一页</font>&nbsp;"
response.write "<font color=""#808080"">尾页</font>&nbsp;"  
else     
response.write "<a href=?page="&(ThisPage+1)&"&"&C_ID&">下一页</a>&nbsp;"   
response.write "<a href=?page="&Allpage&"&"&C_ID&">尾页</a>&nbsp;"     
end if
%></td>
                </tr>
                <tr>
                  <td height="25" width="595" colspan="4"><p align="right">　</td>
                </tr>
            </table></td>
          </tr>
        </form>
      </table></td>
      <td width="4" valign="top" bgcolor="#e6e6e6"></td>
    </tr>
    <tr><%Call CloseRs(rs)%>
      <td width="980" height="26" colspan="3"><!--#include file=footer.asp--></td>
    </tr>
  </table>
  </center>
</div>
</body>
</html>

