<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%set rs=server.createobject("ADODB.recordset")
dim key
leixing=request("leixing")
key=request("key")
 %>
  <!--#include file=top.asp-->
<LINK href="/<%=strInstallDir%>css/style.css" type=text/css rel=stylesheet>
<SCRIPT language=JavaScript type=text/JavaScript>
<!--
function MM_swapImgRestore() { //v3.0
  vari,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</SCRIPT>
<META content="Microsoft FrontPage 5.0" name=GENERATOR>

<style type="text/css">
<!--
.style1 {
	color: #555555;
	font-weight: bold;
}
-->
</style>
</HEAD>
<BODY leftMargin=0 topMargin=0>
<table style="margin-bottom:8px;" width="980" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="CCCCCC">
  <tr>
    <td height="25" bgcolor="FFFAE6" style="padding-left:6px;">
   <span><A href="sdlist.asp?<%=C_ID%>"><b>全部行业</b></A></FONT></span>
<%=V_vassal_c_t(0,c1,c2,c3,0,0,10)%>       
</td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF" style="padding:8px;"><%=W_vassal_c_t(0,1,c1,c2,c3,t1,t2,t3,0,1,0,8,1,22)%></td>
  </tr>
</table>
<table style="margin-bottom:8px;" width="980" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="CCCCCC">
  <tr>
    <td height="25" bgcolor="FFFAE6" style="padding-left:6px;"> 地区导航：<a href="/<%=strInstallDir%>">总站-></a></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF" style="padding:8px;"><%=W_vassal_c_t(0,1,c1,c2,c3,t1,t2,t3,0,1,0,10,0,22)%></td>
  </tr>
</table>
<table width="980" height="250" border="0" align="center" cellpadding="0" cellspacing="0">

  <tr>
    <td width="214" height="62" valign="top"">
  	<table width="100%" border="0" cellspacing="0" cellpadding="0" class="dragTable">
        <tr> 
          <td class="head"> 
            <h3 class="L"></h3>
            <span class="TAG">商家店铺分类</span> 
            <h3 class="R"></h3>
          </td>
        </tr>
        <tr> 
          <td class="middle" align="left"><%=m_more(0,0,1,1,0,0,1,0,1,11,9,10,8)%></td>
        </tr>
        <tr> 
          <td class="foot"> 
            <h3 class="L"></h3>
            <h3 class="R"></h3>
          </td>
        </tr>
      </table></td>
    <td valign="top" bgcolor="#FFFFFF"><table width="100%" height="4" border="0" cellpadding="0" cellspacing="0" bgcolor="#E6E6E6">
      <tr>
        <td></td>
      </tr>
    </table>
	<!---上部通栏广告开始---->
  <table align="center" bgcolor="#ffffff">
  <tr>  
   <td ><script language=javascript src="<%=adjs_path("ads/js","fygg",c1&"_"&c2&"_"&c3)%>"></script></td>
  </tr>
  </table>
  <!---上部通栏广告结束----> 
<table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr align="center" bgcolor="#FAFAFA">
        
        <td height="27" bgcolor="#FAFAFA"><strong>店铺名称</strong></td>
        <td width="15%" bgcolor="#FAFAFA"><strong>店主</strong></td>
        <td width="15%" bgcolor="#FAFAFA"><strong>站内短信</strong></td>
        <td width="16%" height="27" bgcolor="#FAFAFA"><span class="style1">店铺审核时间</span></td>
        </tr>
      <tr align="center" bgcolor="#FAFAFA">
        <td width="357" height="1" colspan="2" bgcolor="#CCCCCC"></td>
        <td width="16%" height="1" bgcolor="#CCCCCC"></td>
        </tr>
       <%dim page
		set rs=server.createobject("adodb.recordset")
		sql="select * from [DNJY_user] where sdkg=1 and sdname<>''"&tttt&""&ttccdp&" order by id "&DN_OrderType&""
		rs.open sql,conn,3,3
		rs.pagesize=10 '每页显示记录数
		page=cint(request.querystring("page"))
		if page<1 then
		 page=1
		elseif page>rs.recordcount then
		 page=recordcount
		end if
		if not rs.eof then rs.absolutepage=page
		if rs.eof Then
		Response.Write "<tr><td colspan=7 bgcolor=#ffffff>&middot;<font color=#ff0000>暂无"
        IF c1>0 Then Response.Write " "&conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)&""
		IF t1>0 Then Response.Write " "&conn.Execute("Select name From DNJY_hy_type Where id="&t1&" and twoid="&t2&" and threeid="&t3&"")(0)&"类 "
        Response.Write "店铺!</font></td></tr>"	
		End if
		for i=1 to rs.pagesize		
		if rs.eof then exit for
		 %>
              <tr bgcolor="f5f5f5">
              
                <td height="27" background="img/22.gif"><%dim rs1,sql1,id_1
set rs1=server.createobject("adodb.recordset")
sql1 = "select * from [DNJY_user] where sdname order by id "&DN_OrderType&""
rs1.open sql1,conn,1,1
id_1=rs1("id")
On Error Resume Next
Call CloseRs(rs1)
vip=rs("vip")
%>
[<font color="#FF0000">
<%
leixing=rs("vip")
if leixing=0 then
response.write "<font color=""#0000FF"">普通用户</font>"
else
response.write "<font color=""#ff0000"">VIP用户</font>"

end if%>
</font>]<img src="images/dd.gif"><a  href="user/com.asp?com=<%=rs("username")%>" target="_blank"><font size=3  face="黑体" color="red"><u><%=mid(rs("sdname"),1,15)%></u></font></a> <%=rs("jf")%>' <%if rs("vip")=1 then
response.write "<img src=""images/jian.gif"" alt=""特约商家"">"
end if%></td>
                <td height="27" background="img/22.gif"><div align="center"><%=rs("name")%></div></td>
                <td height="27" background="img/22.gif"><div align="center"><span  style="text-align: right; width: 100%"><INPUT class="inputb" TYPE=button VALUE="发站内短信" ONCLICK="window.open('messh.asp?name=<%=rs("username")%>','Sample','toolbar=yes,location=yes,directories=yes,status=yes,menubar=yes,scrollbars=yes,resizable=yes,copyhistory=no')"></span> </div></td>
                <td height="27" align="center" background="img/22.gif" bgcolor="f5f5f5"><%=dicksj2(rs("dpdata"))%></td>
              </tr>
        <tr bgcolor="f5f5f5">
        <td colspan="3" valign="top" bgcolor="#FFFFFF"><table width="98%" border="0" align="right" cellpadding="0" cellspacing="0">
          <tr>
<%response.write "<td align=""middle"" width=120 height=120 rowspan=2>"
if right(rs("logo"),4)=".gif" or right(rs("logo"),5)=".jpeg" or right(rs("logo"),4)=".jpg"then
response.write "<a target=""_blank"" href=""user/com.asp?com="&rs("username")&"""><img src="""&rs("logo")&""" align=""middle"" border=0 width=120 height=90 alt=""店标""></a>"
else
response.write "<a target=""_blank"" href=""user/com.asp?com="&rs("username")&"""><img src=""images/nopic.gif"" align=""middle"" border=0 alt=""没有店标""></a>"
end if
response.write "</td>"%><td>&nbsp;&nbsp;</td>
            <td height="66" valign="top"><span class="style1">店铺简介</span>：<%if len(rs("sdjj"))<180 then%><%=rs("sdjj")%><%else%><%=mid(rs("sdjj"),1,180)%>...<%end if%><br>
<%belongtype="<a href=""sdlist.asp?t="&rs("type_oneid")&",0,0&"&C_ID&""">"&rs("type_one")&"</a>"
	IF rs("type_two")<>"" and not isnull(rs("type_two")) Then belongtype=belongtype&" / <a href=""sdlist.asp?t="&rs("type_oneid")&","&rs("type_twoid")&",0&"&C_ID&""">"&rs("type_two")&"</a>"
	IF rs("type_three")<>"" and not isnull(rs("type_three")) Then belongtype=belongtype&" / <a href=""sdlist.asp?t="&rs("type_oneid")&","&rs("type_twoid")&","&rs("type_threeid")&"&"&C_ID&""">"&rs("type_three")&"</a>"	
response.write "行业类别："&belongtype&"<br>"%>
            
<%response.write "联系电话：<b>"&rs("dianhua")&"</b><br>联 系 人：<b>"&rs("name")&"</b><br>店铺地址：<b>"&rs("dizhi")&"</b>"%>
            </td>
          </tr>
        </table></td>
        </tr>
      <% 
		rs.movenext
		next
        %>
    </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="768"><table width="100%" height="32" border="0" cellpadding="0">
            <tr>
              <td width="308"><strong>&nbsp;共有<font color=#ff0000><%= rs.recordcount %></font>条记录  每页<font color=#ff0000><%= rs.pagesize %></font>条记录 分<font color=#ff0000><%= rs.pagecount %></font>页显示/当前第<font color=#ff0000><%= page %></font>页</strong></td>
              <td align="right" width="250"><strong>分页： <% if Page<2 then     
response.write "<font color=""#808080"">首页</font>&nbsp;"
response.write "<font color=""#808080"">上一页</font>&nbsp;"     
else %><a href="sdlist.asp?<%=C_ID%>">首页</a> <a href="sdlist.asp?page=<%=page-1%>&<%=C_ID%>">上一页</a><%end if%><%if rs.pagecount-Page<1 then     
response.write "<font color=""#808080"">下一页</font>&nbsp;"
response.write "<font color=""#808080"">尾页</font>&nbsp;"  
else %> <a href="sdlist.asp?page=<%= page+1 %>&<%=C_ID%>">下一页</a> <a href="sdlist.asp?page=<%= rs.pagecount %>&<%=C_ID%>">尾页</a><%end if%>&nbsp;</strong></td>
            </tr>
          </table></td>
        </tr>
      </table>
    <% rs.close
	set rs=Nothing%></td>
    <td width="4" valign="top" bgcolor="#E6E6E6"></td>
  </tr>
</table>

<TABLE cellSpacing=0 cellPadding=0 width=980 align=center border=0>
  <TBODY>
  <TR>
    <TD width="769">
    <p align="center"><!--#include file=footer.asp--></TD></TR>                         
  <TR>
<TD width="769"></TD></TR></TBODY></TABLE>
</BODY></HTML>