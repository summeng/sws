<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="top.asp"-->


<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<script type="text/javascript" src="admin/inc/js_left.js"></script>
<link rel="stylesheet" href="css/news.css" type="text/css" />
</head>
<%Dim oyaya_keywords,oyaya_description,author,page,ppage,pages,j,sqlt,rst,classid,rc

n=request("n")
If n="" Then n="0,0,0"
n=split(n,",")
If n(0)="" Then n(0)=0
n1=strint(n(0))
If ubound(n)<1 Then
n2=0
else
n2=strint(n(1))
End if
If ubound(n)<2 Then
n3=0
else
n3=strint(t(2))
End If

IF n1=0 then
nn=""
elseIF n3>0 then
nn=" and c_oneid="&n1&" and c_twoid="&n2&" and c_threeid="&n3
elseIF n2>0 then
nn=" and c_oneid="&n1&" and c_twoid="&n2
Else
nn=" and c_oneid="&n1
End If

i=clng(request("i"))
%>
<table width="980" height="100%" border="0" align="center" cellpadding="0" cellspacing="0" class=table-zuoyou bgcolor="#FFFFFF">
  <tr> 
    <td width="218" align="center" valign="top">
	<table width="100%" border="0" cellpadding="0" cellspacing="0"  class="ty1">
  <tr>
    <td height="30" align="center"class="ty3"><b>资讯栏目</b></td>
  </tr></table>
 <table width="100%" border="0" cellpadding="0" cellspacing="0"  class="ty1">
        <tr>
          <td>
     <%=news_class(1,1,1,0,0,0,1,12,11,9,15,13,10,"news.asp")%></td>
        </tr>
      </table>
	        <table width="220" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
	<!--#include file="news_left.asp"-->
	</td>

<td height="100%" align="center" valign="top">
<table width="750" border="0" cellpadding="0" cellspacing="0" class="ty1">
<%
page=clng(request.querystring("page"))
If request("ppage")<>"" Then page=clng(request("ppage"))
Call OpenConn
set rs=server.createobject("adodb.recordset")
sql="select * from DNJY_News where ifshow=1"&nn&""&ttccnews&" order by newstop "&DN_OrderType&",id "&DN_OrderType&""

rs.open sql,conn,1,1
if rs.eof and rs.bof then
Response.Write "<p><br>暂无"
IF c1>0 Then Response.Write conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
Response.Write "内容!"
else 
rs.pagesize=30
if page=0 then page=1 
pages=rs.pagecount
if page > pages then page=pages
rs.absolutepage=page  
for j=1 to rs.pagesize 
%>
        <tr> 
          <td height="24" align="center" ><%if rs("newstop")=1 then%><img src="img/ding_ico.gif" width="16" height="16" alt="置顶文章"><%else%><img src="img/pu_ico.gif" width="16" height="16" alt="普通文章"><%end if%></td>
          <td height="24" colspan="2" align="left" style="border-bottom: #999999 1px dotted"> <% If rs("img")=True then response.write "<img src='img/pic.gif' border=0 alt='有相册'>" end if %>
            <a class=<%=rs("oColor")%> href="news_show.asp?id=<%=rs("id")%>&n=<%=n1%>,<%=n2%>,<%=n3%>&<%=C_ID%>" target="_blank"><span class="<%=rs("oStyle")%>"><%= rs("title") %></span></a> <%if rs("newstop")=1 then%><img src='img/top4.gif' border=0 alt='置顶文章'><%end if%>&nbsp;<%if rs("tuijian")=1 then%><img src='img/jian.gif' border=0 alt='推荐文章'><%end if%> <font color="#999999"> (阅读<font color="#ff0000"><%= rs("hits") %></font>次)&nbsp; [<%if rs("infotime")=date() then%><font color="#ff0000"><%else%><font color="#999999"><%end if%><%= dicksj2(rs("infotime")) %></font>] </font>    
          </td>
        </tr>
        <%
rs.movenext
if rs.eof then exit for
next
response.write"<tr valign='bottom'><td height='100%' colspan='3' align='center'>" 
   response.write"<form method=post action='news.asp?n="&n1&","&n2&","&n3&"&page="&page&"&"&C_ID&"'>"  
      if page<2 then      
    response.write "首页 上一页&nbsp;"
  else
    response.write "<a href=news.asp?n="&n1&","&n2&","&n3&"&page=1&"&C_ID&">首页</a>&nbsp;"
    response.write "<a href=news.asp?n="&n1&","&n2&","&n3&"&page=" & page-1 & "&"&C_ID&">上一页</a>&nbsp;"
  end if
  if rs.pagecount-page<1 then
    response.write "下一页 尾页"
  else
    response.write "<a href=news.asp?n="&n1&","&n2&","&n3&"&page=" & (page+1) & "&"&C_ID&">"
    response.write "下一页</a> <a href=news.asp?n="&n1&","&n2&","&n3&"&page="&rs.pagecount&"&"&C_ID&">尾页</a>"
  end if
   response.write "&nbsp;页次：<strong><font color=red>"&page&"</font>/"&rs.pagecount&"</strong>页 "
    response.write "&nbsp;共<b><font color='#ff0000'>"&rs.recordcount&"</font></b>条记录 <b>"&rs.pagesize&"</b>条记录/页"
   response.write " 转到：<input type='text' name='ppage' size=4 maxlength=10 class=input value="&page&">"
   response.write " <input class=input type='submit'  value=' goto '  name='cndok'></span></p>"     
response.write"</form></td></tr>"
end if
Call CloseRs(rs)
%>

      </table>
 <!--JS动态地区调用方式（无JS文件用总站广告代替，只能动态页面用）代码开始-->
  <script src="<%=adjs_path("ads/js","tail",c1&"_"&c2&"_"&c3)%>"></script>
  <!--广告代码结束-->
    </td>
  </tr>
</table>
<!--#include file=footer.asp-->
</body>
</html>


                                                              
                                                              
                                                              