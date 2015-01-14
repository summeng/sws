<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="Include/ContentAutoPage.asp"-->
<!--#include file="top.asp"-->
 <!--#include file="Include/err.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<script type="text/javascript" src="Include/onews.js"></script>
<script type="text/javascript" src="admin/inc/js_left.js"></script>
<link rel="stylesheet" href="css/news.css" type="text/css" />
<script language="javascript" src="Include/news_jquery-1.4.2.min.js"></script>
<script type="text/javascript" src="Include/news_jquery.ad-gallery.js"></script>
<style type="text/css">
<!--
img {width:expression(this.width > 730 ? "730px" : this.width);overflow:hidden;max-width: 730px}
-->
</style>
</head>
<%Call OpenConn
'＝＝＝＝＝保存评论＝＝＝＝＝＝＝＝
Dim action,pinglunname,pingluncontent
if request("action")="save" then
id=clng(request.querystring("id"))
	if id="" then
	response.write "<script language='javascript'>alert('对不起，参数不正确！');history.back();</script>" 
	Call CloseDB(conn)
	response.end
end If

if trim(request.form("validate_codee"))="" then
	response.write "<script language='javascript'>alert('请填验证算式！');history.back();</script>" 
	Call CloseDB(conn)
	response.end
end if 

if CInt(trim(request.form("validate_codee")))<>CInt(Session("ttpSN")) then
call errdick(45)
Session("ttpSN")=""
response.end
end if

if len(request.form("pingluncontent"))>100 then
	   response.write"<script>alert('对不起，评论的内容不能超过100个字符!请重新输入');history.back();</script>"
	   response.end
end If

action=request.querystring("action")
if action="save" Then   
	pinglunname=Replace(Request.Form("pinglunname"),"'","＇") 
	pingluncontent=Replace(Request.Form("pingluncontent"),"'","＇") 
	set rs=server.createobject("adodb.recordset")
	rs.open "select * from DNJY_news_pinglun",conn,1,3
	rs.addnew
	rs("id")=id
	rs("pinglunname")=replace(replace(replace(replace(pinglunname,"'","‘"),"<","&lt;"),">","&gt;")," ","&nbsp;")
	rs("pingluncontent")=replace(replace(replace(replace(pingluncontent,"'","‘"),"<","&lt;"),">","&gt;")," ","&nbsp;")
	If plsh=0 Then rs("sh")=1
	rs.update
Call CloseRs(rs)
If plsh=0 Then
	response.Write "<script language=javascript>alert('恭喜您，您的评论已成功提交。');location.replace('news_show.asp?id="&owen&"');</script>"
	response.end
else
	response.Write "<script language=javascript>alert('恭喜您，您的评论已成功提交。需审核才能显示!');location.replace('news_show.asp?id="&owen&"');</script>"
   response.end
end If
end if
end If
'＝＝＝＝＝＝＝＝＝＝保存评论END＝＝＝＝＝＝＝

Dim owen,rsnews,sql1,rs2,sql2,titleonews,Bimg,Strsimg
randomize timer
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

owen=clng(request.querystring("id"))
if request.querystring("id")="" then
	response.write "<script language='javascript'>alert('对不起，参数不正确！');history.back();</script>" 
	Call CloseDB(conn)
	response.end
end If%>
<body>
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0">
<TR>
<TD width="760" valign="top">
   <table width="760"  cellpadding="1" cellspacing="1" class="tablest">
	<TR>
		<TD>
<%titleonews=title
set rsnews=server.createobject("adodb.recordset")
sql="select * from DNJY_news where ifshow="&DN_True&" and id="&owen&""
rsnews.open sql,conn,1,1
if rsnews.eof then
Response.Write "<p><br><CENTER>暂无"
IF c1>0 Then Response.Write conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
Response.Write "内容!"
rsnews.close
set rsnews=Nothing
response.end
Else
sql="update DNJY_news set hits=hits+1 where id="&owen
conn.execute sql
if rsnews("img")=True And rsnews("Photo")<>"" then'有相关相册时
Bimg=Mid(rsnews("Photo"),2)
end If
%>
        <tr>
         <td align="center" class="inc_name" colspan="2"><div style='line-height: 300%;width: 100%; font-size:20pt; font-family: 黑体; color:#86CAED; filter:shadow(color=#000000,direction=240)'><b><%= rsnews("title") %></b></div></td>
        </tr>
        <tr> 
          <td width="10%" height="30" align="center" style="border-bottom: 1 solid #666666">文章来源：<%= rsnews("come") %>  &nbsp; 发布者：<%= rsnews("newsuser") %> 
          &nbsp; 发布时间：<%= rsnews("infotime") %> &nbsp; 阅读：<font color="#ff0000"><%= rsnews("hits") %></font>&nbsp;&nbsp;次双击滚屏</td>
        </tr>
<tr><td align="center" height="20">&nbsp;</td></tr>
<%If Bimg<>"" Then%>
<tr><td align='center'> <div id='gallery' class='ad-gallery'>   <div class='ad-image-wrapper'>   </div>   <div class='ad-controls' id=a1></div>  <div class='ad-nav'>    <div class='ad-thumbs'> <ul class='ad-thumb-list'>
<%If Bimg<>"" Then
Bimg=split(Bimg,"|")
For Each Strsimg In Bimg%>
<li><a href='<%=Strsimg%>'><img src='<%=Strsimg%>' width='90' height='70'></a></li>
<%Next
End if%>
</ul></div> </div></div></td></tr>
<tr><td align="center" height="20"><hr>&nbsp;</td></tr>
<%End if%>
<tr><td>
<div style="margin:10px;font-size:13;letter-spacing:8px;line-height:23px"><%=ContentPagination(rsnews("content"))%></div>  
</td></tr>
<%end If%>
</TD>
	</TR>
   </TABLE>
<TABLE width="750">
<TR>
	<TD width="750"><script src="<%=adjs_path("ads/js","fygg",c1&"_"&c2&"_"&c3)%>"></script></TD>
</TR>
</TABLE>

</TD>

<TD width="3"></TD>

<TD width="215" valign="top">
  <table width="215" border="0" cellpadding="0" cellspacing="0" class="ty1">
  <tr>
    <td align="center" height="30" class="ty3"><b><A href="news.asp?n=0,0,0&<%=C_ID%>">文章栏目</a></b></td>
  </tr>
  <tr>
    <td colspan="2" align="center">
<%=news_class(1,1,1,0,0,0,1,12,11,9,15,13,10,"news.asp")%>
</td>
  </tr>     
</table>
      <table width="215" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
  <table width="215" border="0" cellpadding="0" cellspacing="0" class="ty1">
    <tr>
    <td align="center" height="30" class="ty3"><b>推荐资讯</b></td>
  </tr>

  <tr>
    <td colspan="2" align="center"><%=news(c1,c2,c3,1,0,0,0,1,0,0,0,0,0,8,1,24,13,20)%></td>
  </tr>     
</table>
      <table width="215" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
<table width="215" border="0" cellpadding="0" cellspacing="0"  class="ty1">
  <tr>
    <td align="center" height="30" class="ty3"><b>热门点击</b></td>
  </tr>
  <tr>
    <td height="22" align="center"><%=news(c1,c2,c3,3,0,0,0,1,0,0,0,0,0,8,1,24,13,20)%></td>
  </tr>
</table>
</TD>

</TR>
</TABLE>

<table width="980" border="0" align="center" cellpadding="0" cellspacing="0" >
        <tr>
          <td colspan="2" align="right"> <img src="img/printer.gif" width="16" height="14" align="absmiddle"> <a href="javascript:window.print()">打印本页</a> | <img src="img/close.gif" width="14" height="14" align="absmiddle"> <a href="javascript:window.close()">关闭窗口</a> | <a href="javascript:window.open('http://shuqian.qq.com/post?from=3&title='+encodeURIComponent(document.title)+'&uri='+encodeURIComponent(document.location.href)+'&jumpback=2&noui=1','favit','width=930,height=470,left=50,top=50,toolbar=no,menubar=no,location=no,scrollbars=yes,status=yes,resizable=yes');void(0)" style="background:url('http://shuqian.qq.com/img/add.gif') no-repeat 0px 0px;height:23px;width:60px;padding:2px 2px 0px 20px;font-size:12px;">QQ书签</a></td>
        </tr>
        <hr width="980">
        <tr>
<td width="49%" height="25">
<%Dim rs1
set rs1=Server.CreateObject("ADODB.Recordset")
sql1="select top 1 * from DNJY_news where ifshow=1 and id<"&owen&""&nn&" order by id"&DN_OrderType&""
rs1.open sql1,conn,1,1
if rs1.eof and rs1.bof then
%>
<div align="left">&nbsp;上一篇文章：暂时没有</div>
<%else%>
	 <div align="left">&nbsp;上一篇文章：<a title="<%=rs1("title")%>" href="news_show.asp?id=<%=rs1("id")%>&n=<%=n1%>,<%=n2%>,<%=n3%>&<%=C_ID%>"><%=rs1("title")%></a></div>
<%end If
		rs1.close
		set rs1=nothing%>
		  </td>
          <td width="49%">
<%
set rs2=Server.CreateObject("ADODB.Recordset")
sql2="select top 1 * from DNJY_news where ifshow=1 and id>"&owen&""&nn&""
rs2.open sql2,conn,1,1
if rs2.eof and rs2.bof then
%>
<div align="left">&nbsp;&nbsp;下一篇文章：暂时没有</div>
<%else%>
	 <div align="left">&nbsp;&nbsp;下一篇文章：<a title="<%=rs2("title")%>" href="news_show.asp?id=<%=rs2("id")%>&n=<%=n1%>,<%=n2%>,<%=n3%>&<%=C_ID%>"><%=rs2("title")%></a></div>
 <%end If
 Call CloseRs(rs2)%>
		  </td>
        </tr>
</table>

<%If newspl<3 then%>	  
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0" >

	    <tr>
          <td height="25" colspan="4" align="center"><b><font color="#FF0000">您的评论将被网络上成千上万的读者所共享，请对自己言论负责。</font></b></td>
        </tr>
        <tr>
          <td width="500">
		  <table  border="0" cellpadding="2" cellspacing="1" bgcolor="#ffffff">
<%If newspl=0 Or (newspl=1 And request.cookies("dick")("username")<>"") then%>
		  <form name="pinglunform" method="post" action="?id=<%=owen%>&action=save">
            <tr bgcolor="#f5f5f5">
              <td width="150">您的姓名：</td>
              <td><%If request.cookies("dick")("username")="" then%><input name="pinglunname" type="text" id="pinglunname" size="10" maxlength="8"><%else%><%=request.cookies("dick")("username")%><input type="hidden" name="pinglunname" value="<%=request.cookies("dick")("username")%>"><%End if%></td>
            </tr>
            <tr bgcolor="#f5f5f5">
              <td>评论正文：</td>
              <td><textarea name="pingluncontent" cols="40" rows="4" id="pingluncontent"></textarea></td>
            </tr>

<tr bgcolor="#f5f5f5">                                                
  <td height="25"> 算式验证：</td>
      <td height="25" width="408"><img src="tt_getcodee.asp" alt="验证码,看不清楚?请点击刷新验证码" height="10" style="cursor : pointer;" onClick="this.src='tt_getcodee.asp?t='+(new Date().getTime());" /><input name="validate_codee" type="text" size="4" maxlength="4"  onKeyUp="if(isNaN(value)){alert('只允许输入数字');value='';}"   onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d|^\.]/g,''))"> &nbsp;&nbsp; <img src=images/memo.gif alt="验证码,看不清楚?请点击刷新验证码"></td>
  </tr>  
        
            <tr bgcolor="#f5f5f5">
              <td colspan="3"><div align="center"><%If plsh=1 then%><font color="#FF0000">发表评论需审核才能显示</font><%End if%>&nbsp;&nbsp;
                  <input type="submit" name="submit" value="提 交" onClick="return check();">
                &nbsp;
                <input type="reset" name="submit2" value="清 除">
              </div></td>
            </tr></form>
<%ElseIf (request.cookies("dick")("username")="" or request.cookies("dick")("domain")="" or request.cookies("dick")("id")="") And newspl=1 Then
response.write "要<a href=""login.asp?"&C_ID&""" target=_blank><font color=""#0000FF"">登录</font></a>才能评论"
ElseIf newspl=2 Then
response.write "系统已禁止新评论"
End if%>
          </table>	  
		  </td>
          <td width="55%" align="center">
		  <table height="155" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF">
    <%Dim pages
		set rs1=server.CreateObject("adodb.recordset") 
		rs1.open "select top 5 * from DNJY_news_pinglun where sh=1 and id="&owen&"",conn,1,1
		pages=rs1.recordcount
	%>
		    <tr>
              <td width="7%" height="23" align="center" bgcolor="#F3F3F3"><img src="img/review.gif" align="absmiddle" /></td>
              <td width="14%" align="left" bgcolor="#F3F3F3">评论人</td>
              <td width="60%" align="left" bgcolor="#F3F3F3">评论内容摘要(共 <%=pages%> 楼) <a href="user/Commentnews.asp?newsid=<%=owen%>&title=<%=rsnews("title")%>" target=blank>查看完整内容</a></td>
              <td width="20%" align="left" bgcolor="#F3F3F3">评论时间</td>
            </tr>
        <%
		if rs1.eof and rs1.bof then 
		response.write""
		else
		do while not rs1.eof  
		%>
            <tr>
              <td height="20" align="center" bgcolor="#F3F3F3"><img src="img/b.gif" width="10" height="10" /></td>
              <td align="left" bgcolor="#F3F3F3"><b><%=rs1("pinglunname")%></b></td>
              <td align="left" bgcolor="#F3F3F3"><%=left(rs1("pingluncontent"),20)%></td>
              <td align="left" bgcolor="#F3F3F3"><%=rs1("pinglundate")%></td>
            </tr>
            <%rs1.movenext
		loop
		rs1.close
		set rs1=nothing
		end if%>
            <tr>
              <td colspan="5" align="center" bgcolor="#F3F3F3"><font color="#FF0000">本站发表读者评论，并不代表我们赞同或者支持读者的观点。我们的立场仅限于传播更多读者感兴趣的信息。</font> </td>
            </tr>
          </table>
		</td>
        </tr><%rsnews.close
		set rsnews=nothing%>
      </table>
<%End if%>
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0" >
<div align="center"><b>本类最新文章</b></div>
<%Dim proCount,intCurPage,k
sql="select top 6 * from DNJY_News where ifshow=1"&nn&" order by id "&DN_OrderType&""
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
	if rs.eof then 
	Response.Write " <div align='center'>暂无新闻!</div>"
	end if%>
<table width="980"  border="0" cellspacing="0" cellpadding="0" id="table1" align=center>
<%IF not rs.eof Then
	proCount=rs.recordcount
	rs.PageSize=30					  '定义显示数目
		if not IsEmpty(request.QueryString("ToPage")) then
	       ToPage=CInt(request.QueryString("ToPage"))
		   if ToPage>rs.PageCount then
					rs.AbsolutePage=rs.PageCount
					intCurPage=rs.PageCount
		   elseif ToPage<=0 then
					rs.AbsolutePage=1
					intCurPage=1
				else
					rs.AbsolutePage=ToPage
					intCurPage=ToPage
				end if
			else
			        rs.AbsolutePage=1
					intCurPage=1
		  end if
	intCurPage=CInt(intCurPage)
k=1
do while Not rs.eof and  k<3
%>
<TR>
<%
 for n=1 to 3
%>
   <td width="300" height="23"><img src='img/news_ico.gif' width=15 height=11 border=0> <%if rs("img")=True Then Response.Write"<img src='images/num/pic.gif' width='13' height='13' border=0 alt='有相册'>"%><a class="<%=rs("oColor")%>" href="news_show.asp?id=<%=rs("id")%>&n=<%=n1%>,<%=n2%>,<%=n3%>&<%=C_ID%>" target="_blank" title="<%=rs("Title")%>"><font class="<%=rs("oStyle")%>"><% = left(rs("title"),14) %></font></a><% if Len(rs("title"))>12 then %>…<% end if %></td>
<%rs.movenext
if rs.eof then exit for
if rs.eof then exit do
Next
%>
 </tr>
<%
k=k+1
Loop
Call CloseRs(rs)
Call CloseDB(conn)
End if%>
</table></td>
  </tr>
</table>
<%title=titleonews%>
<TABLE cellSpacing=0 cellPadding=0 width=980 align=center border=0>
  <TBODY>
  <TR>
    <TD width="769">
    <p align="center"><!--#include file=footer.asp--></TD></TR>                         
  <TR>
<TD width="769"></TD></TR></TBODY></TABLE>
<p>
</body>
 <script type="text/javascript">
  $(function() {
    var galleries = $('.ad-gallery').adGallery();
    $('#switch-effect').change(
      function() {
        galleries[0].settings.effect = $(this).val();
        return false;
      }
    );
    $('#toggle-slideshow').click(
      function() {
        galleries[0].slideshow.toggle();
        return false;
      }
    );
    $('#toggle-description').click(
      function() {
        if(!galleries[0].settings.description_wrapper) {
          galleries[0].settings.description_wrapper = $('#descriptions');
        } else {
          galleries[0].settings.description_wrapper = false;
        }
        return false;
      }
    );
  });
  </script>
</html>
