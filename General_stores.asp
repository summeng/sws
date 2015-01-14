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
<%Call OpenConn
  public Function RemoveHTML(strHTML) 
  ON ERROR RESUME Next
  Dim objRegExp, strOutput
    Set objRegExp = New Regexp
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    objRegExp.Pattern = "<.+?>"
    strOutput = objRegExp.Replace(strHTML, "")
    strOutput = Replace(strOutput, "<", "〈")
    strOutput = Replace(strOutput, ">", "〉")
    RemoveHTML = strOutput   
    Set objRegExp = Nothing
 End Function  
dim k,m,com,rs3,e
m=CheckStr(trim(request("com")))
set rs3 = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_user] where username='"&m&"'" 
rs3.open sql,conn,1,3
if rs3.eof or rs3.bof then
response.write "<script language=JavaScript>" & chr(13) & "alert('相关用户资料没有通过管理员的审核或者已经被删除！')</script>"
response.redirect "../"
end if
if rs3("sdkg")<>1 then
response.write "<script language='javascript'>"
	response.write "alert('抱歉，本店铺尚未审核或店主暂时关闭！');"
	response.write "location.href='javascript: window.close()';"
	response.write "</script>"
	response.end
end if
if rs3("vip")=1 then
response.redirect "user/com.asp?com="&m&""
end if
rs3("tong")=rs3("tong")+1
rs3.update
%>
<%dim f,rs0
set rs0=server.createobject("adodb.recordset")
set rs0=conn.execute("select count(id) from [DNJY_info] where username='"&m&"'")
f=rs0(0)
rs0.close
set rs0=nothing
%>
 
 <DIV class=h></DIV>
<div align="center">
 <!---通栏广告开始---->  
  <table width="980" align="center" bgcolor="#ffffff" class="td">
  <tr>  
   <td ><script language=javascript src="<%=adjs_path("ads/js","bottom",c1&"_"&c2&"_"&c3)%>"></script></td>
  </tr>
  </table>
  <!---通栏广告结束---->
<table border="0" width="980" cellspacing="0" cellpadding="0" class="td" bgcolor="#ffffff">
      
<tr>
<td width="193" valign="top" bgcolor="#F2F6E2">
<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table2" height="550">
<tr>
<td height="14">
<img border="0" src="img/t_01.gif" width="193" height="14"></td>
</tr>
<tr>
<td background="img/t_02.gif" height="91">&nbsp;<font color="#808080">商家模式：</font>普通会员<br>&nbsp;&nbsp;&nbsp;&nbsp; <font color="#808080">店主：</font><%=rs3("name")%>(<%=rs3("username")%>)<br>&nbsp;&nbsp;&nbsp;&nbsp; <font color="#808080">积分：</font><%=rs3("jf")%><br>&nbsp;&nbsp;<font color="#808080"> 物品数：</font><%=f%><br>&nbsp;<font color="#808080">创店时间：</font><%=formatdatetime(rs3("zcdata"),1)%><br>&nbsp;<font color="#808080">商家地址：</font><%=left(rs3("dizhi"),10)%><br>&nbsp;<a href="#" ONCLICK="window.open('user/user_mail.asp?username=<%=rs3("username")%>&email=<%=rs3("email")%>&mailzt=<%=rs3("sdname")%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=650,height=300,left=300,top=20')">给店家发邮件</a><%if rs3("qq")<>"" then %>&nbsp; <a target=blank href=http://wpa.qq.com/msgrd?V=1&Uin=<%=rs3("qq")%>&Menu=yes><img border="0" SRC=http://wpa.qq.com/pa?p=1:<%=rs3("qq")%>:7 alt="用QQ给他发信息"></a><%end if%><br> <a href="messh.asp?name=<%=rs3("username")%>&<%=C_ID%>">发站内短信给他</a><br><a href="user/collection_shops.asp?scid=<%=rs3("username")%>&name=<%=rs3("name")%>&sdname=<%=rs3("sdname")%>&mpname=<%=rs3("mpname")%>" target="_blank">会员收藏</a><br>
<font color="#666666">&nbsp;已有</font>[<%=rs3("tong")%>]<font color="#666666">人访问您的店铺</font></td>
</tr>
<tr>
<td height="17">
<img border="0" src="img/t_03.gif" width="193" height="17"></td>
</tr>
<tr>
<td bgcolor="#FFFFFF">　</td>
</tr>
<tr>
<td height="329">
<img border="0" src="img/t_07.gif" width="193" height="329"></td>
</tr>
<tr>
<td height="64" valign="top"><%if request.cookies("dick")("username")<>"" then%>&nbsp; 
<font color="#808080">地址：<%=rs3("dizhi")%><br>&nbsp; 电话：<%=rs3("dianhua")%><br>&nbsp; 传真：<%=rs3("fax")%><br>&nbsp; E-mail：<%=rs3("email")%></font><%else%><font color="#808080">浏览非<a target="_blank" href="help.asp?id=1&<%=C_ID%>">[VIP会员]店铺</a>需要登陆</font><a href="login.asp?<%=C_ID%>"><font color="#808080">[<%=title%>]</font></a><font color="#808080">才能看到联系方式→</font><a href="login.asp"><font color="#CC9900">[我要登陆]</font></a><%end if%>
</td>
</tr>
</table>
</td>
<td>
<img border="0" src="img/b.gif" width="10" height="1"></td>
<td valign="top">
<div align="center">
<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table3">
<tr>
<td height="50" colspan="3">&nbsp;<span class=font1 style="height: 30; padding-top: 8px"><font size="5" color="#0066CC" face="黑体">[<%=rs3("sdname")%>]</font></span>&nbsp;&nbsp; 
<br>
<marquee onMouseOver="this.stop()" onMouseOut="this.start()" scrollamount="3" scrolldelay="121" behavior="scroll"  height="14" style="font-size: 10pt; color: #835E07" width="546">最新公告：<%if IsNull(rs3("comgg")) then%>欢迎光临<%=rs3("sdname")%>主页<%else%><%=rs3("comgg")%><%End if%></marquee></td>
</tr>
<tr>
<td width="12" class="tdcom3">
<img border="0" src="img/t_08.gif" width="18" height="18"></td>
<td width="11%" class="tdcom3">
&nbsp;关于我们</td>
<td width="86%" class="tdcom1">
　</td>
</tr>
<tr>
<td colspan="3" width="760">
<div align="center">

<table><tr><td ><div class="pic">
<TABLE style="BORDER-COLLAPSE: collapse"   borderColor=#999999 cellSpacing=0 cellPadding=4 align=right border=1>
<TBODY>
<TR>
<TD align=middle><p align="center"><%if rs3("logo")<>"" then %><img border="0" src="<%=rs3("logo")%>" title="店标"  width="150" height="120"><%else%><img border="0" src="images/nopic.gif" title="店标"  width="120" height="90"><%end if%><br>
<a href="javascript:window.external.AddFavorite('http://<%=BaseUrl%>General_stores.asp?com=<%=rs3("username")%>&<%=C_ID%>', '<%=rs3("sdname")%>-<%=title%>')"><img border="0" src="img/bookmark.gif"><font color="#835E07">收藏此店铺</font></a><br></TD></TR></TBODY></TABLE>

</div><u><b><font color="#808080"><CENTER><%=rs3("sdname")%>简介</font></b></u></CENTER><p><%=HTMLDecode(rs3("sdjj"))%><p>
&nbsp;&nbsp;&nbsp;&nbsp;<b>行业：</b><%belongtype="<a href=""sdlist.asp?t="&rs3("type_oneid")&",0,0"">"&rs3("type_one")&"</a>"
	IF rs3("type_two")<>"" and not isnull(rs3("type_two")) Then belongtype=belongtype&" / <a href=""sdlist.asp?t="&rs3("type_oneid")&","&rs3("type_twoid")&",0"">"&rs3("type_two")&"</a>"
	IF rs3("type_three")<>"" and not isnull(rs3("type_three")) Then belongtype=belongtype&" / <a href=""sdlist.asp?t="&rs3("type_oneid")&","&rs3("type_twoid")&","&rs3("type_threeid")&""">"&rs3("type_three")&"</a>"
	response.write belongtype%><p>
&nbsp;&nbsp;&nbsp;&nbsp;<b>主营：</b><%=rs3("zhuying")%><p>
</td>
</tr></table>
</div>
</td>

</tr>
</table>
<div align="center">
<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table4">
<tr>
<td colspan="3">
　</td>
</tr>
<tr>
<td width="35" class="tdcom3">
<img border="0" src="img/t_08.gif" width="18" height="18"></td>
<td width="11%" class="tdcom3">
<span class=font1>&nbsp;新品展示</span></td>
<td width="86%" class="tdcom1">
　</td>
</tr>
<tr>
<td width="560" colspan="3">
<div align="center">
	<TABLE border=0 width="100%">
<tr>
<% 
dim j,rstj_a,sqltj_a
set rstj_a=server.createobject("adodb.recordset")
sqltj_a="select top "&huiyuansp&" * from DNJY_info where yz=1  and username='"&m&"' order by fbsj "&DN_OrderType&",ID "&DN_OrderType&"" 	      
rstj_a.open sqltj_a,conn,1,1        
while not rstj_a.eof      
%>
<%for j=0 to 4                                      
  if rstj_a.eof then                                      
  exit for                                      
  end if                                      
  %> 
<td align="center" class=tmpstr valign="top" style="; border:1px solid #CCCCCC; background-color:#F7F7F7" width="25%">
<a title="<%=rstj_a("biaoti")%>" target="_blank" href="<%=x_path("",rstj_a("id"),"",rstj_a("Readinfo"))%>"><%if rstj_a("tupian")<>"0" and rstj_a("tupian")<>"" then%><IMG src="UploadFiles/infopic/<%=rstj_a("tupian")%>" width="130" height="110" border=1 style="border: 1px solid #FFFFFF; ; padding-left:2px; padding-right:2px" hspace="0" ><br><u><b><%=CutStr(rstj_a("biaoti"),16,"")%></b></u><font size="1">></font></a><%else%><img src="images/nophoto2.gif" alt="无图片" width="130" height="110" border="0"><br><u><b><%=CutStr(rstj_a("biaoti"),16,"")%></b></u><font size="1">></font><%end if%></td>
<%     
rstj_a.movenext                          
next  
%> </tr>
<%wend          
rstj_a.close
set rstj_a=nothing%> 
</TABLE></div>
  <table width="760" align="center" bgcolor="#ffffff" class="td">
  <tr>  
   <td ><script language=javascript src="<%=adjs_path("ads/js","xxfl1",c1&"_"&c2&"_"&c3)%>"></script></td>
  </tr>
  </table>
</td>
</tr>
<tr>
<td width="560" colspan="3">　</td>
</tr>
<tr>
<td width="560" colspan="3">　</td>
</tr>
</table>
</div>
</div>
</td>
</tr>
</table><%conn.execute "UPDATE DNJY_user SET sdhits=sdhits+1 WHERE username='"&m&"'" '点击计数%>
</div><%Call CloseRs(rs3)
Call CloseDB(conn)%>
<!--#include file=footer.asp-->
</body>

</html>
