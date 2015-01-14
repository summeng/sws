<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<%'用户登录＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
sub userlogin()%>
<script language="javascript">
<!--
function userLogin_onsubmit() {
if (document.userLogin.username2.value=="" )
	{
	  alert("请填用户名！")
	  document.userLogin.username2.focus()
	  return false
	 }
if (document.userLogin.password.value=="" )
	{
	  alert("请填密码！")
	  document.userLogin.password.focus()
	  return false
	 }
if (document.userLogin.yzm.value=="" )
	{
	  alert("请填验证码！")
	  document.userLogin.yzm.focus()
	  return false
	 }
  
}
// --></script>
<table  border=0 style="border-collapse: collapse"  cellpadding=5 height="150"><tr><TD>
<% if request.cookies("dick")("username")=""  And session("openid")="" then %>
<form action="login_save.asp?<%=C_ID%>" method="post" name="userLogin" language="javascript" onSubmit="return userLogin_onsubmit()">
登录帐号 <input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" id="username2" maxLength="20" size="15" name="username2"><br>
密&nbsp;&nbsp;&nbsp;&nbsp;码 <input style="BACKGROUND-COLOR: #eeeeee; BORDER-BOTTOM: #FFFFFF 1px solid; BORDER-LEFT: #a2a2a2 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #a2a2a2 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;" id="password2" type="password" maxLength="32" size="15" name="password"><br>
验 证 码 <input type="text" name="yzm" size="4" /><img src="tt_getcode.asp" alt="验证码,看不清楚?请点击刷新验证码" height="18" style="cursor : pointer;" onClick="this.src='tt_getcode.asp?t='+(new Date().getTime());" />&nbsp;&nbsp;<input style="BACKGROUND-COLOR: #CDE4CF; BORDER-BOTTOM: #a2a2a2 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #a2a2a2 1px solid; BORDER-TOP: #ffffff 1px solid; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px;"   type="submit" value="登陆" name="I3"></form>
<TABLE  width="100%" align="center">
<TR><%Dim qq_login
If API_QQEnable=True And (API_SinaEnable=false or API_AlipayEnable=false) Then
qq_login="qq_login"
Else
qq_login="qqSmall"
End if%>
	<TD align="center"><%If API_QQEnable=True then%><%response.write " <a title=""使用QQ号快速登录"" href=""/" & strInstallDir & "api/qq/redirect_to_login.asp""><img align=""absmiddle"" src=""/" & strInstallDir &"api/img/"&qq_login&".png"" border=""0"" align=""absmiddle""/></a>"%><%End if%><%If API_SinaEnable=True then%>&nbsp;<%response.write " <a title=""使用新浪微博帐号登录"" href=""/" & strInstallDir & "api/sina/deal.asp""><img align=""absmiddle"" src=""/" & strInstallDir &"api/img/sinaSmall.png"" border=""0"" align=""absmiddle""/></a>"%><%End if%><%If API_AlipayEnable=True then%>&nbsp;<%response.write " <a title=""使用支付宝帐号登录"" href=""/" & strInstallDir & "api/alipay/alipayapi.asp""><img align=""absmiddle"" src=""/" & strInstallDir &"api/img/alipay_button.gif"" border=""0"" align=""absmiddle""/></a>"%><%End if%></TD>
</TR>
</TABLE>
<BUTTON onclick="window.open('user/RetakePassword_a.asp','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=420,height=175,left=300,top=300');" style="BACKGROUND-COLOR: #CDE4CF; BORDER-BOTTOM: 1px solid #a2a2a2; BORDER-LEFT: 1px solid #ffffff; BORDER-RIGHT: 1px solid #a2a2a2; BORDER-TOP: 1px solid #ffffff; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px; width:100px"><div align="center">忘记密码找回</div></BUTTON>&nbsp;&nbsp;&nbsp;<BUTTON onclick="window.open('<%=reg%>?<%=C_ID%>');" style="BACKGROUND-COLOR: #CDE4CF; BORDER-BOTTOM: 1px solid #a2a2a2; BORDER-LEFT: 1px solid #ffffff; BORDER-RIGHT: 1px solid #a2a2a2; BORDER-TOP: 1px solid #ffffff; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px; width:100px"><div align="center">免费注册会员</div></BUTTON>
<%ElseIf session("openid")<>"" And Session("Access_Token")<>"" And request.cookies("dick")("domain")="" Then
session("binding")="yes"
response.write "<center>欢迎QQ的 <font color=#FF0000> "&Session("nickname")&"</font> 光临本站</center><br>"
response.write "<center>请将会员号与QQ号关联<br>&nbsp;&nbsp;方便以后快捷登录 &nbsp;&nbsp;&nbsp;&nbsp;<BUTTON onclick=""window.open('/"&strInstallDir&"api/qq/redirect_to_login.asp');"" style=""BACKGROUND-COLOR: #CDE4CF; BORDER-BOTTOM: 1px solid #a2a2a2; BORDER-LEFT: 1px solid #ffffff; BORDER-RIGHT: 1px solid #a2a2a2; BORDER-TOP: 1px solid #ffffff; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 19px; width:60px""><div align=""center"">立即关联</div></BUTTON></center><br>"
response.write "<center><a href=""user/userout.asp?"&C_ID&""">退出</a></center><br>"
else %>
                     
 <%Dim vip,sj,jf,hb,zcdata,dlcs,msgrs,messyd,mess,username
username=request.cookies("dick")("username")
Call OpenConn
 set rs=server.createobject("adodb.recordset")
sql = "select * from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,1
if rs.eof Then
Response.cookies("dick")=""
Response.cookies("dick")("username")=""
else
if isnull(rs("vipdq")) Then
else
sj=DateDiff("d",now(),""&rs("vipdq")&"")
end If
vip=rs("vip")
jf=rs("jf")
hb=rs("hb")
zcdata=rs("zcdata")
dlcs=rs("dlcs")
If dlcs=1 Then
messyd=1
Else
messyd=0
End if
mess=rs("mess")
Call CloseRs(rs)
response.write " "&title&"欢迎您：<font color=#FF0000>"&username&"</font><br>"
if vip>0 then
  response.write "您的会员级别：<font color=#FF0000>VIP会员</font>"
if sj>0 then
response.Write "<font color=#FF0000> (资格剩余"&sj&"</b>天)</font><br>"
elseif sj=0 then
response.Write "<font color=""#414141"">(资格今日到期)</font><br>"
elseif sj<0 then
response.Write "<font color=""#FF0000"">(资格已经过期)</font><br>"
end If
else
response.write "您的会员级别：<font color=#FF0000>普通会员</font><br>"
end if%> 
 积分：<font color=#FF0000><%=jf%>分</font>
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rmb_mc%>：<font color=#FF0000><%=hb%>个</font><br>
 <%if username<>"" then
Set msgrs = conn.Execute("select count(*) as cmsg from DNJY_Message where isnew='0' and name='"&username&"'")
if cint(msgrs("cmsg"))>0 Or mess=0 then
response.write "<b><a target=_top style='color=red;' href=messs.asp?"&C_ID&"><img src='img/notice.gif' width=15 height=15 border=0>您有新的消息("&msgrs("cmsg")+messyd&")</a></b>"
response.write "<bgsound src='images/mail.wav' loop='1'>"
end If
msgrs.close
set msgrs=nothing
end If%>
<center><BUTTON onclick="location.href='user.asp?<%=C_ID%>';" style="BACKGROUND-COLOR: #CDE4CF; BORDER-BOTTOM: 1px solid #a2a2a2; BORDER-LEFT: 1px solid #ffffff; BORDER-RIGHT: 1px solid #a2a2a2; BORDER-TOP: 1px solid #ffffff; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 22px; width:80px"><div align="center">会员中心</div></BUTTON> &nbsp;&nbsp;&nbsp;<BUTTON onclick="location.href='user/userout.asp?<%=C_ID%>';" style="BACKGROUND-COLOR: #CDE4CF; BORDER-BOTTOM: 1px solid #a2a2a2; BORDER-LEFT: 1px solid #ffffff; BORDER-RIGHT: 1px solid #a2a2a2; BORDER-TOP: 1px solid #ffffff; COLOR: #333333; FONT-SIZE: 9pt; HEIGHT: 22px; width:80px"><div align="center">安全退出</div></BUTTON></center>
<%
end If
end If %></div>
	  </td>
      </tr>
    </table>
<%
end sub%>

<%'首页幻灯广告＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
sub hd()%>
<%dim rs3
Call OpenConn
set rs3=server.createobject("ADODB.recordset")
Set rs3 = conn.Execute("select * from DNJY_adhd") %>
<a target=_self href="javascript:goUrl()"> 
<span class="f14b" >
<script type="text/javascript">
imgUrl1="<%=rs3("pic1")%>";
imgtext1="<%=rs3("tit1")%>"
imgLink1=escape("<%=rs3("pic1_lnk")%>");
imgUrl2="<%=rs3("pic2")%>";
imgtext2="<%=rs3("tit2")%>"
imgLink2=escape("<%=rs3("pic2_lnk")%>");
imgUrl3="<%=rs3("pic3")%>";
imgtext3="<%=rs3("tit3")%>"
imgLink3=escape("<%=rs3("pic3_lnk")%>");
imgUrl4="<%=rs3("pic4")%>";
imgtext4="<%=rs3("tit4")%>"
imgLink4=escape("<%=rs3("pic4_lnk")%>");
imgUrl5="<%=rs3("pic5")%>";
imgtext5="<%=rs3("tit5")%>"
imgLink5=escape("<%=rs3("pic5_lnk")%>");

 var focus_width=<%=rs3("width")%>
 var focus_height=<%=rs3("height")%>
 var text_height=18
 var swf_height = focus_height+text_height 
 var pics=imgUrl1+"|"+imgUrl2+"|"+imgUrl3+"|"+imgUrl4+"|"+imgUrl5
 var links=imgLink1+"|"+imgLink2+"|"+imgLink3+"|"+imgLink4+"|"+imgLink5
 var texts=imgtext1+"|"+imgtext2+"|"+imgtext3+"|"+imgtext4+"|"+imgtext5 
 document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ focus_width +'" height="'+ swf_height +'">');
 document.write('<param name="allowScriptAccess" value="sameDomain"><param name="movie" value="images/focus.swf"><param name="quality" value="high"><param name="bgcolor" value="#F0F0F0">');
 document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
 document.write('<param name="FlashVars" value="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'">');
 document.write('<embed src="pixviewer.swf" wmode="opaque" FlashVars="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'" menu="false" bgcolor="#F0F0F0" quality="high" width="'+ focus_width +'" height="'+ focus_height +'" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />');  document.write('</object>');
 </script>
</span></a><span id=focustext class=f14b></span>
<%rs3.close
set rs3=Nothing
end sub%>

<%'首页图片标题＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function f_tpbt(c1,c2,c3)
Dim rs,sql
Call OpenConn
set rs=server.createobject("ADODB.recordset")
sql = "select top 1 * from DNJY_info where yz=1 and homeEliteImage<>''"&ttcc&" order by fbsj "&DN_OrderType&""
rs.open sql,conn,1,1
if rs.eof then
response.write "还没生成有图片标题！"
Else
response.write "<IMG src=""img/girl_left.gif""><a target=""_blank"" href="""&x_path("",rs("id"),"",rs("Readinfo"))&""" title="&rs("name")&"发布于"&month(rs("fbsj"))&"月"&day(rs("fbsj"))&"日"&">"
response.write "<IMG src="""&rs("homeEliteImage")&"""></a>"
End if
Call CloseRs(rs)

end Function%>

<%'首页图片信息推荐＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function f_tptj(c1,c2,c3,xxsx,mh,zs)
'c1,c2,c3为地区，xxsx信息属性（1最新、2竞价、3推荐、4热门、5置顶）,每行数量，总数%>
<table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="139">
         <table width="100%" height="3" border="0" cellpadding="0" cellspacing="0">
           <tr>
             <td></td>
           </tr>
         </table>
         <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
<%dim ae,ak,rstp
ae=0
ak=0
Call OpenConn
set rstp = Server.CreateObject("ADODB.RecordSet")
sql="select  * from [DNJY_info] where c=1 and yz=1"&ttcc&""
if xxsx=2 Then
sql=sql&" and hfxx=1 and hfje>0 and dqsj>"&DN_Now&""
end If
if xxsx=3 Then
sql=sql&" and tj=2"
end If
if xxsx=4 Then
sql=sql&" and llcs>="&hotbz&""
end If
if xxsx=5 Then
sql=sql&" and b>0"
end If
sql=sql&" order by b "&DN_OrderType&""
if xxsx=1 Then
sql=sql&",fbsj "&DN_OrderType&""
end If
if xxsx=2 Then
sql=sql&",hfje/fbts "&DN_OrderType&",fbsj "&DN_OrderType&",id "&DN_OrderType&""
end If
if xxsx=3 Then
sql=sql&",fbsj "&DN_OrderType&""
end If
if xxsx=4 Then
sql=sql&",llcs "&DN_OrderType&""
end If
'sql=sql&",fbsj "&DN_OrderType&""

rstp.open sql,conn,3,3
if rstp.eof Then
Response.Write "<p><br>暂无"
IF c1>0 Then Response.Write conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
if xxsx=1 Then Response.Write "最新"
if xxsx=2 Then Response.Write "竞价"
if xxsx=3 Then Response.Write "推荐"
if xxsx=4 Then Response.Write "热门"
if xxsx=5 Then Response.Write "置顶"
Response.Write "图片信息!"
else
do while not rstp.eof
%>
    <td valign="top"><div align="center"> <a title="点击查看信息&gt;&gt;&gt;" href="<%=x_path("",rstp("id"),"",rstp("Readinfo"))%>">
          </a>
        <table width="150" height="120"  border="0" cellpadding="0" cellspacing="0" background="images/tupian.jpg">
          <tr>
            <td width="7" height="110"><div align="center"></div></td>
            <td width="57"><a title="点击查看信息&gt;&gt;&gt;" href="<%=x_path("",rstp("id"),"",rstp("Readinfo"))%>" target="_blank"><img src="UploadFiles/infopic/<%=rstp("tupian")%>" border=1 width=110"  height="100" style="border: 1px solid #000000" alt="<%=rstp("biaoti")%>"></a></td>
          </tr>
        </table>
        <table width="130" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><a href="<%=x_path("",rstp("id"),"",rstp("Readinfo"))%>" target="_blank"><%=CutStr(rstp("biaoti"),16,"..")%></a></td>
          </tr>
        </table>
    </div>      
</td>
    <%ae=ae+1
   ak=ak+1
If ak=mh then
	response.write "</tr>"
	ak=0
End If               
                rstp.movenext
				if ae>=zs then exit do
                Loop
                End if
rstp.close
set rstp=Nothing
                 %>
  </tr>
</table>  </td>
  </tr>
</table>
<%end Function%>


<%'滚动信息＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
sub gtxx()%>
<!---本站滚动竞价信息开始---->
<div style="width: 590px;overflow: hidden; background-color: #Fff;" onmouseover="clearInterval(timer1);" onmouseout="go();"><div style="position: relative;top: 0px;left: 0px;white-space: nowrap;color: #FFF;" id="news"><span id="nbo" style="color: #CFC;">
<%Dim yss
Call OpenConn
set rs=server.createobject("adodb.recordset")
sql = "select top 5 * from DNJY_info where yz=1 and hfxx=1 and hfje>0 and dqsj>"&DN_Now&" order by  hfje/fbts "&DN_OrderType&",b "&DN_OrderType&",fbsj "&DN_OrderType&""
rs.open sql,conn,3,3
if rs.eof Then
Response.Write "<img src=""img/icon1.gif"" border=""0""/><font color=""#000000"">暂无"
IF c1>0 Then Response.Write conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
Response.Write "信息-_-</font>"
else
do while not rs.eof
%>
<a target="_blank" href="<%=x_path("",rs("id"),"",rs("Readinfo"))%>">&nbsp;&nbsp;&nbsp;<img src="img/icon1.gif" border="0"/><%=CutStr(rs("biaoti"),60,"..")%></a>
<%rs.movenext
Loop
End if
Call CloseRs(rs)

set rs=nothing%></span><script language="javascript">document.write(nbo.innerHTML);   //重复一次新闻内容
document.write(nbo.innerHTML);   //重复一次新闻内容
document.write(nbo.innerHTML);   //重复一次新闻内容
document.write(nbo.innerHTML);   //重复一次新闻内容
</script></div></div><script language=javascript>function newsScroll() {   //实现不间断滚动
news.style.pixelLeft=(news.style.pixelLeft-1)%nbo.offsetWidth;}function go() { timer1=setInterval('newsScroll()',1)   //更改第二个参数可以改变速度，值越小，速度越快。
}go();</script>
<!---本站滚动竞价信息结束---->
<%end sub%>


<%'首页方框显示信息＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function fkss(c1,c2,c3,btc,lx,lg,gd,sj,tp,hits,hy,jj,btw,nyw,ls,zs,fw)
'c1c2c3为地区，btc标题颜色，lx信息类型，lg信息LOGO,gd固顶，sj时间，tp图片标志，hits点击量，hy会员类型，jj显示竞价，btw标题长度，nyw简介长度，ls列数，zs显示总数，fw信息范围%>
<TABLE>
<%dim leixing,v,m,b,bb,rs6,sql6,vip
Call OpenConn
set rs = Server.CreateObject("ADODB.RecordSet")
sql = "select * from DNJY_info where yz=1"&ttcc&""
If fw=1 Then
sql=sql&" and hfxx=1 and hfje>0 "
End if
if jiaoyi=0 then
sql=sql&" and zt="&jiaoyi
end If
if overdue=0 Then
sql=sql&" and dqsj>"&DN_Now&""
end If
sql=sql&" order by "
If gd=1 Then sql=sql&"b "&DN_OrderType&","
sql=sql&"hfje/fbts "&DN_OrderType&",fbsj "&DN_OrderType&",id "&DN_OrderType&""
rs.open sql,conn,3,3
if rs.eof Then
Response.Write "<p><br>暂无"
IF c1>0 Then Response.Write conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
Response.Write "信息!"
else
v=0
m=0
Do while Not rs.eof and not rs.bof 
b=trim(rs("b"))
bb=len(b)
If rs("username")<>"" then
Set rs6= Server.CreateObject("ADODB.Recordset")
sql6="select vip,name from [DNJY_user] where username='"&rs("username")&"'"
rs6.open sql6,conn,1,1
if rs6.eof Then
vip=0
else
vip=rs6("vip")
End if
rs6.close
set rs6=Nothing
End if%>
<td>
 <table border="0" cellpadding="0" cellspacing="0" width="189">
    <tr class="title">
   <td width="189" height="25" rowspan="2" align="center" background="img/aaa_01.gif" color=#0037a7 style="background-color:#EFF7FE; <%If rs("hfxx")=1 then%>font-weight:bold;<%End if%> font-style:normal; "><%if rs("Readinfo")=1 Then Response.Write "<img src=""img/lock.jpg"" alt=""普通会员权限浏览"">"%><%if rs("Readinfo")=2 Then Response.Write "<img src=""img/lockvip.jpg"" alt=""VIP会员权限浏览"">"%><a href=<%=x_path("",rs("id"),"",rs("Readinfo"))%> target="_blank" ><%if rs("a")<>"0" And btc=1 then%><font color=#<%=rs("a")%>><%=CutStr(rs("biaoti"),btw,"..")%><%else%><%=CutStr(rs("biaoti"),btw,"..")%><%End if%></a></td>
  </tr>
   <tr> </tr>
   <tr class="text">   
   <td  height="135" valign="top" background="img/aaa_02.gif" style="color:#333; background-color:; ">
   <TABLE height="135">
   <TR><TD width="5"></TD><TD width="179" height="135"><%if rs("c")=1 and not rs("tupian")="0" And rs("hfxx")=1 And lg=1 then%><a href=<%=x_path("",rs("id"),"",rs("Readinfo"))%> target="_blank"><img src="UploadFiles/infopic/<%=rs("tupian")%>"  width="175" height="135" border="0" ></a><%else%>

<%
vip=0'登录会员
If request.cookies("dick")("username")<>"" Then vip=conn.Execute("Select vip From DNJY_user Where username='"&request.cookies("dick")("username")&"'")(0)
If rs("Readinfo")=1 And request.cookies("dick")("username")="" And request.cookies("dick")("username")<>rs("username") Then
response.write "<font color=#666666>普通会员权限阅读</font>"
elseIf rs("Readinfo")=2 And vip<1 And request.cookies("dick")("username")<>rs("username") Then
response.write "<font color=#666666>VIP会员权限阅读</font>"
else
if lx=1 then%><span style="color:#808080;">[<a target="_blank" href="list.asp?<%=CT_ID%>&leixing=<%=rs("leixing")%>"><%=rs("leixing")%></a>]</span><%End if%>
<%=CutStr(RemoveHTML_ttkj(rs("memo")),nyw,"")%><a href=<%=x_path("",rs("id"),"",rs("Readinfo"))%> target="_blank"> >>></a><br />联系人:<%=rs("name")%>&nbsp;<%if hy=1 then%>[<%if rs("username")<>"" then%><%if vip=1 then%>VIP会员<%else%><span title="注册会员发布的信息" style="cursor:help ">普通会员</span><%End if%><%else%><span title="游客发布的信息" style="cursor:help ">游客</span><%End if%>]<%End if%><br />电话:<%=rs("dianhua")%><br />手机:<%=rs("CompPhone")
%><%End if%><%End if%></TD><TD width="5">
</TD></TR>
   </TABLE>
</td>
    </tr>
   <tr>
  <td height="25" align="center" background="img/aaa_03.gif">
<%if b<>0 And gd=1 then
response.write "<img src=""images/num/jsq.gif"" alt=""置顶"&b&"级"">"
for i=1 to bb
response.write "<img src=""images/num/"&Mid(b,i,1)&".gif"" alt=""置顶"&b&"级"">"
next
end if%>&nbsp;<%if sj=1 then%><span title="信息有效期" style="cursor:help "><font color="#009900">剩<b><%=Fix(rs("dqsj")-now())%></b>天</font></span>&nbsp;<%End if%><%if rs("tupian")<>"0" And rs("hfxx")=1 And tp=1 then%><img src="images/num/pic.gif" alt="有图片">&nbsp;<%End if%><%if hits=1 then%><span title="点击量" style="cursor:help "><%=rs("llcs")%>次</span>&nbsp;<%End if%><%If rs("hfxx")=1 And jj=1 then%><span title="竞价剩<%=Fix(rs("dqsj")-now())%>天,平均每天<%=FormatNumber(rs("hfje")/rs("fbts"),2,-1)%>元" style="cursor:help"><%=FormatNumber(rs("hfje")/rs("fbts"),2,-1)%></span><%End if%></td>
 </tr>
 </table>
</td>
<%v=v+1
m=m+1
If m=ls then
response.write "</tr>"
m=0
End If
rs.movenext
if v>=zs then exit do
Loop
End If '配判断是否为空
response.write "</TABLE>"
Call CloseRs(rs)
end function


'电子杂志显示＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function magazine(c1,c2,c3,mh,zs,tw,th,btw)
'c1、c2、c3为地区，每行数量，总数，图宽，图高，标题长度
Dim temptext,rsmp,aemp,akmp,dd
temptext=temptext&"<table width=""100%"" height=""100%""  border=""0"" cellpadding=""0"" cellspacing=""0"">"
temptext=temptext&"<tr><td height=""90""><table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr>"
aemp=0
akmp=0
Call OpenConn
set rsmp = Server.CreateObject("ADODB.RecordSet")
sql="select  * from DNJY_zz_edition where state="&DN_True&" and thumbnail<>'' order by addtime "&DN_OrderType&",id "&DN_OrderType&""
rsmp.open sql,conn,3,3
if rsmp.eof Then
temptext=temptext&"<p><br>暂无"
IF c1>0 Then dd=conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
temptext=temptext&""&dd&"内容!"
else
do while not rsmp.eof
temptext=temptext&"<td valign=""top""><div align=""center""><table width="""&tw&""" height="""&th&""" border=""0"" cellspacing=""2"" cellpadding=""3"" class=""unnamed1"">"
temptext=temptext&"<tr><td width="""&tw&""" height="""&th&"""  valign=""top""><a href=""magazine/magazine_show.asp?id="&rsmp("id")&""" target=""_blank""><img src="""
if rsmp("thumbnail")<>"" then
temptext=temptext&rsmp("thumbnail")&""
Else
temptext=temptext&"magazine/img/ad.jpg"
End If
temptext=temptext&""" width="""&tw&""" height="""&th&"""></a></td></tr>"
temptext=temptext&"<tr><td><a title=""点击查看详情&gt;&gt;&gt;"" href=""magazine/magazine_show.asp?id="&rsmp("id")&""" arget=""_blank""><b>"&left(rsmp("editionname"),btw)&"</b></a></td></tr></table></div></td>"
aemp=aemp+1
akmp=akmp+1
If akmp=mh then
temptext=temptext&"</tr>"
akmp=0
End If               
rsmp.movenext
if aemp>=zs then exit do
  Loop
End if
rsmp.close
set rsmp=Nothing
temptext=temptext&"</tr></table> </td></tr></table>"
magazine=temptext
end function



'店铺显示方式一＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
 function f_dmtj(c1,c2,c3,h,l,btw,hg,zt,gd,v,tw,tpw,tph)
'参数依次：一级地区，二级地区，三级地区，h显示个数，l列数，btw标题长度，hg行高,zt字体,zh标题字号，gd启用固顶(0不固1固)，v限VIP显示（0不限1限），tw图文显示（0文字，1图文），tpw图片宽度，tph图片高度
dim rst_p,sqltj_a,dd,iii,i,temptext,uv
uv=""
If v=1 Then uv=" and vip=1"
temptext="<TABLE border=0 width=""100%"" style=""border-collapse: collapse"" cellpadding=0  >"
temptext=temptext&"<tr>"
Call OpenConn
set rst_p=server.createobject("adodb.recordset")
sqltj_a="select  username,sdname,logo,sdgd from [DNJY_user]  where sdname<>'' and sdkg=1"&uv&""&ttccdp&" order by "
If gd=1 then sqltj_a=sqltj_a&"sdgd "&DN_OrderType&","
sqltj_a=sqltj_a&"dldata "&DN_OrderType&",ID "&DN_OrderType&"" 	      
rst_p.open sqltj_a,conn,3,3
iii=0
if rst_p.eof Then
temptext=temptext&"<tr><td>暂无"
IF c1>0 Then dd=conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
temptext=temptext&""&dd&"" 
temptext=temptext&"内容!"
Else	
	Do While not rst_p.eof
	i=i+1
	IF iii Mod l=0 then TempText=TempText&  "<tr height="""&hg&""">"
If tw=0 Then temptext=temptext&"<td valign=""top"">"
If tw=1 Then temptext=temptext&"<td align=center valign=""top"">"
temptext=temptext&"&nbsp;<a title="""&rst_p("sdname")&""" target=""_blank"" href=""user/com.asp?com="&rst_p("username")&""">"
If tw=1 then
If rst_p("logo")<>"" Then
temptext=temptext&"<IMG  src="""&rst_p("logo")&""" width="&tpw&" height="&tph&" border=1 style=""border: 1px solid #C0C0C0;"" hspace=0 vspace=1>&nbsp;&nbsp;<br>"
else
temptext=temptext&"<IMG src=""images/nopic.gif"" width="&tpw&" height="&tph&" border=1 style=""border: 1px solid #C0C0C0;"" hspace=0 vspace=1>&nbsp;&nbsp;<br>"
End If
End If
if rst_p("sdgd")=1 And gd=1 Then
TempText=TempText&"<img src=""images/top.gif"" width=""9"" height=""15"" border=0 alt=""固顶"">"
end if
temptext=temptext&"<FONT style=""font-size:"&zt&"pt"">"&CutStr(rst_p("sdname"),btw,"..")&"</font></a>"
temptext=temptext&"</td>"
	iii=iii+1
	IF iii Mod l=0 then TempText=TempText&  "</tr>"
    if i>=h then exit do 
	rst_p.movenext
	Loop	
	IF iii Mod l<>0 then
	Do While not iii Mod l=0
	TempText=TempText&  "<td></td>"
	iii=iii+1
	Loop
	TempText=TempText&  "</tr>"
	End IF
	End If	
	TempText=TempText&  "</table>"
	rst_p.close:Set rst_p=Nothing
    
	f_dmtj=TempText

end Function%>

<%'店铺显示方式二＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function f_dmtj2(c1,c2,c3,h,l,btw,hg,zt,qz,gd,v,tw,tpw,tph)
'参数依次：一级地区，二级地区，三级地区，显示个数，列数，标题长度，行高,zt字体,qz标题前缀(0不显1显红勾2显蓝勾3显蓝箭)，gd启用固顶(0不固1固)，v限VIP显示（0不限1限）,tw图文显示（0文字，1图文），tpw图片宽度，tph图片高度
dim rst_p,sqltj_a,dd,iii,i,temptext,uv,dplogo
uv=""
If v=1 Then uv=" and vip=1"
temptext="<TABLE border=0 height=68 width=""100%"" style=""border-collapse: collapse"" cellpadding=0  >"
temptext=temptext&"<tr>"
Call OpenConn
set rst_p=server.createobject("adodb.recordset")
sqltj_a="select  username,sdname,logo,dizhi,dianhua,CompPhone,sdgd from [DNJY_user]  where sdname<>'' and sdkg=1"&uv&""&ttccdp&" order by "
If gd=1 Then sqltj_a=sqltj_a&"sdgd "&DN_OrderType&","
sqltj_a=sqltj_a&"dldata "&DN_OrderType&",ID "&DN_OrderType&"" 	      
rst_p.open sqltj_a,conn,3,3
iii=0
if rst_p.eof Then
temptext=temptext&"<tr><td>暂无"
IF c1>0 Then dd=conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
temptext=temptext&""&dd&"" 
temptext=temptext&"内容!"
Else	
	Do While not rst_p.eof
	i=i+1
	IF iii Mod l=0 then TempText=TempText&  "<tr height="""&hg&""">"
If tw=1 Then
dplogo=rst_p("logo")
If IsNull(dplogo) Then dplogo="images/nopic.gif"
temptext=temptext&"<td width=82 height=83 align=center style=""border-bottom:1px #CCCCCC dotted;text-align:left"" valign=""top""><a title="&rst_p("sdname")&" target=""_blank"" href=user/com.asp?com="&rst_p("username")&"><div style=""border:1px  #d6d6d6 solid; padding:1px;"">"
temptext=temptext&"<object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000""  codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0""  width=""100"" height=""90"">"
temptext=temptext&"<param name=""movie"" value=""img/swfoto.swf?image="&dplogo&""">" 
temptext=temptext&"<param name=""wmode"" value=""transparent"">"
temptext=temptext&"<embed  src=""img/swfoto.swf?image="""&dplogo&""" width=""100"" height=""90"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash""></embed>"
temptext=temptext&"</object></div></a></td>"
End if
temptext=temptext&"<td width=""175"" style=""border-bottom:1px #CCCCCC dotted;text-align:left"" valign=""top"">"
temptext=temptext&"<a title="&rst_p("sdname")&" target=""_blank"" href=user/com.asp?com="&rst_p("username")&"><div style=""font-size:"&zt&"px; "">&nbsp;"
If qz=1 Then TempText=TempText&"&nbsp;<IMG src=""images/qyimg/qyls/arrow_gray.gif"" border=0> "
If qz=2 Then TempText=TempText&"&nbsp;<IMG src=""images/qyimg/qyls/rowup.gif"" border=0> "
If qz=3 Then TempText=TempText&"&nbsp;<IMG src=""images/qyimg/qyls/article_common.gif"" border=0> "
if rst_p("sdgd")=1 And gd=1 Then
TempText=TempText&"<img src=""images/top.gif"" width=""9"" height=""15"" border=0 alt=""固顶"">"
end if
temptext=temptext&""&CutStr(rst_p("sdname"),btw,"..")&"</a></div><div style="" font-size:"&zt&"px; "">&nbsp;地址:"&CutStr(rst_p("dizhi"),35,"..")&"</div><div style="" font-size:"&zt&"px; "">&nbsp;电话:"&rst_p("dianhua")&"</div><div style="" font-size:"&zt&"px; "">&nbsp;手机:"&rst_p("CompPhone")&"</div>"
TempText=TempText&"</td>"

	iii=iii+1
	IF iii Mod l=0 then TempText=TempText&  "</tr>"
    if i>=h then exit do 
	rst_p.movenext
	Loop	
	IF iii Mod l<>0 then
	Do While not iii Mod l=0
	TempText=TempText&  "<td></td>"
	iii=iii+1
	Loop
	TempText=TempText&  "</tr>"
	End IF
	End If
	TempText=TempText&  "</table>"
	rst_p.close:Set rst_p=Nothing
    
	f_dmtj2=TempText

end Function%>

<%'首页企业名片（文字标签）＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function f_qymp(c1,c2,c3,h,l,btw,hg,zyw,dzw,wid,btb,bj,btc,qt,gd)     
'参数依次：一级地区，二级地区，三级地区，h显示条数，l列数，btw标题长度，hg行高，zyw主营长度，dzw地址长度，wid表格宽度，btb标题加粗，bj信笺背景(0为要信笺1不要信笺保格式2不要信笺不保留格式)，btc标题颜色，qt显示主营地址电话（0不显1显），gd更多（0不显1显）
dim rst_p,sqltj_a,dd,iii,i,temptext,btbb,infoft
If btb=1 Then
btb="<b>"
btbb="</b>"
Else
btb=""
btbb=""
End If
If bj=0 Then
infoft="infoft1"
elseIf bj=1 Then
infoft="infoft4"
Else
infoft=""
End if

temptext="<TABLE border=0  width="""&wid&""" style=""border-collapse: collapse"" cellpadding=0  >"
temptext=temptext&"<tr>"
Call OpenConn
set rst_p=server.createobject("adodb.recordset")
sqltj_a="select  username,mpname,mpjj,mpfg,logo,zhuying,dizhi,dianhua from [DNJY_user]  where mpname<>'' and mpkg=1"&ttccmp&"  order by mpgd "&DN_OrderType&",dldata "&DN_OrderType&",ID "&DN_OrderType&""
rst_p.open sqltj_a,conn,3,3
if rst_p.eof Then
temptext=temptext&"<p><br>暂无"
IF c1>0 Then dd=conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
temptext=temptext&""&dd&"" 
temptext=temptext&"内容!"
TempText=TempText&  "</table>"
Else	
	iii=0
	Do While not rst_p.eof
	i=i+1
	IF iii Mod l=0 then TempText=TempText&  "<tr height="""&hg&""">"
	temptext=temptext&"<td  align=""left"" valign=""top"">"
	temptext=temptext&"<table width=100% border=0 cellspacing=0 cellpadding=0 class="""&infoft&""">"
    temptext=temptext&"<tr><td>"
    temptext=temptext&"<A target=""_blank"" href=""company.asp?username="&rst_p("username")&"""><font color=#"&btc&">"&btb&""&CutStr(rst_p("mpname"),btw,"..")&""&btbb&"</font></a>"
If qt=1 then
    temptext=temptext&"<br>主营: "&CutStr(rst_p("zhuying"),zyw,"..")&"<br>"
    temptext=temptext&"地址："&CutStr(rst_p("dizhi"),dzw,"..")&"<br>"
    temptext=temptext&"电话："&rst_p("dianhua")&"<br>"
End if
    temptext=temptext&"</td></tr></table>" 
    temptext=temptext&"</td>"
	iii=iii+1
	IF iii Mod l=0 then TempText=TempText&  "</tr>"
    if i>=h then exit do 
	rst_p.movenext
	Loop	
	IF iii Mod l<>0 then
	Do While not iii Mod l=0
	TempText=TempText&  "<td></td>"
	iii=iii+1
	Loop
	TempText=TempText&  "</tr>"
	End IF
	TempText=TempText&  "</table>"
	
	If gd=1 Then temptext=temptext&"<table align=""right"" class="""&infoft&"""><tr><td ><a href=""qyindex.asp?ttop=4&"&C_ID&""" target=""_blank"">更多>>></a></td></tr></table>"
	End IF
	rst_p.close:Set rst_p=Nothing
    
	f_qymp=TempText

end Function

'首页企业名片图片＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function qymp(c1,c2,c3,mh,zs,btw,zyw)
'c1、c2、c3为地区，每行数量，总数，标题长度，主营长度%>
<table width="100%" height="100%"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="90">
         <table width="100%" height="3" border="0" cellpadding="0" cellspacing="0">
           <tr>
             <td></td>
           </tr>
         </table>
         <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
<%Dim rsmp,aemp,akmp
aemp=0
akmp=0
Call OpenConn
set rsmp = Server.CreateObject("ADODB.RecordSet")
sql="select  * from [DNJY_user] where mpkg=1 and mpname<>''"&ttccmp&" order by mpgd "&DN_OrderType&",id "&DN_OrderType&""
rsmp.open sql,conn,3,3
if rsmp.eof Then
Response.Write "<p><br>暂无"
IF c1>0 Then Response.Write conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
Response.Write "内容!"
else
do while not rsmp.eof
%>
    <td valign="top"><div align="center"><table width="150" height="120" border="0" cellspacing="2" cellpadding="3" class="unnamed1">                    <tr> 
                      <td width="150" height="120"  valign="top"><a href="company.asp?username=<%=rsmp("username")%>" target="_blank"><img src="<%if right(rsmp("logo"),4)=".gif" or right(rsmp("logo"),5)=".jpeg" or right(rsmp("logo"),4)=".jpg" then%>
					  <%=rsmp("logo")%>
					  <%else%>images/nomp.gif<%End if%>" width="150" height="120" alt="<%=left(rsmp("mpname"),btw)%>，电话：<%=rsmp("dianhua")%>，主营：<%=left(rsmp("zhuying"),zyw)%>"></a></td></tr>
					  <tr><td><a title="点击查看详情&gt;&gt;&gt;" href="company.asp?username=<%=rsmp("username")%>" target="_blank"><b><%=left(rsmp("mpname"),btw)%></b></a></td></tr></table>
    </div>      
</td>
<%aemp=aemp+1
akmp=akmp+1
If akmp=mh then
response.write "</tr>"
akmp=0
End If               
                rsmp.movenext
				if aemp>=zs then exit do
                Loop
                End if
                rsmp.close
                set rsmp=Nothing
                %>
  </tr>
</table>  </td>
  </tr>
</table>
<%end Function%>


<%'分类信息列表＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function xxfl(c1,c2,c3,lx,tm,ts,ls,btw,btc,gd,tp,dj,rq,tw,bj)
'c1、c2、c3为地区，lx信息类型（0不显1显），tm显示分类个数，ts显示条数，ls列数，btw标题长度，btc标题颜色（0不显1显）,gd启用固顶（0不选1选，tp图片标志（0不显1显），dj点击数（0不显1显），rq日期（0不显1显）,tw头条图文（0不选1选），bj信笺背景(0为要信笺1不要信笺保格式2不要信笺不保留格式)）
%>
<table border=0 width="100%" >
<tr>
<%Dim rsclass,sql9,k,i,j,leixing,m,n,infoft,g,bb
If bj=0 Then
infoft="infoft1"
elseIf bj=1 Then
infoft="infoft4"
Else
infoft=""
End If
Call OpenConn
set rsclass=server.createobject("adodb.recordset")
rsclass.Open "select * from DNJY_type where twoid=0 and threeid=0 order by indexid",conn,1,1
if rsclass.eof Then
Response.Write "<p><br>暂无分类信息"
else
m=0
n=0
Do while Not rsclass.eof%>
<%'//判断每两行显示一个广告
If m Mod ls*2=0 Then
n=n+1%>
<table width="100%" cellspacing="0" cellpadding="0" align="center" bgcolor="#FFFFFF">
<tr><td  valign="top" colspan="2" align="center"><script src="<%=adjs_path("ads/js","xxfl"&n&"",c1&"_"&c2&"_"&c3)%>"></script></td></tr>
<table>
<%m=0'下面的shadow1在style.css中css设置板块宽度
End if%><td valign=top width="100%" ><div class="shadow1">
<%k=0	
set rs = Server.CreateObject("ADODB.RecordSet")
sql = "select * from DNJY_info where yz=1 and type_oneid="&rsclass("id")&""&ttcc&""
if jiaoyi=0 then
sql=sql&" and zt="&jiaoyi
end If
if overdue=0 Then
sql=sql&" and dqsj>"&DN_Now&""
end If
'sql=sql&" and leixing<>'出售'"'车网，启用此行
sql=sql&" order by"
If gd=1 Then sql=sql&" b "&DN_OrderType&","
sql=sql&" fbsj "&DN_OrderType&""
rs.open sql,conn,3,3
%>

<table width="100%" cellspacing="0" cellpadding="0" align="center" bgcolor="#FFFFFF" class="stock">
<tr><td  valign="top" class="td2">
<A href="list.asp?t=<%=rsclass("id")%>,0,0&<%=C_ID%>&dick=1" style="cursor:hand"><b><font color="#ffffff"><span class="infoft2"><%=rsclass("name")%></span></font></b></A> <font color="#999999">(本类有<%=rs.Recordcount%>条信息)</font>
</td><td class="td2" align="right"><A href="list.asp?t=<%=rsclass("id")%>,0,0&<%=C_ID%>&dick=1" style="cursor:hand">更多>>></A></td></tr>
 <tr><%if rs.eof Then
Response.Write "<tr><td><p><br>暂无"
IF c1>0 Then Response.Write conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
Response.Write "信息!</td></tr>"
Else
j=1%>
<td><%do while not rs.eof
b=trim(rs("b"))
bb=len(b)%></td>
</tr>
     <tr>
       <td width="10%" colspan="2">
<%if k=0 And tw=1 then%>
<TABLE>
	   <TR>
		<TD><%Response.Write"<a href="&x_path("",rs("id"),"",rs("Readinfo"))&" target=_blank><img style="" border:2 double #E2E2E2"" border=1  WIDTH=""70"" HEIGHT=""70"" onload='DrawImage(this)' src="
if rs("c")=1 and not rs("tupian")="0" Then
Response.Write"UploadFiles/infopic/"&rs("tupian")&">"
Else
Response.Write"images/nopoto.gif>"
end If%></a></TD>
		<TD><%Response.Write"<div style=""font-size=13px;text-align:center"">["&rs("leixing")&"]<a style=""color:#"&rs("a")&";font-size=13px;text-align:center"" target=_blank href='list.asp?t="&rsclass("id")&",0,0&"&C_ID&"&leixing="&rs("leixing")&"'>"&CutStr(trim(rs("biaoti")),btw-2,"..")&"</a></div>"%><div>&nbsp;&nbsp;&nbsp;&nbsp;<%=CutStr(rs("memo"),150,"..")%>(<FONT color=#999999><%=FormatDate(rs("fbsj"),10)%></font>)</div></TD>
	   </TR>
	   </TABLE>	
<%else%>   
	   <div align="left" class="<%=infoft%>">
<%If lx=1 then%>[<a target="_blank" href="list.asp?t=<%=rsclass("id")%>,0,0&<%=C_ID%>&leixing=<%=rs("leixing")%>"><%=rs("leixing")%></a>]<%End if%>
<%if rs("Readinfo")=1 Then Response.Write "<img src=""img/lock.jpg"" alt=""普通会员权限浏览"">"%><%if rs("Readinfo")=2 Then Response.Write "<img src=""img/lockvip.jpg"" alt=""VIP会员权限浏览"">"%>
<%if tp=1 and rs("tupian")<>"0" Then%><img src="images/num/pic.gif" alt="有图片"><%End if%><a target="_blank" href="<%=x_path("",rs("id"),"",rs("Readinfo"))%>" title="<%=rs("biaoti")&"||"&rs("name")%>发布于<%=FormatDateTime(rs("fbsj"),vbShortDate)%>"><%If btc=1 then%><FONT color=#<%=rs("a")%>><%End if%><%=CutStr(rs("biaoti"),btw,"..")%><%If btc=1 then%></font><%End if%></a>
<%if b<>0 And gd=1 then
Response.Write "<img src=""images/num/jsq.gif"" alt=""置顶"&b&"级"">"
for g=1 to bb
Response.Write "<img src=""images/num/"&Mid(b,g,1)&".gif"" alt=""置顶"&b&"级"">"
Next
end If

If dj=1 then%><FONT color=#999999><%=rs("llcs")%>|</font><%End if%>
<FONT color=#999999><%=FormatDate(rs("fbsj"),rq)%></font></div></td></tr>
 <%
End If
k=k+1 
rs.movenext
if k>=ts then exit do
Loop
End If

Call CloseRs(rs)%> </table></div></td>
<% i=i+1
If i=ls then
	response.write "</tr>"
	i=0
	End If
m=m+1
rsclass.movenext
if m>=tm then exit do
Loop
End if
rsclass.close
set rsclass=Nothing
%>
</tr>
   </table>
<%end Function

'新闻资讯栏目（调用一、二、三级分类）
Function news_class(by1,by2,by3,ts1,ts2,ts3,ls,zt1,zt2,zt3,hg1,hg2,hg3,file)
 '标签说明,依次为：tc1,tc2,tc3一二三级分类ID号（0 为全部），by1是否显示一级分类（0不显1显)，by2是否显示二级分类（0不显1显)，by3是否显示三级分类（0不显1显)，ts1一级分类显示个数(0为全部），ts2二级分类显示个数（0为全部），ts3三级分类显示个数（0为全部）,ls一级分类显示列数（0不分列)，zt1一级分类字号，zt2二级分类字号，zt3三级分类字号，hg1一级分类行距，hg2二级分类行距，hg3三级分类行距，file前往页面文件名
 Dim n1,n2,n3,temptext,rsi,rst3,ti,ts,tj,tt,titlecolor,titlecolor2,titlecolor3,lname,kg
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
End if
    temptext="<TABLE border=""0"" cellspacing=""1"" cellpadding=""0"">"
    temptext=temptext&"<tr><td valign=""top"">"
	 Call OpenConn
	set rs=server.CreateObject("adodb.recordset")
	rs.Open "select * from DNJY_news_class where oneid>0 and twoid=0 and threeid=0 order by sorting",conn,1,1
    if rs.eof Then
    Response.Write "<font color=000000>暂无分类或未设置首页显示</font>"
    else
	tj=0
	do while not rs.eof
	  titlecolor="1835D1"
      If rs("oneid")=n1 Then titlecolor="ff0000"
      temptext=temptext&"<TABLE border=""0"" cellspacing=""1"" cellpadding=""0"">"
	  temptext=temptext&"<TR height="""&hg1&"""><TD></TD><TD>"
	  If by1=1 then
	  temptext=temptext&"<A href="""&file&"?"&C_ID&"&n="&rs("oneid")&","&rs("twoid")&","&rs("threeid")&"""><FONT style=""font-size:"&zt1&"pt"" color=""#"&titlecolor&"""><B>"&rs("name")&"</B></FONT></A>"
	 End if
	  temptext=temptext&"</TD></TR>"	  
       set rsi=conn.execute("select * from DNJY_news_class where oneid="&rs("oneid")&" and twoid<>0 and threeid=0 order by  sorting,id,twoid")
	       If by2=1 then
			temptext=temptext&"<tr height="""&hg1&"""><td>&nbsp;&nbsp;</td><td  width="""&920/ls&""">"
			ts=0
			do while not rsi.eof
			titlecolor2="1835D1"
            If rsi("oneid")=n1 And rsi("twoid")=n2 Then titlecolor2="ff0000"
            temptext=temptext&"<A href="""&file&"?"&C_ID&"&n="&rs("oneid")&","&rsi("twoid")&","&rsi("threeid")&"""> <FONT style=""font-size:"&zt2&"pt"" color=""#"&titlecolor2&""">"&rsi("name")&"</FONT></A>"
			If by3=0 then
		    temptext=temptext&"<FONT color=""#CCCCCC"">|</FONT>"
            Else
            temptext=temptext&"<br>"
            set rst3=conn.execute("select * from DNJY_news_class where oneid="&rs("oneid")&" and twoid="&rsi("twoid")&" and threeid<>0 order by  sorting,oneid,twoid,threeid")
			if Not rst3.eof Then
		    temptext=temptext&"<table><tr height="""&hg3&"""><td></td><td>└"
			tt=0
			do while not rst3.eof 
			titlecolor3="7DA2F3"
            If rsi("oneid")=n1 And rsi("twoid")=n2 And rst3("threeid")=n3 Then titlecolor3="ff0000"
			temptext=temptext&"<A href="""&file&"?n="&rs("oneid")&","&rsi("twoid")&","&rst3("threeid")&"""><FONT style=""font-size:"&zt3&"pt"" color=""#"&titlecolor3&""">"&rst3("name")&"</FONT></A>"
			temptext=temptext&"<FONT color=""#CCCCCC"">|</FONT>"
			tt=tt+1
			if tt>=ts3 And ts3>0 Then
            temptext=temptext&"..."
			exit Do
			End if
			rst3.movenext
		   Loop
		  temptext=temptext&"</td></tr></table>"
		   End if
		   rst3.close
           End If

           ts=ts+1
		   if ts>=ts2 And ts2>0 Then
		   temptext=temptext&"<A href="""&file&"?"&C_ID&"&n="&rs("oneid")&","&rs("twoid")&","&rs("threeid")&""">>>></A>"
		   exit do 
           End if
           rsi.movenext
		   Loop
		   End if
		   temptext=temptext&"</td></tr></table>"
		   ti=ti+1
		   if ti mod ls=0 then
		   temptext=temptext&"</td></tr><tr height="""&hg1&"""><td valign='top'>"
		   else
		   temptext=temptext&"</td><td valign='top'>"
		   end if
		   rsi.close
           set rsi=nothing
           tj=tj+1
		   if tj>=ts1 And ts1>0 then exit do 
		   rs.movenext
		   Loop
end if
Call CloseRs(rs)
temptext=temptext&"</span></td></tr></TABLE>"		  
news_class=temptext
End Function


'新闻资讯标题式显示＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function news(c1,c2,c3,lx,n1,n2,n3,gd,tj,tp,hits,ma,d,h,l,btw,zt,hg)         
'参数依次：c1,c2,c3一级地区，二级地区，三级地区，lx类型（0全部、1为推荐、2为图文、3为热门），n1,n2,n3栏目类别（根据ID号定）0或空为全部，gd启用固顶（0不固1固），tj推荐标志（0不显1显），tp图片标志（0不显1显），hits点击数（0不显1显）,ma移动（0不显1显），d时间（0-15种格式，0不显），h显示条数，l列数，btw标题长度，zt字号，hg行高
	Dim ii,i,Sql,TempText,newstop,rs,co,nn
	Dim id,title,oColor,infotime,firstImageName,img,lxx,dd,idd
	
	IF lx=1 then
	lxx="and tuijian=1"	
	elseif lx=2 then
	lxx=" and img="&DN_True&"" 
    elseif lx=0 then
    lxx=""	
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
	Call OpenConn
    Set Rs=server.createobject("aDodb.recordset")               
	Sql="select id,title,hits,oColor,infotime,newstop,img,tuijian,c_oneid,c_twoid,c_threeid from DNJY_News where ifshow=1"&nn&""&lxx&" "&ttccnews&" order by "
	If gd=1 Then sql=sql&"newstop "&DN_OrderType&","
	if lx=3 Then
	sql=sql&"hits "&DN_OrderType&","
	End if
	sql=sql&"id "&DN_OrderType&""
	Rs.open Sql,conn,1,1
	TempText=""
if rs.eof Then
TempText=" <p><br>暂无"
IF c1>0 Then dd=conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
TempText=TempText&  ""&dd&"内容!"
Else
	IF ma=1 Then TempText= "<marquee scrollamount=""1"" scrolldelay=""50"" onmouseover=""this.stop()"" onmouseout=""this.start()"" direction=up style=""width:100%;padding:0"" height='95%'>"
	TempText=TempText&"<table width=""100%""  border=""0"" cellspacing=""2"" cellpadding=""0"" >"
	dim iii
	iii=0
	Do While not Rs.eof
i=i+1
Select Case rs("oColor")
Case "a01"
co="FF0000"
Case "a02"
co="0000FF"
Case "a03"
co="00FFFF"
Case "a04"
co="FF9900"
Case "a05"
co="339966"
Case Else
co="000000"
End Select

	IF iii Mod l=0 then TempText=TempText&  "<tr height="""&hg&""">"
	TempText=TempText&  "<td>"
    IF rs("newstop")=1 And gd=1 Then TempText=TempText&  "<img src=""img/top4.gif"" width=""15"" height=""13"" border=0 alt=""固顶"">"
    IF rs("tuijian")=1 And tj=1 Then TempText=TempText&  "<img src=""img/jian.gif"" width=""15"" height=""13"" border=0 alt=""推荐"">"
    IF rs("img")=True And tp=1 Then TempText=TempText&  "<img src=""images/num/pic.gif"" width=""13"" height=""13"" border=0 alt=""相册"">"
	TempText=TempText&  "<a title="""&rs(1)&""" href="""&S_path("",rs(0),"",""&n1&"")&"&n="&rs(8)&","&rs(9)&","&rs(10)&"&"&C_ID&""" target=""_blank""><SPAN style=""color:"&co&";font-size="&zt&"px"">"&CutStr(rs(1),btw,"..")&"</SPAN></a> "
    IF hits=1 Then TempText=TempText&  " <font color=""#999999"" >["&rs(2)&"]</font>" 
	TempText=TempText&  " <font color=""#999999"" >"&FormatDate(rs(4),d)&"</font>"    
	TempText=TempText&  "</TD>"
	iii=iii+1
	IF iii Mod l=0 then TempText=TempText&  "</tr>"
    if i>=h then exit do 
	Rs.movenext
	Loop	
	IF iii Mod l<>0 then
	Do While not iii Mod l=0
	TempText=TempText&  "<td></td>"
	iii=iii+1
	Loop
	TempText=TempText&  "</tr>"
	End IF
	TempText=TempText&  "</table>"
	IF ma=1 Then TempText=TempText& "</marquee>"
	End IF
Call CloseRs(rs)
	news=TempText
end Function

'新闻资讯相册列表＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function news_photo(c1,c2,c3,n1,n2,n3,gd,tj,img,bt,btw,tw,th,ls,zs,fy,page,s1,s2)
'c1,c2,c3为地区，n1,n2,n3为分类ID号，gd是否显示固顶，tj是否显示推荐,img是否显示无图文章0不显1显，bt是否显示标题0不显1显，btw标题长度，tw图片宽度，th图片高度，ls每行列量，zs每页个数，是否显示翻页0不显1显，page翻页号,s1搜索类型，s2搜索关键词
dim TempText,idimg,nn,totalPut,CurrentPage,TotalPages,j,k,filename,MaxPerPage,gd1,gd2,tj2,ss
Session("n1")=n1
Session("n2")=n2
Session("n3")=n3
filename=Request.ServerVariables("PATH_INFO")
MaxPerPage=zs
if Not isempty(SafeRequest("page",1)) then
  currentPage=Cint(SafeRequest("page",1))
  else
  currentPage=1
end if 
	If img=0 Or img="" Then
	idimg=" and img=1"
	Else
	idimg=""
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
	
	ss=""
	If s1="1" Then
	ss=" and title like '%"&s2&"%'"
	ElseIf s1="2" Then
	ss=" and content like '%"&s2&"%'"
	Else
	ss=""
	End if	
	
Call OpenConn
If gd=1 Then
gd1=" newstop desc,"
Else
gd1=""
End if
set rs = server.CreateObject("adodb.recordset")
rs.open "select * from DNJY_news where ifshow=1"&ss&""&idimg&""&nn&""&ttccnews&" order by"&gd1&" id "&DN_OrderType&"",conn,1,1
if rs.eof And rs.bof then'-----a1-----
TempText="<p align='center' class='contents'>没有记录！</p>"
else'------a2--------
	totalPut=rs.recordcount
    if currentpage<1 then
    currentpage=1
    end if
    if (currentpage-1)*MaxPerPage>totalput then
	   	if (totalPut mod MaxPerPage)=0 then
	    currentpage= totalPut \ MaxPerPage
	   	else
	    currentpage= totalPut \ MaxPerPage + 1
	   	end if
    end If
if currentPage=1 then'--------b1------
i=0
k=0
TempText=TempText&"<table width='100%' border='0' cellspacing='1' cellpadding='1' align='center'><tr>"
do while not rs.eof
TempText=TempText&"<td width='"&tw&"' height='"&th+35&"'><a href='news_show.asp?Id="&rs("Id")&"&"&C_ID&"&n="&rs("c_oneid")&","&rs("c_twoid")&","&rs("c_threeid")&"' target=_blank><img border=0 width='"&tw&"' height='"&th&"' src='"
If Instr(rs("s_photo"),"|") Then
If Instr(Mid(rs("s_photo"),2),"|") Then
TempText=TempText&"/"&strInstallDir&""&cut_string(Mid(rs("s_photo"),2),"|")&"'>"
else
TempText=TempText&"/"&strInstallDir&""&Mid(rs("s_photo"),2)&"'>"
End if
Else
TempText=TempText&"/"&strInstallDir&"images/photono.gif'>"
End If
IF rs("newstop")=1 And gd=1 Then gd2="<img src=""img/top4.gif"" width=""15"" height=""13"" border=0 alt=""固顶"">"
if rs("tuijian")=1 And tj=1 Then tj2="<img src='img/jian.gif' border=0 alt='推荐文章'>"
If bt=1 Then TempText=TempText&"<br><center>"&gd2&""&tj2&""&replace(CutStr(rs("title"),btw,".."),s2,"<font color=#FF0000>" &s2& "</font>")&"</center>"
TempText=TempText&"</a></td>"
i=i+1	
k=k+1
If k=ls then
TempText=TempText&"</tr>" 'N个后换行
k=0
End If
gd2=""
if i>=MaxPerPage then exit Do
rs.movenext
Loop
Call CloseRs(rs)
TempText=TempText&"</tr></table>"
            	If fy=1 Then'--1----
				If totalput Mod maxperpage=0 Then  
					n= totalput \ maxperpage  
				Else
					n= totalput \ maxperpage+1  
				End If
TempText=TempText&"<form method=Post action= "&filename&"?"&C_ID&"&n="&Session("tc1")&","&Session("tc2")&","&Session("tc3")&"&s1="&s1&"&s2="&s2&"><p align='center'>"
				If CurrentPage<2 Then
TempText=TempText&"首页 上页" 
				Else  
TempText=TempText&"<a href="&filename&"?page=1&"&C_ID&"&n="&Session("tc1")&","&Session("tc2")&","&Session("tc3")&"&s1="&s1&"&s2="&s2&">首页</a> <a href="&filename&"?page="&CurrentPage-1&"&"&C_ID&"&n="&Session("tc1")&","&Session("tc2")&","&Session("tc3")&"&s1="&s1&"&s2="&s2&">上页</a>" 
				End If	
				If n-currentpage<1 Then 
TempText=TempText&" 下页 末页" 
				Else
TempText=TempText&"<a href="&filename&"?page="&(CurrentPage+1)&"&"&C_ID&"&n="&Session("tc1")&","&Session("tc2")&","&Session("tc3")&"&s1="&s1&"&s2="&s2&"> 下页</a> <a href="&filename&"?page="&n&"&"&C_ID&"&n="&Session("tc1")&","&Session("tc2")&","&Session("tc3")&"&s1="&s1&"&s2="&s2&">末页</a>" 
				End If
TempText=TempText&" &nbsp;&nbsp;第<b>"&CurrentPage&"</b>页 共<b>"&n&"</b>页 共<b>"&totalput&"</b>个记录　每页<b>"&maxperpage&"</b>个记录 转到第<input type='text' name='page' size=2 maxlength=10 class=smallInput value="&currentpage&">页<input type='submit'  class='contents' value='跳转' name='cndok'></form>"
                End if'----1----
       	else'---------b2-------------
          				if (currentPage-1)*MaxPerPage<totalPut then'----------c1--------
            				rs.move  (currentPage-1)*MaxPerPage
            				dim bookmark
            				bookmark=rs.bookmark
i=0
k=0
TempText=TempText&"<table width='100%' border='0' cellspacing='1' cellpadding='1' align='center'><tr>"
do while not rs.eof
TempText=TempText&"<td width='"&tw&"'><a href='news_show.asp?Id="&rs("Id")&"&"&C_ID&"&n="&rs("c_oneid")&","&rs("c_twoid")&","&rs("c_threeid")&"' target=_blank><img border=0 width='"&tw&"' height='"&th&"' src='"
If Instr(rs("s_photo"),"|") Then
TempText=TempText&"/"&strInstallDir&""&cut_string(rs("s_photo"),"|")&"'>"
Else
TempText=TempText&"/"&strInstallDir&"images/photono.gif'>"
End If
IF rs("newstop")=1 And gd=1 Then gd2="<img src=""img/top4.gif"" width=""15"" height=""13"" border=0 alt=""固顶"">"
if rs("tuijian")=1 And tj=1 Then tj2="<img src='img/jian.gif' border=0 alt='推荐文章'>"
If bt=1 Then TempText=TempText&"<br><center>"&gd2&""&tj2&""&replace(CutStr(rs("title"),btw,".."),s2,"<font color=#FF0000>" &s2& "</font>")&"</center>"
TempText=TempText&"</a></td>"
i=i+1	
k=k+1
If k=ls then
TempText=TempText&"</tr>" 'N个后换行
k=0
End If
gd2=""
if i>=MaxPerPage then exit Do
rs.movenext
Loop
Call CloseRs(rs)
TempText=TempText&"</tr></table>"
             	If fy=1 Then'---2---
				If totalput Mod maxperpage=0 Then  
					n= totalput \ maxperpage  
				Else
					n= totalput \ maxperpage+1  
				End If
TempText=TempText&"<form method=Post action= "&filename&"?"&C_ID&"&n="&Session("tc1")&","&Session("tc2")&","&Session("tc3")&"&s1="&s1&"&s2="&s2&"><p align='center'>"
				If CurrentPage<2 Then
TempText=TempText&"首页 上页" 
				Else  
TempText=TempText&"<a href="&filename&"?page=1&"&C_ID&"&n="&Session("tc1")&","&Session("tc2")&","&Session("tc3")&"&s1="&s1&"&s2="&s2&">首页</a> <a href="&filename&"?page="&CurrentPage-1&"&"&C_ID&"&n="&Session("tc1")&","&Session("tc2")&","&Session("tc3")&"&s1="&s1&"&s2="&s2&">上页</a>" 
				End If	
				If n-currentpage<1 Then 
TempText=TempText&" 下页 末页" 
				Else
TempText=TempText&"<a href="&filename&"?page="&(CurrentPage+1)&"&"&C_ID&"&n="&Session("tc1")&","&Session("tc2")&","&Session("tc3")&"&s1="&s1&"&s2="&s2&"> 下页</a> <a href="&filename&"?page="&n&"&"&C_ID&"&n="&Session("tc1")&","&Session("tc2")&","&Session("tc3")&"&s1="&s1&"&s2="&s2&">末页</a>" 
				End If
TempText=TempText&" &nbsp;&nbsp;第<b>"&CurrentPage&"</b>页 共<b>"&n&"</b>页 共<b>"&totalput&"</b>个记录　每页<b>"&maxperpage&"</b>个记录 转到第<input type='text' name='page' size=2 maxlength=10 class=smallInput value="&currentpage&">页<input type='submit'  class='contents' value='跳转' name='cndok'></form>"
                End if'---2-----
        				else'-------------c2----------------
currentPage=1
i=0
k=0
TempText=TempText&"<table width='100%' border='0' cellspacing='1' cellpadding='1' align='center'><tr>"
do while not rs.eof
TempText=TempText&"<td width='"&tw&"'><a href='news_show.asp?Id="&rs("Id")&"&"&C_ID&"&n="&rs("c_oneid")&","&rs("c_twoid")&","&rs("c_threeid")&"' target=_blank><img border=0 width='"&tw&"' height='"&th&"' src='"
If Instr(rs("s_photo"),"|") Then
TempText=TempText&"/"&strInstallDir&""&cut_string(rs("s_photo"),"|")&"'>"
Else
TempText=TempText&"/"&strInstallDir&"images/photono.gif'>"
End If
IF rs("newstop")=1 And gd=1 Then gd2="<img src=""img/top4.gif"" width=""15"" height=""13"" border=0 alt=""固顶"">"
if rs("tuijian")=1 And tj=1 Then tj2="<img src='img/jian.gif' border=0 alt='推荐文章'>"
If bt=1 Then TempText=TempText&"<br><center>"&gd2&""&tj&""&replace(CutStr(rs("title"),btw,".."),s2,"<font color=#FF0000>" &s2& "</font>")&"</center>"
TempText=TempText&"</a></td>"
i=i+1	
k=k+1
If k=ls then
TempText=TempText&"</tr>" 'N个后换行
k=0
End If
gd2=""
if i>=MaxPerPage then exit Do
rs.movenext
Loop
Call CloseRs(rs)
TempText=TempText&"</tr></table>"
           		If fy=1 Then'---3---
				If totalput Mod maxperpage=0 Then  
					n= totalput \ maxperpage  
				Else
					n= totalput \ maxperpage+1  
				End If
TempText=TempText&"<form method=Post action= "&filename&"?"&C_ID&"&n="&Session("tc1")&","&Session("tc2")&","&Session("tc3")&"&s1="&s1&"&s2="&s2&"><p align='center'>"
				If CurrentPage<2 Then
TempText=TempText&"首页 上页" 
				Else  
TempText=TempText&"<a href="&filename&"?page=1&"&C_ID&"&n="&Session("tc1")&","&Session("tc2")&","&Session("tc3")&"&s1="&s1&"&s2="&s2&">首页</a> <a href="&filename&"?page="&CurrentPage-1&"&"&C_ID&"&n="&Session("tc1")&","&Session("tc2")&","&Session("tc3")&"&s1="&s1&"&s2="&s2&">上页</a>" 
				End If	
				If n-currentpage<1 Then 
TempText=TempText&" 下页 末页" 
				Else
TempText=TempText&"<a href="&filename&"?page="&(CurrentPage+1)&"&"&C_ID&"&n="&Session("tc1")&","&Session("tc2")&","&Session("tc3")&"&s1="&s1&"&s2="&s2&"> 下页</a> <a href="&filename&"?page="&n&"&"&C_ID&"&n="&Session("tc1")&","&Session("tc2")&","&Session("tc3")&"&s1="&s1&"&s2="&s2&">末页</a>" 
				End If
TempText=TempText&" &nbsp;&nbsp;第<b>"&CurrentPage&"</b>页 共<b>"&n&"</b>页 共<b>"&totalput&"</b>个记录　每页<b>"&maxperpage&"</b>个记录 转到第<input type='text' name='page' size=2 maxlength=10 class=smallInput value="&currentpage&">页<input type='submit'  class='contents' value='跳转' name='cndok'></form>"
                End if'----3----
	      				end if'---------c3-----
	   				end if'--------b3------------     
end If'--------a3-----------------
news_photo=TempText
Call CloseDB(conn)
end Function



'最新回复＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function hf(c1,c2,c3,ts,lx,bj,tp,btw,btc,dj,rq)
'c1、c2、c3为地区，ts显示条数，lx信息类型，bj信笺背景(0为要信笺1不要信笺保格式2不要信笺不保留格式)，tp图片标志，btw标题长度，btc标题颜色，dj点击数，rq日期（0-15种格式，0不显）
%>
<table width="100%">
<%dim m,infoft
If bj=0 Then
infoft="infoft1"
elseIf bj=1 Then
infoft="infoft4"
Else
infoft=""
End if
m=0
Call OpenConn
set rs = Server.CreateObject("ADODB.RecordSet")
sql = "select * from DNJY_info where yz=1 and hfcs>=1"&ttcc&""
if jiaoyi=0 then
sql=sql&" and zt="&jiaoyi
end If
if overdue=0 Then
sql=sql&" and dqsj>"&DN_Now&""
end If
sql=sql&" order by hfsj "&DN_OrderType&""
rs.open sql,conn,3,3
if rs.eof Then
Response.Write "<p><br>暂无"
IF c1>0 Then Response.Write conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
Response.Write "回复信息!"
else
do while not rs.eof%>
<tr>
       <td><div class="<%=infoft%>"><%If lx=1 then%>[<a target="_blank" href="list.asp?t=<%=rsclass("id")%>,0,0&<%=C_ID%>&leixing=<%=rs("leixing")%>"><%=rs("leixing")%></a>]<%End if%><%if rs("Readinfo")=1 Then Response.Write "<img src=""img/lock.jpg"" alt=""普通会员权限浏览"">"%><%if rs("Readinfo")=2 Then Response.Write "<img src=""img/lockvip.jpg"" alt=""VIP会员权限浏览"">"%><%if tp=1 and rs("tupian")<>"0" Then%><img src="images/num/pic.gif" alt="有图片"><%End if%><a  href=<%=x_path("",rs("id"),"",rs("Readinfo"))%> target="_blank"><%If btc=1 then%><FONT color=#<%=rs("a")%>><%End if%><%=CutStr(rs("biaoti"),btw,"..")%><%If btc=1 then%></font><%End if%></a>
<%If dj=1 then%><FONT color=#999999><%=rs("llcs")%>|</font><%End if%>
<FONT color=#999999><%=FormatDate(rs("fbsj"),rq)%></font></a>
</td>
     </tr>
               <%m=m+1
                if m>=ts then exit do
                rs.movenext
                Loop%>
<tr><td align="right"><div class="<%=infoft%>"><a href="list.asp?xxsx=5&<%=CT_ID%>" target=_blank>更多>>></a></div></td></tr>
                <%End if
Call CloseRs(rs)

                %>            
				</table>
<%end function%>


<%' 客户留言＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function book(c1,c2,c3,ts,bj,hf,btw,zz,rq)
'c1、c2、c3为地区，ts显示条数，bj信笺背景(0为要信笺1不要信笺保格式2不要信笺不保留格式)，hf回复状态,btw标题长度，zz作者，rq日期（0-15种格式，0不显）%>
<table width="100%">
<%Dim b,infoft
If bj=0 Then
infoft="infoft1"
elseIf bj=1 Then
infoft="infoft4"
Else
infoft=""
End if
b=0
Call OpenConn
set rs=server.createobject("adodb.recordset")
rs.open "select * from DNJY_gbook where hf=1 and known=0"&ttccbook&" order by fbsj "&DN_OrderType&"",conn,1,1
if rs.eof then
Response.Write "<p><br>暂无"
IF c1>0 Then Response.Write conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
Response.Write "用户留言!"
else%>
<%do while not rs.eof%>
    <tr>
       <td><div class="<%=infoft%>"><%If rs("hf")=1 And hf=1 Then response.write "<font color=#FF0000>[回]</font>"%><A href="javascript:win=open('user/book.asp?id=<%=rs("id")%>','offer','width=450,height=500,status=no,menubar=no,scrollbars=yes,top=0,left=0'); win.focus()"><%=CutStr(rs("gbook1"),btw,"..")%>
<%If rs("username")<>"" And zz=1 then%><FONT color=#999999><%=(rs("username"))%>|</font><%End if%>
<FONT color=#999999><%=FormatDate(rs("fbsj"),rq)%></font>	</a>
</td></tr>
<%b=b+1
 if b>=ts then exit do
 rs.movenext
 Loop%>
 <tr><td align="right"><div class="<%=infoft%>"><a href="Message_board.asp?<%=C_ID%>" target=_blank>更多>>></a></div></td></tr>
 <%end If%>
 </table>
<%Call CloseRs(rs)

end function
%>


<%'信息目录＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
sub classlist()%>
<TABLE width="980"  border="0" align="center" cellpadding="0" cellspacing="0">
  <TR>
    <TD><IMG src="img/news_0034.gif" width="980" height="35"></TD>
  </TR>
</TABLE>
<table width="980"  border="0" align="center" cellpadding="0" cellspacing="2" bgcolor="0AAEDF">
  <TR>
    <TD bgcolor="#FFFFFF"><TABLE width="100%" border="0" cellspacing="8" cellpadding="0">
      <tr>
        <td valign='top' width="30%">
          <P style="line-height: 180%">
            <%Dim rsi,i
			i=0
			Call OpenConn
			set rs=server.CreateObject("adodb.recordset")
			rs.Open "select * from DNJY_type where twoid=0 and threeid=0 order by indexid",conn,1,1
			if rs.eof Then
Response.Write "<p><br>暂无"
IF c1>0 Then Response.Write conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
Response.Write "信息目录!"
else
			do while not rs.eof %>
      <TABLE width="100%" border="0" cellspacing="8" cellpadding="0"><TR><TD  width="5%"><IMG src="img/news_0035.gif">
                  <TABLE  border="0" cellpadding="0" cellspacing="0" align="center" id="table17">
                    <TR>
                      <TD height="8"></TD>
                    </TR>
                  </TABLE></TD><TD><A target="_blank" href="list.asp?t=<%=rs("id")%>,<%=rs("twoid")%>,<%=rs("threeid")%>"><FONT style="font-size:10.5pt" color="#1835D1"><B><%=rs("name")%></B></FONT></A></TD></TR>
            <%set rsi=server.CreateObject("adodb.recordset")
			set rsi=conn.execute("select * from DNJY_type where id="&rs("id")&" and twoid<>0 and threeid=0 order by  indexid,id,twoid")
			response.write "<tr><td></td><td>"
			do while not rsi.eof %>
            <A target="_blank" href="list.asp?t=<%=rsi("id")%>,<%=rsi("twoid")%>,<%=rsi("threeid")%>"> <FONT style="font-size:9pt" color="#1835D1"><%=rsi("name")%></FONT></A><FONT color="#CCCCCC">|</FONT>
            <%rsi.movenext
		   loop
		   response.write "</td></tr></table>"
		   i=i+1
		   if i mod 2=0 then
		   response.Write "</td></tr><tr><td valign='top' width='30%'>"
		   else
		   response.Write "</td><td valign='top' width='30%'>"
		   end if
		   rsi.close
		   set rsi=nothing
		   rs.movenext
		   Loop
		   end if
Call CloseRs(rs)
 
		   %>
            </span></td>
      </tr>
    </TABLE>
  
</TABLE>
      <TABLE  border="0" cellpadding="0" cellspacing="0" align="center" id="table17">
        <TR>
          <TD height="8"></TD>
        </TR>
      </TABLE>
      <TITLE></TITLE>

<%end sub%>

<%sub XinxiData()
id=strint(id)
Call OpenConn
set rs=server.createobject("adodb.recordset")
sql="select * from DNJY_info where yz=1"
if jiaoyi=0 then
sql=sql&" and zt="&jiaoyi
end If
if overdue=0 Then
sql=sql&" and dqsj>"&DN_Now&""
end If
sql=sql&" and id="&id
rs.open sql,conn,1,3
if rs.eof Then
Response.Write ("<script language=javascript>alert('操作成功，但此信息已过期，或已经删除，或未通过审核，无法生成htm文件!');history.go(-1);</script>")
Call CloseRs(rs)
response.end
end If
CT_ID=CT_ID
username=rs("username")
leixin=rs("leixing")
biaoti=rs("biaoti")
fbsj=rs("fbsj")
userip=rs("ip")
memo=rs("memo")
hfcs=rs("hfcs")
dianhua=rs("dianhua")
mobile=rs("CompPhone")
userqq=rs("qq")
email=rs("email")
xxmpname=rs("mpname")
dizhi=rs("dizhi")
youbian=rs("youbian")
wcolor=rs("wcolor")
ncolor=rs("ncolor")
c1=rs("city_oneid")
c2=rs("city_twoid")
c3=rs("city_threeid")
Readinfo=rs("Readinfo")
scjiage=check_jiage(rs("scjiage"))
jiage=check_jiage(rs("jiage"))
if rs("tj")>0 Then
biaotixs="<font color=""#FF0000""><img border=""0"" src=""images/jian.gif"" alt=""推荐信息""></font>"
end If
if rs("a")="0" Then
biaotixs=biaotixs&""&rs("biaoti")&"</b>"
Else
biaotixs=biaotixs&"<font color=""#"&rs("a")&"""><b>"&rs("biaoti")&"</b></font>"
end If
biaotixs=biaotixs&"</font>"

If rs("username")<>"" And rs("username")<>"游客" then
dim rs6,sql6,vip,sdname,sdkg,sdfg,mpname,mpkg,mpfg
Set rs6= Server.CreateObject("ADODB.Recordset")
sql6="select vip,name,sdname,sdfg,sdkg,mpname,mpfg,mpkg from [DNJY_user] where username='"&rs("username")&"'"
rs6.open sql6,conn,1,1
if Not rs6.eof Then
vip=rs6("vip")
sdname=rs6("sdname")
sdkg=rs6("sdkg")
mpname=rs6("mpname")
mpkg=rs6("mpkg")
mpfg=rs6("mpfg")
end If
rs6.close
set rs6=Nothing
end If

if sdkg=1 And sdname<>"" Then
sdmba="<a href=""user/com.asp?com="&rs("username")&""" target=""_blank""><img src=""img/dian.gif"" title=""此会员已开有店铺"" ></a>"
end If
if mpkg=1 And mpname<>"" then
sdmba=sdmba+" <a href=""company.asp?username="&rs("username")&""" target=""_blank""><img src=""img/qy.gif"" title=""此会员已开有企业名片"" ></a>"
end If
if (sdkg=0 And  mpkg=0) Or (sdkg="" And  mpkg="") Then
sdmba="无店铺和企业名片或未经审核"
end If


If IsNull(rs("username")) or rs("username")="游客" Then
usernameid="游客"
Else
usernameid="<a href=preview.asp?username="&rs("username")&"&id="&id&"&Readinfo="&Readinfo&">"&rs("username")&" 阅其信息</a> <INPUT TYPE=button VALUE=""评价"" ONCLICK=""window.open('user/evaluation.asp?username="&rs("username")&"','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbarsper=yes,scrollbars=yes,resizable=no,copyhistory=no,width=400,height=300,left=300,top=100')"" style=""color: #0060C5; background-color: #E1E2DC"">"
End if


if vip=1 Then
namea="<img src=""images/vip.gif"" alt=""验证VIP用户"" width=""13"" height=""13"" border=""0"">&nbsp;"
end If
namea=namea&""&rs("name")&""

If Filterhtm=1 And usernameid="游客" Then
memo=HTMLEncode(memo)'过滤HTM代码
ElseIf Filterhtm=2 And vip=0 Then
memo=HTMLEncode(memo)'过滤HTM代码
ElseIf Filterhtm>2 Then
memo=HTMLEncode(memo)'过滤HTM代码
else
memo=HTMLDecode(memo)'还原HTM代码
End if

if rs("c")=0 And shows8=1 Then
xxtp="<script src=""ads/ad.asp?adid=info4&c="&c1&","&c2&","&c3&"""></script>"
ElseIf rs("c")=0 And shows8=0 Then
xxtp="<script src=""ads/ad.asp?adid=info4&c=0,0,0""></script>"
else
if rs("c")=1 and not rs("tupian")="0" then
xxtp="<a target=""_blank"" title=""点击放大&gt;&gt;&gt;"" href=""UploadFiles/infopic/"&rs("tupian")&"""><img src=""UploadFiles/infopic/"&rs("tupian")&""" alt=""点击放大"" width=""250"" height=""200"" border=""0""><br><center>[信息图片 点击放大]</a>"
end if
end if


	belongtype="<a href=""list.asp?t="&rs("type_oneid")&",0,0&"&C_ID&""">"&rs("type_one")&"</a>"
	IF rs("type_two")<>"" and not isnull(rs("type_two")) Then belongtype=belongtype&" / <a href=""list.asp?t="&rs("type_oneid")&","&rs("type_twoid")&",0&"&C_ID&""">"&rs("type_two")&"</a>"
	IF rs("type_three")<>"" and not isnull(rs("type_three")) Then belongtype=belongtype&" / <a href=""list.asp?t="&rs("type_oneid")&","&rs("type_twoid")&","&rs("type_threeid")&"&"&C_ID&""">"&rs("type_three")&"</a>"	


    belongcity="<a href=""index.asp?c="&rs("city_oneid")&",0,0"">"&rs("city_one")&"</a>"
    IF rs("city_two")<>"" and not isnull(rs("city_two")) Then belongcity=belongcity&" / <a href=""index.asp?c="&rs("city_oneid")&","&rs("city_twoid")&",0"">"&rs("city_two")&"</a>"
	IF rs("city_three")<>"" and not isnull(rs("city_three")) Then belongcity=belongcity&" / <a href=""index.asp?c="&rs("city_oneid")&","&rs("city_twoid")&","&rs("city_threeid")&""">"&rs("city_three")&"</a>"

If request.cookies("dick")("username")<>rs("username") Then'自己发布的全权阅读
Dim rsvip,sqlvip,vipuser
Set rsvip= Server.CreateObject("ADODB.Recordset")
sqlvip="select vip from [DNJY_user] where username='"&request.cookies("dick")("username")&"'"
rsvip.open sqlvip,conn,1,1
if Not rsvip.eof Then vipuser=rsvip("vip")
rsvip.close
set rsvip=Nothing
If Readinfo=1 And request.cookies("dick")("username")="" Then
dianhua="普通会员权限阅读"
mobile="普通会员权限阅读"
email="普通会员权限阅读"
userqq="普通会员权限阅读"
youbian="普通会员权限阅读"
dizhi="普通会员权限阅读"
contentss="......>>>普通会员权限阅读"
If content_length>0 Then
memo=CutStr(memo,content_length,contentss)
else
memo="普通会员权限阅读"
End if
End If
cksj=""
If Readinfo=2 And vipuser<1 Then
dianhua="VIP会员权限阅读"
mobile="VIP会员权限阅读"
email="VIP会员权限阅读"
userqq="VIP会员权限阅读"
youbian="VIP会员权限阅读"
dizhi="VIP会员权限阅读"
contentss="......>>>VIP会员权限阅读"
cksj="<a href=""#"" ONCLICK=window.open('user/user_ckkf.asp?username="&request.cookies("dick")("username")&"&id="&id&"&key=cs','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=350,height=300,left=300,top=20')>&nbsp;<FONT color=""0000ff"">耗"&jf_ck&"点积分查手机号码</font></a>"
If content_length>0 Then
memo=CutStr(memo,content_length,contentss)
else
memo="VIP会员权限阅读"
End if
End If
End If
Call CloseRs(rs)
end sub


'信息类型，数据库中修改★★
sub leixs()
Dim rslx,sqllx,leixing,x
Call OpenConn
set rslx=server.createobject("adodb.recordset")
sqllx = "select leixing from DNJY_config "
rslx.open sqllx,conn,1,1
leixing=split(rslx("leixing"),"|")
response.write "<select name=""leixing""><option value=>类型</option>"
for x=0 to ubound(leixing)
response.write "<option value="""&leixing(x)&""">"&leixing(x)&"</option>"
next
response.write "</select>"
rslx.close
set rslx=Nothing
end sub


'信息显示（标题式）最新、1竞价、2推荐、3热门＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function f_info(xxsx,tt1,tt2,tt3,c1,c2,c3,xxlx,gd,xxtj,tp,jj,gdxx,ys,xj,hits,xxsj,btw,s,l,zt,hg)
               
'信息标题显示标签说明。依次：xxsx信息属性（1最新、2竞价、3推荐、4热门），tt1一级信息分类（0不分、其它按该类ID号），tt2二级信息分类（0不分、其它按该类ID号），tt3三级信息分类（0不分、其它按该类ID号），c1一级地区分类（0不分、其它按该类ID号），c2二级地区分类（0不分、其它按该类ID号），c3三级地区分类（0不分、其它按该类ID号），xxlx信息类型（0不显1显），gd启用固顶（0不固1固），xxtj推荐标志（0不显1显），tp图片标志（0不显1显），jj竞价提示（0不显1显），gdxx更多（0不显1显），ys标题颜色（0不显1显），xj信笺背景（0不显1显），hits点击数（0不显1显），xxsj时间（0-15种格式，0不显），btw标题长度（以字节计），s显示条数，l列数，zt字号，hg行高
dim t,TempText,g,ThisPage,Pagesize,Allrecord,Allpage,leixing,iii,i,xt,b,tj,bb,a,c,biaoti,sj2,gdxxx,yss,b2
t=0
If xj=1 Then xj="style=""BACKGROUND-COLOR: ffffff""   height=""26"" background=""img/22.gif"""
gdxxx=""
If gdxx=1 And xxsx=0 Then gdxxx="&xxsx=0"
If gdxx=1 And xxsx=1 Then gdxxx="&xxsx=1"
If gdxx=1 And xxsx=2 Then gdxxx="&xxsx=2"
If gdxx=1 And xxsx=3 Then gdxxx="&xxsx=3"
If gdxx=1 And xxsx=4 Then gdxxx="&xxsx=4"
If gdxx=1 And xxsx=5 Then gdxxx="&xxsx=5"

    Dim tttt
	IF tt1=0 then
	tttt=""
	elseIF tt3>0 then
	tttt=" and type_oneid="&tt1&" and type_twoid="&tt2&" and type_threeid="&tt3
	elseIF tt2>0 then
	tttt=" and type_oneid="&tt1&" and type_twoid="&tt2
	Else
	tttt=" and type_oneid="&tt1
	End If
Call OpenConn
set rs=server.createobject("ADODB.recordset")
sql = "select * from DNJY_info where yz=1"
if jiaoyi=0 Then sql=sql&" and zt="&jiaoyi
if overdue=0 Then sql=sql&" and dqsj>"&DN_Now&""
sql=sql&""&ttcc
sql=sql&""&tttt
if xxsx=2 Then sql=sql&" and hfxx=1 and hfje>0 "
if xxsx=3 Then sql=sql&" and tj>0"
if xxsx=4 Then sql=sql&" and llcs>="&hotbz&""
'sql=sql&" and leixing<>'出售'"'车网，启用此行

sql=sql&" order by "
if gd=1 Then sql=sql&"b "&DN_OrderType&","'加置顶优先条件
if xxsx=2 Then
sql=sql&"hfje/fbts "&DN_OrderType&"," 
elseif xxsx=4 Then
sql=sql&"llcs "&DN_OrderType&","'车网，注释掉这行
end If
sql=sql&"fbsj "&DN_OrderType&"" 


Select Case xxsx
	Case 1
	b2="最新"
	Case 2
	b2="竞价"
	Case 3	
	b2="推荐"
	Case 4	
	b2="热门"
	Case 5	
	b2="回复"
	Case Else	
	b2=""
	End Select

rs.open sql,conn,1,1
if rs.eof Then
TempText="暂无"
IF c1>0 Then TempText=TempText&conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
TempText=TempText&b2&"信息"
else

iii=0
TempText="<table width=""100%""  border=""0"" cellspacing=""2"" cellpadding=""0"" >"
do while not rs.eof
i=i+1
b=trim(rs("b"))
tj=trim(rs("tj"))
bb=len(b)

IF iii Mod l=0 Then TempText=TempText&"<tr height="""&hg&""">"
If xxlx=1 then
TempText=TempText&"<td "&xj&">[<a target=""_blank"" href='list.asp?"&CT_ID&"&leixing="&rs("leixing")&"&xxsx="&xxsx&"'>"&rs("leixing")&"</a>]</td>"
End if
TempText=TempText&"<td "&xj&">"
if rs("Readinfo")=1 Then TempText=TempText&"<img src=""img/lock.jpg"" alt=""普通会员权限浏览"">" 
if rs("Readinfo")=2 Then TempText=TempText&"<img src=""img/lockvip.jpg"" alt=""VIP会员权限浏览"">" 
if rs("tupian")<>"0" And tp=1 Then TempText=TempText&"<img src=""images/num/pic.gif"" alt=""有图片"">" 
TempText=TempText&"<a target=""_blank"" href="""&x_path("",rs("id"),"",rs("Readinfo"))&""" title="&rs("biaoti")&"||"&rs("name")&"发布于"&FormatDateTime(rs("fbsj"),vbShortDate)&">"
if rs("a")<>"0" And ys=1 Then
yss=rs("a")
Else
yss="000000"
End if
TempText=TempText&"<SPAN style=""color:#"&yss&";font-size="&zt&"px"">"&CutStr(rs("biaoti"),btw,"..")&"</SPAN>"
TempText=TempText&"</a>"
if b<>0 And gd=1 then
TempText=TempText&"<img src=""images/num/jsq.gif"" alt=""置顶"&b&"级"">"
for g=1 to bb
TempText=TempText&"<img src=""images/num/"&Mid(b,g,1)&".gif"" alt=""置顶"&b&"级"">"
Next
end If
if tj<>0 And xxtj=1 Then TempText=TempText&"<img src=""images/jian.gif"" width=""15"" height=""15"" alt=""本站推荐"">"
if rs("hfxx")=1 And rs("hfje")>0 And rs("dqsj")>now() And jj=1 Then TempText=TempText&"<span title=竞价信息平均每天"&FormatNumber(rs("hfje")/rs("fbts"),2,-1)&"元 style=""cursor:help""><img src=""img/jinja.gif"" width=""16"" height=""16"" ></span>("&FormatNumber(rs("hfje")/rs("fbts"),2,-1)&")"  
IF hits=1 Then TempText=TempText&  " <font color=""#999999"" ><span title=此信息有"&rs("llcs")&"人关注过 style=""cursor:help"">["&rs("llcs")&"]</span></font>" 
If xxsj>0 Then TempText=TempText&"<td "&xj&">"
TempText=TempText&  " <font color=""#999999"" >"&FormatDate(rs("fbsj"),xxsj)&"</font>" 
If xxsj>0 Then TempText=TempText&  "</td>"
TempText=TempText&"</TD>"
	iii=iii+1
	IF iii Mod l=0 then TempText=TempText&  "</tr>"
    if i>=s then exit do 
	Rs.movenext
	Loop	
	IF iii Mod l<>0 then
	Do While not iii Mod l=0
	TempText=TempText&  "<td></td>"
	iii=iii+1
	Loop
	TempText=TempText&  "</tr>"
	End IF
	TempText=TempText&  "</table>"
	If gdxx=1 Then TempText=TempText&  "<table align=""right""><tr><td><A href=""list.asp?"&CT_ID&""&gdxxx&""">更多>>></a></td></tr></table>"
end if
Call CloseRs(rs)

f_info=TempText
end Function

'信息显示（大图文式）最新、1竞价、2推荐、3热门＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function tw_info(xxsx,tt1,tt2,tt3,c1,c2,c3,xxlx,gd,xxtj,tp,hits,xxsj,scj,bzj,btw,ys,zt,tw,th,dw,dh,sl,l,fy)
'信息标题显示标签说明。依次：xxsx信息属性（1最新、2竞价、3推荐、4热门），tt1一级信息分类（0不分、其它按该类ID号），tt2二级信息分类（0不分、其它按该类ID号），tt3三级信息分类（0不分、其它按该类ID号），c1一级地区分类（0不分、其它按该类ID号），c2二级地区分类（0不分、其它按该类ID号），c3三级地区分类（0不分、其它按该类ID号），xxlx信息交易类型（0不显1显），gd启用固顶（0不固1固），xxtj推荐标志（0不显1显），tp图片标志（0不显1显），hits点击数（0不显1显），xxsj时间（0-15种格式，0不显），scj市场价（0不显1显）,bzj本站价（0不显1显），btw标题长度（以字节计）,ys标题颜色（0不显1显），zt字号，tw图片宽度,th图片高度,dw单元宽度,dh单元高度,sl每页数量,l列数，fy分页（0不显1显）

dim t,TempText,i,gdxxx,gdxx,b2
t=0
gdxxx=""
If gdxx=1 And xxsx=0 Then gdxxx="&xxsx=0"
If gdxx=1 And xxsx=1 Then gdxxx="&xxsx=1"
If gdxx=1 And xxsx=2 Then gdxxx="&xxsx=2"
If gdxx=1 And xxsx=3 Then gdxxx="&xxsx=3"
If gdxx=1 And xxsx=4 Then gdxxx="&xxsx=4"
If gdxx=1 And xxsx=5 Then gdxxx="&xxsx=5"

    Dim tttt
	IF tt1=0 then
	tttt=""
	elseIF tt3>0 then
	tttt=" and type_oneid="&tt1&" and type_twoid="&tt2&" and type_threeid="&tt3
	elseIF tt2>0 then
	tttt=" and type_oneid="&tt1&" and type_twoid="&tt2
	Else
	tttt=" and type_oneid="&tt1
	End If
Call OpenConn

dim totalPut,CurrentPage,j,MaxPerPage
 Tmpstr = Request.ServerVariables("PATH_INFO")
Tmpstr = Split(Tmpstr,"/")
ScriptName = Lcase(Tmpstr(UBound(Tmpstr)))

				MaxPerPage=sl '每页显示个数   				
    			if Not isempty(SafeRequest("page",1)) then
      				currentPage=Cint(SafeRequest("page",1))
   				else
      				currentPage=1
   				end if 

set rs=server.createobject("ADODB.recordset")
sql = "select * from DNJY_info where yz=1"
if jiaoyi=0 Then sql=sql&" and zt="&jiaoyi
if overdue=0 Then sql=sql&" and dqsj>"&DN_Now&""
sql=sql&""&ttcc
sql=sql&""&tttt
if xxsx=2 Then sql=sql&" and hfxx=1 and hfje>0 "
if xxsx=3 Then sql=sql&" and tj>0"
if xxsx=4 Then sql=sql&" and llcs>="&hotbz&""
leixing=request("leixing")
lx=request("leixing")
If leixing<>"" Then sql=sql&" And leixing='"&leixing&"'"
sql=sql&" order by "
if gd=1 Then sql=sql&"b "&DN_OrderType&","'加置顶优先条件
if xxsx=2 Then
sql=sql&"hfje/fbts "&DN_OrderType&"," 
elseif xxsx=4 Then
sql=sql&"llcs "&DN_OrderType&","
end If
sql=sql&"fbsj "&DN_OrderType&"" 

Select Case xxsx
	Case 1
	b2="最新"
	Case 2
	b2="竞价"
	Case 3	
	b2="推荐"
	Case 4	
	b2="热门"
	Case 5	
	b2="回复"
	Case Else	
	b2=""
	End Select
rs.open sql,conn,1,1
if rs.eof Then
TempText="<TR><TD>暂无"
IF c1>0 Then TempText=TempText&conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
TempText=TempText&b2&lx&"信息</TD></TR>"
else'正式显示开始

	  				totalPut=rs.recordcount
      				if currentpage<1 then
          				currentpage=1
      				end if
      				if (currentpage-1)*MaxPerPage>totalput then
	   					if (totalPut mod MaxPerPage)=0 then
	     					currentpage= totalPut \ MaxPerPage
	   					else
	      					currentpage= totalPut \ MaxPerPage + 1
	   					end if
      				end if
       				if currentPage=1 Then
'------------------------------------------------
				dim k
	   			i=0
                k=0
TempText=TempText&"<table width=100% border=0 cellspacing=1 cellpadding=1 align=center><tr>"
do while not rs.eof        
TempText=TempText&"<td width="&dw&"><TABLE WIDTH="&dw&" HEIGHT="&dh&" BORDER=0 CELLPADDING=0 CELLSPACING=0 align=center>"
TempText=TempText&"<TR><TD align=center ><a href="&x_path("",rs("id"),"",rs("Readinfo"))&" target=_blank><img style="" border:2 double #E2E2E2"" border=1  WIDTH="&tw&" HEIGHT="&th&" onload='DrawImage(this)' src="
if rs("c")=1 and not rs("tupian")="0" Then
TempText=TempText&"'UploadFiles/infopic/"&rs("tupian")&"'"
Else
TempText=TempText&"images/nopicture.jpg"
end If
TempText=TempText&"></a></TD></TR>"
TempText=TempText&"<TR><TD valign=""top"" bgcolor=""#FFFFFF"">"
IF rs("b")>0 And gd=1 Then
TempText=TempText&"<img border=0 src=img/top4.gif width=""15"" height=""13"" alt=""固顶"">"
End If
IF rs("tj")>0 And xxtj=1 Then
TempText=TempText&"<img border=0 src=img/jian.gif width=""15"" height=""13"" alt=""推荐"">"
End if
IF rs("tupian")<>"0" And tp=1 Then
TempText=TempText& "<img src=""images/num/pic.gif"" width=""13"" height=""13"" border=0 alt=""图片"">"
End if
TempText=TempText& "<a href="""&x_path("",rs("id"),"",rs("Readinfo"))&""" target=_blank title="&rs("biaoti")&"""><SPAN style=""color:#"&ys&";font-size="&zt&"px"">"&CutStr(trim(rs("biaoti")),btw,"..")&"</SPAN></a></TD></TR>"
IF xxlx=1 Then TempText=TempText&"<TR><TD valign=""top"" bgcolor=""#FFFFFF"">交易类型："&rs("leixing")&"</TD></TR>"
IF scj=1 Then TempText=TempText&"<TR><TD valign=""top"" bgcolor=""#FFFFFF"">市场价格：<s>￥"&check_jiage(rs("scjiage"))&"</s></TD></TR>"
IF bzj=1 Then TempText=TempText&"<TR><TD valign=""top"" bgcolor=""#FFFFFF"">本站价格：￥"&check_jiage(rs("jiage"))&"</TD></TR>"
IF hits=1 Then TempText=TempText&"<TR><TD valign=""top"" bgcolor=""#FFFFFF"">点击次数："&rs("llcs")&"</TD></TR>"
IF xxsj>0 Then TempText=TempText&"<TR><TD valign=""top"" bgcolor=""#FFFFFF"">发布日期："&FormatDate(rs("fbsj"),xxsj)&"</TD></TR>"
TempText=TempText&"</TABLE></td>"
i=i+1	
	k=k+1
    If k=l then
	TempText=TempText&"</tr>" 'L个后换行
	k=0
	End If
	if i>=MaxPerPage then exit Do'每页显示个数，与上面MaxPerPage设定参数
    rs.movenext
    loop
TempText=TempText&"</tr></table></td></tr>" 
If fy=1 then
TempText=TempText&"<tr><td colspan=""4"">"
Dim n,totalnumber,filename
totalnumber=totalput
maxperpage=MaxPerPage
filename=""&ScriptName&""
				If totalnumber Mod maxperpage=0 Then  
					n= totalnumber \ maxperpage  
				Else
					n= totalnumber \ maxperpage+1  
				End If 
            TempText=TempText&"<form method=Post action= '"& filename &"?"&CT_ID&"&xxsx="&xxsx&"&leixing="&lx&"'><p align='center'>"               
				If CurrentPage<2 Then  
              TempText=TempText&"首页 上页 "                     
				Else   
         TempText=TempText&"<a href='"&filename&"?page=1&"&CT_ID&"&xxsx="&xxsx&"&leixing="&lx&"'>首页</a> " 
         TempText=TempText&"<a href='"& filename &"?page="& CurrentPage-1&"&"&CT_ID&"&xxsx="&xxsx&"&leixing="&lx&"'>上页</a> "             
				End If	
				If n-currentpage<1 Then  
            TempText=TempText&"下页 末页 "                
				Else   
      TempText=TempText&"<a href='"&filename&"?page="&(CurrentPage+1)&"&"&CT_ID&"&xxsx="&xxsx&"&leixing="&lx&"'>下页</a> " 
      TempText=TempText&"<a href='"&filename&"?page="&n&"&"&CT_ID&"&xxsx="&xxsx&"&leixing="&lx&"'>末页</a> "                
				End If 
      TempText=TempText&"第 <b>"& CurrentPage&"</b> 页 共 <b>"& n&"</b> 页 共&nbsp;<b>"& totalnumber&"</b>&nbsp;件商品　每页 <b>"& maxperpage&" </b> 件商品 转到第：<input type='text' name='page' size=2 maxlength=10 class=smallInput value='"&currentpage&"'>页<input type='submit'  class='contents' value='跳转' name='cndok'></form>"
TempText=TempText&"</td></tr>"  
End If
'-------------------------------                    
       				else
          				if (currentPage-1)*MaxPerPage<totalPut then
            				rs.move  (currentPage-1)*MaxPerPage
            				dim bookmark
            				bookmark=rs.bookmark
'------------------------------				
	   			i=0
                k=0
TempText=TempText&"<table width=100% border=0 cellspacing=1 cellpadding=1 align=center><tr>"
do while not rs.eof        
TempText=TempText&"<td width="&dw&"><TABLE WIDTH="&dw&" HEIGHT="&dh&" BORDER=0 CELLPADDING=0 CELLSPACING=0 align=center>"
TempText=TempText&"<TR><TD align=center ><a href="&x_path("",rs("id"),"",rs("Readinfo"))&") target=_blank><img style="" border:2 double #E2E2E2"" border=1  WIDTH="&tw&" HEIGHT="&th&" onload='DrawImage(this)' src="
if rs("c")=1 and not rs("tupian")="0" Then
TempText=TempText&"'UploadFiles/infopic/"&rs("tupian")&"'"
Else
TempText=TempText&"images/nopicture.jpg"
end If
TempText=TempText&"></a></TD></TR>"
TempText=TempText&"<TR><TD valign=""top"" bgcolor=""#FFFFFF"">"
IF rs("b")>0 And gd=1 Then
TempText=TempText&"<img border=0 src=img/top4.gif width=""15"" height=""13"" alt=""固顶"">"
End If
IF rs("tj")>0 And xxtj=1 Then
TempText=TempText&"<img border=0 src=img/jian.gif width=""15"" height=""13"" alt=""推荐"">"
End if
IF rs("tupian")<>"0" And tp=1 Then
TempText=TempText& "<img src=""images/num/pic.gif"" width=""13"" height=""13"" border=0 alt=""图片"">"
End if
TempText=TempText& "<a href="""&x_path("",rs("id"),"",rs("Readinfo"))&""" target=_blank title="&rs("biaoti")&"""><SPAN style=""color:#"&ys&";font-size="&zt&"px"">"&CutStr(trim(rs("biaoti")),btw,"..")&"</SPAN></a></TD></TR>"
IF xxlx=1 Then TempText=TempText&"<TR><TD valign=""top"" bgcolor=""#FFFFFF"">交易类型："&rs("leixing")&"</TD></TR>"
IF scj=1 Then TempText=TempText&"<TR><TD valign=""top"" bgcolor=""#FFFFFF"">市场价格：<s>￥"&check_jiage(rs("scjiage"))&"</s></TD></TR>"
IF bzj=1 Then TempText=TempText&"<TR><TD valign=""top"" bgcolor=""#FFFFFF"">本站价格：￥<b><font color='#ff0000'>"&check_jiage(rs("jiage"))&"</font></b></TD></TR>"
IF hits=1 Then TempText=TempText&"<TR><TD valign=""top"" bgcolor=""#FFFFFF"">点击次数："&rs("llcs")&"</TD></TR>"
IF xxsj>0 Then TempText=TempText&"<TR><TD valign=""top"" bgcolor=""#FFFFFF"">发布日期："&FormatDate(rs("fbsj"),xxsj)&"</TD></TR>"
TempText=TempText&"</TABLE></td>"
i=i+1	
	k=k+1
    If k=l then
	TempText=TempText&"</tr>" 'L个后换行
	k=0
	End If
	if i>=MaxPerPage then exit Do'每页显示个数，与上面MaxPerPage设定参数
    rs.movenext
    loop
TempText=TempText&"</tr></table></td></tr>" 
If fy=1 then
TempText=TempText&"<tr><td colspan=""4"">"
totalnumber=totalput
maxperpage=MaxPerPage
filename=""&ScriptName&""
				If totalnumber Mod maxperpage=0 Then  
					n= totalnumber \ maxperpage  
				Else
					n= totalnumber \ maxperpage+1  
				End If 
            TempText=TempText&"<form method=Post action=' "& filename &"?"&CT_ID&"&xxsx="&xxsx&"&leixing="&lx&"'><p align='center'>"               
				If CurrentPage<2 Then  
              TempText=TempText&"首页 上页 "                     
				Else   
         TempText=TempText&"<a href='"&filename&"?page=1&"&CT_ID&"&xxsx="&xxsx&"&leixing="&lx&"'>首页</a> " 
         TempText=TempText&"<a href='"& filename &"?page="& CurrentPage-1&"&"&CT_ID&"&xxsx="&xxsx&"&leixing="&lx&"'>上页</a> "             
				End If	
				If n-currentpage<1 Then  
            TempText=TempText&"下页 末页 "                
				Else   
      TempText=TempText&"<a href='"&filename&"?page="&(CurrentPage+1)&"&"&CT_ID&"&xxsx="&xxsx&"&leixing="&lx&"'>下页</a> " 
      TempText=TempText&"<a href='"&filename&"?page="&n&"&"&CT_ID&"&xxsx="&xxsx&"&leixing="&lx&"'>末页</a> "                
				End If 
      TempText=TempText&"第 <b>"& CurrentPage&"</b> 页 共 <b>"& n&"</b> 页 共&nbsp;<b>"& totalnumber&"</b>&nbsp;件商品　每页 <b>"& maxperpage&" </b> 件商品 转到第：<input type='text' name='page' size=2 maxlength=10 class=smallInput value='"&currentpage&"'>页<input type='submit'  class='contents' value='跳转' name='cndok'></form>"
TempText=TempText&"</td></tr>"
End If
'-------------------------------             				
        				else
	        				currentPage=1
'-------------------------------
	   			i=0
                k=0
TempText=TempText&"<table width=100% border=0 cellspacing=1 cellpadding=1 align=center><tr>"
do while not rs.eof        
TempText=TempText&"<td width="&dw&"><TABLE WIDTH="&dw&" HEIGHT="&dh&" BORDER=0 CELLPADDING=0 CELLSPACING=0 align=center>"
TempText=TempText&"<TR><TD align=center ><a href="&x_path("",rs("id"),"",rs("Readinfo"))&" target=_blank><img style="" border:2 double #E2E2E2"" border=1  WIDTH="&tw&" HEIGHT="&th&" onload='DrawImage(this)' src="
if rs("c")=1 and not rs("tupian")="0" Then
TempText=TempText&"'UploadFiles/infopic/"&rs("tupian")&"'"
Else
TempText=TempText&"images/nopicture.jpg"
end If
TempText=TempText&"></a></TD></TR>"
TempText=TempText&"<TR><TD valign=""top"" bgcolor=""#FFFFFF"">"
IF rs("b")>0 And gd=1 Then
TempText=TempText&"<img border=0 src=img/top4.gif width=""15"" height=""13"" alt=""固顶"">"
End If
IF rs("tj")>0 And xxtj=1 Then
TempText=TempText&"<img border=0 src=img/jian.gif width=""15"" height=""13"" alt=""推荐"">"
End if
IF rs("tupian")<>"0" And tp=1 Then
TempText=TempText& "<img src=""images/num/pic.gif"" width=""13"" height=""13"" border=0 alt=""图片"">"
End if
TempText=TempText& "<a href="""&x_path("",rs("id"),"",rs("Readinfo"))&" target=_blank title="&rs("biaoti")&"""><SPAN style=""color:#"&ys&";font-size="&zt&"px"">"&CutStr(trim(rs("biaoti")),btw,"..")&"</SPAN></a></TD></TR>"
IF xxlx=1 Then TempText=TempText&"<TR><TD valign=""top"" bgcolor=""#FFFFFF"">交易类型："&rs("leixing")&"</TD></TR>"
IF scj=1 Then TempText=TempText&"<TR><TD valign=""top"" bgcolor=""#FFFFFF"">市场价格：<s>￥"&check_jiage(rs("scjiage"))&"</s></TD></TR>"
IF bzj=1 Then TempText=TempText&"<TR><TD valign=""top"" bgcolor=""#FFFFFF"">本站价格：￥"&check_jiage(rs("jiage"))&"</TD></TR>"
IF hits=1 Then TempText=TempText&"<TR><TD valign=""top"" bgcolor=""#FFFFFF"">点击次数："&rs("llcs")&"</TD></TR>"
IF xxsj>0 Then TempText=TempText&"<TR><TD valign=""top"" bgcolor=""#FFFFFF"">发布日期："&FormatDate(rs("fbsj"),xxsj)&"</TD></TR>"
TempText=TempText&"</TABLE></td>"
i=i+1	
	k=k+1
    If k=l then
	TempText=TempText&"</tr>" 'L个后换行
	k=0
	End If
	if i>=MaxPerPage then exit Do'每页显示个数，与上面MaxPerPage设定参数
    rs.movenext
    loop
TempText=TempText&"</tr></table></td></tr>"
If fy=1 then
TempText=TempText&"<tr><td colspan=""4"">"
totalnumber=totalput
maxperpage=MaxPerPage
filename=""&ScriptName&""
				If totalnumber Mod maxperpage=0 Then  
					n= totalnumber \ maxperpage  
				Else
					n= totalnumber \ maxperpage+1  
				End If 
            TempText=TempText&"<form method=Post action= '"& filename &"?"&CT_ID&"&xxsx="&xxsx&"&leixing="&lx&"'><p align='center'>"               
				If CurrentPage<2 Then  
              TempText=TempText&"首页 上页 "                     
				Else   
         TempText=TempText&"<a href='"&filename&"?page=1&"&CT_ID&"&xxsx="&xxsx&"&leixing="&lx&"'>首页</a> " 
         TempText=TempText&"<a href='"& filename &"?page="& CurrentPage-1&"&"&CT_ID&"&xxsx="&xxsx&"&leixing="&lx&"'>上页</a> "             
				End If	
				If n-currentpage<1 Then  
            TempText=TempText&"下页 末页 "                
				Else   
      TempText=TempText&"<a href='"&filename&"?page="&(CurrentPage+1)&"&"&CT_ID&"&xxsx="&xxsx&"&leixing="&lx&"'>下页</a> " 
      TempText=TempText&"<a href='"&filename&"?page="&n&"&"&CT_ID&"&xxsx="&xxsx&"&leixing="&lx&"'>末页</a> "                
				End If 
      TempText=TempText&"第 <b>"& CurrentPage&"</b> 页 共 <b>"& n&"</b> 页 共&nbsp;<b>"& totalnumber&"</b>&nbsp;件商品　每页 <b>"& maxperpage&" </b> 件商品 转到第：<input type='text' name='page' size=2 maxlength=10 class=smallInput value='"&currentpage&"'>页<input type='submit'  class='contents' value='跳转' name='cndok'></form>"
TempText=TempText&"</td></tr>"  
End If
'-------------------------------           					
	      				end if
	   				end if
   				end if'正式显示结束

tw_info=TempText
 Call CloseRs(rs)
end Function

'信息、地区分类目录列表====================================================
Function F_vassal_c_t(c1,c2,c3,t1,t2,t3,q,n,c,l,h,s)
'标签说明,依次为：一级地区，二级地区，三级地区，信息一级分类，信息二级分类，信息三级分类，标题圆点标志,显示所在分类信息条数（0不显1显),颜色（0不显1显),列数,行高,显示信息或地区分类（0地区1信息)
Dim rs,sql
Call OpenConn
 Set Rs=server.createobject("aDodb.recordset") 
	'If Not IsObject(Conn) or Conn Is NoThing Then dblink
	Dim ci,t,ii,cc,temptext
	IF s=1 Then
	IF t1=0 then exit Function
	IF t2>0 Then
	t="id="&t1&" and twoid="&t2&" and threeid<>0"
	Else
	t="id="&t1&" and threeid=0"
	End IF
	Else
	IF c2>0 Then
	ci="id="&c1&" and twoid="&c2&" and threeid<>0"
	ElseIF c1>0 Then
	ci="id="&c1&" and threeid=0"
	Else
	ci="twoid=0"
	End IF
	End IF
	

	ii=0
	temptext="<table align=center width=""100%"" border=0 cellpadding=0 cellspacing=0>"
	IF s=1 Then
	Set Rs=conn.execute("select twoid,threeid,[name],[titlecolor],id from DNJY_type where "&t&" and twoid>0 order by indexid")
	do while not rs.eof
	IF ii mod l=0 Then temptext=temptext&"<TR height="&h&" align=center>"
	temptext=temptext&"<td>"
	IF q=1 Then temptext=temptext&"·"
	IF c=1 Then temptext=temptext&"<font color="&rs(3)&">"
	temptext=temptext&"<a href=list.asp?c="&c1&","&c2&","&c3&"&t="&t1&","&rs(0)&","&rs(1)&"> "&rs(2)&" </a>"
	IF c=1 Then temptext=temptext&"</font>"
	IF n=1 Then
	IF t2=0 Then
	Sql="Select count(id) from DNJY_info Where type_oneid="&t1&" and type_twoid="&rs(0)&" "&ttcc&""
	Else
	Sql="Select count(id) from DNJY_info Where type_oneid="&t1&" and type_twoid="&t2&" and type_threeid="&rs(1)&" "&ttcc&""
	End IF
	temptext=temptext&"("&conn.execute(sql)(0)&")"
	End IF
	temptext=temptext&"</td>"
	Rs.movenext
	ii=ii+1
	IF ii mod l=0 Then temptext=temptext&"</TR>"
	loop
	Else
	Set Rs=conn.execute("select twoid,threeid,[city],id from DNJY_city where "&ci&" order by indexid")
	do while not rs.eof
	IF ii mod l=0 Then temptext=temptext&"<TR height="&h&" align=center>"
	temptext=temptext&"<td>"
	IF q=1 Then temptext=temptext&"·"
	temptext=temptext&"<a href=list.asp?c="&rs(3)&","&rs(0)&","&rs(1)&"&t="&t1&","&t2&","&t3&">"&rs(2)&"</a></td>"
	Rs.movenext
	ii=ii+1
	IF ii mod l=0 Then temptext=temptext&"</TR>"
	loop
	
	End IF
	Call CloseRs(rs)
    
	
	IF ii mod l<>0 Then
	do while ii mod l<>0
	temptext=temptext&"<TD>&nbsp;</TD>"
	ii=ii+1
	loop
	temptext=temptext&"</TR>"
	End IF
	F_vassal_c_t=TempText&"</table>"
End Function

'地区分类列表（调用一、二、三级分类）
Function c_more(ty1,by1,by2,by3,ts1,ts2,ts3,n1,n2,ls,zt1,zt2,zt3,hg1,hg2,hg3)
 '标签说明,依次为：ty1一级分类ID号（0为全部），by1是否显示一级分类（0不显1显)，by2是否显示二级分类（0不显1显)，by3是否显示三级分类（0不显1显)，ts1一级分类显示个数(0为全部），ts2二级分类显示个数（0为全部），ts3三级分类显示个数（0为全部）,n1显示一级分类属下分类个数（0不显1显),n2显示二级分类属下分类个数（0不显1显),ls1一级分类显示列数（0不分列)，zt1一级分类字号，zt2二级分类字号，zt3三级分类字号，hg1一级分类行距，hg2二级分类行距，hg3三级分类行距
 Dim temptext,rsi,rst3,ti,ts,tj,tt,idd,titlecolor,titlecolor2,lname,kg
    temptext="<TABLE border=""0"" cellspacing=""1"" cellpadding=""0"">"
      temptext=temptext&"<tr><td valign=""top"">"         
     idd=" and id="&ty1
	 If ty1=0 Then idd=""
	 Call OpenConn
			set rs=server.CreateObject("adodb.recordset")
			rs.Open "select * from DNJY_city where twoid=0 and threeid=0"&idd&" order by indexid",conn,1,1
if rs.eof Then
Response.Write "<font color=000000>暂无分类或未设置首页显示</font>"
else
			tj=0
			do while not rs.eof
	  titlecolor="1835D1"
      If rs(0)=t1 Then titlecolor="ff0000"
      temptext=temptext&"<TABLE border=""0"" cellspacing=""1"" cellpadding=""0"">"
	  temptext=temptext&"<TR height="""&hg1&"""><TD></TD><TD>"
	  If by1=1 Then
	  temptext=temptext&"<A target=""_blank"" href="""
	  If rs("dname")<>"" And shows1=1 Then
	  temptext=temptext&"http://"&rs("dname")&""
	  Else
	  temptext=temptext&"index.asp?c="&rs("id")&","&rs("twoid")&","&rs("threeid")&""
	  End If
	 temptext=temptext&"""><FONT style=""font-size:"&zt1&"pt"" color=""#1835D1""><B>"&rs("city")&"</B></FONT></A>"
	 If n1=1 then
	 Sql="Select count(i) from DNJY_city Where id="&rs("i")&" and twoid>0 and threeid=0"
	 temptext=temptext&"<FONT color=""ff0000""><span title=属下有"&conn.execute(sql)(0)&"个分类 style=""cursor:help "">("&conn.execute(sql)(0)&")</span></FONT>"
	 End IF 
	 End if
	  temptext=temptext&"</TD></TR>"	  
       set rsi=conn.execute("select * from DNJY_city where id="&rs("id")&" and twoid<>0 and threeid=0 order by  indexid,id,twoid")	  
	       If by2=1 then
			temptext=temptext&"<tr height="""&hg1&"""><td>&nbsp;&nbsp;</td><td  width="""&920/ls&""">"
			ts=0
			do while not rsi.eof
			titlecolor2="1835D1"
            If rsi("id")=t1 And rsi("twoid")=t2 Then titlecolor2="ff0000"
			temptext=temptext&"<A target=""_blank"" href="""
			If rsi("dname")<>"" And shows1=1 Then
			temptext=temptext&"http://"&rsi("dname")&""
			Else
			temptext=temptext&"index.asp?c="&rsi("id")&","&rsi("twoid")&","&rsi("threeid")&""
			End If
			temptext=temptext&"""><FONT style=""font-size:"&zt2&"pt"" color=""#1835D1"">"&rsi("city")&"</FONT></A>"
            If n2=1 then		    
	        Sql="Select count(i) from DNJY_city Where id="&rsi("id")&" and twoid="&rsi("twoid")&" and threeid>0"
	        temptext=temptext&"<FONT color=""ff0000""><span title=属下有"&conn.execute(sql)(0)&"个分类 style=""cursor:help "">("&conn.execute(sql)(0)&")</span></FONT>"
			End If

			If by3=0 then
		    temptext=temptext&"<FONT color=""#CCCCCC"">|</FONT>"
            Else
            temptext=temptext&"<br>"
            set rst3=conn.execute("select * from DNJY_city where id="&rs("id")&" and twoid="&rsi("twoid")&" and threeid<>0 order by  indexid,id,twoid,threeid")
			if Not rst3.eof Then
		    temptext=temptext&"<table><tr height="""&hg3&"""><td></td><td>└"
			tt=0			
			do while not rst3.eof
			temptext=temptext&"<A target=""_blank"" href="""
			If rst3("dname")<>"" And shows1=1 Then
			temptext=temptext&"http://"&rst3("dname")&""
			Else
			temptext=temptext&"index.asp?c="&rst3("id")&","&rst3("twoid")&","&rst3("threeid")&""
			End If
			temptext=temptext&"""><FONT style=""font-size:"&zt3&"pt"" color=""#9CA6B5"">"&rst3("city")&"</FONT></A>"
			temptext=temptext&"<FONT color=""#CCCCCC"">|</FONT>"
			tt=tt+1			
			if tt>=ts3 And ts3>0 Then
            temptext=temptext&"..."
			exit Do
			End if
			rst3.movenext
		   Loop
		  temptext=temptext&"</td></tr></table>"
		   End if
		   rst3.close
           End If
           ts=ts+1
		   if ts>=ts2 And ts2>0 Then
		   temptext=temptext&"..."
		   exit do 
           End If          
           rsi.movenext
		   Loop
		   End if
		   temptext=temptext&"</td></tr></table>"
		   ti=ti+1
		   if ti mod ls=0 then
		   temptext=temptext&"</td></tr><tr height="""&hg1&"""><td valign='top'>"
		   else
		   temptext=temptext&"</td><td valign='top'>"
		   end if
		   rsi.close
           set rsi=nothing
           tj=tj+1
		   if tj>=ts1 And ts1>0 then exit Do		   
		   rs.movenext
		   Loop
end if
Call CloseRs(rs)

           temptext=temptext&"</span></td></tr></TABLE>"
c_more=temptext
End Function

'网店或名片一级栏目导航（只显示第一级分类）
Function V_vassal_c_t(lb,c1,c2,c3,tj,s,zh)
'标签说明,依次为：网店或名片（0网店1名片),一级地区，二级地区，三级地区,显示所在分类名片条数（0不显1显),显示记录数（0全部其它按指定数)，字号大小
Dim rs,sql,dqDocument,titlecolor,TempText,show
If lb=0 Then
dqDocument="sdindex.asp"
show=ttccdp
Else
dqDocument="qyindex.asp"
show=ttccmp
End If
Call OpenConn
set rs=server.CreateObject("adodb.recordset")
sql="select id,[name],titlecolor from DNJY_hy_type where id>0 and indexshow='yes' and twoid=0 and threeid=0 order by indexid"               
set rs=server.createobject("adodb.recordset")               
rs.open sql,conn,1,1
i=0                                                                                            
do while not rs.eof
titlecolor=""
If rs(0)=t1 Then titlecolor="ff0000"
TempText=TempText&"<span><A href="""&dqDocument&"?t="&rs(0)&",0,0&"&C_ID&"""><FONT style=""font-size:"&zh&"pt"" color=#"&titlecolor&">"&rs(1)&"</FONT></A>"
     If tj=1 then
	 Sql="Select count(id) from DNJY_user Where type_oneid="&rs("id")&""&show&""
	 temptext=temptext&"<FONT color=""ff0000"">("&conn.execute(sql)(0)&")</FONT>"
	 End IF
temptext=temptext&" ｜</span>"
i=i+1
if i>=s And s>0 then exit do 
rs.movenext 
loop
Call CloseRs(rs)

V_vassal_c_t=TempText
End Function

'网店名片或地区分类目录列表====================================================
Function W_vassal_c_t(lb,wj,c1,c2,c3,t1,t2,t3,q,n,c,l,s,h)
'标签说明,依次为：网店或名片（0网店,1名片),使用所在文件名(0频道首页,1为列表页),一级地区，二级地区，三级地区，一级分类，二级分类，三级分类，标题圆点标志,显示所在分类名片条数（0不显1显),颜色（0不显1显),列数,显示名片或地区分类（0地区1名片),行高
Dim rs,sql,dqDocument
Call OpenConn
 Set Rs=server.createobject("aDodb.recordset") 
	Dim ci,t,ii,cc,temptext,show
	IF s=1 Then
	IF t1=0 then exit Function
	IF t2>0 Then
	t="id="&t1&" and twoid="&t2&""
	Else
	t="id="&t1&" and threeid=0"
	End IF
	Else
	IF c2>0 Then
	ci="id="&c1&" and twoid="&c2&""
	ElseIF c1>0 Then
	ci="id="&c1&" and threeid=0"
	Else
	ci="twoid=0"
	End IF
	End IF
	
If lb=0 And wj=0 Then
dqDocument="sdindex.asp"
show=ttccdp
elseIf lb=0 And wj=1 Then
dqDocument="sdclass.asp"
show=ttccdp
End if
If lb=1 And wj=0 Then
dqDocument="qyindex.asp"
show=ttccmp
elseIf lb=1 And wj=1 Then
dqDocument="qyclass.asp"
show=ttccmp
End if

	ii=0
	temptext="<table align=center width=""100%"" border=0 cellpadding=0 cellspacing=0>"
	IF s=1 Then
	Set Rs=conn.execute("select twoid,threeid,[name],[titlecolor],id from DNJY_hy_type where "&t&" and twoid>0 order by indexid")
	do while not rs.eof
	IF ii mod l=0 Then temptext=temptext&"<TR height="&h&" align=center>"
	temptext=temptext&"<td>"
	IF q=1 Then temptext=temptext&"·"
	IF c=1 Then temptext=temptext&"<font color="&rs(3)&">"
	temptext=temptext&"<a href="&dqDocument&"?c="&c1&","&c2&","&c3&"&t="&t1&","&rs(0)&","&rs(1)&"> "&rs(2)&" </a>"
	IF c=1 Then temptext=temptext&"</font>"
	IF n=1 Then
	IF t2=0 Then
	Sql="Select count(id) from DNJY_user Where type_oneid="&t1&" and type_twoid="&rs(0)&" "&show&""
	Else
	Sql="Select count(id) from DNJY_user Where type_oneid="&t1&" and type_twoid="&t2&" and type_threeid="&rs(1)&" "&show&""
	End IF
	temptext=temptext&"("&conn.execute(sql)(0)&")"
	End IF
	temptext=temptext&"</td>"
	Rs.movenext
	ii=ii+1
	IF ii mod l=0 Then temptext=temptext&"</TR>"
	loop
	Else
	Set Rs=conn.execute("select twoid,threeid,[city],id from DNJY_city where "&ci&" order by indexid")
	do while not rs.eof
	IF ii mod l=0 Then temptext=temptext&"<TR height="&h&" align=center>"
	temptext=temptext&"<td>"
	IF q=1 Then temptext=temptext&"·"
	temptext=temptext&"<a href="&dqDocument&"?c="&rs(3)&","&rs(0)&","&rs(1)&"&t="&t1&","&t2&","&t3&">"&rs(2)&"</a></td>"
	Rs.movenext
	ii=ii+1
	IF ii mod l=0 Then temptext=temptext&"</TR>"
	loop
	
	End IF
Call CloseRs(rs)

	
	IF ii mod l<>0 Then
	do while ii mod l<>0
	temptext=temptext&"<TD>&nbsp;</TD>"
	ii=ii+1
	loop
	temptext=temptext&"</TR>"
	End IF
	W_vassal_c_t=TempText&"</table>"
End Function

'网店名片分类目录一（只显示二级）＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
Function m_more(lb,ty1,by1,by2,ts1,ts2,n1,n2,ls1,zt1,zt2,hg1,hg2)  
 '标签说明,依次为：lb网店或名片（0网店1名片),ty1一级分类ID号（0 为全部），by1是否显示一级分类（0不显1显)，by2是否显示二级分类（0不显1显)，ts1一级分类显示个数(0为全部），ts2二级分类显示个数（0为全部）,显示一级分类属下名片条数（0不显1显),显示二级分类属下名片条数（0不显1显),ls1一级分类显示列数（0不分列)，zt1一级分类字号，zt2二级分类字号，hg1一级分类行距，hg2二级分类行距
 Dim temptext,rsi,ti,ts,tj,idd,dqDocument,titlecolor,titlecolor2,lname,kg,k,show
If lb=0 Then
dqDocument="sdclass.asp"
lname="sdname"
kg="sdkg"
show=ttccdp
else
dqDocument="qyclass.asp"
lname="mpname"
kg="mpkg"
show=ttccmp
End If

    temptext="<TABLE border=""0"" cellspacing=""0"" cellpadding=""0"">"
      temptext=temptext&"<tr><td valign=""top"">"         
     idd=" and id="&ty1
	 If ty1=0 Then idd=""
	 If by2=0 Then k="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
	 Call OpenConn
			set rs=server.CreateObject("adodb.recordset")
			rs.Open "select * from DNJY_hy_type where twoid=0 and threeid=0"&idd&" order by indexid",conn,1,1
			tj=0
			do while not rs.eof
	  titlecolor="1835D1"
      If rs(0)=t1 Then titlecolor="ff0000"
      temptext=temptext&"<TABLE border=""0"" cellspacing=""0"" cellpadding=""0"">"
	  temptext=temptext&"<TR height="""&hg1&"""><TD></TD><TD>"
	  If by1=1 then
	  temptext=temptext&""&k&"<A href="""&dqDocument&"?"&C_ID&"&t="&rs("id")&","&rs("twoid")&","&rs("threeid")&"""><FONT style=""font-size:"&zt1&"pt"" color=""#"&titlecolor&"""><B>"&rs("name")&"</B></FONT></A>"
     If n1=1 then
	 Sql="Select count(id) from DNJY_user Where "&lname&"<>'' and "&kg&"=1 and type_oneid="&rs("id")&""&show&""
	 temptext=temptext&"<FONT color=""ff0000"">("&conn.execute(sql)(0)&")</FONT>"
	 End IF
	 End if
	  temptext=temptext&"</TD></TR>"	  
       set rsi=conn.execute("select * from DNJY_hy_type where id="&rs("id")&" and twoid<>0 and threeid=0 order by  indexid,id,twoid")
	       If by2=1 then
			temptext=temptext&"<tr height="""&hg1&"""><td></td><td>"
			ts=0
			do while not rsi.eof
			titlecolor2="1835D1"
            If rsi("id")=t1 And rsi("twoid")=t2 Then titlecolor2="ff0000"
            temptext=temptext&"<A href="""&dqDocument&"?"&C_ID&"&t="&rsi("id")&","&rsi("twoid")&","&rsi("threeid")&"""> <FONT style=""font-size:"&zt2&"pt"" color=""#"&titlecolor2&""">"&rsi("name")&"</FONT></A>"
            If n2=1 then
		    Sql="Select count(id) from DNJY_user Where "&lname&"<>'' and "&kg&"=1 and type_oneid="&rs("id")&" and type_twoid="&rsi("twoid")&""&show&""
	        temptext=temptext&"<FONT color=""ff0000"">("&conn.execute(sql)(0)&")</FONT>"
			End if
		    temptext=temptext&"<FONT color=""#CCCCCC"">|</FONT>"
           ts=ts+1
		   if ts>=ts2 And ts2>0 Then
		   temptext=temptext&"<A href="""&dqDocument&"?"&C_ID&"&t="&rs("id")&","&rs("twoid")&","&rs("threeid")&""">>>></A>"
		   exit do 
           End if
           rsi.movenext
		   Loop
		   End if
		   temptext=temptext&"</td></tr></table>"
		   ti=ti+1
		   if ti mod ls1=0 then
		   temptext=temptext&"</td></tr><tr height="""&hg1&"""><td valign='top'>"
		   else
		   temptext=temptext&"</td><td valign='top'>"
		   end if
		   rsi.close
           set rsi=nothing
           tj=tj+1
		   if tj>=ts1 And ts1>0 then exit do 
		   rs.movenext
		   Loop
Call CloseRs(rs)

           temptext=temptext&"</span></td></tr></TABLE>"		  
m_more=temptext
End Function

'网店名片分类目录二（只显示一个一级分类及其属下二级）＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
Function n_more(lb,by1,by2,ts1,ts2,n1,n2,ls1,zt1,zt2,hg1,hg2)  
 '标签说明,依次为：lb网店或名片（0网店1名片),by1是否显示一级分类（0不显1显)，by2是否显示二级分类（0不显1显)，ts1一级分类显示个数(0为全部），ts2二级分类显示个数（0为全部）,n1显示一级分类属下名片条数（0不显1显),n2显示二级分类属下名片条数（0不显1显),ls1一级分类显示列数（0不分列)，zt1一级分类字号，zt2二级分类字号，hg1一级分类行距，hg2二级分类行距
 Dim temptext,rsi,ti,ts,tj,idd,dqDocument,titlecolor,titlecolor2,lname,kg,show
If lb=0 Then
dqDocument="sdclass.asp"
lname="sdname"
kg="sdkg"
show=ttccdp
else
dqDocument="qyclass.asp"
lname="mpname"
kg="mpkg"
show=ttccmp
End If

    temptext="<TABLE border=""0"" cellspacing=""1"" cellpadding=""0"">"
      temptext=temptext&"<tr><td valign=""top"">"       
Call OpenConn
			set rs=server.CreateObject("adodb.recordset")
			rs.Open "select * from DNJY_hy_type where id="&type_oneid&" and twoid=0 and threeid=0 order by indexid",conn,1,1
			tj=0
			do while not rs.eof
	  titlecolor="1835D1"
      If rs(0)=t1 Then titlecolor="ff0000"
      temptext=temptext&"<TABLE border=""0"" cellspacing=""1"" cellpadding=""0"">"
	  temptext=temptext&"<TR height="""&hg1&"""><TD></TD><TD>"
	  If by1=1 then
	  temptext=temptext&"<A href="""&dqDocument&"?"&C_ID&"&t="&rs("id")&","&rs("twoid")&","&rs("threeid")&"""><FONT style=""font-size:"&zt1&"pt"" color=""#"&titlecolor&"""><B>"&rs("name")&"</B></FONT></A>"
     If n1=1 then
	 Sql="Select count(id) from DNJY_user Where "&lname&"<>'' and "&kg&"=1 and type_oneid="&rs("id")&""&show&""
	 temptext=temptext&"<FONT color=""ff0000"">("&conn.execute(sql)(0)&")</FONT>"
	 End IF
	 End if
	  temptext=temptext&"</TD></TR>"	  
       set rsi=conn.execute("select * from DNJY_hy_type where id="&rs("id")&" and twoid<>0 and threeid=0 order by  indexid,id,twoid")
	       If by2=1 then
			temptext=temptext&"<tr height="""&hg1&"""><td></td><td>"
			ts=0
			do while not rsi.eof
			titlecolor2="1835D1"
            If rsi("id")=t1 And rsi("twoid")=t2 Then titlecolor2="ff0000"
            temptext=temptext&"<A href="""&dqDocument&"?"&C_ID&"&t="&rsi("id")&","&rsi("twoid")&","&rsi("threeid")&"""> <FONT style=""font-size:"&zt2&"pt"" color=""#"&titlecolor2&""">"&rsi("name")&"</FONT></A>"
            If n2=1 then
		    Sql="Select count(id) from DNJY_user Where "&lname&"<>'' and "&kg&"=1 and type_oneid="&rs("id")&" and type_twoid="&rsi("twoid")&""&show&""
	        temptext=temptext&"<FONT color=""ff0000"">("&conn.execute(sql)(0)&")</FONT>"
			End if
		    temptext=temptext&"<FONT color=""#CCCCCC"">|</FONT>"
           ts=ts+1
		   if ts>=ts2 And ts2>0 Then
		   temptext=temptext&"<A href="""&dqDocument&"?"&C_ID&"&t="&rs("id")&","&rs("twoid")&","&rs("threeid")&""">>>></A>"
		   exit do 
           End if
           rsi.movenext
		   Loop
		   End if
		   temptext=temptext&"</td></tr></table>"
		   ti=ti+1
		   if ti mod ls1=0 then
		   temptext=temptext&"</td></tr><tr height="""&hg1&"""><td valign='top'>"
		   else
		   temptext=temptext&"</td><td valign='top'>"
		   end if
		   rsi.close
           set rsi=nothing
           tj=tj+1
		   if tj>=ts1 And ts1>0 then exit do 
		   rs.movenext
		   Loop
Call CloseRs(rs)

           temptext=temptext&"</span></td></tr></TABLE>"		  
n_more=temptext
End Function

'网店名片栏目内容列表＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function dq_more(t1,t2,t3,lx,ts,qz,gd,px,btw,zt,hg,dj,rq,d)
't1一级栏目ID号，t2二级栏目ID号，t3三级栏目ID号，lx类型(0网店1名片)，ts显示条数，qz标题前缀(0不显1显红勾2显蓝勾3显蓝箭)，gd固顶标志(0不显1显)，px排序(0为开店时间,1为固顶,2为点击数,btw标题长度，zt字体,hg行高，dj点击数(0不显1显)，rq日期(0不显1显)，d是否显示更多(0不显1显)

Dim c,t,bigclassname,oColor,name,TempText,gdd,djj,wjm,show
c=request("c")
If request("c")="" Then c=c4
c=split(c,",")
If c(0)="" Then c(0)=0
c1=strint(c(0))
If c(1)="" Then c(1)=0
c2=strint(c(1))
If c(2)="" Then c(2)=0
c3=strint(c(2))
	
	IF t1=0 then
	t=""
	elseIF t3>0 then
	t=" and type_oneid="&t1&" and type_twoid="&t2&" and type_threeid="&t3
	elseIF t2>0 then
	t=" and type_oneid="&t1&" and type_twoid="&t2
	Else
	t=" and type_oneid="&t1
	End If
If lx=0 Then
lx="sdkg=1 and sdname"
bigclassname="网店"
gdd="sdgd"
djj="sdhits"
name="sdname"
wjm="user/com.asp"
show=ttccdp
Else
lx="mpkg=1 and mpname"
bigclassname="名片"
gdd="mpgd"
djj="mphits"
name="mpname"
wjm="company.asp"
show=ttccmp
End if

 TempText=TempText&"<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"" bordercolor=""#111111"" style=""border-collapse: collapse""><tr><TD>"
Dim ae
ae=0
Call OpenConn
set rs=server.createobject("adodb.recordset")
sql = "select * from DNJY_user where "&lx&"<>''"&t&""&show&" order by"
If px=0 Then sql=sql&" dpdata "&DN_OrderType&""
If px=1 Then sql=sql&" "&gdd&" "&DN_OrderType&",id "&DN_OrderType&""
If px=2 Then sql=sql&" "&djj&" "&DN_OrderType&",id "&DN_OrderType&""
rs.open sql,conn,3,3
if rs.eof Then
TempText=TempText&"<p><br>暂无"
IF c1>0 Then TempText=TempText&""& conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)&""
TempText=TempText&""&bigclassname&"!"
else
do while not rs.eof
TempText=TempText&"<tr height="""&hg&"""><td valign=""bottom""><div>"
If qz=1 Then TempText=TempText&"&nbsp;<IMG src=""images/qyimg/qyls/arrow_gray.gif"" border=0> "
If qz=2 Then TempText=TempText&"&nbsp;<IMG src=""images/qyimg/qyls/rowup.gif"" border=0> "
If qz=3 Then TempText=TempText&"&nbsp;<IMG src=""images/qyimg/qyls/article_common.gif"" border=0> "
if rs(""&gdd&"")=1 And gd=1 Then
TempText=TempText&"<img src=""images/top.gif"" width=""9"" height=""15"" border=0 alt=""固顶"">"
end if
TempText=TempText&"<a href="&wjm&"?username="&Cstr(rs("username"))&"&com="&Cstr(rs("username"))&"&"&T_ID&"&"&C_ID&" target=_blank title=此条被浏览过"&rs(""&djj&"")&"次>"
TempText=TempText&"<FONT style=""font-size:"&zt&"pt"">"&left(trim(rs(""&name&"")),btw)&"</font>"
TempText=TempText&"</a>"
If dj=1 Or rq=1 Then TempText=TempText&"<FONT color=#999999>["
If dj=1 Then
TempText=TempText&""&rs(""&djj&"")&""
End if
If rq=1 Then TempText=TempText&"|"&dicksj2(rs("dpdata"))&""
If dj=1 Or rq=1 Then TempText=TempText&"<FONT color=#999999>]</font>"
TempText=TempText&"</td></tr></TD></tr>"
                ae=ae+1
                if ae>=ts then exit do
                rs.movenext
                Loop
                End if
Call CloseRs(rs)

                If d=1 Then TempText=TempText&"<tr><td align=""right""><a href=""news.asp?"&C_ID&""" target=_blank>更多>>></a></TD></tr>"
				TempText=TempText&"</table>"
dq_more=TempText
end Function


'网店名片栏目内容分类导航＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function dq_more_dh(lx,ls,tm,ts,qz,gd,px,btw,zt,hg,dj,rq,d)
'lx类型(0网店1名片)，ls显示列数，tm显示栏目个数，ts显示条数，qz标题前缀(0不显1显红勾2显蓝勾3显蓝箭)，gd固顶标志(0不显1显)，px排序(0为开店时间,1为固顶,2为点击数,btw标题长度，zt字体,hg行高，dj点击数(0不显1显)，rq日期(0不显1显)，d是否显示更多(0不显1显)

Dim c,bigclassname,oColor,name,TempText,gdd,djj,wjm,gdclass,ae,show
c=request("c")
If request("c")="" Then c=c4
c=split(c,",")
If c(0)="" Then c(0)=0
c1=strint(c(0))
If c(1)="" Then c(1)=0
c2=strint(c(1))
If c(2)="" Then c(2)=0
c3=strint(c(2))
	
If lx=0 Then
lx="sdkg=1 and sdname"
bigclassname="网店"
gdd="sdgd"
djj="sdhits"
name="sdname"
wjm="user/com.asp"
gdclass="sdclass.asp"
show=ttccdp
Else
lx="mpkg=1 and mpname"
bigclassname="名片"
gdd="mpgd"
djj="mphits"
name="mpname"
wjm="company.asp"
gdclass="qyclass.asp"
show=ttccmp
End if

Dim rsclass,i,m
Call OpenConn
set rsclass=server.createobject("adodb.recordset")
rsclass.Open "select * from DNJY_hy_type where twoid=0 and threeid=0 order by indexid",conn,1,1
if rsclass.eof Then
Response.Write "<font color=000000>暂无分类</font>"
else
TempText=TempText&"<TABLE height=184 cellSpacing=0 cellPadding=0 width=980 align=center border=0 class=""dq3"">"
TempText=TempText&"<TBODY><TR><TD width=980 style=""font-size: 12px; color: #000000; font-family: Verdana, Arial, Helvetica, sans-serif"">"
TempText=TempText&"<TABLE><TBODY><TR>"

Do while Not rsclass.eof



TempText=TempText&"<TD align=middle width=250 style=""font-size: 12px; color: #000000; font-family: Verdana, Arial, Helvetica, sans-serif"">"
TempText=TempText&"<TABLE cellSpacing=0 cellPadding=0 width=242 align=center bgColor=#ffffff border=0>"
TempText=TempText&"<TBODY><TR><TD vAlign=top align=left style=""font-size: 12px; color: #000000; font-family: Verdana, Arial, Helvetica, sans-serif"">"
TempText=TempText&"<IMG height=3 src=""images/qyimg/qyls/prd_rig01.gif"" width=3></TD>"
TempText=TempText&"<TD class=tab_top_1c width=""100%""><IMG src=""images/qyimg/qyls/spacer.gif""></TD>"
TempText=TempText&"<TD vAlign=top align=right style=""font-size: 12px; color: #000000; font-family: Verdana, Arial, Helvetica, sans-serif"">"
TempText=TempText&"<IMG height=3 src=""images/qyimg/qyls/prd_rig02.gif"" width=3></TD></TR>"
TempText=TempText&"<TR vAlign=bottom><TD align=middle colSpan=3 height=20 style=""font-size: 12px; color: #000000; font-family: Verdana, Arial, Helvetica, sans-serif"">"
TempText=TempText&"<a href="&gdclass&"?id="&rsclass("id")&"&t="&rsclass("id")&",0,0&"&C_ID&"><font class=f14><b>"&rsclass("name")&"</b></font></a></strong></font></TD></TR>"
TempText=TempText&"<TR><TD class=dot_xline align=middle background=""images/qyimg/qyls/dot_xline.gif"" colSpan=3 height=1></TD></TR></TBODY></TABLE>"
TempText=TempText&"<div align=""center""><TABLE cellSpacing=0 cellPadding=2 width=242 bgColor=#ffffff border=0><TBODY><TR class=""dq2""><TD colSpan=4 height=145 style=""font-size: 12px; color: #000000; font-family: Verdana, Arial, Helvetica, sans-serif"" valign=""top"">"


 TempText=TempText&"<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"" bordercolor=""#111111"" style=""border-collapse: collapse""><tr><TD>"
ae=0
set rs=server.createobject("adodb.recordset")
sql = "select * from DNJY_user where "&lx&"<>'' and type_oneid="&rsclass("id")&""&show&" order by"
If px=0 Then sql=sql&" dpdata "&DN_OrderType&""
If px=1 Then sql=sql&" "&gdd&" "&DN_OrderType&",id "&DN_OrderType&""
If px=2 Then sql=sql&" "&djj&" "&DN_OrderType&",id "&DN_OrderType&""
rs.open sql,conn,3,3
if rs.eof Then
TempText=TempText&"<p><br>暂无"
IF c1>0 Then TempText=TempText&""& conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)&""
TempText=TempText&""&bigclassname&"!"
else
do while not rs.eof
TempText=TempText&"<tr height="""&hg&"""><td valign=""bottom""><div>"
If qz=1 Then TempText=TempText&"&nbsp;<IMG src=""images/qyimg/qyls/arrow_gray.gif"" border=0> "
If qz=2 Then TempText=TempText&"&nbsp;<IMG src=""images/qyimg/qyls/rowup.gif"" border=0> "
If qz=3 Then TempText=TempText&"&nbsp;<IMG src=""images/qyimg/qyls/article_common.gif"" border=0> "
if rs(""&gdd&"")=1 And gd=1 Then
TempText=TempText&"<img src=""images/top.gif"" width=""9"" height=""15"" border=0 alt=""固顶"">"
end if
TempText=TempText&"<a href="&wjm&"?username="&Cstr(rs("username"))&"&com="&Cstr(rs("username"))&"&"&T_ID&"&"&C_ID&" target=_blank title=此条被浏览过"&rs(""&djj&"")&"次>"
TempText=TempText&"<FONT style=""font-size:"&zt&"pt"">"&left(trim(rs(""&name&"")),btw)&"</font>"
TempText=TempText&"</a>"
If dj=1 Or rq=1 Then TempText=TempText&"<FONT color=#999999>["
If dj=1 Then
TempText=TempText&""&rs(""&djj&"")&""
End if
If rq=1 Then TempText=TempText&"|"&dicksj2(rs("dpdata"))&""
If dj=1 Or rq=1 Then TempText=TempText&"<FONT color=#999999>]</font>"
TempText=TempText&"</td></tr></TD></tr>"
                ae=ae+1
                if ae>=ts then exit do
                rs.movenext
                Loop
                End if
                rs.close
                set rs=nothing
                If d=1 Then TempText=TempText&"<tr><td align=""right""><a href="""&gdclass&"?id="&rsclass("id")&"&t="&rsclass("id")&",0,0&"&C_ID&""" >更多>>></a></TD></tr>"
				TempText=TempText&"</table>"

TempText=TempText&"</TD></TR>"
TempText=TempText&"<TR><TD vAlign=bottom align=left style=""font-size: 12px; color: #000000; font-family: Verdana, Arial, Helvetica, sans-serif"" width=""61"">"
TempText=TempText&"<IMG height=3 src=""images/qyimg/qyls/prd_rig03.gif"" width=3></TD>"
TempText=TempText&"<TD class=tab_bot_1c width=""99"" colspan=""2""><IMG src=""images/qyimg/qyls/spacer.gif""></TD>"
TempText=TempText&"<TD vAlign=bottom align=right> <IMG height=3 src=""images/qyimg/qyls/prd_rig04.gif"" width=3></TD></TR></TBODY></TABLE></div></TD>"
i=i+1
If i=ls then
TempText=TempText&"</tr><TD align=middle width=2 style=""font-size: 12px; color: #000000; font-family: Verdana, Arial, Helvetica, sans-serif""></TD></TR>"
	i=0
	End If
m=m+1
rsclass.movenext
if m>=tm then exit do
Loop
end if
rsclass.close
set rsclass=Nothing

TempText=TempText&"<TD align=middle width=2 style=""font-size: 12px; color: #000000; font-family: Verdana, Arial, Helvetica, sans-serif""></TD></TR></TBODY></TABLE>"
dq_more_dh=TempText
end function


'店铺新闻共享－－－－－－－－－－－－－－－－－－－－－－－－－－
function vipnews(zz,t,lx,gd,tj,tp,hits,d,btw,h,l,zt,hg,gdnews)

'参数依次：zz作者（指定作者，""为全部），t栏目（新闻中心、荣誉资质、求贤纳士、下载中心，根据ID号定,0为全部），lx类型（0全部、1为推荐、2为热门），gd启用固顶（0不固1固），tj推荐标志（0不显1显），tp图片标志（0不显1显），hits点击数（0不显1显），d时间（0-15种格式，0不显），btw标题长度，显示条数，l列数，zt字号，hg行高，gdnews是否显示更多（0不显1显）
	Dim ii,i,Sql,TempText,newstop,rs,username,ClassID
	Dim id,title,firstImageName,lxx
	IF zz="" Then
	username=""
	else
	username=" and username='"&zz&"'"    	
	End If
    IF t=0 Then
	ClassID=""
	else
	ClassID=" and ClassID="&t&""    	
	End If		
	IF lx=1 then
	lxx=" and tj="&DN_True&""	
	elseif lx=0 then
    lxx=""	
	End If
	Call OpenConn
    Set Rs=server.createobject("aDodb.recordset")               
	Sql="select * from DNJY_userNews where newsShared=0 and ifshow=1"&ClassID&""&username&""&lxx&" order by "
	If gd=1 Then sql=sql&"newstop "&DN_OrderType&","
	if lx=3 Then
	sql=sql&"hits "&DN_OrderType&","
	End if
	sql=sql&"id "&DN_OrderType&""
	Rs.open Sql,conn,1,1
	TempText=""
if rs.eof Then
TempText=" <p><br>暂无内容!"
Else
	
	TempText=TempText&"<table width=""100%""  border=""0"" cellspacing=""2"" cellpadding=""0"" >"
	dim iii
	iii=0
	Do While not Rs.eof
if conn.Execute("Select sdkg From DNJY_user Where username='"&rs("username")&"'")(0)=1 then
    i=i+1
	IF iii Mod l=0 then TempText=TempText&  "<tr height="""&hg&""">"
	TempText=TempText&  "<td><img src=""images/qyimg/qyls/rowup.gif""> "
    IF rs("newstop")=1 And gd=1 Then TempText=TempText&  "<img src=""img/top4.gif"" width=""15"" height=""13"" border=0 alt=""固顶"">"
    IF rs("tj")=true And tj=1 Then TempText=TempText&  "<img src=""img/jian.gif"" width=""15"" height=""13"" border=0 alt=""推荐"">"
    IF rs("SavePathFileName")<>"" And  tp=1 Then TempText=TempText&  "<img src=""images/num/pic.gif"" width=""13"" height=""13"" border=0 alt=""图片"">"
	TempText=TempText&  "<a title="""&rs("title")&""" href=""user/article.asp?id="&rs("id")&"&com="&rs("username")&""" target=""_blank""><SPAN style=""color:""000000"";font-size="&zt&"px"">"&CutStr(rs("title"),btw,"..")&"</SPAN></a> "
    IF hits=1 Then TempText=TempText&  " <font color=""#999999"" >["&rs("hits")&"]</font>" 
	TempText=TempText&  " <font color=""#999999"" >"&FormatDate(rs("addtime"),d)&"</font>"    
	TempText=TempText&  "</TD>"
	iii=iii+1
	IF iii Mod l=0 then TempText=TempText&  "</tr>"
    if i>=h then exit Do
    End if 
	Rs.movenext
	Loop	
	IF iii Mod l<>0 then
	Do While not iii Mod l=0
	TempText=TempText&  "<td></td>"
	iii=iii+1
	Loop
	TempText=TempText&  "</tr>"
	End IF
	TempText=TempText&  "</table>"
	If gdnews=1 Then TempText=TempText&  "<tr><td align=""right"" colSpan=6><A href=""vipnews.asp?zz="&zz&"&t="&t&"&lx="&lx&""">更多>>></a></td></tr>"
	End IF
    Call CloseRs(rs)
	vipnews=TempText

end Function
'店铺新闻共享END－－－－－－－－－－－－－－－－－－－－－－－－－－


'店企搜索条=================================================================================
Function dq_search(lb,t,c)
'参数依次为：lb为类别（0店铺，1名片），t为信息分类显示级数（1为一级，2为二级，3为三级），c为地区分类显示级数（1为一级，2为二级，3为三级）。建议用二级以下！
	
	dim TempText,wjm,ss
	If lb=0 Then
    wjm="sdclass.asp"
	ss="店铺"
	Else
	wjm="qyclass.asp"
	ss="名片"
	End if
	TempText="<FORM action="""&wjm&"?"&C_ID&""" method=""post"" name=""search""><TABLE width=""100%""  border=0 cellpadding=0 cellspacing=1><TR><TD width=""65%""> <b>搜索"&ss&"：</b>&nbsp; <INPUT name=keyword onFocus=""this.value=''"" size=28   value=请输入搜索关键字> <SELECT size=1 name=item><OPTION value=全部 selected>全部</OPTION>"
	Call OpenConn
	If T=1 Then
		Sql="select id,twoid,threeid,name from [DNJY_hy_type] where id>0 And Twoid=0 order by id,twoid,threeid"
	ElseIf T=2 Then
		Sql="select id,twoid,threeid,name from [DNJY_hy_type] where id>0 And Threeid=0 order by id,twoid,threeid"
	Else
		Sql="select id,twoid,threeid,name from [DNJY_hy_type] where id>0 order by id,twoid,threeid"
	End If
	Set Rs=Conn.Execute(Sql)
	do while not Rs.eof
		if Rs(1)=0 and Rs(2)=0 then
			TempText=TempText&"<OPTION value="""&Rs(0)&""" style=""background-color:#F8F4F5 ;color: #FF0000"">"&Rs(3)&"</OPTION>"
		elseif Rs(1)>0 and Rs(2)=0 and t>1 then
			TempText=TempText&"<OPTION value="""&Rs(0)&"a"&Rs(1)&""">&nbsp;&nbsp;"&Rs(3)&"</OPTION>"
		elseif Rs(1)>0 and Rs(2)>0 and t>2 then
			TempText=TempText&"<OPTION value="""&Rs(0)&"a"&Rs(1)&"a"&Rs(2)&""">&nbsp;&nbsp;&nbsp;&nbsp;"&Rs(3)&"</OPTION>"
		end if
		Rs.movenext                                        
	loop
	Call CloseRs(rs)
Dim ccc
If c3>0 Then
ccc=" id="&c1&" and Twoid="&c2&""
ElseIf c2>0 And c3=0 Then
ccc=" id="&c1&" and Twoid="&c2&""
ElseIf c1>0 And c2=0 Then
ccc=" id="&c1&" and threeid=0"
ElseIf c1=0 Then
ccc=" Twoid=0 and threeid=0"
End if

TempText=TempText&"</SELECT><SELECT name=icity size=1 id=icity><OPTION value=总站 selected>所有城市</OPTION>"
'搜索地区判别,屏蔽的为不定地区============================= 
Sql="select id,twoid,threeid,city from [DNJY_city] where "&ccc&" order by indexid asc,id,twoid,threeid"
	'If C=1 Then
	'Sql="select id,twoid,threeid,city from [DNJY_city] where id>0 And Twoid=0 order by indexid asc, id,twoid,threeid"
	'ElseIf C=2 Then
	'	Sql="select id,twoid,threeid,city from [DNJY_city] where id>0 And threeid=0 order by indexid asc,id,twoid,threeid"
	'Else
	'	Sql="select id,twoid,threeid,city from [DNJY_city] where id>0 order by indexid asc,id,twoid,threeid"
	'End If
'=======================================================
	Set Rs=Conn.Execute(Sql)
	do while not Rs.eof
'搜索地区判别,屏蔽的为不定地区================================= 
		if c3>0 And  Rs(2)=c3 Then
		TempText=TempText&"<OPTION value="""&Rs(0)&"a"&Rs(1)&"a"&Rs(2)&""" selected>&nbsp;&nbsp;&nbsp;&nbsp;"&Rs(3)&"</OPTION>"
		ElseIf c2>0 And c3=0 And  Rs(2)=0 And Rs(1)=c2 Then
		TempText=TempText&"<OPTION value="""&Rs(0)&"a"&Rs(1)&"a"&Rs(2)&""" selected>&nbsp;&nbsp;&nbsp;&nbsp;"&Rs(3)&"</OPTION>"
        ElseIf c1>0 And c2=0 And Rs(1)=0 And Rs(2)=0 And Rs(0)=c1 Then
		TempText=TempText&"<OPTION value="""&Rs(0)&"a"&Rs(1)&"a"&Rs(2)&""" selected>&nbsp;&nbsp;&nbsp;&nbsp;"&Rs(3)&"</OPTION>"
		Else
		TempText=TempText&"<OPTION value="""&Rs(0)&"a"&Rs(1)&"a"&Rs(2)&""">&nbsp;&nbsp;&nbsp;&nbsp;"&Rs(3)&"</OPTION>"
		end if
		'if Rs(1)=0 and Rs(2)=0 then
		'TempText=TempText&"<OPTION value="""&Rs(0)&""" style=""background-color:#F8F4F5 ;color: #FF0000"">"&Rs(3)&"</OPTION>"
		'elseif Rs(1)>0 and Rs(2)=0 and c>1 then
		'TempText=TempText&"<OPTION value="""&Rs(0)&"a"&Rs(1)&""">&nbsp;&nbsp;"&Rs(3)&"</OPTION>"
		'elseif Rs(1)>0 and Rs(2)>0 and c>2 then
		'TempText=TempText&"<OPTION value="""&Rs(0)&"a"&Rs(1)&"a"&Rs(2)&""">&nbsp;&nbsp;&nbsp;&nbsp;"&Rs(3)&"</OPTION>"
		'end if
'===========================================================
		Rs.movenext                                        
	loop                                           
Call CloseRs(rs)

	TempText=TempText&"</SELECT> <INPUT  type=submit value=好手气 style=""font-size: 12px"" name=search></TD></TR></TABLE></FORM>"
	dq_search=TempText
End Function



'信息栏目一（标签方式一,只调用一、二级分类）
 Function menu(ej,hs1,hs2,btw1,btw2,l,lk,hg,gdxx,lmt)
 '标签说明：二级分类（0不显1显），一级分类个数，二级分类个数，一级分类名称长度，二级分类名称长度，ls列数,lk列宽，hg行高，gdxx显示更多（0不显1显）,lmt栏目图标（0不显1显）

dim rs,rs2,TempText,i,j,iii
iii=0
TempText="<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"" >"
If lmt=1 then
TempText=TempText&"<tr><td height=""19"" background=""img/dhmenu.gif"" style=""padding-left:30px;padding-top:10px;font-size:14px;color:#963230;font-weight:600"">分类索引</td></tr>"
End if
TempText=TempText&"<tr><td valign=""top"" style=""padding:10px;""><table class=style1 width=""100%"" >"
Call OpenConn
set rs=server.CreateObject("adodb.recordset")
sql="select id,[name],titlecolor from DNJY_type where id>0 and indexshow='yes' and twoid=0 and threeid=0 order by indexid,id"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.eof Then
Response.Write "<font color=000000>暂无分类或未设置首页显示</font>"
else
i=0
	do while not rs.eof
	i=i+1
IF iii Mod l=0 then TempText=TempText&  "<tr height="""&hg&""">"
temptext=temptext&"<td  valign=""top"">"
TempText=TempText&"<table><tr><td><span class=style1><img src='images/icon1.gif'></span></td></tr></table><table><tr><td  width="""&lk&"""><a class='style1' href=list.asp?t="&rs(0)&"&"&C_ID&"><font color="&rs("titlecolor")&"><b>"&rs("name")&"</b></font></a>"
If ej=1 Then '－－－－－－－－－
set rs2=server.CreateObject("adodb.recordset")
sql="select id,twoid,[name],titlecolor from DNJY_type where id="&rs("id")&" and indexshow='yes' and twoid>0 and threeid=0 order by indexid,twoid"
set rs2=server.createobject("adodb.recordset")
rs2.open sql,conn,1,1
if rs2.eof Then
TempText=TempText&"&nbsp;<font color=000000>暂无下级分类或未设置首页显示</font>"
Else
TempText=TempText&"<table><tr><td ><font color=000000>|</font>"
j=0
do while not rs2.eof
TempText=TempText&"<a class='style1' href=list.asp?t="&rs2(0)&","&rs2(1)&",0&"&C_ID&"><font color="&rs2("titlecolor")&">"&rs2("name")&"</font></a><font color=000000>|</font>"

j=j+1
rs2.movenext
if j>=hs2 then exit Do
Loop
TempText=TempText&"</td></tr></table>"
end If
end If'－－－－
TempText=TempText&"</td></tr></table>"
temptext=temptext&"</td>"
	iii=iii+1
	IF iii Mod l=0 then TempText=TempText&  "</tr>"
    if i>=hs1 then exit do 
	rs.movenext
	Loop	
	IF iii Mod l<>0 then
	Do While not iii Mod l=0
	TempText=TempText&  "<td></td>"
	iii=iii+1
	Loop
	TempText=TempText&  "</tr>"
	End IF
	End If
	TempText=TempText&  "</td></tr></table>"
	If gdxx=1 Then TempText=TempText&  "<tr><td align=""right"" colSpan=6><A href=""type_list.asp?"&CT_ID&"&ttop=6"">更多>>></a></td></tr>"
	TempText=TempText&  "</table>"
	rs.close:Set rs=Nothing
	rs2.close:Set rs2=Nothing
    
	menu=TempText

end Function

'信息栏目二（标签方式二,只调用一、二级分类）
 Function t_more(ty1,by1,by2,ts1,ts2,ls1,zt1,zt2,hg1,hg2)  
 '标签说明,依次为：ty1信息一级分类ID号（0 为全部），by1是否显示一级分类（0不显1显)，by2是否显示二级分类（0不显1显)，ts1一级分类显示个数(0为全部），ts2二级分类显示个数（0为全部）,ls1一级分类显示列数（0不分列)，zt1一级分类字号，zt2二级分类字号，hg1一级分类行距，hg2二级分类行距
 Dim temptext,rsi,ti,ts,tj,idd
    temptext="<TABLE border=""0"" cellspacing=""1"" cellpadding=""0"">"
      temptext=temptext&"<tr><td valign=""top"">"         
     idd=" and id="&ty1
	 If ty1=0 Then idd=""
	 Call OpenConn
			set rs=server.CreateObject("adodb.recordset")
			rs.Open "select * from DNJY_type where twoid=0 and threeid=0"&idd&" order by indexid",conn,1,1
if rs.eof Then
Response.Write "<font color=000000>暂无分类或未设置首页显示</font>"
else
			tj=0
			do while not rs.eof
      temptext=temptext&"<TABLE border=""0"" cellspacing=""1"" cellpadding=""0"">"
	  temptext=temptext&"<TR height="""&hg1&"""><TD></TD><TD>"
	  If by1=1 then
	  temptext=temptext&"<A target=""_blank"" href=""list.asp?t="&rs("id")&","&rs("twoid")&","&rs("threeid")&"""><FONT style=""font-size:"&zt1&"pt"" color=""#1835D1""><B>"&rs("name")&"</B></FONT></A>"
	  End if
	  temptext=temptext&"</TD></TR>"	  
       set rsi=conn.execute("select * from DNJY_type where id="&rs("id")&" and twoid<>0 and threeid=0 order by  indexid,id,twoid")
	       If by2=1 then
			temptext=temptext&"<tr height="""&hg1&"""><td></td><td>"
			ts=0

			do while not rsi.eof
            temptext=temptext&"<A target=""_blank"" href=""list.asp?t="&rsi("id")&","&rsi("twoid")&","&rsi("threeid")&"""> <FONT style=""font-size:"&zt2&"pt"" color=""#1835D1"">"&rsi("name")&"</FONT></A><FONT color=""#CCCCCC"">|</FONT>"
           ts=ts+1
		   if ts>=ts2 And ts2>0 then exit do 
           rsi.movenext
		   Loop
		   
		   End if
		   temptext=temptext&"</td></tr></table>"
		   ti=ti+1
		   if ti mod ls1=0 then
		   temptext=temptext&"</td></tr><tr height="""&hg1&"""><td valign='top'>"
		   else
		   temptext=temptext&"</td><td valign='top'>"
		   end if
		   rsi.close
           set rsi=nothing
           tj=tj+1
		   if tj>=ts1 And ts1>0 then exit do 
		   rs.movenext
		   Loop
End if		   
Call CloseRs(rs)

           temptext=temptext&"</span></td></tr></TABLE>"
 
		   t_more=temptext
End Function

'信息栏目三（标签方式三,只调用一、二级分类）
Function s_more(ty1,by1,by2,ts1,ts2,n1,n2,ls1,zt1,zt2,hg1,hg2)  
 '标签说明,依次为：ty1一级分类ID号（0 为全部），by1是否显示一级分类（0不显1显)，by2是否显示二级分类（0不显1显)，ts1一级分类显示个数(0为全部），ts2二级分类显示个数（0为全部）,n1显示一级分类属下信息条数（0不显1显),n2显示二级分类属下信息条数（0不显1显),ls1一级分类显示列数（0不分列)，zt1一级分类字号，zt2二级分类字号，hg1一级分类行距，hg2二级分类行距
 Dim temptext,rsi,ti,ts,tj,idd,titlecolor,titlecolor2,lname,kg
    temptext="<TABLE border=""0"" cellspacing=""1"" cellpadding=""0"">"
      temptext=temptext&"<tr><td valign=""top"">"         
     idd=" and id="&ty1
	 If ty1=0 Then idd=""
	 Call OpenConn
			set rs=server.CreateObject("adodb.recordset")
			rs.Open "select * from DNJY_type where twoid=0 and threeid=0"&idd&" order by indexid",conn,1,1
if rs.eof Then
Response.Write "<font color=000000>暂无分类或未设置首页显示</font>"
else
			tj=0
			do while not rs.eof
	  titlecolor="1835D1"
      If rs(0)=t1 Then titlecolor="ff0000"
      temptext=temptext&"<TABLE border=""0"" cellspacing=""1"" cellpadding=""0"">"
	  temptext=temptext&"<TR height="""&hg1&"""><TD></TD><TD>"
	  If by1=1 then
	  temptext=temptext&"<A href=""list.asp?"&C_ID&"&t="&rs("id")&","&rs("twoid")&","&rs("threeid")&"""><FONT style=""font-size:"&zt1&"pt"" color=""#"&titlecolor&"""><B>"&rs("name")&"</B></FONT></A>"
     If n1=1 then
	 Sql="Select count(id) from DNJY_info Where yz=1 and type_oneid="&rs("id")&""&ttcc&""
	 temptext=temptext&"<FONT color=""ff0000"">("&conn.execute(sql)(0)&")</FONT>"
	 End IF
	 End if
	  temptext=temptext&"</TD></TR>"	  
       set rsi=conn.execute("select * from DNJY_type where id="&rs("id")&" and twoid<>0 and threeid=0 order by  indexid,id,twoid")
	       If by2=1 Then
       
			temptext=temptext&"<tr height="""&hg1&"""><td></td><td>"
			ts=0
			do while not rsi.eof
			titlecolor2="1835D1"
            If rsi("id")=t1 And rsi("twoid")=t2 Then titlecolor2="ff0000"
            temptext=temptext&"<A href=""list.asp?"&C_ID&"&t="&rsi("id")&","&rsi("twoid")&","&rsi("threeid")&"""> <FONT style=""font-size:"&zt2&"pt"" color=""#"&titlecolor2&""">"&rsi("name")&"</FONT></A>"
            If n2=1 then
		    Sql="Select count(id) from DNJY_info Where  yz=1 and type_oneid="&rs("id")&" and type_twoid="&rsi("twoid")&""&ttcc&""
	        temptext=temptext&"<FONT color=""ff0000"">("&conn.execute(sql)(0)&")</FONT>"
			End if
		    temptext=temptext&"<FONT color=""#CCCCCC"">|</FONT>"
           ts=ts+1
		   if ts>=ts2 And ts2>0 Then
		   temptext=temptext&"<A href=""list.asp?"&C_ID&"&t="&rs("id")&","&rs("twoid")&","&rs("threeid")&""">>>></A>"
		   exit do 
           End if
           rsi.movenext
		   Loop
		   End If
   
		   temptext=temptext&"</td></tr></table>"
		   ti=ti+1
		   if ti mod ls1=0 then
		   temptext=temptext&"</td></tr><tr height="""&hg1&"""><td valign='top'>"
		   else
		   temptext=temptext&"</td><td valign='top'>"
		   end if
		   rsi.close
           set rsi=nothing
           tj=tj+1
		   if tj>=ts1 And ts1>0 then exit do 
		   rs.movenext
		   Loop
 End if
Call CloseRs(rs)

           temptext=temptext&"</span></td></tr></TABLE>"		  
s_more=temptext
End Function

'信息栏目四（标签方式四,调用一、二、三级分类）
Function x_more(ty1,by1,by2,by3,ts1,ts2,ts3,n1,n2,n3,ls,zt1,zt2,zt3,hg1,hg2,hg3)
 '标签说明,依次为：ty1一级分类ID号（0 为全部），by1是否显示一级分类（0不显1显)，by2是否显示二级分类（0不显1显)，by3是否显示三级分类（0不显1显)，ts1一级分类显示个数(0为全部），ts2二级分类显示个数（0为全部），ts3三级分类显示个数（0为全部）,n1显示一级分类属下信息条数（0不显1显),n2显示二级分类属下信息条数（0不显1显),n3显示三级分类属下信息条数（0不显1显),ls一级分类显示列数（0不分列)，zt1一级分类字号，zt2二级分类字号，zt3三级分类字号，hg1一级分类行距，hg2二级分类行距，hg3三级分类行距
 Dim temptext,rsi,rst3,ti,ts,tj,tt,idd,titlecolor,titlecolor2,lname,kg
    temptext="<TABLE border=""0"" cellspacing=""1"" cellpadding=""0"">"
      temptext=temptext&"<tr><td valign=""top"">"         
     idd=" and id="&ty1
	 If ty1=0 Then idd=""
	 Call OpenConn
			set rs=server.CreateObject("adodb.recordset")
			rs.Open "select * from DNJY_type where twoid=0 and threeid=0"&idd&" order by indexid",conn,1,1
if rs.eof Then
Response.Write "<font color=000000>暂无分类或未设置首页显示</font>"
else
			tj=0
			do while not rs.eof
	  titlecolor="1835D1"
      If rs(0)=t1 Then titlecolor="ff0000"
      temptext=temptext&"<TABLE border=""0"" cellspacing=""1"" cellpadding=""0"">"
	  temptext=temptext&"<TR height="""&hg1&"""><TD></TD><TD>"
	  If by1=1 then
	  temptext=temptext&"<A href=""list.asp?"&C_ID&"&t="&rs("id")&","&rs("twoid")&","&rs("threeid")&"""><FONT style=""font-size:"&zt1&"pt"" color=""#"&titlecolor&"""><B>"&rs("name")&"</B></FONT></A>"
     If n1=1 then
	 Sql="Select count(id) from DNJY_info Where yz=1 and type_oneid="&rs("id")&""&ttcc&""
	 temptext=temptext&"<FONT color=""ff0000""><span title=属下有"&conn.execute(sql)(0)&"条信息 style=""cursor:help "">("&conn.execute(sql)(0)&")</span></FONT>"
	 End IF
	 End if
	  temptext=temptext&"</TD></TR>"	  
       set rsi=conn.execute("select * from DNJY_type where id="&rs("id")&" and twoid<>0 and threeid=0 order by  indexid,id,twoid")
	       If by2=1 then
			temptext=temptext&"<tr height="""&hg1&"""><td>&nbsp;&nbsp;</td><td  width="""&920/ls&""">"
			ts=0
			do while not rsi.eof
			titlecolor2="1835D1"
            If rsi("id")=t1 And rsi("twoid")=t2 Then titlecolor2="ff0000"
            temptext=temptext&"<A href=""list.asp?"&C_ID&"&t="&rsi("id")&","&rsi("twoid")&","&rsi("threeid")&"""> <FONT style=""font-size:"&zt2&"pt"" color=""#"&titlecolor2&""">"&rsi("name")&"</FONT></A>"
            If n2=1 then
		    Sql="Select count(id) from DNJY_info Where  yz=1 and type_oneid="&rs("id")&" and type_twoid="&rsi("twoid")&""&ttcc&""
	        temptext=temptext&"<FONT color=""ff0000""><span title=属下有"&conn.execute(sql)(0)&"条信息 style=""cursor:help "">("&conn.execute(sql)(0)&")</span></FONT>"
			End If

			If by3=0 then
		    temptext=temptext&"<FONT color=""#CCCCCC"">|</FONT>"
            Else
            temptext=temptext&"<br>"
            set rst3=conn.execute("select * from DNJY_type where id="&rs("id")&" and twoid="&rsi("twoid")&" and threeid<>0 order by  indexid,id,twoid,threeid")
			if Not rst3.eof Then
		    temptext=temptext&"<table><tr height="""&hg3&"""><td></td><td>└"
			tt=0
			do while not rst3.eof 
			temptext=temptext&"<A target=""_blank"" href=""list.asp?t="&rst3("id")&","&rst3("twoid")&","&rst3("threeid")&"""><FONT style=""font-size:"&zt3&"pt"" color=""#9CA6B5"">"&rst3("name")&"</FONT></A>"			
			If n3=1 then
		    Sql="Select count(id) from DNJY_info Where  yz=1 and type_oneid="&rs("id")&" and type_twoid="&rsi("twoid")&" and type_threeid="&rst3("threeid")&""&ttcc&""
	        temptext=temptext&"<FONT color=""ff0000""><span title=属下有"&conn.execute(sql)(0)&"条信息 style=""cursor:help "">("&conn.execute(sql)(0)&")</span></FONT>"
			End If
			temptext=temptext&"<FONT color=""#CCCCCC"">|</FONT>"
			tt=tt+1
			if tt>=ts3 And ts3>0 Then
            temptext=temptext&"..."
			exit Do
			End if
			rst3.movenext
		   Loop
		  temptext=temptext&"</td></tr></table>"
		   End if
		   rst3.close
           End If

           ts=ts+1
		   if ts>=ts2 And ts2>0 Then
		   temptext=temptext&"<A href=""list.asp?"&C_ID&"&t="&rs("id")&","&rs("twoid")&","&rs("threeid")&""">>>></A>"
		   exit do 
           End if
           rsi.movenext
		   Loop
		   End if
		   temptext=temptext&"</td></tr></table>"
		   ti=ti+1
		   if ti mod ls=0 then
		   temptext=temptext&"</td></tr><tr height="""&hg1&"""><td valign='top'>"
		   else
		   temptext=temptext&"</td><td valign='top'>"
		   end if
		   rsi.close
           set rsi=nothing
           tj=tj+1
		   if tj>=ts1 And ts1>0 then exit do 
		   rs.movenext
		   Loop
end if
Call CloseRs(rs)

           temptext=temptext&"</span></td></tr></TABLE>"		  
x_more=temptext
End Function

'信息栏目页显示方式，参考海洋分类信息网＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
Function TableHead()
	Dim ShowType,OrderType,PageType,Url
	ShowType = StrInt(Request.QueryString("ShowType"))
	OrderType = StrInt(Request.QueryString("OrderType"))
	PageType = StrInt(Request.QueryString("PageType"))
	Url="list.asp?"&CT_ID&"&xxsx="&xxsx
	If ShowType=0 Then
		Response.Write "&nbsp;显示方式：<a href="""&Url&"&OrderType="&OrderType&"&PageType="&PageType&"&ShowType=1""><img src=""img/t_p_11.gif"" width=""14"" height=""14"" style=""margin-bottom:-2px"" alt=""文本方式"" ></a> "&_
		"<img src=""img/t_p_00.gif"" width=""14"" height=""14"" style=""margin-bottom:-2px"" alt=""图文方式""> "&_
		"<a href="""&Url&"&OrderType="&OrderType&"&PageType="&PageType&"&ShowType=2""><img src=""img/t_p_31.gif"" width=""14"" height=""14"" style=""margin-bottom:-2px"" alt=""图片方式""></a> "
	ElseIf ShowType=1 Then
		Response.Write "&nbsp;显示方式：<img src=""img/t_p_1.gif"" width=""14"" height=""14"" style=""margin-bottom:-2px"" alt=""文本方式""> "&_
		"<a href="""&Url&"&OrderType="&OrderType&"&PageType="&PageType&"""><img src=""img/t_p_0.gif"" width=""14"" height=""14"" style=""margin-bottom:-2px"" alt=""图文方式""></a> "&_
		"<a href="""&Url&"&OrderType="&OrderType&"&PageType="&PageType&"&ShowType=2""><img src=""img/t_p_31.gif"" width=""14"" height=""14"" style=""margin-bottom:-2px"" alt=""图片方式""></a> "
	ElseIf ShowType=2 Then
		Response.Write "&nbsp;显示方式：<a href="""&Url&"&OrderType="&OrderType&"&PageType="&PageType&"&ShowType=1""><img src=""img/t_p_11.gif"" width=""14"" height=""14"" style=""margin-bottom:-2px"" alt=""文本方式""></a> "&_
		"<a href="""&Url&"&OrderType="&OrderType&"&PageType="&PageType&"""><img src=""img/t_p_0.gif"" width=""14"" height=""14"" style=""margin-bottom:-2px"" alt=""图文方式""></a> "&_
		"<img src=""img/t_p_3.gif"" width=""14"" height=""14"" style=""margin-bottom:-2px"" alt=""图片方式""> "
	End If
	Response.Write "&nbsp;&nbsp;每页显示数量："
	If PageType=0 Then
		Response.Write "&nbsp;<img src=""img/t_p_20.gif"" width=""18"" height=""15"" style=""margin-bottom:-2px"">"&_
		"<a href="""&Url&"&OrderType="&OrderType&"&PageType=1&ShowType="&ShowType&""">&nbsp;<img src=""img/t_p_401.gif"" width=""18"" height=""15"" style=""margin-bottom:-2px""></a>"&_
		"<a href="""&Url&"&OrderType="&OrderType&"&PageType=2&ShowType="&ShowType&""">&nbsp;<img src=""img/t_p_801.gif"" width=""18"" height=""15"" style=""margin-bottom:-2px""></a>"
	ElseIf PageType=1 Then
		Response.Write "<a href="""&Url&"&OrderType="&OrderType&"&ShowType="&ShowType&""">&nbsp;<img src=""img/t_p_201.gif"" width=""18"" height=""15"" style=""margin-bottom:-2px""></a>"&_
		"&nbsp;<img src=""img/t_p_40.gif"" width=""18"" height=""15"" style=""margin-bottom:-2px"">"&_
		"<a href="""&Url&"&OrderType="&OrderType&"&PageType=2&ShowType="&ShowType&""">&nbsp;<img src=""img/t_p_801.gif"" width=""18"" height=""15"" style=""margin-bottom:-2px""></a>"
	ElseIf PageType=2 Then
		Response.Write "<a href="""&Url&"&OrderType="&OrderType&"&ShowType="&ShowType&""">&nbsp;<img src=""img/t_p_201.gif"" width=""18"" height=""15"" style=""margin-bottom:-2px""></a>"&_
		"<a href="""&Url&"&OrderType="&OrderType&"&PageType=1&ShowType="&ShowType&""">&nbsp;<img src=""img/t_p_401.gif"" width=""18"" height=""15"" style=""margin-bottom:-2px""></a>"&_
		"&nbsp;<img src=""img/t_p_80.gif"" width=""18"" height=""15"" style=""margin-bottom:-2px"">"
	End If
	Response.Write "&nbsp;&nbsp;&nbsp;信息排序方式：<select style=""height:18px;"" onchange=""location='"&Url&"&PageType="&PageType&"&ShowType="&ShowType&"&OrderType='+this.value;""><option value="""""
	If OrderType=0 Then Response.Write " selected=""selected"" "
	Response.Write ">选择排序方式</option><option value=""1"""
	If OrderType=1 Then Response.Write " selected=""selected"" "
	Response.Write " >按发布时间 升序</option><option value=""2"""
	If OrderType=2 Then Response.Write " selected=""selected"" "
	Response.Write " >按发布时间 降序</option><option value=""3"""
	If OrderType=3 Then Response.Write " selected=""selected"" "
	Response.Write " >按浏览量 升序</option><option value=""4"""
	If OrderType=4 Then Response.Write " selected=""selected"" "
	Response.Write " >按浏览量 降序</option></select>"
End Function

'信息显示（详细式）==============================================================
Function f_typeinfo(xxsx,c1,c2,c3,t1,t2,t3,p,xxlx,tw,s,pw,ph,btw,zs)
'标签说明,依次为：xxsx信息属性（1最新、2竞价、3推荐、4热门、5回复），c1一级地区，c2二级地区，c3三级地区，t1信息一级分类，t2信息二级分类，t3信息三级分类，p分页ID号，xxlx信息类型（0不显1显)，tw图文显示（0不显1显)，s每页显示条数，pw图片宽度（宽和高大于0显示），ph图片高度，btw标题长度，zs正文截取长度（字节计）
%><%
dim ThisPage,Pagesize,Allrecord,Allpage,leixing,hfcs,tj
	IF p<0 Then p=0
	If xxsx=""  Then xxsx=0
	dim c,TempText,t,n,b,bb,bbb,num1,bg
	
	IF c1=0 then
	c=""
	elseIF c3>0 then
	c=" and city_oneid="&c1&" and city_twoid="&c2&" and city_threeid="&c3
	elseIF c2>0 then
	c=" and city_oneid="&c1&" and city_twoid="&c2
	Else
	c=" and city_oneid="&c1
	End IF
	
	IF t1=0 then
	t=""
	elseIF t3>0 then
	t=" and type_oneid="&t1&" and type_twoid="&t2&" and type_threeid="&t3
	elseIF t2>0 then
	t=" and type_oneid="&t1&" and type_twoid="&t2
	Else
	t=" and type_oneid="&t1
	End If
	
if request("page")="" then
  ThisPage=1		
else
  ThisPage=request("page")
end If

Dim lx
leixing=request("leixing")
lx=request("leixing")
If leixing="" Then
leixing=""
Else
leixing=" And leixing='"&leixing&"'"
End If
Call OpenConn
set rs=server.createobject("adodb.recordset")
tj=0
sql = "select * from DNJY_info where  yz=1"&leixing&""
if jiaoyi=0 then
sql=sql&" and zt="&jiaoyi
end If
if overdue=0 Then
sql=sql&" and dqsj>"&DN_Now&""
end If
sql=sql&""&c&t&""
if xxsx=2 Then
sql=sql&" and hfxx=1 and hfje>0 and dqsj>"&DN_Now&""
end If
if xxsx=3 Then
sql=sql&" and tj>0"
end If
if xxsx=4 Then
sql=sql&" and llcs>="&hotbz&""
end If
if xxsx=5 Then
sql=sql&" and hfcs>=1"
end If

sql=sql&" order by b "&DN_OrderType&""

if xxsx=1 Then
sql=sql&",fbsj "&DN_OrderType&""
elseif xxsx=2 Then
sql=sql&",hfje/fbts "&DN_OrderType&",fbsj "&DN_OrderType&",id "&DN_OrderType&""
elseif xxsx=3 Then
sql=sql&",fbsj "&DN_OrderType&""
elseif xxsx=4 Then
sql=sql&",llcs "&DN_OrderType&""
elseif xxsx=5 Then
sql=sql&",hfsj "&DN_OrderType&""
Else
sql=sql&",fbsj "&DN_OrderType&""
end If

rs.open sql,conn,1,1
rs.Pagesize=s
Pagesize=rs.Pagesize
Allrecord=rs.Recordcount
Allpage=rs.Pagecount
if ThisPage<1 then                           
ThisPage=1                           
end if   

if rs.eof then
temptext=TempText&"<tr><td>暂无本类信息！</td></tr>"
Else
rs.move (ThisPage-1)*Pagesize

do while not rs.eof
			
b=trim(rs("b"))
bb=len(b)

num1 = Num1+1 '//判断数字奇偶定背景色
If num1 Mod 2=0 Then
bg="F1F9FE"
Else
bg="F9F6FF"
End if
TempText=TempText&"<tr>"

temptext=temptext&"<td bgcolor=""#"&bg&""" align=""left"" width=""46%"" height=""30"">"
If xxlx=1 Then temptext=temptext&"[<a href='list.asp?"&CT_ID&"&leixing="&rs("leixing")&"&xxsx="&xxsx&"'>"&rs("leixing")&"</a>]"
if rs("Readinfo")=1 Then TempText=TempText&"<img src=""img/lock.jpg"" alt=""普通会员权限浏览"">" 
if rs("Readinfo")=2 Then TempText=TempText&"<img src=""img/lockvip.jpg"" alt=""VIP会员权限浏览"">" 
if rs("tupian")<>"0" then
temptext=temptext&"<img src=""images/num/pic.gif"" alt=""有图片"">" 
end if
temptext=temptext&"<a target=""_blank"" href="""&x_path("",rs("id"),"",rs("Readinfo"))&""" title="&rs("biaoti")&">"
if rs("a")="0" and rs("b")="0" and rs("c")<>1 then
temptext=temptext&"<b><u>"&CutStr(rs("biaoti"),btw,"..")&"</u></b>"
end if
if rs("a")<>"0" or rs("b")<>"0" or rs("c")=1 then
temptext=temptext&"<font color=#"&rs("a")&"><b><u>"&CutStr(rs("biaoti"),btw,"..")&"</u></b></font>"
end if
temptext=temptext&"</a>"
if b<>0 then
temptext=temptext&"<img src=""images/num/jsq.gif"" alt=""置顶"&b&"级"">"
for i=1 to bb
temptext=temptext&"<img src=""images/num/"&Mid(b,i,1)&".gif"" alt=""置顶"&b&"级"">"
next
end if
if rs("tj")>0 then
temptext=temptext&"<img src=""images/jian.gif"" alt=""推荐信息"" width=""15"" height=""15"">" 
end If
if rs("hfxx")=1 And rs("dqsj")>now() Then TempText=TempText&"<span title=竞价信息平均每天"&FormatNumber(rs("hfje")/rs("fbts"),2,-1)&"元 style=""cursor:help""><img src=""img/jinja.gif"" width=""16"" height=""16"" ></span>"
temptext=temptext&"</td>"
temptext=temptext&"<td bgcolor=""#"&bg&""" align=""middle"" width=""8%"" height=""22"" >"
temptext=temptext&check_jiage(rs("jiage"))
temptext=temptext&"</td>"
temptext=temptext&"<td bgcolor=""#"&bg&""" align=""middle"" width=""7%"" height=""22"">"&rs("llcs")&"/"&rs("hfcs")&"</td>"
temptext=temptext&"<td bgcolor=""#"&bg&""" align=""middle"" width=""16%"" height=""22"" >"
temptext=temptext&""&dicksj2(rs("fbsj"))&"</td>"
temptext=temptext&"<td  align=""middle"" width=""6%"" height=""22""  bgcolor=""#"&bg&""">"
dim wcf
wcf=rs("zt")
if wcf=1 then
temptext=temptext&"<font color=""#ff0000"">完成</font>"

elseif wcf=0 then
temptext=temptext&"<font color=""#0066FF"">未完</font>"

end if
temptext=temptext&"<td  align=""middle"" width=""9%"" height=""22""  bgcolor=""#"&bg&""">"
dim sj
sj=DateDiff("d",now(),""&rs("dqsj")&"")
if sj>0 then
temptext=temptext&"<font color=""#414141"">剩"&sj&"</b>天</font>"
elseif sj=0 then
temptext=temptext&"<font color=""#414141"">今日到期</font>"
elseif sj<0 then
temptext=temptext&"<font color=""#414141"">已失效</font>"
end if
temptext=temptext&"</td></tr>"
If tw=1 then'图文方式显示
temptext=temptext&"<tr>"
temptext=temptext&"<td width=""95%"" colspan=""4"" bgcolor=""#"&bg&""" height=""60""><span style=""word-break:break-all"">"

Dim vip,memo
memo=CutStr(RemoveHTML_ttkj(trim(rs("memo"))),zs,"..")
If Filterhtm=1 And rs("username")="游客" Then
memo=memo
ElseIf Filterhtm=2 And vip=0 Then
memo=memo
ElseIf Filterhtm>2 Then
memo=memo
else
memo=HTMLDecode(memo)
End if
If request.cookies("dick")("username")<>rs("username") Then'自己发布的全权阅读
vip=0'登录会员
If request.cookies("dick")("username")<>"" Then vip=conn.Execute("Select vip From DNJY_user Where username='"&request.cookies("dick")("username")&"'")(0)
If rs("Readinfo")=1 And request.cookies("dick")("username")="" Then
memo="<font color=#666666>普通会员权限阅读</font>"
End If
If rs("Readinfo")=2 And vip<1 Then
memo="<font color=#666666>VIP会员权限阅读</font>"
End If
End If
temptext=temptext&"<font color=#666666>"&memo&"</font>"

temptext=temptext&"</span>"   
temptext=temptext&"</td>"
temptext=temptext&"<td bgcolor=""#"&bg&""" colspan=""2"">"
If pw>0 and ph>0 Then
IF rs("tupian")<>"0" Then
		TempText=TempText&"<a target=""_blank"" title=""点击放大"" href=UploadFiles/infopic/"&rs("tupian")&"><img src=""UploadFiles/infopic/"&rs("tupian")&""" width="&pw&" height="&ph&"></a>"
	Else
		TempText=TempText&"<img src=""images/nopoto.gif"" width="&pw&" height="&ph&">"
	End IF
End If
TempText=TempText&"</td></tr>"     
End if 
TempText=TempText&"<tr><td colspan=""8"" height=""1"" background=""img/dotline.gif""></td></tr>"
tj=tj+1
rs.movenext
if tj>=Pagesize then exit do
loop
end if
Call CloseRs(rs)


temptext=temptext&"</table></div>"
temptext=temptext&"<div align=""center"">"
temptext=temptext&"<table border=""0"" cellpadding=""0""  width='' style=""border-collapse: collapse"" height=""32""><tr>"
temptext=temptext&"<td height=""25"" width=""151""><p align=""center""><strong> 共有&nbsp;<font color=""#CC5200"">"&Allrecord&"</font>&nbsp;条记录</strong></td>"
temptext=temptext&"<td height=""25"" width=""151""><p align=""center""><strong>每页&nbsp;<font color=""#CC5200"">"&Pagesize&"</font>&nbsp;条记录</strong></td>"
temptext=temptext&"<td height=""25"" width=""126""><p align=""center""><strong>共 <font color=""#CC5200"">"&Allpage&"</font> 页</strong></td>"
temptext=temptext&"<td height=""25"" width=""118""><p align=""center""><strong>现在是第 <font color=""#CC5200"">"&ThisPage&"</font> 页</strong></td>"
temptext=temptext&"<td height=""25"" width=""185""><p align=""right""><strong>"
                      
if ThisPage<2 then     
temptext=temptext&"<font color=""#808080"">首页</font>&nbsp;"
temptext=temptext&"<font color=""#808080"">上一页</font>&nbsp;"     
else     
temptext=temptext&"<a href=?page=1&"&CT_ID&"&xxsx="&xxsx&"&leixing="&lx&"&tw="&tw&">首页</a>&nbsp;"
temptext=temptext&"<a href=?page="&ThisPage-1&"&"&CT_ID&"&xxsx="&xxsx&"&leixing="&lx&"&tw="&tw&">上一页</a>&nbsp;"     
end if
if Allpage-ThisPage<1 then     
temptext=temptext&"<font color=""#808080"">下一页</font>&nbsp;"
temptext=temptext&"<font color=""#808080"">尾页</font>&nbsp;"  
else     
temptext=temptext&"<a href=?page="&(ThisPage+1)&"&"&CT_ID&"&xxsx="&xxsx&"&leixing="&lx&"&tw="&tw&">下一页</a>&nbsp;"   
temptext=temptext&"<a href=?page="&Allpage&"&"&CT_ID&"&xxsx="&xxsx&"&leixing="&lx&"&tw="&tw&">尾页</a>&nbsp;"     
end if
temptext=temptext&"</strong></td></tr>"
f_typeinfo=TempText
End Function



'信息搜索条=================================================================================
Function f_search(t,c)
'参数依次为：信息分类显示级数（1为一级，2为二级，3为三级），地区分类显示级数（1为一级，2为二级，3为三级）。建议用二级以下！
	'If Not IsObject(Conn) or Conn Is NoThing Then dblink
	dim TempText
	TempText="<FORM action=""search.asp?"&C_ID&""" method=""post"" name=""search""><TABLE width=""100%""  border=0 cellpadding=0 cellspacing=1><TR><TD width=""65%""> 搜索信息 <INPUT name=keyword onFocus=""this.value=''"" size=20   value=> <SELECT size=1 name=item><OPTION value=全部 selected>全部</OPTION>"
	Call OpenConn
	If T=1 Then
		Sql="select id,twoid,threeid,name from [DNJY_type] where id>0 And Twoid=0 order by id,twoid,threeid"
	ElseIf T=2 Then
		Sql="select id,twoid,threeid,name from [DNJY_type] where id>0 And Threeid=0 order by id,twoid,threeid"
	Else
		Sql="select id,twoid,threeid,name from [DNJY_type] where id>0 order by id,twoid,threeid"
	End If
	Set Rs=Conn.Execute(Sql)
	do while not Rs.eof
		if Rs(1)=0 and Rs(2)=0 then
			TempText=TempText&"<OPTION value="""&Rs(0)&""" style=""background-color:#F8F4F5 ;color: #FF0000"">"&Rs(3)&"</OPTION>"
		elseif Rs(1)>0 and Rs(2)=0 and t>1 then
			TempText=TempText&"<OPTION value="""&Rs(0)&"a"&Rs(1)&""">&nbsp;&nbsp;"&Rs(3)&"</OPTION>"
		elseif Rs(1)>0 and Rs(2)>0 and t>2 then
			TempText=TempText&"<OPTION value="""&Rs(0)&"a"&Rs(1)&"a"&Rs(2)&""">&nbsp;&nbsp;&nbsp;&nbsp;"&Rs(3)&"</OPTION>"
		end if
		Rs.movenext                                        
	loop
	Call CloseRs(rs)
Dim ccc
If c3>0 Then
ccc=" id="&c1&" and Twoid="&c2&""
ElseIf c2>0 And c3=0 Then
ccc=" id="&c1&" and Twoid="&c2&""
ElseIf c1>0 And c2=0 Then
ccc=" id="&c1&" and threeid=0"
ElseIf c1=0 Then
ccc=" Twoid=0 and threeid=0"
End if

TempText=TempText&"</SELECT><SELECT name=icity size=1 id=icity><OPTION value=总站 selected>所有城市</OPTION>"
'===========================搜索条地区判别开始，与修改根目录search.asp相结合================================ 
'Sql="select id,twoid,threeid,city from [DNJY_city] where "&ccc&" order by indexid asc,id,twoid,threeid"'搜索地区判别数据库定义，定地区显示，与不定地区不能同时开
'------------搜索地区判别数据库定义,不定地区，显示全部开始-----------------------
	If C=1 Then
	Sql="select id,twoid,threeid,city from [DNJY_city] where id>0 And Twoid=0 order by  id,twoid,threeid"
	ElseIf C=2 Then
		Sql="select id,twoid,threeid,city from [DNJY_city] where id>0 And threeid=0 order by id,twoid,threeid"
	Else
		Sql="select id,twoid,threeid,city from [DNJY_city] where id>0 order by id,twoid,threeid"
	End If
'------------搜索地区判别数据库定义,不定地区，显示全部END-----------------------
	Set Rs=Conn.Execute(Sql)
	do while not Rs.eof
'------------搜索地区判别读取地区名,定地区显示开始，与不定地区不能同时开------------
		'if c3>0 And  Rs(2)=c3 Then
		'TempText=TempText&"<OPTION value="""&Rs(0)&"a"&Rs(1)&"a"&Rs(2)&""" selected>&nbsp;&nbsp;&nbsp;&nbsp;"&Rs(3)&"</OPTION>"
		'ElseIf c2>0 And c3=0 And  Rs(2)=0 And Rs(1)=c2 Then
		'TempText=TempText&"<OPTION value="""&Rs(0)&"a"&Rs(1)&"a"&Rs(2)&""" selected>&nbsp;&nbsp;&nbsp;&nbsp;"&Rs(3)&"</OPTION>"
        'ElseIf c1>0 And c2=0 And Rs(1)=0 And Rs(2)=0 And Rs(0)=c1 Then
		'TempText=TempText&"<OPTION value="""&Rs(0)&"a"&Rs(1)&"a"&Rs(2)&""" selected>&nbsp;&nbsp;&nbsp;&nbsp;"&Rs(3)&"</OPTION>"
		'Else
		'TempText=TempText&"<OPTION value="""&Rs(0)&"a"&Rs(1)&"a"&Rs(2)&""">&nbsp;&nbsp;&nbsp;&nbsp;"&Rs(3)&"</OPTION>"
		'end if
'------------搜索地区判别读取地区名,定地区显示END-----------------------
'------------搜索地区判别读取地区名,不定地区，显示全部开始-----------------------
		if Rs(1)=0 and Rs(2)=0 then
		TempText=TempText&"<OPTION value="""&Rs(0)&""" style=""background-color:#F8F4F5 ;color: #FF0000"">"&Rs(3)&"</OPTION>"
		elseif Rs(1)>0 and Rs(2)=0 and c>1 then
		TempText=TempText&"<OPTION value="""&Rs(0)&"a"&Rs(1)&""">&nbsp;&nbsp;"&Rs(3)&"</OPTION>"
		elseif Rs(1)>0 and Rs(2)>0 and c>2 then
		TempText=TempText&"<OPTION value="""&Rs(0)&"a"&Rs(1)&"a"&Rs(2)&""">&nbsp;&nbsp;&nbsp;&nbsp;"&Rs(3)&"</OPTION>"
		end if
'------------搜索地区判别读取地区名,不定地区，显示全部END-----------------------
'===========================搜索条地区判别END================================
		Rs.movenext                                        
	loop                                           
	Call CloseRs(rs)
TempText=TempText&"</SELECT>"

Dim rslx,sqllx,leixing,x
set rslx=server.createobject("adodb.recordset")
sqllx = "select leixing from DNJY_config "
rslx.open sqllx,conn,1,1
leixing=split(rslx("leixing"),"|")
TempText=TempText&"<select name=""leixing""><option value="""">交易类型</option>"
for x=0 to ubound(leixing)
TempText=TempText&"<option value="""&leixing(x)&""">"&leixing(x)&"</option>"
next
TempText=TempText&"</select>"
rslx.close
set rslx=Nothing

	TempText=TempText&" <INPUT  type=submit value=搜索 style=""font-size: 12px"" name=search></TD></TR></TABLE></FORM>"
	f_search=TempText
End Function


'都市114＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
Function ds114(c1,c2,c3,mh,zs)
'c1、c2、c3为地区，每行个数，总数
Dim rs114,ae,ak,dq,TempText
ae=0
ak=0
TempText="<table width=""98%"" border=""0"" cellpadding=""0"" cellspacing=""0"" bgcolor=""#FFFFFF"" align=""center""><tr>"
Call OpenConn
set rs114=server.CreateObject("adodb.recordset")
rs114.Open "select * from DNJY_114 where known=0 and dqsj>"&DN_Now&""&ttcc114&" order by hfje/xsts "&DN_OrderType&",id",conn,1,1
if rs114.eof Then
TempText=TempText&"<p><br>暂无"
IF c1>0 Then dq=conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
TempText=TempText&""&dq&"内容!"
Else
do while not rs114.eof
TempText=TempText&"<td valign=""top""><div class=""articleline114sy"">"
If rs114("co")<>"0" Then
TempText=TempText&"<font color=#"&rs114("co")&">"
End if
TempText=TempText&"<img src=""img/dot1.gif"">"&rs114("d_name")&":"&rs114("d_tel")&""
If rs114("co")<>"0" Then
TempText=TempText&"</font>"
End if
TempText=TempText&"</div></td>"
ae=ae+1
ak=ak+1
If ak=mh then
TempText=TempText& "</tr>"
ak=0
End If               
                rs114.movenext
				if ae>=zs then exit do
                Loop
                End if
                rs114.close
                set rs114=Nothing
                
TempText=TempText&"</tr></table>"
ds114=TempText
End Function

'便民服务＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
Function bmfw(c1,c2,c3,zs,mh,tabw,hg,zh,btw,bj,b,u,gd)
'c1、c2、c3为地区，总数（0为不限制），每行个数，表格宽度，行高，字号，标题长度，背景（0不要1要），粗体（0不要1要），下划线（0不要1要），更多（0不要1要）

Dim rsbm,ae,ak,dq,TempText,bb,bbb,uu,uuu,articleline5
ae=0
ak=0
If b=1 Then
bb="<b>"
bbb="</b>"
Else
bb=""
bbb=""
End If
If u=1 Then
uu="<u>"
uuu="</u>"
Else
uu=""
uuu=""
End If
articleline5=""
If bj=1 then articleline5="class=""articleline5"""
TempText="<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"" bgcolor=""#FFFFFF""><tr>"
Call OpenConn
set rsbm=server.CreateObject("adodb.recordset")
rsbm.Open "select * from DNJY_bmcs where bmname<>''"&ttccbmxx&" order by paixu "&DN_OrderType&",id",conn,1,1
if rsbm.eof Then
TempText=TempText&"<p><br>暂无"
IF c1>0 Then dq=conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
TempText=TempText&""&dq&"内容!"
Else
do while not rsbm.eof
TempText=TempText&"<td valign=""top""><div style=""height:"&hg&"px;width:"&tabw&"px; font-size:"&zh&"px;margin-left:2px;margin-top:3px;padding:4px;float:left"" "&articleline5&">"

TempText=TempText&""&bb&""&uu&"<a href="""&rsbm("link")&""" target=_blank>"&Left(rsbm("bmname"),btw)&"</a>"&uuu&""&bbb&""

TempText=TempText&"</div></td>"
ae=ae+1
ak=ak+1
If ak=mh then
TempText=TempText& "</tr>"
ak=0
End If               
                rsbm.movenext
				if ae>=zs And zs<>0 then exit do
                Loop
                End if
                rsbm.close
                set rsbm=Nothing
                
TempText=TempText&"</tr></table>"
If gd=1 Then TempText=TempText&  "<table width=""100%"" ><tr><td align=""right"" "&articleline5&"><A href=""bmcslist.asp?"&CT_ID&""">更多>>></a></td></tr></table>"
bmfw=TempText
End Function

'本站公告显示＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
function announce(c1,c2,c3,gd,hits,ma,d,h,l,btw,zt,hg)         
'参数依次：c1,c2,c3一级地区，二级地区，三级地区，gd启用固顶（0不固1固），hits点击数（0不显1显）,ma移动（0不显1显），d时间（0-15种格式，0不显），h显示条数，l列数，btw标题长度，zt字号，hg行高
	Dim id,title,ii,i,Sql,TempText,rs,dd
	Call OpenConn
    Set Rs=server.createobject("aDodb.recordset")               
	Sql="select * from DNJY_announce Where activity=1"&ttccnews&" order by "
	If gd=1 Then sql=sql&" Solidtop "&DN_OrderType&","
	if hits=1 Then
	sql=sql&" hits "&DN_OrderType&","
	End if
	sql=sql&" id "&DN_OrderType&""
	Rs.open Sql,conn,1,1
	TempText=""
if rs.eof Then
TempText=" <p><br>暂无"
IF c1>0 Then dd=conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
TempText=TempText&""&dd&"内容!"
Else
	IF ma=1 Then TempText= "<marquee scrollamount=""1"" scrolldelay=""50"" onmouseover=""this.stop()"" onmouseout=""this.start()"" direction=up style=""width:100%;padding:0"" height='95%'>"
	TempText=TempText&"<table width=""100%""  border=""0"" cellspacing=""2"" cellpadding=""0"" >"
	dim iii
	iii=0
	Do While not Rs.eof
i=i+1
	IF iii Mod l=0 then TempText=TempText&  "<tr height="""&hg&""">"
	TempText=TempText&  "<td>"
    IF rs("Solidtop")=1 And gd=1 Then TempText=TempText&  "<img src=""img/top4.gif"" width=""15"" height=""13"" border=0 alt=""固顶"">"
	TempText=TempText&  "<a href=""#"" ONCLICK=""window.open('../announce_show.asp?id="&rs("id")&"','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=500,height=450,left=300,top=20')""><SPAN style=""color:000000;font-size="&zt&"px"">"&CutStr(rs("title"),btw,"..")&"</SPAN></a> "
    IF hits=1 Then TempText=TempText&  " <font color=""#999999"" >["&rs("hits")&"]</font>" 
	TempText=TempText&  " <font color=""#999999"" >"&FormatDate(rs("addtime"),d)&"</font>"    
	TempText=TempText&  "</TD>"
	iii=iii+1
	IF iii Mod l=0 then TempText=TempText&  "</tr>"
    if i>=h then exit do 
	Rs.movenext
	Loop	
	IF iii Mod l<>0 then
	Do While not iii Mod l=0
	TempText=TempText&  "<td></td>"
	iii=iii+1
	Loop
	TempText=TempText&  "</tr>"
	End IF
	TempText=TempText&  "</table>"
	IF ma=1 Then TempText=TempText& "</marquee>"
	End IF
Call CloseRs(rs)
	announce=TempText
end Function


'滚动公告＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
sub gtgg()%>
<!---本站滚动公告开始---->
<div style="width: 590px;overflow: hidden; background-color: #Fff;" onmouseover="clearInterval(timer1);" onmouseout="go();"><div style="position: relative;top: 0px;left: 0px;white-space: nowrap;color: #FFF;" id="news"><span id="nbo" style="color: #CFC;">
<%Dim yss,dd
Call OpenConn
set rs=server.createobject("adodb.recordset")
'Sql="select id,title,hits,oColor,infotime,newstop,firstImageName,img,tuijian from DNJY_News where [bigcid]="&t&" "&lxx&" "&ttccnews&" order by newstop "&DN_OrderType&""
sql = "select top 5 * from DNJY_announce where activity=1"&ttccnews&" order by Solidtop "&DN_OrderType&""
rs.open sql,conn,3,3
if rs.eof Then
Response.Write "<img src=""img/icon1.gif"" border=""0""/><font color=""#000000"">暂无"
IF c1>0 Then dd=conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
Response.Write ""&dd&"内容-_-</font>"
else
do while not rs.eof
Response.Write "&nbsp;&nbsp;&nbsp;<img src=""img/icon1.gif"" border=""0""/><a href=""#"" ONCLICK=""window.open('../announce_show.asp?id="&rs("id")&"','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=500,height=450,left=300,top=20')"">"&rs("title")&"</a>" 
rs.movenext
Loop
End if
Call CloseRs(rs)
%></span><script language="javascript">document.write(nbo.innerHTML);   //重复一次新闻内容
document.write(nbo.innerHTML);   //重复一次新闻内容
document.write(nbo.innerHTML);   //重复一次新闻内容
document.write(nbo.innerHTML);   //重复一次新闻内容
</script></div></div><script language=javascript>function newsScroll() {   //实现不间断滚动
news.style.pixelLeft=(news.style.pixelLeft-1)%nbo.offsetWidth;}function go() { timer1=setInterval('newsScroll()',1)   //更改第二个参数可以改变速度，值越小，速度越快。
}go();</script>
<!---本站滚动公告结束---->
<%end Sub


'友情链接＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
Function f_link(c1,c2,c3,linktype,btw,hg,LineNum,MaxLine,LineLogo,LogoNum,WritingNum)
'c1、c2、c3分别为地区，linktype为显示方式（0为图文分开显示，1为综合显示,2为仅文字）,btw为网站名称长度,hg为高度，LineNum为每行显示个数,MaxLine为综合显示最多显示个数，LineLogo为综合显示LOGO显示行数，LogoNum为LOGO独显个数,WritingNum为文字独显个数
dim rsttlj,ttlj_i,ttlj_j,ttlj_k,ttlj_m,ttname,ttlogo,ttinfo,TempText,bb
if LogoNum=0 Then LogoNum=10000
If WritingNum=0 Then WritingNum=10000
If linktype=0 Or linktype=2 Then'图片文字分别进行
If linktype<>2 Then'图片文字分别进行
'图片部分
  ttlj_i=0
  ttlj_m=0
  ttlj_k=0
TempText= "<table border=""1"" align=""center"" cellspacing=""0"" width=""970"" bgcolor=""#CDE4CF"" bordercolorlight=""#008000"" bordercolordark=""#FFFFFF"" cellpadding=""1"">"
Call OpenConn
set rsttlj=server.createobject("adodb.recordset")
sql="select * from DNJY_link where known=1 and img=1 "&ttcclink&" order by linktop "&DN_OrderType&",paixu "&DN_OrderType&",id" 
rsttlj.open sql,conn,1,1
if rsttlj.eof then
TempText=TempText&"<p><br>暂无"
IF c1>0 Then
bb= conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
End IF
TempText=TempText&""&bb&"图片链接!"
end if
 
do while not rsttlj.eof and ttlj_m<LogoNum
TempText=TempText&"<tr>"
for ttlj_j=1 to LineNum
if rsttlj.eof then
TempText=TempText&"<td align=""center"" width=""970""> <center><a href=""user/addlink.asp?adminlink=2"" target=""_blank"">"
TempText=TempText&"<img src='images/your_link.gif' width=88 height=31 border=0 alt=""申请您的座位"">"
TempText=TempText&"</a></td>"
else
TempText=TempText&"<td width=""970"" align=""center"">"
if ttlj_i<LogoNum*LineNum  then'0000000-----------------------
 ttlogo=rsttlj("logo")
 ttinfo=rsttlj("info")
 ttname=rsttlj("wzname")
 TempText=TempText&"<center><a href="""&rsttlj("url")&""" title="""&rsttlj("wzname")&""" target=_blank>"
if ttlogo="0"  or ttlogo="images/wu.gif" then '如果没有图片显示默认图片===========
TempText=TempText&"<table width=""88"" height=""31""  background=""Images/nologo2.gif""><tr><td align=""center""><a href="""&rsttlj("url")&""" title="""&rsttlj("wzname")&""" target=_blank>"

if len(rsttlj("wzname"))>btw Then'判断标题长度
if rsttlj("c")=1 Then
TempText=TempText&"<font color=#FF0000>"
end If
TempText=TempText&""&Left(rsttlj("wzname"),btw)&""
if rsttlj("c")=1 Then
TempText=TempText&"</font >"
end If
Else
if rsttlj("c")=1 Then
TempText=TempText&"<font color=#FF0000>"
end If
TempText=TempText&""&rsttlj("wzname")&""
if rsttlj("c")=1 Then
TempText=TempText&"</font >"
end If
end If'判断标题长度end
TempText=TempText&"</a></td></tr></table>"

Else'如果有图片显示图片=======================
if ttlj_j = LogoNum * LineNum Then
TempText=TempText&"<a href=""user/link_gd.asp?"&C_ID&""" target=""_blank""><img src=""img/gdlj.gif"" width=88 height=31 border=0></a>"
Else
TempText=TempText&"<img src="""&rsttlj("logo")&""" width=88 height=31 border=0 alt="""&rsttlj("wzname")&""">"
end if
end if
TempText=TempText&"</a>"'========================
end If'-------------------------------
 ttlj_i=ttlj_i+1
 ttlj_k=ttlj_k+1
rsttlj.movenext
end if
Next

TempText=TempText&"</tr>"
ttlj_m=ttlj_m+1
loop
rsttlj.close
set rsttlj=nothing
TempText=TempText&"</table>"
'图片部分end
end if
'文字部分
  ttlj_i=0
  ttlj_m=0
  ttlj_k=0
TempText= TempText&"<table border=""1"" align=""center"" cellspacing=""0"" width=""970"" bgcolor=""#CDE4CF"" bordercolorlight=""#008000"" bordercolordark=""#FFFFFF"" cellpadding=""1"">"
set rsttlj=server.createobject("adodb.recordset")
sql="select top "&WritingNum&" * from DNJY_link where known=1 and img=0 "&ttcclink&" order by linktop "&DN_OrderType&",paixu "&DN_OrderType&",id" 
rsttlj.open sql,conn,1,1
if rsttlj.eof then
TempText=TempText&"<p><br>暂无"
IF c1>0 Then
bb= conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
End IF
TempText=TempText&""&bb&"文字链接!"
end if

do while not rsttlj.eof and ttlj_m<WritingNum
TempText=TempText&"<tr>"
for ttlj_j=1 to LineNum
if rsttlj.eof Then
TempText=TempText&"<td align=""center"" width=""970""> <center><a href=""user/addlink.asp?adminlink=2"" target=""_blank"">"
TempText=TempText&"您的位置"
TempText=TempText&"</a></td>"
else
TempText=TempText&"<td width=""970"" align=""center"">"
if ttlj_i<WritingNum*LineNum  then
 ttlogo=rsttlj("logo")
 ttinfo=rsttlj("info")
 ttname=rsttlj("wzname")
TempText=TempText&"<table><tr><td align=""center"">"
if ttlj_j = WritingNum * LineNum then
TempText=TempText&"<a href=""user/link_gd.asp?"&C_ID&""" target=""_blank"" >更多链接>></a>"
else
TempText=TempText&"<a href="""&rsttlj("url")&""" title="""&rsttlj("wzname")&"--"&rsttlj("info")&""" target=_blank>"
if len(rsttlj("wzname"))>btw Then
if rsttlj("c")=1 Then
TempText=TempText&"<font color=#FF0000>"
end If
TempText=TempText&""&Left(rsttlj("wzname"),btw)&""
if rsttlj("c")=1 Then
TempText=TempText&"</font>"
end If
Else
if rsttlj("c")=1 Then
TempText=TempText&"<font color=#FF0000>"
end If
TempText=TempText&""&rsttlj("wzname")&""
if rsttlj("c")=1 Then
TempText=TempText&"</font >"
end If
end If
TempText=TempText&"</a>"
end If
TempText=TempText&"</td></tr></table>"
 end if
 ttlj_i=ttlj_i+1
 ttlj_k=ttlj_k+1
rsttlj.movenext
end if
next
TempText=TempText&"</tr>"
ttlj_m=ttlj_m+1
loop
rsttlj.close
set rsttlj=nothing
TempText=TempText&"</table>"
'文字部分end
Else'综合方式开始
TempText= "<table border=""1"" cellspacing=""0"" width=""970"" bgcolor=""#CDE4CF"" bordercolorlight=""#008000"" bordercolordark=""#FFFFFF"" cellpadding=""1"">"
set rsttlj=server.createobject("adodb.recordset")
sql="select * from DNJY_link where known=1 "&ttcclink&" order by linktop "&DN_OrderType&",paixu "&DN_OrderType&",id" 
rsttlj.open sql,conn,1,1
if rsttlj.eof then
TempText=TempText&"<p><br>暂无"
IF c1>0 Then
bb= conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
End IF
TempText=TempText&""&bb&"链接!"
end if
  ttlj_i=0
  ttlj_m=0
  ttlj_k=0
if MaxLine=0 then
MaxLine=10000
end if
do while not rsttlj.eof and ttlj_m<MaxLine
TempText=TempText&"<tr>"
for ttlj_j=1 to LineNum
if rsttlj.eof then
TempText=TempText&"<td align=""center"" width=""970""> <center><a href=""user/addlink.asp?adminlink=2"" target=""_blank"">"
if ttlj_k=<LineLogo*LineNum then
TempText=TempText&"<img src='images/your_link.gif' width=88 height=31 border=0 alt=""申请您的座位"">"
else
TempText=TempText&"您的位置"
end if
TempText=TempText&"</a></td>"
else
TempText=TempText&"<td width=""970"" align=""center"">"
if ttlj_i<LineLogo*LineNum  then
 ttlogo=rsttlj("logo")
 ttinfo=rsttlj("info")
 ttname=rsttlj("wzname")
 TempText=TempText&"<center><a href="""&rsttlj("url")&""" title="""&rsttlj("info")&""" target=_blank>"
if ttlogo="0"  or ttlogo="images/wu.gif" then 
TempText=TempText&"<table width=""88"" height=""31""  background=""Images/nologo2.gif""><tr><td align=""center""  width=""88""><a href="""&rsttlj("url")&""" title="""&rsttlj("info")&""" target=_blank>"

if len(rsttlj("wzname"))>btw Then
if rsttlj("c")=1 Then
TempText=TempText&"<font color=#FF0000>"
end If
TempText=TempText&""&Left(rsttlj("wzname"),btw)&""
if rsttlj("c")=1 Then
TempText=TempText&"</font >"
end If
Else
if rsttlj("c")=1 Then
TempText=TempText&"<font color=#FF0000>"
end If
TempText=TempText&""&rsttlj("wzname")&""
if rsttlj("c")=1 Then
TempText=TempText&"</font >"
end If
end If

TempText=TempText&"</a></td></tr></table>"
Else
if ttlj_j = MaxLine * LineNum Then
TempText=TempText&"<a href=""user/link_gd.asp?"&C_ID&""" target=""_blank""><img src=""img/gdlj.gif"" width=88 height=31 border=0></a>"
Else
TempText=TempText&"<img src="""&rsttlj("logo")&""" width=88 height=31 border=0 alt="""&rsttlj("info")&""">"
end if
end if
TempText=TempText&"</a>"
else
TempText=TempText&"<table><tr><td align=""center"">"
if ttlj_j = MaxLine * LineNum then
TempText=TempText&"<a href=""user/link_gd.asp?"&C_ID&""" target=""_blank"" >更多链接>></a>"
else
TempText=TempText&"<a href="""&rsttlj("url")&""" title="""&rsttlj("wzname")&"--"&rsttlj("info")&""" target=_blank>"
if len(rsttlj("wzname"))>btw Then
if rsttlj("c")=1 Then
TempText=TempText&"<font color=#FF0000>"
end If
TempText=TempText&""&Left(rsttlj("wzname"),btw)&""
if rsttlj("c")=1 Then
TempText=TempText&"</font>"
end If
Else
if rsttlj("c")=1 Then
TempText=TempText&"<font color=#FF0000>"
end If
TempText=TempText&""&rsttlj("wzname")&""
if rsttlj("c")=1 Then
TempText=TempText&"</font >"
end If
end If
TempText=TempText&"</a>"
end If
TempText=TempText&"</td></tr></table>"
 end if
 ttlj_i=ttlj_i+1
 ttlj_k=ttlj_k+1
rsttlj.movenext
end if
next
TempText=TempText&"</tr>"
ttlj_m=ttlj_m+1
loop
rsttlj.close
set rsttlj=nothing

TempText=TempText&"</table>"
End if'综合部分end
f_link=TempText
End Function

'*************************************************
'标签名：DNJY_link
'作  用：文字友情链接调用
'参  数：c1,c2,c3-----地区
'        lx ----0为文字，1为图片
'        Num----显示个数,0为不限制
'        color ----文字颜色
'*************************************************
function DNJY_link(c1,c2,c3,lx,Num,color)
dim link,TempText,i,img,bb,wz
If lx=0 Then
img="img=0"
wz="文字"
elseIf lx=1 Then img="img=1"
wz="图片"
End If
Call OpenConn
set link=server.CreateObject("adodb.recordset")
link.open "select * from DNJY_link where known=1 and "&img&""&ttcclink&" order by linktop "&DN_OrderType&",paixu "&DN_OrderType&",id",conn,1,1
if link.eof and link.bof Then
TempText=TempText&"<p><br>暂无"
IF c1>0 Then
bb= conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
End IF
TempText=TempText&""&bb&wz&"链接!"
end if
do while not link.eof
If lx=1 Then
temptext=temptext& "<a href="""&link("Url")&""" title="""&link("info")&""" target=""_blank""><img src="""&link("logo")&""" width=""88"" height=""31""/></a>&nbsp;"
else
TempText = TempText & "<a href=""" & link("Url") & """ target=""_blank"" title=""" & link("info") & """><span style='color:#"&color&"'>" & link("wzname") & "</span></a> | "
End if
i=i+1
link.movenext
if i>=Num And Num>0 then exit do
Loop
if right(TempText,3)=" | " then
TempText=left(TempText,len(TempText)-3)
end if
link.close
set link=Nothing

DNJY_link=TempText
end function


Private Sub HomeHtml()'admin目录下生成静态首页
Dim beforeURL,mudi,homeFileStr
	beforeURL="http://"& request.ServerVariables("Server_NAME") & request.ServerVariables("SCRIPT_NAME")
	beforeURL=Left(beforeURL,InstrRev(beforeURL,"/") - 1)
	beforeURL=Left(beforeURL,InstrRev(beforeURL,"/"))
	homeFileStr = CD_GetCode(beforeURL &"index.asp","GB2312")
	If InStr(homeFileStr,"[DN"&"JY]")<1 Then homeFileStr="<script language='javascript' type='text/javascript'>document.location.href='index.asp';</script>"
	Call CD_WriteFile(homeFileStr & Chr(13) &"<!-- DNJY Home Html For "& Now() &" -->",Server.MapPath("../index.html"))
End Sub

Private Sub s_HomeHtml()'根目录下生成静态首页
Dim beforeURL,mudi,homeFileStr
	beforeURL="http://"& request.ServerVariables("Server_NAME") & request.ServerVariables("SCRIPT_NAME")
	beforeURL=Left(beforeURL,InstrRev(beforeURL,"/") - 1)
	beforeURL=Left(beforeURL,InstrRev(beforeURL,"/"))
	homeFileStr = CD_GetCode(beforeURL &"index.asp","GB2312")
	If InStr(homeFileStr,"[DN"&"JY]")<1 Then homeFileStr="<script language='javascript' type='text/javascript'>document.location.href='index.asp';</script>"
	Call CD_WriteFile(homeFileStr & Chr(13) &"<!-- DNJY Home Html For "& Now() &" -->",Server.MapPath("index.html"))
End Sub

%>

