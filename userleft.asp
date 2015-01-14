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
dim sdname,mpname,mpfg,sj,vip
username=request.cookies("dick")("username")
Call OpenConn
set rs=server.createobject("adodb.recordset")
sql = "select * from [DNJY_user] where username='"&username&"'"
rs.open sql,conn,1,1
sdname=rs("sdname")
mpname=rs("mpname")
mpfg=rs("mpfg")
vip=rs("vip")
if rs("vipdq")<>"" Then
sj=DateDiff("d",now(),""&rs("vipdq")&"")
end If
%>
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=title%>-我的状态</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
-->
</style>

<table width="214" border="0" cellPadding="0" cellSpacing="0" borderColor="#111111" style="border-collapse:collapse">
<td width="215" height="360" vAlign="top">
      <table width="210" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"><a href="<%=adduserinfo%>?<%=CT_ID%>"><img src="img/addxx.jpg" width=214 height=58 border=0></a></td>
        </tr>
      </table>
      <TABLE cellSpacing=0 cellPadding=0 border=0 class="ty1">        
        <TR>
          <TD align=middle  class="ty3" height=30><b>会 员 中 心</b></TD></TR>
        <TR>
          <TD width="210" height=200 valign="top">	
	<table width="80%" border="0" align="center" cellpadding="2" cellspacing="0" class="font_10_e_black">
      <tr>
        <td height="10" colspan="2"><%response.write "会员<font color=#FF0000>"&username&"</font>的管理中心<br>"%>
  <%if vip>0 then
  response.write "我的会员级别：<font color=#FF0000>VIP会员</font><br>"
if sj>0 then
response.Write "我的VIP会员资格<font color=#FF0000> 剩余"&sj&"</b>天</font><br>"
elseif sj=0 then
response.Write "我的VIP会员资格<font color=""#414141"">今日到期</font><br>"
elseif sj<0 then
response.Write "我的VIP会员资格<font color=""#FF0000""> 已经过期</font><br>"
end If
else
response.write "我的会员级别：<font color=#FF0000>普通会员</font><br>"
response.write "建议您点击升级为<a href='upgradevip.asp?"&C_ID&"'><font color=#0000ff>VIP会员</font></a>"
end if%> </td>
      </tr>

      <tr>
        <td align="center" valign="top">[ <a href="user.asp?<%=C_ID%>">我的状态</a></td>
        <td width="84"><div align="center"><a href="user_cwmx.asp?<%=C_ID%>" >财务明细</a> ]</div></td>
      </tr>
      <tr>
        <td align="center" valign="top"><div align="center">[ <a href="user_zhcz.asp?<%=C_ID%>">帐户充值</a></div></td>
        <td width="84"><div align="center"><a href="user_order.asp?<%=C_ID%>">订单管理</a> ]</div></td>
      </tr>

      <tr>
        <td align="center" valign="top"><div align="center">[ <a href="user_props.asp?<%=C_ID%>" >购买道具</a></div></td>
        <td width="84"><div align="center"><a href="props.asp?<%=C_ID%>">装备转换</a> ]</div></td>
      </tr>
      <tr>
        <td align="center" valign="top"><div align="center">[ <a href="user_xxgl.asp?<%=C_ID%>">信息管理</a></div></td>
        <td width="84"><div align="center"><a href="user_hfgl.asp?<%=C_ID%>">管理回复</a> ]</div></td>
      </tr>
      <tr>
        <td align="center" valign="top"><div align="center">[ <a href="UserNews_AddSelClass.asp?<%=C_ID%>">发布文章</a></div></td>
        <td width="84"><div align="center"><a href="UserNews_ModSelClass.asp?<%=C_ID%>">管理文章</a> ]</div></td>
      </tr>
       <tr>
        <td align="center" valign="top"><div align="center">[ <a href="messs.asp?<%=C_ID%>">接收短信</a></div></td>
        <td width="84"><div align="center"> <a href="messh.asp?<%=C_ID%>">发送短信</a> ]</div></td>
      </tr>	  
      <tr>
        <td align="center" valign="top"><div align="center">[ <% if sdname<>"" then %><a href="user/com.asp?com=<%=rs("username")%>&<%=C_ID%>" target="_blank">我的网店</a><%Else%><font color="#FF0000">暂无网店</font><%end if%></div></td>
        <td width="84"><div align="center"><% if mpname<>"" then %><a href="company.asp?username=<%=rs("username")%>&<%=C_ID%>" target="_blank">我的名片</a><%Else%><font color="#FF0000">暂无名片</font><%end if%> ]</div></td>
      </tr>   
      <tr>
        <td align="center" valign="top"><div align="center">[ <a href="usersdeditzl.asp?<%=C_ID%>">管理店铺</a></div></td>
        <td width="84"><div align="center"> <a href="usersdeditzl.asp?<%=C_ID%>">管理名片</a> ]</div></td>
      </tr> 
      <tr>
        <td align="center" valign="top"><div align="center">[ <a href="certificate.asp?<%=C_ID%>">证照管理</a></div></td>
        <td width="84"><div align="center"><% if sdname<>"" Or mpname<>"" then %><a href="comgg.asp?<%=C_ID%>">店企标志</a><%else%><font color="#FF0000">店企标志</font><%end if%> ]</div></td>
      </tr> 	  
 
      <tr>
        <td align="center" valign="top"><div align="center">[ <a href="user_data.asp?<%=C_ID%>">修改资料</a></div></td>
        <td width="84"><div align="center"><a href="user_pass.asp?<%=C_ID%>">修改密码</a> ]</div></td>
      </tr>

       <tr>
        <td align="center" valign="top"><div align="center">[ <a href="user_favorites.asp?<%=C_ID%>">信息收藏</a></div></td>
        <td width="84"><div align="center"> <a href="user_favshops.asp?<%=C_ID%>">店铺收藏</a> ]</div></td>
      </tr>
      <tr>
        <td align="center" valign="top"><div align="center">[ <a href="#" ONCLICK="window.open('user/contribute.asp','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=780,height=700,left=300,top=20')">我要投稿</a></div></td>
        <td width="84"><div align="center"><a href="manuscript.asp?<%=C_ID%>">已投稿件</a> ]</div></td>
      </tr>
      <tr>
        <td align="center" valign="top"><div align="center">[ <a href="Message_board.asp?mybook=ok&<%=C_ID%>">我的留言</a></div></td>
        <td width="84"><div align="center"><a href="usermail_loglist.asp?isse=0&<%=C_ID%>">邮件日志</a> ]</div></td>
      </tr>
      <tr>
        <td align="center" valign="top"><div align="center">[ <a href="user_link.asp?<%=C_ID%>">友情链接</a></div></td>
        <td width="84"><div align="center"><a href="userrecommend.asp">推荐会员</a> ]</div></td>
      </tr>
	  <%If  isnull(rs("QQopenid")) Then
	  session("binding")="yes"%>
      <tr>
        <td align="center" valign="top" colspan="2"><div align="center"><a title="使用QQ号快速登录" href="/<%=strInstallDir%>api/qq/redirect_to_login.asp"  target="_blank"><font color=#FF0000><b>绑定QQ号快捷登录</b></font></a></td>
      </tr>
	  <%End if%>
      <tr>
        <td height="10" colspan="2"><div align="center"><a href="user_sk.asp"><font color=#FF0000><b>本站收款帐号</b></font></a></div></td>
      </tr>
      <tr>
        <td align="center" valign="top" colspan="2"><div align="center"><a href="user/userout.asp?<%=C_ID%>"><b>安全退出</b></a></td>
      </tr>    </table></td>
    
  </tr>

</table>
