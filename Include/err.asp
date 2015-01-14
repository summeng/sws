<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%sub errdick(id)%>
<title>错误页面</title>
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
<link href="/<%=strInstallDir%>css/css.CSS" rel="stylesheet" type="text/css" />
<div align="center">
  <center>


<TABLE border=0 align="center" cellSpacing=0>
  <TBODY>
    <TR>
      <TD  background="/<%=strInstallDir%>images/ereer.gif" width="750" height="450" align="center">
	  <TABLE height="150">
	  <TR>
		<TD> </TD>
	  </TR>
	  </TABLE>
<TABLE  cellSpacing=0 cellPadding=0 border=0 >
          <TBODY>
            <TR><TD width="250"> &nbsp;</td>
<TD width="300">
<table cellSpacing="0" cellPadding="0" width="500" border="0" style="border-collapse: collapse" bordercolor="#111111">
  <tr>
    <td width="300">
    <table cellSpacing="0" cellPadding="3" width="100%" border="0">
      <tr>
        <td width="100%">　</td>
      </tr>
      <tr>
        <td>
<%'客服QQ此文件不需要作任何修改
Dim n,ppp,qq,mymail,qq_name
Call OpenConn
ppp=Conn.Execute("Select WebSetting From [DNJY_config]")(0)
ppp=Split(ppp,"|")
qq=ppp(10)
mymail=ppp(4)

if qq<>"" Then 
qq=replace(qq,"，",",")
if isnull(qq_name) or qq_name="" then qq_name=qq
qq_name=replace(qq_name,"，",",")
qq_name=split(qq_name,",")
qq=split(qq,",")
end If%>
          <span class="STYLE1">
          <%
if id=0 then
response.write "<li>非法提交！"
elseif id=1 then
response.write "<li>邮件地址错误！"
elseif id=2 then
response.write "<li>非法字符,请重新输入！"
elseif id=3 then
response.write "<li>不能使用中文用户名！"
elseif id=4 then
response.write "<li>内容输入不完整！"
elseif id=5 then
response.write "<li>邮政编码输入错误！"
elseif id=6 then
response.write "<li>身份证号码填写错误！"
elseif id=7 then
response.write "<li>用户不存在或密码错误！"
elseif id=8 then
response.write "<li>错误次数太多，程序已经停止对你的审核！"
elseif id=9 then
response.write "<li>操作错误，找不到用户资料！"
elseif id=10 then
response.write "<li>一级分类没有选择！"
elseif id=11 then
response.write "<li>二级分类没有选择！"
elseif id=12 then
response.write "<li>发布信息时标题不能为空！"
elseif id=13 then
response.write "<li>标题中含有非法或敏感字符，请注意检查！"
elseif id=14 then
response.write "<li>价格输入错误！"
elseif id=15 then
response.write "<li>交易地区错误，并且需要选择二级地名，否则用户搜索不到你交易的物品！"
elseif id=16 then
response.write "<li>发布信息的说明没有填写！"
elseif id=17 then
response.write "<li>联系人必须填写，你可以更改为其它名字！"
elseif id=18 then
response.write "<li>EMAIL必须填写，你可以更改为其它邮件地址！"
elseif id=19 then
response.write "<li>请检查你的邮件地址！"
elseif id=20 then
response.write "<li>联系方式/电话不能为空！"
elseif id=21 then
response.write "<li>标题过长，不能多于30个字符！"
elseif id=22 then
response.write "<li>你输入的用户名已经注册过，如果你已经注册了，请找回密码！"
elseif id=23 then
response.write "<li>你输入的邮件地址已经注册过，如果你已经注册了，请找回密码！"
elseif id=24 then
response.write "<li>未知登陆错误！"
elseif id=25 then
response.write "<li>你没有足够的标题变色道具，请不要修改程序的运行参数！"
elseif id=26 then
response.write "<li>你没有足够的信息置顶道具，请不要修改程序的运行参数！"
elseif id=27 then
response.write "<li>你没有足够的信息贴图道具，请不要修改程序的运行参数！"
elseif id=28 then
response.write "<li>你没有足够的信息验证道具，请不要修改程序的运行参数！"
elseif id=29 then
response.write "<li>系统保护：你提交数据太快，系统中止运行，请等待5分钟！"
elseif id=30 then
response.write "<li>密码不一致或错误！"
elseif id=31 then
response.write "<li>请输入你的QQ号码！"
elseif id=32 then
response.write "<li>请输入你的详细地址！"
elseif id=33 then
response.write "<li>登陆帐号不能包含禁用或敏感字符！"
elseif id=34 then
response.write "<li>订阅资讯必须填写正确邮箱！"
elseif id=35 Then
 response.write "提示：您的IP地址<font color=red> "&Request.serverVariables("REMOTE_ADDR")&" </font>涉嫌非法猜解会员密码，已被系统限制访问，请与管理员联系，并告知被锁定的IP地址。<br><br>客服信箱,点击发信：<img src=/"&strInstallDir&"images/email.gif border=0 align=center> <a href='mailto:"&mymail&"?subject=我的IP "&Request.serverVariables("REMOTE_ADDR")&" 被"&webname&"锁定'>"&mymail&"</a><br>客服QQ,点击临时会话："
for n=0 to UBound(qq)
 response.write "<a class='c' target=blank href='tencent://message/?uin="&qq(n)&"&Site="&webname&"&Menu=yes'>"&qq(n)&"</a> "
Next

 elseif id=36 Then
  response.write "提示：您的IP地址<font color=red> "&Request.serverVariables("REMOTE_ADDR")&" </font>涉嫌灌水，已被系统限制访问，请与管理员联系，并告知被锁定的IP地址。<br><br>客服信箱,点击发信：<img src=/"&strInstallDir&"images/email.gif border=0 align=center> <a href='mailto:"&mymail&"?subject=我的IP "&Request.serverVariables("REMOTE_ADDR")&" 被"&webname&"锁定'>"&mymail&"</a><br>客服QQ,点击临时会话："
for n=0 to UBound(qq)
 response.write "<a class='c' target=blank href='tencent://message/?uin="&qq(n)&"&Site="&webname&"&Menu=yes'>"&qq(n)&"</a> "
Next
 elseif id=37 Then
 response.write "提示：您的IP地址<font color=red> "&Request.serverVariables("REMOTE_ADDR")&" </font>涉嫌SQL注入危害本站，已被系统限制访问，请与管理员联系，并告知被锁定的IP地址。<br><br>客服信箱,点击发信：<img src=/"&strInstallDir&"images/email.gif border=0 align=center> <a href='mailto:"&mymail&"?subject=我的IP "&Request.serverVariables("REMOTE_ADDR")&" 被"&webname&"锁定'>"&mymail&"</a><br>客服QQ,点击临时会话："
for n=0 to UBound(qq)
 response.write "<a class='c' target=blank href='tencent://message/?uin="&qq(n)&"&Site="&webname&"&Menu=yes'>"&qq(n)&"</a> "
Next 
elseif id=38 then
response.write "<li>网址不用带http://！"
elseif id=39 then
response.write "<li>数字验证码不对！"
elseif id=40 then
response.write "<li>文字验证码不对！"
elseif id=41 then
response.write "<li>星期验证不对！"
elseif id=42 then
response.write "<li>你发布的信息或曾经有违反本站规则之处！"
elseif id=43 Then
 response.write "提示：您的IP地址<font color=red> "&Request.serverVariables("REMOTE_ADDR")&" </font>已被系统限制发布信息，请与管理员联系，并告知被锁定的IP地址。<br><br>客服信箱,点击发信：<img src=/"&strInstallDir&"images/email.gif border=0 align=center> <a href='mailto:"&mymail&"?subject=我的IP "&Request.serverVariables("REMOTE_ADDR")&" 被"&webname&"锁定'>"&mymail&"</a><br>客服QQ,点击临时会话："
for n=0 to UBound(qq)
 response.write "<a class='c' target=blank href='tencent://message/?uin="&qq(n)&"&Site="&webname&"&Menu=yes'>"&qq(n)&"</a> "
Next
elseif id=44 then
response.write "<li>问答验证不对！"
elseif id=45 then
response.write "<li>算式验证码答案不对！"
elseif id=46 Then
 response.write "提示：您的IP地址<font color=red> "&Request.serverVariables("REMOTE_ADDR")&" </font>已被系统限制访问，请与管理员联系，并告知被锁定的IP地址。<br><br>客服信箱,点击发信：<img src=/"&strInstallDir&"images/email.gif border=0 align=center> <a href='mailto:"&mymail&"?subject=我的IP "&Request.serverVariables("REMOTE_ADDR")&" 被"&webname&"锁定'>"&mymail&"</a><br>客服QQ,点击临时会话："
for n=0 to UBound(qq)
 response.write "<a class='c' target=blank href='tencent://message/?uin="&qq(n)&"&Site="&webname&"&Menu=yes'>"&qq(n)&"</a> "
Next
elseif id=47 then
response.write "<li>联系人或地址不能包含禁用或敏感字符！"
elseif id=48 then
response.write "<li>联系地址不能包含禁用字符！"
elseif id=49 then
response.write "<li>信息详细内容不能包含禁用或敏感字符！"
elseif id=50 then
response.write "<li>网站当前只读状态，不能注册会员、发布信息和留言等！若要紧急发布信息请点击QQ临时会话联系客服："
for n=0 to UBound(qq)
 response.write "<a class='c' target=blank href='tencent://message/?uin="&qq(n)&"&Site="&webname&"&Menu=yes'>"&qq(n)&"</a> "
Next

end if
%>
          </span></td>
      </tr>
      <tr>
        <td>　</td>
      </tr>
      <tr>
        <td><div align="center"><a href="javascript:history.back(-1);">
          <img height="15" src="/<%=strInstallDir%>images/_back.gif" width="56" border="0">返回</a></div></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
  </center>
</div>
<%end sub%>