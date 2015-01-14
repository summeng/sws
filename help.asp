<!--#include file="conn.asp"-->
<!--#include file="top.asp"-->
<!--#include file="SqlIn.Asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%Dim Data
Call OpenConn
ID=Clng(Strint(Request.QueryString("ID")))
Set Rs=Conn.Execute("Select Text From DNJY_help Where OneID="&city_oneid&" and TwoID="&city_twoid&" and ThreeID="&city_threeid&"")
If Not Rs.Eof Then
	Data=Split(Rs(0),"∮∮∮")
Else
Data=Split(Application(CacheName&"_help"),"∮∮∮")
End If
Call CloseRs(rs)
Call CloseDB(conn)%>
<link href="Css/help.css" rel="stylesheet" type="text/css" />
<div class="content xw">
	<div class="l leftBar">
		<ul class="listMenu">
			<li <%If ID=0 Then%>class="current"<%End If%>><a href="/<%=strInstallDir%>help.asp?ID=0&<%=C_ID%>">关于本站</a></li>
			<li <%If ID=1 Then%>class="current"<%End If%>><a href="/<%=strInstallDir%>help.asp?id=1&<%=C_ID%>">会员帮助</a></li>
			<li <%If ID=2 Then%>class="current"<%End If%>><a href="/<%=strInstallDir%>help.asp?id=2&<%=C_ID%>">服务条款</a></li>
			<li <%If ID=3 Then%>class="current"<%End If%>><a href="/<%=strInstallDir%>help.asp?id=3&<%=C_ID%>">免责声明</a></li>
			<li <%If ID=4 Then%>class="current"<%End If%>><a href="/<%=strInstallDir%>help.asp?id=4&<%=C_ID%>">广告服务</a></li>
			<li <%If ID=5 Then%>class="current"<%End If%>><a href="/<%=strInstallDir%>help.asp?id=5&<%=C_ID%>">联系我们</a></li>

		</ul>
	</div>
	<div class="r rightTxt">
	<%Select Case ID
	Case 0%>
		<div class="view">
			<h2 style="line-height:170%">关于本站</h2>
			<div class="txt"><%=Data(0)%></div>
		</div>
		<%CAse 1%>
		<div class="view">
			<h2 style="line-height:170%">会员帮助</h2>
			<div class="txt"><%=Data(1)%></div>
		</div>
<TABLE>
<TR>
	<TD height=10></TD>
</TR>
</TABLE>
<TABLE style="BORDER-COLLAPSE: collapse" border=0 cellSpacing=0 borderColor=#111111 cellPadding=0 width=700 height=240>
<TBODY>
<TR>
<TD bgColor=#ffcc33 height=5 width="100%">
<DIV class=bd align=center><IMG border=0 src="img/vip1.jpg" width=760 height=240></DIV></TD></TR></TBODY></TABLE>
<TABLE style="BORDER-BOTTOM: #cc9900 1px dotted; BORDER-LEFT: #cc9900 1px dotted; BORDER-TOP: #cc9900 1px dotted; BORDER-RIGHT: #cc9900 1px dotted" id=table1 border=0 cellSpacing=1 width="100%">
<TBODY>
<TR>
<TD>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><B>① <U>本站功能</U></B></P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px">在本站您可发布查询供求信息、申请网络店铺、发布企业名片等</P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><B>② <U>注册用户</U></B></P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px">本站用户分为<FONT color=#ff0000><B>游客</B></FONT>、<FONT color=#ff0000><B>普通会员</B></FONT>和<FONT color=#ff0000><B>VIP会员</B></FONT>三种。游客是指不注册，直接发布和查询信息的网友。免费注册即可成为本站普通会员，可使用道具功能，拥有强大的后台管理平台。交纳会员费即可成为vip会员，除拥有普通会员权限外，信用度更高，可开店铺、企业名片、信息免审核、获得本站赠送
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
等。具体条件和权限见下表：</P></TD></TR></TBODY></TABLE>
<DIV align=center>
<TABLE style="BORDER-COLLAPSE: collapse" border=1 cellSpacing=0 borderColor=#fff2d7 cellPadding=0 width="98%">
<TBODY>
<TR>
<TD width="8%">
<DIV align=center><FONT color=#b548e1>会员级别</DIV></FONT></TD>
<TD width="8%">
<DIV align=center><FONT color=#b548e1>付人民币 (元/月)</DIV></FONT></TD>
<TD width="5%">
<DIV align=center><FONT color=#b548e1>道具使用</DIV></FONT></TD>
<TD width="5%">
<DIV align=center><FONT color=#b548e1>上传图片</DIV></FONT></TD>
<TD width="8%">
<DIV align=center><FONT color=#b548e1>开店铺名片</DIV></FONT></TD>
<TD width="8%">
<DIV align=center><FONT color=#b548e1>店铺商品展示(条)</DIV></FONT></TD>
<TD width="8%">
<DIV align=center><FONT color=#b548e1>店铺推荐</DIV></FONT></TD>
<TD width="8%">
<DIV align=center><FONT color=#b548e1>个人信息显示(条)</DIV></FONT></TD>
<TD width="6%">
<DIV align=center><FONT color=#b548e1>HTM代码过虑</DIV></FONT></TD>
<TD width="6%">
<DIV align=center><FONT color=#b548e1>信息审验</DIV></FONT></TD>
<TD width="6%">
<DIV align=center><FONT color=#b548e1>信息推荐</DIV></FONT></TD>
<TD width="6%">
<DIV align=center><FONT color=#b548e1>发布企业名片</DIV></FONT></TD>
<TD width="8%">
<DIV align=center><FONT color=#b548e1>发布竞价信息</DIV></FONT></TD>
<TD width="4%">
<DIV align=center><FONT color=#b548e1>修改信息</DIV></FONT></TD>
<TD width="6%">
<DIV align=center><FONT color=#b548e1>获赠送
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
 </DIV></FONT></TD></TR>
<TR>
<TD width="8%">
<DIV align=center><FONT color=#008080>游 客</FONT></DIV>
<TD width="8%">
<DIV align=center>×</DIV></TD>
<TD width="5%">
<DIV align=center>×</DIV></TD>
<TD width="5%">
<DIV align=center>×</DIV></TD>
<TD width="8%">
<DIV align=center>×</DIV></TD>
<TD width="8%">
<DIV align=center>×</DIV></TD>
<TD width="8%">
<DIV align=center>×</DIV></TD>
<TD width="8%">
<DIV align=center>最新
<SCRIPT src="user/helppz.asp?action=huiyuanxx"></SCRIPT>
</DIV></TD>
<TD width="6%">
<DIV align=center>√</DIV></TD>
<TD width="8%">
<DIV align=center>▲</DIV></TD>
<TD width="6%">
<DIV align=center>×</DIV></TD>
<TD width="6%">
<DIV align=center>×</DIV></TD>
<TD width="8%">
<DIV align=center>×</DIV></TD>
<TD width="4%">
<DIV align=center>×</DIV></TD>
<TD width="6%">
<DIV align=center>×</DIV></TD></TR>
<TR>
<TD width="8%">
<DIV align=center><FONT color=#008080>普通会员</FONT></DIV></TD>
<TD width="8%">
<DIV align=center>×</DIV></TD>
<TD width="5%">
<DIV align=center>√</DIV></TD>
<TD width="5%">
<DIV align=center>√</DIV></TD>
<TD width="8%">
<DIV align=center>√</DIV></TD>
<TD width="8%">
<DIV align=center>最新
<SCRIPT src="user/helppz.asp?action=huiyuansp"></SCRIPT>
</DIV></TD>
<TD width="8%">
<DIV align=center>×</DIV></TD>
<TD width="8%">
<DIV align=center>最新
<SCRIPT src="user/helppz.asp?action=huiyuanxx"></SCRIPT>
</DIV></TD>
<TD width="6%">
<DIV align=center>×</DIV></TD>
<TD width="6%">
<DIV align=center>▲</DIV></TD>
<TD width="6%">
<DIV align=center>×</DIV></TD>
<TD width="6%">
<DIV align=center>×</DIV></TD>
<TD width="8%">
<DIV align=center>√</DIV></TD>
<TD width="4%">
<DIV align=center>√</DIV></TD>
<TD width="6%">
<DIV align=center>×</DIV></TD></TR>
<TR>
<TD width="8%">
<DIV align=center><FONT color=#008080>VIP会员</FONT></DIV></TD>
<TD width="8%">
<DIV align=center>
<SCRIPT src="user/helppz.asp?action=vipje"></SCRIPT>
</DIV></TD>
<TD width="5%">
<DIV align=center>√</DIV></TD>
<TD width="5%">
<DIV align=center>√</DIV></TD>
<TD width="8%">
<DIV align=center>√专页</DIV></TD>
<TD width="8%">
<DIV align=center>不限</DIV></TD>
<TD width="8%">
<DIV align=center>√</DIV></TD>
<TD width="8%">
<DIV align=center>不限</DIV></TD>
<TD width="6%">
<DIV align=center>×</DIV></TD>
<TD width="6%">
<DIV align=center>×</DIV></TD>
<TD width="6%">
<DIV align=center>√</DIV></TD>
<TD width="6%">
<DIV align=center>√</DIV></TD>
<TD width="8%">
<DIV align=center>√</DIV></TD>
<TD width="4%">
<DIV align=center>√</DIV></TD>
<TD width="6%">
<DIV align=center>√</DIV></TD></TR></TBODY></TABLE></DIV>注： a、带▲为当系统开启审核制度时要审核。审核制度分级进行。 b、开店铺或企业名片需身份验证，必须向本站提供个人身份证复印件，且有固定电话联系方式，由本站验证后备案，并在网站上显示为验证会员。 c、首页最新信息栏按“置顶&gt;发布时间”优先顺序；首页中部方框信息按“置顶&gt;竞价高低&gt;发布时间”优先顺序，其中只有竞价信息和VIP会员信息才能显示图片。
<P>
<TABLE style="BORDER-COLLAPSE: collapse" border=0 cellSpacing=0 borderColor=#111111 cellPadding=0 width="100%" height=17>
<TBODY>
<TR>
<TD height=13 vAlign=top width="31%">
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><B>③ <U>发布信息</U></B></P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px">点击[发布信息]，提醒登陆发布，登陆→发布信息即可；也可以选择我要快速免费发布发布信息，但其信息需要通过审核,且不能上传图片。</P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><B>④ <U>取消信息</U></B></P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px">若你的信息不考虑生效，可登录后自助删除，已完成交易的，登录后，进入信息管理－点击管理－点击修改，将该条信息设置为“已完成交易”。已成交的信息将不再显示，以免信息遗留受干扰。</P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><B>⑤ <U>竞价信息</U></B></P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px">竞价信息可在显眼位置显示。发布竞价信息时，在您的帐户有足够金额的前提下，选择发布天数，填写竞价总金额，系统自动扣除费用。每天均价越高，所发布的信息越靠前。</P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><B>⑥ <U>道具功能介绍</U></B></P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px">本站使用四个道具，即标题变色道具（改变帖子主题颜色，吸引眼球），信息置顶道具（<FONT face=宋体>能发布信息置顶，有效时间为24小时；使用个数越多，位置越高</FONT>；个数相同，时间越早，位置越高），内容贴图道具（<FONT face=宋体>可以发和信息相关的图片</FONT>），通过验证道具（<FONT face=宋体>可不通过管理审核，直接发布</FONT>）。四种道具需用
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
购买，属可选收费项目。详见下表。</FONT></P></TD>
<TD height=13 vAlign=top width="69%"><A href="upgradevip.asp"><IMG border=0 src="images/vip3.gif"></A></TD></TR></TBODY></TABLE>
<DIV align=center>
<TABLE style="BORDER-COLLAPSE: collapse" border=1 cellSpacing=0 borderColor=#fff2d7 cellPadding=0 width="98%">
<TBODY>
<TR>
<TD>
<DIV align=center>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px">道具名</P></DIV></TD>
<TD>
<DIV align=center>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px">实现功能</P></DIV></TD>
<TD>
<DIV align=center>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px">需
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
</P></DIV></TD></TR>
<TR>
<TD width="24%">
<DIV align=center>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><FONT color=#008080>标题变色道具</FONT></FONT></P></DIV></TD>
<TD width="60%">
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><FONT face=宋体>改变标题颜色</FONT></P></TD>
<TD width="15%">
<DIV align=center>
<SCRIPT src="user/helppz.asp?action=g_a"></SCRIPT>
</DIV></TD></TR>
<TR>
<TD>
<DIV align=center>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><FONT color=#008080>信息置顶道具</FONT></FONT></P></DIV></TD>
<TD>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><FONT face=宋体>能使发布信息置顶，有效时间为1个置顶/天；使用个数越多，位置越高。个数相同，时间越早，位置越高。</FONT></P></TD>
<TD>
<DIV align=center>
<SCRIPT src="user/helppz.asp?action=g_b"></SCRIPT>
</DIV></TD></TR>
<TR>
<TD>
<DIV align=center>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><FONT color=#008080>内容贴图道具</FONT></FONT></P></DIV></TD>
<TD>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><FONT face=宋体>可以发布和信息相关的图片</FONT></P></TD>
<TD>
<DIV align=center>
<SCRIPT src="user/helppz.asp?action=g_c"></SCRIPT>
</DIV></TD></TR>
<TR>
<TD>
<DIV align=center>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><FONT color=#008080>通过验证道具</FONT></FONT></P></DIV></TD>
<TD>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><FONT face=宋体>可不通过管理审核，直接发布</FONT></P></TD>
<TD>
<DIV align=center>
<SCRIPT src="user/helppz.asp?action=g_d"></SCRIPT>
</DIV></TD></TR></TBODY></TABLE></DIV>
<TABLE style="BORDER-COLLAPSE: collapse" border=1 cellSpacing=0 borderColor=#fff2d7 cellPadding=0 width="98%">
<TBODY>
<TR>
<TD>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><B>⑦ <U>积分、
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
、道具怎么来？</U></B></P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><br>▲积分增减。用户积分是用户登录本站的次数、发贴次数、在本站消费的
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
数量、推荐注册本站会员的贡献积分，由系统自动记分，具体： 注册获赠
<SCRIPT src="user/helppz.asp?action=jf_1"></SCRIPT>
点积分 发布一条信息得
<SCRIPT src="user/helppz.asp?action=jf_2"></SCRIPT>
点积分，删除一条信息扣
<SCRIPT src="user/helppz.asp?action=jf_2"></SCRIPT>
点积分，管理员删除信息加倍扣分。
新闻资讯投稿一篇得
<SCRIPT src="user/helppz.asp?action=tgjf"></SCRIPT>
点积分，删除新闻资讯扣
<SCRIPT src="user/helppz.asp?action=tgjf"></SCRIPT>点积分。
 登陆一次得
<SCRIPT src="user/helppz.asp?action=jf_3"></SCRIPT>
点积分 消费
<SCRIPT src="user/helppz.asp?action=jf_4"></SCRIPT>
个
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
得1点积分 ，推荐注册本站会员每个获得<SCRIPT src="user/helppz.asp?action=jf_5"></SCRIPT>点积分。<br>▲用
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
可购买道具。
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
是通过注册会员获赠、积分转换或人民币购买。具体： 免费注册本站会员可获赠
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>

<SCRIPT src="user/helppz.asp?action=z_hb"></SCRIPT>
个 你可以用
<SCRIPT src="user/helppz.asp?action=jf_hb"></SCRIPT>
点积分转换为一个
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
 您可花1元人民币购买
<SCRIPT src="user/helppz.asp?action=rmb_hb"></SCRIPT>
个
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
 。<br>▲用道具能改变标题颜色、信息置顶、内容贴图、免审核验证。道具是通过注册会员获赠或用
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
购买。具体： 注册获赠标题变色道具
<SCRIPT src="user/helppz.asp?action=z_a"></SCRIPT>
个 注册获赠信息置顶道具
<SCRIPT src="user/helppz.asp?action=z_b"></SCRIPT>
个 注册获赠内容贴图道具
<SCRIPT src="user/helppz.asp?action=z_c"></SCRIPT>
个 注册获赠通过验证道具
<SCRIPT src="user/helppz.asp?action=z_d"></SCRIPT>
个 一个标题变色道具价格：
<SCRIPT src="user/helppz.asp?action=g_a"></SCRIPT>
个
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
 一个信息置顶道具价格：
<SCRIPT src="user/helppz.asp?action=g_b"></SCRIPT>
个
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
 一个内容贴图道具价格：
<SCRIPT src="user/helppz.asp?action=g_c"></SCRIPT>
个
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
 一个通过验证道具价格：
<SCRIPT src="user/helppz.asp?action=g_d"></SCRIPT>
个
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>。

<P></P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px">▲道具使用：发布信息时可直接使用，如果道具不足，可登陆本站→到用户中心→
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
兑换道具（
<SCRIPT src="user/helppz.asp?action=rmb_mc"></SCRIPT>
不足可积分转换或用人民币购买）→用道具发布和管理您的个性信息。 
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px"><B>⑧ <U>个人资金帐户及充值</U></B></P>
<P style="LINE-HEIGHT: 150%; MARGIN: 0px 15px">▲每个会员都有一个人民币资金帐户，可对帐户进行充值，并利用帐户资金在本站进行消费。帐户资金只能在本站消费，不得取现金或退款。</P></TD></TR></TBODY></TABLE>

		<%CAse 2%>
		<div class="view">
			<h2 style="line-height:170%">服务条款</h2>
			<div class="txt"><%=Data(2)%></div>
		</div>
		<%Case 3%>
		<div class="view">
			<h2 style="line-height:170%">免责声明</h2>
			<div class="txt"><%=Data(3)%></div>
		</div>
		<%Case 4%>
		<div class="view">
			<h2 style="line-height:170%">广告服务</h2>
			<div class="txt"><%=Data(4)%></div>
		</div>
		<%Case Else %>
		<div class="view">
			<h2 style="line-height:170%">联系我们</h2>
			<div class="txt"><%=Data(5)%></div>
		</div>
		<%End Select%>
	</div>
	<span class="cf"></span>
</div>

<!--#include file="footer.asp"-->
</body>
</html>
