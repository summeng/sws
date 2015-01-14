<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file=top.asp-->
<!--#include file="Include/err.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%if lmkg2="1" then
call errdick(50)
Call CloseDB(conn)
response.end
end If
Dim tjname
tjname=trim(request("tjname"))%>
<link rel="stylesheet" href="/<%=strInstallDir%>css/css.css" type="text/css">

<center>
<TABLE width=980 border=0 cellpadding=3 cellspacing=1 align=center  bgcolor=#FFFFFF    class="grayline"  background="/img/hd_normal_bg3.gif">
<tr><td height=30>目前位置：<a href=index.asp?<%=C_ID%>>首页</a> > 会员注册第一步</td></tr>
</TABLE>
<TABLE width=980 border=0 cellpadding=3 cellspacing=1 align=center  bgcolor=#FFFFFF height=300    class="grayline">

<tr><td>
<table border=0 width=547 align=center cellpadding=0 cellspacing=0  background=images/regbg2.gif>
<tr><td width=550 height=120><img src=img/reg.jpg border=0></td></tr>
<tr><td width=547 height=60 valign=middle bgcolor=#FFFFFF>
<br>· 您在本站注册，视为对本站各种条款、规则、协议、声明、约定、公告的无条件接受和认可<br><br>
</td></tr>
<tr><td width=547 height=12><img src=images/regbg1.gif border=0></td></tr>
<tr><td width=547 height=250  background=images/regbg2.gif>
　欢迎您注册为本站会员，您在访问本站（含论坛）时，请您自觉遵守以下条款。<br>

一、不得利用本站危害国家安全、泄露国家秘密，不得侵犯国家社会集体的和公民的合法权益，不得利用本站制作、复制和传播下列信息：<br> 
　　（一）煽动抗拒、破坏宪法和法律、行政法规实施的；<br>
　　（二）煽动颠覆国家政权，推翻社会主义制度的；<br>
　　（三）煽动分裂国家、破坏国家统一的；<br>
　　（四）煽动民族仇恨、民族歧视，破坏民族团结的；<br>
　　（五）捏造或者歪曲事实，散布谣言，扰乱社会秩序的；<br>
　　（六）宣扬封建迷信、淫秽、色情、赌博、暴力、凶杀、恐怖、教唆犯罪的；<br>
　　（七）公然侮辱他人或者捏造事实诽谤他人的，或者进行其他恶意攻击的；<br>
　　（八）损害国家机关信誉的；<br>
　　（九）其他违反宪法和法律行政法规的；<br>
　　（十）发布虚假、伪劣商品及其它服务信息。<br>

二、互相尊重，对自己的言论和行为负责。<br>
三、同意并遵守以下附件中的<%=title%>用户协议及其相关的条款和条件。<br>
<br>附件：<A href="help.asp?id=2&<%=C_ID%>" target=_blank><%=title%>用户协议</A>及其相关的条款和条件。</td></tr>

<tr><td width=547 height=12><img src=images/regbg3.gif border=0></td></tr></table>
<%Response.Cookies("reg")("regkey")="ok" %>
<table border=0 width=547 align=center cellpadding=0 cellspacing=0>
<tr><td height="100" align="center">
<a href=<%=Regto%>?tjname=<%=tjname%>&<%=C_ID%>><img border=0 src="images/agree.gif" title="我同意"></a>
&nbsp;&nbsp;&nbsp;&nbsp;
<a href=index.asp?<%=C_ID%>><img border=0 src="images/agree_no.gif" title="不同意，再考虑一下"></a>
</form>
</td>
</tr>
 
</table>
</td></tr>
</table>
</center>
</BODY>
</HTML>

<%Call OpenConn%>