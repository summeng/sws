<!--#include file="../conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file=cookies.asp-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<%call checkmanage("14")%>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<style type="text/css">
<!--
.style1 {
	color: #FFFFFF;
	font-weight: bold;
}
body {
	background-color: #E3EBF9;
}
-->
</style>
<style type="text/css">
/*提示信息*/
.info {position:relative;background:#;color:#666; text-decoration:none;font-size:12px;width:0px;height:0px;text-align:center;border:0px solid #ccc;line-height:18px;}/*设置链接的属性,一定要设置为relative才能使提示层跟着链接走*/
.info:hover {background:#eee;color:#333;}
.info span {display: none }/*设置正常下的span为隐藏状态*/
.info:hover span /*设置hover下的span属性为呈现状态,并设置提示层的位置*/{display:block;position:absolute;top:0px;left:25px;width:200px;
border:1px solid #DC5826; background:#FFF3B4; color:#333;padding:5px;text-align:left;}
</style>
<%Call OpenConn
dim rows,action,ip,i,url,rsip
rows=request.Form("rows")
if rows="" or rows=0 then rows=10
action=request.Form("ok")
if action="" then 
	set rs=server.createobject("adodb.recordset")
	sql="select * from [DNJY_config]"
	rs.open sql,conn,1,3
%>
<div id="dHTMLToolTip" style="position: absolute; visibility: hidden; width:10; height: 10; z-index: 1000; left: 0; top: 0"></div>

<form action=admin_ip.asp method=post name=addip>
<table width="98%" border="1"  style="border-collapse: collapse; border-style: dotted; border-width: 0px"  bordercolor="#333333" cellspacing="0" cellpadding="2">

<tr><td width=25% align=right>    
锁定IP列表：<br>(对下面红色选项有效) &nbsp;</td><td><TEXTAREA alt="请输入限制访问本站的IP地址<br>如：205.158.40.15<br>每个IP地址占一行" NAME="IP" ROWS="<%=rows%>" COLS="25" style="overflow:auto;">
<%
if rs("IP")<>"" then response.write Replace(rs("IP"),"@",vbCRLF)
%>
</TEXTAREA><br>※提示：若输入IP地址的一部分，如：220.50，那么任何包含220.50的IP地址都将禁止访问本站。</td></TR>
<tr><td colspan=2 width="568" align="center"><INPUT name="ok" TYPE="hidden" value="ip"><INPUT name=action TYPE="submit" value="确认修改"></form></table>

<table width="98%" border="1"  style="border-collapse: collapse; border-style: dotted; border-width: 0px"  bordercolor="#333333" bgcolor="#cccccc" cellspacing="0" cellpadding="2"><form action=admin_ip.asp method=post name=setup>
              <tr> 
            <td height="22" align="right"><font color=red>是否锁定IP限制访问网站：</font></td>
            <td height="22" colspan="3"><input type="radio" name="lockip" value="1" <%if rs("lockip")="1" then%>checked<%end if%>>是 <input type="radio" name="lockip" value="0" <%if rs("lockip")<>"1" then%>checked<%end if%>>否&nbsp;&nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>锁定后IP列表中的IP地址将无法访问本站</span></a></td>
          </tr>
              <tr> 
            <td height="22" align="right"><font color=red>是否锁定IP限制发布信息：</font></td>
            <td height="22" colspan="3"><input type="radio" name="addip" value="1" <%if rs("addip")="1" then%>checked<%end if%>>是 <input type="radio" name="addip" value="0" <%if rs("addip")<>"1" then%>checked<%end if%>>否&nbsp;&nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>锁定后IP列表中的IP地址将无法发布信息</span></a></td>
          </tr>
              <tr> 
            <td height="22" align="right"><font color=red>会员登录错误次数限制：</font></td>
            <td height="22" colspan="3"><INPUT NAME="hydlip" TYPE="text" ID="hydlip" SIZE="3" VALUE="<% = rs("hydlip") %>">&nbsp;&nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>连续登陆X次密码出错，锁定IP，0为不限制</span></a></td>
          </tr>
              <tr> 
            <td height="22" align="right"><font color=red>会员取回密码错误次数限制：</font></td>
            <td height="22" colspan="3"><INPUT NAME="gpwip" TYPE="text" ID="gpwip" SIZE="3" VALUE="<% = rs("gpwip") %>">&nbsp;&nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>连续取回X次密码出错，锁定IP，0为不限制</span></a></td>
          </tr> 
              <tr> 
            <td height="22" align="right"><font color=red>后台登录错误次数限制：</font></td>
            <td height="22" colspan="3"><INPUT NAME="htdlip" TYPE="text" ID="htdlip" SIZE="3" VALUE="<% = rs("htdlip") %>">&nbsp;&nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>登录后台连续出错，锁定IP，0为不限制。</span></a></td>
          </tr>           
              <tr> 
            <td height="22" align="right"><font color=red>连续留言次数限制：</font></td>
            <td height="22" colspan="3"><INPUT NAME="bookip" TYPE="text" ID="bookip" SIZE="3" VALUE="<% = rs("bookip") %>">&nbsp;&nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>一个会话期20分钟内连续留言次数，超过锁定IP，0为不限制</span></a></td>
          </tr>
    
              <tr> 
            <td height="22" align="right"><font color=red>连续发布信息次数限制：</font></td>
            <td height="22" colspan="3"><INPUT NAME="infoip" TYPE="text" ID="infoip" SIZE="3" VALUE="<% = rs("infoip") %>">&nbsp;&nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>一个会话期20分钟内连续发布信息次数，超过锁定IP，0为不限制</span></a></td>
          </tr>
              <tr> 
            <td height="22" align="right"><font color=red>连续出现SQL注入行为次数限制：</font></td>
            <td height="22" colspan="3"><INPUT NAME="sqlip" TYPE="text" ID="sqlip" SIZE="3" VALUE="<% = rs("sqlip") %>">&nbsp;&nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>一个会话期20分钟内连续出现SQL注入行为次数，超过锁定IP，0为不限制</span></a></td>
          </tr>
              <tr> 
            <td height="22" align="right">&nbsp;</td>
            <td height="22" colspan="3">&nbsp;</td>
          </tr>

              <tr> 
            <td height="22" align="right">用户注册间隔时间限制：</td>
            <td height="22" colspan="3"><INPUT NAME="times" TYPE="text" ID="times" SIZE="3" VALUE="<% = rs("times") %>"> 小时内只能注册<INPUT NAME="zcNum" TYPE="text" ID="zcNum" SIZE="3" VALUE="<% = rs("zcNum") %>">&nbsp;&nbsp;次<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>同一ip限定小时内注册次数</span></a></td>
          </tr>
              <tr> 
            <td height="22" align="right">发布信息间隔时间限制：</td>
            <td height="22" colspan="3"><INPUT NAME="infosj" TYPE="text" ID="infosj" SIZE="3" VALUE="<% = rs("infosj") %>">&nbsp;&nbsp;秒<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>连续发布信息间隔时间，单位为秒</span></a></td>
          </tr>
              <tr> 
            <td height="22" colspan="4" align="center"><font color=DD5A26>有的服务器安装有敏感字符限制软件，可能误报无法保存，遇此情况可清空或部分删除下面限制的字符，由服务器负责过滤。</font></td>
          </tr>
              <tr> 
            <td height="22" align="right">信息标题字符限制：</td>
            <td height="22" colspan="3"><INPUT NAME="killword" TYPE="text" ID="killword" SIZE="80" VALUE="<% = rs("killword") %>">&nbsp;&nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>主要作用是限制发布非法关键字，凡信息标题含有限制字符的都不能发布。，字符间用‘|'分隔，注意不能以‘|'结尾！，有的服务器安装有敏感字符限制可能误报无法保存，可删除</span></a></td>
          </tr>
              <tr> 
            <td height="22" align="right">信息详细内容字符限制：</td>
            <td height="22" colspan="3"><INPUT NAME="killmemo" TYPE="text" ID="killmemo" SIZE="80" VALUE="<% = rs("killmemo") %>">&nbsp;&nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>主要作用是限制发布非法关键字，凡信息详细内容含有限制字符的都不能发布。字符间用‘|'分隔，注意不能以‘|'结尾！</span></a></td>
          </tr>  		  
              <tr> 
            <td height="22" align="right">信息联系人和地址字符限制：</td>
            <td height="22" colspan="3"><INPUT NAME="killname" TYPE="text" ID="killname" SIZE="80" VALUE="<% = rs("killname") %>">&nbsp;&nbsp;<a class="info" href="#"><img src="../images/memo.gif"  BORDER="0" /><span>主要作用是限制发布非法关键字，凡信息联系人含有限制字符的都不能发布。字符间用‘|'分隔，注意不能以‘|'结尾！</span></a></td>
          </tr>  
          
                                   
<tr><td colspan=4 width="568" align="center"><INPUT name="ok" TYPE="hidden" value="ok"><INPUT name=action TYPE="submit" value="保存设置"></form></table>


<DIV style="position:absolute; top:120; left:500"> 
<table border=0 width=100><tr><td width=50% align=center>
<form action=admin_ip.asp method=post name=setup2>
<INPUT name="rows" TYPE="hidden" value="<%=rows+10%>">
<input type=image src=images/jia.gif border=0 alt="增加编辑区域">
</form>
</td>
<td width=100 align=center>
<form action=admin_ip.asp method=post name=setup3>
<INPUT name="rows" TYPE="hidden" value="<%=rows-10%>">
<input type=image src=images/jian.gif border=0 alt="减少编辑区域">
</form>
</td></tr></table>
</td></tr>
</table>


<%Call CloseRs(rs)
conn.close
set conn=nothing
end If

if action="ok" then
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from DNJY_config"
rs.open sql,conn,1,3
rs("lockip")=request.form("lockip")
rs("addip")=request.form("addip")
rs("hydlip")=request.form("hydlip")
rs("gpwip")=request.form("gpwip")
rs("htdlip")=request.form("htdlip")
rs("bookip")=request.form("bookip")
rs("infoip")=request.form("infoip")
rs("sqlip")=request.form("sqlip")
rs("infosj")=request.form("infosj")
rs("killword")=request.form("killword")
rs("killmemo")=request.form("killmemo")
rs("killname")=request.form("killname")
rs("times")=request.form("times")
rs("zcNum")=request.form("zcNum")

rs.update
url="admin_ip.asp"
Call CloseRs(rs)
response.Write "<script language=javascript>alert('操作成功！将自动生成配置文件config.asp');location.href='admin_config.asp?page="&url&"';</script>"
end If

if action="ip" Then
Set rsip=Server.CreateObject("ADODB.Recordset")
sql="select * from DNJY_config"
rsip.open sql,conn,1,3
ip=replace(request.form("ip")," ","")
ip=replace(ip,vbCRLF,"@")
for i=1 to 5
ip=replace(ip,"@@","@")
next
if right(ip,1)="@" then ip=left(ip,len(ip)-1)
if left(ip,1)="@" then ip=right(ip,len(ip)-1)
rsip("IP")=ip
rsip.update
rsip.close
set rsip=nothing
Call CloseDB(conn)
response.Write "<script language=javascript>alert('操作成功！');location.href='admin_ip.asp';</script>"
end If
%>
