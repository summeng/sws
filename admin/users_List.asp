<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
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
<%Dim page
call checkmanage("04")%>
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" type="text/css" href="../css/css.css">
<BODY>
<!---全选验证删除开始--->
<SCRIPT language=JavaScript>
function showoperatealert(id){
		if (id==1)
        {
		{
		  	thisForm.target='_self';
			thisForm.action="user_yz.asp?qxyz=uyes";
			thisForm.submit();
		}
		}
		if (id==2)
        {
		{
		  	thisForm.target='_self';
			thisForm.action="user_yz.asp?qxyz=uno";
			thisForm.submit();
		}
		}
		if (id==3)
        {
		{
		  	thisForm.target='_self';
			thisForm.action="user_yz.asp?qxyz=myes";
			thisForm.submit();
		}
		}
		if (id==4)
        {
		{
		  	thisForm.target='_self';
			thisForm.action="user_yz.asp?qxyz=mno";
			thisForm.submit();
		}
		}		
		if (id==5)
        {
		{
		  	thisForm.target='_self';
			thisForm.action="user_del.asp";
			thisForm.submit();
		}
		}
		if (id==6)
        {
		{
		  	thisForm.target='_blank';
			thisForm.action="sendmail.asp?userqf=ok";
			thisForm.submit();
		}
		}		
}
//-->
    </SCRIPT>
<!---全选验证删除结束--->
<script>
function loadThreadFollow(t_id,b_id){//弹出隐藏
	var targetImg =eval("document.all.followImg" + t_id);
	var targetDiv =eval("document.all.follow" + t_id);
	
	if ("object"==typeof(targetImg)){
		if (targetDiv.style.display!=''){
			targetDiv.style.display="";
			
		}else{
			targetDiv.style.display="none";
		}
	}
}
function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.name != 'chkall' )
       e.checked = form.chkall.checked;
    }
  }
function DoEmpty(params)
{
if (confirm("真的要删除这条记录吗？删除后将无法恢复！"))
window.location = params ;
}
function test()
{
  if(!confirm('是否确定进行批量操作？操作后不能恢复！')) return false;
}
//-->
</script>
<style type="text/css">
<!--
body {
	background-color: #E3EBF9;
}
.style1 {color: #FFFFFF}
-->
</style><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" height="1">
<TR>
    <TD height="20" bgcolor="#799AE1" align="center" colspan="13"><FONT color="#FFFFFF" style="font-size:14px">注 册 用 户 管 理</FONT></TD>
 </TR>
  <tr>    
    <td width="30%">
    <FORM name=sou1 action="?dick=1" method=POST>
    <p align="center">按帐号：
      <input type="text" name="T1" size="20">
    <input type="submit" value="查找" name="B1">
    </form>
    </td>
    <td width="30%">
    <FORM name=sou2 action="?dick=2" method=POST>
    <p align="center">按姓名：
      <input type="text" name="T2" size="20">
    <input type="submit" value="查找" name="B1">
    </form>
    </td>
<td width="40%">
<%set rs=server.createobject("adodb.recordset")
sql = "select * from [DNJY_user] where (mpname<>'' and mpkg=2) or (sdname<>'' and sdkg=2 ) or useryz=0 or (vip=1 and vipdq<"&DN_Now&")"
rs.open sql,conn,1,1
if rs.eof then
response.write "<font color=#0000FF>目前没有待审核事项！</font>"
Else
response.write "<b><img src='../img/notice.gif' width=15 height=15 border=0><font color=#FF0000>铃声响有待审核事项！</font></b>"
response.write "(<a href='?dick=14' title='指等待审核的会员（当启用会员审核时）'>待审会员</a>--<a href='?dick=15' title='指申请或修改店铺状态需审核的会员'>待审店企</a>--<a href='?dick=4' title='指VIP资格到期的会员'>到期VIP</a>)<bgsound src='../img/msg.wav' loop='1'>"
end If
Call CloseRs(rs)%>
 </td>
  </tr>
</table>
<div align="center"><a href="users_List.asp">全部会员</a>--<a href="?dick=3" title='指具备VIP资格的会员'>VIP会员</a>--<a href="?dick=4" title='指VIP资格期限到的会员'>VIP到期会员</a>--<a href="?dick=5" title='指10后到期的VIP会员'>VIP有效期不足10天会员</a>--<a href="?dick=6" title='指一般注册的会员'>普通会员</a>--<a href="?dick=7" title='指开有网络店铺的会员'>店铺会员</a>--<a href="?dick=8" title='指开有企业名片的会员'>名片会员</a>--<a href="?dick=9" title='指发布信息较多的会员'>活跃会员</a>--<a href="?dick=10" title='指<%=rmb_mc%>、人民币余额较多的会员'>富翁会员</a>--<a href="?dick=11" title='指信息、网店、名片均为0的会员'>懒惰会员</a>--<a href="?dick=12" title='指最近登录时间距今百天以上的会员'>失踪百天会员</a>--<a href="?dick=13" title='指最近登录时间距今百天以上且信息、网店、名片均为0的会员'>失踪百天懒惰会员</a></div>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#ADAED6" width="100%" height="30">
<FORM name=thisForm action="user_del.asp" method=POST>
<%Call OpenConn
dim k,dick,rs1
dim ThisPage,Pagesize,Allrecord,Allpage
if request("page")="" then
  ThisPage=1		
else
  ThisPage=request("page")
end if
dick=request("dick")
set rs = Server.CreateObject("ADODB.RecordSet")
set rs1 = Server.CreateObject("ADODB.RecordSet")
Select Case dick
Case "1"'帐号搜索
sql = "select * from [DNJY_user] where username like '%"&trim(request("T1"))&"%' order by id "&DN_OrderType&""
Case "2"'姓名搜索
sql = "select * from [DNJY_user] where name like '%"&trim(request("T2"))&"%' order by id "&DN_OrderType&""
Case "3"'VIP会员
sql = "select * from [DNJY_user] where vip=1 order by id "&DN_OrderType&""
Case "4"'VIP到期会员
sql = "select * from [DNJY_user] where vip=1 and datediff("&DN_DatePart_D&",vipdq,"&DN_Now&")>=0 order by id "&DN_OrderType&""
Case "5"'VIP有效期不足10天会员
sql = "select * from [DNJY_user] where vip=1 and DateDiff("&DN_DatePart_D&",vipdq,"&DN_Now&")*-1<10  order by id "&DN_OrderType&""
Case "6"'普通会员
sql = "select * from [DNJY_user] where vip=0 order by id "&DN_OrderType&""
Case "7"'店铺会员
sql = "select * from [DNJY_user] where sdname<>'' order by id "&DN_OrderType&""
Case "8"'名片会员
sql = "select * from [DNJY_user] where mpname<>'' order by id "&DN_OrderType&""
Case "9"'活跃会员
sql = "select * from [DNJY_user] where xxts>0 order by xxts "&DN_OrderType&""
Case "10"'富翁会员
sql = "select * from [DNJY_user] where hb>0 or Amount>0 order by hb "&DN_OrderType&",Amount "&DN_OrderType&""
Case "11"'懒惰会员
sql = "select * from [DNJY_user] where xxts=0 and sdname='' and mpname='' order by id "&DN_OrderType&""
Case "12"'失踪百天会员
'dldata is not Null 日期dldata不空
'dldata is null 日期dldata为空
sql = "select * from [DNJY_user] where DateDiff("&DN_DatePart_D&",dldata,"&DN_Now&")>=100 or dldata is null order by id "&DN_OrderType&""
Case "13"'失踪百天懒惰会员
sql = "select * from [DNJY_user] where xxts=0 and sdname='' and mpname=''  and DateDiff("&DN_DatePart_D&",dldata,"&DN_Now&")>=100 order by id "&DN_OrderType&""
Case "14"'待审会员
sql = "select * from [DNJY_user] where useryz=0 order by id "&DN_OrderType&""
Case "15"'待审店企
sql = "select * from [DNJY_user] where (mpname<>'' and mpkg=2) or (sdname<>'' and sdkg=2) order by id "&DN_OrderType&""
Case Else'全部会员
sql="select * from [DNJY_user] order by id "&DN_OrderType&""
End Select
rs.open sql,conn,1,3
if rs.eof then
response.write "<li>还没有用户数据！"
response.end
end if
rs.Pagesize=20
Pagesize=rs.Pagesize
Allrecord=rs.Recordcount
Allpage=rs.Pagecount
if ThisPage<1 then                           
ThisPage=1
end if
On Error Resume Next
rs.move (ThisPage-1)*Pagesize
k=0
%>
  <tr bgcolor="#BDBEDE">
    <td width="3%" height="28" align="center" bordercolor="#666666">编号</td>
	<td width="2%" height="28" align="center" bordercolor="#666666">选择</td>
    <td width="5%" height="28" align="center" bordercolor="#666666">登陆帐号</td>
    <td width="6%" height="28" align="center" bordercolor="#666666">姓名</div></td>
    <td width="4%" height="28" align="center" bordercolor="#666666">审核状态</td>
    <td width="4%" height="28" align="center" bordercolor="#666666">邮件确认</td>
	<td width="4%" height="28" align="center" bordercolor="#666666">VIP</td>
    <td width="4%" height="28" align="center" bordercolor="#666666">信息条数</td>
	<td width="4%" height="28" align="center" bordercolor="#666666">积分</td>
    <td width="4%" height="28" align="center" bordercolor="#666666"><%=rmb_mc%><br>(个)</td>
    <td width="4%" height="28" align="center" bordercolor="#666666">帐户金额<br>(RMB)</td>
    <td width="4%" height="28" align="center" bordercolor="#666666">网店/状态</td>
	<td width="4%" height="28" align="center" bordercolor="#666666">名片/状态</td>
    <td width="3%" height="28" align="center" bordercolor="#666666">颜色道具</td>
    <td width="3%" height="28" align="center" bordercolor="#666666">置顶道具</td>
    <td width="3%" height="28" align="center" bordercolor="#666666">图片道具</td>
    <td width="3%" height="28" align="center" bordercolor="#666666">验证道具</td>
    <td width="3%" height="28" align="center" bordercolor="#666666">守信次数</td>
    <td width="3%" height="28" align="center" bordercolor="#666666">失信次数</td>
    <td width="3%" height="28" align="center" bordercolor="#666666">参评网友(人次)</td>
	<td width="3%" height="28" align="center" bordercolor="#666666">平均得分</td>
    <td width="5%" height="28" align="center" bordercolor="#666666">注册时间</td>
    <td width="5%" height="28" align="center" bordercolor="#666666">VIP到期时间</td>
    <td width="2%" height="28" align="center" bordercolor="#666666">操作</td>	
  </tr>
  <%Dim id
  do while not rs.eof
  id=rs("id")
  %>
  <tr bgcolor="#FFFFFF">
    <td width="3%" height="25" bgcolor="#FaFaFa">
    <p align="center"><%=k+1%></td>
    <td width="2%" height="25" align="center" bgcolor="#FFFFFF">
    <input type="checkbox" name="selectedid" value="<%=trim(rs("username"))%>"></td>
    <td width="5%" height="25" bgcolor="#FaFaFa"><p align="center"><a href=../preview.asp?username=<%= rs("username") %> target="_blank"  title='查<%=rs("username")%>的综合信息'><%=rs("username")%></a><%If (rs("mpname")<>"" and rs("mpkg")=2) or (rs("sdname")<>"" and rs("sdkg")=2) Or rs("useryz")=0 then%><img src="img/warning.gif" alt="此会员有待审核事项"><%End if%>|<a href="infolist.asp?dick=18&key1=<%= rs("username")%>&city_one=<%=city_oneid%>&city_two=<%=city_twoid%>&city_three=<%=city_threeid%>&type_one=<%=type_oneid%>&type_two=<%=type_twoid%>&type_three=<%=type_threeid%>" title='查<%=rs("username")%>发布的供求信息'>信息</a></td>
    <td width="6%" height="25"><div align="center"><a href=../preview.asp?username=<%= rs("username") %> target="_blank"><%=rs("name")%></a></div></td>
    <td width="4%" height="25"><p align="center">
<%if rs("useryz")=1 then
response.write "<a href='user_yz.asp?selectedid="&rs("username")&"&qxyz=uno&page="&ThisPage&"'><font color=""#ff0000""><strong><span title=已审核>√</span></strong></font></a>"
else
response.write "<font color=""#0066FF""><strong><span title=未审核><a href='user_yz.asp?selectedid="&rs("username")&"&qxyz=uyes&page="&ThisPage&"'>ⅹ</span></strong></font></a>"
end if%></td>
    <td width="4%" height="25"><p align="center">
<%if rs("mailyz")=1 then
response.write "<a href='user_yz.asp?selectedid="&rs("username")&"&qxyz=mno&page="&ThisPage&"'><font color=""#ff0000""><strong><span title=已确认>√</span></strong></font></a>"
else
response.write "<a href='user_yz.asp?selectedid="&rs("username")&"&qxyz=myes&page="&ThisPage&"'><font color=""#0066FF""><strong><span title=未确认>ⅹ</span></strong></font></a>"
end if%></td>
    <td width="4%" height="25">
      <p align="center">
  <%dim vip
vip=rs("vip")
if vip=1 then
response.write "<font color=""#ff0000""><strong>√</strong></font>"
else
response.write "<font color=""#0066FF""><strong>ⅹ</strong></font>"

end if%>
      </td>
	<td width="4%" height="25" bgcolor="#FaFaFa"><p align="center"><%=rs("xxts")%></td>
    <td width="4%" height="25" bgcolor="#FaFaFa"><p align="center"><%=rs("jf")%></td>
    <td width="4%" height="25"><p align="center"><%=rs("hb")%></td>
    <td width="4%" height="25"><p align="center"><%=rs("Amount")%></td>

    <td width="4%" height="25"><p align="center"><%
	dim sd
sd=rs("sdname")
if sd<>"" then
response.write "<font color=""#ff0000""><strong><span title=已开店>√</span></strong></font>"
else
response.write "<font color=""#0066FF""><strong><span title=未开店>ⅹ</span></strong></font>"

end if%>/<%
	dim sdkg
sdkg=rs("sdkg")
if sdkg=1 then
response.write "<font color=""#ff0000""><strong><span title=开放状态>√</span></strong></font>"
elseif sdkg=2 Then
response.write "<font color=""#0000ff""><strong><span title=申请开放>！</span></strong></font>"
else
response.write "<font color=""#0066FF""><strong><span title=关闭状态>ⅹ</span></strong></font>"

end if%></td>
    <td width="4%" height="25"><p align="center"><%
	dim mp
mp=rs("mpname")
if mp<>"" then
response.write "<font color=""#ff0000""><strong><span title=已打名片>√</span></strong></font>"
else
response.write "<font color=""#0066FF""><strong><span title=未打名片>ⅹ</span></strong></font>"

end if%>/<%
	dim mpkg
mpkg=rs("mpkg")
if mpkg=1 then
response.write "<font color=""#ff0000""><strong><span title=开放状态>√</span></strong></font>"
elseif mpkg=2 Then
response.write "<font color=""#0000ff""><strong><span title=申请开放>！</span></strong></font>"
else
response.write "<font color=""#0066FF""><strong><span title=关闭状态>ⅹ</span></strong></font>"

end if%></td>
    <td width="3%" height="25" align="center" bgcolor="#FaFaFa">
    <%=rs("a")%></td>
    <td width="3%" height="25" align="center" bgcolor="#FFFFFF">
    <%=rs("b")%></td>
    <td width="3%" height="25" align="center" bgcolor="#FaFaFa">
    <%=rs("c")%></td>
    <td width="3%" height="25" align="center" bgcolor="#FFFFFF">
    <%=rs("d")%></td>
    <td width="3%" height="25" align="center" bgcolor="#FFFFFF">
    <%=rs("goodfaith")%></td>
    <td width="3%" height="25" align="center" bgcolor="#FFFFFF">
    <%=rs("bpromises")%></td>
    <td width="3%" height="25" align="center" bgcolor="#FFFFFF">
<%Dim df,cs
df=rs("wevaluation")
cs=rs("participants")
If rs("wevaluation")=0 Then df=5
If rs("participants")=0 Then cs=1%>
    <%=cs%></td>
    <td width="3%" height="25" align="center" bgcolor="#FFFFFF">
    <%=Formatnumber(df/cs,1,0,0,0)%></td>
    <td width="5%" height="25" bgcolor="#FaFaFa"><p align="center"><%=rs("zcdata")%></td>
    <td width="5%" height="25" bgcolor="#FaFaFa"><p align="center"><%=rs("vipdq")%></td>
    <td width="2%" height="25" align="center" bgcolor="#FaFaFa">
    <font color="#FF0000">
    <span id="followImg<%=k%>" style="CURSOR: hand" onclick="loadThreadFollow(<%=k%>,5)">
    面版</span></font></td>

  </tr>
  <tr style="display:none" id="follow<%=k%>">
    <td width="95%" height="25" colspan="25">    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="100%">
      <tr bgcolor="#CCCCFF"><td><div align="right"><%If rs("tjname")<>"" then%>其注册推荐人：<a href=../preview.asp?username=<%= rs("tjname") %> target="_blank"><%=rs("tjname")%></a>&nbsp;&nbsp;&nbsp;<%End if%><a href="#" onClick="window.open('user_tjname.asp?username=<%=rs("username")%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=200,height=590,left=300,top=100')">其发展的会员</a>&nbsp;&nbsp;&nbsp;最后登录IP：<a href="http://www.ip138.com/ips.asp?ip=<%=rs("userip")%>&action=2" target="_blank"><%=rs("userip")%></a>&nbsp;&nbsp;&nbsp;｜<font color="#666666"><%if sd<>"" then%><a href="#" onclick="window.open('user_editdp.asp?id=<%=id%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=510,height=520,left=300,top=100')"><font color="#FF0000">网店审核</font></a><%end if%> &nbsp;<%if mp<>"" then%><a href="#" onclick="window.open('user_editdp.asp?id=<%=id%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=550,height=520,left=300,top=100')"><font color="#FF0000">名片审核</font></a><%end if%> <a href="#" onClick="window.open('user_editdata.asp?username=<%=rs("username")%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=550,height=650,left=300,top=100')">修改资料</a> <a href="#" onClick="window.open('user_edithy.asp?id=<%=id%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=420,height=180,left=300,top=100')">修改级别</a> <a href="#" onClick="window.open('user_editcurrency.asp?id=<%=id%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=180,height=150,left=300,top=100')">修改积分<%=rmb_mc%></a> <a href="#" onClick="window.open('user_editprops.asp?id=<%=id%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=200,height=150,left=300,top=100')">修改道具</a> <a href="#" onClick="window.open('user_evaluation.asp?id=<%=id%>&username=<%=rs("username")%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=200,height=150,left=300,top=100')">评价</a> <a href="#" onClick="window.open('admin_Amount_add.asp?dick=uyz&username=<%=rs("username")%>&aliname=<%=rs("name")%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=480,height=400,left=300,top=100')"><font color="#FF0000">财务操作</font></a> <a href="#" onClick="window.open('user_mail.asp?username=<%=rs("username")%>&email=<%=rs("email")%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width=700,height=700,left=300,top=30')">给其邮件</a> <a href='message2.asp?username=<%=trim(rs("username"))%>'>给其短信</a></font><font color="#0000FF"> <font color="#0000FF"><a href="javascript:DoEmpty('user_del.asp?selectedid=<%=trim(rs("username"))%>')">删此会员</a></font></font><font color="#666666">
<%dim m
set rs1=conn.execute("select count(id) from [DNJY_info] where username='"&rs("username")&"'")
m=rs1(0)
rs1.close%>
<a target="_blank" href="user_center.asp?id=<%=id%>">代管会员中心</a>[发有<%=m%>条信息]&nbsp;</font></div></td>
      </tr>
    </table></td>
  </tr>
    <%
    k=k+1
    rs.movenext
    if k>=Pagesize then exit do
	loop
	rs.close
    set rs=nothing
    set rs1=nothing
    Call CloseDB(conn)
	%>
  <tr bgcolor="#eeeeee">
    <td width="100%" height="35" colspan="25">
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
<tr> 
<td width="595" height="25" align="center">
  共有&nbsp;<font color="#CC5200"><%=Allrecord%></font>&nbsp;条记录 共 <font color="#CC5200"><%=Allpage%></font> 页 现在是第 <font color="#CC5200"><%=ThisPage%></font> 页
<%
if ThisPage<2 then     
response.write "<font color=""#808080"">首页</font>&nbsp;"
response.write "<font color=""#808080"">上一页</font>&nbsp;"     
else     
response.write "<a href=?page=1&t1="&trim(request("t1"))&"&dick="&request("dick")&">首页</a>&nbsp;"
response.write "<a href=?page="&ThisPage-1&"&t1="&trim(request("t1"))&"&dick="&request("dick")&">上一页</a>&nbsp;"     
end if
if Allpage-ThisPage<1 then     
response.write "<font color=""#808080"">下一页</font>&nbsp;"
response.write "<font color=""#808080"">尾页</font>&nbsp;"  
else     
response.write "<a href=?page="&(ThisPage+1)&"&t1="&trim(request("t1"))&"&dick="&request("dick")&">下一页</a>&nbsp;"   
response.write "<a href=?page="&Allpage&"&t1="&trim(request("t1"))&"&dick="&request("dick")&">尾页</a>&nbsp;"     
end if
Session("page") =ThisPage%></td>

</tr>
<tr> 
<td width="700" height="25" align="center">
  <input onClick=CheckAll(this.form) type=checkbox value=on name=chkall>
  全选记录  
  <input onClick=javascript:showoperatealert(1) type="submit" value="审核通过" name="B1">
  <input onClick=javascript:showoperatealert(2) type="submit" value="取消审核" name="B2">
  <input onClick=javascript:showoperatealert(3) type="submit" value="通过邮件确认" name="B3">
  <input onClick=javascript:showoperatealert(4) type="submit" value="取消邮件确认" name="B4">  
  <input onClick=javascript:showoperatealert(6) type="submit" value="发送邮件" name="B6">
  <input name='submit1' type='submit' value=' 批量删除 ' onClick="return test();"  style="color: #FF0000">
</td>
</tr>		  
</table>    </td>
  </tr>
  </form>
</table>
  </center>
</div>
