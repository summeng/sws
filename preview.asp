<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
 <!--#include file=top.asp-->
 
 <%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<link rel="stylesheet" type="text/css" href="1.CSS">
</head>
<%
dim rsper,k,sdname,sdkg,sdfg,mpname,mpkg,mpfg,vip,guest,permissions
username=trim(request("username"))
Readinfo=trim(request("Readinfo"))
If trim(request("Readinfo"))="" Then Readinfo=0
Call OpenConn
set rsper = Server.CreateObject("ADODB.RecordSet")
sql="select * from [DNJY_user] where username='"&username&"'" 
rsper.open sql,conn,1,1
if rsper.eof then
response.write "<p><center><li>哇！估计你找错地方喽！用户没有注册怎么办！"
response.end
end if
vip=rsper("vip")
sdkg=rsper("sdkg")
sdfg=rsper("sdfg")
sdname=rsper("sdname")
mpkg=rsper("mpkg")
mpfg=rsper("mpfg")
mpname=rsper("mpname")

If request.cookies("dick")("username")<>username Then'自己发布的全权阅读
guest=0'登录会员
If request.cookies("dick")("username")<>"" Then guest=conn.Execute("Select vip From DNJY_user Where username='"&request.cookies("dick")("username")&"'")(0)
If Readinfo=1 And request.cookies("dick")("username")="" Then
permissions="普通会员权限阅读"
End If
If Readinfo=2 And guest<1 Then
permissions="VIP会员权限阅读"
End If
End If%>
<body topmargin="0" background="images/bg1.gif">
<div align="center">

</div>
<table cellSpacing="0" cellPadding="0" width="980" align="center" bgColor="#ffffff" border="0" height="410">
  <tr>
    <td width="980" height="331">
    <table cellSpacing="0" cellPadding="0" width="980" border="0" height="406">
      <tr>
        <td vAlign="top" width="210" height="406">
        <table cellSpacing="0" cellPadding="0" width="97%" border="0">
                   <tr>
            <td>
           <!--#include file=left.asp--></td>
          </tr>
        </table>
        </td>
		<td width="5">&nbsp;</td>
        <td vAlign="top" width="760" height="406">
        <div align="center">        
        <table border="0" cellpadding="0" cellspacing="0"   width="100%" height="222">
		  <table width="760" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="ty1">
          <tr>
            <td width="100%" height="50" background="images/user_1.gif" colspan="6">　
              <div align="center"><b>会员详细信息仅供交易参考</b></div></td>
          </tr>
<tr>
<td>
<table width="750" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" align="center">
<TR>
	<TD valign="top" width="450"><TABLE>
	<tr>
    <td align="right">会员ID号：</td><td><%=rsper("username")%></td>
    </tr>
	<tr>
    <td align="right">会员姓名：</td><td><%if Readinfo>0 then%><%=permissions%><%else%><%=rsper("name")%>&nbsp;<%if rsper("sex")=1 then%>先生<%else%>女士<%end if%><%end if%></td>
    </tr>
	<tr>
    <td align="right">VIP 会员：</td><td><%if rsper("vip")=1 then%>是<%else%>否<%end if%>&nbsp;&nbsp;&nbsp;&nbsp;<%if sdkg=1 And sdname<>"" then%> <a  href="user/com.asp?username=<%=rsper("username")%>&com=<%=rsper("username")%>&<%=C_ID%>" target="_blank"><img src="img/dian.gif" title="此会员已开有店铺" ></a><%end if%> &nbsp;&nbsp;&nbsp;<%if mpkg=1 And mpname<>"" then%><a  href="company.asp?username=<%=rsper("username")%>&<%=C_ID%>" target="_blank"><img src="img/qy.gif" title="此会员已开有企业名片" ></a><%end if%></td>
    </tr>
	<tr>
    <td align="right">注册时间：</td><td><%=rsper("zcdata")%></td>
    </tr>
	<tr>
    <td align="right">站长评价：</td><td><img src="img/hp.jpg"> 守信记录 <%=rsper("goodfaith")%> 次&nbsp;&nbsp;| <img src="img/cp.jpg"> 失信记录 <%=rsper("bpromises")%> 次</td>
    </tr>
	<tr>
    <td align="right" valign="top">网友打分：</td><td>

<%Dim df,cs,pj
df=rsper("wevaluation")
cs=rsper("participants")
If rsper("wevaluation")=0 Then df=5
If rsper("participants")=0 Then cs=1
pj=Formatnumber(df/(cs*5)*100,0,0,0,0)%>
	<style type="text/css"> 
body{ 
text-align:center; 
} 
.graph{ 
width:355px; 
border:1px solid #F8B3D0; 
height:15px; 
} 
#bar{ 
display:block; 
background:#FFE7F4; 
float:left; 
height:100%; 
text-align:center; 
} 
#barNum{ 
position:absolute; 
} 
</style> 
<script type="text/javascript"> 
function $(obj){ 
return document.getElementById(obj); 
} 
function go(){ 
$("bar").style.width = parseInt($("bar").style.width) + 1 + "%"; 
$("bar").innerHTML = $("bar").style.width; 
if($("bar").style.width =="<%=pj%>%"){ 
window.clearInterval(bar); 
} 

} 
var bar = window.setInterval("go()",30); 
window.onload = function(){ 
bar; 
} 
</script> 
<div class="graph"> 
<strong id="bar" style="width:1%;"></strong> 
</div>
<img src="img/pj.jpg" width="370"><br>网友参评 <%=cs%> 人次&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;平均得分 <%=Formatnumber(df/cs,1,0,0,0)%> 分&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<INPUT TYPE=button VALUE="我要评价" ONCLICK="window.open('user/evaluation.asp?username=<%=rsper("username")%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbarsper=yes,scrollbars=yes,resizable=no,copyhistory=no,width=400,height=300,left=300,top=100')" style="color: #0060C5; background-color: #E1E2DC">	
	</td>
    </tr>

	</TABLE>		  
	</TD>
	<TD valign="top"><TABLE>
	<tr>
    <td align="right">腾讯QQ号：</td><td><%if Readinfo>0 then%><%=permissions%><%else%><%=rsper("qq")%><%end if%></td>
    </tr>
	<tr>
    <td align="right">联系电话：</td><td><%if Readinfo>0 then%><%=permissions%><%else%><%=rsper("dianhua")%><%end if%></td>
    </tr>
	<tr>
    <td align="right">联系手机：</td><td><%if Readinfo>0 then%><%=permissions%><%else%><%=rsper("CompPhone")%><%end if%></td>
    </tr>
	<tr>
    <td align="right">站内短信：</td><td><%if Readinfo>0 then%><%=permissions%><%else%><a href='messh.asp?username2=<%=trim(rsper("username"))%>&<%=C_ID%>'>发站内短信给他</a><%end if%></td>
    </tr>
	<tr>
    <td align="right">电子信箱：</td><td><%if Readinfo>0 then%><%=permissions%><%else%><INPUT TYPE=button VALUE="发邮件交流" ONCLICK="window.open('user/user_mail.asp?username=<%=rsper("username")%>&email=<%=rsper("email")%>&mailzt=供求信息','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbarsper=yes,scrollbars=yes,resizable=no,copyhistory=no,width=650,height=300,left=300,top=100')" style="color: #666666; background-color: #E1E2DC"><%end if%></td>
    </tr>
	<tr>
    <td align="right">公司名称：</td><td><%if Readinfo>0 then%><%=permissions%><%else%><%end if%></td>
    </tr>
	<tr>
    <td align="right">邮政编码：</td><td><%if Readinfo>0 then%><%=permissions%><%else%><%=rsper("youbian")%><%end if%></td>
    </tr>
	<tr>
    <td align="right">通信地址：</td><td><%if Readinfo>0 then%><%=permissions%><%else%><%=rsper("dizhi")%><%end if%></td>
    </tr>
	</TABLE>
	</TD>
</TR>
</TABLE>
</td>
</tr>          
           <tr>
            <td width="100%" height="12" colspan="6"></td>
          </tr>
            <td width="100%" height="12" colspan="6"><div align="center"><font color="#ff0000"><%=rsper("name")%>(<%=rsper("username")%>)</font>发布的信息<%If rsper("vip")=0 then%> （此用户已发布<%=rsper("xxts")%>条信息，但不是VIP会员，最多只显示最新<%=huiyuanxx%>条信息）<%end if%></div></td>          
          <%rsper.close%>          
            <td width="100%" height="24" background="images/user_2.gif" colspan="6"><a name="1"></a></td>
          </tr>
          <tr>
            <td width="100%" height="24" colspan="6">
            <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="99%" height="100%">
<%
dim ThisPage,Pagesize,Allrecord,Allpage,b,bb
k=0
if request("page")="" then
  ThisPage=1		
else
  ThisPage=request("page")
end If
if request.cookies("administrator")<>"" then
if vip=1 then
sql = "select id,biaoti,a,b,c,fbsj,llcs,hfcs,Readinfo from DNJY_info where username='"&username&"' order by b "&DN_OrderType&",id "&DN_OrderType&""
else
sql = "select top "&huiyuanxx&" id,biaoti,a,b,c,fbsj,llcs,hfcs,Readinfo from DNJY_info where username='"&username&"' order by b "&DN_OrderType&",id "&DN_OrderType&""
end If
Else
if vip=1 then
sql = "select id,biaoti,a,b,c,fbsj,llcs,hfcs,Readinfo from DNJY_info where yz=1 and username='"&username&"' order by b "&DN_OrderType&",id "&DN_OrderType&""
else
sql = "select top "&huiyuanxx&" id,biaoti,a,b,c,fbsj,llcs,hfcs,Readinfo from DNJY_info where yz=1 and username='"&username&"' order by b "&DN_OrderType&",id "&DN_OrderType&""
end If
End if
rsper.open sql,conn,1,1
if rsper.eof then
response.write "该用户还没有发布信息！"
else
rsper.Pagesize=10
Pagesize=rsper.Pagesize
Allrecord=rsper.Recordcount
Allpage=rsper.Pagecount
if ThisPage<1 then                           
ThisPage=1
end if
'On Error Resume Next
rsper.move (ThisPage-1)*Pagesize
%>
              <tr>
                <td width="7%" align="center" height="25" style="border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1" bordercolor="#F4F4F2">编号</td>
                <td width="66%" align="center" height="25" style="border-top-style: solid; border-top-width: 1" bordercolor="#F4F4F2">&nbsp; 信&nbsp;&nbsp; 
                息&nbsp;&nbsp; 标&nbsp;&nbsp; 题</td>
                <td width="14%" align="center" height="25" style="border-top-style: solid; border-top-width: 1" bordercolor="#F4F4F2">
                查/回</td>
                <td width="13%" align="center" height="25" style="border-right-style: solid; border-right-width: 1; border-top-style: solid; border-top-width: 1" bordercolor="#F4F4F2">发布时间</td>
              </tr>
<%
do while not rsper.eof
b=trim(rsper("b"))
bb=len(b)
%>

              <tr>
                <td width="7%" align="center" height="25" style="border-left-style: solid; border-left-width: 1" bordercolor="#F4F4F2"><%=k+1%></td>
                <td width="66%" align="center" height="25" style="border-bottom-style: none; border-bottom-width: medium" bordercolor="#F4F4F2"><p align="left"><%if rsper("c")=1 then%><img src="images/num/pic.gif"><%end if%><a target="_blank" href="<%=x_path("",rsper("id"),"",rsper("Readinfo"))%>"><%if rsper("a")="0" then%><%=rsper("biaoti")%><%else%><font color="#<%=rsper("a")%>"><%=rsper("biaoti")%></font><%end if%></a><%if b<>0 then%><img src="images/num/jsq.gif"><%for i=1 to bb%><img src="images/num/<%=Mid(b,i,1)%>.gif"><%next%><%end if%></td>
                <td width="14%" align="center" height="25" style="border-bottom-style: none; border-bottom-width: medium" bordercolor="#F4F4F2"><%=rsper("llcs")%>/<%=rsper("hfcs")%></td>
                <td width="13%" align="center" height="25" style="border-right-style: solid; border-right-width: 1; border-bottom-style: none; border-bottom-width: medium" bordercolor="#F4F4F2"><%=datevalue(rsper("fbsj"))%></td>
              </tr>
<%
k=k+1
rsper.movenext
if k>=Pagesize then exit do
loop
%>
              <tr>
                <td width="100%" align="center" height="13" colspan="4" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1" bordercolor="#F4F4F2">
<%
if Allrecord>rsper.Pagesize then
call pages()
end if
sub pages()
%>
<table border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="569">
<tr> 
<td height="24" width="569" colspan="4">
<p align="center">
　</td>
</tr>
<tr> 
<td height="24" width="151" align="center">
共有&nbsp;<font color="#CC5200"><%=Allrecord%></font>&nbsp;条记录</td>
<td height="24" width="126" align="center">
共 <font color="#CC5200"><%=Allpage%></font> 页</td>
<td height="24" width="118" align="center">
现在是第 
                <font color="#CC5200"><%=ThisPage%></font> 页</td>
<td height="24" width="174" align="center">
<%
if ThisPage<2 then     
response.write "<font color=""#808080"">首页</font>&nbsp;"
response.write "<font color=""#808080"">上一页</font>&nbsp;"     
else     
response.write "<a href=?page=1&username="&username&"#1>首页</a>&nbsp;"
response.write "<a href=?page="&ThisPage-1&"&username="&username&"#1>上一页</a>&nbsp;"     
end if
if Allpage-ThisPage<1 then     
response.write "<font color=""#808080"">下一页</font>&nbsp;"
response.write "<font color=""#808080"">尾页</font>&nbsp;"  
else     
response.write "<a href=?page="&(ThisPage+1)&"&username="&username&"#1>下一页</a>&nbsp;"   
response.write "<a href=?page="&Allpage&"&username="&username&"#1>尾页</a>&nbsp;"     
end if
%></td>
</tr>
</table>
<%end sub%>
                </td>
              </tr>
              <tr>
                <td width="100%" align="center" height="12" colspan="4" style="border-left-style: none; border-left-width: medium; border-right-style: none; border-right-width: medium; border-top-style: solid; border-top-width: 1; border-bottom-style: none; border-bottom-width: medium" bordercolor="#F4F4F2">
                </td>
              </tr>
<%end if%>
            </table>
            </td>
          </tr>
        </table>
          </center>
        </div>
        </td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="772" height="23">
    <tr>
      <td width="100%" height="16" bgcolor="#FFFFFF"></td>
    </tr>
    <tr>
      <td width="100%" height="7" bgcolor="#FFFFFF"> </td>
    </tr>
    </table>
  </center>
</div>
</body><%rsper.close
set rsper=nothing%>
<!--#include file=footer.asp-->
</html>
