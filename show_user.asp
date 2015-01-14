<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="top.asp"-->
<%dim zuo1,zuo2,zuo3,info1,info2,info3,info4,info5,info6,info7,info8,info9,xxtp,biaoti,leixin,biaotixs,fbsj,scjiage,jiage,sdmba,usernameid,namea,dianhua,mobile,userqq,email,dizhi,youbian,memo,hfcs,dqsj,sj
%>
<!--
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室版权所有
'官方客服中心 http://www.ip126.com
'技术支持论坛 http://www.ip126.com/bbs
'QQ:530051328 mail:bdunion@126.com
'=====================================--->

<script language="JavaScript">
//弹出回复标题
function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.name != 'chkall' )
       e.checked = form.chkall.checked;
    }
  }
 function loadThreadFollow(t_id,b_id){
	var targetImg =eval("document.all.followImg" + t_id);
	var targetDiv =eval("document.all.follow" + t_id);

	if ("object"==typeof(targetImg)){
		if (targetDiv.style.display!='block'){
			targetDiv.style.display="block";

		}else{
			targetDiv.style.display="none";
		}
	}
}
</script>
<script LANGUAGE="JavaScript">
<!--
function checkMobile(){
	var sMobile = document.mobileform.mobile.value
	if(!(/^13[0-9]\d{4,8}$/.test(sMobile))){
		alert("不是完整的11位手机号或者正确的手机号前七位");
		document.mobileform.mobile.focus();
		return false;
	}
	window.open('', 'mobilewindow', 'height=197,width=350,status=yes,toolbar=no,menubar=no,location=no')
}
//-->
</script>


<html>
<head>
<title><%=biaoti%>-<%=title%></title>
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="keywords" content="<%=keywords%>">
<meta name="description" content="<%=description%>">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script type="text/javascript" src="Include/copy.js"></script>
</head>
<body>
<table width="980" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="214" height="100%" valign="top"><!--#include file=left.asp--></td>
	<td width="5">&nbsp;</td>　
<%id=trim(request("id"))
set rs=server.createobject("adodb.recordset")
sql = "select * from DNJY_info where id="&id&" and username='"&request.cookies("dick")("username")&"'"
rs.open sql,conn,1,3
if rs.eof Then
Response.Write "<CENTER><p>没有符合条件信息</CENTER>"
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
dizhi=rs("dizhi")
youbian=rs("youbian")
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
dizhi=rs("dizhi")
youbian=rs("youbian")
dqsj=rs("dqsj")
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
dim rs6,sql6,vip,sdkg,sdname,mpkg,mpname,mpfg
Set rs6= Server.CreateObject("ADODB.Recordset")
sql6="select vip,name,sdkg,sdname,mpkg,mpname,mpfg from [DNJY_user] where username='"&rs("username")&"'"
rs6.open sql6,conn,1,1
vip=rs6("vip")
sdkg=rs6("sdkg")
sdname=rs6("sdname")
mpkg=rs6("mpkg")
mpname=rs6("mpname")
mpfg=rs6("mpfg")
rs6.close
set rs6=Nothing
end If

if sdkg=1 And sdname<>"" Then
sdmba="<a href=""user/com.asp?com="&rs("username")&""" target=""_blank""><img src=""img/dian.gif"" title=""此会员已开有店铺"" ></a>"
end if
if mpkg=1 And mpname<>"" then
sdmba=sdmba+" <a href=""company.asp?username="&rs("username")&""" target=""_blank""><img src=""img/qy.gif"" title=""此会员已开有企业名片"" ></a>"
end if
if sdkg=0 And  mpkg=0 Then
sdmba="无店铺和企业名片或未经审核"
end if



If IsNull(rs("username")) Then
usernameid="非注册会员"
Else
usernameid=""&rs("username")&" <a href=preview.asp?username="&rs("username")&"&id="&id&">阅其所有信息</a>"
End if

if vip=1 Then
namea="<img src=""images/vip.gif"" alt=""验证VIP用户"" width=""13"" height=""13"" border=""0"">&nbsp;"
end If
namea=namea&""&rs("name")&""
if rs("c")=0 then
xxtp="<script language=javascript src=""adjs_path(""ads/js"",""info4"","&c1&"_"&c2&"_"&c3&")""></script>"
else 
if rs("c")=1 and not rs("tupian")="0" then
xxtp="<a target=""_blank"" title=""点击放大&gt;&gt;&gt;"" href=""UploadFiles/infopic/"&rs("tupian")&"""><img src=""UploadFiles/infopic/"&rs("tupian")&""" alt=""点击放大"" width=""250"" height=""200"" border=""0""><br><center>[信息图片 点击放大]</a>"
end if
end if


	belongtype="<a href=""list.asp?t="&rs("type_oneid")&"&"&C_ID&""">"&rs("type_one")&"</a>"
	IF rs("type_two")<>"" and not isnull(rs("type_two")) Then belongtype=belongtype&" / <a href=""list.asp?t="&rs("type_oneid")&","&rs("type_twoid")&"&"&C_ID&""">"&rs("type_two")&"</a>"
	IF rs("type_three")<>"" and not isnull(rs("type_three")) Then belongtype=belongtype&" / <a href=""list.asp?t="&rs("type_oneid")&","&rs("type_twoid")&","&rs("type_threeid")&"&"&C_ID&""">"&rs("type_three")&"</a>"	


    belongcity="<a href=""index.asp?c="&rs("city_oneid")&""">"&rs("city_one")&"</a>"
    IF rs("city_two")<>"" and not isnull(rs("city_two")) Then belongcity=belongcity&" / <a href=""index.asp?c="&rs("city_oneid")&","&rs("city_twoid")&""">"&rs("city_two")&"</a>"
	IF rs("city_three")<>"" and not isnull(rs("city_three")) Then belongcity=belongcity&" / <a href=""index.asp?c="&rs("city_oneid")&","&rs("city_twoid")&","&rs("city_threeid")&""">"&rs("city_three")&"</a>"%>
    <td width="760" valign="top" bgcolor="#FFFFFF">
<table align="center" bgcolor="#ffffff" class="ty1"><tr><td ><script language=javascript src="<%=adjs_path("ads/js","info1",c1&"_"&c2&"_"&c3)%>"></script></td></tr></table>
      <table width="760" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
      </table>
 <table width="760" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" class="ty1">
      <tr>
        <td bgcolor="#FFFFFF" ><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
          <tr>
            <td height="5"></td>
            </tr>
          <tr>
            <td height="28" ><table width="100%" border="0" cellspacing="0" cellpadding="5">
 <tr>
<td height="38" colspan="4" align="center" style="border-top: 1px solid #C0C0C0; padding-left: 0; padding-right: 0; padding-top: 1px; padding-bottom: 1px">
<font size="3"><b>

<%
response.write "["&leixin&"] "&biaotixs&""%>
</td>
</tr>
<tr><td height="27" colspan="4" style="border-bottom: 1px solid #C0C0C0; padding: 1px">
<p align="center"><b>发布时间：</b><%=fbsj%> <%if rs("dqsj")<>"" Then
sj=DateDiff("d",now(),""&rs("dqsj")&"")
if sj>0 then
response.Write "<font color=#FF0000> 剩余"&sj&"</b>天</font>"
elseif sj=0 then
response.Write "<font color=""#414141"">今日到期</font>"
elseif sj<0 then
response.Write "<font color=""#FF0000""> 已经过期</font>"
end If
end If%>　<b>浏览/回复：</b><%response.Write rs("llcs")%>次/<%response.Write rs("hfcs")%>次  不良记录：<SCRIPT src="user/xinfodt.asp?action=Bad&id=<%=id%>"></SCRIPT>条 <a href="Bad_info_list.asp?infoid=<%=cstr(id)%>&biaoti=<%=biaoti%>" target=blank>查</a> [<a title=凭删除密码可以删除此信息 href="#" ONCLICK="window.open('../user_delxx.asp?id=<%=id%>&key=del','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=300,height=200,left=300,top=100')">删除</a>]</td></tr>
            </table></td>
            </tr></table>

              <table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="260" valign="top"> <table width="250" height="161" border="0" cellpadding="0" cellspacing="2" bgcolor="#E4E4E4">
                      <tr> 
                        <td height="250" align="center" bgcolor="#FFFFFF"><%=xxtp%></td>
                      </tr>
                    </table></td>
                  <td valign="top"><table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#E4E4E4">
                      <tr bgcolor="#FFFFFF"> 
                        <td width="50%">&nbsp;信息类别：<%=belongtype%></td>
                        <td>&nbsp;联 系 人：<%=namea%></td>
                      </tr>
                      <tr bgcolor="#FFFFFF"> 
                        <td>&nbsp;交易类型：[<%=leixin%>]</td>
                        <td>&nbsp;会员ID号：<%=usernameid%></td>
                      </tr>
                      <tr bgcolor="#FFFFFF"> 
                        <td>&nbsp;交易状态：<SCRIPT src="user/xinfodt.asp?action=zt&id=<%=id%>"></SCRIPT></td>
                        <td>&nbsp;固定电话：<%=dianhua%></td>
						
                      </tr>
                      <tr bgcolor="#FFFFFF"> 
                        <td>&nbsp;市场价格：<%=scjiage%>/本站价格：<%=jiage%></td>
                         <form action="http://www.ip138.com:8080/search.asp" method="get" onSubmit="return checkMobile();" target="mobilewindow" name="mobileform"><td>&nbsp;移动电话：<input type="hidden" name="action" value="mobile">
			  <input type="text" name="mobile" size="16" value="<%=mobile%>" style="border: 1px solid #FFFFFF">
			  <INPUT type="submit" value="属地" name="Submit">
                        </td></form>
                      </tr>
                      <tr bgcolor="#FFFFFF"> 
                        <td>&nbsp;交易地区：<%=belongcity%></td>
                        <td>&nbsp;电子邮件：<%=email%> <a href="mailto:<%=email%>?subject=我在<%=title%>看了您发布的信息后与您联系">发信给他</a></td>
                      </tr>
                      <tr bgcolor="#FFFFFF"> 
                        <td>&nbsp;商家店企：<%=sdmba%></td>
<td>&nbsp;Q Q 号码：<%=userqq%><a target=blank href=http://wpa.qq.com/msgrd?V=1&Uin=<%=userqq%>&Site=<%=web%>&Menu=yes><img border='0' SRC=http://wpa.qq.com/pa?p=1:<%=userqq%>:7 alt=''给我留言'></a></td>
                      </tr>
                      <tr bgcolor="#FFFFFF"> 
                        <form method="get" action="http://www.ip138.com/ips1388.asp" name="ipform" onsubmit="return checkIP();" target="_blank">
	<td>&nbsp;I P 地址：<input type="text" name="ip" size="15" value="<%=userip%>" style="border: 1px solid #FFFFFF"> <input type="submit" value="属地"><input type="hidden" name="action" value="2"></td></td></form>
                        <td>&nbsp;邮政编码：<%=youbian%></td>
                      </tr>
                      <tr bgcolor="#FFFFFF"> 
                        <td height="32" colspan="2">&nbsp;通讯地址：<%=dizhi%></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td colspan="2" height="7"></td>
                </tr>
              </table>
 			  <table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td valign="top" bgcolor="#FFFFFF" class="a14">
				  <table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#E4E4E4">
                      <tr bgcolor="#FFFFFF"> 
                        <td width="50%" bgcolor="#FCFCFC"><div style="margin:9px; font-size:14px; letter-spacing:1px; line-height:140%"><%=HTMLDecode(memo)%></div></td></tr></table></td>
                </tr>
                <tr> 
                  <td height="7"></td>
                </tr>
              </table>
<table width="760" border="1" cellpadding="0" cellspacing="0" align="center">
        <tr>
          <td  style="border-style:none; padding:1px;" valign="top"><script language=javascript src="<%=adjs_path("ads/js","tail",c1&"_"&c2&"_"&c3)%>"></script></td>
</tr></table>             
<%Call CloseRs(rs)%>
<table width="760" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
              <tr>
                <td colspan="3" bgcolor="#FFFFFF" align="center">
                <INPUT class="inputb" TYPE=button VALUE="发站内短信给用户" ONCLICK="window.open('messh.asp?name=<%=username%>','Sample')">
                <INPUT class="inputb" TYPE=button VALUE="加入我的收藏" ONCLICK="window.open('user/collection.asp?id=<%=Id%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=170,height=80,left=300,top=100')">
                <input type="button" name="Submit" onClick='copyToClipBoard()' value="复制推荐给好友">
				<INPUT class="inputb" TYPE=button VALUE="举报该信息" ONCLICK="window.open('user/Bad_info.asp?id=<%=Id%>','Sample','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=500,height=600,left=300,top=100')">
				</td>

              </tr>
              <tr>
                <td height="10" colspan="3" align="center">加入网摘：| <a href="javascript:window.open('http://shuqian.qq.com/post?from=3&title='+encodeURIComponent(document.title)+'&uri='+encodeURIComponent(document.location.href)+'&jumpback=2&noui=1','favit','width=930,height=470,left=50,top=50,toolbar=no,menubar=no,location=no,scrollbars=yes,status=yes,resizable=yes');void(0)" style="background:url('http://shuqian.qq.com/img/add.gif') no-repeat 0px 0px;height:23px;width:60px;padding:2px 2px 0px 20px;font-size:12px;">QQ书签</a> | <A title=功能强大的网络收藏夹，一秒钟操作就可以轻松实现保存带来的价值、分享带来的快乐 href="javascript:d=document;t=d.selection?(d.selection.type!='None'?d.selection.createRange().text:''):(d.getSelection?d.getSelection():'');void(keyit=window.open('http://www.365key.com/storeit.aspx?t='+escape(d.title)+'&u='+escape(d.location.href)+'&c='+escape(t),'keyit','scrollbars=no,width=475,height=575,left=75,top=20,status=no,resizable=yes'));keyit.focus();"><FONT color=darkorchid>365K<FONT color=#57c200>e</FONT>y</FONT></A> |  <A title="收藏到POCO 网摘http://share.poco.cn" href="javascript:d=document;t=d.selection?(d.selection.type!='None'?d.selection.createRange().text:''):(d.getSelection?d.getSelection():'');void(vivi=window.open('http://my.poco.cn/fav/storeIt.php?t='+escape(d.title)+'&u='+escape(d.location.href)+'&c='+escape(t),'_blank','scrollbars=no,width=475,height=575,left=75,top=20,status=no,resizable=yes'));"><FONT color=#4cb509>POCO网摘</FONT></A> |<A style="FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #000000; TEXT-DECORATION: none" href="javascript:window.open('http://cang.baidu.com/do/add?it='+encodeURIComponent(document.title.substring(0,76))+'&iu='+encodeURIComponent(location.href)+'&fr=ien#nw=1','_blank','scrollbars=no,width=600,height=450,left=75,top=20,status=no,resizable=yes'); void 0"><SPAN style="PADDING-RIGHT: 5px; PADDING-LEFT: 5px; FONT-SIZE: 12px; PADDING-BOTTOM: 0px; MARGIN-LEFT: 10px; CURSOR: pointer; PADDING-TOP: 5px">百度网摘</A> |</SPAN></td>
              </tr>

              <tr>
                <td height="10" colspan="3" bgcolor="#FFFFFF"></td>
              </tr>
<script language="javascript">
<!--
//验证邮箱正确性
function checkemail(mdz){
var str=mdz;
//在JavaScript中，正则表达式只能使用"/"开头和结束，不能使用双引号
var Expression=/\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*/;
var objExp=new RegExp(Expression);
if(objExp.test(str)==true){
return true;
}else{
return false;
}
}

function form_onsubmit() {
    if((!checkemail(form.mdz.value)) && (document.form.mdz.value!=0))
	{
    alert("您输入Email地址不正确，请重新输入!");
    document.form.mdz.focus();
    return false;
    }
if (document.form.neirong.value==0)
	{
	  alert("请填回复内容")
	  document.form.neirong.focus()
	  return false
	 }
}
// --></script>


            </table>

  <!---信息详细页底部一广告开始---->
<script language=javascript src="<%=adjs_path("ads/js","info9",c1&"_"&c2&"_"&c3)%>"></script>
  <!---信息详细页底部一广告结束---->



  </td>
            </tr>

        </table></td>
      </tr>
    </table></td>
    <td width="4" bgcolor="#E6E6E6"></td>
  </tr>
</table>
<table align="center" bgcolor="#ffffff">
<tr><td>
<script language=javascript src="<%=adjs_path("ads/js","tc",c1&"_"&c2&"_"&c3)%>"></script>
</td></tr>
</table>
</body>
</html>

<!--#include file="footer.asp" -->