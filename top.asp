<!--#include file="config.asp"-->
<!--#include file="Include/tt_auto_lock.asp" -->
<!--#include file="Include/filename.asp"-->
<!--#include file="Include/head.asp"-->
<!--#include file="Include/Function.asp"-->
<!--#include file="Include/label.asp"-->
<!--#include file =Include/Version.asp -->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室版权所有
'官方客服中心 http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<head>
<title><%=ttkj_title%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<Meta name="Author" Content="<%=title%>,<%=web%>">   
<meta name="keywords" content="<%=ttkj_keywords%>">
<meta name="description" content="<%=ttkj_description%>">
<link rel="icon" href="/favicon.ico" type="image/x-icon"/>
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon"/>
<link href="/<%=strInstallDir%>skin/ltby/style.css" rel="stylesheet" type="text/css" >
<link href="/<%=strInstallDir%>skin/ltby/css.css" rel="stylesheet" type="text/css">
<link href="/<%=strInstallDir%>css/style.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="Include/hdxf.js"></script>
<SCRIPT src="Include/sinaflash.js" type="text/javascript"></SCRIPT>
<SCRIPT src="Include/sinaflashfp.js" type="text/javascript"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
//屏蔽可忽略的js脚本错误
function killErr(){
	return true;
}
window.onerror=killErr;
</SCRIPT>
<script type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
<SCRIPT language=javascript>
 window.onload=function(){
var _1=document.getElementById("h_qiehuan").getElementsByTagName("li");
for(var i=0;i<_1.length;i++){
_1[i].onmouseover=function(){
this.className+=(this.className.length>0?" ":"")+"cityhover";
if(document.getElementById("item")){document.getElementById("item").style.display="none";}
if(document.getElementById("icity")){document.getElementById("icity").style.display="none";}
};
_1[i].onmouseout=function(){
this.className=this.className.replace(new RegExp("( ?|^)cityhover\\b"),"");
if(document.getElementById("item")){document.getElementById("item").style.display="";}
if(document.getElementById("icity")){document.getElementById("icity").style.display="";}
};
}
}; 
function secBoard(n)
{
    for(i=1;i<secTable.cells.length-1;i++)
	
      secTable.cells[i].className="sec1";
    secTable.cells[n+1].className="sec2";
    for(i=0;i<mainTable.tBodies.length;i++)
      mainTable.tBodies[i].style.display="none";
    mainTable.tBodies[n-1].style.display="block";
	secTable.cells[1].className="";
  }
</SCRIPT>
<script language="javascript">function sendmail (address){ address = address.split("&");address = address.join("@");	window.open('mailto:'+address,'null','width=50,height=50,scrollbars=no,resizable=no,toolbar=no,location=no,directories=no,status=no,copyhistory=no');}</script>

<style type="text/css">
<!--
.STYLE1 {color: #000099}
-->
</style>
</head>
<%if lmkg1="0" Then%>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="980" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td>
      <div align="center">
        <p> </p><p> </p>
      <img src="images/stop_1.gif" width="433" height="207"><br><img src="images/stop_2.gif" width="357" height="52"><a href="http://<%=web%>"><img src="images/stop_3.gif" width="76" height="52" border="0"></a>
      <p> </p>
<font color="red" size=+2><%=(stopsm)%></font> 
      </div>
    </td>
   
  </tr>
</table>
</BODY>
</HTML>
<%response.End
End if%>

<table width="980" border="0" cellspacing="0" cellpadding="0" align="center" height="30" background="skin/ltby/indextop.gif">
  <tr>
    <td><table width="980" height="29" border="0" align="center" cellpadding="0" cellspacing="0" background="skin/ltby/indextop.gif">
      <tr>
        <td width="350">&nbsp;&nbsp;<span class="logintt"><script type="text/javascript">showcal(0)</script></span></td>
        <td width="260"></td>
        <td width="350"><table width="350" border="0" cellpadding="0" cellspacing="0">
          <tbody>
            <tr>
              <td width="350" class="grow_W"><div align="center"><font 
            color="#333333"><a href="city_list.asp?<%=CT_ID%>">城市联盟</a></font><font 
            color="#333333">- </font> <a href="type_list.asp?<%=CT_ID%>">信息目录</a><font 
            color="#333333"> - </font> <a href="#" onClick="javascript:window.external.AddFavorite('http://<%=web%>', '<%=title%>')" target="_self">收藏本站</a></div></td>
              
<td width="40" align="right"><a href="/<%=strInstallDir%>xml/Rss<%=c1%>_<%=c2%>_<%=c3%>.xml" target="_blank"><img border="0" src="img/rss.gif" width="31" height="15" alt="Rss2.0聚合服务页，当您安装了Rss阅读器后，您可以订阅本站最新信息更新"></a></td>
            </tr>
          </tbody>
        </table></td>
      </tr>
    </table></td>
  </tr>
</table>

<table width="980" border="0" cellspacing="0" cellpadding="0" align="center">
<div id="header">
<div class="logo"><a href="/<%=strInstallDir%>" target="_self"><img src="<%=l_path("UploadFiles/logo",c1&"_"&c2&"_"&c3)%>" alt="<%=webname%>" width="200" height="80" border="0"></a></div>
<div class="fyad">
<!--翻牌方式一广告代码开始-->
<script src="UploadFiles/adfp/sinaflash.js"></script>
<!--广告代码结束--></div>
<div class="ico_guide user"><a href="user.asp?<%=C_ID%>">会员中心</a></div>
<div class="ico_guide qymp"><a href="qyindex.asp?<%=C_ID%>">企业名片</a></div>
<div class="ico_guide vipnews"><a href="vipnews.asp?<%=C_ID%>">商家资讯</a></div>
<div class="ico_guide sjdp"><a href="sdindex.asp?<%=C_ID%>">商家店铺</a></div>
<div class="ico_guide magazine"><a href="magazine/magazine_show.asp?<%=C_ID%>">电子杂志</a></div>
<div class="ico_guide news"><a href="news.asp?<%=C_ID%>">新闻资讯</a></div>
<div class="ico_guide fbxx"><a href="xxindex.asp?<%=C_ID%>">供求频道</a></div>
<%if HtmlHome=1 then%>
<div class="ico_guide web"><a href="/<%=strInstallDir%>index.html">首 页</a></div>
<%else%>
<div class="ico_guide web"><a href="/<%=strInstallDir%>index.asp?<%=C_ID%>">首 页</a></div>
<%end if%>

</div>

</table>
<table width="980" height="85" border="0" align="center" cellpadding="0" cellspacing="0" background="skin/ltby/menubg.png">
<div id="menu">
  <tr>
    <td width="10" rowspan="2" valign="top"><img src="skin/ltby/menuleft.png" width="10" height="85" /></td>
    <td width="86" height="40"><img src="skin/ltby/menu_1.jpg" width="86" height="40" border="0" usemap="#Map" /></td>
    <td width="874">
	<table border="0" align="center" cellpadding="0" cellspacing="0">
      <%
'栏目分类
Call OpenConn
Dim sm
set rs=server.CreateObject("adodb.recordset")
sql="select id,[name],titlecolor from DNJY_type where id>0 and indexshow='yes' and twoid=0 and threeid=0 order by indexid"               
set rs=server.createobject("adodb.recordset")               
rs.open sql,conn,1,1
if rs.eof Then
Response.Write "<font color=000000>暂无分类或未设首页显示！</font>"
else
i=0                                                                                            
do while not rs.eof
Dim st
st=""
If rs("id")=t1 And t2=0 And t3=0 Then st="style=""color:red"";"
Response.Write "<td align=""center"" class=""bar_muen"" style=""color:#FFFFFF;width:75px"" >"
Response.Write "<A href=""list.asp?t="&rs(0)&",0,0&"&C_ID&"""><FONT "&st&" color="&rs(2)&"><b>"&rs(1)&"</b></FONT></A> "
Response.Write "</td><td valign=""top""><img src=""skin/ltby/menuline.png"" width=""3"" height=""40"" /></td>"
i=i+1
rs.movenext 
loop
if rs.recordcount<12 then
sm=12-rs.recordcount
rs.close:set rs=nothing
set rs=server.createobject("adodb.recordset")
sql="select id,twoid,titlecolor,[name] from DNJY_type where indexshow='yes' and id>0 and twoid>0 and threeid=0 order by indexid" 
rs.open sql,conn,1,1
do while not rs.eof
st=""
If rs("id")=t1 And rs("twoid")=t2 And t3=0 Then st="style=""color:red"";"
Response.Write "<td align=""center"" class=""bar_muen"" style=""color:#FFFFFF;width:75px"" >"
Response.Write "<A href=""list.asp?t="&rs(0)&","&rs(1)&",0&"&C_ID&"""><FONT "&st&" color="&rs(2)&"><b>"&rs(3)&"</b></FONT></A> "
Response.Write "</td><td valign=""top""><img src=""skin/ltby/menuline.png"" width=""3"" height=""40"" /></td>"
i=i+1  
sm=sm-1
if sm=0 then exit do 
rs.movenext 
Loop
Response.Write "</tr>"
if sm>0 then
rs.close:set rs=nothing
set rs=server.createobject("adodb.recordset")
sql="select id,twoid,threeid,[name],titlecolor from DNJY_type where id>0 and indexshow='yes' and twoid>0 and threeid>0 order by indexid" 
rs.open sql,conn,1,1 
do while sm=0 or not rs.eof
st=""
If rs("id")=t1 And rs("twoid")=t2 And rs("threeid")=t3 Then st="style=""color:red"";"
Response.Write "<td align=""center"" class=""bar_muen"" style=""color:#FFFFFF;width:75px"" >"
Response.Write "<A href=""list.asp?t="&rs(0)&","&rs(1)&","&rs(2)&"&"&C_ID&"""><FONT "&st&" color="&rs(4)&"><b>"&rs(3)&"</b></FONT></A> "
Response.Write "</td><td valign=""top""><img src=""skin/ltby/menuline.png"" width=""3"" height=""40"" /></td>"
i=i+1  
sm=sm-1
if sm=0 then exit do
rs.movenext 
loop
end If
end if
end if
rs.close:set rs=nothing
%>
    </table></td>
    <td width="9" rowspan="2" valign="top"><img src="skin/ltby/menuright.png" width="9" height="85" /></td>
  </tr>
  <tr>
    <td height="45" colspan="2"><table width="960" height="40" border="0" align="center" cellpadding="0" cellspacing="0" scrolling="no" scroll="no" topmargin="0" leftmargin="0" bottommargin="0" rightmargin="0" >
     
      <tr>
        <td width="100%" height="33" align="center"  valign="middle"><%if lmkg3="1" Then%><table height="0" border="0" align="left" cellPadding="0" cellSpacing="0">
        <tr>
          <td width="86"><FONT color="#800000"> <SPAN style="font-size: 10pt; font-weight: 700"><%=currentCity%></SPAN></FONT></td><td width="65" valign="margin-top"><%=qiehuan%></td><td width="5"></td>
          <td height="22"><%Response.Write navigate%></td><td width="5"></td>
        </tr>
      </table><%else%>欢迎光临<%=webname%><%end if%></td>

        </tr>
    </table></td>
  </tr></div>
</table>
<map name="Map" id="Map">
  <area shape="rect" coords="6,3,81,42" href="list.asp?<%=C_ID%>" />
</map>
<%dim count,username,rslx,sqllx,x,lx
Call CloseDB(conn)%>