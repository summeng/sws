<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="config.asp"-->
<!--#include file="Include/head.asp"-->
<!--#include file="Include/Function.asp"-->
<!--#include file="Include/label.asp"-->
<%'======================================
'程序名称：新闻（信息）图文幻灯片播放插件
'特点：可显示任意格式图片
'适用范围：天天供求信息发布管理系统或其它以数据库形式调用的新闻（信息）图文幻灯显示
'运行环境：asp+access
'调用方式：将本文件放到WEB根目录，在任何页面位置加上：<iframe marginwidth="0" marginheight="0" src="bd_hd_player.asp?MaxNum=10&PicWidth=280&PicHeight=250&Title=20&TextHeight=15" frameborder="0" width="280" height="250" scrolling="No"  align="center"></iframe>
'参数说明：MaxNum -- 图片显示数量
'          TextHeight -- 标题文字高度
'　　　　　PicWidth -- 图片宽度
'　　　　　PicHeight -- 图片高度
'　　　　　Title -- 标题字符长度
'
MaxNum=trim(request("MaxNum"))
PicWidth=trim(request("PicWidth"))
PicHeight=trim(request("PicHeight"))
Title=strint(request("Title"))
TextHeight=trim(request("TextHeight"))

'开发：天天网络科技(http://www.ip126.com)
'免费使用，请保留版权信息
'联系信箱：bbfhq@sohu.com
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" >
<title>新闻（信息）图文幻灯显示</title>
<style type="text/css">
body,th,td{font-size:12px}
a:link {
	text-decoration: none;color:#000000;
}
a:visited {
	text-decoration: none;color:#000000;
}
a:hover {
	text-decoration: none;color:#FF0000;
}
a:active {
	text-decoration: none;color:#FF0000;
}
</style>
</head>

<body>
<%Call OpenConn
set rsbdhd=server.createobject("ADODB.recordset")
sqlbdhd="select  top 100 * from [DNJY_info] where yz=1 and (tupian <>'0')"&ttcc&" order by fbsj "&DN_OrderType&""
rsbdhd.open sqlbdhd,conn,3,3
if rsbdhd.eof Then
Response.Write "<p><br>暂无"
IF c1>0 Then Response.Write conn.Execute("Select city From DNJY_city Where id="&c1&" and twoid="&c2&" and threeid="&c3&"")(0)
Response.Write "最新图片信息!"
Call CloseRs(rsbdhd)
else%>

<SCRIPT LANGUAGE="JavaScript">
<!-- Hide
function killErrors() {
return true;
}
window.onerror=killErrors;
// -->
</SCRIPT>
<script language=JavaScript>
  var imgUrl=new Array();
  var imgLink=new Array();
  var ImgTitle=new Array();
  var adNum=0;
<%dim rsbdhd,sqlbdhd,MaxNum,TextHeight,PicWidth,PicHeight,AdID,AdUrl,AdTitle,AdPic,k


//以下三行为数据库调用，可根据你数据库的情况修改
set rsbdhd=server.createobject("ADODB.recordset")
sqlbdhd="select top 100 * from [DNJY_info] where yz=1 and (tupian <>'0') order by fbsj "&DN_OrderType&""
rsbdhd.open sqlbdhd,conn,3,3

	k=rsbdhd.RecordCount
	If K>MaxNum Then k=MaxNum
	i=1
do while not rsbdhd.eof
	AdID = rsbdhd("id") //ID为信息ID号字段，可根据你数据库的情况修改
	AdTitle = CutStr(rsbdhd("biaoti"),Title,"...") //bobai为信息标题字段，可根据你数据库的情况修改
	AdPic = "UploadFiles/infopic/"& rsbdhd("tupian") //图片物理路径，UploadFiles/infopic目录和tupian字段可根据你数据库及网站目录的情况修改
	AdUrl = ""&x_path("",rsbdhd("id"),"",rsbdhd("Readinfo"))&"" //show.asp为超链接文件，可根据你网站文件的情况修改

	//以下代码请勿乱改，否则可能显示不正常===========================================================
	Response.Write"imgUrl["&i&"]=""" & AdPic & """;" & vbCrLf
	Response.Write"ImgTitle["&i&"]=""" & AdTitle & """;" & vbCrLf
	Response.Write"imgLink["&i&"]=""" & AdUrl & """;" & vbCrLf
	If i>=k Then Exit Do	
	i=i+1	
	rsbdhd.movenext
    loop
Call CloseRs(rsbdhd)
%> 

function setTransition(){
  if (document.all){
     imgUrlrotator.filters.revealTrans.Transition=Math.floor(Math.random()*23);
     imgUrlrotator.filters.revealTrans.apply();
  }
}

function playTransition(){
  if (document.all)
     imgUrlrotator.filters.revealTrans.play()
}

function nextAd(){
  if(adNum<<%=k%>)adNum++;
     else adNum=1;
  setTransition();
  document.images.imgUrlrotator.src=imgUrl[adNum];
  document.all.tags("div")[0].innerText=ImgTitle[adNum]
  playTransition();
  theTimer=setTimeout("nextAd()", 8000);
}

function jump2url(){
  jumpUrl=imgLink[adNum];
  jumpTarget='_blank';
  if (jumpUrl != ''){
     if (jumpTarget != '')window.open(jumpUrl,jumpTarget);
     else location.href=jumpUrl;
  }
}
function displayStatusMsg() { 
  status=imgLink[adNum];
  document.returnValue = true;
}
window.onload=nextAd;/*这行代码作用是兼容IE7/8*/
//-->
</script>
<a onMouseOver="displayStatusMsg();return document.returnValue" href="javascript:jump2url()"><img style="FILTER: revealTrans(duration=2,transition=20)" height=<%=PicHeight%> src="javascript:nextAd()" width=<%=PicWidth%> border=0 name=imgUrlrotator>
<div style=" background-color:#FFFFFF;width:<%=PicWidth-1%>px; text-align:center;text-valign:middle;cursor:hand;line-height:<%=TextHeight%>px;"></div></a>
<%End If
Call CloseDB(conn)%>
</body>
</html>
