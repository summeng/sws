<!--#include file="../conn.asp"-->
<!--#include file="../SqlIn.Asp"-->
<!--#include file=cookies.asp-->
<!--#include file="../config.asp"-->
<!--#include file="../Include/head.asp"-->
<!--#include file="../Include/Function.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室出品
'客服中心http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>

<html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<title><%=title%>--后台管理系统</title>
<SCRIPT language=javascript>
 <!--
 var startTime,endTime;
 var d=new Date();
 startTime=d.getTime();
 //-->
 </SCRIPT>
 
 <body style="MARGIN: 0px" scroll=no >
<script>
if(self!=top){top.location=self.location;}
function switchSysBar(){
if (switchPoint.innerText==3){
switchPoint.innerText=4
document.all("frmTitle").style.display="none"
}else{
switchPoint.innerText=3
document.all("frmTitle").style.display=""
}}
</script>

<style type="text/css">.navPoint {COLOR: white; CURSOR: hand; FONT-FAMILY: Webdings; FONT-SIZE: 9pt}
</style>

<table border="0" cellPadding="0" cellSpacing="0" height="100%" width="100%">
  <tr>
    <td align="middle" noWrap vAlign="center" id="frmTitle">
    
    <iframe frameBorder="0" id="left" name="left" scrolling=auto src="menu.asp" style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 180px; Z-INDEX: 2">
    </iframe>
    
    

    </td>
    <td bgcolor="A4B6D7" style="WIDTH: 9pt">
    <table border="0" cellPadding="0" cellSpacing="0" height="100%">
      <tr>
        <td style="HEIGHT: 100%" onclick="switchSysBar()">
        <font style="FONT-SIZE: 9pt; CURSOR: default; COLOR: #ffffff">
        <span class="navPoint" id="switchPoint" title="切换">3</span><br>
        </font></td>
      </tr>
    </table>
    </td>
    <td style="WIDTH: 100%">
<iframe frameBorder="0" id="topFrame" name="topFrame" scrolling="auto" src="top.asp"  style="HEIGHT: 30; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1">
    </iframe>
<iframe frameBorder="0" id="right" name="right" scrolling="auto" src="right.asp" style="HEIGHT: 96%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1">
    </iframe>

</head>
    </td>
  </tr>
</table>
</html>
<script>
if(window.screen.width<'1024'){switchSysBar()}
</script>