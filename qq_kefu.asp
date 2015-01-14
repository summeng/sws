<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!-- saved from url=(0022)http://www.soomac.org/ -->
<HTML><HEAD><TITLE>右窗口浮动广告 兼容IE FF等</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%'QQ客服在线，适应新老标准网页声明
Dim myqq,kqq,kqq_name
kqq=qq
kqq_name=qq_name
'此文件不需要作任何修改
if kqq<>"" And (kefu=1 Or kefu=4) then%>
<SCRIPT language=javascript>
<!-- //横幅
var isDOM = (document.getElementById ? true : false);
var isIE4 = ((document.all && !isDOM) ? true : false);
var isNS4 = (document.layers ? true : false);
function getRef(id) {
        if (isDOM) return document.getElementById(id);
        if (isIE4) return document.all[id];
        if (isNS4) return document.layers[id];
}
var isNS = navigator.appName == "Netscape";
function moveRightEdge() {
        var yMenuFrom, yMenuTo, yOffset, timeoutNextCheck;
        if (isNS4) {
                yMenuFrom   = divMenu.top;
                yMenuTo     = windows.pageYOffset + 90;   //榜首位置
        } else if (isDOM) {
                yMenuFrom   = parseInt (divMenu.style.top, 10);
                yMenuTo     = (isNS ? window.pageYOffset : document.body.scrollTop) + 90; //榜首位置
        }
        timeoutNextCheck = 500;
        if (yMenuFrom != yMenuTo) {
                yOffset = Math.ceil(Math.abs(yMenuTo - yMenuFrom) / 10);
                if (yMenuTo < yMenuFrom)
                        yOffset = -yOffset;
                if (isNS4)
                        divMenu.top += yOffset;
                else if (isDOM)
                        divMenu.style.top = parseInt (divMenu.style.top, 10) + yOffset+"px";
                        timeoutNextCheck = 10;
        }
        setTimeout ("moveRightEdge()", timeoutNextCheck);
}
//-->
</SCRIPT>
</HEAD>
<BODY>
<table width=1500>
<tr>
	<td><DIV id=divMenu style="Z-INDEX:100;FILTER:alpha(opacity=90);WIDTH:150px;HEIGHT:220px;VISIBILITY:visible;POSITION:absolute;position:absolute">
<%Dim getType,boxstr
getType=request.ServerVariables("HTTP_USER_AGENT")
If instr(getType,"MSIE")>0 Then
response.write"<div ID=""floater"" style=""right: 5px; top: 0px; width: 48px; height: 48px"">" 
elseif instr(getType,"Firefox")>0 Or instr(getType,"Maxthon")>0 Or instr(getType,"Chrome")>0 Or instr(getType,"TaoBrowser")>0 then
response.write"<div ID=""floater"" style=""right: 80px; top: 0px; width: 48px; height: 48px"">"
Else
response.write"<div ID=""floater"" style=""right: 5px; top: 0px; width: 48px; height: 48px"">" 
end if%>

<TABLE cellSpacing=0 cellPadding=0 width=100% align=center border=0 id="table2">
  <TBODY>
  <TR>
    <TD>
	<IMG src="images/qq/up<%=kefuskin%>.gif" 
      width=118 useMap=#m_about02Map border=0></TD></TR></TBODY></TABLE>
<DIV id=tracq1>
<TABLE cellSpacing=0 cellPadding=0 width=100 align=center border=0 id="table3">
  <TBODY>
  <TR>
    <TD class=cpx12hei vAlign=center align=middle width=118 
    background="images/qq/mid<%=kefuskin%>.gif" 
    height=80>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0 id="table5">
        <TBODY>
        <TR>
  <TD height=5>
<SPAN id=ad_01>少等载入中...</SPAN>
</TD></TR></TBODY></TABLE></TD></TR>
<TR>
<TD>
<IMG src="images/qq/down<%=kefuskin%>.gif" align=center border=0></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>	  
</DIV></DIV></DIV></DIV><p>   
<SPAN class=spanclass id=span_ad_01>
<!--以下内容可以更改为你自己的任意内容-->
<%kqq=replace(kqq,"，",",")
if isnull(kqq_name) Or kqq_name="" then kqq_name=kqq
kqq_name=replace(kqq_name,"，",",")
kQQ_NAME=split(kqq_name,",")
kQQ=split(kqq,",")
for N=0 to UBound(kQQ)
MyQQ=MyQQ+kQQ(N)+":"
next%>
<table border="0" width="110" cellspacing="0" cellpadding="0">
<%for N=0 to UBound(kQQ)%>
<script>
document.write("<tr><td >&nbsp;&nbsp;<a class='c' target=blank href='tencent://message/?uin=<%=kQQ(n)%>&Site=<%=title%>&Menu=yes'><%if qqshow=3 or qqshow=5 or qqshow=6 then%><img border='0' SRC=http://wpa.qq.com/pa?p=1:<%=kQQ(n)%>:<%=qqimg%> alt='客服QQ'><br><%end if%><%if qqshow=0 or qqshow=4 or qqshow=6 then%>&nbsp;&nbsp;<%=kQQ(n)%><br><%end if%><%if qqshow=1 or qqshow=4 or qqshow=5 or qqshow=6 then%>&nbsp;&nbsp;<%=kQQ_name(n)%><%end if%></a></td></tr>");
</script>
<%next%>
</table> 
<!--可更改内容结束-->
</SPAN>
<SCRIPT>ad_01.innerHTML=span_ad_01.innerHTML;span_ad_01.innerHTML="";</SCRIPT>
</DIV>
<SCRIPT language=javascript>
<!-- //横幅
if (isNS4) {
        var divMenu = document["divMenu"];
        divMenu.top = top.pageYOffset + 110;
        divMenu.visibility = "visible";
        moveRightEdge();
} else if (isDOM) {
        var divMenu = getRef('divMenu');
        divMenu.style.top = (isNS ? window.pageYOffset : document.documentElement.scrollTop) + 110+"px";
        divMenu.style.visibility = "visible";
        moveRightEdge();
}
//-->
</SCRIPT>
<SCRIPT language=javascript>
<!--DIV居中代码
varcc=document.getElementById("divMenu");
divMenu.style.marginLeft=document.body.offsetWidth-divMenu.offsetWidth-40+"px";
//下面offsetHeight属性在FireFox下有问题改成clientHeight属性了// 另外document.body无效时换 document.documentElement
divMenu.style.marginTop=document.body.clientHeight-divMenu.clientHeight-120+"px";

function CloseQQ()
{
divMenu.style.display="none";
return true; 
}
//-->
</SCRIPT> </td>
</tr>
</table>
</BODY>
</HTML>
<%end if%>