<!--#include file="conn.asp"-->
<!--#include file="SqlIn.Asp"-->
<!--#include file="top.asp"-->
<%
'=====================================
'天天供求信息网站管理系统
'天天网络科技工作室版权所有
'官方客服中心 http://www.ip126.com
'QQ:530051328 mail:bdunion@126.com
'=====================================
%>
<!-- [DNJY] -->
<%dim b,bb,tj,dh
id=trim(request("id"))%>
<body topmargin="0" leftmargin="0" style="text-align:center;">
<SCRIPT LANGUAGE="JavaScript">
<!--//目的是为了做风格方便
document.write('<div style="clear:both;"></div><div class="wrap">');
//-->
</SCRIPT>
<body topmargin="0" leftmargin="0" style="text-align:center;">
      <table width="980" align="center" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td class="middle" align="center" valign="top"><script src="<%=adjs_path("ads/js","toppic",c1&"_"&c2&"_"&c3)%>"></script></td>
        </tr>
      </table>
<table width="10" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="6"></td>
  </tr>
</table>
<table align="center" bgcolor="#ffffff" width="980" class="ty1">
<TR>
	<TD>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" class="MainTable" id="Index_Search">
  <tr> 
    <td width="75%" height="60" valign="top" class="Main">
      <table width="100%" border="0" cellspacing="0" cellpadding="0" class="tablest">
        <tr> 
          <td class="middle" align="center" valign="top"> 
            <div class="top"><span style="float:left;padding-top:2px;"><%=f_search(1,1)%></div>
        </td>
        </tr>
      </table>
    </td>
    <td width="25%" height="60" valign="top" class="Side" align="right">
<a href="<%if request.cookies("dick")("username")<>"" then %><%=adduserinfo%>?<%=CT_ID%><%Elseif addxinxi=1 then%><%=addinfo%>?<%=CT_ID%><%else%>level.asp?<%=CT_ID%><%End if%>"><img src="skin/ltby/addxx.jpg" width=190 height=52 border=0></a>
    </td>
  </tr>
</table>
<div style="width:470px;float:left;padding-right:5px;">
<!---头条图片标题开始---->
<table width="100%" cellspacing="0" cellpadding="0" >
<tr><td width="490" > <%=f_tpbt(c1,c2,c3)%> </td>
</tr></table>
<!---头条图片标题结束---->
<table width="100%" border="0" cellspacing="0" cellpadding="0" id="IndexMainNews" style="margin-top:5px;">
              <tr> 
                <td></td>
                <td>
 <!--标签显示代码开始-->
<SCRIPT language=javascript>
function secBoardq(n)
{
for(i=0;i<secTableq.cells.length;i++) {
secTableq.cells[i].className="sec11";
secTableq.cells[n].className="sec22";}
for(i=0;i<mainTableq.tBodies.length;i++) {
mainTableq.tBodies[i].style.display="none";
mainTableq.tBodies[n].style.display="block";}
}
</SCRIPT>
            <table width="470" height="153"  align="center" cellpadding="0" cellspacing="0"  class="tablest">
              <tr>
                <td height="32" valign="top" style="border-style:none;" class="sectitlebg"><table width="470" height="32" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr id="secTableq">
                      <td width="105" height="30" class="sec22" style="border-style:none" onmouseover=secBoardq(0) ><a href="list.asp?<%=CT_ID%>&xxsx=1"><div align="center">最新信息 </div></a></td>
                      <td width="105" class="sec11" style="border-style:none" onmouseover=secBoardq(1)><a href="list.asp?<%=CT_ID%>&xxsx=2"><div align="center">竞价信息</div></a></td>
                      <td width="105" class="sec11" style="border-style:none" onmouseover=secBoardq(2)><a href="list.asp?<%=CT_ID%>&xxsx=3"><div align="center">推荐信息</div></a></td>
                    <td width="105" class="sec11" style="border-style:none" onmouseover=secBoardq(3)><a href="list.asp?<%=CT_ID%>&xxsx=4"><div align="center">热门信息</div></a></td>
                      <td width="200" class="sec11" style="border-style:none"></td>
                    </tr>
                  </table></td>
              </tr>
              <tr>
                <td height="320" style="border-style:none;" valign="top"><table id="mainTableq" width="470" border="0" align="center" cellpadding="0" cellspacing="0">
                    <!--最新信息-->
                    <TBODY style="DISPLAY: block">
                      <tr>
                        <td width="470" style="padding:6px;"><%=f_info(1,0,0,0,c1,c2,c3,1,1,1,1,13,1,1,1,1,0,40,12,1,13,11)%></td>
                      </tr>
                    </TBODY>
                    <!--竞价信息-->
                    <TBODY style="DISPLAY: none">
                      <tr>
                        <td style="padding:6px;"><%=f_info(2,0,0,0,c1,c2,c3,1,1,0,1,13,1,1,1,1,0,38,12,1,13,11)%></td>
                      </tr>
                    </TBODY>
                    <!--推荐信息-->
                    <TBODY style="DISPLAY: none">
                      <tr>
                        <td style="padding:6px;" valign="top"><%=f_info(3,0,0,0,c1,c2,c3,1,1,0,1,13,0,1,1,1,1,40,12,1,13,11)%></td>
                      </tr>
                    </TBODY>
					<!--热门信息-->
                    <TBODY style="DISPLAY: none">
                      <tr>
                        <td style="padding:6px;" valign="top"><%=f_info(4,0,0,0,c1,c2,c3,1,1,0,1,13,0,1,1,1,1,40,12,1,13,11)%></td>
                      </tr>
                    </TBODY>
                  </table></td>
              </tr>
            </table>
      <!--标签显示代码结束-->
</td>
                <td></td>
              </tr>
            </table>         
</div>

<div style="width:490px;float:right;">
<table width="100%" cellspacing="0" cellpadding="0" >
<tr><td width="490" ><script language=javascript src="<%=adjs_path("ads/js","zmd",c1&"_"&c2&"_"&c3)%>"></script> </td>
</tr></table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0" class="dragTable" style="margin-top:10px;">
              <tr> 
                <td class="head"> 
                  <h3 class="L"></h3>
                  <span class="TAG">图片信息</span> 
                  <h3 class="R"></h3>
                  <table border="0" cellspacing="0" cellpadding="0" style="float:right;" class="rmenu">
                    <tr>
                      <td class="LL"></td>
                      <td class="CC">
                        <div class="ch" id="showpicmu1" onclick="showpic(1)">最新</div>
                        <div id="showpicmu2" onclick="showpic(2)">推荐</div>
                        <div id="showpicmu3" onclick="showpic(3)">竞价</div>
                        <div id="showpicmu4" onclick="showpic(4)">热门</div>
                        <div id="showpicmu5" onclick="showpic(5)">置顶</div>
                      </td>
                      <td class="RR"></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr> 
                <td class="middle" align="left"> 
<div id="showpic1"><table width='100%' border='0' cellspacing='0' cellpadding='0'><tr><td width=20%><%=f_tptj(c1,c2,c3,1,3,9)%></td></tr></table> </div>
<div id="showpic2"><table width='100%' border='0' cellspacing='0' cellpadding='0'><tr><td width=20%><%=f_tptj(c1,c2,c3,3,3,9)%></td></tr></table></div>
<div id="showpic3"><table width='100%' border='0' cellspacing='0' cellpadding='0'><tr><td width=20%><%=f_tptj(c1,c2,c3,2,3,9)%></td></tr></table></div>
<div id="showpic4"><table width='100%' border='0' cellspacing='0' cellpadding='0'><tr><td width=20%><%=f_tptj(c1,c2,c3,4,3,9)%></td></tr></table></div>
<div id="showpic5"><table width='100%' border='0' cellspacing='0' cellpadding='0'><tr><td width=20%><%=f_tptj(c1,c2,c3,5,3,9)%></td></tr></table></div>
    <SCRIPT LANGUAGE="JavaScript">
<!--
function showpic(t){
	for(var i=1;i<6 ;i++ ){
		if(i==t){
			document.getElementById("showpic"+i).style.display='';			
			if(i==5){
				document.getElementById("showpicmu"+i).className='ch f';
			}else{
				document.getElementById("showpicmu"+i).className='ch';
			}
		}else{
			document.getElementById("showpic"+i).style.display='none';			
			if(i==5){
				document.getElementById("showpicmu"+i).className='f';
			}else{
				document.getElementById("showpicmu"+i).className='';
			}
		}
	}
}
showpic(1);
//-->
</SCRIPT>
                </td>
              </tr>
              <tr> 
                <td class="foot"> 
                  <h3 class="L"></h3>
                  <h3 class="R"></h3>
                </td>
              </tr>
            </table>
</div>
</div>
</TD>
</TR>
</TABLE>

<table width="10" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="6"></td>
  </tr>
</table>
      <table width="980" align="center" border="0" cellspacing="0" cellpadding="0" class="ty1">
        <tr> 
          <td class="middle" align="center" valign="top"><script src="<%=adjs_path("ads/js","bigpic",c1&"_"&c2&"_"&c3)%>"></script></td>
        </tr>
      </table>
<table width="10" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="6"></td>
  </tr>
</table>
 <table align="center" bgcolor="#ffffff" width="980" class="ty1">
<tr><td width="980"><div class="td2">
  <strong>竞价信息↓</strong>                       <span style="color:#999999"><A href="list.asp?xxsx=2&<%=C_ID%>">更多>>></a></span>
</div></td></tr>
  <tr>
   <td ><%=fkss(c1,c2,c3,1,1,1,1,1,1,0,1,1,20,50,5,10,1)%></td>
  </tr>
  </table>
 
      <table width="980" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td  height="7"></td>
        </tr>
</table>
<TABLE width="980" align="center" border="0" cellspacing="0" cellpadding="0">
<TR>
	<TD width="200" class="ty1" valign="top" align="center"><%=s_more(0,1,1,0,0,1,0,1,11,9,10,8)%></TD>
	<td width="5"></td>
	<TD width="775" class="ty1" valign="top" align="center"><%=xxfl(c1,c2,c3,1,12,8,2,38,1,1,1,0,5,0,1)%></TD>
</TR>
</TABLE> 
<table width="10" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="6"></td>
  </tr>
</table>
      <table width="980" align="center" border="0" cellspacing="0" cellpadding="0" class="ty1">
        <tr> 
          <td class="middle" align="center" valign="top"><script language=javascript src="<%=adjs_path("ads/js","bottom",c1&"_"&c2&"_"&c3)%>"></script></td>
        </tr>
      </table>
<!--#include file="footer.asp" -->
<!--#include file="qq_kefu.asp"-->
<script language='javascript' src='/Resetting.js'></script>